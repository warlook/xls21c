#include "xls21c_addin.h"
#include "stdafx.h"
#include <vector>
#include <cctype>

struct CellRange {
    int startRow{0};
    int startCol{0};
    int endRow{-1};
    int endCol{-1};
    bool valid{false};
};

static int colNameToIndex(const std::string& s, size_t& pos) {
    int col = 0;
    while (pos < s.size() && std::isalpha(static_cast<unsigned char>(s[pos]))) {
        char c = static_cast<char>(std::toupper(static_cast<unsigned char>(s[pos])));
        col = col * 26 + (c - 'A' + 1);
        ++pos;
    }
    return col > 0 ? col - 1 : -1;
}

static CellRange parseRange(const std::string& range) {
    CellRange r;
    if (range.empty()) return r;

    size_t pos = 0;
    int sc = colNameToIndex(range, pos);
    if (sc < 0 || pos >= range.size() || !std::isdigit(static_cast<unsigned char>(range[pos]))) return r;
    int sr = 0;
    while (pos < range.size() && std::isdigit(static_cast<unsigned char>(range[pos]))) {
        sr = sr * 10 + (range[pos] - '0');
        ++pos;
    }
    if (sr < 1) return r;

    if (pos >= range.size()) {
        r.startRow = sr - 1;
        r.startCol = sc;
        r.endRow = r.startRow;
        r.endCol = r.startCol;
        r.valid = true;
        return r;
    }

    if (range[pos] != ':') return r;
    ++pos;

    int ec = colNameToIndex(range, pos);
    if (ec < 0 || pos >= range.size() || !std::isdigit(static_cast<unsigned char>(range[pos]))) return r;
    int er = 0;
    while (pos < range.size() && std::isdigit(static_cast<unsigned char>(range[pos]))) {
        er = er * 10 + (range[pos] - '0');
        ++pos;
    }
    if (er < 1) return r;

    r.startRow = sr - 1;
    r.startCol = sc;
    r.endRow = er - 1;
    r.endCol = ec;
    r.valid = true;
    return r;
}

Xls21cAddin::Xls21cAddin() = default;
Xls21cAddin::~Xls21cAddin() = default;

static std::vector<u16s> utf8ToUtf16(const std::string& utf8) {
    std::vector<u16s> result;
    result.reserve(utf8.size());
    size_t i = 0;
    while (i < utf8.size()) {
        unsigned char c = static_cast<unsigned char>(utf8[i]);
        unsigned int cp = 0;
        if (c < 0x80) {
            cp = c;
            i += 1;
        } else if ((c & 0xE0) == 0xC0 && i + 1 < utf8.size()) {
            cp = ((c & 0x1F) << 6) | (static_cast<unsigned char>(utf8[i + 1]) & 0x3F);
            i += 2;
        } else if ((c & 0xF0) == 0xE0 && i + 2 < utf8.size()) {
            cp = ((c & 0x0F) << 12) | ((static_cast<unsigned char>(utf8[i + 1]) & 0x3F) << 6) |
                 (static_cast<unsigned char>(utf8[i + 2]) & 0x3F);
            i += 3;
        } else if ((c & 0xF8) == 0xF0 && i + 3 < utf8.size()) {
            cp = ((c & 0x07) << 18) | ((static_cast<unsigned char>(utf8[i + 1]) & 0x3F) << 12) |
                 ((static_cast<unsigned char>(utf8[i + 2]) & 0x3F) << 6) |
                 (static_cast<unsigned char>(utf8[i + 3]) & 0x3F);
            i += 4;
        } else {
            cp = c;
            i += 1;
        }
        if (cp <= 0xFFFF) {
            result.push_back(static_cast<u16s>(cp));
        } else {
            cp -= 0x10000;
            result.push_back(static_cast<u16s>(0xD800 + (cp >> 10)));
            result.push_back(static_cast<u16s>(0xDC00 + (cp & 0x3FF)));
        }
    }
    return result;
}

static void writeEscapedValueTable(chunked_string_builder<u16s>& builder, const std::string& s) {
    auto chars = utf8ToUtf16(s);
    builder << u"{\"S\",\""_ss;
    for (u16s ch : chars) {
        if (ch == u'\"') {
            builder << u"\"\""_ss;
        } else {
            lstringu<1> c{expr_char<u16s>{ch}};
            builder << c;
        }
    }
    builder << u"\"},"_ss;
}

static void writeEscapedJson(chunked_string_builder<u16s>& builder, const std::string& s) {
    auto chars = utf8ToUtf16(s);
    builder << u"{\"#type\":\"jxs:string\",\"#value\":\""_ss;
    for (u16s ch : chars) {
        if (ch == u'\\') {
            builder << u"\\\\"_ss;
        } else if (ch == u'\"') {
            builder << u"\\\""_ss;
        } else {
            lstringu<1> c{expr_char<u16s>{ch}};
            builder << c;
        }
    }
    builder << u"\"}"_ss;
}

bool Xls21cAddin::OpenFile(tVariant* params, unsigned count) {
    if (params[0].vt != VTYPE_PWSTR) {
        return error(u"Ожидается строковый параметр", u"String parameter needed");
    }

    stru filePath = varToTextU(params[0]);
    std::wstring wpath(filePath.length(), L'\0');
    for (size_t i = 0; i < filePath.length(); ++i) {
        wpath[i] = static_cast<wchar_t>(filePath[i]);
    }

    if (!reader_.open(wpath)) {
        return error(u"Не удалось открыть файл", u"Failed to open file");
    }

    lastError_.make_empty();
    return true;
}

bool Xls21cAddin::CloseFile(tVariant* params, unsigned count) {
    reader_.close();
    lastError_.make_empty();
    return true;
}

bool Xls21cAddin::GetSheetName(tVariant& retVal, tVariant* params, unsigned count) {
    if (!reader_.isOpen()) {
        return error(u"Файл не открыт", u"File not opened");
    }
    if (params[0].vt != VTYPE_I4) {
        return error(u"Ожидается целочисленный параметр", u"Integer parameter needed");
    }

    int index = params[0].intVal;
    std::string name = reader_.sheetName(index);

    if (name.empty() && (index < 0 || index >= reader_.sheetCount())) {
        return error(u"Неверный индекс листа", u"Invalid sheet index");
    }

    retVal.vt = VTYPE_PWSTR;
    auto name16 = utf8ToUtf16(name);
    retVal.wstrLen = static_cast<int>(name16.size());
    memoryManager_->AllocMemory(reinterpret_cast<void**>(&retVal.pwstrVal),
                                static_cast<unsigned long>((name16.size() + 1) * sizeof(WCHAR_T)));
    auto* dst = reinterpret_cast<WCHAR_T*>(retVal.pwstrVal);
    for (size_t i = 0; i < name16.size(); ++i) {
        dst[i] = static_cast<WCHAR_T>(name16[i]);
    }
    dst[name16.size()] = 0;

    lastError_.make_empty();
    return true;
}

static void appendValueTableValue(chunked_string_builder<u16s>& text, const CellValue& cell) {
    switch (cell.type) {
    case CellValue::Empty:
        text << u"{\"L\"},"_ss;
        break;
    case CellValue::String:
        if (cell.text.empty()) {
            text << u"{\"L\"},"_ss;
        } else {
            writeEscapedValueTable(text, cell.text);
        }
        break;
    case CellValue::Number:
        text << u"{\"N\","_ss + cell.number + u"},"_ss;
        break;
    case CellValue::Boolean:
        text << u"{\"B\","_ss + (cell.boolean ? u"true"_ss : u"false"_ss) + u"},"_ss;
        break;
    case CellValue::Error:
        text << u"{\"L\"},"_ss;
        break;
    }
}

static bool readSheetAsValueTable(const SheetData& sheet, const CellRange& range, tVariant& retVal, IMemoryManager* mm) {
    chunked_string_builder<u16s> text;

    int sheetColCount = 0;
    if (!sheet.rows.empty()) {
        sheetColCount = static_cast<int>(sheet.rows[0].size());
    }

    int dataColCount;
    int firstRow, lastRow;
    int firstCol, lastCol;

    if (range.valid) {
        firstCol = std::max(0, range.startCol);
        lastCol = std::min(sheetColCount - 1, range.endCol);
        firstRow = std::max(0, range.startRow);
        lastRow = static_cast<int>(sheet.rows.size()) - 1;
        if (range.endRow < lastRow) lastRow = range.endRow;
        dataColCount = (lastCol >= firstCol) ? (lastCol - firstCol + 1) : 0;
    } else {
        firstCol = 0;
        lastCol = sheetColCount - 1;
        firstRow = 0;
        lastRow = static_cast<int>(sheet.rows.size()) - 1;
        dataColCount = sheetColCount;
    }

    unsigned totalColCount = static_cast<unsigned>(dataColCount + 1);

    text << u"{\"#\",acf6192e-81ca-46ef-93a6-5a6968b78663,{9,{"_ss + totalColCount + u","_ss;

    text << u"{0,\"НомерСтроки\",{\"Pattern\"},\"НомерСтроки\",0},"_ss;
    for (int i = 0; i < dataColCount; ++i) {
        lstringu<20> colName{u"k_"_ss + (i + 1)};
        text << u"{"_ss + (i + 1) + u",\"" + colName + u"\",{\"Pattern\"},\"" + colName + u"\",0},"_ss;
    }

    text << u"},{2,"_ss + totalColCount + u","_ss;
    for (unsigned i = 0; i < totalColCount; ++i) {
        lstringu<10> idx{eeu + i};
        text << idx + u","_ss + idx + u","_ss;
    }
    text << u"{1,"_ss;

    int rowCount = 0;
    for (int row = firstRow; row <= lastRow; ++row) {
        if (row < static_cast<int>(sheet.rows.size())) {
            bool hasData = false;
            for (int col = firstCol; col <= lastCol; ++col) {
                if (col < static_cast<int>(sheet.rows[row].size()) &&
                    sheet.rows[row][col].type != CellValue::Empty &&
                    !(sheet.rows[row][col].type == CellValue::String && sheet.rows[row][col].text.empty())) {
                    hasData = true;
                    break;
                }
            }
            if (hasData) ++rowCount;
        }
    }

    int startOfRowCount = static_cast<int>(text.length());
    text << expr_spaces<u16s, 20>{} + u","_ss;

    int outRow = 0;
    for (int row = firstRow; row <= lastRow; ++row) {
        if (row >= static_cast<int>(sheet.rows.size())) break;
        bool hasData = false;
        for (int col = firstCol; col <= lastCol; ++col) {
            if (col < static_cast<int>(sheet.rows[row].size()) &&
                sheet.rows[row][col].type != CellValue::Empty &&
                !(sheet.rows[row][col].type == CellValue::String && sheet.rows[row][col].text.empty())) {
                hasData = true;
                break;
            }
        }
        if (!hasData) continue;

        if (outRow > 0) {
            text << u"0},"_ss;
        }
        text << u"{2,"_ss + outRow + u","_ss + totalColCount + u","_ss;
        text << u"{\"N\","_ss + (outRow + 1) + u"},"_ss;
        for (int col = firstCol; col <= lastCol; ++col) {
            if (col < static_cast<int>(sheet.rows[row].size())) {
                appendValueTableValue(text, sheet.rows[row][col]);
            } else {
                text << u"{\"L\"},"_ss;
            }
        }
        ++outRow;
    }

    if (rowCount == 0) {
        text << u"{2,"_ss + totalColCount + u","_ss;
        for (unsigned i = 0; i < totalColCount; ++i) {
            lstringu<10> idx{eeu + i};
            text << idx + u","_ss + idx + u","_ss;
        }
        text << u"{1,0},2,-1},{0,0}}}"_ss;
    } else {
        text << u"0}},1,2},{0,0}}}"_ss;
    }

    retVal.vt = VTYPE_PWSTR;
    retVal.wstrLen = static_cast<int>(text.length());
    mm->AllocMemory(reinterpret_cast<void**>(&retVal.pwstrVal),
                     static_cast<unsigned long>((text.length() + 1) * sizeof(WCHAR_T)));
    *text.place(retVal.pwstrVal) = 0;

    if (rowCount > 0) {
        WCHAR_T* buf = retVal.pwstrVal;
        WCHAR_T num[12];
        int pos = 11;
        num[pos] = 0;
        unsigned n = rowCount;
        do {
            num[--pos] = L'0' + (n % 10);
            n /= 10;
        } while (n > 0);
        int digits = 11 - pos;
        WCHAR_T* dst = buf + startOfRowCount + 20 - digits;
        for (int i = 0; i < digits; ++i) {
            dst[i] = num[pos + i];
        }
    }

    return true;
}

static bool readSheetAsJson(const SheetData& sheet, const CellRange& range, tVariant& retVal, IMemoryManager* mm) {
    chunked_string_builder<u16s> text;

    text << u"{\"#type\":\"jv8:Array\",\"#value\":[{\"#type\":\"jv8:Array\",\"#value\":["_ss;

    int sheetColCount = 0;
    if (!sheet.rows.empty()) {
        sheetColCount = static_cast<int>(sheet.rows[0].size());
    }

    int firstRow, lastRow;
    int firstCol, lastCol;
    int dataColCount;

    if (range.valid) {
        firstCol = std::max(0, range.startCol);
        lastCol = std::min(sheetColCount - 1, range.endCol);
        firstRow = std::max(0, range.startRow);
        lastRow = static_cast<int>(sheet.rows.size()) - 1;
        if (range.endRow < lastRow) lastRow = range.endRow;
        dataColCount = (lastCol >= firstCol) ? (lastCol - firstCol + 1) : 0;
    } else {
        firstCol = 0;
        lastCol = sheetColCount - 1;
        firstRow = 0;
        lastRow = static_cast<int>(sheet.rows.size()) - 1;
        dataColCount = sheetColCount;
    }

    for (int i = 0; i < dataColCount; ++i) {
        if (i > 0) text << u","_ss;
        lstringu<20> colName{u"k_"_ss + (i + 1)};
        text << u"{\"#type\":\"jxs:string\",\"#value\":\""_ss + colName + u"\"}"_ss;
    }

    text << u"]}"_ss;

    for (int row = firstRow; row <= lastRow; ++row) {
        if (row >= static_cast<int>(sheet.rows.size())) break;
        bool hasData = false;
        for (int col = firstCol; col <= lastCol; ++col) {
            if (col < static_cast<int>(sheet.rows[row].size()) &&
                sheet.rows[row][col].type != CellValue::Empty &&
                !(sheet.rows[row][col].type == CellValue::String && sheet.rows[row][col].text.empty())) {
                hasData = true;
                break;
            }
        }
        if (!hasData) continue;

        text << u",{\"#type\":\"jv8:Array\",\"#value\":["_ss;
        for (int col = firstCol; col <= lastCol; ++col) {
            if (col != firstCol) text << u","_ss;
            if (col < static_cast<int>(sheet.rows[row].size())) {
                const auto& cell = sheet.rows[row][col];
                switch (cell.type) {
                case CellValue::Empty:
                    text << u"{\"#type\":\"jv8:Null\",\"#value\":\"\"}"_ss;
                    break;
                case CellValue::String:
                    if (cell.text.empty()) {
                        text << u"{\"#type\":\"jv8:Null\",\"#value\":\"\"}"_ss;
                    } else {
                        writeEscapedJson(text, cell.text);
                    }
                    break;
                case CellValue::Number:
                    text << u"{\"#type\":\"jxs:decimal\",\"#value\":"_ss + cell.number + u"}"_ss;
                    break;
                case CellValue::Boolean:
                    text << u"{\"#type\":\"jxs:boolean\",\"#value\":"_ss + (cell.boolean ? u"true"_ss : u"false"_ss) + u"}"_ss;
                    break;
                case CellValue::Error:
                    text << u"{\"#type\":\"jv8:Null\",\"#value\":\"\"}"_ss;
                    break;
                }
            } else {
                text << u"{\"#type\":\"jv8:Null\",\"#value\":\"\"}"_ss;
            }
        }
        text << u"]}"_ss;
    }

    text << u"]}"_ss;

    retVal.vt = VTYPE_PWSTR;
    retVal.wstrLen = static_cast<int>(text.length());
    mm->AllocMemory(reinterpret_cast<void**>(&retVal.pwstrVal),
                     static_cast<unsigned long>((text.length() + 1) * sizeof(WCHAR_T)));
    *text.place(retVal.pwstrVal) = 0;

    return true;
}

bool Xls21cAddin::ReadSheet(tVariant& retVal, tVariant* params, unsigned count) {
    if (!reader_.isOpen()) {
        return error(u"Файл не открыт", u"File not opened");
    }
    if (params[0].vt != VTYPE_I4) {
        return error(u"Первый параметр должен быть целым числом", u"First param must be integer");
    }
    if (params[1].vt != VTYPE_PWSTR) {
        return error(u"Второй параметр должен быть строкой", u"Second param must be string");
    }

    int index = params[0].intVal;
    stru format = varToTextU(params[1]);

    CellRange range;
    if (count >= 3 && params[2].vt == VTYPE_PWSTR) {
        stru rangeStr = varToTextU(params[2]);
        std::string rangeUtf8;
        rangeUtf8.reserve(rangeStr.length());
        for (size_t i = 0; i < rangeStr.length(); ++i) {
            rangeUtf8 += static_cast<char>(rangeStr[i]);
        }
        range = parseRange(rangeUtf8);
        if (!range.valid && !rangeUtf8.empty()) {
            return error(u"Неверный формат диапазона", u"Invalid range format");
        }
    }

    if (index < 0 || index >= reader_.sheetCount()) {
        return error(u"Неверный индекс листа", u"Invalid sheet index");
    }

    const SheetData* sheet = reader_.sheet(index);
    if (!sheet) {
        return error(u"Не удалось получить данные листа", u"Failed to get sheet data");
    }

    lastError_.make_empty();

    if (format.equal_ia(u"ValueTable") || format.equal_iu(u"ТаблицаЗначений")) {
        return readSheetAsValueTable(*sheet, range, retVal, memoryManager_);
    } else if (format.equal_ia(u"JSON")) {
        return readSheetAsJson(*sheet, range, retVal, memoryManager_);
    } else {
        return error(u"Неизвестный формат для результата", u"Unknown result format");
    }
}
