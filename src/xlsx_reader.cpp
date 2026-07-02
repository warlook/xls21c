#include "xlsx_reader.h"
#include "xls_helper.h"
#include <OpenXLSX.hpp>
#include <algorithm>
#include <cstdio>
#include <cstring>
#include <vector>

#ifdef _WIN32
#include <windows.h>
#endif

XlsxReader::~XlsxReader() {
    close();
}

bool XlsxReader::open(const std::wstring& filePath) {
    close();
    filePath_ = filePath;

    try {
        FILE* f = _wfopen(filePath.c_str(), L"rb");
        if (!f) {
            lastError_ = "Cannot open file";
            return false;
        }
        unsigned char sig[8] = {};
        fread(sig, 1, 8, f);
        fclose(f);

        if (sig[0] == 0x50 && sig[1] == 0x4B) {
            return readXlsx(filePath);
        } else if (sig[0] == 0xD0 && sig[1] == 0xCF) {
            return readXls(filePath);
        } else {
            lastError_ = "Unknown file format";
            return false;
        }
    } catch (const std::exception& e) {
        lastError_ = e.what();
        return false;
    }
}

void XlsxReader::close() {
    sheets_.clear();
    fileOpen_ = false;
    filePath_.clear();
    lastError_.clear();
}

std::string XlsxReader::sheetName(int index) const {
    if (index < 0 || index >= static_cast<int>(sheets_.size())) {
        return "";
    }
    return sheets_[index].name;
}

const SheetData* XlsxReader::sheet(int index) const {
    if (index < 0 || index >= static_cast<int>(sheets_.size())) {
        return nullptr;
    }
    return &sheets_[index];
}

static std::string wstringToUtf8(const std::wstring& ws) {
    std::string result;
    result.reserve(ws.size() * 3);
    for (wchar_t wc : ws) {
        if (wc < 0x80) {
            result += static_cast<char>(wc);
        } else if (wc < 0x800) {
            result += static_cast<char>(0xC0 | (wc >> 6));
            result += static_cast<char>(0x80 | (wc & 0x3F));
        } else {
            result += static_cast<char>(0xE0 | (wc >> 12));
            result += static_cast<char>(0x80 | ((wc >> 6) & 0x3F));
            result += static_cast<char>(0x80 | (wc & 0x3F));
        }
    }
    return result;
}

static bool copyFileToTemp(const std::wstring& src, std::string& tempPath) {
    FILE* fsrc = _wfopen(src.c_str(), L"rb");
    if (!fsrc) return false;

    char tempDir[MAX_PATH];
    GetTempPathA(MAX_PATH, tempDir);
    char tempName[MAX_PATH];
    GetTempFileNameA(tempDir, "xls", 0, tempName);

    std::string tmp(tempName);
    size_t dotPos = tmp.rfind('.');
    if (dotPos != std::string::npos) {
        tmp = tmp.substr(0, dotPos) + ".xlsx";
    } else {
        tmp += ".xlsx";
    }
    tempPath = tmp;

    FILE* fdst = fopen(tmp.c_str(), "wb");
    if (!fdst) { fclose(fsrc); return false; }

    char buf[8192];
    size_t n;
    while ((n = fread(buf, 1, sizeof(buf), fsrc)) > 0) {
        fwrite(buf, 1, n, fdst);
    }
    fclose(fsrc);
    fclose(fdst);
    return true;
}

bool XlsxReader::readXlsx(const std::wstring& path) {
    std::string tempPath;
    bool useTemp = false;

    try {
        std::string openPath;
        if (wstringToUtf8(path).find_first_of(static_cast<unsigned char>(0x80)) != std::string::npos) {
            if (copyFileToTemp(path, tempPath)) {
                openPath = tempPath;
                useTemp = true;
            } else {
                openPath = wstringToUtf8(path);
            }
        } else {
            openPath = wstringToUtf8(path);
        }

        OpenXLSX::XLDocument doc;
        doc.open(openPath);

        auto workbook = doc.workbook();
        auto sheetNames = workbook.sheetNames();

        sheets_.clear();
        sheets_.reserve(sheetNames.size());

        for (const auto& name : sheetNames) {
            SheetData sheetData;
            sheetData.name = name;

            auto worksheet = workbook.worksheet(name);
            uint32_t rowCount = worksheet.rowCount();
            uint32_t colCount = worksheet.columnCount();

            if (rowCount > 0 && colCount > 0) {
                sheetData.rows.reserve(rowCount);

                for (uint32_t row = 1; row <= rowCount; ++row) {
                    std::vector<CellValue> rowData;
                    rowData.reserve(colCount);

                    for (uint32_t col = 1; col <= colCount; ++col) {
                        auto cell = worksheet.cell(row, col);
                        CellValue cv;

                        auto vt = cell.value().type();
                        if (vt == OpenXLSX::XLValueType::Empty) {
                            cv.type = CellValue::Empty;
                        } else if (vt == OpenXLSX::XLValueType::String) {
                            cv.type = CellValue::String;
                            cv.text = cell.value().get<std::string>();
                        } else if (vt == OpenXLSX::XLValueType::Integer) {
                            cv.type = CellValue::Number;
                            cv.number = static_cast<double>(cell.value().get<int64_t>());
                        } else if (vt == OpenXLSX::XLValueType::Float) {
                            cv.type = CellValue::Number;
                            cv.number = cell.value().get<double>();
                        } else if (vt == OpenXLSX::XLValueType::Boolean) {
                            cv.type = CellValue::Boolean;
                            cv.boolean = cell.value().get<bool>();
                        } else if (vt == OpenXLSX::XLValueType::Error) {
                            cv.type = CellValue::Error;
                        }

                        rowData.push_back(std::move(cv));
                    }

                    sheetData.rows.push_back(std::move(rowData));
                }
            }

            sheets_.push_back(std::move(sheetData));
        }

        doc.close();
        if (useTemp) {
            remove(tempPath.c_str());
        }
        fileOpen_ = true;
        return true;
    } catch (const std::exception& e) {
        if (useTemp) {
            remove(tempPath.c_str());
        }
        lastError_ = e.what();
        return false;
    }
}

bool XlsxReader::readXls(const std::wstring& path) {
    FILE* f = _wfopen(path.c_str(), L"rb");
    if (!f) {
        lastError_ = "Cannot open XLS file";
        return false;
    }

    fseek(f, 0, SEEK_END);
    long fileSize = ftell(f);
    fseek(f, 0, SEEK_SET);

    std::vector<unsigned char> buffer(fileSize);
    size_t readBytes = fread(buffer.data(), 1, fileSize, f);
    fclose(f);

    if (static_cast<long>(readBytes) != fileSize) {
        lastError_ = "Failed to read XLS file";
        return false;
    }

    XlsData xlsData = {};
    if (xls_read_from_buffer(buffer.data(), buffer.size(), &xlsData) != 0) {
        lastError_ = "Failed to parse XLS file";
        xls_free_data(&xlsData);
        return false;
    }

    sheets_.clear();
    sheets_.reserve(xlsData.sheet_count);

    for (int si = 0; si < xlsData.sheet_count; si++) {
        XlsSheet* s = &xlsData.sheets[si];
        SheetData sheetData;
        sheetData.name = s->name ? s->name : ("Sheet" + std::to_string(si + 1));

        if (s->row_count > 0 && s->col_count > 0) {
            sheetData.rows.resize(s->row_count);
            for (int r = 0; r < s->row_count; r++) {
                sheetData.rows[r].resize(s->col_count);
                for (int c = 0; c < s->col_count; c++) {
                    CellValue cv;
                    XlsCellData* cd = &s->cells[r * s->col_count + c];
                    switch (cd->type) {
                    case XLCELL_NUMBER:
                        cv.type = CellValue::Number;
                        cv.number = cd->num;
                        break;
                    case XLCELL_BOOLEAN:
                        cv.type = CellValue::Boolean;
                        cv.boolean = (cd->bool_val != 0);
                        break;
                    case XLCELL_STRING:
                        cv.type = CellValue::String;
                        cv.text = cd->str ? cd->str : "";
                        break;
                    default:
                        cv.type = CellValue::Empty;
                        break;
                    }
                    sheetData.rows[r][c] = std::move(cv);
                }
            }
        }

        sheets_.push_back(std::move(sheetData));
    }

    xls_free_data(&xlsData);
    fileOpen_ = true;
    return true;
}
