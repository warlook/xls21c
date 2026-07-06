#include "xls_helper.h"
#include <xls.h>
#include <stdlib.h>
#include <string.h>

#ifdef _MSC_VER
#include <windows.h>
#endif

#ifdef _MSC_VER
#define strdup _strdup
#endif

static const unsigned int cp1251_to_unicode[128] = {
    0x0402, 0x0403, 0x201A, 0x0453, 0x201E, 0x2026, 0x2020, 0x2021,
    0x20AC, 0x2030, 0x0409, 0x2039, 0x040A, 0x040C, 0x040B, 0x040F,
    0x0452, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
    0x0000, 0x2122, 0x0459, 0x203A, 0x045A, 0x045C, 0x045B, 0x045F,
    0x00A0, 0x040E, 0x045E, 0x0408, 0x00A4, 0x0490, 0x00A6, 0x00A7,
    0x0401, 0x00A9, 0x0404, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x0407,
    0x00B0, 0x00B1, 0x0406, 0x0456, 0x0491, 0x00B5, 0x00B6, 0x00B7,
    0x0451, 0x2116, 0x0454, 0x00BB, 0x0458, 0x0405, 0x0455, 0x0457,
    0x0410, 0x0411, 0x0412, 0x0413, 0x0414, 0x0415, 0x0416, 0x0417,
    0x0418, 0x0419, 0x041A, 0x041B, 0x041C, 0x041D, 0x041E, 0x041F,
    0x0420, 0x0421, 0x0422, 0x0423, 0x0424, 0x0425, 0x0426, 0x0427,
    0x0428, 0x0429, 0x042A, 0x042B, 0x042C, 0x042D, 0x042E, 0x042F,
    0x0430, 0x0431, 0x0432, 0x0433, 0x0434, 0x0435, 0x0436, 0x0437,
    0x0438, 0x0439, 0x043A, 0x043B, 0x043C, 0x043D, 0x043E, 0x043F,
    0x0440, 0x0441, 0x0442, 0x0443, 0x0444, 0x0445, 0x0446, 0x0447,
    0x0448, 0x0449, 0x044A, 0x044B, 0x044C, 0x044D, 0x044E, 0x044F
};

static int utf8_codepoint_len(unsigned int cp) {
    if (cp < 0x80) return 1;
    if (cp < 0x800) return 2;
    if (cp < 0x10000) return 3;
    return 4;
}

static void utf8_encode(unsigned int cp, char* out, int* len) {
    if (cp < 0x80) { out[0] = (char)cp; *len = 1; }
    else if (cp < 0x800) { out[0] = (char)(0xC0 | (cp >> 6)); out[1] = (char)(0x80 | (cp & 0x3F)); *len = 2; }
    else if (cp < 0x10000) { out[0] = (char)(0xE0 | (cp >> 12)); out[1] = (char)(0x80 | ((cp >> 6) & 0x3F)); out[2] = (char)(0x80 | (cp & 0x3F)); *len = 3; }
    else { out[0] = (char)(0xF0 | (cp >> 18)); out[1] = (char)(0x80 | ((cp >> 12) & 0x3F)); out[2] = (char)(0x80 | ((cp >> 6) & 0x3F)); out[3] = (char)(0x80 | (cp & 0x3F)); *len = 4; }
}

static unsigned int utf8_decode(const char* s, int* len) {
    unsigned char c = (unsigned char)s[0];
    if (c < 0x80) { *len = 1; return c; }
    if ((c & 0xE0) == 0xC0) { *len = 2; return ((c & 0x1F) << 6) | (s[1] & 0x3F); }
    if ((c & 0xF0) == 0xE0) { *len = 3; return ((c & 0x0F) << 12) | ((s[1] & 0x3F) << 6) | (s[2] & 0x3F); }
    *len = 4; return ((c & 0x07) << 18) | ((s[1] & 0x3F) << 12) | ((s[2] & 0x3F) << 6) | (s[3] & 0x3F);
}

static int has_high_bytes(const char* s) {
    for (const unsigned char* p = (const unsigned char*)s; *p; p++) {
        if (*p >= 0x80) return 1;
    }
    return 0;
}

static char* fix_cp1251_encoding(const char* s) {
    if (!s) return NULL;
    if (!has_high_bytes(s)) return strdup(s);

    int src_len = (int)strlen(s);
    char* buf = (char*)malloc(src_len * 3 + 1);
    if (!buf) return strdup(s);

    int out_pos = 0;
    for (int i = 0; i < src_len; i++) {
        unsigned char b = (unsigned char)s[i];
        unsigned int cp;
        if (b < 0x80) {
            cp = b;
        } else {
            cp = cp1251_to_unicode[b - 0x80];
        }
        int elen = 0;
        utf8_encode(cp, buf + out_pos, &elen);
        out_pos += elen;
    }
    buf[out_pos] = 0;
    return buf;
}

void xls_free_data(XlsData* data) {
    if (!data) return;
    for (int s = 0; s < data->sheet_count; s++) {
        XlsSheet* sh = &data->sheets[s];
        if (sh->cells) {
            for (int i = 0; i < sh->row_count * sh->col_count; i++) {
                free(sh->cells[i].str);
            }
            free(sh->cells);
        }
        free(sh->name);
    }
    free(data->sheets);
    data->sheets = NULL;
    data->sheet_count = 0;
}

static int xls_read_from_buffer_impl(const unsigned char* buf, size_t len, XlsData* out) {
    xls_error_t err;
    xlsWorkBook* wb = xls_open_buffer(buf, len, "UTF-8", &err);
    if (!wb) return -1;

    xls_parseWorkBook(wb);

    if (wb->codepage == 0 || wb->codepage == 1252) {
        wb->codepage = 1251;
        wb->converter = NULL;
    }

    out->sheet_count = (int)wb->sheets.count;
    out->sheets = (XlsSheet*)calloc(wb->sheets.count, sizeof(XlsSheet));
    if (!out->sheets) {
        xls_close_WB(wb);
        return -1;
    }

    for (DWORD si = 0; si < wb->sheets.count; si++) {
        xlsWorkSheet* ws = xls_getWorkSheet(wb, si);
        if (!ws) continue;

        xls_parseWorkSheet(ws);

        XlsSheet* s = &out->sheets[si];
        char* raw_name = wb->sheets.sheet[si].name ? wb->sheets.sheet[si].name : "";
        s->name = fix_cp1251_encoding(raw_name);

        int rowCount = ws->rows.lastrow + 1;
        int colCount = ws->rows.lastcol + 1;
        s->row_count = rowCount;
        s->col_count = colCount;
        s->cells = NULL;

        if (rowCount <= 0 || colCount <= 0) {
            xls_close_WS(ws);
            continue;
        }

        int totalCells = rowCount * colCount;
        s->cells = (XlsCellData*)calloc(totalCells, sizeof(XlsCellData));
        if (!s->cells) {
            xls_close_WS(ws);
            continue;
        }

        for (int row = 0; row < rowCount; row++) {
            if (row >= (int)ws->rows.lastrow + 1) break;
            xlsRow* r = &ws->rows.row[row];
            if (!r) continue;
            for (int col = 0; col < colCount; col++) {
                if (col > r->lcell) break;
                xlsCell* cell = &r->cells.cell[col];
                if (!cell) continue;

                XlsCellData* cd = &s->cells[row * colCount + col];
                if (cell->str && strlen(cell->str) > 0) {
                    cd->str = fix_cp1251_encoding(cell->str);
                } else {
                    cd->str = NULL;
                }
                cd->type = cd->str ? XLCELL_STRING : XLCELL_EMPTY;
            }
        }

        xls_close_WS(ws);
    }

    xls_close_WB(wb);
    return 0;
}

int xls_read_from_buffer(const unsigned char* buf, size_t len, XlsData* out) {
#ifdef _MSC_VER
    int result = -1;
    __try {
        result = xls_read_from_buffer_impl(buf, len, out);
    } __except(EXCEPTION_EXECUTE_HANDLER) {
        memset(out, 0, sizeof(XlsData));
        return -1;
    }
    return result;
#else
    return xls_read_from_buffer_impl(buf, len, out);
#endif
}
