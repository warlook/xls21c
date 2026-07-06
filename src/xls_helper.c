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
        s->name = strdup(wb->sheets.sheet[si].name ? wb->sheets.sheet[si].name : "");

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
                cd->str = (cell->str && strlen(cell->str) > 0) ? strdup(cell->str) : NULL;
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
