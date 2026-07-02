#pragma once

#ifdef __cplusplus
extern "C" {
#endif

enum XlsCellType { XLCELL_EMPTY, XLCELL_STRING, XLCELL_NUMBER, XLCELL_BOOLEAN, XLCELL_ERROR };

typedef struct {
    enum XlsCellType type;
    char* str;
    double num;
    int bool_val;
} XlsCellData;

typedef struct {
    char* name;
    int row_count;
    int col_count;
    XlsCellData* cells;
} XlsSheet;

typedef struct {
    int sheet_count;
    XlsSheet* sheets;
} XlsData;

int xls_read_from_buffer(const unsigned char* buf, size_t len, XlsData* out);
void xls_free_data(XlsData* data);

#ifdef __cplusplus
}
#endif
