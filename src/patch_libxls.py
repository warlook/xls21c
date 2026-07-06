import sys, os

def patch_file(path, old, new):
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    if old not in content:
        print(f"WARNING: pattern not found in {path}: {old[:60]}...")
        return False
    content = content.replace(old, new, 1)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Patched: {path}")
    return True

src = sys.argv[1]

# 1. lcell fix in xls_addRow
patch_file(os.path.join(src, 'xls.c'),
    'tmp->lcell=row->lcell;',
    'tmp->lcell=row->lcell > pWS->rows.lastcol ? row->lcell : pWS->rows.lastcol;')

# 2. preparse: ignore cell-too-small errors
patch_file(os.path.join(src, 'xls.c'),
    """        case XLS_RECORD_BOOLERR:
            if (xls_isCellTooSmall(pWS->workbook, &tmp, buf)) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
            if (pWS->rows.lastcol""",
    """        case XLS_RECORD_BOOLERR:
            xls_isCellTooSmall(pWS->workbook, &tmp, buf);
            if (pWS->rows.lastcol""")

# 3. xls_addCell: skip size check for LABELSST and LABEL, only reject tiny LABEL records
patch_file(os.path.join(src, 'xls.c'),
    """    if (xls_isCellTooSmall(pWS->workbook, bof, buf))
        return NULL;

\t// printf("ROW: %u COL: %u""",
    """    if (bof->id != XLS_RECORD_LABELSST && bof->id != XLS_RECORD_LABEL && xls_isCellTooSmall(pWS->workbook, bof, buf))
        return NULL;
    if ((bof->id == XLS_RECORD_LABEL || bof->id == XLS_RECORD_LABELSST) && bof->size < 9)
        return NULL;

\t// printf("ROW: %u COL: %u""")

# 4. xls_addCell: handle out-of-bounds column
patch_file(os.path.join(src, 'xls.c'),
    """    if (col >= row->cells.count) {
        if (xls_debug) fprintf(stderr, "Error: Column index out of bounds\\n");
        return NULL;
    }""",
    """    if (col >= row->cells.count) {
        static struct st_cell_data empty_cell = {0};
        return &empty_cell;
    }""")

# 5. main loop: don't abort on addCell failure
patch_file(os.path.join(src, 'xls.c'),
    """            if ((cell = xls_addCell(pWS, &tmp, buf)) == NULL) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }""",
    """            xls_addCell(pWS, &tmp, buf);""")

# 6. xlstool.c: use codepage for all versions
patch_file(os.path.join(src, 'xlstool.c'),
    'const char *from_encoding = pWB->is5ver ? encoding_for_codepage(pWB->codepage) : "ISO-8859-1";',
    'const char *from_encoding = encoding_for_codepage(pWB->codepage);')

# 6b. xlstool.c: keep latin1 shortcut for all BIFF8 (fix_cp1251_encoding handles encoding in wrapper)
# No patch needed - latin1 conversion + fix_cp1251_encoding works correctly

# 7. xlstool.c: LABELSST always uses 4-byte SST index
patch_file(os.path.join(src, 'xlstool.c'),
    """    case XLS_RECORD_LABELSST:
        offset = label[0] + (label[1] << 8);
        if(!pWB->is5ver) {
            offset += ((DWORD)label[2] << 16);
            offset += ((DWORD)label[3] << 24);
        }""",
    """    case XLS_RECORD_LABELSST:
        offset = label[0] + (label[1] << 8) + ((DWORD)label[2] << 16) + ((DWORD)label[3] << 24);""")
