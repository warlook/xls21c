#pragma once
#include "addin_base.h"
#include "xlsx_reader.h"
#include "version.h"

#define ADDIN_NAME u"xls21c"

class Xls21cAddin : public V8Addin<Xls21cAddin> {
protected:
    XlsxReader reader_;

public:
    Xls21cAddin();
    virtual ~Xls21cAddin();

    REGISTER_EXTENSION(ADDIN_NAME);
    
    EXT_PROP_RO(Version, u"Версия") {
        lstringu<40> v{ssa{P_VERSION}};
        value.vt = VTYPE_PWSTR;
        value.pwstrVal = copyText(v);
        value.wstrLen = (int)v.length();
        return true;
    }

    EXT_PROP_RW(ThrowErrorDescription, u"БросатьОписаниеОшибки") {
        value.vt = VTYPE_BOOL;
        value.bVal = throwErrors_;
        return true;
    }
    EXT_PROP_WRITE(ThrowErrorDescription) {
        if (value.vt == VTYPE_BOOL) {
            throwErrors_ = value.bVal;
            return true;
        }
        return error(u"Ожидается булево значение", u"Boolean value needed");
    }

    EXT_PROP_RO(ErrorDescription, u"ОписаниеОшибки") {
        value.vt = VTYPE_PWSTR;
        value.pwstrVal = copyText(lastError_);
        value.wstrLen = (int)lastError_.length();
        return true;
    }

    EXT_PROP_RO(IsFileOpen, u"ФайлОткрыт") {
        value.vt = VTYPE_BOOL;
        value.bVal = reader_.isOpen();
        return true;
    }

    EXT_PROP_RO(SheetCount, u"КоличествоЛистов") {
        value.vt = VTYPE_I4;
        value.intVal = reader_.sheetCount();
        return true;
    }

    EXT_PROC(OpenFile, u"ОткрытьФайл", 1);
    EXT_PROC(CloseFile, u"ЗакрытьФайл", 0);
    EXT_FUNC(GetSheetName, u"ПолучитьИмяЛиста", 1);
    EXT_FUNC_WITH_DEFVAL(ReadSheet, u"ПрочитатьЛист", 3);
    EXT_DEFVAL_FOR(ReadSheet) {
        value->vt = VTYPE_PWSTR;
        value->pwstrVal = nullptr;
        value->wstrLen = 0;
        return true;
    }

    END_EXTENSION();
};
