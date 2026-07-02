#pragma once
#include <string>
#include <vector>

struct CellValue {
    enum Type { Empty, String, Number, Boolean, Error };
    Type type{Empty};
    std::string text;
    double number{0};
    bool boolean{false};
};

struct SheetData {
    std::string name;
    std::vector<std::vector<CellValue>> rows;
};

class XlsxReader {
public:
    XlsxReader() = default;
    ~XlsxReader();

    bool open(const std::wstring& filePath);
    void close();
    bool isOpen() const { return fileOpen_; }

    int sheetCount() const { return static_cast<int>(sheets_.size()); }
    std::string sheetName(int index) const;
    const SheetData* sheet(int index) const;
    std::string lastError() const { return lastError_; }

private:
    bool fileOpen_{false};
    std::wstring filePath_;
    std::vector<SheetData> sheets_;
    std::string lastError_;

    bool readXlsx(const std::wstring& path);
    bool readXls(const std::wstring& path);
};
