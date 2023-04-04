![Status](https://github.com/flyingelo/ms-office-excel-cpp-interface/actions/workflows/main.yml/badge.svg)

# ms-office-excel-cpp-interface
C++ APIs for Microsoft Excel

For more information, see:
https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/automate-excel-from-c

This module is a C++ wrapper around those APIs. For example you can get a value of
the cell `C12` in the spreadsheet file `test.xlsx` in the worksheet `Sheet1` as follows:

```C++
#include "ExcelInterface.hpp"

office::excel::MicrosoftExcel excel;
const std::string fileName{"test.xlsx"};
auto& workbook = excel.openWorkbook(fileName);
const auto& worksheet = workbook.findWorksheet("Sheet1");
const auto& cell = worksheet.getCell("C12");
const auto value = cell.getValue();
```

Build command:
`scons`
