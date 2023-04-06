![Status](https://github.com/flyingelo/ms-office-excel-cpp-interface/actions/workflows/main.yml/badge.svg)

# ms-office-excel-cpp-interface
C++ APIs for Microsoft Excel

For more information, see:
https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/automate-excel-from-c

This module is a C++ wrapper around some of those APIs.

For example, suppose we want to write a value of `1` into cell `C12` in the spreadsheet file `test.xlsx` in the worksheet `Sheet1`.
This is done in the following code snippet:

```C++
#include "ExcelInterface.hpp"
#include <filesystem>

const std::filesystem::path spreadsheetName("test.xlsx");
office::excel::MicrosoftExcel excel;
auto& workbook = excel.getWorkbook();
auto& worksheet = workbook.findWorksheet("Sheet1");
auto& cell = worksheet.getCell("C12");
cell.setValue(1);
workbook.saveAs(spreadsheetName);
```

Now, suppose we want to read the value of cell `C12` in the spreadsheet file `test.xlsx` in the worksheet `Sheet1`.
This can be done as follows:

```C++
#include "ExcelInterface.hpp"
#include <filesystem>

const std::filesystem::path spreadsheetName{"test.xlsx"};
office::excel::MicrosoftExcel excel;
const auto& workbook = excel.openWorkbook(spreadsheetName);
const auto& worksheet = workbook.findWorksheet("Sheet1");
const auto& cell = worksheet.getCell("C12");
const auto value = cell.getValueInt64();
```

Scons is used to build the project.

Build command:
`scons`

Clean command:
`scons -c`

To run tests:
`Tests\runTests.bat`
