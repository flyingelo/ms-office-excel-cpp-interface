#include "ExcelInterface.hpp"

#include <filesystem>
#include <iostream>
#include <string>
#include <vector>

int main(int argc, const char *argv[])
{
  try
  {
    bool debug{false};
    for (int i = 1; i < argc; ++i)
    {
      if (std::string(argv[i]) == "--debug")
      {
        debug = true;
      }
    }

    std::filesystem::path spreadsheetName = "test.xlsx";

    const bool keepExcelAlive{true};
    office::excel::MicrosoftExcel excel(keepExcelAlive);
    excel.makeVisible();
    excel.openSpreadsheet(spreadsheetName);
    excel.selectWorksheet("Sheet1");
    excel.setCellValue("C12", "12345");
    excel.save();

    // goal:
    /*
    {
      auto& workbook = excel.openWorkbook(name);
      auto& worksheet = workbook.selectWorksheet("Sheet1");
      auto& cell = worksheet.getCell("C12");
      cell.setValue("1234");
      workbook.save();
    }
    */

  }
  catch (const std::exception &e)
  {
    std::cout << "ERROR: " << e.what() << std::endl;
    return EXIT_FAILURE;
  }

  return EXIT_SUCCESS;
}
