#include <cmath>
#include <filesystem>
#include <iostream>
#include <string>
#include <vector>

#include "ExcelInterface.hpp"

constexpr int intValue{ 42 };
constexpr double doubleValue{ 55.6 };

static void createTestSpreadsheet() {
  const std::filesystem::path spreadsheetName("test.xlsx");

  const bool keepExcelAlive{ false };

  std::cout << "Start Excel... ";
  office::excel::MicrosoftExcel excel(keepExcelAlive);
  std::cout << "ok\n";

  std::cout << "Get new workbook... ";
  auto& workbook = excel.getWorkbook();
  std::cout << "ok\n";

  std::cout << "Get worksheet... ";
  auto& worksheet = workbook.findWorksheet("Sheet1");
  std::cout << "ok\n";

  std::cout << "Get cell... ";
  auto& cellC12 = worksheet.getCell("C12");
  std::cout << "ok\n";

  std::cout << "Set cell value... ";
  cellC12.setValue("X");
  std::cout << "ok\n";

  worksheet.getCell("D12").setValue(intValue);
  worksheet.getCell("E12").setValue(doubleValue);

  std::cout << "Save workbook... ";
  office::excel::SaveAsArguments saveArgs(spreadsheetName);
  saveArgs.saveConflictResolution =
    office::excel::SaveAsArguments::SaveConflictResolution::OverwriteFile;
  workbook.saveAs(saveArgs);
  std::cout << "ok\n";
}

static void checkSpreadsheetValues() {
  const std::filesystem::path spreadsheetName("test.xlsx");
  const bool keepExcelAlive{ false };
  office::excel::MicrosoftExcel excel(keepExcelAlive);
  auto& workbook = excel.openWorkbook(spreadsheetName);
  auto& worksheet = workbook.findWorksheet("Sheet1");
  const auto c12value = worksheet.getCell("C12").getValue();
  if (c12value != "X") {
    throw std::runtime_error("Unexpected value at cell C12. Expected X, got " +
      c12value);
  }
  const auto d12value = worksheet.getCell("D12").getValueInt64();
  if (d12value != intValue) {
    throw std::runtime_error("Unexpected value at cell D12. Expected 42, got " +
      std::to_string(d12value));
  }
  const auto e12value = worksheet.getCell("E12").getValueDouble();
  constexpr double tolerance{ 1E-12 };
  if (std::abs(e12value - doubleValue) > tolerance) {
    throw std::runtime_error(
      "Unexpected value at cell E12. Expected 55.6, got " +
      std::to_string(e12value));
  }
}

static void makeExcelVisibleTest() {
  // not sure how to actually check that it pops up,
  // but assume that if it doesn't crash, it's ok

  const bool keepExcelAlive{ false };
  office::excel::MicrosoftExcel excel(keepExcelAlive);

  std::cout << "Making Excel visible... ";
  excel.makeVisible();
  std::cout << "ok\n";
}

int main(int argc, const char* argv[]) {
  try {

    const std::vector<std::string> arguments(argv, argv + argc);

    bool debug{ false };

    for (const auto& argument : arguments) {
      if (argument == "--debug") {
        debug = true;
      }
    }

    if (!debug) {
      createTestSpreadsheet();
      checkSpreadsheetValues();
      makeExcelVisibleTest();
    }
    else {
      std::cout << "Debug mode, not running tests\n";
    }

  }
  catch (const std::exception& e) {
    std::cout << "ERROR: " << e.what() << std::endl;
    return EXIT_FAILURE;
  }

  return EXIT_SUCCESS;
}
