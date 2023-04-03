#include <chrono>
#include <cmath>
#include <filesystem>
#include <iostream>
#include <string>
#include <utility>
#include <vector>

#include "ExcelInterface.hpp"

constexpr std::int64_t intValue{ 42 };
const std::string spreadsheetFileName{ "perf_test.xlsx" };

struct Timer
{
  Timer(): m_startTime(std::chrono::high_resolution_clock::now())
  {
  }

  Timer(std::string name):
    m_name(std::move(name)),
    m_startTime(std::chrono::high_resolution_clock::now())
  {
  }

  ~Timer() {
    const auto endTime = std::chrono::high_resolution_clock::now();
    const std::chrono::duration<double> diff = endTime - m_startTime;
    std::cout << m_name << ": " << diff.count() << " [s]\n";
  }

  Timer(Timer&) = delete;
  Timer& operator=(Timer&) = delete;

  Timer(Timer&&) = delete;
  Timer& operator=(Timer&&) = delete;

private:

  std::string m_name;
  std::chrono::time_point<std::chrono::high_resolution_clock> m_startTime;
};

static void writeSpreadsheet(std::uint32_t rows, std::uint16_t cols) {

  const Timer timer{ "writeSpreadsheet" };

  const std::filesystem::path spreadsheetName(spreadsheetFileName);

  const bool keepExcelAlive{ false };

  office::excel::MicrosoftExcel excel(keepExcelAlive);
  auto& workbook = excel.getWorkbook();
  auto& worksheet = workbook.findWorksheet("Sheet1");
  for (std::uint32_t row = 0; row < rows; ++row) {
    for (std::uint16_t col = 0; col < cols; ++col) {
      worksheet.getCell(row, col).setValue(intValue);
    }
  }

  office::excel::SaveAsArguments saveArgs(spreadsheetName);
  saveArgs.saveConflictResolution =
    office::excel::SaveAsArguments::SaveConflictResolution::OverwriteFile;
  workbook.saveAs(saveArgs);
}

static void readSpreadsheet(std::uint32_t rows, std::uint16_t cols) {

  const Timer timer{ "readSpreadsheet" };

  const std::filesystem::path spreadsheetName(spreadsheetFileName);
  const bool keepExcelAlive{ false };
  office::excel::MicrosoftExcel excel(keepExcelAlive);
  auto& workbook = excel.openWorkbook(spreadsheetName);
  auto& worksheet = workbook.findWorksheet("Sheet1");
  for (std::uint32_t row = 0; row < rows; ++row) {
    for (std::uint16_t col = 0; col < cols; ++col) {
      if (worksheet.getCell(row, col).getValueInt64() != intValue) {
        throw std::runtime_error("Value read from spreadsheet is not correct");
      }
    }
  }
}

int main(int argc, const char* argv[]) {
  try {

    constexpr std::uint32_t defaultMaxRows{100U};
    constexpr std::uint16_t defaultMaxCols{100U};

    std::uint32_t rows{defaultMaxRows};
    std::uint16_t cols{defaultMaxCols};
    const std::vector<std::string> arguments(argv, argv + argc);
    for (std::size_t i = 0; i < arguments.size(); ++i) {
      if (arguments.at(i) == "--rows") {
        rows = static_cast<std::uint32_t>(std::stoi(arguments.at(i + 1)));
      }
      else if (arguments.at(i) == "--cols") {
        cols = static_cast<std::uint16_t>(std::stoi(arguments.at(i + 1)));
      }
    }

    std::cout << "==================================================\n";
    std::cout << "Running Excel performance tests\n";
    std::cout << "Max rows: " << rows << "\n";
    std::cout << "Max cols: " << cols << "\n";

    writeSpreadsheet(rows, cols);
    readSpreadsheet(rows, cols);

    std::cout << "All Excel tests performance finished\n";
    std::cout << "==================================================\n";

  }
  catch (const std::exception& e) {
    std::cout << "ERROR: " << e.what() << std::endl;
    return EXIT_FAILURE;
  }

  return EXIT_SUCCESS;
}
