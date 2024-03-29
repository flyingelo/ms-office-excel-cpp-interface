
#pragma once

#include <filesystem>
#include <memory>
#include <string>

#include "Worksheet.hpp"

struct IDispatch;

namespace office::excel {

  struct SaveAsArguments {
    enum class SaveConflictResolution { OverwriteFile = 2, UserResolution = 1 };

    std::filesystem::path fileName;
    enum SaveConflictResolution saveConflictResolution {
      SaveConflictResolution::UserResolution
    };

    SaveAsArguments(const std::filesystem::path& fileName)
      : fileName(std::filesystem::absolute(fileName)) {}
    SaveAsArguments(const std::string& fileName)
      : SaveAsArguments(std::filesystem::path(fileName)) {}
    SaveAsArguments(const std::wstring& fileName)
      : SaveAsArguments(std::filesystem::path(fileName)) {}

    SaveAsArguments() = default;
    SaveAsArguments(SaveAsArguments&) = default;
    SaveAsArguments(SaveAsArguments&&) = default;
    SaveAsArguments& operator=(SaveAsArguments&&) = default;
    SaveAsArguments& operator=(const SaveAsArguments&) = default;

    ~SaveAsArguments() = default;
  };

  class Workbook {
  public:
    using WorksheetName = std::string;

    using WorkbookDispatch = IDispatch*;

    Workbook(IDispatch* dispatch);

    Workbook() = delete;

    // avoid copying this class, unless a specific need arises
    Workbook(Workbook&) = delete;
    Workbook& operator=(Workbook&) = delete;

    Workbook(Workbook&&) = default;
    Workbook& operator=(Workbook&&) = default;

    ~Workbook();

    Worksheet& findWorksheet(const WorksheetName& name);
    const Worksheet& findWorksheet(const WorksheetName& name) const;

    void selectWorksheet(const Worksheet& worksheet);

    void save();

    void saveAs(const SaveAsArguments& arguments);

  private:
    using WorksheetsDispatch = IDispatch*;

    std::unique_ptr<Worksheet> m_worksheet;
    WorksheetsDispatch m_worksheetsDispatch{ nullptr };
    WorkbookDispatch m_workbookDispatch;
  };

}  // namespace office::excel
