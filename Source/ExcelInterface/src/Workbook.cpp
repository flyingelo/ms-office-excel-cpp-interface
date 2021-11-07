#include "../Workbook.hpp"

#include <ole2.h>

#include <memory>

#include "Ole.hpp"

namespace office::excel {

inline std::wstring to_wstring(const std::string& src) {
  std::wstring trg(src.size(), L' ');
  std::copy(std::begin(src), std::end(src), std::begin(trg));
  return trg;
}

Workbook::Workbook(IDispatch* workbookDispatch)
    : m_workbookDispatch(workbookDispatch) {
  // Get Worksheets collection
  {
    VARIANT result = getArgumentResult();
    AutoWrap(DISPATCH_PROPERTYGET, &result, m_workbookDispatch,
             std::wstring(L"Worksheets").data(), 0);
    m_worksheetsDispatch = result.pdispVal;
  }
}

Workbook::~Workbook() {
  if (m_worksheetsDispatch != nullptr) {
    m_worksheetsDispatch->Release();
  }

  if (m_workbookDispatch != nullptr) {
    m_workbookDispatch->Release();
  }
}

void Workbook::save() {
  try {
    VARIANT result = getArgumentResult();
    AutoWrap(DISPATCH_PROPERTYGET, &result, m_workbookDispatch,
             std::wstring(L"Save").data(), 0);
  } catch (const std::exception& e) {
    throw std::runtime_error("Workbook::save failed. " + std::string(e.what()));
  }
}

void Workbook::saveAs(const SaveAsArguments& arguments) {
  try {
    auto fileNameArg = getArgumentString(arguments.fileName.wstring());

    if (arguments.saveConflictResolution ==
        SaveAsArguments::SaveConflictResolution::OverwriteFile) {
      // cannot get conflict resolution argument to work properly,
      // so to avoid user dialog to overwrite the existing file or not,
      // just delete the file beforehand, if it already exists or bail.
      if (std::filesystem::exists(arguments.fileName)) {
        std::filesystem::remove(arguments.fileName);
      }
    }

    VARIANT result = getArgumentResult();
    AutoWrap(DISPATCH_PROPERTYGET, &result, m_workbookDispatch,
             std::wstring(L"SaveAs").data(), 1, fileNameArg.variant);
  } catch (const std::exception& e) {
    throw std::runtime_error("Workbook::save failed. " + std::string(e.what()));
  }
}

Worksheet& Workbook::findWorksheet(const std::string& worksheetName) {
  try {
    // Find the desired worksheet. TODO: do not SELECT it

    auto sheetNameArg = getArgumentString(to_wstring(worksheetName));
    VARIANT result = getArgumentResult();
    AutoWrap(DISPATCH_PROPERTYGET, &result, m_worksheetsDispatch,
             std::wstring(L"Item").data(), 1, sheetNameArg.variant);

    auto isReplace = getArgumentBool(true);
    AutoWrap(DISPATCH_PROPERTYGET, nullptr, result.pdispVal,
             std::wstring(L"Select").data(), 1, isReplace.variant);

    m_worksheet = std::make_unique<Worksheet>(result.pdispVal);
    return *m_worksheet.get();

  } catch (const std::exception& e) {
    throw std::runtime_error(
        "MicrosoftExcel::selectWorksheet failed. Worksheet: " + worksheetName +
        ". " + std::string(e.what()));
  }
}

}  // namespace office::excel
