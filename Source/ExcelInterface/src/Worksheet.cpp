#include "../Worksheet.hpp"

#include <ole2.h>

#include <memory>

#include "Ole.hpp"
#include "Utilities.hpp"

namespace office::excel {

  Worksheet::Worksheet(WorksheetDispatch worksheetDispatch)
    : m_worksheetDispatch(worksheetDispatch) {}

  Worksheet::~Worksheet() {
    if (m_worksheetDispatch != nullptr) {
      m_worksheetDispatch->Release();
    }
  }

  Worksheet::WorksheetDispatch Worksheet::getDispatch() const noexcept {
    return m_worksheetDispatch;
  }

  Worksheet::WorksheetName Worksheet::getName() const {
    try {
      VARIANT result = getArgumentResult();
      AutoWrap(DISPATCH_PROPERTYGET, &result, m_worksheetDispatch,
        std::wstring(L"Name").data(), 0);
      return to_string(result.bstrVal);
    }
    catch (const std::runtime_error& e) {
      throw std::runtime_error("Worksheet::getName failed. " +
        std::string(e.what()));
    }
  }

  Cell& Worksheet::getCell(const std::string& cellRange) {
    if (m_cells.find(cellRange) == std::end(m_cells)) {
      try {
        auto parm = getArgumentString(to_wstring(cellRange));
        VARIANT result = getArgumentResult();
        AutoWrap(DISPATCH_PROPERTYGET, &result, m_worksheetDispatch,
          std::wstring(L"Range").data(), 1, parm.variant);
        m_cells.insert(
          std::make_pair(cellRange, std::make_unique<Cell>(result.pdispVal)));
      }
      catch (const std::runtime_error& e) {
        throw std::runtime_error("MicrosoftExcel::getCell failed. Cell range: " +
          cellRange + ". " + std::string(e.what()));
      }
    }

    return *m_cells.at(cellRange).get();
  }

}  // namespace office::excel
