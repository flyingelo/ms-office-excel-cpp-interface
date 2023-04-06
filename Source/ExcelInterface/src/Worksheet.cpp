#include "../Worksheet.hpp"

#include <ole2.h>

#include <memory>

#include "getCellCoords.hpp"
#include "Ole.hpp"
#include "Utilities.hpp"

namespace office::excel {

  using CellCoords = std::string;
  using CellMap = std::map<std::uintptr_t, std::map<CellCoords, std::unique_ptr<Cell>>>;
  static CellMap cellMap;

  static Cell* getCellFromMap(Worksheet::WorksheetDispatch worksheetDispatch, const CellCoords& cellCoords) {
    auto& worksheetCells = cellMap.at(reinterpret_cast<std::uintptr_t>(worksheetDispatch));
    if (worksheetCells.find(cellCoords) == std::end(worksheetCells)) {
      try {
        auto parm = getArgumentString(to_wstring(cellCoords));
        VARIANT result = getArgumentResult();
        AutoWrap(DISPATCH_PROPERTYGET, &result, worksheetDispatch,
          std::wstring(L"Range").data(), 1, parm.variant);
        worksheetCells.insert(
          std::make_pair(cellCoords, std::make_unique<Cell>(result.pdispVal)));
      }
      catch (const std::runtime_error& e) {
        throw std::runtime_error("MicrosoftExcel::getCell failed. Cell range: " +
          cellCoords + ". " + std::string(e.what()));
      }
    }

    return worksheetCells.at(cellCoords).get();
  }

  Worksheet::Worksheet(WorksheetDispatch worksheetDispatch)
    : m_worksheetDispatch(worksheetDispatch) {
    cellMap.insert(std::make_pair(reinterpret_cast<std::uintptr_t>(m_worksheetDispatch), std::map<std::string, std::unique_ptr<Cell>>()));
  }

  Worksheet::~Worksheet() {
    if (m_worksheetDispatch != nullptr) {
      cellMap.erase(reinterpret_cast<std::uintptr_t>(m_worksheetDispatch));
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

  Cell& Worksheet::getCell(const std::string& cellCoords) {
    return *getCellFromMap(m_worksheetDispatch, cellCoords);
  }

  const Cell& Worksheet::getCell(const std::string& cellCoords) const {
    return *getCellFromMap(m_worksheetDispatch, cellCoords);
  }

  Cell& Worksheet::getCell(std::uint32_t row, std::uint16_t column) {
    return getCell(utils::getCellCoords(row, column));
  }

  const Cell& Worksheet::getCell(std::uint32_t row, std::uint16_t column) const {
    return getCell(utils::getCellCoords(row, column));
  }

}  // namespace office::excel
