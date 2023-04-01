#include "../Cell.hpp"

#include <ole2.h>

#include <memory>
#include <string>
#include <vector>

#include "Ole.hpp"
#include "Utilities.hpp"

namespace office::excel {

  Cell::Cell(CellDispatch cellDispatch): m_cellDispatch(cellDispatch) {}

  Cell::~Cell() {
    if (m_cellDispatch != nullptr) {
      m_cellDispatch->Release();
    }
  }

  void Cell::setValue(const std::string& value) {
    try {
      auto valueArg = getArgumentString(to_wstring(value));
      AutoWrap(DISPATCH_PROPERTYPUT, nullptr, m_cellDispatch,
        std::wstring(L"Value").data(), 1, valueArg.variant);
    }
    catch (const std::runtime_error& e) {
      throw std::runtime_error("Cell::setValue failed. Value: " + value + ". " +
        std::string(e.what()));
    }
  }

  void Cell::setValue(std::int32_t value) {
    try {
      auto valueArg = getArgumentInt32(value);
      AutoWrap(DISPATCH_PROPERTYPUT, nullptr, m_cellDispatch,
        std::wstring(L"Value").data(), 1, valueArg.variant);
    }
    catch (const std::runtime_error& e) {
      throw std::runtime_error(
        "Cell::setValue failed. Value: " + std::to_string(value) + ". " +
        std::string(e.what()));
    }
  }

  void Cell::setValue(std::int64_t value) {
    try {
      auto valueArg = getArgumentInt64(value);
      AutoWrap(DISPATCH_PROPERTYPUT, nullptr, m_cellDispatch,
        std::wstring(L"Value").data(), 1, valueArg.variant);
    }
    catch (const std::runtime_error& e) {
      throw std::runtime_error(
        "Cell::setValue failed. Value: " + std::to_string(value) + ". " +
        std::string(e.what()));
    }
  }

  void Cell::setValue(double value) {
    try {
      auto valueArg = getArgumentDouble(value);
      AutoWrap(DISPATCH_PROPERTYPUT, nullptr, m_cellDispatch,
        std::wstring(L"Value").data(), 1, valueArg.variant);
    }
    catch (const std::runtime_error& e) {
      throw std::runtime_error(
        "Cell::setValue failed. Value: " + std::to_string(value) + ". " +
        std::string(e.what()));
    }
  }

  std::string Cell::getValue() const {
    try {
      VARIANT result = getArgumentResult();
      AutoWrap(DISPATCH_PROPERTYGET, &result, m_cellDispatch,
        std::wstring(L"Value").data(), 0);
      if (result.vt == VARENUM::VT_BSTR) {
        return to_string(result.bstrVal);
      }
      else if (result.vt == VARENUM::VT_R8) {
        return std::to_string(result.dblVal);
      }
      else if (result.vt == VARENUM::VT_I8) {
        return std::to_string(result.llVal);
      }
      else {
        throw std::runtime_error("Cell::getValue: unsupported type " +
          std::to_string(result.vt));
      }
    }
    catch (const std::runtime_error& e) {
      throw std::runtime_error("Cell::getValue failed. " + std::string(e.what()));
    }
  }

  double Cell::getValueDouble() const {
    try {
      VARIANT result = getArgumentResult();
      AutoWrap(DISPATCH_PROPERTYGET, &result, m_cellDispatch,
        std::wstring(L"Value").data(), 0);
      if (result.vt == VARENUM::VT_R8) {
        return result.dblVal;
      }
      else if (result.vt == VARENUM::VT_I8) {
        return static_cast<double>(result.llVal);
      }
      else {
        throw std::runtime_error("Cell::getValueDouble: unsupported type " +
          std::to_string(result.vt));
      }
    }
    catch (const std::runtime_error& e) {
      throw std::runtime_error("Cell::getValueDouble failed. " +
        std::string(e.what()));
    }
  }

  std::int64_t Cell::getValueInt64() const {
    try {
      VARIANT result = getArgumentResult();
      AutoWrap(DISPATCH_PROPERTYGET, &result, m_cellDispatch,
        std::wstring(L"Value").data(), 0);
      if (result.vt == VARENUM::VT_R8) {
        return static_cast<std::int64_t>(result.dblVal);
      }
      else if (result.vt == VARENUM::VT_I8) {
        return result.llVal;
      }
      else {
        throw std::runtime_error("Cell::getValueInt64: unsupported type " +
          std::to_string(result.vt));
      }
    }
    catch (const std::runtime_error& e) {
      throw std::runtime_error("Cell::getValueInt64 failed. " +
        std::string(e.what()));
    }
  }

}  // namespace office::excel
