#include "../ExcelInterface.hpp"

#include <ole2.h>

#include <iostream>
#include <memory>
#include <string>
#include <vector>

#include "Ole.hpp"

namespace office::excel {

inline std::string to_string(const std::wstring &src) {
  std::string trg(src.size(), ' ');
  for (std::size_t i = 0; i < src.size(); ++i) {
    if (static_cast<int>(src[i]) < 255) {
      trg[i] = static_cast<char>(src[i]);
    } else {
      trg[i] = '#';
    }
  }

  return trg;
}

struct CLSIDContainer {
  CLSIDContainer() {
    if (FAILED(CoInitialize(nullptr))) {
      throw std::runtime_error(
          "Failed to initialize because CoInitialize() function call failed.");
    }
    if (FAILED(CLSIDFromProgID(L"Excel.Application", &m_clsid))) {
      throw std::runtime_error(
          "Failed to initialize because CLSIDFromProgID() function call "
          "failed.");
    }
  }

  ~CLSIDContainer() { CoUninitialize(); }

  // avoid moving or copying this struct, unless a specific need arises
  CLSIDContainer(CLSIDContainer &) = delete;
  CLSIDContainer(CLSIDContainer &&) = delete;
  CLSIDContainer &operator=(CLSIDContainer &&) = delete;
  CLSIDContainer &operator=(CLSIDContainer &) = delete;

  const CLSID &getCLSID() const { return m_clsid; }

 private:
  CLSID m_clsid;
};

MicrosoftExcel::MicrosoftExcel(bool keepAlive)
    : m_keepAlive(keepAlive),
      m_clsidContainer(std::make_unique<CLSIDContainer>()) {
  if (FAILED(CoCreateInstance(m_clsidContainer->getCLSID(), NULL,
                              CLSCTX_LOCAL_SERVER, IID_IDispatch,
                              (void **)&m_excelApp))) {
    throw std::runtime_error("Excel not registered properly");
  }

  // Get Workbooks collection
  {
    VARIANT result = getArgumentResult();
    AutoWrap(DISPATCH_PROPERTYGET, &result, m_excelApp,
             std::wstring(L"Workbooks").data(), 0);
    m_workbooks = result.pdispVal;
  }
}

MicrosoftExcel::~MicrosoftExcel() {
  try {
    if (!m_keepAlive) {
      // Tell Excel to quit (i.e. App.Quit)
      AutoWrap(DISPATCH_METHOD, NULL, m_excelApp, std::wstring(L"Quit").data(),
               0);
    }

    if (m_workbooks != nullptr) m_workbooks->Release();
    if (m_excelApp != nullptr) m_excelApp->Release();
    // do not delete m_dispatch - someone else appears to own it :(
    // for the same reason, m_dispatch cannot be a unique_ptr.
  } catch (const std::exception &e) {
    std::cout << "~MicrosoftExcel exception: " << e.what() << '\n';
  }
}

void MicrosoftExcel::makeVisible() {
  try {
    // Make it visible (i.e. app.visible = 1)
    auto x = getArgumentInt32(1);
    AutoWrap(DISPATCH_PROPERTYPUT, NULL, m_excelApp,
             std::wstring(L"Visible").data(), 1, x.variant);
  } catch (const std::exception &e) {
    throw std::runtime_error("MicrosoftExcel::makeVisible failed. " +
                             std::string(e.what()));
  }
}

Workbook &MicrosoftExcel::openWorkbook(
    const OpenSpreadsheetArguments &arguments) {
  try {
    // Call Workbooks.Open() to open workbook...
    VARIANT result = getArgumentResult();
    auto fileNameArg = getArgumentString(arguments.fileName.wstring());
    auto updateLinksArg = getEmptyArgument();
    auto formatArg = getEmptyArgument();
    auto readOnlyArg = getArgumentBool(arguments.readOnly);
    auto passwordArg = arguments.password.empty()
                           ? getEmptyArgument()
                           : getArgumentString(arguments.password);
    AutoWrap(DISPATCH_PROPERTYGET, &result, m_workbooks,
             std::wstring(L"Open").data(), 5, passwordArg.variant,
             formatArg.variant, readOnlyArg.variant, updateLinksArg.variant,
             fileNameArg.variant);
    m_workbook = std::make_unique<Workbook>(result.pdispVal);
    return *m_workbook.get();
  } catch (const std::exception &e) {
    throw std::runtime_error("MicrosoftExcel::openWorkbook failed. Workbook: " +
                             to_string(arguments.fileName) + ". " +
                             std::string(e.what()));
  }
}

Workbook &MicrosoftExcel::getWorkbook() {
  try {
    // Call Workbooks.Add() to open workbook...
    VARIANT result = getArgumentResult();
    AutoWrap(DISPATCH_PROPERTYGET, &result, m_workbooks,
             std::wstring(L"Add").data(), 0);
    m_workbook = std::make_unique<Workbook>(result.pdispVal);
    return *m_workbook.get();
  } catch (const std::exception &e) {
    throw std::runtime_error("MicrosoftExcel::getWorkbook failed. " +
                             std::string(e.what()));
  }
}

}  // namespace office::excel
