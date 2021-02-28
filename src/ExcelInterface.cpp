#include "../ExcelInterface.hpp"

#include <ole2.h>

#include <iostream>
#include <memory>
#include <string>
#include <vector>

namespace office::excel {


// This probably requires some work.

inline std::wstring to_wstring(const std::string &src)
{
  std::wstring trg(src.size(), L' ');
  std::copy(std::begin(src), std::end(src), std::begin(trg));
  return trg;
}

inline std::string to_string(const std::wstring &src)
{
  std::string trg(src.size(), ' ');
  for (std::size_t i = 0; i < src.size(); ++i) {
    if (static_cast<int>(src[i]) < 255) {
      trg[i] = static_cast<char>(src[i]);
    }
    else {
      trg[i] = '#';
    }
  }

  return trg;
}

static DISPPARAMS getDispParams(
    WORD autoType,
    unsigned int cArgs,
    VARIANTARG *rgvarg,
    unsigned int cNamedArgs,
    DISPID *rgdispidNamedArgs)
{
  DISPPARAMS dp;
  dp.cArgs = cArgs;
  dp.rgvarg = rgvarg;
  if (autoType & DISPATCH_PROPERTYPUT)
  {
    dp.cNamedArgs = cNamedArgs;
    dp.rgdispidNamedArgs = rgdispidNamedArgs;
  }
  else
  {
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = nullptr;
  }
  return dp;
}

// AutoWrap() - Automation helper function...
static void AutoWrap(WORD autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, unsigned int cArgs...)
{
  // Begin variable-argument list...
  va_list marker;
  va_start(marker, static_cast<int>(cArgs));

  if (pDisp == nullptr)
  {
    throw std::runtime_error("NULL IDispatch passed to AutoWrap()");
  }

  // Variables used...
  DISPID dispidNamed = DISPID_PROPERTYPUT;
  DISPID dispID;

  // Get DISPID for name passed...
  if (FAILED(pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID)))
  {
    throw std::runtime_error("IDispatch::GetIDsOfNames failed.");
  }

  // Allocate memory for arguments...
  std::vector<VARIANT> pArgs(cArgs);
  // Extract arguments...
  for (auto &pArg : pArgs)
  {
    pArg = va_arg(marker, VARIANT);
  }

  pArgs.emplace_back();

  // Build DISPPARAMS
  DISPPARAMS dp = getDispParams(autoType, cArgs, pArgs.data(), 1, &dispidNamed);

  // Make the call!
  if (FAILED(pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL)))
  {
    throw std::runtime_error("IDispatch::Invoke failed.");
  }

  // End variable-argument section...
  va_end(marker);
}

struct CLSIDContainer
{

  CLSIDContainer()
  {
    if (FAILED(CoInitialize(nullptr)))
    {
      throw std::runtime_error("Failed to initialize because CoInitialize() function call failed.");
    }
    if (FAILED(CLSIDFromProgID(L"Excel.Application", &m_clsid)))
    {
      throw std::runtime_error("Failed to initialize because CLSIDFromProgID() function call failed.");
    }
  }

  ~CLSIDContainer()
  {
    CoUninitialize();
  }

  // avoid moving or copying this struct, unless a specific need arises
  CLSIDContainer(CLSIDContainer &) = delete;
  CLSIDContainer(CLSIDContainer &&) = delete;
  CLSIDContainer &operator=(CLSIDContainer &&) = delete;
  CLSIDContainer &operator=(CLSIDContainer &) = delete;

  const CLSID &getCLSID() const
  {
    return m_clsid;
  }

private:
  CLSID m_clsid;
};

struct VariantContainer
{
  VariantContainer() = default;
  VariantContainer(VariantContainer &) = delete;
  VariantContainer(VariantContainer &&) = default;
  VariantContainer &operator=(VariantContainer &&) = default;
  VariantContainer &operator=(VariantContainer &) = delete;

  ~VariantContainer()
  {
    VariantClear(&variant);
  }

  VARIANT variant;
};

static VariantContainer getEmptyArgument()
{
  // make empty argument:
  // https://support.microsoft.com/en-us/topic/office-automation-using-visual-c-67da40c2-7671-f700-474d-36ac522d76f2
  VariantContainer varOpt;
  varOpt.variant.vt = VT_ERROR;
  varOpt.variant.scode = DISP_E_PARAMNOTFOUND;
  return varOpt;
}

static VariantContainer getArgumentInt32(int value)
{
  VariantContainer var;
  var.variant.vt = VT_I4;
  var.variant.lVal = value;
  return var;
}

static VariantContainer getArgumentString(const std::wstring &value)
{
  VariantContainer container;
  container.variant.vt = VT_BSTR;
  container.variant.bstrVal = ::SysAllocString(value.c_str());
  return container;
}

static VariantContainer getArgumentBool(bool value)
{
  // set read-only mode
  VariantContainer var;
  var.variant.vt = VT_BOOL;
  var.variant.bVal = static_cast<BYTE>(value);
  return var;
}

static VARIANT getArgumentResult()
{
  VARIANT result;
  VariantInit(&result);
  return result;
}

// ===================

MicrosoftExcel::MicrosoftExcel(bool keepAlive) : m_keepAlive(keepAlive),
                                                 m_clsidContainer(std::make_unique<CLSIDContainer>())
{
  if (FAILED(CoCreateInstance(m_clsidContainer->getCLSID(), NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&m_excelApp)))
  {
    throw std::runtime_error("Excel not registered properly");
  }
}

MicrosoftExcel::~MicrosoftExcel()
{
  try
  {
    if (!m_keepAlive)
    {
      // Tell Excel to quit (i.e. App.Quit)
      AutoWrap(DISPATCH_METHOD, NULL, m_excelApp, L"Quit", 0);
    }

    if (m_worksheet != nullptr)
      m_worksheet->Release();
    if (m_worksheets != nullptr)
      m_worksheets->Release();
    if (m_workbook != nullptr)
      m_workbook->Release();
    if (m_workbooks != nullptr)
      m_workbooks->Release();
    if (m_excelApp != nullptr)
      m_excelApp->Release();
    // do not delete m_dispatch - someone else appears to own it :(
    // for the same reason, m_dispatch cannot be a unique_ptr.
  }
  catch (...)
  {
  }
}

void MicrosoftExcel::makeVisible()
{
  try
  {
    // Make it visible (i.e. app.visible = 1)
    auto x = getArgumentInt32(1);
    AutoWrap(DISPATCH_PROPERTYPUT, NULL, m_excelApp, L"Visible", 1, x.variant);
  }
  catch (const std::exception &e)
  {
    throw std::runtime_error("MicrosoftExcel::makeVisible failed. " + std::string(e.what()));
  }
}

void MicrosoftExcel::openSpreadsheet(const OpenSpreadsheetArguments& arguments)
{
  try
  {
    // Get Workbooks collection
    {
      VARIANT result = getArgumentResult();
      AutoWrap(DISPATCH_PROPERTYGET, &result, m_excelApp, L"Workbooks", 0);
      m_workbooks = result.pdispVal;
    }

    // Call Workbooks.Open() to open workbook...
    {
      VARIANT result = getArgumentResult();

      auto fileNameArg = getArgumentString(arguments.fileName.wstring());
      auto updateLinksArg = arguments.updateLinks == OpenSpreadsheetArguments::UpdateLinks::Default ? getEmptyArgument() : getArgumentInt32(static_cast<int>(arguments.updateLinks));
      auto formatArg = arguments.format == OpenSpreadsheetArguments::Format::Default ? getEmptyArgument() : getArgumentInt32(static_cast<int>(arguments.format));
      auto readOnlyArg = getArgumentBool(arguments.readOnly);
      auto passwordArg = arguments.password.empty() ? getEmptyArgument() : getArgumentString(arguments.password);
      AutoWrap(DISPATCH_PROPERTYGET, &result, m_workbooks, L"Open", 5, passwordArg.variant, formatArg.variant, readOnlyArg.variant, updateLinksArg.variant, fileNameArg.variant);  

      m_workbook = result.pdispVal;
    }
  }
  catch (const std::exception &e)
  {
    throw std::runtime_error("MicrosoftExcel::openSpreadsheet failed. Spreadsheet: " + to_string(arguments.fileName) + ". " + std::string(e.what()));
  }
}

void MicrosoftExcel::save()
{
  try
  {
    VARIANT result = getArgumentResult();
    AutoWrap(DISPATCH_PROPERTYGET, &result, m_workbook, L"Save", 0);
  }
  catch (const std::exception &e)
  {
    throw std::runtime_error("MicrosoftExcel::save failed. " + std::string(e.what()));
  }
}

void MicrosoftExcel::selectWorksheet(const std::string &worksheetName)
{
  try
  {
    std::wstring wworksheetName = to_wstring(worksheetName);

    // Get Worksheets collection
    {
      VARIANT result = getArgumentResult();
      AutoWrap(DISPATCH_PROPERTYGET, &result, m_workbook, L"Worksheets", 0);
      m_worksheets = result.pdispVal;
    }

    // Select the desired worksheet
    {
      auto sheetNameArg = getArgumentString(wworksheetName);
      VARIANT result = getArgumentResult();
      AutoWrap(DISPATCH_PROPERTYGET, &result, m_worksheets, L"Item", 1, sheetNameArg.variant);
      m_worksheet = result.pdispVal;

      auto isReplace = getArgumentBool(true);
      AutoWrap(DISPATCH_PROPERTYGET, nullptr, m_worksheet, L"Select", 1, isReplace.variant);
    }
  }
  catch (const std::exception &e)
  {
    throw std::runtime_error("MicrosoftExcel::selectWorksheet failed. Worksheet: " + worksheetName + ". " + std::string(e.what()));
  }
}

void MicrosoftExcel::setCellValue(const std::string &cellRange, const std::string &value)
{
  try
  {
    // Get Range object for the Range A1:O15...
    using Range = IDispatch *;

    std::wstring wcellRange = to_wstring(cellRange);
    std::wstring wvalue = to_wstring(value);

    Range range;
    {
      auto parm = getArgumentString(wcellRange);
      VARIANT result = getArgumentResult();
      AutoWrap(DISPATCH_PROPERTYGET, &result, m_worksheet, L"Range", 1, parm.variant);
      range = result.pdispVal;
    }

    auto valueArg = getArgumentString(wvalue);
    AutoWrap(DISPATCH_PROPERTYPUT, NULL, range, L"Value", 1, valueArg.variant);

    range->Release();
  }
  catch (const std::runtime_error &e)
  {
    throw std::runtime_error("MicrosoftExcel::setCellValue failed. Cell range: " + cellRange + ". Value: " + value + ". " + std::string(e.what()));
  }
}

}
