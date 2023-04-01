#include "Ole.hpp"

namespace office::excel {

  static DISPPARAMS getDispParams(WORD autoType, unsigned int cArgs,
    VARIANTARG* rgvarg, unsigned int cNamedArgs,
    DISPID* rgdispidNamedArgs) {
    DISPPARAMS dp;
    dp.cArgs = cArgs;
    dp.rgvarg = rgvarg;
    if (autoType & DISPATCH_PROPERTYPUT) {
      dp.cNamedArgs = cNamedArgs;
      dp.rgdispidNamedArgs = rgdispidNamedArgs;
    }
    else {
      dp.cNamedArgs = 0;
      dp.rgdispidNamedArgs = nullptr;
    }
    return dp;
  }

  void AutoWrap(WORD autoType, VARIANT* pvResult, IDispatch* pDisp,
    LPOLESTR ptName, unsigned int cArgs...) {
    // Begin variable-argument list...
    va_list marker;
    va_start(marker, static_cast<int>(cArgs));

    if (pDisp == nullptr) {
      throw std::runtime_error("NULL IDispatch passed to AutoWrap()");
    }

    // Variables used...
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;

    // Get DISPID for name passed...
    if (FAILED(pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT,
      &dispID))) {
      throw std::runtime_error("IDispatch::GetIDsOfNames failed.");
    }

    // Allocate memory for arguments...
    std::vector<VARIANT> pArgs(cArgs);
    // Extract arguments...
    for (auto& pArg : pArgs) {
      pArg = va_arg(marker, VARIANT);
    }

    pArgs.emplace_back();

    // Build DISPPARAMS
    DISPPARAMS dp = getDispParams(autoType, cArgs, pArgs.data(), 1, &dispidNamed);

    // Make the call!
    if (FAILED(pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType,
      &dp, pvResult, NULL, NULL))) {
      throw std::runtime_error("IDispatch::Invoke failed.");
    }

    // End variable-argument section...
    va_end(marker);
  }

  VariantContainer getEmptyArgument() {
    // make empty argument:
    // https://support.microsoft.com/en-us/topic/office-automation-using-visual-c-67da40c2-7671-f700-474d-36ac522d76f2
    VariantContainer varOpt;
    varOpt.variant.vt = VT_ERROR;
    varOpt.variant.scode = DISP_E_PARAMNOTFOUND;
    return varOpt;
  }

  VariantContainer getArgumentInt32(std::int32_t value) {
    VariantContainer var;
    var.variant.vt = VT_I4;
    var.variant.lVal = value;
    return var;
  }

  VariantContainer getArgumentInt64(std::int64_t value) {
    VariantContainer var;
    var.variant.vt = VT_I8;
    var.variant.llVal = value;
    return var;
  }

  VariantContainer getArgumentDouble(double value) {
    VariantContainer var;
    var.variant.vt = VT_R8;
    var.variant.dblVal = value;
    return var;
  }

  VariantContainer getArgumentString(const std::wstring& value) {
    VariantContainer container;
    container.variant.vt = VT_BSTR;
    container.variant.bstrVal = ::SysAllocString(value.c_str());
    return container;
  }

  VariantContainer getArgumentBool(bool value) {
    // set read-only mode
    VariantContainer var;
    var.variant.vt = VT_BOOL;
    var.variant.bVal = static_cast<BYTE>(value);
    return var;
  }

  VARIANT getArgumentResult() {
    VARIANT result;
    VariantInit(&result);
    return result;
  }

}  // namespace office::excel
