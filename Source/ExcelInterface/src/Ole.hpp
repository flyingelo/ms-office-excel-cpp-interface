
#pragma once

#include <ole2.h>

#include <iostream>
#include <memory>
#include <string>
#include <vector>

#include "../ExcelInterface.hpp"

namespace office::excel {

struct VariantContainer {
  VariantContainer() = default;
  VariantContainer(VariantContainer &) = delete;
  VariantContainer(VariantContainer &&) = default;
  VariantContainer &operator=(VariantContainer &&) = default;
  VariantContainer &operator=(VariantContainer &) = delete;

  ~VariantContainer() { VariantClear(&variant); }

  VARIANT variant;
};

void AutoWrap(WORD autoType, VARIANT *pvResult, IDispatch *pDisp,
              LPOLESTR ptName, unsigned int cArgs...);

VariantContainer getEmptyArgument();

VariantContainer getArgumentInt32(std::int32_t value);

VariantContainer getArgumentInt64(std::int64_t value);

VariantContainer getArgumentDouble(double value);

VariantContainer getArgumentString(const std::wstring &value);

VariantContainer getArgumentBool(bool value);

VARIANT getArgumentResult();

}  // namespace office::excel
