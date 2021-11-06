#include "../Cell.hpp"

#include "Ole.hpp"

#include <ole2.h>

#include <iostream>
#include <memory>
#include <string>
#include <vector>

namespace office::excel {

// TODO: get rid of this
static inline std::wstring to_wstring(const std::string &src)
{
  std::wstring trg(src.size(), L' ');
  std::copy(std::begin(src), std::end(src), std::begin(trg));
  return trg;
}

Cell::Cell(IDispatch* range) :
  m_range(range)
{
}

Cell::~Cell()
{
  m_range->Release();
}

void Cell::setValue(const std::string &value)
{
  try
  {
    auto valueArg = getArgumentString(to_wstring(value));
    AutoWrap(DISPATCH_PROPERTYPUT, NULL, m_range, std::wstring(L"Value").data(), 1, valueArg.variant);
  }
  catch (const std::runtime_error &e)
  {
    throw std::runtime_error("Cell::setValue failed. Value: " + value + ". " + std::string(e.what()));
  }
}

}
