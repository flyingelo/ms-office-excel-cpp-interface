
#pragma once

#include <filesystem>
#include <memory>
#include <string>

struct IDispatch;

namespace office::excel {

class Cell
{
public:

  Cell(IDispatch *);

  Cell() = delete;

  // avoid copying this class, unless a specific need arises
  Cell(Cell&) = delete; 
  Cell& operator=(Cell&) = delete;

  Cell(Cell&&) = default;
  Cell& operator=(Cell&&) = default;

  ~Cell();

  void setValue(const std::string& value);
  //std::string getValue() const;

private:  

  IDispatch* m_range;

};

}
