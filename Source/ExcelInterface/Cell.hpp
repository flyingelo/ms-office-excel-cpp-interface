
#pragma once

#include <filesystem>
#include <memory>
#include <string>

struct IDispatch;

namespace office::excel {

class Cell {
 public:
  using CellDispatch = IDispatch*;

  Cell(CellDispatch);

  Cell() = delete;

  // avoid copying this class, unless a specific need arises
  Cell(Cell&) = delete;
  Cell& operator=(Cell&) = delete;

  Cell(Cell&&) = default;
  Cell& operator=(Cell&&) = default;

  ~Cell();

  void setValue(std::int32_t);
  void setValue(std::int64_t);
  void setValue(double);
  void setValue(const std::string&);

  std::string getValue() const;

  double getValueDouble() const;

  std::int64_t getValueInt64() const;

 private:
  CellDispatch m_cellDispatch{nullptr};
};

}  // namespace office::excel
