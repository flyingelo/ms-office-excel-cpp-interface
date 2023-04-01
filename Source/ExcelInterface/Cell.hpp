
#pragma once

#include <filesystem>
#include <memory>
#include <string>

struct IDispatch;

namespace office::excel {

  class Cell {
  public:
    using CellDispatch = IDispatch*;

    Cell(CellDispatch dispatch);

    Cell() = delete;

    // avoid copying this class, unless a specific need arises
    Cell(Cell&) = delete;
    Cell& operator=(Cell&) = delete;

    Cell(Cell&&) = default;
    Cell& operator=(Cell&&) = default;

    ~Cell();

    void setValue(std::int32_t value);
    void setValue(std::int64_t value);
    void setValue(double value);
    void setValue(const std::string& value);

    [[nodiscard]] std::string getValue() const;

    [[nodiscard]] double getValueDouble() const;

    [[nodiscard]] std::int64_t getValueInt64() const;

  private:
    CellDispatch m_cellDispatch{ nullptr };
  };

}  // namespace office::excel
