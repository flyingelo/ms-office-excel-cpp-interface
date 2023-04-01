
#pragma once

#include <map>
#include <memory>
#include <string>
#include <vector>

#include "Cell.hpp"

struct IDispatch;

namespace office::excel {

  class Worksheet {
  public:
    using WorksheetName = std::string;
    using WorksheetDispatch = IDispatch*;

    Worksheet(WorksheetDispatch dispatch);

    Worksheet() = delete;

    // avoid copying this class, unless a specific need arises
    Worksheet(Worksheet&) = delete;
    Worksheet& operator=(Worksheet&) = delete;

    Worksheet(Worksheet&&) = default;
    Worksheet& operator=(Worksheet&&) = default;

    ~Worksheet();

    [[nodiscard]] WorksheetName getName() const;

    Cell& getCell(const std::string& cellCoords);

    [[nodiscard]] WorksheetDispatch getDispatch() const noexcept;

  private:
    std::map<std::string, std::unique_ptr<Cell>> m_cells;
    WorksheetDispatch m_worksheetDispatch{ nullptr };
  };

}  // namespace office::excel
