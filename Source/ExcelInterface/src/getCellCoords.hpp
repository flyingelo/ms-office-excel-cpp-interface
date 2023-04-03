#pragma once

#include <iostream>
#include <stdexcept>
#include <string>

namespace office::excel::utils {

  static std::string getCellCoords(const std::uint32_t row, const std::uint16_t col) {

    constexpr std::uint16_t abcSize{ 26U };

    // create lambda function to get the column letter

    auto getColumnChar = [&](std::uint16_t remainder) {
      auto columnDigit = remainder % abcSize;
      return static_cast<char>(columnDigit + 'A');
    };

    constexpr std::uint32_t maxRow{1048575U};
    constexpr std::uint16_t maxColumn{16383U};

    if (row > maxRow) {
      throw std::runtime_error("Row number " + std::to_string(row) + " is too large. Maximum row number is " +
        std::to_string(maxRow));
    }

    if (col > maxColumn) {
      throw std::runtime_error("Column number " + std::to_string(col) + " is too large. Maximum column number is " +
        std::to_string(maxColumn));
    }

    std::string columnLetters;

    if (col >= (abcSize * (abcSize + 1))) {
      const std::uint16_t offset = col - (abcSize * (abcSize + 1));
      const std::uint16_t remainder = offset / (abcSize * abcSize);
      columnLetters.push_back(getColumnChar(remainder));
    }

    if (col >= abcSize) {
      const std::uint16_t offset = col - abcSize;
      const std::uint16_t remainder = offset / abcSize;
      columnLetters.push_back(getColumnChar(remainder));
    }

    const std::uint16_t offset = col;
    const std::uint16_t remainder = offset;
    columnLetters.push_back(getColumnChar(remainder));

    return columnLetters + std::to_string(row + 1); // Excel uses 1-based indexing for rows and columns
  }

}
