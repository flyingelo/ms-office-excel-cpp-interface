
#include <cmath>
#include <filesystem>
#include <iostream>
#include <stdexcept>
#include <string>
#include <vector>

#include "src/getCellCoords.hpp"

template <typename T>
void assert(const T& actual, const T& expected) {
  if (expected != actual) {
    throw std::runtime_error("Expected " + std::to_string(expected) + ", got " + std::to_string(actual));
  }
}

template <>
void assert(const std::string& actual, const std::string& expected) {
  if (expected != actual) {
    throw std::runtime_error("Expected " + expected + ", got " + actual);
  }
}

int main(int argc, const char* argv[]) {
  try {
    std::cout << "==================================================\n";
    std::cout << "Running coords tests\n";

    const std::vector<std::string> arguments(argv, argv + argc);

    constexpr std::uint16_t colZ{25U};
    constexpr std::uint16_t colAA{26U};
    constexpr std::uint16_t colAB{27U};
    constexpr std::uint16_t colUJ{555U};
    constexpr std::uint16_t colZT{695U};
    constexpr std::uint16_t colZZ{701U};
    constexpr std::uint16_t colAAA{702U};
    constexpr std::uint16_t colALM{1000U};
    constexpr std::uint16_t colKMD{7777U};
    constexpr std::uint16_t colMax{16383U};
    constexpr std::uint16_t colTooLarge{colMax + 1};

    constexpr std::uint32_t row1000{1000U};
    constexpr std::uint32_t row5555{5555U};
    constexpr std::uint32_t rowMax{1048575U};
    constexpr std::uint32_t rowTooLarge{rowMax + 1};

    assert<std::string>(office::excel::utils::getCellCoords(0U, 0U), std::string("A1"));
    assert<std::string>(office::excel::utils::getCellCoords(1U, colZ), std::string("Z2"));
    assert<std::string>(office::excel::utils::getCellCoords(2U, colAA), std::string("AA3"));
    assert<std::string>(office::excel::utils::getCellCoords(2U, colAB), std::string("AB3"));
    assert<std::string>(office::excel::utils::getCellCoords(2U, colUJ), std::string("UJ3"));
    assert<std::string>(office::excel::utils::getCellCoords(2U, colZT), std::string("ZT3"));
    assert<std::string>(office::excel::utils::getCellCoords(0U, colZZ), std::string("ZZ1"));
    assert<std::string>(office::excel::utils::getCellCoords(0U, colAAA), std::string("AAA1"));
    assert<std::string>(office::excel::utils::getCellCoords(row1000, colALM), std::string("ALM1001"));
    assert<std::string>(office::excel::utils::getCellCoords(row5555, colKMD), std::string("KMD5556"));
    assert<std::string>(office::excel::utils::getCellCoords(rowMax, colMax), std::string("XFD1048576"));

    bool exceptionThrown{ false };
    try {
      office::excel::utils::getCellCoords(rowTooLarge, 0U);
    }
    catch (const std::exception&) {
      exceptionThrown = true;
    }

    if (!exceptionThrown) {
      throw std::runtime_error("Expected exception not thrown");
    }

    exceptionThrown = false;
    try {
      office::excel::utils::getCellCoords(0U, colTooLarge);
    }
    catch (const std::exception&) {
      exceptionThrown = true;
    }

    if (!exceptionThrown) {
      throw std::runtime_error("Expected exception not thrown");
    }

    std::cout << "All coords tests passed\n";
    std::cout << "==================================================\n";
  }
  catch (const std::exception& e) {
    std::cout << "ERROR: " << e.what() << std::endl;
    return EXIT_FAILURE;
  }

  return EXIT_SUCCESS;
}
