#pragma once

#include <algorithm>
#include <string>

static inline std::wstring to_wstring(const std::string& src) {
  std::wstring trg(src.size(), L' ');
  std::copy(std::begin(src), std::end(src), std::begin(trg));
  return trg;
}

static inline std::string to_string(const std::wstring& src) {
  constexpr int maxlen{ 255 };
  std::string trg(src.size(), ' ');
  for (std::size_t i = 0; i < src.size(); ++i) {
    if (static_cast<int>(src[i]) < maxlen) {
      trg[i] = static_cast<char>(src[i]);
    }
    else {
      trg[i] = '#';
    }
  }

  return trg;
}
