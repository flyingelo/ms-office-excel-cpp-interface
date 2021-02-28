#include <filesystem>
#include <memory>
#include <string>

struct IDispatch;

namespace office::excel {

struct CLSIDContainer;

struct OpenSpreadsheetArguments {

  enum class UpdateLinks {
    Default = -1,
    DoNotUpdateLinks = 0,
    UpdateLinks = 3    
  };

  enum class Format {
    Default = -1,
    Tabs = 1,
    Commas = 2,
    Spaces = 3,
    Semicolons = 4,
    Nothing = 5,
    CustomCharacter = 6
  };

  std::filesystem::path fileName;
  enum UpdateLinks updateLinks{UpdateLinks::Default};
  bool readOnly{false};
  enum Format format{Format::Default};
  std::wstring password;
  
  OpenSpreadsheetArguments(const std::filesystem::path& fileName) :
    fileName(std::filesystem::absolute(fileName)) {}
  OpenSpreadsheetArguments(const std::string& fileName) :
    OpenSpreadsheetArguments(std::filesystem::path(fileName)) {}
  OpenSpreadsheetArguments(const std::wstring& fileName) :
    OpenSpreadsheetArguments(std::filesystem::path(fileName)) {}

  OpenSpreadsheetArguments() = default;
  OpenSpreadsheetArguments(OpenSpreadsheetArguments&) = default;
  OpenSpreadsheetArguments(OpenSpreadsheetArguments&&) = default;
  OpenSpreadsheetArguments& operator=(OpenSpreadsheetArguments&&) = default;
  OpenSpreadsheetArguments& operator=(OpenSpreadsheetArguments&) = default;
};

class MicrosoftExcel
{
public:

  MicrosoftExcel(bool keepAlive);

  MicrosoftExcel() = delete;

  // avoid copying this class, unless a specific need arises
  MicrosoftExcel(MicrosoftExcel&) = delete; 
  MicrosoftExcel& operator=(MicrosoftExcel&) = delete;

  MicrosoftExcel(MicrosoftExcel&&) = default;
  MicrosoftExcel& operator=(MicrosoftExcel&&) = default;

  ~MicrosoftExcel();

  void makeVisible();

  // https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
  void openSpreadsheet(const OpenSpreadsheetArguments&);

  void selectWorksheet(const std::string& worksheetName);
  void setCellValue(const std::string& cellRange, const std::string& value);
  void save();

private:

  using CLSIDMember = std::unique_ptr<CLSIDContainer>;
  using ExcelApp = IDispatch*;
  using Workbooks = IDispatch*;
  using Workbook = IDispatch*;
  using Worksheets = IDispatch*;
  using Worksheet = IDispatch*;
  
  bool m_keepAlive{false};
  CLSIDMember m_clsidContainer;
  ExcelApp m_excelApp{nullptr};
  Workbooks m_workbooks{nullptr};
  Workbook m_workbook{nullptr};
  Worksheets m_worksheets{nullptr};
  Worksheet m_worksheet{nullptr};  
};

}
