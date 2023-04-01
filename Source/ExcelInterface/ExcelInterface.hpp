
#pragma once

#include <filesystem>
#include <memory>
#include <string>

#include "Workbook.hpp"

struct IDispatch;

namespace office::excel {

    struct CLSIDContainer;

    struct OpenSpreadsheetArguments {
        std::filesystem::path fileName;
        bool readOnly{ false };
        std::wstring password;

        OpenSpreadsheetArguments(const std::filesystem::path& fileName)
            : fileName(std::filesystem::absolute(fileName)) {}
        OpenSpreadsheetArguments(const std::string& fileName)
            : OpenSpreadsheetArguments(std::filesystem::path(fileName)) {}
        OpenSpreadsheetArguments(const std::wstring& fileName)
            : OpenSpreadsheetArguments(std::filesystem::path(fileName)) {}

        OpenSpreadsheetArguments() = default;
        OpenSpreadsheetArguments(OpenSpreadsheetArguments&) = default;
        OpenSpreadsheetArguments(OpenSpreadsheetArguments&&) = default;
        OpenSpreadsheetArguments& operator=(OpenSpreadsheetArguments&&) = default;
        OpenSpreadsheetArguments& operator=(const OpenSpreadsheetArguments&) =
            default;

        ~OpenSpreadsheetArguments() = default;
    };

    class MicrosoftExcel {
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
        Workbook& openWorkbook(const OpenSpreadsheetArguments&);

        Workbook& getWorkbook();

    private:
        using CLSIDMember = std::unique_ptr<CLSIDContainer>;
        using ExcelApp = IDispatch*;
        using Workbooks = IDispatch*;

        std::unique_ptr<Workbook> m_workbook;

        bool m_keepAlive{ false };
        CLSIDMember m_clsidContainer;
        ExcelApp m_excelApp{ nullptr };
        Workbooks m_workbooks{ nullptr };
    };

}  // namespace office::excel
