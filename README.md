# ms-office-excel-cpp-interface
C++ APIs for Microsoft Excel

For more information, see:
https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/automate-excel-from-c

This module is just a modern C++ wrapper around those APIs

Build command:
`cl /std:c++17 /EHsc /W4 /WX /nologo src/ExcelInterface.cpp tests/Tests.cpp /link "ole32.lib" "oleaut32.lib" /SUBSYSTEM:CONSOLE /OUT:Tests.exe`
