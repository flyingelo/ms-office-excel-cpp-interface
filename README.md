# ms-office-excel-cpp-interface
C++ APIs for Microsoft Excel

Build command:
`cl /std:c++17 /EHsc /W4 /WX /nologo ExcelInterface.cpp Tests.cpp /link "ole32.lib" "oleaut32.lib" /SUBSYSTEM:CONSOLE /OUT:Tests.exe`
