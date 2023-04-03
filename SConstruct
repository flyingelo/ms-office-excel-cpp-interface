
def getCxxFlags():
    return ["/EHsc", "/W4", "/WX", "/std:c++20", "/O2"]

env = Environment()

for cxxFlag in getCxxFlags():
    env.Append(CXXFLAGS = cxxFlag)

env.Replace(VARIANT_DIR = "Build")
env.Replace(SOURCE_DIR = "Source")

VariantDir(
    variant_dir = env['VARIANT_DIR'],
    src_dir = env['SOURCE_DIR'],
    duplicate = 0)

lib = env.Library(
    'Build/ExcelInterface/ExcelInterface',
    source = [
        'Build/ExcelInterface/src/Ole.cpp',
        'Build/ExcelInterface/src/Cell.cpp',
        'Build/ExcelInterface/src/Worksheet.cpp',
        'Build/ExcelInterface/src/Workbook.cpp',
        'Build/ExcelInterface/src/ExcelInterface.cpp'])

Install('lib', lib)
Install('include', [
    'Source/ExcelInterface/Cell.hpp',
    'Source/ExcelInterface/ExcelInterface.hpp',
    'Source/ExcelInterface/Workbook.hpp',
    'Source/ExcelInterface/Worksheet.hpp'])

coordsTest = env.Program(
    'Build/tests/CoordsTests',
    source = ['Build/tests/CoordsTests.cpp'],
    CPPPATH = ['Source/ExcelInterface']
)

test = env.Program(
    'Build/tests/Tests',
    source = ['Build/tests/Tests.cpp'],
    CPPPATH = ['include'],
    LIBPATH = ['lib'],
    LIBS = ['ExcelInterface', 'ole32.lib', 'oleaut32.lib']
)

perfTests = env.Program(
    'Build/tests/PerformanceTests',
    source = ['Build/tests/PerformanceTests.cpp'],
    CPPPATH = ['include'],
    LIBPATH = ['lib'],
    LIBS = ['ExcelInterface', 'ole32.lib', 'oleaut32.lib']
)
