set VCVARS="C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat"

cd build

call %VCVARS% x86

cl.exe /Zi /LD /D RISCY_DEBUG /nologo /MP /W4 /Fe: riscylib32.dll ^
/I ..\common /I ..\hook ^
..\common\*.cpp ..\hook\*.cpp ^
/link /incremental:no /manifest:no /def:..\hook\main.def || goto :exit

cl.exe /Zi /EHsc /MTd /D RISCY_DEBUG /nologo /MP /W4 /Fe: riscysave32.exe ^
/I ..\common /I ..\app ^
..\common\*.cpp ..\app\*.cpp ^
/link /incremental:no /manifest:embed || goto :exit

call %VCVARS% x64

cl.exe /Zi /LD /D RISCY_DEBUG /nologo /MP /W4 /Fe: riscylib64.dll ^
/I ..\common /I ..\hook ^
..\common\*.cpp ..\hook\*.cpp ^
/link /incremental:no /manifest:no /def:..\hook\main.def || goto :exit

cl.exe /Zi /EHsc /MTd /D RISCY_DEBUG /nologo /MP /W4 /Fe: riscysave64.exe ^
/I ..\common /I ..\app ^
..\common\*.cpp ..\app\*.cpp ^
/link /incremental:no /manifest:embed || goto :exit

cl.exe /Zi /LD /D RISCY_DEBUG /nologo /MP /W4 /Fe: riscyext.dll ^
/I ..\common /I ..\shellext ^
..\common\*.cpp ..\shellext\*.cpp ^
/link /incremental:no /manifest:no /def:..\shellext\main.def || goto :exit

:exit
cd ..

