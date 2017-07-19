@echo off

echo import reg

@set baseDir="D:\mywork\WordAddInTest2010\WpsWordAddin\WpsWordAddin\install\bin"

regedit /s  D:\mywork\WordAddInTest2010\WpsWordAddin\WpsWordAddin\install\bin\myreg.reg
 
C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm  %baseDir%\Debug\WpsWordAddin.dll /tlb:%baseDir%\Debug\WpsWordAddin.tlb
@SET GACUTIL="%baseDir%\NETFX 4.0 Tools\gacutil.exe"
Echo Install the dll into GAC
%GACUTIL% -i %baseDir%\Debug\WpsWordAddin.dll
%GACUTIL% -i %baseDir%\Debug\Word.dll
%GACUTIL% -i %baseDir%\Debug\Office.dll

pause