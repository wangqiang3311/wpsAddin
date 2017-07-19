@echo off
regedit /s myreg.reg
C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm   /codebase  D:\mywork\WordAddInTest2010\WpsWordAddin\WpsWordAddin\bin\Debug\WpsWordAddin.dll /tlb:D:\mywork\WordAddInTest2010\WpsWordAddin\WpsWordAddin\bin\Debug\WpsWordAddin.tlb

@SET GACUTIL="C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.6.1 Tools\gacutil.exe"
Echo Install the dll into GAC
%GACUTIL% -i D:\mywork\WordAddInTest2010\WpsWordAddin\WpsWordAddin\bin\Debug\WpsWordAddin.dll

pause