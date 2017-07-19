wpsAddin,本示例代码，基于wps 2016的插件开发，兼容wps 2013，有参考价值，另附部署：

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


myreg.reg文件内容如下：

Windows Registry Editor Version 5.00
[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Word\Addins\WpsWordAddin.WPSWord2016]
"FriendlyName"="WpsWordAddin"
"Description"="wps word示例"
"LoadBehavior"=dword:00000003
"CommandLineSafe"=dword:00000001
[HKEY_CURRENT_USER\Software\Kingsoft\Office\WPS\AddinsWL]
"WpsWordAddin.WPSWord2016"=""