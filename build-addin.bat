cd Addins\PowerPoint\Microsoft\MarkLogic_PowerPointAddin\MarkLogic_PowerPointAddin

rem call "C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\vcvarsall.bat"
call  "C:\Program Files\Microsoft Visual Studio 9.0\VC\vcvarsall.bat"

msbuild /target:publish
