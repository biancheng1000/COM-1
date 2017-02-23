SET Path=%windir%\Microsoft.NET\Framework\v4.0.30319
%Path%\attrib -r ComInteropExample.tlb
%Path%\regasm.exe ComInteropExample.dll /codebase /tlb
pause