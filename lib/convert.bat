@rem Convert the libs from PlatformSDK to COFF format required by the dmd 32bit.

coffimplib "%DXSDK_DIR%\lib\x86\d3d10.lib" d3d10.lib
coffimplib "%DXSDK_DIR%\lib\x86\d3d11.lib" d3d11.lib
coffimplib "%DXSDK_DIR%\lib\x86\d3dx11.lib" d3dx11.lib
coffimplib "%DXSDK_DIR%\lib\x86\d3dcompiler.lib" d3dcompiler.lib
coffimplib "%DXSDK_DIR%\lib\x86\d3dx10.lib" d3dx10.lib
coffimplib "%DXSDK_DIR%\lib\x86\dxgi.lib" dxgi.lib
coffimplib "%DXSDK_DIR%\lib\x86\x3daudio.lib" x3daudio.lib
coffimplib "%DXSDK_DIR%\lib\x86\xinput.lib" xinput.lib
coffimplib "%DXSDK_DIR%\lib\x86\d2d1.lib" d2d1.lib
coffimplib "%DXSDK_DIR%\lib\x86\dwrite.lib" dwrite.lib

SET WindowsSDKDir=%ProgramFiles(x86)%\Microsoft SDKs\Windows\v7.1A\

coffimplib "%WindowsSDKDir%\lib\gdi32.lib"     gdi32.lib
coffimplib "%WindowsSDKDir%\lib\aclui.lib"     aclui.lib
coffimplib "%WindowsSDKDir%\lib\netapi32.lib"  netapi32.lib
coffimplib "%WindowsSDKDir%\lib\netapi.lib"    netapi.lib
coffimplib "%WindowsSDKDir%\lib\oleacc.lib"    oleacc.lib
coffimplib "%WindowsSDKDir%\lib\powrprof.lib"  powrprof.lib
coffimplib "%WindowsSDKDir%\lib\rasapi32.lib"  rasapi32.lib
coffimplib "%WindowsSDKDir%\lib\secur32.lib"   secur32.lib
coffimplib "%WindowsSDKDir%\lib\setupapi.lib"  setupapi.lib
coffimplib "%WindowsSDKDir%\lib\shlwapi.lib"   shlwapi.lib
coffimplib "%WindowsSDKDir%\lib\vfw32.lib"     vfw32.lib
coffimplib "%WindowsSDKDir%\lib\OleAut32.lib"  OleAut32.lib
coffimplib "%WindowsSDKDir%\lib\Ole32.lib"     Ole32.lib
coffimplib "%WindowsSDKDir%\lib\Crypt32.lib"   Crypt32.lib
coffimplib "%WindowsSDKDir%\lib\gdiplus.lib"   gdiplus.lib
