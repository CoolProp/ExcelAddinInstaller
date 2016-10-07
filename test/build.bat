rmdir /s /q build bin
mkdir bin\MicrosoftExcel
mkdir build\32bitDLL
mkdir build\64bitDLL
mkdir build\installer
rmdir /s /q install_root

pushd build\32bitDLL & cmake ..\.. -DCOOLPROP_SHARED_LIBRARY=ON -DCOOLPROP_STDCALL_LIBRARY=ON -G"Visual Studio 10" & popd
pushd build\32bitDLL & cmake --build . --target install --config Release & popd

pushd build\64bitDLL & cmake ..\.. -DCOOLPROP_SHARED_LIBRARY=ON -G"Visual Studio 10 Win64" & popd
pushd build\64bitDLL & cmake --build . --target install --config Release & popd

pushd build\installer & cmake ..\.. -DCOOLPROP_EXCEL_MODULE=ON -G"Visual Studio 10 Win64" & popd
pushd build\installer & cmake --build . --target COOLPROP_EXCEL & popd


pushd build\installer & cmake ..\.. -DCOOLPROP_EXCEL_MODULE=ON -G"Visual Studio 10 Win64" & popd
pushd build\installer & cmake ..\.. -DCOOLPROP_EXCEL_MODULE=ON & popd

cmake ..\.. -DCOOLPROP_EXCEL_MODULE=ON -G"Visual Studio 10 Win64"
cmake --build . --target COOLPROP_EXCEL

cmake ..\CoolProp.win.git -DCOOLPROP_EXCEL_MODULE=ON -G"Visual Studio 10 Win64"
cmake --build . --target COOLPROP_EXCEL

cmake ..\CoolProp.win.git -DCOOLPROP_WINDOWS_PACKAGE=ON -G"Visual Studio 10 Win64"
cmake --build . --target COOLPROP_WINDOWS_PACKAGE_DELETE
cmake --build . --target COOLPROP_WINDOWS_PACKAGE_PREPARE
cmake --build . --target COOLPROP_WINDOWS_PACKAGE_SHARED_LIBRARIES
cmake --build . --target COOLPROP_WINDOWS_PACKAGE_EES
cmake --build . --target COOLPROP_WINDOWS_PACKAGE_EXCEL
cmake --build . --target COOLPROP_WINDOWS_PACKAGE_INSTALLER

