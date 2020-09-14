@echo off
IF "%PROCESSOR_ARCHITECTURE%"=="x86" (set bit=x86) ELSE (set bit=x64)

:echo Notice: do NOT restart your computer before installation finished
echo 注意: 在安装结束之前请不要重启电脑


echo 安装 python 3.8 ...
echo 注意安装时勾选《第 1 页》的"Add Python 3.8 to PATH"
:start /wait .\python-2.7.14-%bit%.msi
start /wait .\Install\python-3.8.5-amd-%bit%.exe

:echo update pip packet ...
:pip install --upgrade

echo 安装 pyserial ...
pip install -U .\Install\pyserial-3.4-py2.py3-none-any.whl

echo 安装 xlrd ...
pip install -U .\Install\xlrd-1.2.0-py2.py3-none-any.whl

echo 安装 xlwt ...
pip install -U .\Install\xlwt-1.3.0-py2.py3-none-any.whl

echo 安装 xlutils ...
pip install -U .\Install\xlutils-2.0.0-py2.py3-none-any.whl


goto 111
echo installing Visual Studio Code ...
start /wait .\VSCodeSetup-%bit%-1.21.1.exe

set PYTHON=
for %%D in (C D E F G H I J K) do (
    IF EXIST %%D:\Python27\python.exe (set PYTHON=%%D:\Python27)
)
IF "%PYTHON%" EQU "" (
    set /p PYTHON="Can not find Python, enter where you have installed it(e.g: %systemdrive%\Python27):"
)
IF NOT EXIST %PYTHON%\python.exe (
    echo %PYTHON%\python.exe none exist, install failed
)
IF NOT EXIST %PYTHON%\python.exe (
    exit /B 1
)
set PATH=%PATH%;%PYTHON%;%PYTHON%\Scripts
setx PATH "%PATH%"

for %%D in (C D E F G H I J K) do (
    IF EXIST %%D:\Program Files\Git\cmd\git.exe (set PATH=%PATH%;%%D:\Program Files\Git\cmd)
)

for %%D in (C D E F G H I J K) do (
    IF EXIST %%D:\Program File\Microsoft VS Code\bin\code.exe (set PATH=%PATH%;%%D:\Program File\Microsoft VS Code\bin)
)

echo installing aos-cube ...
pip install -U aos-cube
IF errorlevel 1 (
    echo install aos-cube failed
) else (
    echo install aos-cube succeed
)

echo installing scons ...
pip install -U setuptools
pip install -U wheel
python -m pip install -U pip
pip install -U scons
IF errorlevel 1 (
    echo install scons failed
) else (
    echo install scons succeed
)

echo installing cpptools extensions ...
echo Pleas close CMD window when finished, and input N to continue
start /wait code --install-extension cpptools.vsix

echo installing alios-studio extensions ...
echo Pleas close CMD window when finished, and input N to continue
start /wait code --install-extension alios-studio.vsix

echo installing MobaXterm ...
start /wait .\MobaXterm_Setup_8.4.msi
echo installing FTDI serial port driver ...
start /wait .\FTDI_VCP_Driver.exe
echo installing CP210x serial port driver ...
start /wait .\CP210x_Windows_Drivers\CP210xVCPInstaller_%bit%.exe
echo installing JLink debugger driver ...
start /wait .\JLink_Windows_V620c.exe
echo installing ST-Link debugger driver ...
start /wait .\ST-LINK_Utility_Setup.exe
echo "最后执行zadig_2.2.exe，在J-Link连接电脑情况下，选择Options - List ALL Devices - J-Link（Interface 2），根据J-Link的类型选择libusb-win32或者libusbK两种修复驱动"
start /wait .\zidag\zadig_2.2.exe

:111
echo installation finished, press restart your computer to make everything effective
pause
