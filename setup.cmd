@ECHO OFF

SET NAME=FaXtract

DEL /F /S /Q build dist winstaller

python setup.py py2exe

IF EXIST C:\PROGRA~1\7-Zip\7z.exe (
    ECHO Creating ZIP file of the installation folder...
    IF EXIST README COPY README dist\README.txt
    IF EXIST includes\%NAME%.ini COPY %NAME%.ini dist\
    CD dist
    IF NOT EXIST ..\winstaller\. MKDIR ..\winstaller
    C:\PROGRA~1\7-Zip\7z.exe a -r -tzip ..\winstaller\%NAME%_EXE.zip *
    CD ..
    ECHO Look for "%NAME%_EXE.zip" in the winstaller folder
)
