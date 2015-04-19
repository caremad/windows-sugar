@echo off

set downloadDir="target"

if exist "%downloadDir%\" (
  choice /m "Download directory %downloadDir%'' already exists; do you want to delete it"
  IF ERRORLEVEL 2 (
    echo Aborting installation
    exit /b
  )
  echo "Delete %downloadDir%"
  @RD /S /Q "%downloadDir%"
)

echo "Proceeding..."s

cscript download.vbs

echo Done!
