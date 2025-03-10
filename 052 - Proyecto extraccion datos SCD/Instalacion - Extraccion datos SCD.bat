powershell Expand-Archive 'Extraccion datos SCD_r02.zip' -DestinationPath '%USERPROFILE%\Documents\'
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\Desktop\Generacion listado GOOSE.lnk');$s.TargetPath='%userprofile%\Documents\Extraccion datos SCD_r02\Extraccion datos SCD\bin\Debug\Extraccion datos SCD.exe';$s.Save()"
powershell Expand-Archive "Planilla template GOOSE.zip" -DestinationPath '%USERPROFILE%\Desktop\'
powershell Expand-Archive "Planilla template GOOSE.zip" -DestinationPath '%USERPROFILE%\Escritorio\'
echo "Ya está listo para usarse la aplicación!"
pause