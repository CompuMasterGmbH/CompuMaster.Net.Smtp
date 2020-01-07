::REM -UpdateNuGetExecutable not required since it's updated by VS.NET mechanisms
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& 'CompuMaster.Net.Smtp\_CreateNewNuGetPackage\DoNotModify\New-NuGetPackage.ps1' -ProjectFilePath '.\CompuMaster.Net.Smtp\CompuMaster.Net.Smtp.vbproj' -verbose -NoPrompt -PushPackageToNuGetGallery"
pause