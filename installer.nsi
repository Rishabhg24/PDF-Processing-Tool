!include "MUI2.nsh"
!include "FileFunc.nsh"
!include "nsDialogs.nsh"
!include "LogicLib.nsh"
!include "WinMessages.nsh"
!include "x64.nsh"
; Using built-in NSISdl plugin for downloading Python

; Define installer expiry date (YYYY-MM-DD format)
!define INSTALLER_EXPIRY_DATE "2025-12-31"

; This installer automatically downloads and installs Python 3.11.8 from python.org
; if Python is not already installed on the system. It also sets up environment variables
; and creates a virtual environment with all required packages.

Name "PDF Processing Tool"
OutFile "PDF_Processing_Tool_Setup.exe"
InstallDir "$LOCALAPPDATA\PDF Processing Tool"
RequestExecutionLevel user

!define MUI_ABORTWARNING
!define MUI_ICON "icon.ico"
!define MUI_UNICON "icon.ico"

ShowInstDetails show
ShowUninstDetails show

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY
Page custom LicenseSelectionPage LicenseSelectionPageLeave
Page custom PasswordPage PasswordPageLeave
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "English"

Var LicenseType
Var CustomDays
Var RadioDemo
Var Radio2Years
Var RadioCustom
Var PROGRESS_STEP
Var CustomDaysLabel
Var PasswordInput
Var PythonFound
Var PythonPath
Var PythonScriptsPath

!define TOTAL_STEPS 13

!define DEMO_PASSWORD "DEMO_FUTUREWAVE_RG"
!define LICENSE_PASSWORD "LICENSEWAVE24"
!define APP_VERSION "1.0.0"

; Function to check if installer has expired
Function CheckInstallerExpiry
    ; Get current date
    ${GetTime} "" "L" $0 $1 $2 $3 $4 $5 $6
    
    ; Current date components
    StrCpy $R0 $2  ; Year
    StrCpy $R1 $1  ; Month  
    StrCpy $R2 $0  ; Day
    
    ; Parse expiry date (2025-12-31 format)
    StrCpy $R3 "2025"  ; Year
    StrCpy $R4 "12"    ; Month
    StrCpy $R5 "31"    ; Day
    
    ; Convert to numbers
    IntOp $R0 $R0 + 0
    IntOp $R1 $R1 + 0
    IntOp $R2 $R2 + 0
    IntOp $R3 $R3 + 0
    IntOp $R4 $R4 + 0
    IntOp $R5 $R5 + 0
    
    ; Compare dates
    ${If} $R0 > $R3
        Goto Expired
    ${ElseIf} $R0 == $R3
        ${If} $R1 > $R4
            Goto Expired
        ${ElseIf} $R1 == $R4
            ${If} $R2 > $R5
                Goto Expired
            ${EndIf}
        ${EndIf}
    ${EndIf}
    
    Goto NotExpired
    
    Expired:
        MessageBox MB_OK|MB_ICONSTOP "This installer has expired.Please contact the developer for an updated version."
        Abort "Installer expired"
    
    NotExpired:
        ; Continue with installation
FunctionEnd

Function LicenseSelectionPage
    !insertmacro MUI_HEADER_TEXT "License Selection" "Choose your license type"

    nsDialogs::Create 1018
    Pop $0
    ${If} $0 == error
        Abort
    ${EndIf}

    ${NSD_CreateLabel} 0 0 100% 12u "Please select your license type:"
    Pop $0

    ${NSD_CreateRadioButton} 0 20u 100% 12u "Demo License (1 Month Trial)"
    Pop $RadioDemo
    ${NSD_Check} $RadioDemo

    ${NSD_CreateRadioButton} 0 40u 100% 12u "2 Years License"
    Pop $Radio2Years

    ${NSD_CreateRadioButton} 0 60u 100% 12u "Custom Days License"
    Pop $RadioCustom

    ${NSD_CreateLabel} 0 90u 100% 12u "Custom Days (if selected above):"
    Pop $CustomDaysLabel

    ${NSD_CreateNumber} 0 110u 50% 12u "30"
    Pop $CustomDays
    EnableWindow $CustomDays 0
    EnableWindow $CustomDaysLabel 0

    ${NSD_OnClick} $RadioDemo OnLicenseRadioClick
    ${NSD_OnClick} $Radio2Years OnLicenseRadioClick
    ${NSD_OnClick} $RadioCustom OnLicenseRadioClick

    nsDialogs::Show
FunctionEnd

Function OnLicenseRadioClick
    ${NSD_GetState} $RadioCustom $0
    ${If} $0 == ${BST_CHECKED}
        EnableWindow $CustomDays 1
        EnableWindow $CustomDaysLabel 1
    ${Else}
        EnableWindow $CustomDays 0
        EnableWindow $CustomDaysLabel 0
    ${EndIf}
FunctionEnd

Function LicenseSelectionPageLeave
    ${NSD_GetState} $RadioDemo $R0
    ${NSD_GetState} $Radio2Years $R1
    ${NSD_GetState} $RadioCustom $R2

    ${If} $R0 == ${BST_CHECKED}
        StrCpy $LicenseType "demo"
    ${ElseIf} $R1 == ${BST_CHECKED}
        StrCpy $LicenseType "2years"
    ${ElseIf} $R2 == ${BST_CHECKED}
        StrCpy $LicenseType "custom"
    ${Else}
        StrCpy $LicenseType "demo"
    ${EndIf}

    ${If} $LicenseType == "custom"
        ${NSD_GetText} $CustomDays $0
        ${If} $0 < 1
            MessageBox MB_OK|MB_ICONEXCLAMATION "Please enter a valid number of days (minimum 1)."
            Abort
        ${EndIf}
        StrCpy $CustomDays $0
    ${EndIf}
FunctionEnd

Function PasswordPage
    !insertmacro MUI_HEADER_TEXT "Password Required" "Enter the password for your selected license type."

    nsDialogs::Create 1018
    Pop $0
    ${If} $0 == error
        Abort
    ${EndIf}

    ${NSD_CreateLabel} 0 0 100% 12u "Please enter the password:"
    Pop $0

    ${NSD_CreatePassword} 0 20u 100% 12u ""
    Pop $PasswordInput

    nsDialogs::Show
FunctionEnd

Function PasswordPageLeave
    ${NSD_GetText} $PasswordInput $0
    StrCpy $1 $0

    ; Determine which license type was selected
    ${If} $LicenseType == "demo"
        StrCpy $2 ${DEMO_PASSWORD}
    ${Else}
        StrCpy $2 ${LICENSE_PASSWORD}
    ${EndIf}

    ${If} $1 != $2
        MessageBox MB_OK|MB_ICONSTOP "Incorrect password. Please try again."
        Abort
    ${EndIf}
FunctionEnd

Function .onInit
    Call CheckInstallerExpiry
    StrCpy $PROGRESS_STEP 0
FunctionEnd

Function UpdateProgress
    IntOp $PROGRESS_STEP $PROGRESS_STEP + 1
    IntOp $R0 $PROGRESS_STEP * 100
    IntOp $R0 $R0 / ${TOTAL_STEPS}
    SendMessage $HWNDPARENT 0x402 0 $R0
    StrCpy $1 "PDF Processing Tool Setup ($R0% Complete)"
    SendMessage $HWNDPARENT ${WM_SETTEXT} 0 "STR:$1"
    DetailPrint "Progress: $R0% complete"
FunctionEnd

Section "Check and Install Python"
    DetailPrint "Checking for existing Python installation..."
    
    ; Search for Python 3.x in registry (HKLM/HKCU, 64/32-bit)
    Push $0
    Push $1
    Push $2
    Push $3
    Push $8
    StrCpy $8 11
    StrCpy $PythonFound 0
    
python_search_loop:
    IntOp $8 $8 + 1
    StrCpy $7 "3.$8"
    ReadRegStr $0 HKLM "SOFTWARE\Python\PythonCore\$7\InstallPath" ""
    StrCmp $0 "" 0 python_found_in_registry
    ReadRegStr $0 HKLM "SOFTWARE\WOW6432Node\Python\PythonCore\$7\InstallPath" ""
    StrCmp $0 "" 0 python_found_in_registry
    ReadRegStr $0 HKCU "SOFTWARE\Python\PythonCore\$7\InstallPath" ""
    StrCmp $0 "" 0 python_found_in_registry
    IntCmp $8 15 0 python_search_loop 0

    ; Try to find Python in default installation location
    ExpandEnvStrings $2 "%LOCALAPPDATA%\Programs\Python"
    FindFirst $3 $8 "$2\python*"
    StrCmp $3 "" python_not_in_registry python_found_in_default
    
python_found_in_registry:
    StrCpy $PythonFound 1
    StrCpy $PythonPath "$0"
    StrCpy $PythonScriptsPath "$0\Scripts"
    DetailPrint "Python found in registry at: $PythonPath"
    Goto python_setup_complete
    
python_found_in_default:
    StrCpy $PythonFound 1
    StrCpy $PythonPath "$2\$8"
    StrCpy $PythonScriptsPath "$2\$8\Scripts"
    DetailPrint "Python found in default location at: $PythonPath"
    Goto python_setup_complete
    
python_not_in_registry:
    DetailPrint "Python not found. Downloading from official website..."
    
    ; Download Python 3.11.8 from official website
    DetailPrint "Downloading Python 3.11.8 from python.org..."
    NSISdl::download "https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe" "$TEMP\python-3.11.8-amd64.exe"
    Pop $0
    
    ${If} $0 == "success"
        DetailPrint "Python download completed successfully"
        
        ; Install Python with proper settings
        DetailPrint "Installing Python 3.11.8..."
        ExecWait '"$TEMP\python-3.11.8-amd64.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 SimpleInstall=1 Include_pip=1 Include_dev=0 Include_launcher=1' $0
        
        ${If} $0 == 0
            DetailPrint "Python installation completed successfully"
            
            ; Wait a moment for installation to complete
            Sleep 2000
            
            ; Try to find the newly installed Python
            StrCpy $8 11
retry_python_registry:
            IntOp $8 $8 + 1
            StrCpy $7 "3.$8"
            ReadRegStr $0 HKLM "SOFTWARE\Python\PythonCore\$7\InstallPath" ""
            StrCmp $0 "" 0 python_installed_successfully
            ReadRegStr $0 HKLM "SOFTWARE\WOW6432Node\Python\PythonCore\$7\InstallPath" ""
            StrCmp $0 "" 0 python_installed_successfully
            ReadRegStr $0 HKCU "SOFTWARE\Python\PythonCore\$7\InstallPath" ""
            StrCmp $0 "" 0 python_installed_successfully
            IntCmp $8 15 0 retry_python_registry 0
            
            ; Try default location
            ExpandEnvStrings $2 "%LOCALAPPDATA%\Programs\Python"
            FindFirst $3 $8 "$2\python*"
            StrCmp $3 "" python_installation_failed python_installed_successfully
            
python_installed_successfully:
            StrCpy $PythonFound 1
            StrCpy $PythonPath "$0"
            StrCpy $PythonScriptsPath "$0\Scripts"
            DetailPrint "Python installed successfully at: $PythonPath"
            
            ; Update PATH environment variable
            DetailPrint "Updating PATH environment variable..."
            ReadRegStr $3 HKCU "Environment" "Path"
            StrCpy $4 "$3;$PythonPath;$PythonScriptsPath"
            WriteRegExpandStr HKCU "Environment" "Path" $4
            SendMessage ${HWND_BROADCAST} ${WM_SETTINGCHANGE} 0 "STR:Environment" /TIMEOUT=5000
            DetailPrint "PATH environment variable updated"
            
        ${Else}
            DetailPrint "Python installation failed with error code: $0"
            Goto python_installation_failed
        ${EndIf}
        
        ; Clean up downloaded installer
        Delete "$TEMP\python-3.11.8-amd64.exe"
        
    ${Else}
        DetailPrint "Python download failed with error: $0"
        Goto python_installation_failed
    ${EndIf}
    
    Goto python_setup_complete
    
python_installation_failed:
    MessageBox MB_OK|MB_ICONSTOP "Python installation failed. Please install Python 3.11 or later manually from python.org and try again."
    Abort "Python installation required"

python_setup_complete:
    ${If} $PythonFound == 1
        DetailPrint "Python setup completed successfully"
        DetailPrint "Python Path: $PythonPath"
        DetailPrint "Python Scripts Path: $PythonScriptsPath"
        
        ; Set up Python virtual environment and install packages
        DetailPrint "Setting up Python virtual environment..."
        nsExec::ExecToLog '"$PythonPath\python.exe" -m venv "$INSTDIR\venv"'
        DetailPrint "Upgrading pip..."
        nsExec::ExecToLog '"$INSTDIR\venv\Scripts\python.exe" -m pip install --upgrade pip'
        DetailPrint "Installing required packages..."
        nsExec::ExecToLog '"$INSTDIR\venv\Scripts\python.exe" -m pip install PyMuPDF pdfplumber PyPDF2 pandas python-docx ttkbootstrap pdf2docx openpyxl Pillow pymongo'
        
        ; Generate license and checksum files
        DetailPrint "Generating license and checksum files..."
        nsExec::ExecToLog '"$INSTDIR\venv\Scripts\python.exe" -c "from checksum_validator import generate_data_file_with_license; generate_data_file_with_license(\"$LicenseType\", \"$CustomDays\")"'
        nsExec::ExecToLog '"$INSTDIR\venv\Scripts\python.exe" -c "from checksum_validator import calculate_checksum, save_checksum; import os; data_file = os.path.join(os.getcwd(), \"app_data.json\"); checksum = calculate_checksum(data_file); save_checksum(data_file, checksum)"'
        
        Call UpdateProgress
    ${Else}
        MessageBox MB_OK|MB_ICONSTOP "Python setup failed. Please check the installation logs."
        Abort "Python setup failed"
    ${EndIf}
    
    Pop $8
    Pop $3
    Pop $2
    Pop $1
    Pop $0
SectionEnd

Section "Install"
    SetOutPath "$INSTDIR"
    DetailPrint "Creating installation directory..."
    CreateDirectory "$INSTDIR"
    CreateDirectory "$INSTDIR\venv"
    Call UpdateProgress

    DetailPrint "Copying application files..."
    File "PDF_PROCESSING_TOOL.py"
    File "license_generator_gui.py"
    File "checksum_validator.py"
    File "icon.ico"
    Call UpdateProgress

    DetailPrint "Creating launch script..."
    FileOpen $0 "$INSTDIR\launch.vbs" w
    FileWrite $0 'Set objShell = CreateObject("WScript.Shell")$\r$\n'
    FileWrite $0 'Set objFSO = CreateObject("Scripting.FileSystemObject")$\r$\n'
    FileWrite $0 'strPath = objFSO.GetParentFolderName(WScript.ScriptFullName)$\r$\n'
    FileWrite $0 'objShell.CurrentDirectory = strPath$\r$\n'
    FileWrite $0 'licenseFile = strPath & "\license.key"$\r$\n'
    FileWrite $0 'If Not objFSO.FileExists(licenseFile) Then$\r$\n'
    FileWrite $0 '    objShell.Run """" & strPath & "\venv\Scripts\pythonw.exe"" """ & strPath & "\license_generator_gui.py""", 1, True$\r$\n'
    FileWrite $0 '    If objFSO.FileExists(licenseFile) Then$\r$\n'
    FileWrite $0 '        objShell.Run """" & strPath & "\venv\Scripts\pythonw.exe"" """ & strPath & "\PDF_PROCESSING_TOOL.py""", 1, False$\r$\n'
    FileWrite $0 '    End If$\r$\n'
    FileWrite $0 'Else$\r$\n'
    FileWrite $0 '    objShell.Run """" & strPath & "\venv\Scripts\pythonw.exe"" """ & strPath & "\PDF_PROCESSING_TOOL.py""", 1, False$\r$\n'
    FileWrite $0 'End If$\r$\n'
    FileClose $0

    DetailPrint "Creating batch launcher for compatibility..."
    FileOpen $0 "$INSTDIR\launch.bat" w
    FileWrite $0 '@echo off$\r$\n'
    FileWrite $0 'start "" "%~dp0launch.vbs"$\r$\n'
    FileClose $0

    DetailPrint "Creating shortcuts..."
    CreateShortCut "$DESKTOP\PDF Processing Tool.lnk" "$INSTDIR\launch.vbs" "" "$INSTDIR\icon.ico"
    CreateDirectory "$SMPROGRAMS\PDF Processing Tool"
    CreateShortCut "$SMPROGRAMS\PDF Processing Tool\PDF Processing Tool.lnk" "$INSTDIR\launch.vbs" "" "$INSTDIR\icon.ico"
    CreateShortCut "$SMPROGRAMS\PDF Processing Tool\License Generator.lnk" "$INSTDIR\launch.vbs" "" "$INSTDIR\icon.ico"
    CreateShortCut "$SMPROGRAMS\PDF Processing Tool\Uninstall.lnk" "$INSTDIR\uninstall.exe"
    Call UpdateProgress

    DetailPrint "Creating uninstaller..."
    WriteUninstaller "$INSTDIR\uninstall.exe"
    Call UpdateProgress

    DetailPrint "Updating registry..."
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PDFProcessingTool" "DisplayName" "PDF Processing Tool"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PDFProcessingTool" "UninstallString" "$INSTDIR\uninstall.exe"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PDFProcessingTool" "DisplayIcon" "$INSTDIR\icon.ico"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PDFProcessingTool" "Publisher" "Rishabh private limited"
    Push $R0
    ${GetTime} "" "L" $0 $1 $2 $3 $4 $5 $6
    StrCpy $R0 "$2$1$0" ; YYYYMMDD
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PDFProcessingTool" "InstallDate" $R0
    Pop $R0
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PDFProcessingTool" "DisplayVersion" "${APP_VERSION}"
    WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PDFProcessingTool" "EstimatedSize" 129024
    Call UpdateProgress

    DetailPrint "Installation complete!"
SectionEnd

Section "Uninstall"
    DetailPrint "Removing shortcuts..."
    Delete "$DESKTOP\PDF Processing Tool.lnk"
    Delete "$SMPROGRAMS\PDF Processing Tool\PDF Processing Tool.lnk"
    Delete "$SMPROGRAMS\PDF Processing Tool\License Generator.lnk"
    Delete "$SMPROGRAMS\PDF Processing Tool\Uninstall.lnk"
    RMDir "$SMPROGRAMS\PDF Processing Tool"

    DetailPrint "Removing application folder..."
    SetShellVarContext current
    RMDir /r /REBOOTOK "$INSTDIR"

    DetailPrint "Cleaning up registry..."
    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PDFProcessingTool"

    DetailPrint "Uninstallation complete!"
SectionEnd

