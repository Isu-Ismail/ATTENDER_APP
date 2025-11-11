; installer.nsi
; This is the NSIS (Nullsoft Scriptable Install System) script.
; Place this file in the ROOT of your repository, next to main.py and app.ico.

;--------------------------------
; Basic Installer Attributes

!define APP_NAME "Attender"
!define COMP_NAME "Your Company Name" ; Change this
!define VERSION "1.0.0" ; This is just a base version, the release will be versioned
!define SETUP_EXE "attender-setup.exe"
!define MAIN_EXE "attender.exe"
!define MAIN_ICON "app.ico"
!define LICENSE_FILE "LICENSE"
!define README_FILE "README.md"

; Output file name
OutFile "dist\${SETUP_EXE}"

; Set the icon for the installer itself
InstallIcon "${MAIN_ICON}"
UninstallIcon "${MAIN_ICON}"

; Default installation directory
InstallDir "$PROGRAMFILES64\${APP_NAME}"

; Request administrator privileges (needed for Program Files)
RequestExecutionLevel admin

;--------------------------------
; Pages
; These are the screens the user will see.

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

;--------------------------------
; Installation Section

Section "Install"
  ; Set the output path
  SetOutPath $INSTDIR
  
  ; Add the main .exe file
  ; This path assumes the .exe is in the 'dist' folder
  ; which is where PyInstaller puts it.
  File "dist\${MAIN_EXE}"
  
  ; Add your other files from the root of the repo
  File "${MAIN_ICON}"
  File "${LICENSE_FILE}"
  File "${README_FILE}"

  ; Create a Start Menu shortcut
  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${MAIN_EXE}" "" "$INSTDIR\${MAIN_ICON}"
  
  ; Create a Desktop shortcut
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${MAIN_EXE}" "" "$INSTDIR\${MAIN_ICON}"

  ; Write uninstaller information to the registry
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName" "${APP_NAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayIcon" "$INSTDIR\${MAIN_ICON}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "Publisher" "${COMP_NAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayVersion" "${VERSION}"
  
  ; Write the uninstaller
  WriteUninstaller "$INSTDIR\uninstall.exe"
  
SectionEnd

;--------------------------------
; Uninstaller Section

Section "Uninstall"
  ; Remove registry keys
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"

  ; Remove files
  Delete "$INSTDIR\${MAIN_EXE}"
  Delete "$INSTDIR\${MAIN_ICON}"
  Delete "$INSTDIR\${LICENSE_FILE}"
  Delete "$INSTDIR\${README_FILE}"
  Delete "$INSTDIR\uninstall.exe"

  ; Remove shortcuts
  Delete "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk"
  Delete "$DESKTOP\${APP_NAME}.lnk"

  ; Remove directories
  RMDir "$SMPROGRAMS\${APP_NAME}"
  RMDir "$INSTDIR"
  
SectionEnd

