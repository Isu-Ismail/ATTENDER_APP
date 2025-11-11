; installer.nsi
; This is the NSIS (Nullsoft Scriptable Install System) script.
; Place this file in the ROOT of your repository.

;--------------------------------
; Basic Installer Attributes

!define APP_NAME "Attender"
!define COMP_NAME "Your Company Name" ; Change this
!define VERSION "1.0.0" 
!define SETUP_EXE "attender-setup.exe"
!define MAIN_EXE "attender.exe"
!define MAIN_ICON "app.ico"
!define LICENSE_FILE "LICENSE"
!define README_FILE "README.md"

; Output file name
OutFile "dist\${SETUP_EXE}"

; --- THIS IS THE FIX ---
; Set the icon for the installer itself
; We are using "Icon" instead of "InstallIcon" for compatibility
Icon "${MAIN_ICON}"
UninstallIcon "${MAIN_ICON}"
; -----------------------

; Default installation directory
InstallDir "$PROGRAMFILES64\${APP_NAME}"

; Request administrator privileges (needed for Program Files)
RequestExecutionLevel admin

;--------------------------------
; Pages

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

;--------------------------------
; Installation Section

Section "Install"
  SetOutPath $INSTDIR
  
  File "dist\${MAIN_EXE}"
  File "${MAIN_ICON}"
  File "${LICENSE_FILE}"
  File "${README_FILE}"

  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${MAIN_EXE}" "" "$INSTDIR\${MAIN_ICON}"
  
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${MAIN_EXE}" "" "$INSTDIR\${MAIN_ICON}"

  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName" "${APP_NAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayIcon" "$INSTDIR\${MAIN_ICON}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "Publisher" "${COMP_NAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayVersion" "${VERSION}"
  
  WriteUninstaller "$INSTDIR\uninstall.exe"
  
SectionEnd

;--------------------------------
; Uninstaller Section

Section "Uninstall"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"

  Delete "$INSTDIR\${MAIN_EXE}"
  Delete "$INSTDIR\${MAIN_ICON}"
  Delete "$INSTDIR\${LICENSE_FILE}"
  Delete "$INSTDIR\${README_FILE}"
  Delete "$INSTDIR\uninstall.exe"

  Delete "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk"
  Delete "$DESKTOP\${APP_NAME}.lnk"

  RMDir "$SMPROGRAMS\${APP_NAME}"
  RMDir "$INSTDIR"
  
SectionEnd