; installer.nsi

;--------------------------------
; Basic Installer Attributes
!define APP_NAME "Attender"
!define COMP_NAME "Your Company Name"
!define VERSION "1.0.0"
!define SETUP_EXE "attender-setup.exe"
!define MAIN_EXE "attender.exe"
!define MAIN_ICON "app.ico"
!define LICENSE_FILE "LICENSE"
!define README_FILE "README.md"

; Output file name
OutFile "dist\${SETUP_EXE}"

; Set the icon for the installer itself (using compatible command)
Icon "${MAIN_ICON}"
UninstallIcon "${MAIN_ICON}"

; Default installation directory
InstallDir "$PROGRAMFILES64\${APP_NAME}"

; Request administrator privileges
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
  ; Set the directory to install files to
  SetOutPath $INSTDIR
  
  ; Add your program files
  File "dist\${MAIN_EXE}"
  File "${MAIN_ICON}"
  File "${LICENSE_FILE}"
  File "${README_FILE}"

  ; --- SHORTCUT CREATION ---
  
  ; Create a Start Menu shortcut
  ; This creates a shortcut named "Attender"
  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${MAIN_EXE}" "" "$INSTDIR\${MAIN_ICON}"
  
  ; Create a Desktop (home screen) shortcut
  ; This creates a shortcut named "AttS_APPender"
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${MAIN_EXE}" "" "$INSTDIR\${MAIN_ICON}"

  ; --- END SHORTCUT CREATION ---

  ; Write uninstaller information to the registry
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName" "${APP_NAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayIcon" "$INSTDIR\${MAIN_ICON}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "Publisher" "${COMP_NAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayVersion" "${VERSION}"
  
  ; Write the uninstaller program
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