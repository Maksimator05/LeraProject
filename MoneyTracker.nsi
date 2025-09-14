; Настройка кодировки для русского языка
Unicode true

; Подключаем современный интерфейс
!include "MUI2.nsh"
!include "FileFunc.nsh"

; Название приложения
Name "MoneyTracker"
; Имя выходного файла установщика
OutFile "MoneyTracker_Setup.exe"
; Папка установки по умолчанию
InstallDir "$PROGRAMFILES64\MoneyTracker"
; Запрашиваем права администратора
RequestExecutionLevel admin

; Настройки интерфейса
!define MUI_ABORTWARNING

; Используем стандартные иконки NSIS (убедитесь что они существуют)
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; Страницы установщика
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "license.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

; Страницы удаления
!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

; Русский язык
!insertmacro MUI_LANGUAGE "Russian"

; Переменные
Var AppDataPath

Section "Основные файлы" SecMain
    ; Устанавливаем выходной путь
    SetOutPath "$INSTDIR"

    ; Копируем файлы
    File "dist\MoneyTracker.exe"
    File "money_tracker.db"
    File "settings.csv"
    File "license.txt"

    ; Создаем папку в AppData для базы данных
    StrCpy $AppDataPath "$LOCALAPPDATA\MoneyTracker"
    CreateDirectory "$AppDataPath"

    ; Копируем первоначальную базу в AppData (если ее там нет)
    IfFileExists "$AppDataPath\money_tracker.db" SkipDBCopy
    CopyFiles "$INSTDIR\money_tracker.db" "$AppDataPath\"
    SkipDBCopy:

    ; Создаем ярлыки (используем иконку из EXE-файла)
    CreateShortCut "$SMPROGRAMS\MoneyTracker.lnk" "$INSTDIR\MoneyTracker.exe" "" "$INSTDIR\MoneyTracker.exe" 0
    CreateShortCut "$DESKTOP\MoneyTracker.lnk" "$INSTDIR\MoneyTracker.exe" "" "$INSTDIR\MoneyTracker.exe" 0

    ; Записываем информацию для удаления
    WriteUninstaller "$INSTDIR\Uninstall.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MoneyTracker" \
        "DisplayName" "MoneyTracker"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MoneyTracker" \
        "UninstallString" "$\"$INSTDIR\Uninstall.exe$\""
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MoneyTracker" \
        "DisplayIcon" "$INSTDIR\MoneyTracker.exe,0"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MoneyTracker" \
        "Publisher" "Ваша компания"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MoneyTracker" \
        "DisplayVersion" "2.5"

    ; Рассчитываем размер установки
    ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
    IntFmt $0 "0x%08X" $0
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MoneyTracker" \
        "EstimatedSize" "$0"
SectionEnd

Section "Uninstall"
    ; Удаляем файлы
    Delete "$INSTDIR\MoneyTracker.exe"
    Delete "$INSTDIR\money_tracker.db"
    Delete "$INSTDIR\settings.csv"
    Delete "$INSTDIR\license.txt"
    Delete "$INSTDIR\Uninstall.exe"

    ; Удаляем папку установки
    RMDir "$INSTDIR"

    ; Удаляем ярлыки
    Delete "$SMPROGRAMS\MoneyTracker.lnk"
    Delete "$DESKTOP\MoneyTracker.lnk"

    ; Удаляем запись из реестра
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MoneyTracker"

    ; ВАЖНО: Не удаляем базу данных из AppData чтобы сохранить данные пользователя!
    ; Если хотите удалять и базу - раскомментируйте следующую строку:
    ; RMDir /r "$LOCALAPPDATA\MoneyTracker"
SectionEnd

; Функция инициализации
Function .onInit
    ; Проверяем, не установлена ли уже программа
    ReadRegStr $R0 HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MoneyTracker" "UninstallString"
    StrCmp $R0 "" done

    MessageBox MB_OKCANCEL|MB_ICONEXCLAMATION \
        "MoneyTracker уже установлен. $\n$\nНажмите OK для удаления предыдущей версии или Cancel для отмены." \
        IDOK uninst
    Abort

    uninst:
        ClearErrors
        ExecWait '$R0 _?=$INSTDIR'

        IfErrors no_remove_uninstaller
        Delete $R0
        no_remove_uninstaller:

    done:
FunctionEnd