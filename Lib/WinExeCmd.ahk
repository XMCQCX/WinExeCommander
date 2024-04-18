/*
    Script:    WinExeCmd.ahk
    Author:    Martin Chartier (XMCQCX)
    Date:      2024-04-17
    Version:   1.0.0
    Tested on: Windows 11
    Github:    https://github.com/XMCQCX/WinExeCommander
    AHKForum:
*/

#Include '..\Lib\JSON.ahk'
#Include '..\Lib\GuiButtonIcon.ahk'
#Include '..\Lib\GuiCtrlTips.ahk'
#Include '..\Lib\LV_GridColor.ahk'

;============================================================================================

Class WinExeCmd {

    static m := Map('mEvents', Map(), 'mUser', Map())       ; Main map object for storing events values and user settings.  
    static mTmp := Map('mEvents', Map(), 'mUser', Map())    ; Map object for temporarily storing changes made from the GUI.
    static mDefault := Map()                                ; Default settings
    static mProfiles := Map(), mProfiles.CaseSense := 'off'
    static mNotifyGUIs := Map('bottomRight', Map(), 'topLeft', Map(), 'topCenter', Map(), 'bottomCenter', Map(), 'topRight', Map(), 'bottomLeft', Map())
    static mMsgBoxGUIs := Map()

    ;============================================================================================
    
    static __New()
    {
        this.debugMode         := false
        this.scriptName        := 'WinExeCommander'
        this.scriptVersion     := 'v1.0.0'
        this.gMainTitle        := 'Events Manager - ' this.scriptName ' ' this.scriptVersion
        this.gEventTitle       := 'Event Creator - ' this.scriptName
        this.gProfileTitle     := 'Profiles Manager - ' this.scriptName
        this.gSettingsTitle    := 'Settings - ' this.scriptName
        this.gAboutTitle       := 'About - ' this.scriptName        
        this.gErrorTitle       := 'Error - ' this.scriptName               
        this.gDelConfirmTitle  := 'Delete Confirmation - ' this.scriptName    
        this.gSaveConfirmTitle := 'Profile Saved - ' this.scriptName               
        this.linkGitHub        := 'https://github.com/XMCQCX/WinExeCommander'
        this.linkPayPal        := 'https://www.paypal.com/paypalme/martinchartier/10'

        if A_ScriptName == 'WinExeCmd.ahk' 
            SplitPath(A_ScriptDir,, &parentDir), A_WorkingDir := parentDir
  
        if !FileExist('Icons.dll')
            MsgBox('The script can`'t start because "Icons.dll" is missing.`n`nThe script will terminate.', 'Missing DLL file - ' this.scriptName, 'Iconx'), ExitApp()

        ; Default settings ============================
        this.mDefault['periodWMIprocess']     := 1250
        this.mDefault['delayWinEvent']        := 1000
        this.mDefault['delayDeviceChange']    := 1250
        this.mDefault['periodTimerProcess']   := 1500
        this.mDefault['periodTimerWindow']    := 1500
        this.mDefault['periodTimerDevice']    := 1500
        this.mDefault['state']                := 0,  this.aState                := [0, 1]
        this.mDefault['log']                  := 0,  this.aLog                  := [0, 1]
        this.mDefault['critical']             := 0,  this.aCritical             := [0, 1]
        this.mDefault['notifInfo']            := 1,  this.aNotifInfo            := [0, 1]
        this.mDefault['notifStatus']          := 1,  this.aNotifStatus          := [0, 1]
        this.mDefault['notifProfileStart']    := 1,  this.aNotifProfileStart    := [0, 1]
        this.mDefault['notifProfileMatch']    := 1,  this.aNotifProfileMatch    := [0, 1]
        this.mDefault['profileLoadEvent']     := 0,  this.aProfileLoadEvent     := [0, 1]
        this.mDefault['alwaysOnTop']          := 0,  this.aAlwaysOnTop          := [0, 1]
        this.mDefault['detectHiddenWindows']  := 0,  this.aDetectHiddenWindows  := [0, 1]
        this.mDefault['winActive']            := 0,  this.aWinActive            := [0, 1]
        this.mDefault['winMinMax']            := '', this.aWinMinMax            := ['', 0, 1, -1]
        this.mDefault['modeProcess']          := 2,  this.aModeProcess          := [1, 2, 3]
        this.mDefault['modeDevice']           := 2,  this.aModeDevice           := [1, 2, 3]
        this.mDefault['modeWindow']           := 2,  this.aModeWindow           := [1, 2, 3, 4, 5, 6]
        this.mDefault['winEventPreset']       := 3,  this.aWinEventPreset       := [1, 2, 3, 4, 5] 
        this.mDefault['winTitleMatchMode']    := 2,  this.aWinTitleMatchMode    := [1, 2, 3, 'RegEx']
        this.mDefault['trayMenuIconSize']     := 20, this.aTrayMenuIconSize     := [16, 20, 24, 28, 32]
        this.mDefault['dwFlagsWinEvent']      := '0x0',          this.aDWflagsWinEvent   := ['0x0', '0x1', '0x2']
        this.mDefault['monitoringProcess']    := 'wmi',          this.aMonitoringProcess := ['Timer', 'WMI']
        this.mDefault['monitoringWindow']     := 'winEvent',     this.aMonitoringWindow  := ['Timer', 'WinEvent']
        this.mDefault['monitoringDevice']     := 'deviceChange', this.aMonitoringDevice  := ['Timer', 'DeviceChange']
        this.mDefault['trayIconLeftClick']    := '- None -'
        this.mDefault['trayIconLeftDClick']   := 'Open Events Manager'
        this.mDefault['soundProfile']         := '- None -'
        this.mDefault['soundCreated']         := '- None -'
        this.mDefault['soundTerminated']      := '- None -'
        this.aThemeIcons := ['main', 'loading', 'exit', 'reload', 'about', 'settings', 'tools', 'edit', 'folder', 'profile', 'select', 'events', 'checkmark']
        this.aEventType := ['process', 'window', 'device']
        this.linesFuncFile := StrSplit( FileOpen(this.scriptName '.ahk', 'r').Read() , '`n')

        ;==============================================
        this.mGuiCtrlTips := Map(
            'gEvent.cb_critical', '
            (
                Uncheck: The function's thread is not critical. (Default)
                Check: The function's thread is critical.                
            )',

            'gEvent.window.ddl_mode', '
            (
                1: Call "Function_Created" for every window ID created. Call "Function_Terminated" for every window ID terminated.
                2: Call "Function_Created" only for the initial window ID created. Call "Function_Terminated" only when the last window ID is terminated. (Default)	
                3: Call "Function_Created" for every window ID created. Call "Function_Terminated" only when the last window ID is terminated.		
                4: Call "Function_Created" for every window created. Call "Function_Terminated" for every window terminated.
                5: Call "Function_Created" only for the initial window created. Call "Function_Terminated" only when the last window is terminated.
                6: Call "Function_Created" for every window created. Call "Function_Terminated" only when the last window is terminated.
            )',

            'gEvent.window.ddl_winTitleMatchMode', '
            (
                1: A window's title must start with the specified WinTitle to be a match.
                2: A window's title can contain WinTitle anywhere inside it to be a match. (Default)
                3: A window's title must exactly match WinTitle.
                RegEx: Regular expression WinTitle matching.
            )',

            'gEvent.window.ddl_winMinMax', '
            (
                Null: Not monitoring WinMinMax.
                0: The window is neither minimized nor maximized.
                1: The window is maximized.
                -1: The window is minimized.
            )',   
            
            'gEvent.window.cb_winActive', '
            (
                Uncheck: Not monitoring for active window. (Default)
                Check: Monitoring for active window.
            )',           
            
            'gEvent.window.cb_detectHiddenWindows', '
            (
                Uncheck: Hidden windows are not detected. (Default)
                Check: Hidden windows are detected.
            )' ,

            'gEvent.window.ddl_monitoring', '
            (
                WinEvent: SetWinEventHook. Sets an event hook function for a range of Windows Events. (Default)
                Timer: Check for the existence of the window at a specified time interval.
            )' ,

            ;=======================
            'gEvent.process.ddl_mode', '
            (
                1: Call "Function_Created" for every process ID created. Call "Function_Terminated" for every process ID terminated.
                2: Call "Function_Created" only for the initial process ID created. Call "Function_Terminated" only when the last process ID is terminated. (Default)
                3: Call "Function_Created" for every process ID created. Call "Function_Terminated" only when the last process ID is terminated.	
            )',

            'gEvent.process.ddl_monitoring', '
            (
                WMI: Check for the existence of the process at a specified time interval using WMI Provider Host process. (Default)
                Timer: Check for the existence of the process at a specified time interval.
            )' ,            

            ;=======================            
            'gEvent.device.ddl_mode', '
            (
                1: Call "Function_Created" for every device connected. Call "Function_Terminated" for every device disconnected.
                2: Call "Function_Created" only for the initial device connected. Call "Function_Terminated" only when the last device is disconnected. (Default)
                3: Call "Function_Created" for every device connected. Call "Function_Terminated" only when the last device is disconnected.	
            )',

            'gEvent.device.ddl_monitoring', '
            (
                DeviceChange: Send message notifications when there is a change to the hardware configuration of a device or the computer. (Default)
                Timer: Check for the existence of the device at a specified time interval.
            )' ,                   
        )

        ;==============================================
        this.mWinEvents := Map(
            'EVENT_SYSTEM_SOUND',            Map('hex', '0x1', 'event', 1),
            'EVENT_SYSTEM_ALERT',            Map('hex', '0x2', 'event', 2),
            'EVENT_SYSTEM_FOREGROUND',       Map('hex', '0x3', 'event', 3),
            'EVENT_SYSTEM_MENUSTART',        Map('hex', '0x4', 'event', 4),
            'EVENT_SYSTEM_MENUEND',          Map('hex', '0x5', 'event', 5),
            'EVENT_SYSTEM_MENUPOPUPSTART',   Map('hex', '0x6', 'event', 6),
            'EVENT_SYSTEM_MENUPOPUPEND',     Map('hex', '0x7', 'event', 7),
            'EVENT_SYSTEM_CAPTURESTART',     Map('hex', '0x8', 'event', 8),
            'EVENT_SYSTEM_CAPTUREEND',       Map('hex', '0x9', 'event', 9),
            'EVENT_SYSTEM_MOVESIZESTART',    Map('hex', '0xA', 'event', 10),
            'EVENT_SYSTEM_MOVESIZEEND',      Map('hex', '0xB', 'event', 11),
            'EVENT_SYSTEM_CONTEXTHELPSTART', Map('hex', '0xC', 'event', 12),
            'EVENT_SYSTEM_CONTEXTHELPEND',   Map('hex', '0xD', 'event', 13),
            'EVENT_SYSTEM_DRAGDROPSTART',    Map('hex', '0xE', 'event', 14),
            'EVENT_SYSTEM_DRAGDROPEND',      Map('hex', '0xF', 'event', 15),
            'EVENT_SYSTEM_DIALOGSTART',      Map('hex', '0x10', 'event', 16),
            'EVENT_SYSTEM_DIALOGEND',        Map('hex', '0x11', 'event', 17),
            'EVENT_SYSTEM_SCROLLINGSTART',   Map('hex', '0x12', 'event', 18),
            'EVENT_SYSTEM_SCROLLINGEND',     Map('hex', '0x13', 'event', 19),
            'EVENT_SYSTEM_SWITCHSTART',      Map('hex', '0x14', 'event', 20),
            'EVENT_SYSTEM_SWITCHEND',        Map('hex', '0x15', 'event', 21),
            'EVENT_SYSTEM_MINIMIZESTART',    Map('hex', '0x16', 'event', 22),
            'EVENT_SYSTEM_MINIMIZEEND',      Map('hex', '0x17', 'event', 23)
        )

        aWinEventPresets := [
            [3, 8],
            [3, 8, 16, 20, 22],
            [3, 8, 9, 16, 17, 20, 21, 22, 23],
            [1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 13, 16, 17, 20, 21, 22, 23],
            [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]
        ]

        this.mWinEventPresets := Map()
        this.gMainMenuWinEventsPreset := Menu()

        Loop aWinEventPresets.Length {
            this.mWinEventPresets[index := A_Index] := Array()
            for value in aWinEventPresets[A_Index] {
                for winEventName, v in this.mWinEvents {
                    if (v['event'] = value) {
                        this.mWinEventPresets[index].Push(winEventName)
                        break
                    }
                }
            }

            index = 5 ? this.aWinEventOrder := this.mWinEventPresets[5] : ''
            this.gMainMenuWinEventsPreset.Add('Preset ' A_Index (this.mDefault['winEventPreset'] = A_Index ? ' (Default)' : ''), this.gMain_MenuWinEventPreset_ItemClick.Bind(this))
        }

        ;==============================================
        for _, value in ['aPeriodTimerProcess','aPeriodTimerWindow', 'aPeriodTimerDevice']
            this.%value% := this.CreateIncrementalArray(interval := 250, min := 500, max := 8000)

        this.aPeriodWMIprocess := this.CreateIncrementalArray(250, 500, 8000)
        this.aDelayDeviceChange := this.CreateIncrementalArray(250, 500, 8000)
        this.aDelayWinEvent := this.CreateIncrementalArray(25, 25, 8000)

        ;==============================================
        this.mTools := Map(
            'deviceInfoFinder', Map('name', 'DeviceInfoFinder', 'path', 'Tools\DeviceInfoFinder\DeviceInfoFinder.ahk', 'iconPath', 'Tools\DeviceInfoFinder\DeviceInfoFinder.ico'),
            'autoHotkeyv2Help', Map('name', 'AutoHotkey v2 Help', 'path', A_ProgramFiles '\AutoHotkey\v2\AutoHotkey.chm'),
            'windowSpy', Map('name', 'WindowSpy', 'path', A_ProgramFiles '\AutoHotkey\WindowSpy.ahk')
        )

        this.mIconsUser32 := Map(
            'icon!', 2,
            'icon?', 3,
            'iconx', 4,
            'iconi', 5,
        )         

        ;  Sounds  ====================================
        this.mSounds := Map(
            '- None -', '',
            'soundx', '*16',
            'soundi', '*64'
        )

        for _, path in [A_WinDir '\Media', A_WorkingDir '\Sounds']
            Loop Files path '\*.wav'
                SplitPath(A_LoopFilePath,,,, &fileName), this.mSounds[fileName] := A_LoopFilePath

        this.aSounds := Array()

        for key, value in this.mSounds
            this.aSounds.Push(key)

        ; JSON settings file to Map Object.
        if (FileExist('Settings.json')) {
            this.mInitValues := this.JSONtoMapObject('Settings.json')
        }
        else { ; JSON settings file does not exist.
            this.mInitValues := Map('mEvents', Map(), 'mUser', Map('mWinEventsStates', Map()))
            this.mInitValues['mUser']['periodWMIprocess'] := this.mDefault['periodWMIprocess']

            for winEventName, value in this.mWinEvents {
                if this.HasVal(winEventName, this.mWinEventPresets[this.mDefault['winEventPreset']])
                    this.mInitValues['mUser']['mWinEventsStates'][winEventName] := 1
                else
                    this.mInitValues['mUser']['mWinEventsStates'][winEventName] := 0
            }
        }

        ; Recycle 'EventsLog.txt' if it contains more than x lines.
        if (FileExist('EventsLog.txt')) {
            StrReplace( FileRead('EventsLog.txt') , '`n', '`n',, &cntLines)

            if cntLines > 350000
                FileRecycle('EventsLog.txt')
        }    

        ;==============================================
        if FileExist('EventsBackup.json')
            this.mEventsBackup := JSON.parse( FileOpen('EventsBackup.json', 'r', 'UTF-8').Read() , keepbooltype := false, as_map := true)
        else
            this.mEventsBackup := Map('mEvents', Map())

        ; Convert all profiles in profiles folder to Map Objects.
        Loop Files 'Profiles\*.json' {
            mProfile := this.JSONtoMapObject(A_LoopFilePath)
            SplitPath(A_LoopFilePath,,,, &profileName)
            this.mProfiles[profileName] := mProfile
        }

        ;============================================== 
        for key, value in this.mDefault {
            if !RegExMatch(key, '^(trayIconLeftClick|trayIconLeftDClick|soundCreated|soundTerminated|soundProfile)$')
                if this.mInitValues['mUser'].Has(key) && this.HasVal(value, this.a%key%)
                    this.m['mUser'][key] := this.mInitValues['mUser'][key]
                else
                    this.m['mUser'][key] := this.mDefault[key] 
        }    
  
        for _, value in ['trayIconLeftClick', 'trayIconLeftDClick'] {
            if this.mInitValues['mUser'].Has(value) && this.HasVal(this.mInitValues['mUser'][value], this.Create_Array_aTrayIconLeftClick())
                this.m['mUser'][value] := this.mInitValues['mUser'][value]
            else
                this.m['mUser'][value] := this.mDefault[value]   
        }
            
        for _, value in ['soundCreated', 'soundTerminated', 'soundProfile'] {
            if this.mInitValues['mUser'].Has(value) && this.HasVal(this.mInitValues['mUser'][value], this.aSounds)
                this.m['mUser'][value] := this.mInitValues['mUser'][value]
            else
                this.m['mUser'][value] := this.mDefault[value] 
        } 
        
        ;==============================================
        for _, value in ['monitoringWMIProcess', 'monitoringTimerProcess', 'monitoringWinEvent', 'monitoringTimerWindow', 'monitoringDeviceChange',
        'monitoringTimerDevice', 'monitoringTrayClick', 'profileMatch', 'isCriticalMethodExecuting'] {
            this.%value% := false
        }        

        ; Customization Icons/Themes ==================
        this.mIcons := Map()  

        this.mIcons['gMain+']      := 'HICON:*' LoadPicture('Icons.dll', 'Icon14 w32', &imageType)
        this.mIcons['gMainX']      := 'HICON:*' LoadPicture('Icons.dll', 'Icon15 w32', &imageType)
        this.mIcons['gMainPaypal'] := 'HICON:*' LoadPicture('Icons.dll', 'Icon21 w32', &imageType)       
        this.mIcons['gMainAOT0']   := 'HICON:*' LoadPicture('Icons.dll', 'Icon31 w48', &imageType)
        this.mIcons['gMainAOT1']   := 'HICON:*' LoadPicture('Icons.dll', 'Icon32 w48', &imageType)        
        this.mIcons['downArrow']   := 'HICON:*' LoadPicture('Icons.dll', 'Icon23 w32', &imageType)
        this.mIcons['?Small']      := 'HICON:*' LoadPicture('Icons.dll', 'Icon24 w32', &imageType)
        this.mIcons['btnGreen']    := 'HICON:*' LoadPicture('Icons.dll', 'Icon17 w32', &imageType)
        this.mIcons['btnGrey']     := 'HICON:*' LoadPicture('Icons.dll', 'Icon19 w32', &imageType)
        this.mIcons['floppy']      := 'HICON:*' LoadPicture('Icons.dll', 'Icon22 w32', &imageType)       
        this.mIcons['unCheckAll']  := 'HICON:*' LoadPicture('Icons.dll', 'Icon34 w32', &imageType)
        this.mIcons['checkAll']    := 'HICON:*' LoadPicture('Icons.dll', 'Icon35 w32', &imageType)
        this.mIcons['selectAll']   := 'HICON:*' LoadPicture('Icons.dll', 'Icon37 w32', &imageType)
        this.mIcons['clear']       := 'HICON:*' LoadPicture('Icons.dll', 'Icon38 w32', &imageType)
        this.mIcons['g?']          := 'HICON:*' LoadPicture('Icons.dll', 'Icon24 w48', &imageType)
        this.mIcons['gI']          := 'HICON:*' LoadPicture('Icons.dll', 'Icon25 w48', &imageType)
        this.mIcons['gX']          := 'HICON:*' LoadPicture('Icons.dll', 'Icon26 w48', &imageType)
        this.mIcons['g!']          := 'HICON:*' LoadPicture('Icons.dll', 'Icon36 w48', &imageType)
        this.mIcons['gMenuX']      := 'HICON:*' LoadPicture('Icons.dll', 'Icon15 w32', &imageType)    
        this.mIcons['gMenuGitHub'] := 'HICON:*' LoadPicture('Icons.dll', 'Icon20 w32', &imageType)
        this.mIcons['gMenuPayPal'] := 'HICON:*' LoadPicture('Icons.dll', 'Icon21 w32', &imageType)
        this.mIcons['exe']         := 'HICON:*' LoadPicture('Icons.dll', 'Icon28 w48', &imageType)
        this.mIcons['window']      := 'HICON:*' LoadPicture('Icons.dll', 'Icon29 w48', &imageType)
        this.mIcons['tMenuGray']   := 'HICON:*' LoadPicture('Icons.dll', 'Icon27 w' this.m['mUser']['trayMenuIconSize'], &imageType)
        this.mIcons['tMenuTrans']  := 'HICON:*' LoadPicture('Icons.dll', 'Icon16 w' this.m['mUser']['trayMenuIconSize'], &imageType)

        ;==============================================
        this.mThemeIcons := Map()
        this.aThemeName := Array()
        
        Loop files, 'Themes\*', 'D' 
        {
            mThemeIcons := Map(), iconMissing := false
            SplitPath(A_LoopFilePath, &themeName)

            for _, iconName in this.aThemeIcons { 
                if FileExist('Themes\' themeName '\' iconName '.ico') {
                    mThemeIcons[iconName] := 'HICON:*' LoadPicture('Themes\' themeName '\' iconName '.ico', 'w' this.m['mUser']['trayMenuIconSize'], &imageType)
                } 
                else {
                    iconMissing := true
                    break
                }
            }

            if (!iconMissing) {
                this.aThemeName.Push(themeName)
                this.mThemeIcons[themeName] := mThemeIcons
                this.mThemeIcons[themeName]['gAbout']        := 'HICON:*' LoadPicture('Themes\' themeName '\main.ico', 'w64', &imageType)
                this.mThemeIcons[themeName]['gMainAbout']    := 'HICON:*' LoadPicture('Themes\' themeName '\about.ico', 'w32', &imageType)
                this.mThemeIcons[themeName]['gMainSettings'] := 'HICON:*' LoadPicture('Themes\' themeName '\settings.ico', 'w32', &imageType)
                this.mThemeIcons[themeName]['gMainEdit']     := 'HICON:*' LoadPicture('Themes\' themeName '\edit.ico', 'w32', &imageType)
                this.mThemeIcons[themeName]['gMainProfile']  := 'HICON:*' LoadPicture('Themes\' themeName '\profile.ico', 'w32', &imageType)                
                this.mThemeIcons[themeName]['gMenuEdit']     := 'HICON:*' LoadPicture('Themes\' themeName '\edit.ico', 'w32', &imageType)
                this.mThemeIcons[themeName]['gMenuFolder']   := 'HICON:*' LoadPicture('Themes\' themeName '\folder.ico', 'w32', &imageType)
                this.mThemeIcons[themeName]['gMenuAbout']    := 'HICON:*' LoadPicture('Themes\' themeName '\about.ico', 'w32', &imageType)
                this.mThemeIcons[themeName]['gMenuSettings'] := 'HICON:*' LoadPicture('Themes\' themeName '\settings.ico', 'w32', &imageType)
                this.mThemeIcons[themeName]['trayMain']      := 'HICON:*' LoadPicture('Themes\' themeName '\main.ico', 'w32', &imageType)
                this.mThemeIcons[themeName]['trayLoading']   := 'HICON:*' LoadPicture('Themes\' themeName '\loading.ico', 'w32', &imageType)
                this.mThemeIcons[themeName]['notifStatus']   := 'HICON:*' LoadPicture('Themes\' themeName '\main.ico', 'w64', &imageType)                
            }
        }

        ;==============================================
        this.aThemeName.Push('Reptilian')
        this.mThemeIcons['Reptilian'] := Map()
        this.mThemeIcons['Reptilian']['gAbout']        := 'HICON:*' LoadPicture('Icons.dll', 'Icon1 w64', &imageType)
        this.mThemeIcons['Reptilian']['gMainAbout']    := 'HICON:*' LoadPicture('Icons.dll', 'Icon5 w32', &imageType)
        this.mThemeIcons['Reptilian']['gMainSettings'] := 'HICON:*' LoadPicture('Icons.dll', 'Icon6 w32', &imageType)
        this.mThemeIcons['Reptilian']['gMainEdit']     := 'HICON:*' LoadPicture('Icons.dll', 'Icon8 w32', &imageType)
        this.mThemeIcons['Reptilian']['gMainProfile']  := 'HICON:*' LoadPicture('Icons.dll', 'Icon10 w32', &imageType)        
        this.mThemeIcons['Reptilian']['gMenuEdit']     := 'HICON:*' LoadPicture('Icons.dll', 'Icon8 w32', &imageType)
        this.mThemeIcons['Reptilian']['gMenuFolder']   := 'HICON:*' LoadPicture('Icons.dll', 'Icon9 w32', &imageType)
        this.mThemeIcons['Reptilian']['gMenuAbout']    := 'HICON:*' LoadPicture('Icons.dll', 'Icon5 w32', &imageType)
        this.mThemeIcons['Reptilian']['gMenuSettings'] := 'HICON:*' LoadPicture('Icons.dll', 'Icon6 w32', &imageType)
        this.mThemeIcons['Reptilian']['trayMain']      := 'HICON:*' LoadPicture('Icons.dll', 'Icon1 w32', &imageType)
        this.mThemeIcons['Reptilian']['trayLoading']   := 'HICON:*' LoadPicture('Icons.dll', 'Icon2 w32', &imageType)
        this.mThemeIcons['Reptilian']['notifStatus']   := 'HICON:*' LoadPicture('Icons.dll', 'Icon1 w64', &imageType)

        for _, iconName in this.aThemeIcons
            this.mThemeIcons['Reptilian'][iconName] := 'HICON:*' LoadPicture('Icons.dll', 'Icon' A_Index ' w' this.m['mUser']['trayMenuIconSize'], &imageType)

        if this.mInitValues['mUser'].Has('themeName') && this.HasVal(this.mInitValues['mUser']['themeName'], this.aThemeName)
            this.m['mUser']['themeName'] := this.mInitValues['mUser']['themeName']
        else
            this.m['mUser']['themeName'] := this.mInitValues['mUser']['themeName'] := 'Reptilian'

        ;==============================================
        A_TrayMenu.Delete
        A_TrayMenu.Add('Events', this.trayMenuEvents := Menu())
        A_TrayMenu.Add('Events Manager', this.gMain_Show.Bind(this))
        A_TrayMenu.Add()
        A_TrayMenu.Add('Select Profile', this.trayMenuSelectProfile := Menu())
        A_TrayMenu.Add('All Events Enable', this.TrayMenuProfile_ItemClick.Bind(this))
        A_TrayMenu.Add('All Events Disable', this.TrayMenuProfile_ItemClick.Bind(this))
        A_TrayMenu.Add()
        A_TrayMenu.Add('Tools', this.trayMenuTools := Menu())
        A_TrayMenu.Add('Open Script Folder', (*) => Run(A_WorkingDir))
        A_TrayMenu.Add('Edit Script', (*) => Edit())
        A_TrayMenu.Add('Settings', this.gSettings_Show.Bind(this))
        A_TrayMenu.Add('About', this.gAbout_Show.Bind(this))
        A_TrayMenu.Add('Reload', (*) => Reload())
        A_TrayMenu.Add('Exit', (*) => ExitApp())

        aKeyToDelete := Array()
        
        for key, mTool in this.mTools {
            if FileExist(mTool['path'])
                this.trayMenuTools.Add(mTool['name'], this.RunTool.Bind(this, mTool['path']))
            else
                aKeyToDelete.Push(key)
        }

        for _, key in aKeyToDelete
            this.mTools.Delete(key)
            
        ;==============================================
        for _, value in ['exit', 'reload', 'about', 'tools', 'settings', 'events']
            A_TrayMenu.SetIcon(value, this.mThemeIcons[this.m['mUser']['themeName']][value],, this.m['mUser']['trayMenuIconSize'])

        A_TrayMenu.SetIcon('Events Manager', this.mThemeIcons[this.m['mUser']['themeName']]['main'],, this.m['mUser']['trayMenuIconSize'])
        A_TrayMenu.SetIcon('Select Profile', this.mThemeIcons[this.m['mUser']['themeName']]['profile'],, this.m['mUser']['trayMenuIconSize'])
        A_TrayMenu.SetIcon('Edit Script', this.mThemeIcons[this.m['mUser']['themeName']]['edit'],, this.m['mUser']['trayMenuIconSize'])
        A_TrayMenu.SetIcon('Open Script Folder', this.mThemeIcons[this.m['mUser']['themeName']]['folder'],, this.m['mUser']['trayMenuIconSize'])

        for _, value in ['All Events Enable', 'All Events Disable']
            A_TrayMenu.SetIcon(value, this.mIcons['tMenuTrans'],, this.m['mUser']['trayMenuIconSize'])

        for _, mTool in this.mTools {
            if mTool.Has('iconPath') && FileExist(mTool['iconPath'])
                this.trayMenuTools.SetIcon(mTool['name'], 'HICON:*' LoadPicture(mTool['iconPath'], 'w' this.m['mUser']['trayMenuIconSize'], &imageType),, this.m['mUser']['trayMenuIconSize'])
            else
                this.trayMenuTools.SetIcon(mTool['name'], this.mIcons['tMenuTrans'],, this.m['mUser']['trayMenuIconSize'])
        }

        ; Hotkeys =====================================
        HotIfWinActive(this.gMainTitle ' ahk_class AutoHotkeyGUI')
        Hotkey('Delete', this.gMain_lv_event_DeleteEvent.Bind(this))
        Hotkey('^a', this.gMain_lv_event_SelectAll.Bind(this))

        ;============================================== 
        this.callbackNotifyIcon := this.AHK_NOTIFYICON.Bind(this)
        this.callbackDeviceChange := this.WM_DEVICECHANGE.Bind(this)
        this.wmi := ComObjGet('winmgmts:')
        this.sink := ComObject('WbemScripting.SWbemSink')
        this.hWinEventHook := 0, this.hookProcAdr := 0
        OnError(this.OnError.Bind(this))
        this.SetMonitoringValues(this.mInitValues, '__New')
        OnExit(this.OnExit.Bind(this))
    }

    ;=============================================================================================

    static OnExit(exitReason, exitCode)
    {
        (this.debugMode && IsSet(Debug)) ? Debug(exitReason, 'exitReason') : ''

        this.GUI_IfExistReturn_Destroy(['gMain', 'gEvent', 'gSettings', 'gAbout'], 'destroy')
        this.TrayMenu_Enable_Disable('disable')

        for _, value in ['wmi', 'winEvent', 'deviceChange', 'trayClick']
            this.SetMonitoring%value%(0)

        this.WriteValuesToJSON()
    }

    ;============================================================================================

    static OnError(thrown, mode)
    {
        (this.debugMode && IsSet(Debug)) ? Debug(thrown, 'thrown') : ''

        Thread('NoTimers')
        this.GUI_IfExistReturn_Destroy(['gMain', 'gEvent', 'gSettings', 'gAbout'], 'destroy')
        this.TrayMenu_Enable_Disable('disable')
        A_TrayMenu.SetIcon('All Events Enable', this.mIcons['tMenuGray'],, this.m['mUser']['trayMenuIconSize'])
        A_TrayMenu.SetIcon('All Events Disable', this.mIcons['tMenuGray'],, this.m['mUser']['trayMenuIconSize'])
        
        for _, value in ['wmi', 'winEvent', 'deviceChange', 'trayClick']
            this.SetMonitoring%value%(0)

        this.Notify(this.gErrorTitle, 'Monitoring has been turned off.`nExiting or reloading the script is required.`nClick here to reload.', this.mIcons['gX'],,, 0, (*) => Reload())
    }

    ;============================================================================================

    static WriteValuesToJSON()
    {
        m := Map('mUser', Map())
        m['mEvents'] := this.Copy_Add_mEvents(this.m['mEvents'])
        
        for key, value in this.m['mUser']
            m['mUser'][key] := value  

        strSettingsJSON := JSON.stringify(m, expandlevel := unset, space := "  ")
        try FileRecycle('Settings.json')
        fileSettingsJSON := FileOpen('Settings.json', 'w', 'UTF-8')
        fileSettingsJSON.Write(strSettingsJSON)        

        ;==============================================
        for eventName, mEvent in m['mEvents']
            this.mEventsBackup['mEvents'][eventName] := mEvent

        strEventsBakJSON := JSON.stringify(this.mEventsBackup, expandlevel := unset, space := "  ")
        try FileRecycle('EventsBackup.json')
        fileEventsBakJSON := FileOpen('EventsBackup.json', 'w', 'UTF-8')
        fileEventsBakJSON.Write(strEventsBakJSON)
        fileEventsBakJSON.Close()      
    }    

    ;=============================================================================================

    static JSONtoMapObject(filePath)
    {
        SplitPath(filePath,,,, &profileName)
        mProfile := JSON.parse( FileOpen(filePath, 'r', 'UTF-8').Read() , keepbooltype := false, as_map := true) ; mProfile CaseSense is On.

        for _, value in ['mEvents', 'mUser']
            if !mProfile.Has(value)
                mProfile[value] := Map()

        if !mProfile['mUser'].Has('mWinEventsStates')
            mProfile['mUser']['mWinEventsStates'] := Map()

        for winEventName, value in this.mWinEvents {
            if this.HasVal(winEventName, this.mWinEventPresets[this.mDefault['winEventPreset']])
                mProfile['mUser']['mWinEventsStates'][winEventName] := 1
            else
                mProfile['mUser']['mWinEventsStates'][winEventName] := 0  
        }  

        for winEventName, state in mProfile['mUser']['mWinEventsStates']
            if !mProfile['mUser']['mWinEventsStates'].Has(winEventName) || !this.HasVal(state, [0, 1])
                mProfile['mUser']['mWinEventsStates'][winEventName] := 0

        for _, value in ['periodWMIprocess', 'delayDeviceChange']
            if !mProfile['mUser'].Has(value) || !this.HasVal(mProfile['mUser'][value], this.a%value%)
                mProfile['mUser'][value] := this.mDefault[value]
        
        ;==============================================
        aInvalidEvent := Array()

        for eventName, mEvent in mProfile['mEvents'] 
        {
            if (!this.GetEventType(mEvent) || !this.IsValidEventName(eventName)) {                
                aInvalidEvent.Push(eventName)
                continue
            }

            for _, value in ['state', 'critical', 'log', 'notifInfo', 'notifStatus']
                if !mEvent.Has(value) || !this.HasVal(mEvent[value], this.a%value%)
                    mEvent[value] := this.mDefault[value] 

            for _, value in ['soundCreated', 'soundTerminated']
                if !mEvent.Has(value) || !this.HasVal(mEvent[value], this.aSounds)
                    mEvent[value] := this.mDefault[value]

            if !mEvent.Has('function') || !this.IsValidFunctionName(mEvent['function'])
                mEvent['function'] := ''            

            switch {
                case mEvent.Has('window'):
                {
                    for _, value in ['winTitle', 'winClass', 'processName', 'processPath']
                        if !mEvent['window'].Has(value)
                            mEvent['window'][value] := ''
                                      
                    if (!mEvent['window']['winTitle'] && !mEvent['window']['winClass'] && !mEvent['window']['processName'] && !mEvent['window']['processPath']) {
                        aInvalidEvent.Push(eventName)
                        continue
                    }    
                                       
                    if !mEvent.Has('mode') || !this.HasVal(mEvent['mode'], this.aModeWindow)
                        mEvent['mode'] := this.mDefault['modeWindow']

                    if !mEvent.Has('monitoring') || !this.HasVal(mEvent['monitoring'], this.aMonitoringWindow, caseSensitive := false)
                        mEvent['monitoring'] := this.mDefault['monitoringWindow']

                    if (!mEvent.Has('period')) {
                        if mEvent['monitoring'] == 'timer'
                            mEvent['period'] := this.mDefault['periodTimerWindow']
                        
                        if mEvent['monitoring'] == 'winEvent'
                            mEvent['period'] := this.mDefault['delayWinEvent']
                    }

                    if mEvent['monitoring'] == 'timer' && !this.HasVal(mEvent['period'], this.aPeriodTimerWindow)
                        mEvent['period'] := this.mDefault['periodTimerWindow']

                    if mEvent['monitoring'] == 'winEvent' && !this.HasVal(mEvent['period'], this.aDelayWinEvent)
                        mEvent['period'] := this.mDefault['delayWinEvent']

                    for _, value in ['winTitleMatchMode', 'detectHiddenWindows', 'winActive', 'winMinMax']
                        if !mEvent['window'].Has(value) || !this.HasVal(mEvent['window'][value], this.a%value%)
                            mEvent['window'][value] := this.mDefault[value]                        
                }

                case mEvent.Has('process'):
                {
                    for _, value in ['processName', 'processPath']
                        if !mEvent['process'].Has(value)
                            mEvent['process'][value] := ''
                                      
                    if (!mEvent['process']['processName'] && !mEvent['process']['processPath']) {
                        aInvalidEvent.Push(eventName)
                        continue
                    }                       

                    if !mEvent.Has('mode') || !this.HasVal(mEvent['mode'], this.aModeProcess)
                        mEvent['mode'] := this.mDefault['modeProcess']

                    if !mEvent.Has('monitoring') || !this.HasVal(mEvent['monitoring'], this.aMonitoringProcess, caseSensitive := false)
                        mEvent['monitoring'] := this.mDefault['monitoringProcess']

                    if !mEvent.Has('period') {
                        if mEvent['monitoring'] == 'timer'
                            mEvent['period'] := this.mDefault['periodTimerProcess']

                        if mEvent['monitoring'] == 'wmi'
                            mEvent['period'] := this.mDefault['periodWMIprocess']
                    }
                    
                    if mEvent['monitoring'] == 'timer' && !this.HasVal(mEvent['period'], this.aPeriodTimerProcess)
                        mEvent['period'] := this.mDefault['periodTimerProcess']

                    if mEvent['monitoring'] == 'wmi' && !this.HasVal(mEvent['period'], this.aPeriodWMIprocess)
                        mEvent['period'] := this.mDefault['periodWMIprocess']
                }                

                case mEvent.Has('device'):
                {
                    for _, value in ['deviceName', 'deviceId']
                        if !mEvent['device'].Has(value)
                            mEvent['device'][value] := ''
                                      
                    if (!mEvent['device']['deviceName'] && !mEvent['device']['deviceId']) {
                        aInvalidEvent.Push(eventName)
                        continue
                    }                     
                    
                    if !mEvent.Has('mode') || !this.HasVal(mEvent['mode'], this.aModeDevice)
                        mEvent['mode'] := this.mDefault['modeDevice']

                    if !mEvent.Has('monitoring') || !this.HasVal(mEvent['monitoring'], this.aMonitoringDevice, caseSensitive := false)
                        mEvent['monitoring'] := this.mDefault['monitoringDevice']

                    if (!mEvent.Has('period')) {
                        if mEvent['monitoring'] == 'timer'
                            mEvent['period'] := this.mDefault['periodTimerDevice']

                        if mEvent['monitoring'] == 'deviceChange'
                            mEvent['period'] := this.mDefault['delayDeviceChange']
                    }

                    if mEvent['monitoring'] == 'timer' && !this.HasVal(mEvent['period'], this.aPeriodTimerDevice)
                        mEvent['period'] := this.mDefault['periodTimerDevice']

                    if mEvent['monitoring'] == 'deviceChange' && !this.HasVal(mEvent['period'], this.aDelayDeviceChange)
                        mEvent['period'] := this.mDefault['delayDeviceChange']
                }
            }
        }
           
        for _, eventName in aInvalidEvent
            try mProfile['mEvents'].Delete(eventName)        

        return mProfile
    }

    ;============================================================================================

    static SetMonitoringWMI(state)
    {
        if (state = 1 && this.monitoringWMIProcess) || (state = 0 && !this.monitoringWMIProcess)
            return

        if (state = 1) {
            this.aEventsProcess := Array()
            
            for _, mEvent in this.m['mEvents']
                if mEvent.Has('process') && mEvent['monitoring'] == 'wmi' && mEvent['state'] = 1
                    this.aEventsProcess.Push(mEvent['process']['processName'])

            strWMIQuery := 'Within ' Format('{:g}', this.m['mUser']['periodWMIprocess']/ 1000) ' Where TargetInstance ISA `'Win32_Process`' AND ('

            for index, process in this.aEventsProcess
                strWMIQuery .= (index > 1 ? ' OR ' : '') 'TargetInstance.Name = `'' process '`'' (index = this.aEventsProcess.Length ? ')' : '')

            (this.debugMode && IsSet(Debug)) ? Debug(this.aEventsProcess, 'this.aEventsProcess') : ''
            (this.debugMode && IsSet(Debug)) ? Debug(strWMIQuery, 'strWMIQuery') : ''

            ComObjConnect(this.sink, this)
            this.wmi.ExecNotificationQueryAsync(this.sink, 'SELECT * FROM __InstanceCreationEvent ' . strWMIQuery)
            this.wmi.ExecNotificationQueryAsync(this.sink, 'SELECT * FROM __InstanceDeletionEvent ' . strWMIQuery)
            this.monitoringWMIProcess := true
        } 
        else {
            ComObjConnect(this.sink)
            this.sink.Cancel()
            this.monitoringWMIProcess := false
        }
    }

    ;=============================================================================================

    static OnObjectReady(SWbemObjectEx, WbemAsyncContext, thisSWbemSink)
    {
        if (this.isCriticalMethodExecuting) {
            (this.debugMode && IsSet(Debug)) ? Debug('isCriticalMethodExecuting_OnObjectReady', 'str') : ''
            SetTimer(this.OnObjectReady.Bind(this, SWbemObjectEx, WbemAsyncContext, thisSWbemSink), -75)
            return
        }

        switch SWbemObjectEx.Path_.Class {
            case '__InstanceCreationEvent': processStatus := 'created'
            case '__InstanceDeletionEvent': processStatus := 'terminated'
        }

        ti := SWbemObjectEx.TargetInstance
        (this.debugMode && IsSet(Debug)) ? Debug(TI.Name, 'TI.Name') : ''

        for eventName, mEvent in this.m['mEvents']
            if mEvent.Has('process') && mEvent['monitoring'] == 'wmi' && mEvent['state'] = 1 && mEvent['process']['processName'] = ti.Name
                SetTimer(this.ProcessStatusFinderUpdater.Bind(this, mEvent, eventName), -mEvent['period'])
    }

    ;=============================================================================================

    static ProcessStatusFinderUpdater(mEvent, eventName)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        aEventCreated := Array(), aEventTerminated := Array(), aProcessCreated := Array(), aProcessTerminated := Array()
        aProcessFinder := this.ProcessFinder(mEvent['process']['processName'], mEvent['process']['processPath'])

        for _, process in aProcessFinder {
            pidFound := false
            for index, proc in mEvent['processMatch'] {
                if process['pid'] = proc['pid'] {
                    pidFound := true
                    break
                }
            }

            if !pidFound {
                switch {
                    case RegExMatch(mEvent['mode'], '^(1|3)$'): mEvent['processMatch'].Push(process), aEventCreated.Push(process), mEvent['status'] := 'created'
                    case mEvent['mode'] = 2:                    mEvent['processMatch'].Push(process), aProcessCreated.Push(process)
                }
            }            
        }

        ;==============================================
        switch mEvent['mode'] {
            case 1:
            {
                for index, process in mEvent['processMatch']
                    if !ProcessExist(process['pid'])
                        mEvent['processMatch'].RemoveAt(index), aEventTerminated.Push(process)

                if !mEvent['processMatch'].Length
                    mEvent['status'] := 'terminated'
            }            

            ;==============================================
            case 2:
            {    
                for index, process in mEvent['processMatch']
                    if !ProcessExist(process['pid'])
                        mEvent['processMatch'].RemoveAt(index), aProcessTerminated.Push(process)

                if mEvent['status'] == 'terminated' && mEvent['processMatch'].Length
                    aEventCreated.Push(mEvent['processMatch'][1]), mEvent['status'] := 'created'

                if mEvent['status'] == 'created' && !mEvent['processMatch'].Length && aProcessTerminated.Length
                    aEventTerminated.Push(aProcessTerminated[1]), mEvent['status'] := 'terminated'
            }

            ;==============================================
            case 3:
            {
                for index, process in mEvent['processMatch']
                    if !ProcessExist(process['pid'])
                        mEvent['processMatch'].RemoveAt(index), aProcessTerminated.Push(process)

                if mEvent['status'] == 'created' && !mEvent['processMatch'].Length && aProcessTerminated.Length
                    aEventTerminated.Push(aProcessTerminated[1]), mEvent['status'] := 'terminated'
            }
        }

        ;==============================================
        for _, status in ['created', 'terminated'] {
            if aEvent%status%.Length {
                for index, mValue in aEvent%status% {
                    for _, prop in ['period', 'mode', 'monitoring', 'function', 'critical', 'log', 'notifInfo', 'notifStatus']
                        mValue[prop] := mEvent[prop]

                    mValue['eventName'] := eventName
                    mValue['eventType'] := this.GetEventType(mEvent)
                    mValue['status'] := status
                    this.gMain_lv_event_ModifyIcon(eventName, mEvent['processMatch'].Length) 

                    if mEvent['log']
                        this.LogEventToFile(this.DisplayObj(mValue)) 

                    if mEvent['notifInfo']
                        this.Notify(eventName, this.DisplayObj(mValue), mValue['processPath'],, 'topLeft', '10000')

                    if mEvent['notifStatus']             
                        this.Notify(this.scriptName, 'Event: ' eventName ' ' StrTitle(status), this.mThemeIcons[this.m['mUser']['themeName']]['notifStatus'])
                    
                    SetTimer(this.Sound.Bind(this, this.mSounds[mEvent['sound' StrTitle(status)]]), -1)     
                    
                    if (IsSet(%mEvent['function'] '_' status%)) {
                        if mEvent['critical']
                            %mEvent['function'] '_' status%(mValue)  
                        else                   
                            SetTimer(%mEvent['function'] '_' status%.Bind(mValue), -1)  
                    }                    
                }
            }
        }

        this.SetCritical_Off(wasNotCritical)
    }

    ;=============================================================================================
    ; Based on WTSEnumProcesses by SKAN.    https://www.autohotkey.com/boards/viewtopic.php?t=4365
    ; Returns an array containing map objects with all existing processes that match the specified parameters. If there are no matching processes, an empty array is returned.
    static ProcessFinder(processName:='', processPath:='')
    {      
        aProcessFinder := Array(), tPtr := 0, pPtr := 0, nTTL := 0

        if not DllCall("Wtsapi32\WTSEnumerateProcesses", "Ptr", 0, "Int", 0, "Int", 1, "PtrP", &pPtr, "PtrP", &nTTL)
            return aProcessFinder 

        tPtr := pPtr

        if !processName && processPath
            SplitPath(processPath, &processName)    
        
        Loop nTTL 
        {
            mProcessMatch := Map(), strProcessPath := ''
            strPid := NumGet(tPtr + 4, "UInt")
            strProcessname := StrGet(NumGet(tPtr + 8, "UPtr"))
            tPtr += (A_PtrSize = 4 ? 16 : 24) ; sizeof(WTS_PROCESS_INFO)

            if !strPid 
            || !RegExMatch(strProcessname, 'i)\.exe$')
            || processName && (processName != strProcessName)
                continue          

            try strProcessPath := ProcessGetPath(strPid)
            if processPath && (processPath != strProcessPath)
                continue
 
            strCmdLine := this.GetCommandLine(strPid,, &imagePath:='') 
            strElevated := this.IsProcessElevated(strPid)
            ; strPPID := removed, ProcessGetParent was slow.
            
            for _, value in ['processName', 'processPath', 'pid', 'cmdLine', 'elevated']
                mProcessMatch[value] := %'str' value% 
            
            aProcessFinder.Push(mProcessMatch)
        }

        DllCall("Wtsapi32\WTSFreeMemory", "Ptr", pPtr)
        return aProcessFinder
    }

    ;============================================================================================
    ; IsProcessElevated by jNizM.    https://github.com/jNizM/ahk-scripts-v2/blob/main/src/ProcessThreadModule/IsProcessElevated.ahk
    static IsProcessElevated(ProcessID)
    {
        static INVALID_HANDLE_VALUE              := -1
        static PROCESS_QUERY_INFORMATION         := 0x0400
        static PROCESS_QUERY_LIMITED_INFORMATION := 0x1000
        static TOKEN_QUERY                       := 0x0008
        static TOKEN_QUERY_SOURCE                := 0x0010
        static TokenElevation                    := 20

        if !ProcessID
            return 0

        hProcess := DllCall("OpenProcess", "UInt", PROCESS_QUERY_INFORMATION, "Int", False, "UInt", ProcessID, "Ptr")
        if (!hProcess) || (hProcess = INVALID_HANDLE_VALUE)
        {
            hProcess := DllCall("OpenProcess", "UInt", PROCESS_QUERY_LIMITED_INFORMATION, "Int", False, "UInt", ProcessID, "Ptr")
            if (!hProcess) || (hProcess = INVALID_HANDLE_VALUE)
                return 0
        }

        if !(DllCall("advapi32\OpenProcessToken", "Ptr", hProcess, "UInt", TOKEN_QUERY | TOKEN_QUERY_SOURCE, "Ptr*", &hToken := 0))
        {
            DllCall("CloseHandle", "Ptr", hProcess)
            return 0
        }

        if !(DllCall("advapi32\GetTokenInformation", "Ptr", hToken, "Int", TokenElevation, "UInt*", &IsElevated := 0, "UInt", 4, "UInt*", &Size := 0))
        {
            DllCall("CloseHandle", "Ptr", hToken)
            DllCall("CloseHandle", "Ptr", hProcess)
            return 0
        }

        DllCall("CloseHandle", "Ptr", hToken)
        DllCall("CloseHandle", "Ptr", hProcess)
        return IsElevated
    }

    ;=============================================================================================

    ; GetCommandLine by teadrinker (Based on Sean and SKAN v1 code) https://www.autohotkey.com/boards/viewtopic.php?p=526409#p526409 https://www.autohotkey.com/board/topic/15214-getcommandline/
    static GetCommandLine(PID:="", setDebugPrivilege := false, &imagePath?)
    {
        static flags := (PROCESS_QUERY_INFORMATION := 0x400) | (PROCESS_VM_READ := 0x10), STATUS_SUCCESS := 0, setDebug := false

        if (setDebugPrivilege && !setDebug) {
            RunAsAdmin()
            if !(setDebug := !SetPrivilege("SeDebugPrivilege"))
                Return
        }
        hProc := DllCall("OpenProcess", "UInt", flags, "Int", 0, "UInt", PID, "Ptr")
        (A_Is64bitOS && DllCall("IsWow64Process", "Ptr", hProc, "UIntP", &IsWow64 := 0))
        if (!A_Is64bitOS || IsWow64)
            PtrSize := 4, PtrType := "UInt", pPtr := "UIntP", offsetCMD := 0x40
        else
            PtrSize := 8, PtrType := "Int64", pPtr := "Int64P", offsetCMD := 0x70

        hModule := DllCall("GetModuleHandle", "str", "Ntdll", "Ptr")
        failed := ""
        if (A_PtrSize < PtrSize) {    ; script 32, dest proc 64
            if !QueryInformationProcess := DllCall("GetProcAddress", "Ptr", hModule, "AStr", "NtWow64QueryInformationProcess64", "Ptr")
                failed := "NtWow64QueryInformationProcess64"
            if !ReadProcessMemory := DllCall("GetProcAddress", "Ptr", hModule, "AStr", "NtWow64ReadVirtualMemory64", "Ptr")
                failed := "NtWow64ReadVirtualMemory64"
            info := 0, szPBI := 48, offsetPEB := 8
        }
        else {
            if !QueryInformationProcess := DllCall("GetProcAddress", "Ptr", hModule, "AStr", "NtQueryInformationProcess", "Ptr")
                failed := "NtQueryInformationProcess"
            ReadProcessMemory := "ReadProcessMemory"
            if (A_PtrSize > PtrSize)    ; script 64, dest proc 32
                info := 26, szPBI := 8, offsetPEB := 0
            else    ; script and dest proc the same bitness
                info := 0, szPBI := PtrSize * 6, offsetPEB := PtrSize
        }
        if failed {
            DllCall("CloseHandle", "Ptr", hProc)
            Return
        }
        PBI := Buffer(48, 0)
        if DllCall(QueryInformationProcess, "Ptr", hProc, "UInt", info, "Ptr", PBI, "UInt", szPBI, "UIntP", &bytes := 0) != STATUS_SUCCESS {
            DllCall("CloseHandle", "Ptr", hProc)
            Return
        }
        pPEB := NumGet(PBI.Ptr + offsetPEB, PtrType)
        DllCall(ReadProcessMemory, "Ptr", hProc, PtrType, pPEB + PtrSize * 4, pPtr, &pRUPP := 0, PtrType, PtrSize, "UIntP", &bytes := 0)
        DllCall(ReadProcessMemory, "Ptr", hProc, PtrType, pRUPP + offsetCMD, "UShortP", &szCMD := 0, PtrType, 2, "UIntP", &bytes)
        DllCall(ReadProcessMemory, "Ptr", hProc, PtrType, pRUPP + offsetCMD + PtrSize, pPtr, &pCMD := 0, PtrType, PtrSize, "UIntP", &bytes)

        buff := Buffer(szCMD, 0)
        DllCall(ReadProcessMemory, "Ptr", hProc, PtrType, pCMD, "Ptr", buff, PtrType, szCMD, "UIntP", &bytes)
        CMD := StrGet(buff, "UTF-16")

        if IsSet(imagePath) {
            DllCall(ReadProcessMemory, "Ptr", hProc, PtrType, pRUPP + offsetCMD - PtrSize * 2, "UShortP", &szPATH := 0, PtrType, 2, "UIntP", &bytes)
            DllCall(ReadProcessMemory, "Ptr", hProc, PtrType, pRUPP + offsetCMD - PtrSize, pPtr, &pPATH := 0, PtrType, PtrSize, "UIntP", &bytes)

            buff := Buffer(szPATH, 0)
            DllCall(ReadProcessMemory, "Ptr", hProc, PtrType, pPATH, "Ptr", buff, PtrType, szPATH, "UIntP", &bytes)
            imagePath := StrGet(buff, "UTF-16") . (IsWow64 ? " *32" : "")
        }
        DllCall("CloseHandle", "Ptr", hProc)
        Return CMD

        ;==============================================

        SetPrivilege(privilege, enable := true) {
            static TOKEN_ADJUST_PRIVILEGES := 0x20, SE_PRIVILEGE_ENABLED := 0x2, SE_PRIVILEGE_REMOVED := 0x4
            DllCall("Advapi32\OpenProcessToken", "Ptr", -1, "UInt", TOKEN_ADJUST_PRIVILEGES, "PtrP", &token := 0)
            DllCall("Advapi32\LookupPrivilegeValue", "Ptr", 0, "Str", privilege, "Int64P", &luid := 0)
            TOKEN_PRIVILEGES := Buffer(16, 0)
            NumPut("UInt", 1, TOKEN_PRIVILEGES)
            NumPut("Int64", luid, TOKEN_PRIVILEGES, 4)
            NumPut("UInt", enable ? SE_PRIVILEGE_ENABLED : SE_PRIVILEGE_REMOVED, TOKEN_PRIVILEGES, 12)
            DllCall("Advapi32\AdjustTokenPrivileges", "Ptr", token, "Int", !enable, "Ptr", TOKEN_PRIVILEGES, "UInt", 0, "Ptr", 0, "Ptr", 0)
            res := A_LastError
            DllCall("CloseHandle", "Ptr", token)
            Return res    ; success — 0
        }

        RunAsAdmin() {
            isRestarted := !!RegExMatch(DllCall("GetCommandLine", "str"), " /restart(?!\S)")
            if !(A_IsAdmin || isRestarted) {
                try Run '*RunAs ' . (A_IsCompiled ? '"' . A_ScriptFullPath . '" /restart'
                    : '"' . A_AhkPath . '" /restart "' . A_ScriptFullPath . '"')
                ExitApp
            }
        }
    }

    ;============================================================================================

    static SetMonitoringWinEvent(state)
    {
        if (state = 1 && this.monitoringWinEvent) || (state = 0 && !this.monitoringWinEvent)
            return

        if (state = 1) {
            for _, value in ['winEventMinHex', 'winEventMaxHex', 'aWinEventExclude']
                (this.debugMode && IsSet(Debug)) ? Debug(this.%value%, 'this.' value) : ''
            
            this.hookProcAdr := CallbackCreate(this.CaptureWinEvent.Bind(this), 'F', 7)
            this.hWinEventHook := this.SetWinEventHook(this.winEventMinHex, this.winEventMaxHex, 0, this.hookProcAdr, 0, 0, this.m['mUser']['dwFlagsWinEvent'])
            this.monitoringWinEvent := true
        } 
        else {
            (this.hookProcAdr) ? (CallbackFree(this.hookProcAdr), this.hookProcAdr := 0) : ''
            DllCall('UnhookWinEvent', 'Ptr', this.hWinEventHook), this.hWinEventHook := 0
            this.monitoringWinEvent := false
        }
    }

    ;=============================================================================================

    static SetWinEventHook(eventMin, eventMax, hmodWinEventProc, lpfnWinEventProc, idProcess, idThread, dwFlags)
    {
        return DllCall('SetWinEventHook', 'Uint', eventMin, 'Uint', eventMax, 'Uint', hmodWinEventProc, 'Uint', lpfnWinEventProc, 'Uint', idProcess, 'Uint', idThread, 'Uint', dwFlags)
    }

    ;=============================================================================================

    static CaptureWinEvent(hWinEventHook, event, *)
    {
        if this.HasVal(event, this.aWinEventExclude)
            return

        if (this.isCriticalMethodExecuting) {
            (this.debugMode && IsSet(Debug)) ? Debug('isCriticalMethodExecuting_CaptureWinEvent', 'str') : ''
            SetTimer(this.CaptureWinEvent.Bind(this, hWinEventHook, event), -75)
            return
        }

        ; (this.debugMode && IsSet(Debug)) ? Debug(event, 'event') : ''

        for eventName, mEvent in this.m['mEvents']
            if mEvent.Has('window') && mEvent['monitoring'] == 'winEvent' && mEvent['state'] = 1
                SetTimer(this.WindowStatusFinderUpdater.Bind(this, mEvent, eventName), -mEvent['period'])
    }

    ;=============================================================================================

    static WindowStatusFinderUpdater(mEvent, eventName)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        ; (this.debugMode && IsSet(Debug)) ? Debug(mEvent, 'mEvent') : ''

        dhwPrev := A_DetectHiddenWindows
        DetectHiddenWindows(mEvent['window']['detectHiddenWindows'])
        aEventCreated := Array(), aEventTerminated := Array(), aWindowCreated := Array(), aWindowTerminated := Array()
        
        aWindowFinder := this.WindowFinder(
            mEvent['window']['winTitle'], 
            mEvent['window']['winClass'], 
            mEvent['window']['processName'], 
            mEvent['window']['processPath'], 
            mEvent['window']['winTitleMatchMode'], 
            mEvent['window']['detectHiddenWindows'], 
            mEvent['window']['winActive'], 
            mEvent['window']['winMinMax']
        )

        switch { 
            case RegExMatch(mEvent['mode'], '^(4|5|6)$') || mEvent['window']['winActive'] = 1 || RegExMatch(mEvent['window']['winMinMax'], '^(0|1|-1)$'):
            {
                for _, status in ['created', 'terminated'] { 
                    for index, window in (status == 'created' ? aWindowFinder : mEvent['windowMatch']) {
                        winFound := false
                        for _, win in (status == 'created' ? mEvent['windowMatch'] : aWindowFinder) {
                            if window['winTitle']    == win['winTitle'] 
                            && window['winClass']    == win['winClass'] 
                            && window['id']          =  win['id']
                            && window['processName'] =  win['processName'] 
                            && window['processPath'] == win['processPath'] 
                            && window['pid']         =  win['pid'] {
                                winFound := true
                                break
                            }
                        }

                        if !winFound {
                            switch status {
                                case 'created':    mEvent['windowMatch'].Push(window), aWindowCreated.Push(window)
                                case 'terminated': mEvent['windowMatch'].RemoveAt(index), aWindowTerminated.Push(window)
                            }
                        }
                    }
                }
            }
            
            case RegExMatch(mEvent['mode'], '^(1|2|3)$'):
            {
                for _, window in aWindowFinder {
                    winIdFound := false
                    for index, win in mEvent['windowMatch'] {
                        if window['id'] = win['id'] {
                            winIdFound := true
                            break
                        }
                    }

                    if !winIdFound {
                        switch {
                            case RegExMatch(mEvent['mode'], '^(1|3)$'): mEvent['windowMatch'].Push(window), aEventCreated.Push(window), mEvent['status'] := 'created'
                            case mEvent['mode'] = 2:                    mEvent['windowMatch'].Push(window), aWindowCreated.Push(window)
                        }
                    }
                }

                for index, win in mEvent['windowMatch'] {
                    for _, window in aWindowFinder {
                        if window['id'] = win['id']
                        && (window['winTitle']   !== win['winTitle'] 
                        || window['winClass']    !== win['winClass'] 
                        || window['id']          !=  win['id']
                        || window['processName'] !=  win['processName'] 
                        || window['processPath'] !== win['processPath'] 
                        || window['pid']         !=  win['pid']) {
                            mEvent['windowMatch'].RemoveAt(index), mEvent['windowMatch'].Push(window)
                            break
                        }
                    }
                }
            }
        }

        ;==============================================
        switch {
            case mEvent['mode'] = 5 || mEvent['window']['winActive'] = 1 || RegExMatch(mEvent['window']['winMinMax'], '^(0|1|-1)$'):
            { 
                if mEvent['status'] == 'terminated' && mEvent['windowMatch'].Length
                    aEventCreated.Push(mEvent['windowMatch'][1]), mEvent['status'] := 'created'

                if mEvent['status'] == 'created' && !mEvent['windowMatch'].Length && aWindowTerminated.Length
                    aEventTerminated.Push(aWindowTerminated[1]), mEvent['status'] := 'terminated'
            }

            case mEvent['mode'] = 1:
            { 
                for index, window in mEvent['windowMatch']
                    if !WinExist('ahk_id ' window['id'])
                        mEvent['windowMatch'].RemoveAt(index), aEventTerminated.Push(window)

                if !mEvent['windowMatch'].Length
                    mEvent['status'] := 'terminated'
            }

            case mEvent['mode'] = 2:
            {   
                for index, window in mEvent['windowMatch']
                    if !WinExist('ahk_id ' window['id'])
                        mEvent['windowMatch'].RemoveAt(index), aWindowTerminated.Push(window)

                if mEvent['status'] == 'terminated' && mEvent['windowMatch'].Length
                    aEventCreated.Push(mEvent['windowMatch'][1]), mEvent['status'] := 'created'

                if mEvent['status'] == 'created' && !mEvent['windowMatch'].Length && aWindowTerminated.Length
                    aEventTerminated.Push(aWindowTerminated[1]), mEvent['status'] := 'terminated'
            }

            case mEvent['mode'] = 3:
            { 
                for index, window in mEvent['windowMatch']
                    if !WinExist('ahk_id ' window['id'])
                        mEvent['windowMatch'].RemoveAt(index), aWindowTerminated.Push(window)

                if mEvent['status'] == 'created' && !mEvent['windowMatch'].Length && aWindowTerminated.Length
                    aEventTerminated.Push(aWindowTerminated[1]), mEvent['status'] := 'terminated'
            }

            case mEvent['mode'] = 4:
            { 
                if aWindowCreated.Length
                    aEventCreated := aWindowCreated

                if aWindowTerminated.Length
                    aEventTerminated := aWindowTerminated
            }

            case mEvent['mode'] = 6:
            { 
                if aWindowCreated.Length
                    aEventCreated := aWindowCreated, mEvent['status'] := 'created'

                if mEvent['status'] == 'created' && !mEvent['windowMatch'].Length && aWindowTerminated.Length
                    aEventTerminated.Push(aWindowTerminated[1]), mEvent['status'] := 'terminated'
            }
        }

        ;==============================================
        for _, status in ['created', 'terminated'] {
            if aEvent%status%.Length {
                for index, mValue in aEvent%status% {
                    for _, prop in ['period', 'mode', 'monitoring', 'function', 'critical', 'log', 'notifInfo', 'notifStatus']
                        mValue[prop] := mEvent[prop]

                    mValue['eventName'] := eventName
                    mValue['eventType'] := this.GetEventType(mEvent)
                    mValue['status'] := status
                    this.gMain_lv_event_ModifyIcon(eventName, mEvent['windowMatch'].Length)  
                    
                    if mEvent['log']
                        this.LogEventToFile(this.DisplayObj(mValue))

                    if mEvent['notifInfo']
                        this.Notify(eventName, this.DisplayObj(mValue), mValue['processPath'],, 'topLeft', '10000')     
                    
                    if mEvent['notifStatus']             
                        this.Notify(this.scriptName, 'Event: ' eventName ' ' StrTitle(status), this.mThemeIcons[this.m['mUser']['themeName']]['notifStatus'])

                    SetTimer(this.Sound.Bind(this, this.mSounds[mEvent['sound' StrTitle(status)]]), -1)

                    if (IsSet(%mEvent['function'] '_' status%)) {
                        if mEvent['critical']
                            %mEvent['function'] '_' status%(mValue)  
                        else                   
                            SetTimer(%mEvent['function'] '_' status%.Bind(mValue), -1)  
                    }                    
                }
            }
        }

        DetectHiddenWindows(dhwPrev)
        this.SetCritical_Off(wasNotCritical)
    }

    ;=============================================================================================
    ; Returns an array containing map objects with all existing windows that match the specified parameters. If there are no matching windows, an empty array is returned.
    static WindowFinder(winTitle:='', winClass:='', processName:='', processPath:='', strWinTitleMatchMode := 2, strDetectHiddenWindows := 0, strWinActive:='', strWinMinMax:='')
    {
        tmmPrev := A_TitleMatchMode
        dhwPrev := A_DetectHiddenWindows
        SetTitleMatchMode(strWinTitleMatchMode)
        DetectHiddenWindows(strDetectHiddenWindows)
        aWindowFinder := Array()

        switch {
            case (winTitle && !winClass): winCriteria := winTitle
            case (winClass && !winTitle): winCriteria := 'ahk_class ' winClass
            case (winTitle && winClass):  winCriteria := winTitle ' ahk_class ' winClass
        }

        for strId in WinGetList(IsSet(winCriteria) ? winCriteria : '') 
        {
            mWindowMatch := Map(), strWinTitle := strWinClass := strProcessName := strPid := strProcessPath := winActiveId := str_WinMinMax := ''

            try strWinTitle := WinGetTitle('ahk_id ' strId)

            try strWinClass := WinGetClass('ahk_id ' strId)
            if (winClass && (winClass !== strWinClass)) || strWinClass == 'tooltips_class32'
                continue

            try strProcessName := WinGetProcessName('ahk_id ' strId)
            if processName && (processName != strProcessName)
                continue

            try strPid := WinGetPID('ahk_id ' strId)

            try strProcessPath := ProcessGetPath(strPid)
            if processPath && (processPath !== strProcessPath)
                continue

            if (strWinActive = 1) {
                try winActiveId := WinExist('A')
                if winActiveId != strId
                    continue
            }

            if (RegExMatch(strWinMinMax , '^(-1|0|1)$')) {
                try str_WinMinMax := WinGetMinMax('ahk_id ' strId)
                if str_WinMinMax != strWinMinMax
                    continue
            }

            strElevated := this.IsProcessElevated(strPid)

            for _, value in ['id', 'pid', 'winTitle', 'winClass', 'processName', 'processPath', 'winTitleMatchMode', 'detectHiddenWindows', 'winActive', 'winMinMax', 'elevated']
                mWindowMatch[value] := %'str' value%

            aWindowFinder.Push(mWindowMatch)
        }

        SetTitleMatchMode(tmmPrev)
        DetectHiddenWindows(dhwPrev)
        return aWindowFinder
    }

    ;=============================================================================================

    static SetMonitoringDeviceChange(state)
    {
        if (state = 1 && this.monitoringDeviceChange) || (state = 0 && !this.monitoringDeviceChange)
            return

        if state = 1
            OnMessage(0x219, this.callbackDeviceChange), this.monitoringDeviceChange := true
        else
            OnMessage(0x219, this.callbackDeviceChange, 0), this.monitoringDeviceChange := false
    }

    ;=============================================================================================

    static WM_DEVICECHANGE(wParam, lParam, msg, hwnd)
    {
        (this.debugMode && IsSet(Debug)) ? Debug(wParam, 'WM_DEVICECHANGE_wParam') : ''

        if (this.isCriticalMethodExecuting) {
            (this.debugMode && IsSet(Debug)) ? Debug('isCriticalMethodExecuting_WM_DEVICECHANGE', 'str') : ''
            SetTimer(this.WM_DEVICECHANGE.Bind(this, wParam, lParam, msg, hwnd), -75)
            return
        }

        for eventName, mEvent in this.m['mEvents']
            if mEvent.Has('device') && mEvent['monitoring'] == 'deviceChange' && mEvent['state'] = 1
                SetTimer(this.DeviceStatusFinderUpdater.Bind(this, mEvent, eventName), -this.m['mUser']['delayDeviceChange'])
    }

    ;============================================================================================

    static DeviceStatusFinderUpdater(mEvent, eventName)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        ; (this.debugMode && IsSet(Debug)) ? Debug(mEvent, 'DeviceStatusFinderUpdater') : ''

        aEventCreated := Array(), aEventTerminated := Array(), aDeviceCreated := Array(), aDeviceTerminated := Array()
        aDeviceFinder := this.DeviceFinder(mEvent['device']['deviceName'], mEvent['device']['deviceId'], 'getDeviceIconLarge')

        for _, device in aDeviceFinder {
            iDfound := false
            for _, dev in mEvent['deviceMatch'] {
                if device['deviceId'] == dev['deviceId'] {
                    iDfound := true
                    break
                }
            }

            if !iDfound {
                switch {
                    case RegExMatch(mEvent['mode'], '^(1|3)$'): mEvent['deviceMatch'].Push(device), aEventCreated.Push(device), mEvent['status'] := 'created'
                    case mEvent['mode'] = 2:                    mEvent['deviceMatch'].Push(device), aDeviceCreated.Push(device)
                }
            }            
        }

        ;==============================================
        switch mEvent['mode'] {
            case 1:
            {
                for index, device in mEvent['deviceMatch']
                    if !this.DeviceFinder(, device['deviceId']).Length
                        mEvent['deviceMatch'].RemoveAt(index), aEventTerminated.Push(device)

                if !mEvent['deviceMatch'].Length
                    mEvent['status'] := 'terminated'
            }            

            case 2:
            {    
                for index, device in mEvent['deviceMatch']
                    if !this.DeviceFinder(, device['deviceId']).Length
                        mEvent['deviceMatch'].RemoveAt(index), aDeviceTerminated.Push(device)

                if mEvent['status'] == 'terminated' && mEvent['deviceMatch'].Length
                    aEventCreated.Push(mEvent['deviceMatch'][1]), mEvent['status'] := 'created'

                if mEvent['status'] == 'created' && !mEvent['deviceMatch'].Length && aDeviceTerminated.Length
                    aEventTerminated.Push(aDeviceTerminated[1]), mEvent['status'] := 'terminated'
            }

            case 3:
            {
                for index, device in mEvent['deviceMatch']
                    if !this.DeviceFinder(, device['deviceId']).Length
                        mEvent['deviceMatch'].RemoveAt(index), aDeviceTerminated.Push(device)

                if mEvent['status'] == 'created' && !mEvent['deviceMatch'].Length && aDeviceTerminated.Length
                    aEventTerminated.Push(aDeviceTerminated[1]), mEvent['status'] := 'terminated'
            }
        }

        ;==============================================
        for _, status in ['created', 'terminated'] {
            if aEvent%status%.Length {
                for index, mValue in aEvent%status% {
                    for _, prop in ['period', 'mode', 'monitoring', 'function', 'critical', 'log', 'notifInfo', 'notifStatus']
                        mValue[prop] := mEvent[prop]

                    mValue['eventName'] := eventName
                    mValue['eventType'] := this.GetEventType(mEvent)
                    mValue['status'] := status
                    this.gMain_lv_event_ModifyIcon(eventName, mEvent['deviceMatch'].Length) 

                    if mEvent['log']
                        this.LogEventToFile(this.DisplayObj(mValue))

                    if mEvent['notifInfo']
                        this.Notify(eventName, this.DisplayObj(mValue), 'HICON:*' mValue['hIconLarge'],, 'topLeft', '10000')                      

                    if mEvent['notifStatus']             
                        this.Notify(this.scriptName, 'Event: ' eventName ' ' StrTitle(status), this.mThemeIcons[this.m['mUser']['themeName']]['notifStatus'])
                    
                    SetTimer(this.Sound.Bind(this, this.mSounds[mEvent['sound' StrTitle(status)]]), -1)

                    if (IsSet(%mEvent['function'] '_' status%)) {
                        if mEvent['critical']
                            %mEvent['function'] '_' status%(mValue)  
                        else                   
                            SetTimer(%mEvent['function'] '_' status%.Bind(mValue), -1)  
                    }                    
                }
            }
        }

        this.SetCritical_Off(wasNotCritical)
    }

    ;=============================================================================================
    /*
    Based on EnumDeviceInfo by teadrinker and JEE_DeviceList by jeeswg.
    https://www.autohotkey.com/boards/viewtopic.php?t=69380  https://www.autohotkey.com/boards/viewtopic.php?f=82&t=120154  https://www.autohotkey.com/boards/viewtopic.php?t=121125&p=537515
    Returns an array containing map objects with the device(s) matching the specified parameters. If there are no matching device, an empty array is returned.
    */
    static DeviceFinder(deviceName:='', deviceId:='', param:='')
    {
        static flags := (DIGCF_PRESENT := 0x2) | (DIGCF_ALLCLASSES := 0x4)
                , PKEY_Device_DeviceDesc   := '{A45C254E-DF1C-4EFD-8020-67D146A850E0} 2'
                , PKEY_Device_FriendlyName := '{A45C254E-DF1C-4EFD-8020-67D146A850E0} 14'
                , PKEY_Device_InstanceId   := '{78C34FC8-104A-4ACA-9EA4-524D52996E57} 256'
                , cxSmallIcon := SysGet(SM_CXSMICON := 49), cySmallIcon := SysGet(SM_CYSMICON := 50)
                , cxLargeIcon := SysGet(SM_CXICON := 11),  cyLargeIcon := SysGet(SM_CYICON := 12)

        aDeviceFinder := Array(), friendlyName := deviceDesc := instanceId := ''        
    
        if !(hModule := DllCall('kernel32\LoadLibrary', 'Str','setupapi.dll', 'Ptr'))
        || !(hDevInfo := DllCall('SetupAPI\SetupDiGetClassDevs', 'Ptr', 0, 'Ptr', 0, 'Ptr', 0, 'UInt', flags))
            return      
    
        SP_DEVINFO_DATA := Buffer(size := 24 + A_PtrSize, 0)
        NumPut('UInt', size, SP_DEVINFO_DATA)
    
        DllCall('Propsys\PSPropertyKeyFromString', 'Str', PKEY_Device_FriendlyName, 'Ptr', friendlyName := Buffer(20, 0))
        DllCall('Propsys\PSPropertyKeyFromString', 'Str', PKEY_Device_DeviceDesc, 'Ptr', deviceDesc := Buffer(20, 0))
        DllCall('Propsys\PSPropertyKeyFromString', 'Str', PKEY_Device_InstanceId, 'Ptr', instanceId := Buffer(20, 0))
    
        while DllCall('SetupAPI\SetupDiEnumDeviceInfo', 'Ptr', hDevInfo, 'UInt', A_Index - 1, 'Ptr', SP_DEVINFO_DATA) 
        {
            mDevice := Map()
            getDevicePropFn := GetDeviceProp.Bind(hDevInfo, SP_DEVINFO_DATA)
            devId := getDevicePropFn(instanceId)
     
            if deviceId && (deviceId !== devId)
                continue 
    
            devName := getDevicePropFn(friendlyName) || getDevicePropFn(deviceDesc)
    
            if deviceName && (deviceName !== devName)
                continue      
            
            if param == 'getDeviceIconSmall'
                mDevice['hIconSmall'] := GetDeviceIcon(hDevInfo, SP_DEVINFO_DATA, cxSmallIcon, cySmallIcon)
    
            if param == 'getDeviceIconLarge'
                mDevice['hIconLarge'] := GetDeviceIcon(hDevInfo, SP_DEVINFO_DATA, cxLargeIcon, cyLargeIcon)
    
            mDevice['deviceId'] := devId
            mDevice['deviceName'] := devName
            aDeviceFinder.Push(mDevice)
        }
    
        DllCall('SetupAPI\SetupDiDestroyDeviceInfoList', 'Ptr', hDevInfo)
        DllCall('kernel32\FreeLibrary', 'Ptr', hModule)
        return aDeviceFinder
        
        ;==============================================
        GetDeviceProp(hDevInfo, SP_DEVINFO_DATA, PROPERTYKEY) {
            DllCall('SetupAPI\SetupDiGetDeviceProperty', 'Ptr', hDevInfo, 'Ptr', SP_DEVINFO_DATA, 'Ptr', PROPERTYKEY,
                                                            'UIntP', &propType := 0, 'Ptr', 0, 'UInt', 0, 'UIntP', &reqSize := 0, 'UInt', 0)
            VarSetStrCapacity(&propStr, reqSize)
            DllCall('SetupAPI\SetupDiGetDeviceProperty', 'Ptr', hDevInfo, 'Ptr', SP_DEVINFO_DATA, 'Ptr', PROPERTYKEY,
                                                            'UIntP', &propType, 'Str', propStr, 'UInt', reqSize, 'Ptr', 0, 'UInt', 0)
            return propStr
        }
    
        GetDeviceIcon(hDevInfo, SP_DEVINFO_DATA, cxIcon, cyIcon) {
            DllCall('SetupAPI\SetupDiLoadDeviceIcon', 'Ptr', hDevInfo, 'Ptr', SP_DEVINFO_DATA,
                                                        'UInt', cxIcon, 'UInt', cyIcon, 'UInt', 0, 'PtrP', &hIcon := 0)
            return hIcon
        }
    }

    ;============================================================================================

    static SetPeriodWMI(period)
    {
        if !(period := this.FindClosestNumberInArray(period, this.aPeriodWMIprocess)) || period = this.m['mUser']['periodWMIprocess']
            return               
        
        m := Map('mUser', Map())
        m['mEvents'] := this.Copy_Add_mEvents(this.m['mEvents'])
        m['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(this.m['mUser']['mWinEventsStates'])
        m['mUser']['periodWMIprocess'] := period
        this.SetMonitoringValues(m)
    }

    ;============================================================================================

    static SetProfile(profileName) => this.ProfileSetValues(profileName)
    static TrayMenuProfile_ItemClick(profileName, *) => this.ProfileSetValues(profileName)

    ;=============================================================================================

    static ProfileSetValues(profileName)
    {
        if profileName == this.profileMatch || (!this.mProfiles.Has(profileName) && !RegExMatch(profileName, '^All Events (Enable|Disable)$'))
            return

        m := Map('mUser', Map())

        if (RegExMatch(profileName, '^All Events (Enable|Disable)$')) {
            m['mEvents'] := this.Copy_Add_mEvents(this.m['mEvents'])
            m['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(this.m['mUser']['mWinEventsStates'])
            m['mUser']['periodWMIprocess'] := this.m['mUser']['periodWMIprocess']
        
            for eventName, mEvent in this.m['mEvents']
                m['mEvents'][eventName]['state'] := (profileName == 'All Events Enable') ? 1 : 0
        } 
        else {
            if (this.m['mUser']['profileLoadEvent']) {
                m['mEvents'] := this.Copy_Add_mEvents(this.mProfiles[profileName]['mEvents'])
                m['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(this.mProfiles[profileName]['mUser']['mWinEventsStates'])
                m['mUser']['periodWMIprocess'] := this.mProfiles[profileName]['mUser']['periodWMIprocess']
            }
            else {
                m['mEvents'] := this.Copy_Add_mEvents(this.m['mEvents'])
                m['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(this.m['mUser']['mWinEventsStates'])
                m['mUser']['periodWMIprocess'] := this.m['mUser']['periodWMIprocess']                
                
                for eventName, mEvent in this.m['mEvents']
                    if !this.mProfiles[profileName]['mEvents'].Has(eventName)
                        m['mEvents'] := this.Copy_Add_mEvents(Map(eventName, mEvent), m['mEvents'])
            }

            for _, mEvent in m['mEvents']
                mEvent['state'] := 0

            for eventName, mEvent in this.mProfiles[profileName]['mEvents']
                if mEvent['state'] = 1 && m['mEvents'].Has(eventName)
                    m['mEvents'][eventName]['state'] := 1                      
        }

        this.SetMonitoringValues(m)
    }

    ;=============================================================================================

    static SetEvent(state, eventName)
    {
        this.ValidateParams(
            {num: 1, name: 'state',     value: state,     expect: 'Integer'},
            {num: 2, name: 'eventName', value: eventName, expect: 'String|Integer'}
        )

        if !RegExMatch(state, '^(1|0)$') || (this.m['mEvents'].Has(eventName) && this.m['mEvents'][eventName]['state'] = state)
            return

        this.EventSetValues(eventName)
    }

    ;=============================================================================================

    static TrayMenuEvent_ItemClick(eventName, *) => this.EventSetValues(eventName)

    ;=============================================================================================

    static EventSetValues(eventName)
    {
        if !this.m['mEvents'].Has(eventName)
            return

        m := Map('mUser', Map())
        m['mEvents'] := this.Copy_Add_mEvents(this.m['mEvents'])
        m['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(this.m['mUser']['mWinEventsStates'])
        state := this.m['mEvents'][eventName]['state']
        m['mEvents'][eventName]['state'] := state ? 0 : 1
        m['mUser']['periodWMIprocess'] := this.m['mUser']['periodWMIprocess']
        this.SetMonitoringValues(m)
    }

    ;============================================================================================

    static SetMonitoringValues(m, param:='')
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        ; (this.debugMode && IsSet(Debug)) ? Debug(m, 'SetMonitoringValues_m') : ''
        
        for _, value in ['gMain', 'gSettings']
            try this.%value%.Opt('+Disabled')

        this.GUI_IfExistReturn_Destroy(['gEvent', 'gSettings', 'gAbout', 'gProfile', 'gDelProfile', 'gDelConfirm'], 'destroy')
        this.TrayMenu_Enable_Disable('disable')

        for _, value in ['wmi', 'winEvent', 'deviceChange', 'trayClick']
            this.SetMonitoring%value%(0)

        ;==============================================
        mWinEventValues := this.WinEventCalculation(m['mUser']['mWinEventsStates'])

        for _, value in ['winEventMinHex', 'winEventMaxHex', 'aWinEventExclude']
            this.%value% := mWinEventValues[value]
            
        for _, mEvent in this.m['mEvents']
            if mEvent['monitoring'] == 'timer'
                SetTimer(mEvent['boundFuncTimer'], 0)

        for _, mEvent in m['mEvents']
            if mEvent.Has('process') && mEvent['monitoring'] == 'wmi'
                mEvent['period'] := m['mUser']['periodWMIprocess']

        ;==============================================
        this.m['mEvents'] := this.Copy_Add_mEvents(m['mEvents'])
        this.m['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(m['mUser']['mWinEventsStates'])
        this.m['mUser']['periodWMIprocess'] := m['mUser']['periodWMIprocess']
               
        this.GetEventsFunctionLineNumber(this.m['mEvents'])

        ;==============================================
        for _, mEvent in this.m['mEvents'] {
            mEvent[this.GetEventType(mEvent) 'Match'] := Array()
            mEvent['status'] := 'terminated'

            switch {
                case mEvent.Has('window'):
                {
                    aWindowFinder := this.WindowFinder(
                        mEvent['window']['winTitle'], 
                        mEvent['window']['winClass'], 
                        mEvent['window']['processName'], 
                        mEvent['window']['processPath'], 
                        mEvent['window']['winTitleMatchMode'], 
                        mEvent['window']['detectHiddenWindows'], 
                        mEvent['window']['winActive'], 
                        mEvent['window']['winMinMax']
                    )
                    
                    if aWindowFinder.Length
                        mEvent['windowMatch'] := aWindowFinder, mEvent['status'] := 'created'
                }

                case mEvent.Has('process'):
                {
                    aProcessFinder := this.ProcessFinder(mEvent['process']['processName'], mEvent['process']['processPath'])

                    if aProcessFinder.Length
                        mEvent['processMatch'] := aProcessFinder, mEvent['status'] := 'created'
                }

                case mEvent.Has('device'):
                {
                    aDeviceFinder := this.DeviceFinder(mEvent['device']['deviceName'], mEvent['device']['deviceId'], 'getDeviceIconLarge')

                    if aDeviceFinder.Length
                        mEvent['deviceMatch'] := aDeviceFinder, mEvent['status'] := 'created'
                } 
            }
        }

        ;==============================================
        for eventName, mEvent in this.m['mEvents']
            mEvent['boundFuncTimer'] := this.%this.GetEventType(mEvent)%StatusFinderUpdater.Bind(this, mEvent, eventName)

        for _, mEvent in this.m['mEvents']
            if mEvent['monitoring'] == 'timer' && mEvent['state'] = 1
                SetTimer(mEvent['boundFuncTimer'], mEvent['period'])

        for _, value in ['window', 'process', 'device'] {
            this.monitoringTimer%value% := false
            for _, mEvent in this.m['mEvents'] {
                if (mEvent.Has(value) && mEvent['monitoring'] == 'timer' && mEvent['state'] = 1) {
                    this.monitoringTimer%value% := true
                    break
                }
            }
        }        

        ;==============================================
        mMonitoring := Map()
    
        for _, mEvent in this.m['mEvents'] {
            if (mEvent['state'] = 1) {
                switch {
                    case mEvent.Has('window'):  (mEvent['monitoring'] == 'winEvent') ? mMonitoring['winEvent'] := true :''
                    case mEvent.Has('process'): (mEvent['monitoring'] == 'wmi') ? mMonitoring['wmi'] := true :''
                    case mEvent.Has('device'):  (mEvent['monitoring'] == 'deviceChange') ? mMonitoring['deviceChange'] := true :''
                }
            }
        }

        for _, value in ['wmi', 'winEvent', 'deviceChange']
            this.SetMonitoring%value%(mMonitoring.Has(value))

        this.SetMonitoringTrayClick((this.m['mUser']['trayIconLeftClick'] !== '- None -' || this.m['mUser']['trayIconLeftDClick'] !== '- None -'))

        ;==============================================
        ; Warning: If two events have the same values.
        aDuplicateEvent := Array()
        
        for eventName, mEvent in this.m['mEvents'] {
            for key, value in this.m['mEvents'] {
                if eventName !== key && this.GetEventType(mEvent) == this.GetEventType(value) {
                    switch {
                        case mEvent.Has('window'):
                        {
                            if mEvent['window']['winTitle']    == value['window']['winTitle'] 
                            && mEvent['window']['winClass']    == value['window']['winClass']
                            && mEvent['window']['processName'] == value['window']['processName'] 
                            && mEvent['window']['processPath'] == value['window']['processPath']
                            && mEvent['window']['winActive']   =  value['window']['winActive'] 
                            && mEvent['window']['winMinMax']   =  value['window']['winMinMax']
                            && !this.HasVal(eventName, aDuplicateEvent) {
                                this.MsgBox('Events: "' eventName '" and "' key '" contains the same "WinTitle", "WinClass", "Process Name", "Process Path", "WinActive" and "WinMinMax". It`'s recommended that you delete one of them.', 
                                'Duplicate Event - ' this.scriptName, this.mIcons['g!'], '+AlwaysOnTop -MinimizeBox -MaximizeBox -SysMenu',,,,,
                                [{name:'*OK', callback:'this.MsgBox_Destroy'}])                                 
                                aDuplicateEvent.push(key)
                                break
                            }
                        }                        
                        case mEvent.Has('process'):
                        {
                            if mEvent['process']['processName'] == value['process']['processName'] 
                            && mEvent['process']['processPath'] == value['process']['processPath']
                            && !this.HasVal(eventName, aDuplicateEvent) {
                                this.MsgBox('Events: "' eventName '" and "' key '" contains the same "Process Name" and "Process Path". It`'s recommended that you delete one of them.', 
                                'Duplicate Event - ' this.scriptName, this.mIcons['g!'], '+AlwaysOnTop -MinimizeBox -MaximizeBox -SysMenu',,,,,
                                [{name:'*OK', callback:'this.MsgBox_Destroy'}])                                
                                aDuplicateEvent.push(key)
                                break
                            }
                        }
                        case mEvent.Has('device'):
                        {
                            if mEvent['device']['deviceName'] == value['device']['deviceName']
                            && mEvent['device']['deviceId']   == value['device']['deviceId']
                            && !this.HasVal(eventName, aDuplicateEvent) {
                                this.MsgBox('Events: "' eventName '" and "' key '" contains the same "Device Name" and "Device ID". It`'s recommended that you delete one of them.', 
                                'Duplicate Event - ' this.scriptName, this.mIcons['g!'], '+AlwaysOnTop -MinimizeBox -MaximizeBox -SysMenu',,,,,
                                [{name:'*OK', callback:'this.MsgBox_Destroy'}])
                                aDuplicateEvent.push(key)
                                break
                            }
                        }
                    }
                }
            }
        }

        ;==============================================
        this.profileMatch := this.IsEventsValuesMatchProfile('m')
        (this.debugMode && IsSet(Debug)) ? Debug(this.profileMatch, 'this.profileMatch') : ''

        if (this.profileMatch) {
            SetTimer( this.Sound.Bind(this, this.mSounds[this.m['mUser']['soundProfile']]) , -1 )

            if param  == '__New' && this.m['mUser']['notifProfileStart']
            || param !== '__New' && this.m['mUser']['notifProfileMatch']            
                this.Notify(this.scriptName, 'Profile: ' this.profileMatch, this.mThemeIcons[this.m['mUser']['themeName']]['notifStatus'])
        }
  
        ;==============================================
        if WinExist(this.gMainTitle ' ahk_class AutoHotkeyGUI') && param !== 'gMain_btn_OK' 
        {
            this.mTmp['mEvents'] := this.Copy_Add_mEvents(m['mEvents'])
            this.mTmp['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(m['mUser']['mWinEventsStates'])
            this.mTmp['mUser']['periodWMIprocess'] := m['mUser']['periodWMIprocess']

            this.gMain.lv_event.Delete()
            this.gMain.lv_event.SetImageList(this.imageList)
            this.gMain_lv_event_AddRow_SetIcon()
            this.gMain_lv_event_ModifyCol()
            this.gMain.lv_winEvent.Modify(0, '-Check')

            Loop this.gMain.lv_winEvent.GetCount()
                if this.m['mUser']['mWinEventsStates'][this.gMain.lv_winEvent.GetText(A_Index , 1)] = 1
                    this.gMain.lv_winEvent.Modify(A_Index, 'Check')

            this.DDLchoose(this.m['mUser']['periodWMIprocess'], this.aPeriodWMIprocess, this.gMain.ddl_periodWMIprocess)
            this.gMain_pic_SetColor()
            this.gMain_txt_SetTextBottom()
            this.Create_gMainMenuLoadProfile('SetMonitoringValues', this.profileMatch)
            this.Create_gMainMenuWinEventsPreset()
            ControlSetEnabled(0, this.gMain.btn_apply.hwnd)
        }
        
        this.Create_TrayMenuEvents()
        this.Create_TrayMenuSelectProfile()
        this.TrayMenu_Enable_Disable('enable')
        this.SetIconTip()

        for _, value in ['gMain', 'gSettings']
            try this.%value%.Opt('-Disabled')

        this.SetCritical_Off(wasNotCritical)
    }

    ;=============================================================================================

    static IsEventsValuesMatchProfile(m:='')
    {
        aEventsState1 := Array()
        
        for eventName, mEvent in this.%m%['mEvents']
            if mEvent['state'] = 1
                aEventsState1.Push(eventName)

        for profileName, mProfile in this.mProfiles 
        {           
            allEventMatch := allWinEventMatch := periodWMIprocessMatch := false
            cntEventMatch := cntWinEventMatch := 0

            if (this.m['mUser']['profileLoadEvent']) {
                for eventName, mEvent in mProfile['mEvents'] {
                    if this.%m%['mEvents'].Has(eventName) 
                    && this.GetEventType(mEvent) == this.GetEventType(this.%m%['mEvents'][eventName])
                    && mEvent['state']           =  this.%m%['mEvents'][eventName]['state'] 
                    && mEvent['mode']            =  this.%m%['mEvents'][eventName]['mode']
                    && mEvent['period']          =  this.%m%['mEvents'][eventName]['period']
                    && mEvent['monitoring']      == this.%m%['mEvents'][eventName]['monitoring']
                    && mEvent['log']             =  this.%m%['mEvents'][eventName]['log'] 
                    && mEvent['critical']        =  this.%m%['mEvents'][eventName]['critical']
                    && mEvent['notifInfo']       =  this.%m%['mEvents'][eventName]['notifInfo']
                    && mEvent['notifStatus']     =  this.%m%['mEvents'][eventName]['notifStatus']
                    && mEvent['soundCreated']    == this.%m%['mEvents'][eventName]['soundCreated']
                    && mEvent['soundTerminated'] == this.%m%['mEvents'][eventName]['soundTerminated'] {

                        switch {
                            case mEvent.Has('window'):
                            {  
                                if mEvent['window']['winTitle']    == this.%m%['mEvents'][eventName]['window']['winTitle']
                                && mEvent['window']['winClass']    == this.%m%['mEvents'][eventName]['window']['winClass']
                                && mEvent['window']['processName'] == this.%m%['mEvents'][eventName]['window']['processName']
                                && mEvent['window']['processPath'] == this.%m%['mEvents'][eventName]['window']['processPath']
                                && mEvent['window']['winActive']   =  this.%m%['mEvents'][eventName]['window']['winActive']
                                && mEvent['window']['winMinMax']   =  this.%m%['mEvents'][eventName]['window']['winMinMax']
                                    cntEventMatch++
                            }

                            case mEvent.Has('process'):
                            { 
                                if mEvent['process']['processName'] = this.%m%['mEvents'][eventName]['process']['processName'] 
                                && mEvent['process']['processPath'] == this.%m%['mEvents'][eventName]['process']['processPath']
                                    cntEventMatch++
                            }
                            
                            case mEvent.Has('device'): 
                            { 
                                if mEvent['device']['deviceName'] == this.%m%['mEvents'][eventName]['device']['deviceName'] 
                                && mEvent['device']['deviceId']   == this.%m%['mEvents'][eventName]['device']['deviceId']
                                    cntEventMatch++
                            }
                        }
                    }
                }

                if this.%m%['mEvents'].Count = mProfile['mEvents'].Count && this.%m%['mEvents'].Count = cntEventMatch
                    allEventMatch := true

                for winEventName, value in mProfile['mUser']['mWinEventsStates']
                    if this.%m%['mUser']['mWinEventsStates'][winEventName] = value
                        cntWinEventMatch++
    
                if cntWinEventMatch = this.mWinEvents.Count
                    allWinEventMatch := true
    
                if mProfile['mUser']['periodWMIprocess'] = this.%m%['mUser']['periodWMIprocess']
                    periodWMIprocessMatch := true
            }
            else {
                aEventsPState1 := Array()
                
                for eventName, mEvent in mProfile['mEvents']
                    if mEvent['state'] = 1
                        aEventsPState1.Push(eventName)

                if this.IsArrayContainsSameValues(aEventsState1, aEventsPState1)
                    allEventMatch := true          
            }

            if !this.m['mUser']['profileLoadEvent']
                allWinEventMatch := periodWMIprocessMatch := true

            if allEventMatch && allWinEventMatch && periodWMIprocessMatch && this.%m%['mEvents'].Count
                return profileName
        }

        if this.%m%['mEvents'].Count && this.%m%['mEvents'].Count = aEventsState1.Length
            return 'All Events Enable'

        if this.%m%['mEvents'].Count && !aEventsState1.Length
            return 'All Events Disable'

        return 0
    }

    ;============================================================================================

    static Create_gMainMenuLoadProfile(fromMethod:='', profileMatch:='')
    {
        this.gMainMenuLoadProfile := Menu()

        for _, value in ['All Events Enable', 'All Events Disable'] {
            this.gMainMenuLoadProfile.Add(value, this.gMain_MenuProfile_ItemClick.Bind(this))
            this.gMainMenuLoadProfile.SetIcon(value, this.mIcons['tMenuTrans'])
        }
        
        for profileName, mProfile in this.mProfiles {
            this.gMainMenuLoadProfile.Add(profileName, this.gMain_MenuProfile_ItemClick.Bind(this))
            this.gMainMenuLoadProfile.SetIcon(profileName, this.mIcons['tMenuTrans'])
        }

        if fromMethod !== 'SetMonitoringValues'
            profileMatch := this.IsEventsValuesMatchProfile('mTmp')

        if profileMatch
            this.gMainMenuLoadProfile.SetIcon(profileMatch, this.mThemeIcons[this.m['mUser']['themeName']]['select'])
    }

    ;============================================================================================

    static Create_gMainMenuWinEventsPreset()
    {        
        this.gMainMenuWinEventsPreset := Menu()

        Loop this.mWinEventPresets.Count
            this.gMainMenuWinEventsPreset.Add('Preset ' A_Index (this.mDefault['winEventPreset'] = A_Index ? ' (Default)' : ''), this.gMain_MenuWinEventPreset_ItemClick.Bind(this))

        for index, aWinEventName in this.mWinEventPresets {
            if index = this.mDefault['winEventPreset']
                index := index ' (Default)'

            this.gMainMenuWinEventsPreset.SetIcon('Preset ' index, this.mIcons['tMenuTrans'])
        }

        if index := this.IsWinEventValuesMatchPreset()
            this.gMainMenuWinEventsPreset.SetIcon('Preset ' index, this.mThemeIcons[this.m['mUser']['themeName']]['select'])
    }

    ;============================================================================================

    static IsWinEventValuesMatchPreset()
    {
        aWinEventsState1 := Array()
        
        for winEventName, state in this.mTmp['mUser']['mWinEventsStates']
            if state = 1
                aWinEventsState1.Push(winEventName)

        for index, aWinEventName in this.mWinEventPresets {
            if this.IsArrayContainsSameValues(aWinEventsState1, aWinEventName) {
                if index = this.mDefault['winEventPreset']
                    index := index ' (Default)'
            
                return Index
            }
        }
    }

    ;============================================================================================

    static WinEventCalculation(mWinEventsStates)
    {
        ; if all WinEvents are set to 0. Set to Preset 1.
        cntState0 := 0
        for winEventName, state in mWinEventsStates
            state = 0 ? cntState0++ : ''

        if cntState0 = this.mWinEvents.Count
            for _, winEventName in this.mWinEventPresets[1]
                mWinEventsStates[winEventName] := 1

        ;==============================================
        for _, value in this.aWinEventOrder {
            for winEventName, state in mWinEventsStates {
                if (value == winEventName && state = 1) {
                    winEventMinName := winEventName
                    winEventMinHex := this.mWinEvents[winEventName]['hex']
                    break 2
                }
            }
        }

        for _, value in this.aWinEventOrder {
            for winEventName, state in mWinEventsStates {
                if (value == winEventName && state = 1) {
                    winEventMaxName := winEventName
                    winEventMaxHex := this.mWinEvents[winEventMaxName]['hex']
                }
            }
        }

        for index, winEventName in this.aWinEventOrder {
            if winEventName == winEventMinName
                eventMinIndex := index

            if winEventName == winEventMaxName
                eventMaxIndex := index
        }

        aWinEventExclude := Array()
        
        for index, winEventName in this.aWinEventOrder
            if index > eventMinIndex && index < eventMaxIndex
                for k, state in mWinEventsStates
                    if k == winEventName && state = 0
                        aWinEventExclude.Push(this.mWinEvents[k]['event'])

        return mWinEventValues := Map('winEventMinHex', winEventMinHex, 'winEventMaxHex', winEventMaxHex, 'mWinEventsStates', mWinEventsStates, 'aWinEventExclude', aWinEventExclude)
    }      

    ;=============================================================================================

    static gMain_Show(*)
    {
        if this.GUI_IfExistReturn_Destroy(['gMain', 'gEvent', 'gSettings', 'gAbout'], 'ifExistReturn')
            return  

        wasNotCritical := this.SetCritical_On(A_IsCritical)
        
        this.gMain := Gui('+Resize +MinSize', this.gMainTitle)
        this.gMain.OnEvent('Size', this.gMain_Size.Bind(this))
        this.gMain.OnEvent('Close', this.gMain_Close.Bind(this))
        this.gMain.SetFont('s10')
        
        this.gMain.Tips := GuiCtrlTips(this.gMain)
        this.gMain.Tips.SetDelayTime('AUTOPOP', 30000)

        this.gMain.btnTopSize := 35
        this.gMain.btnTopWidth := 120
        this.gMain.btnTopHeight := 35
        this.gMain.lv_eventWidth := 625
        this.gMain.lv_winEventWidth := 300
        this.gMain.lv_height := 445
        this.btnMarginX := this.gMain.MarginX/2
     
        this.gMain.btn_addEvent := this.gMain.Add('Button', 'w' this.gMain.btnTopWidth ' h' this.gMain.btnTopHeight, 'Add Event')
        GuiButtonIcon(this.gMain.btn_addEvent, this.mIcons['gMain+'], 1, 's20 a0 l10')
        this.gMain.btn_addEvent.OnEvent('Click', this.gEvent_show.Bind(this, this.gMain_CreateDefaultEventMap(), '', 'add'))
        this.gMain.btn_editEvent := this.gMain.Add('Button', 'x+' this.btnMarginX ' w' this.gMain.btnTopSize ' h' this.gMain.btnTopSize)
        this.gMain.btn_editEvent.OnEvent('Click', this.gMain_btn_editEvent.Bind(this))
        GuiButtonIcon(this.gMain.btn_editEvent, this.mThemeIcons[this.m['mUser']['themeName']]['gMainEdit'], 1, 's20')
        this.gMain.Tips.SetTip(this.gMain.btn_editEvent, 'Edit selected event')                
        this.gMain.btn_deleteEvent := this.gMain.Add('Button', 'x+' this.btnMarginX ' w' this.gMain.btnTopSize ' h' this.gMain.btnTopSize)
        this.gMain.btn_deleteEvent.OnEvent('Click', this.gMain_lv_event_DeleteEvent.Bind(this))
        GuiButtonIcon(this.gMain.btn_deleteEvent, this.mIcons['gMainX'], 1, 's20')
        this.gMain.Tips.SetTip(this.gMain.btn_deleteEvent, 'Delete selected event(s)')
        this.gMain.btn_selectAll := this.gMain.Add('Button', 'x+' this.btnMarginX ' w' this.gMain.btnTopSize ' h' this.gMain.btnTopSize)
        this.gMain.btn_selectAll.OnEvent('Click', this.gMain_lv_event_SelectAll.Bind(this))
        GuiButtonIcon(this.gMain.btn_selectAll, this.mIcons['selectAll'], 1, 's20')
        this.gMain.Tips.SetTip(this.gMain.btn_selectAll, 'Select All')
        this.gMain.btn_checkAll := this.gMain.Add('Button', 'x+' this.btnMarginX ' w' this.gMain.btnTopSize ' h' this.gMain.btnTopSize)
        this.gMain.btn_checkAll.OnEvent('Click', this.gMain_lv_event_CheckAll_UncheckAll.Bind(this, 'checkAll'))
        GuiButtonIcon(this.gMain.btn_checkAll, this.mIcons['checkAll'], 1, 's20')
        this.gMain.Tips.SetTip(this.gMain.btn_checkAll, 'Check all events')
        this.gMain.btn_unCheckAll := this.gMain.Add('Button', 'x+' this.btnMarginX ' w' this.gMain.btnTopSize ' h' this.gMain.btnTopSize)
        this.gMain.btn_unCheckAll.OnEvent('Click', this.gMain_lv_event_CheckAll_UncheckAll.Bind(this, 'unCheckAll'))
        GuiButtonIcon(this.gMain.btn_unCheckAll, this.mIcons['unCheckAll'], 1, 's20')
        this.gMain.Tips.SetTip(this.gMain.btn_unCheckAll, 'Uncheck all events') 
        this.gMain.btn_manageProfile := this.gMain.Add('Button', 'x+' this.btnMarginX ' w' this.gMain.btnTopSize ' h' this.gMain.btnTopSize)
        this.gMain.btn_manageProfile.OnEvent('Click', this.gProfile_Show.Bind(this))
        GuiButtonIcon(this.gMain.btn_manageProfile, this.mThemeIcons[this.m['mUser']['themeName']]['gMainProfile'], 1, 's20')
        this.gMain.Tips.SetTip(this.gMain.btn_manageProfile, 'Profiles Manager')                      
        this.gMain.btn_loadProfile := this.gMain.Add('Button', 'x+' this.btnMarginX ' w' this.gMain.btnTopWidth ' h' this.gMain.btnTopHeight, 'Load Profile')
        this.gMain.btn_loadProfile.OnEvent('Click', this.gMain_btn_DropDownMenu_Click.Bind(this, 'gMainMenuLoadProfile'))
        GuiButtonIcon(this.gMain.btn_loadProfile, this.mIcons['downArrow'], 1, 's14 a1 r10')                    
        this.gMain.btn_settings := this.gMain.Add('Button', 'x+' this.btnMarginX ' w' this.gMain.btnTopSize ' h' this.gMain.btnTopSize)
        this.gMain.btn_settings.OnEvent('Click', this.gSettings_Show.Bind(this))
        GuiButtonIcon(this.gMain.btn_settings, this.mThemeIcons[this.m['mUser']['themeName']]['gMainSettings'], 1, 's20')
        this.gMain.Tips.SetTip(this.gMain.btn_settings, 'Settings')                
        this.gMain.btn_about := this.gMain.Add('Button', 'x+' this.btnMarginX ' w' this.gMain.btnTopSize ' h' this.gMain.btnTopSize)
        this.gMain.btn_about.OnEvent('Click', this.gAbout_Show.Bind(this))
        GuiButtonIcon(this.gMain.btn_about, this.mThemeIcons[this.m['mUser']['themeName']]['gMainAbout'], 1, 's20') 
        this.gMain.Tips.SetTip(this.gMain.btn_about, 'About')                    
        this.gMain.btn_alwaysOnTop := this.gMain.Add('Picture', 'x+' this.btnMarginX ' ym+3 w28 h28')
        this.gMain.btn_alwaysOnTop.OnEvent('Click', this.gMain_cb_alwaysOnTop_Click.Bind(this))
        this.gMain.Tips.SetTip(this.gMain.btn_alwaysOnTop, 'Cycle Always on Top') 
        this.gMain.btn_presets := this.gMain.Add('Button', 'xm+' this.gMain.lv_eventWidth + this.gMain.MarginX ' ym w' this.gMain.btnTopWidth ' h' this.gMain.btnTopHeight, 'Load Preset')
        this.gMain.btn_presets.OnEvent('Click', this.gMain_btn_DropDownMenu_Click.Bind(this, 'gMainMenuWinEventsPreset'))
        GuiButtonIcon(this.gMain.btn_presets, this.mIcons['downArrow'], 1, 's14 a1 r10')
        this.gMain.Tips.SetTip(this.gMain.btn_presets, 'Load WinEvents Preset')        
        this.gMain.btn_donate := this.gMain.Add('Button', 'x' this.gMain.lv_eventWidth + this.gMain.lv_winEventWidth + this.gMain.MarginX*2 - this.btnMarginX - 24 - 90 ' ym w90 h35', "Donate")
        this.gMain.btn_donate.OnEvent('Click', (*) => Run(this.linkPayPal))
        GuiButtonIcon(this.gMain.btn_donate, this.mIcons['gMainPaypal'], 1, 's20 a0 l10')
        this.gMain.Tips.SetTip(this.gMain.btn_donate, 'Donate')              
        this.gMain.btn_help := this.gMain.Add('Button', 'x' this.gMain.lv_eventWidth + this.gMain.lv_winEventWidth + this.gMain.MarginX*2 - 24 ' ym h24 w24')
        this.gMain.btn_help.OnEvent('Click', (*) => Run(This.linkGitHub))
        GuiButtonIcon(this.gMain.btn_help, this.mIcons['?Small'], 1, 's18')
        this.gMain.Tips.SetTip(this.gMain.btn_help, 'View the documentation on GitHub.')
        ;==============================================        
        this.gMain.lv_event := this.gMain.Add('ListView', 'xm w' this.gMain.lv_eventWidth ' h' this.gMain.lv_height ' Checked Grid Sort Section +BackgroundEAF4FB', ['Event Name', 'Type', 'Function', 'Line', 'Monitoring', 'Period/Delay', 'Mode', 'Options'])
        this.gMain.lv_event.OnEvent('ContextMenu', this.gMain_lv_event_ContextMenu.Bind(this))
        this.gMain.lv_event.OnEvent('DoubleClick', this.gMain_lv_event_DoubleClick.Bind(this))
        this.gMain.lv_event.OnEvent('ItemCheck', this.gMain_lv_event_ItemCheck.Bind(this))
        this.gMain.lv_event.OnEvent('ItemSelect', this.gMain_lv_event_ItemSelect.Bind(this))
        this.gMain.lv_winEvent := this.gMain.Add('ListView', 'x+' this.gMain.MarginX ' w' this.gMain.lv_winEventWidth ' h' this.gMain.lv_height ' Checked Grid NoSortHdr +BackgroundEAF4FB', ['WinEvent'])
        this.gMain.lv_winEvent.OnEvent('ItemCheck', this.gMain_lv_winEvent_ItemCheck.Bind(this))
        ;==============================================
        this.gMain.gbHeight := 85
        this.gMain.gbWidth := 170                      
        this.gMain.gb_process := this.gMain.Add('GroupBox', 'xs ys+' this.gMain.lv_height + this.gMain.MarginY ' w' this.gMain.gbWidth ' h' this.gMain.gbHeight ' Section cBlack', 'Process')
        this.gMain.pic_wmiProcess := this.gMain.Add('Picture', 'xs+10 ys+25 w20 h20')
        this.gMain.txt_wmi := this.gMain.Add('Text', 'x+5 ys+28', 'WMI')
        this.gMain.pic_timerprocess := this.gMain.Add('Picture', 'xs+10 ys+55 w20 h20')
        this.gMain.txt_timerprocess := this.gMain.Add('Text', 'x+5 ys+58', 'Timer')                        
        this.gMain.ddl_periodWMIprocess := this.gMain.Add('DropDownList', 'w55 xs+80 ys+25', this.aPeriodWMIprocess)
        this.gMain.ddl_periodWMIprocess.OnEvent('Change', this.gMain_ddl_periodWMIprocess_CtrlChange.Bind(this))
        this.gMain.Tips.SetTip(this.gMain.ddl_periodWMIprocess, 'WMI period interval')        
        this.gMain.txt_msWMIProcess := this.gMain.Add('Text', 'xs+140 ys+30', 'ms')
        ;==============================================               
        this.gMain.gbWidth2 := 135
        this.gMain.gb_window := this.gMain.Add('GroupBox', 'xs+' this.gMain.gbWidth + this.gMain.MarginX ' ys w' this.gMain.gbWidth2 ' h' this.gMain.gbHeight ' Section cBlack', 'Window')
        this.gMain.pic_winEvent := this.gMain.Add('Picture', 'xs+10 ys+25 w20 h20')
        this.gMain.txt_winEvent := this.gMain.Add('Text', 'x+5 ys+28', 'WinEvent')
        this.gMain.pic_timerWindow := this.gMain.Add('Picture', 'xs+10 ys+55 w20 h20')
        this.gMain.txt_timerWindow := this.gMain.Add('Text', 'x+5 ys+58', 'Timer')                
        this.gMain.gb_device := this.gMain.Add('GroupBox', 'xs+' this.gMain.gbWidth2 + this.gMain.MarginX ' ys w' this.gMain.gbWidth2 ' h' this.gMain.gbHeight ' Section cBlack', 'Device')
        this.gMain.pic_deviceChange := this.gMain.Add('Picture', 'xs+10 ys+25 w20 h20')
        this.gMain.txt_deviceChange := this.gMain.Add('Text', 'x+5 ys+28', 'DeviceChange')
        this.gMain.pic_timerDevice := this.gMain.Add('Picture', 'xs+10 ys+55 w20 h20')
        this.gMain.txt_timerDevice := this.gMain.Add('Text', 'x+5 ys+58', 'Timer')
        ;==============================================
        this.gMain.btnOCAwidth := 140
        this.gMain.btnOCAheight := 40
        this.gMain.btnsOCAwidth := this.gMain.btnOCAwidth*3 + this.gMain.MarginX*2
        btnOCAPos := this.gMain.lv_eventWidth + this.gMain.lv_winEventWidth - this.gMain.btnsOCAwidth + this.gMain.MarginX*2

        this.gMain.btn_ok := this.gMain.Add('Button', 'x' btnOCAPos ' y' this.gMain.btnTopHeight + this.gMain.lv_height + this.gMain.gbHeight - this.gMain.btnOCAheight + this.gMain.MarginY*2 ' w' this.gMain.btnOCAwidth ' h' this.gMain.btnOCAheight ' Default', 'OK')
        this.gMain.btn_ok.OnEvent('Click', this.gMain_btnOK_btnApply_Click.Bind(this))
        this.gMain.btn_cancel := this.gMain.Add('Button', 'x+m w' this.gMain.btnOCAwidth ' h' this.gMain.btnOCAheight, 'Cancel')
        this.gMain.btn_cancel.OnEvent('Click', this.gMain_Close.Bind(this))
        this.gMain.btn_apply := this.gMain.Add('Button', 'x+m w' this.gMain.btnOCAwidth ' h' this.gMain.btnOCAheight, 'Apply')
        this.gMain.btn_apply.OnEvent('Click', this.gMain_btnOK_btnApply_Click.Bind(this))
        this.gMain.txtBottomHeight := 18
        this.gMain.txt_bottom := this.gMain.Add('Text', 'xm h' this.gMain.txtBottomHeight ' w900')
        this.gMain.txt_bottom.SetFont('bold')

        ; MenuBar ===============================
        this.gMain.mb := MenuBar()
        this.gMain.mb.Add('File', this.gMain.mb_file := Menu())
        this.gMain.mb_file.Add('Open Script Folder', (*) => Run(A_WorkingDir))
        this.gMain.mb_file.SetIcon('Open Script Folder', this.mThemeIcons[this.m['mUser']['themeName']]['gMenuFolder'])
        this.gMain.mb_file.Add('Open Startup Folder', (*) => Run(A_Startup))
        this.gMain.mb_file.SetIcon('Open Startup Folder', this.mThemeIcons[this.m['mUser']['themeName']]['gMenuFolder'])
        this.gMain.mb_file.Add('Open Events Log', this.RunTool.Bind(this, 'EventsLog.txt'))
        this.gMain.mb_file.Add()
        this.gMain.mb_file.Add('Close', this.gMain_Close.Bind(this))
        this.gMain.mb.Add('Edit', this.gMain.mb_edit := Menu())
        this.gMain.mb_edit.Add('Select All`tCtrl+A', this.gMain_lv_event_SelectAll.Bind(this))
        this.gMain.mb_edit.SetIcon('Select All`tCtrl+A', this.mIcons['selectAll'])
        this.gMain.mb.Add('Options', this.gMain.mb_options := Menu())
        this.gMain.mb_options.Add('Settings...', this.gSettings_Show.Bind(this))
        this.gMain.mb_options.SetIcon('Settings...', this.mThemeIcons[this.m['mUser']['themeName']]['gMenuSettings'])

        ;==============================================
        this.gMain.mb.Add('Tools', this.gMain.mb_tools := Menu())

        for key, mTool in this.mTools {
            this.gMain.mb_tools.Add(mTool['name'], this.RunTool.Bind(this, mTool['path']))
            
            if mTool.Has('iconPath') && FileExist(mTool['iconPath'])
                this.gMain.mb_tools.SetIcon(mTool['name'], mTool['iconPath'])
        }

        ;==============================================
        this.gMain.mb.Add('Help', this.gMain.mb_help := Menu())
        this.gMain.mb_help.Add('GitHub Repository', (*) => Run(this.linkGitHub))
        this.gMain.mb_help.SetIcon('GitHub Repository', this.mIcons['gMenuGitHub'])
        this.gMain.mb_help.Add('Donate', (*) => Run(this.linkPayPal))
        this.gMain.mb_help.SetIcon('Donate', this.mIcons['gMenuPayPal'])
        this.gMain.mb_help.Add('About', this.gAbout_Show.Bind(this))
        this.gMain.mb_help.SetIcon('About', this.mThemeIcons[this.m['mUser']['themeName']]['gMenuAbout'])
        this.gMain.MenuBar := this.gMain.mb

        this.imageList := IL_Create()
        IL_Add(this.imageList, this.mIcons['btnGreen'])
        IL_Add(this.imageList, this.mIcons['btnGrey'])
        this.gMain.lv_event.SetImageList(this.imageList)                                
        this.gMain_lv_event_AddRow_SetIcon()
        this.gMain_lv_event_ModifyCol()
            
        ;==============================================
        for index, value in this.aWinEventOrder
            this.gMain.lv_winEvent.Add('Checked', value)

        Loop this.gMain.lv_winEvent.GetCount()
            if this.m['mUser']['mWinEventsStates'][this.gMain.lv_winEvent.GetText(A_Index , 1)] = 1
                this.gMain.lv_winEvent.Modify(A_Index, 'Check')

        ;==============================================
        for _, value in ['gb_window', 'gb_process', 'gb_device']
            this.gMain.%value%.SetFont('bold')

        this.mTmp['mEvents'] := this.Copy_Add_mEvents(this.m['mEvents'])
        this.mTmp['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(this.m['mUser']['mWinEventsStates'])
        this.mTmp['mUser']['periodWMIprocess'] := this.m['mUser']['periodWMIprocess']

        this.DDLchoose(this.m['mUser']['periodWMIprocess'], this.aPeriodWMIProcess, this.gMain.ddl_periodWMIprocess)
        this.Create_gMainMenuLoadProfile()
        this.Create_gMainMenuWinEventsPreset()
        this.gMain_pic_SetColor()
        this.gMain_txt_SetTextBottom()
        ControlSetEnabled(0, this.gMain.btn_apply.hwnd)
        LV_GridColor(this.gMain.lv_event, '0xcbcbcb')
        LV_GridColor(this.gMain.lv_winEvent, '0xcbcbcb')
        this.gMain.Show()
        this.gMain_btn_alwaysOnTop_SetValue()
        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gMain_Close(*) => this.gMain.Destroy()

    ;============================================================================================

    static gMain_Size(guiObj, minMax, width, height)
    {
        if minMax = -1
            return

        posYctrl := height - this.gMain.gbHeight - this.gMain.MarginY*4

        this.MoveControls(                   
                                  
            {Control:this.gMain.lv_event, w:width - this.gMain.lv_winEventWidth - this.gMain.MarginX*3, h:height - this.gMain.gbHeight - this.gMain.btnTopHeight - this.gMain.MarginY*7},
            {Control:this.gMain.lv_winEvent, x:width - this.gMain.lv_winEventWidth - this.gMain.MarginX, h:height - this.gMain.gbHeight - this.gMain.btnTopHeight - this.gMain.MarginY*7},
            {Control:this.gMain.gb_process, y:posYctrl},
            {Control:this.gMain.gb_window, y:posYctrl},
            {Control:this.gMain.gb_device, y:posYctrl},
            {Control:this.gMain.pic_wmiProcess, y:posYctrl + 25},
            {Control:this.gMain.txt_wmi, y:posYctrl + 28},
            {Control:this.gMain.pic_timerprocess, y:posYctrl + 55},
            {Control:this.gMain.txt_timerprocess, y:posYctrl + 58},                        
            {Control:this.gMain.ddl_periodWMIprocess, y:posYctrl + 25},            
            {Control:this.gMain.txt_msWMIProcess, y:posYctrl + 30},            
            {Control:this.gMain.pic_winEvent, y:posYctrl + 25},
            {Control:this.gMain.txt_winEvent, y:posYctrl + 28},
            {Control:this.gMain.pic_timerWindow, y:posYctrl + 55},
            {Control:this.gMain.txt_timerWindow, y:posYctrl + 58},
            {Control:this.gMain.pic_deviceChange, y:posYctrl + 25},
            {Control:this.gMain.txt_deviceChange, y:posYctrl + 28},
            {Control:this.gMain.pic_timerDevice, y:posYctrl + 55},
            {Control:this.gMain.txt_timerDevice, y:posYctrl + 58},      
            {Control:this.gMain.btn_presets, x:width - this.gMain.lv_winEventWidth - this.gMain.MarginX},
            {Control:this.gMain.btn_donate, x:width - this.gMain.MarginX - this.btnMarginX - 24 - 90},       
            {Control:this.gMain.btn_help, x:width - this.gMain.MarginX - 24},
            {Control:this.gMain.btn_ok, x:width - this.gMain.btnsOCAwidth - this.gMain.MarginX, y:height - this.gMain.btnOCAheight - this.gMain.txtBottomHeight - this.gMain.MarginY*3},
            {Control:this.gMain.btn_cancel, x:width - this.gMain.btnOCAwidth*2 - this.gMain.MarginX*2, y:height - this.gMain.btnOCAheight - this.gMain.txtBottomHeight - this.gMain.MarginY*3},            
            {Control:this.gMain.btn_apply, x:width - this.gMain.btnOCAwidth - this.gMain.MarginX, y:height - this.gMain.btnOCAheight - this.gMain.txtBottomHeight - this.gMain.MarginY*3},            
            {Control:this.gMain.txt_bottom , y:height - this.gMain.txtBottomHeight - this.gMain.MarginY}
        )

        DllCall('RedrawWindow', 'ptr', this.gMain.hwnd, 'ptr', 0, 'ptr', 0, 'uint', 0x0081)
    }

    ;============================================================================================

    static gMain_btnOK_btnApply_Click(ctrlObj, *)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        
        this.gMain.Opt('+Disabled')

        for _, value in ['Apply', 'OK']
            if ctrlObj.Text == value
                ControlSetEnabled(0, this.gMain.%'btn_' value%.hwnd)

        if (ctrlObj.Text == 'OK' && !ControlGetEnabled(this.gMain.btn_apply.hwnd)) {
            this.gMain_Close()
            this.SetCritical_Off(wasNotCritical)
            return
        }

        m := Map('mUser', Map())
        m['mEvents'] := this.Copy_Add_mEvents(this.mTmp['mEvents'])
        m['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(this.mTmp['mUser']['mWinEventsStates'])
        m['mUser']['periodWMIprocess'] := this.mTmp['mUser']['periodWMIprocess']

        this.SetMonitoringValues(m, 'gMain_btn_' ctrlObj.Text)        
        this.WriteValuesToJSON()
        
        this.gMain.Opt('-Disabled')

        if ctrlObj.Text == 'OK'
            this.gMain_Close()

        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gMain_lv_event_ContextMenu(ctrlObj, item, isRightClick, x, y)
    {
        MouseGetPos(,,, &mouseOverClassNN)
        if item = 0 || InStr(mouseOverClassNN, 'SysHeader')
            return

        eventName := ctrlObj.GetText(item, 1)
        lineNumber := ctrlObj.GetText(item, 4)
        
        ctxMenu := Menu()
        ctxMenu.Add('Edit Event', this.gEvent_show.Bind(this, this.mTmp['mEvents'][eventName], eventName, 'edit'))
        ctxMenu.SetIcon('Edit Event', this.mThemeIcons[this.m['mUser']['themeName']]['gMenuEdit'])

        if IsNumber(lineNumber)
            itemNameEdit := 'Edit Function'
        else
            itemNameEdit := 'Edit Script' 
        
        ctxMenu.Add(itemNameEdit, EditScript.Bind(lineNumber))
        ctxMenu.SetIcon(itemNameEdit, this.mThemeIcons[this.m['mUser']['themeName']]['gMenuEdit'])
        ctxMenu.Add('Delete', this.gMain_lv_event_DeleteEvent.Bind(this))
        ctxMenu.SetIcon('Delete', this.mIcons['gMenuX'])
        ctxMenu.Show(x, y)

        EditScript(lineNumber, *)
        {
            Edit
            if (IsNumber(lineNumber)) 
            {
                SetTimer(() => GoToLine(), -1750)
                
                coordmmPrev := A_CoordModeMouse
                CoordMode('Mouse', 'Screen'), MouseGetPos(&mX, &mY)
                ToolTip('Waiting for the editor to open...', mX, mY)
                CoordMode('Mouse', coordmmPrev)
            }

            GoToLine() {
                activeWinTitle := ''
                try activeWinTitle := WinGetTitle('A')
                
                if (!InStr(activeWinTitle, A_ScriptName)) {
                    ToolTip
                    this.MsgBox('The script file failed to open.', this.gErrorTitle, this.mIcons['gX'], '+AlwaysOnTop -MinimizeBox -MaximizeBox -SysMenu',,,,,
                    [{name:'*OK', callback:'this.MsgBox_Destroy'}])
                    return
                }
                ToolTip

                if (WinActive('ahk_exe SciTE.exe') || WinActive('ahk_exe notepad++.exe') || WinActive('ahk_exe Code.exe')) && InStr(activeWinTitle, A_ScriptName)
                    SendInput('^g' lineNumber '{enter}')
            }
        }
    }

    ;============================================================================================

    static gMain_lv_event_ItemSelect(ctrlObj, item, selected, *) => this.Create_gMainMenuLoadProfile()

    ;============================================================================================

    static gMain_lv_event_SelectAll(*) => this.gMain.lv_event.Modify(0, 'Select')

    ;============================================================================================

    static gMain_CreateDefaultEventMap(*)
    {
        return Map( 
        'period',          this.m['mUser']['delayWinEvent'], 
        'mode',            this.m['mUser']['modeWindow'], 
        'monitoring',      this.m['mUser']['monitoringWindow'], 
        'log',             this.m['mUser']['log'],
        'critical',        this.m['mUser']['critical'],
        'notifInfo',       this.m['mUser']['notifInfo'],
        'notifStatus',     this.m['mUser']['notifStatus'],
        'soundCreated',    this.m['mUser']['soundCreated'],
        'soundTerminated', this.m['mUser']['soundTerminated'],
        'window', Map(
            'winTitle', '', 'winClass', '', 'processName', '', 'processPath', '', 
            'winTitleMatchMode',   this.m['mUser']['winTitleMatchMode'], 
            'detectHiddenWindows', this.m['mUser']['detectHiddenWindows'], 
            'winActive', this.m['mUser']['winActive'], 
            'winMinMax', this.m['mUser']['winMinMax'])
        )
    }

    ;============================================================================================

    static gMain_lv_event_DeleteEvent(*)
    {
        if !selectedContent := ListViewGetContent('Col1 Selected', this.gMain.lv_event.hwnd, this.gMainTitle ' ahk_class AutoHotkeyGUI')
            return

        if (StrSplit(selectedContent, '`n').Length > 1) {
            this.MsgBox('Are you sure you want to delete all selected events?', this.gDelConfirmTitle, this.mIcons['g?'],, this.gMain, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
            [{name: '*Yes', callback: 'this.gMain_lv_event_DeleteEvent_ClickYes'}, 
             {name: 'Cancel', callback: 'this.MsgBox_Destroy'}])
        } 
        else
            this.gMain_lv_event_DeleteEvent_ClickYes()
    }

    ;============================================================================================

    static gMain_lv_event_DeleteEvent_ClickYes(gIndex:='', owner:='', *)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)

        Loop Parse, ListViewGetContent('Col1 Selected', this.gMain.lv_event.hwnd, this.gMainTitle ' ahk_class AutoHotkeyGUI'), '`n'
            this.mTmp['mEvents'].Delete(A_LoopField)

        loop this.gMain.lv_event.GetCount('Selected')
            this.gMain.lv_event.Delete(this.gMain.lv_event.GetNext(0))
   
        this.gMain_lv_event_ModifyCol()
        this.Create_gMainMenuLoadProfile()
        this.gMain_btn_apply_EnableDisable()
        this.gMain_txt_SetTextBottom()
        
        if gIndex 
            this.MsgBox_Destroy(gIndex, owner)
        
        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gMain_btn_editEvent(*)
    {
        Loop Parse, ListViewGetContent('Col1 Selected', this.gMain.lv_event.hwnd, this.gMainTitle ' ahk_class AutoHotkeyGUI'), '`n' {
            this.gEvent_show(this.mTmp['mEvents'][A_LoopField], A_LoopField, 'edit')
            break
        }
    }

    ;============================================================================================

    static gMain_lv_event_DoubleClick(ctrlObj, item, *)
    {
        if item = 0
            return

        eventName := ctrlObj.GetText(item, 1)
        this.gEvent_show(this.mTmp['mEvents'][eventName], eventName, 'edit')
    }

    ;============================================================================================

    static gMain_btn_DropDownMenu_Click(menuName, ctrlObj, *)
    {
        ControlGetPos(&x, &y, &w, &h, ctrlObj.hwnd)
        this.%menuName%.Show(x, y+h)
    }

    ;============================================================================================

    static gMain_MenuProfile_ItemClick(profileName, *)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)

        this.gMain.Opt('+Disabled')
        this.gMain.lv_event.Modify(0, '-Check')

        if RegExMatch(profileName, '^All Events (Enable|Disable)$') {
            for eventName, mEvent in this.mTmp['mEvents']
                mEvent['state'] := (profileName == 'All Events Enable' ? 1 : 0)

            if profileName == 'All Events Enable'
                this.gMain.lv_event.Modify(0, 'Check')
        } 
        else {            
            this.gMain.lv_event.Delete()

            if (this.m['mUser']['profileLoadEvent']) 
            {
                this.mTmp['mEvents'] := this.Copy_Add_mEvents(this.mProfiles[profileName]['mEvents'])
                this.mTmp['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(this.mProfiles[profileName]['mUser']['mWinEventsStates'])
                this.mTmp['mUser']['periodWMIprocess'] := this.mProfiles[profileName]['mUser']['periodWMIprocess']

                for _, mEvent in this.mTmp['mEvents']
                    if (mEvent.Has('process') && mEvent['monitoring'] == 'wmi')
                        mEvent['period'] := this.mProfiles[profileName]['mUser']['periodWMIprocess']

                this.DDLchoose(this.mProfiles[profileName]['mUser']['periodWMIprocess'], this.aPeriodWMIprocess, this.gMain.ddl_periodWMIprocess)

                this.gMain.lv_winEvent.Modify(0, '-Check')
            
                Loop this.gMain.lv_winEvent.GetCount()
                    if this.mProfiles[profileName]['mUser']['mWinEventsStates'][this.gMain.lv_winEvent.GetText(A_Index , 1)] = 1
                        this.gMain.lv_winEvent.Modify(A_Index, 'Check')                
            }

            for _, mEvent in this.mTmp['mEvents']
                mEvent['state'] := 0

            for eventName, mEvent in this.mProfiles[profileName]['mEvents']
                if mEvent['state'] = 1 && this.mTmp['mEvents'].Has(eventName)
                    this.mTmp['mEvents'][eventName]['state'] := 1
                    
            this.GetEventsFunctionLineNumber(this.mTmp['mEvents'])

            for eventName, mEvent in this.mTmp['mEvents'] 
                this.gMain.lv_event.Add('Icon2 Checked', eventName, this.FormatTXT(this.GetEventType(mEvent)), mEvent['function'], mEvent['lineNumber'], this.FormatTxt(mEvent['monitoring']), mEvent['period'], this.FormatTxtMode(mEvent), this.FormatTxtOptions(mEvent))

            this.gMain_lv_event_ModifyCol()

            Loop this.gMain.lv_event.GetCount()
                if this.mTmp['mEvents'][this.gMain.lv_event.GetText(A_Index , 1)]['state'] = 1
                    this.gMain.lv_event.Modify(A_Index, 'Check')                      
        }

        this.Create_gMainMenuLoadProfile()
        this.Create_gMainMenuWinEventsPreset()
        this.gMain_btn_apply_EnableDisable()
        this.gMain_txt_SetTextBottom()

        this.gMain.Opt('-Disabled')
        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gMain_MenuWinEventPreset_ItemClick(itemName, *)
    {
        RegExMatch(itemName, '\d+', &match)

        Loop this.gMain.lv_winEvent.GetCount() {
            winEventName := this.gMain.lv_winEvent.GetText(A_Index, 1)

            if this.HasVal(winEventName, this.mWinEventPresets[Number(match[0])])
                this.gMain.lv_winEvent.Modify(A_Index, 'Check'), this.mTmp['mUser']['mWinEventsStates'][winEventName] := 1
            else
                this.gMain.lv_winEvent.Modify(A_Index, '-Check'), this.mTmp['mUser']['mWinEventsStates'][winEventName] := 0
        }

        this.Create_gMainMenuLoadProfile()
        this.Create_gMainMenuWinEventsPreset()
        this.gMain_btn_apply_EnableDisable()
    }

    ;============================================================================================

    static gMain_lv_event_CheckAll_UncheckAll(check, *)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        this.gMain.Opt('+Disabled')
        this.gMain.lv_event.Modify(0, '-Check')

        for eventName, mEvent in this.mTmp['mEvents']
            mEvent['state'] := (check == 'checkAll') ? 1 : 0

        if check == 'checkAll'
            this.gMain.lv_event.Modify(0, 'Check')
        
        this.Create_gMainMenuLoadProfile()
        this.gMain_btn_apply_EnableDisable()
        this.gMain.Opt('-Disabled')
        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gMain_ddl_periodWMIprocess_CtrlChange(*)
    {
        this.mTmp['mUser']['periodWMIprocess'] := this.gMain.ddl_periodWMIprocess.Text
        this.Create_gMainMenuLoadProfile()
        this.gMain_btn_apply_EnableDisable()
    }

    ;============================================================================================

    static gMain_lv_event_ItemCheck(ctrlObj, item, checked, *)
    {
        this.mTmp['mEvents'][this.gMain.lv_event.GetText(item, 1)]['state'] := checked
        this.Create_gMainMenuLoadProfile()
        this.gMain_btn_apply_EnableDisable()
    }

    ;============================================================================================
    
    static gMain_lv_winEvent_ItemCheck(ctrlObj, item, checked, *)
    {
        this.mTmp['mUser']['mWinEventsStates'][this.gMain.lv_winEvent.GetText(item, 1)] := checked
        this.Create_gMainMenuLoadProfile()
        this.Create_gMainMenuWinEventsPreset()
        this.gMain_btn_apply_EnableDisable()
    }

    ;============================================================================================

    static gMain_btn_apply_EnableDisable(*)
    {
        if this.m['mEvents'].Count != this.mTmp['mEvents'].Count
            eventCountChanged := true

        if this.gMain.ddl_periodWMIprocess.Text != this.m['mUser']['periodWMIprocess']
            periodWMIprocessChanged := true

        for winEventName, state in this.m['mUser']['mWinEventsStates'] {
            if this.mTmp['mUser']['mWinEventsStates'][winEventName] != state {
                winEventStateChanged := true
                break
            } 
        } 

        for eventName, mEvent in this.mTmp['mEvents'] 
        {
            if !this.m['mEvents'].Has(eventName) {
                eventValueChanged := true
                break
            }

            if mEvent['state']           !=  this.m['mEvents'][eventName]['state']
            || this.GetEventType(mEvent) !== this.GetEventType(this.m['mEvents'][eventName])
            || mEvent['function']        !== this.m['mEvents'][eventName]['function'] 
            || mEvent['mode']            !=  this.m['mEvents'][eventName]['mode']
            || mEvent['period']          !=  this.m['mEvents'][eventName]['period']
            || mEvent['monitoring']      !== this.m['mEvents'][eventName]['monitoring']
            || mEvent['log']             !=  this.m['mEvents'][eventName]['log'] 
            || mEvent['critical']        !=  this.m['mEvents'][eventName]['critical']
            || mEvent['notifInfo']       !=  this.m['mEvents'][eventName]['notifInfo']
            || mEvent['notifStatus']     !=  this.m['mEvents'][eventName]['notifStatus']
            || mEvent['soundCreated']    !== this.m['mEvents'][eventName]['soundCreated']
            || mEvent['soundTerminated'] !== this.m['mEvents'][eventName]['soundTerminated'] {
                eventValueChanged := true
                break
            }

            switch {
                case mEvent.Has('window'): 
                {
                    if mEvent['window']['winTitle']           !== this.m['mEvents'][eventName]['window']['winTitle'] 
                    || mEvent['window']['winClass']           !== this.m['mEvents'][eventName]['window']['winClass']
                    || mEvent['window']['processName']        !== this.m['mEvents'][eventName]['window']['processName']
                    || mEvent['window']['processPath']        !== this.m['mEvents'][eventName]['window']['processPath'] 
                    || mEvent['window']['winActive']           != this.m['mEvents'][eventName]['window']['winActive']
                    || mEvent['window']['winMinMax']           != this.m['mEvents'][eventName]['window']['winMinMax']
                    || mEvent['window']['detectHiddenWindows'] != this.m['mEvents'][eventName]['window']['detectHiddenWindows']
                    || mEvent['window']['winTitleMatchMode']   != this.m['mEvents'][eventName]['window']['winTitleMatchMode'] {
                        eventValueChanged := true
                        break
                    }
                }                
                case mEvent.Has('process'): 
                {
                    if mEvent['process']['processName'] !== this.m['mEvents'][eventName]['process']['processName'] 
                    || mEvent['process']['processPath'] !== this.m['mEvents'][eventName]['process']['processPath'] {
                        eventValueChanged := true
                        break
                    }
                }
                case mEvent.Has('device'): 
                {
                    if mEvent['device']['deviceName'] !== this.m['mEvents'][eventName]['device']['deviceName']
                    || mEvent['device']['deviceId']   !== this.m['mEvents'][eventName]['device']['deviceId'] {
                        eventValueChanged := true
                        break
                    }
                }
            }
        }

        if IsSet(eventCountChanged) || IsSet(winEventStateChanged) || IsSet(periodWMIprocessChanged) || IsSet(eventValueChanged)
            ControlSetEnabled(1, this.gMain.btn_apply.hwnd)
        else
            ControlSetEnabled(0, this.gMain.btn_apply.hwnd)
    }

    ;============================================================================================
    
    static gMain_lv_event_AddRow_SetIcon()
    {
        for eventName, mEvent in this.m['mEvents'] {
            if mEvent['state'] = 1 && mEvent['status'] == 'created'
                indexIcon := 1
            else
                indexIcon := 2

            this.gMain.lv_event.Add(
                'Icon' indexIcon ' Checked', 
                eventName, 
                this.FormatTxt(this.GetEventType(mEvent)), 
                mEvent['function'], 
                mEvent['lineNumber'], 
                this.FormatTxt(mEvent['monitoring']), 
                mEvent['period'],
                this.FormatTxtMode(mEvent),
                this.FormatTxtOptions(mEvent)
            )
        }
        
        Loop this.gMain.lv_event.GetCount()
            if this.m['mEvents'][this.gMain.lv_event.GetText(A_Index , 1)]['state'] = 1
                this.gMain.lv_event.Modify(A_Index, 'Check')  
    }

    ;============================================================================================

    static gMain_lv_event_ModifyCol()
    {
        for _, col in [1, 2, 3, 5, 7, 8]
            this.gMain.lv_event.ModifyCol(col, 'AutoHdr')
        
        this.gMain.lv_event.ModifyCol(4, '50 Integer')
        this.gMain.lv_event.ModifyCol(6, '100 Integer')
    }

    ;============================================================================================

    static gMain_lv_event_ModifyIcon(eventName, matchLength)
    {
        if WinExist(this.gMainTitle ' ahk_class AutoHotkeyGUI') {
            Loop this.gMain.lv_event.GetCount() {
                if this.gMain.lv_event.GetText(A_Index , 1) == eventName {
                    if matchLength
                        this.gMain.lv_event.Modify(A_Index, 'Icon1')
                    else
                        this.gMain.lv_event.Modify(A_Index, 'Icon2')
                    break
                }
            }
        }
    }   

    ;============================================================================================

    static gMain_cb_alwaysOnTop_Click(*)
    {
        state := this.m['mUser']['alwaysOnTop']
        this.m['mUser']['alwaysOnTop'] := state ? 0 : 1
        this.gMain_btn_alwaysOnTop_SetValue()
    }

    ;============================================================================================

    static gMain_btn_alwaysOnTop_SetValue(*)
    {    
        if (this.m['mUser']['alwaysOnTop']) {
            WinSetAlwaysOnTop(1, this.gMainTitle ' ahk_class AutoHotkeyGUI')
            this.gMain.btn_alwaysOnTop.Value := this.mIcons['gMainAOT1']
        } 
        else {
            WinSetAlwaysOnTop(0, this.gMainTitle ' ahk_class AutoHotkeyGUI')
            this.gMain.btn_alwaysOnTop.Value := this.mIcons['gMainAOT0']
        }
    }

    ;=============================================================================================

    static gMain_pic_SetColor()
    {
        for _, value in ['TimerProcess', 'TimerWindow', 'TimerDevice', 'WMIProcess', 'winEvent', 'deviceChange']
            this.gMain.pic_%value%.Value := this.monitoring%value% ? this.mIcons['btnGreen'] : this.mIcons['btnGrey']
    }

    ;============================================================================================

    static gMain_txt_SetTextBottom()
    {
        cntWindow := cntProcess := cntDevice := 0

        for eventName, mEvent in this.mTmp['mEvents'] {
            switch {
                case mEvent.Has('window') : cntWindow++
                case mEvent.Has('process'): cntProcess++
                case mEvent.Has('device') : cntDevice++
            }
        }

        cntEvent := this.mTmp['mEvents'].Count

        this.gMain.txt_bottom.Text := '  ' cntEvent (cntEvent > 1 ? ' Events' : ' Event') '   | ' 
        . cntWindow  (cntWindow > 1 ? ' Windows'   : ' Window')  ' | '  
        . cntProcess (cntWindow > 1 ? ' Processes' : ' Process') ' | ' 
        . cntDevice  (cntDevice > 1 ? ' Devices'   : ' Device')  ' | '
        . '   Active Profile:  ' (this.profileMatch ? this.profileMatch : '- None -')
    }    

    ;============================================================================================

    static gProfile_Show(*)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)

        this.gProfile := Gui('-MinimizeBox', this.gProfileTitle)
        this.gProfile.SetFont('s10')
        this.gProfile.Tips := GuiCtrlTips(this.gProfile)

        try this.gMain.Opt('+Disabled')
        try this.gProfile.Opt('+Owner' this.gMain.hwnd)
        
        btnWidth := 140
        btnHeight := 40
        btnMarginX := this.gProfile.MarginX/2
        gbSaveHeight := 65
        gbLoadHeight := 100
        gbNotifyHeight := 135
        ctrlWidth := 210      
        btnSize := 35
        gbDelPfileHeight := gbSaveHeight + gbLoadHeight + gbNotifyHeight + this.gProfile.MarginY*2
        lvHeight := gbDelPfileHeight - this.gProfile.MarginY*5
        gbWidth := this.gProfile.MarginX + ctrlWidth + btnSize + btnMarginX*3
        
        this.gProfile.gb_name := this.gProfile.Add('GroupBox', 'w' gbWidth ' h' gbSaveHeight ' cBlack', 'Save')
        this.gProfile.gb_name.SetFont('bold')
        this.gProfile.combox_profile := this.gProfile.Add('ComboBox', 'xp+' this.gProfile.MarginX ' yp+25 w' ctrlWidth, this.Create_Array_Profiles())                                                       
        this.gProfile.btn_save := this.gProfile.Add('Button', 'x+' btnMarginX ' yp-6 w' btnSize ' h' btnSize ' Default') 
        this.gProfile.btn_save.OnEvent('Click', this.gProfile_btn_save_Click.Bind(this))
        GuiButtonIcon(this.gProfile.btn_save, this.mIcons['floppy'], 1, 's20')
        this.gProfile.Tips.SetTip(this.gProfile.btn_save, 'Save Profile')
        this.gProfile.gb_delete := this.gProfile.Add('GroupBox', 'xm+' gbWidth + this.gProfile.MarginX ' ym w' gbWidth ' h' gbDelPfileHeight ' cBlack', 'Delete')
        this.gProfile.lv := this.gProfile.Add('ListView', 'xp+' this.gProfile.MarginX ' yp+25 w' ctrlWidth ' h' lvHeight ' Grid Sort Section +BackgroundEAF4FB', ['Profile'])
        this.gProfile.lv.OnEvent('ContextMenu', this.gDelProfile_ContextMenu.Bind(this))
        this.gProfile.btn_delete := this.gProfile.Add('Button',  'xs+' ctrlWidth + btnMarginX ' yp-6 w' btnSize ' h' btnSize)
        this.gProfile.btn_delete.OnEvent('Click', this.gDelProfile_lv_Delete.Bind(this))
        GuiButtonIcon(this.gProfile.btn_delete, this.mIcons['gMainX'], 1, 's20')
        this.gProfile.Tips.SetTip(this.gProfile.btn_delete, 'Delete selected Profile(s)') 
        this.gProfile.gb_name := this.gProfile.Add('GroupBox', 'xm ym+' gbSaveHeight + this.gProfile.MarginY ' w' gbWidth ' h' gbLoadHeight ' cBlack', 'When loading profiles, load:')       
        this.gProfile.cb_profileLoadEventStates := this.gProfile.Add('Checkbox', 'xp+' this.gProfile.MarginX ' yp+25 w15 h15 Section')
        this.gProfile.cb_profileLoadEventStates.value := 1
        ControlSetEnabled(0, this.gProfile.cb_profileLoadEventStates.hwnd)
        this.gProfile.Add('Text', 'x+1', ' Events states')          
        this.gProfile.cb_profileLoadEvent := this.gProfile.Add('Checkbox', 'xs', ' All events values')
        this.gProfile.cb_profileLoadEvent.OnEvent('Click', (*) => this.m['mUser']['profileLoadEvent'] := this.gProfile.cb_profileLoadEvent.Value)
        this.gProfile.gb_notify := this.gProfile.Add('GroupBox', 'xm ym+' gbSaveHeight + gbLoadHeight + this.gProfile.MarginY*2 ' w' gbWidth ' h' gbNotifyHeight ' cBlack', 'Notifications')        
        this.gProfile.cb_notifProfileStart := this.gProfile.Add('Checkbox', 'xp+' this.gProfile.MarginX ' yp+25 Section', ' Profile match at startup')
        this.gProfile.cb_notifProfileStart.OnEvent('Click', (*) => this.m['mUser']['notifProfileStart'] := this.gProfile.cb_notifProfileStart.Value)                
        this.gProfile.cb_notifProfileMatch := this.gProfile.Add('Checkbox', 'xs', ' Profile match')
        this.gProfile.cb_notifProfileMatch.OnEvent('Click', (*) => this.m['mUser']['notifProfileMatch'] := this.gProfile.cb_notifProfileMatch.Value)        
        this.gProfile.Add('Text', 'xs', 'Sound when profile match:')
        this.gProfile.ddl_soundProfile := this.gProfile.Add('DropDownList', 'xs w' ctrlWidth, this.aSounds)
        this.gProfile.ddl_soundProfile.OnEvent('Change', this.DDL_soundProfile_CtrlChange.Bind(this, 'gProfile', 'soundProfile'))  
        
        btnPosX := (gbWidth*2 + this.gProfile.MarginX*3)/2 - btnWidth/2       
        
        this.gProfile.btn_close := this.gProfile.Add('Button', 'x' btnPosX ' y' gbDelPfileHeight + this.gProfile.MarginY*3 ' w' btnWidth ' h' btnHeight ' Default', 'Close')        
        this.gProfile.btn_close.OnEvent('Click', this.gProfile_Close.Bind(this))
        this.gProfile.OnEvent('Close', this.gProfile_Close.Bind(this))                   
        this.gProfile.OnEvent('Close', this.gProfile_Close.Bind(this))        
        this.gProfile.MarginY := this.gProfile.MarginY*2
        
        for profileName, mProfile in this.mProfiles
            this.gProfile.lv.Add(, profileName)         

        for _, value in ['profileLoadEvent', 'notifProfileMatch', 'notifProfileStart']
            this.gProfile.cb_%value%.Value := this.m['mUser'][value]       

        for _, value in ['name', 'notify', 'delete']
            this.gProfile.%'gb_' value%.SetFont('bold')
        
        this.DDLchoose(this.m['mUser']['soundProfile'], this.aSounds, this.gProfile.ddl_soundProfile)
        LV_GridColor(this.gProfile.lv, '0xcbcbcb')
        this.GUI_ShowOnSameDisplay(this.gProfile, WinExist(this.gMainTitle ' ahk_class AutoHotkeyGUI'))
       
        if this.m['mUser']['alwaysOnTop']
            WinSetAlwaysOnTop(1, this.gProfileTitle  ' ahk_class AutoHotkeyGUI')

        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gProfile_Close(*) 
    {
        try this.gMain.Opt('-Disabled')
        this.gProfile.Destroy()
    }     

    ;============================================================================================

    static gProfile_btn_save_Click(*)
    {      
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        profileName := Trim(ControlGetText(this.gProfile.combox_profile.hwnd))

        Switch {
            case !profileName: returnFlag := true

            case (profileName && !this.IsValidEventName(profileName)):
            {
                this.MsgBox('Invalid profile name. It cannot contain special characters and must be limited to 50 characters in length.', this.gErrorTitle, this.mIcons['gX'],,
                this.gProfile, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true  
            }
        }
            
        if (IsSet(returnFlag)) {
            this.SetCritical_Off(wasNotCritical)
            return
        }       
        
        DirCreate('Profiles')  
        pathProfile := 'Profiles/' profileName '.json'
        try FileRecycle(pathProfile)    

        ;==============================================
        m := Map('mUser', Map())
        m['mEvents'] := this.Copy_Add_mEvents(this.mTmp['mEvents'])
        m['mUser']['mWinEventsStates'] := this.Copy_mWinEventsStates(this.mTmp['mUser']['mWinEventsStates'])
        m['mUser']['periodWMIprocess'] := this.mTmp['mUser']['periodWMIprocess']
        
        for _, mEvent in m['mEvents'] 
            if (mEvent.Has('process') && mEvent['monitoring'] == 'wmi')
                mEvent['period'] := m['mUser']['periodWMIprocess']

        if this.mProfiles.Has(profileName)
            this.mProfiles.Delete(profileName)

        this.mProfiles[profileName] := m

        ;==============================================
        strJSON := JSON.stringify(m, expandlevel := unset, space := "  ")        
        fileProfileJSON := FileOpen(pathProfile, 'w', 'UTF-8')
        fileProfileJSON.Write(strJSON)
        fileProfileJSON.Close() 
      
        ;==============================================
        this.gProfile_lv_Delete_Update()                 
        this.DDLarrayChange_Choose(profileName, this.Create_Array_Profiles(), this.gProfile.combox_profile)
        this.Create_gMainMenuLoadProfile()
        this.profileMatch := this.IsEventsValuesMatchProfile('m')
        this.Create_TrayMenuSelectProfile()
        this.SetIconTip()
        this.gMain_txt_SetTextBottom()

        this.MsgBox('Successfully saved! The profile includes all events values, WMI period interval and window events (WinEvents) states.', this.gSaveConfirmTitle, this.mIcons['gI'],, 
        this.gProfile, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
        [{name: '*OK', callback: 'this.MsgBox_Destroy'}])

        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gDelProfile_ContextMenu(ctrlObj, item, isRightClick, x, y)
    {
        MouseGetPos(,,, &mouseOverClassNN)
        if item = 0 || InStr(mouseOverClassNN, 'SysHeader')
            return

        ctxMenu := Menu()
        ctxMenu.Add('Delete', this.gDelProfile_lv_Delete.Bind(this))
        ctxMenu.SetIcon('Delete', this.mIcons['gMenuX'])
        ctxMenu.Show(x, y)
    }

    ;============================================================================================

    static gProfile_lv_Delete_Update()
    {
        this.gProfile.lv.Delete()
            
        for profileName, mProfile in this.mProfiles
            this.gProfile.lv.Add('', profileName)
    }    
    
    ;============================================================================================

    static gDelProfile_lv_Delete(*)
    {
        if (StrSplit( ListViewGetContent('Col1 Selected', this.gProfile.lv.hwnd, this.gMainTitle ' ahk_class AutoHotkeyGUI') , '`n').Length > 1) 
        {
            this.MsgBox('Are you sure you want to delete all selected profiles?', this.gDelConfirmTitle, this.mIcons['g?'],, 
            this.gProfile, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
            [{name: '*Yes', callback: 'this.DeleteEntry_ClickYes'}, 
             {name: 'Cancel', callback: 'this.MsgBox_Destroy'}])          

        } else
            this.DeleteEntry_ClickYes()
    }

    ;============================================================================================

    static DeleteEntry_ClickYes(gIndex:='', owner:='', *)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)

        Loop Parse, ListViewGetContent('Col1 Selected', this.gProfile.lv.hwnd, this.gMainTitle ' ahk_class AutoHotkeyGUI'), '`n' {
            this.mProfiles.Delete(A_LoopField)
            Try FileRecycle('Profiles\' A_LoopField '.json')
        }

        profileName := this.gProfile.combox_profile.Text

        this.gProfile_lv_Delete_Update()
        this.DDLarrayChange_Choose(profileName, this.Create_Array_Profiles(), this.gProfile.combox_profile)
        this.Create_gMainMenuLoadProfile()
        this.Create_TrayMenuSelectProfile()
        
        if gIndex 
            this.MsgBox_Destroy(gIndex, owner)
        
        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static DDL_soundProfile_CtrlChange(gName, DDLname, *)
    {
        this.m['mUser'][DDLname] := this.%gName%.ddl_%DDLname%.Text
        SetTimer( this.Sound.Bind(this, this.mSounds[this.%gName%.ddl_%DDLname%.Text]) , -1 )
    }
   
    ;============================================================================================

    static gEvent_Show(mEvent:='', eventName:='', addOrEdit:='', *)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        eventType := this.GetEventType(mEvent)
        
        this.gEvent := Gui('-MinimizeBox', this.gEventTitle)
        this.gEvent.OnEvent('Size', this.gEvent_Size.Bind(this))
        this.gEvent.Opt('+Resize +MinSize810x400')
        this.gEvent.SetFont('s10')
        
        this.gEvent.Tips := GuiCtrlTips(this.gEvent)
        this.gEvent.Tips.SetDelayTime('AUTOPOP', 30000)

        this.gMain.Opt('+Disabled')
        this.gEvent.Opt('+Owner' this.gMain.hwnd)

        this.gEvent.gbWidth := 255
        this.gEvent.gbHeight := 130
        this.gEvent.gbHeightNotif := 110
        this.gEvent.gbWidthMon := 175
        this.gEvent.editWidth := 235
        this.gEvent.lvWidth := 645
        this.gEvent.lvHeight := 540
              
        this.gEvent.radio_window := this.gEvent.Add('Radio','xm+' this.gEvent.gbWidth + this.gEvent.MarginX ' Group -Wrap' (mEvent.Has('window') ? ' Checked' : ''), ' Window')
        this.gEvent.radio_window.OnEvent('Click', this.gEvent_radio_eventType_Click.Bind(this, 'window'))
        this.gEvent.radio_process := this.gEvent.Add('Radio','x+m ym -Wrap' (mEvent.Has('process') ? ' Checked' : ''), ' Process')
        this.gEvent.radio_process.OnEvent('Click', this.gEvent_radio_eventType_Click.Bind(this, 'process'))
        this.gEvent.radio_device := this.gEvent.Add('Radio','x+m ym -Wrap' (mEvent.Has('device') ? ' Checked' : ''), ' Device')
        this.gEvent.radio_device.OnEvent('Click', this.gEvent_radio_eventType_Click.Bind(this, 'device'))
        this.gEvent.btn_help := this.gEvent.Add('Button', 'x' this.gEvent.gbWidth + this.gEvent.lvWidth + this.gEvent.MarginX*3 - 24 ' y5 h24 w24')
        this.gEvent.btn_help.OnEvent('Click', (*) => Run(This.linkGitHub))
        GuiButtonIcon(this.gEvent.btn_help, this.mIcons['?Small'], 1, 's18')
        this.gEvent.Tips.SetTip(this.gEvent.btn_help, 'View the documentation on GitHub.')
        this.gEvent.Add('GroupBox', 'xm y1 w' this.gEvent.gbWidth ' h' this.gEvent.gbHeight ' cBlack')
        this.gEvent.Add('Text', 'xp+10 yp+15 Section', 'Event Name:')
        this.gEvent.edit_eventName := this.gEvent.Add('Edit', 'xs w' this.gEvent.editWidth ' h23 -Wrap', eventName)
        this.gEvent.Add('Text', 'xs', 'Function Name:')
        this.gEvent.cb_critical := this.gEvent.Add('CheckBox', 'x+60', ' Critical')
        this.gEvent.Tips.SetTip(this.gEvent.cb_critical, this.mGuiCtrlTips['gEvent.cb_critical'])
        this.gEvent.cb_critical.Value := mEvent['critical']
        this.gEvent.edit_function := this.gEvent.Add('Edit', 'xs w' this.gEvent.editWidth ' h23 -Wrap', mEvent.Has('function') ? mEvent['function'] : '')
        ;==============================================
        this.gEvent.gb_notif := this.gEvent.Add('GroupBox', 'xm w' this.gEvent.gbWidth ' h' this.gEvent.gbHeightNotif ' cBlack', 'Notification')
        this.gEvent.cb_notifStatus := this.gEvent.Add('CheckBox', 'xp+10 yp+25 Section', ' Status')
        this.gEvent.cb_notifStatus.Value := mEvent['notifStatus']
        this.gEvent.Tips.SetTip(this.gEvent.cb_notifStatus, 'Display GUI with the event status in the bottom right corner.')
        this.gEvent.cb_notifInfo := this.gEvent.Add('CheckBox', 'x+20', ' Info')
        this.gEvent.cb_notifInfo.Value := mEvent['notifInfo']
        this.gEvent.Tips.SetTip(this.gEvent.cb_notifInfo, 'Display GUI with the event information in the top left corner.')
        this.gEvent.cb_log := this.gEvent.Add('CheckBox', 'x+20', ' Log Info')
        this.gEvent.cb_log.Value := mEvent['log']
        this.gEvent.Tips.SetTip(this.gEvent.cb_log, 'Log event information to "EventsLog.txt"')
        this.gEvent.ddl_soundCreated := this.gEvent.Add('DropDownList', 'xs w' this.gEvent.editWidth, this.aSounds)
        this.gEvent.ddl_soundCreated.OnEvent('Change', this.DDL_Sound_CtrlChange.Bind(this, 'gEvent', 'soundCreated'))
        this.DDLchoose(mEvent['soundCreated'], this.aSounds, this.gEvent.ddl_soundCreated)
        this.gEvent.Tips.SetTip(this.gEvent.ddl_soundCreated, 'Sound upon event creation.')
        this.gEvent.ddl_soundTerminated := this.gEvent.Add('DropDownList', 'xs w' this.gEvent.editWidth, this.aSounds)
        this.gEvent.ddl_soundTerminated.OnEvent('Change', this.DDL_Sound_CtrlChange.Bind(this, 'gEvent', 'soundTerminated'))
        this.DDLchoose(mEvent['soundTerminated'], this.aSounds, this.gEvent.ddl_soundTerminated)
        this.gEvent.Tips.SetTip(this.gEvent.ddl_soundTerminated, 'Sound upon event termination.')
        this.gEvent.btn_clear := this.gEvent.Add('Button', 'xs+' this.gEvent.editWidth -20 ' y' this.gEvent.gbHeight + this.gEvent.gbHeightNotif + 20 ' h20 w20')
        this.gEvent.btn_clear.OnEvent('Click', this.gEvent_btn_clear_Click.Bind(this))
        GuiButtonIcon(this.gEvent.btn_clear, this.mIcons['clear'], 1, 's14')
        this.gEvent.Tips.SetTip(this.gEvent.btn_clear, 'Clear all input fields.')
        ;==============================================
        this.gEvent.window := Object()
        this.gEvent.gbHeightWin := 305
        this.gEvent.gbHeightMon := 75
        this.gEvent.window.gb_window := this.gEvent.Add('GroupBox', 'xm ym+' this.gEvent.gbHeight + this.gEvent.gbHeightNotif ' w' this.gEvent.gbWidth ' h' this.gEvent.gbHeightWin ' +Hidden cBlack')
        this.gEvent.window.txt_winTitle := this.gEvent.Add('Text', 'xp+10 yp+15 Section +Hidden', 'WinTitle:')
        this.gEvent.window.edit_winTitle := this.gEvent.Add('Edit', 'xs w' this.gEvent.editWidth ' yp+21 h23 -Wrap +Hidden', mEvent.Has('window') ? mEvent['window']['winTitle'] : '')
        this.gEvent.window.txt_winClass := this.gEvent.Add('Text', 'xs +Hidden', 'WinClass:')
        this.gEvent.window.edit_winClass := this.gEvent.Add('Edit', 'xs w' this.gEvent.editWidth ' yp+21 h23 -Wrap +Hidden', mEvent.Has('window') ? mEvent['window']['winClass'] : '')
        this.gEvent.window.txt_processName := this.gEvent.Add('Text', 'xs +Hidden', 'Process Name:')
        this.gEvent.window.edit_processName := this.gEvent.Add('Edit', 'xs w' this.gEvent.editWidth ' yp+21 h23 -Wrap +Hidden', mEvent.Has('window') ? mEvent['window']['processName'] : '')
        this.gEvent.window.txt_processPath := this.gEvent.Add('Text', 'xs +Hidden', 'Process Path:')
        this.gEvent.window.edit_processPath := this.gEvent.Add('Edit', 'xs w' this.gEvent.editWidth ' yp+21 h23 -Wrap +Hidden', mEvent.Has('window') ? mEvent['window']['processPath'] : '')        
        this.gEvent.window.txt_winTitleMatchMode := this.gEvent.Add('Text', 'xs +Hidden', 'WinTitleMatchMode:')
        this.gEvent.window.ddl_winTitleMatchMode := this.gEvent.Add('DropDownList', 'xs+' this.gEvent.editWidth - 65 ' yp-2 w65 +Hidden', this.aWinTitleMatchMode)
        this.gEvent.Tips.SetTip(this.gEvent.window.ddl_winTitleMatchMode, this.mGuiCtrlTips['gEvent.window.ddl_winTitleMatchMode'])
        this.gEvent.window.txt_winMinMax := this.gEvent.Add('Text', 'xs +Hidden', 'WinMinMax:')
        this.gEvent.window.ddl_winMinMax := this.gEvent.Add('DropDownList', 'xs+' this.gEvent.editWidth - 65 ' yp-2 w65 +Hidden', this.aWinMinMax)
        this.gEvent.window.ddl_winMinMax.OnEvent('Change', this.gEvent_window_ddl_mode_Enable_Disable.Bind(this))
        this.gEvent.Tips.SetTip(this.gEvent.window.ddl_winMinMax, this.mGuiCtrlTips['gEvent.window.ddl_winMinMax'])
        this.gEvent.window.cb_winActive := this.gEvent.Add('CheckBox', 'xs yp+30 w85 h20 +Hidden', ' WinActive')
        this.gEvent.window.cb_detectHiddenWindows := this.gEvent.Add('CheckBox', 'x+20 w125 h20 +Hidden', ' Hidden Windows')
        this.gEvent.window.cb_detectHiddenWindows.OnEvent('Click', this.gEvent_RefreshLV.Bind(this)) 
        this.gEvent.Tips.SetTip(this.gEvent.window.cb_detectHiddenWindows, this.mGuiCtrlTips['gEvent.window.cb_detectHiddenWindows'])               
        this.gEvent.window.cb_winActive.OnEvent('Click', this.gEvent_window_ddl_mode_Enable_Disable.Bind(this))
        this.gEvent.Tips.SetTip(this.gEvent.window.cb_winActive, this.mGuiCtrlTips['gEvent.window.cb_winActive'])                       
        this.gEvent.window.gb_monWindow := this.gEvent.Add('GroupBox', 'xm ym+' this.gEvent.gbHeight + this.gEvent.gbHeightWin + this.gEvent.gbHeightNotif ' w' this.gEvent.gbWidth ' h' this.gEvent.gbHeightMon ' +Hidden cBlack')
        this.gEvent.window.txt_monitoring := this.gEvent.Add('Text', 'xp+10 yp+18 Section +Hidden', 'Monitoring:')
        this.gEvent.window.ddl_monitoring := this.gEvent.Add('DropDownList', 'xs+' this.gEvent.editWidth - 85 ' yp-2 w85 +Hidden', this.aMonitoringWindow)
        this.gEvent.window.ddl_monitoring.OnEvent('Change', this.gEvent_ddl_monitoring_CtrlChange.Bind(this))          
        this.DDLchoose(mEvent.Has('window') ? mEvent['monitoring'] : this.m['mUser']['monitoringWindow'], this.aMonitoringWindow, this.gEvent.window.ddl_monitoring, caseSensitive := false)                
        this.gEvent.Tips.SetTip(this.gEvent.window.ddl_monitoring, this.mGuiCtrlTips['gEvent.window.ddl_monitoring'])                
        this.gEvent.window.txt_period := this.gEvent.Add('Text', 'xs w55 +Hidden Section')
        
        switch addOrEdit {
            case 'add':
            {
                this.gEvent.window.ddl_period := this.gEvent.Add('DropDownList', 'xm+65 yp-2 w55 +Hidden', this.m['mUser']['monitoringWindow'] == 'timer' ? this.aPeriodTimerWindow : this.aDelayWinEvent)
                
                switch this.m['mUser']['monitoringWindow'] {
                    case 'timer':    this.DDLchoose(this.m['mUser']['periodTimerWindow'], this.aPeriodTimerWindow, this.gEvent.window.ddl_period)
                    case 'winEvent': this.DDLchoose(this.m['mUser']['delayWinEvent'], this.aDelayWinEvent, this.gEvent.window.ddl_period)
                }
            } 

            case 'edit':
            {
                if (mEvent.Has('window')) {
                    this.gEvent.window.ddl_period := this.gEvent.Add('DropDownList', 'xm+65 yp-2 w55 +Hidden', mEvent['monitoring'] == 'timer' ? this.aPeriodTimerWindow : this.aDelayWinEvent)
                    this.DDLchoose(mEvent['period'], mEvent['monitoring'] == 'timer' ? this.aPeriodTimerWindow : this.aDelayWinEvent, this.gEvent.window.ddl_period)
                } 
                else {
                    this.gEvent.window.ddl_period := this.gEvent.Add('DropDownList', 'xm+65 yp-2 w55 +Hidden', this.m['mUser']['monitoringWindow'] == 'timer' ? this.aPeriodTimerWindow : this.aDelayWinEvent)
                    
                    switch this.m['mUser']['monitoringWindow'] {
                        case 'timer':    this.DDLchoose(this.m['mUser']['periodTimerWindow'], this.aPeriodTimerWindow, this.gEvent.window.ddl_period)
                        case 'winEvent': this.DDLchoose(this.m['mUser']['delayWinEvent'], this.aDelayWinEvent, this.gEvent.window.ddl_period)
                    }
                }
            } 
        }

        this.gEvent.window.txt_ms := this.gEvent.Add('Text', 'x+5 yp+2 +Hidden', 'ms')
        this.gEvent.window.txt_mode := this.gEvent.Add('Text', 'xs+150 ys +Hidden', 'Mode:')      
        this.gEvent.window.ddl_mode := this.gEvent.Add('DropDownList', 'xs+' this.gEvent.editWidth - 40 ' yp-2 w40 +Hidden', this.aModeWindow)
        this.gEvent.Tips.SetTip(this.gEvent.window.ddl_mode, this.mGuiCtrlTips['gEvent.window.ddl_mode'])

        switch addOrEdit {
            case 'add':  this.DDLchoose(this.m['mUser']['modeWindow'], this.aModeWindow, this.gEvent.window.ddl_mode)
            case 'edit': 
            {
                if mEvent.Has('window') 
                    this.DDLchoose(mEvent['mode'], this.aModeWindow, this.gEvent.window.ddl_mode)
                else
                    this.DDLchoose(this.m['mUser']['modeWindow'], this.aModeWindow, this.gEvent.window.ddl_mode)
            }
        }

        if this.gEvent.window.ddl_monitoring.Text = 'winEvent'
            this.gEvent.window.txt_period.Text := 'Delay: ≈'
        else
            this.gEvent.window.txt_period.Text := 'Period:'

        for _, value in ['winTitleMatchMode', 'winMinMax']
            this.DDLchoose(mEvent.Has('window') ? mEvent['window'][value] : this.m['mUser'][value], this.a%value%, this.gEvent.window.ddl_%value%)
        
        for _, value in ['detectHiddenWindows', 'winActive'] {
            if mEvent.Has('window')
                this.gEvent.window.cb_%value%.Value := mEvent['window'][value]
            else
                this.gEvent.window.cb_%value%.Value := this.m['mUser'][value]
        }

        if this.gEvent.window.cb_winActive.Value = 1 || RegExMatch(this.gEvent.window.ddl_winMinMax.Text, '^(0|1|-1)$')  
            ControlSetEnabled(0, this.gEvent.window.ddl_mode.hwnd)

        for _, value in ['winTitle', 'winClass', 'processName', 'processPath']
            this.gEvent.window.edit_%value%.OnEvent('Change', this.gEvent_RefreshLV.Bind(this))

        for _, value in ['winTitleMatchMode', 'winMinMax']
            this.gEvent.window.ddl_%value%.OnEvent('Change', this.gEvent_RefreshLV.Bind(this))

        ;==============================================
        this.gEvent.process := Object()
        this.gEvent.gbHeightProc := 128
        this.gEvent.process.gb_process := this.gEvent.Add('GroupBox', 'xm ym+' this.gEvent.gbHeight + this.gEvent.gbHeightNotif ' w' this.gEvent.gbWidth ' h' this.gEvent.gbHeightProc ' +Hidden cBlack')
        this.gEvent.process.txt_processName := this.gEvent.Add('Text', 'xp+10 yp+15 Section +Hidden', 'Process Name:')                              
        this.gEvent.process.edit_processName := this.gEvent.Add('Edit', 'xs w' this.gEvent.editWidth ' yp+21 h23 -Wrap +Hidden', mEvent.Has('process') ? mEvent['process']['processName'] : '')
        this.gEvent.process.txt_processPath := this.gEvent.Add('Text', 'xs +Hidden', 'Process Path:')
        this.gEvent.process.edit_processPath := this.gEvent.Add('Edit', 'xs w' this.gEvent.editWidth ' h23 -Wrap +Hidden', mEvent.Has('process') ? mEvent['process']['processPath'] : '')
               
        this.gEvent.gbHeightMon2 := 105
        this.gEvent.process.gb_monProcess := this.gEvent.Add('GroupBox', 'xm ym+' this.gEvent.gbHeight + this.gEvent.gbHeightProc + this.gEvent.gbHeightNotif ' w185 h' this.gEvent.gbHeightMon2 ' +Hidden cBlack')
        this.gEvent.process.txt_monitoring := this.gEvent.Add('Text', 'xp+10 yp+18 Section +Hidden', 'Monitoring:')
        this.gEvent.process.ddl_monitoring := this.gEvent.Add('DropDownList', 'xm+85 yp-2 w65 +Hidden', this.aMonitoringProcess)
        this.gEvent.process.ddl_monitoring.OnEvent('Change', this.gEvent_ddl_monitoring_CtrlChange.Bind(this))
        this.DDLchoose(mEvent.Has('process') ? mEvent['monitoring'] : this.m['mUser']['monitoringProcess'], this.aMonitoringProcess, this.gEvent.process.ddl_monitoring, caseSensitive := false)
        this.gEvent.Tips.SetTip(this.gEvent.process.ddl_monitoring, this.mGuiCtrlTips['gEvent.process.ddl_monitoring'])
        this.gEvent.process.txt_period := this.gEvent.Add('Text', 'xs +Hidden', 'Period:')
        
        switch addOrEdit {
            case 'add':
            {
                this.gEvent.process.ddl_period := this.gEvent.Add('DropDownList', 'xm+85 yp-2 w65 +Hidden', this.m['mUser']['monitoringProcess'] == 'timer' ? this.aPeriodTimerProcess : this.aPeriodWMIprocess)
                
                switch this.m['mUser']['monitoringProcess'] {
                    case 'timer': this.DDLchoose(this.m['mUser']['periodTimerProcess'], this.aPeriodTimerProcess, this.gEvent.process.ddl_period)
                    case 'wmi':   this.DDLchoose(this.m['mUser']['periodWMIprocess'],   this.aPeriodWMIprocess,   this.gEvent.process.ddl_period)
                }
            } 
            
            case 'edit':
            {
                if (mEvent.Has('process')) {
                    this.gEvent.process.ddl_period := this.gEvent.Add('DropDownList', 'xm+85 yp-2 w65 +Hidden', mEvent['monitoring'] == 'timer' ? this.aPeriodTimerProcess : this.aPeriodWMIprocess)
                    this.DDLchoose(mEvent['period'], mEvent['monitoring'] == 'timer' ? this.aPeriodTimerProcess : this.aPeriodWMIprocess, this.gEvent.process.ddl_period)
                }
                else {
                    this.gEvent.process.ddl_period := this.gEvent.Add('DropDownList', 'xm+85 yp-2 w65 +Hidden', this.m['mUser']['monitoringProcess'] == 'timer' ? this.aPeriodTimerProcess : this.aPeriodWMIprocess)

                    switch this.m['mUser']['monitoringProcess'] {
                        case 'timer': this.DDLchoose(this.m['mUser']['periodTimerProcess'], this.aPeriodTimerProcess, this.gEvent.process.ddl_period)
                        case 'wmi':   this.DDLchoose(this.m['mUser']['periodWMIprocess'],   this.aPeriodWMIprocess,   this.gEvent.process.ddl_period)
                    }
                }
            }
        } 

        this.gEvent.process.txt_ms := this.gEvent.Add('Text', 'x+5 yp+6 +Hidden', 'ms')
        this.gEvent.process.txt_mode := this.gEvent.Add('Text', 'xs +Hidden', 'Mode:')
        this.gEvent.process.ddl_mode := this.gEvent.Add('DropDownList', 'xm+85 yp-2 w65 +Hidden', this.aModeProcess)
        this.gEvent.Tips.SetTip(this.gEvent.process.ddl_mode, this.mGuiCtrlTips['gEvent.process.ddl_mode'])

        switch addOrEdit {
            case 'add':  this.DDLchoose(this.m['mUser']['modeProcess'], this.aModeProcess, this.gEvent.process.ddl_mode)
            case 'edit': 
            {
                if mEvent.Has('process') 
                    this.DDLchoose(mEvent['mode'], this.aModeProcess, this.gEvent.process.ddl_mode)
                else
                    this.DDLchoose(this.m['mUser']['modeProcess'], this.aModeProcess, this.gEvent.process.ddl_mode)
            }
        }

        if this.gEvent.process.ddl_monitoring.Text = 'wmi'
            ControlSetEnabled(0, this.gEvent.process.ddl_period.hwnd)

        for _, value in ['processName', 'processPath']
            this.gEvent.process.edit_%value%.OnEvent('Change', this.gEvent_RefreshLV.Bind(this))

        ;==============================================
        this.gEvent.device := Object()
        this.gEvent.device.gb_device := this.gEvent.Add('GroupBox', 'xm ym+' this.gEvent.gbHeight + this.gEvent.gbHeightNotif ' w' this.gEvent.gbWidth ' h' this.gEvent.gbHeightProc ' +Hidden cBlack')
        this.gEvent.device.txt_deviceName := this.gEvent.Add('Text', 'xp+10 yp+15 Section +Hidden', 'Device Name:')                            
        this.gEvent.device.edit_deviceName := this.gEvent.Add('Edit', 'xs w' this.gEvent.editWidth ' yp+21 h23 -Wrap +Hidden', mEvent.Has('device') ? mEvent['device']['deviceName'] : '')
        this.gEvent.device.txt_deviceID := this.gEvent.Add('Text', 'xs +Hidden', 'Device ID:')
        this.gEvent.device.edit_deviceID := this.gEvent.Add('Edit', 'xs w' this.gEvent.editWidth ' h23 -Wrap +Hidden', mEvent.Has('device') ? mEvent['device']['deviceId'] : '')
        this.gEvent.device.gb_monDevice := this.gEvent.Add('GroupBox', 'xm ym+' this.gEvent.gbHeight + this.gEvent.gbHeightProc + this.gEvent.gbHeightNotif ' w210 h' this.gEvent.gbHeightMon2 ' +Hidden cBlack')
        this.gEvent.device.txt_monitoring := this.gEvent.Add('Text', 'xp+10 yp+18 Section +Hidden', 'Monitoring:')
        this.gEvent.device.ddl_monitoring := this.gEvent.Add('DropDownList', 'xm+85 yp-2 w115 +Hidden', this.aMonitoringDevice)
        this.gEvent.device.ddl_monitoring.OnEvent('Change', this.gEvent_ddl_monitoring_CtrlChange.Bind(this))        
        this.DDLchoose(mEvent.Has('device') ? mEvent['monitoring'] : this.m['mUser']['monitoringDevice'], this.aMonitoringDevice, this.gEvent.device.ddl_monitoring, caseSensitive := false)
        this.gEvent.Tips.SetTip(this.gEvent.device.ddl_monitoring, this.mGuiCtrlTips['gEvent.device.ddl_monitoring'])
        this.gEvent.device.txt_period := this.gEvent.Add('Text', 'xs w70 +Hidden')

        switch addOrEdit {
            case 'add':
            {
                this.gEvent.device.ddl_period := this.gEvent.Add('DropDownList', 'xm+85 yp-2 w65 +Hidden', this.m['mUser']['monitoringDevice'] == 'timer' ? this.aPeriodTimerDevice : this.aDelayDeviceChange)
            
                switch this.m['mUser']['monitoringDevice'] {
                    case 'timer':        this.DDLchoose(this.m['mUser']['periodTimerDevice'], this.aPeriodTimerDevice, this.gEvent.device.ddl_period)
                    case 'deviceChange': this.DDLchoose(this.m['mUser']['delayDeviceChange'], this.aDelayDeviceChange, this.gEvent.device.ddl_period)
                }
            }
            
            case 'edit':
            {
                if (mEvent.Has('device')) {
                    this.gEvent.device.ddl_period := this.gEvent.Add('DropDownList', 'xm+85 yp-2 w65 +Hidden', mEvent['monitoring'] == 'timer' ? this.aPeriodTimerDevice : this.aDelayDeviceChange)
                    this.DDLchoose(mEvent['period'], mEvent['monitoring'] == 'timer' ? this.aPeriodTimerDevice : this.aDelayDeviceChange, this.gEvent.device.ddl_period)
                }
                else {
                    this.gEvent.device.ddl_period := this.gEvent.Add('DropDownList', 'xm+85 yp-2 w65 +Hidden', this.m['mUser']['monitoringDevice'] == 'timer' ? this.aPeriodTimerDevice : this.aDelayDeviceChange)
                    
                    switch this.m['mUser']['monitoringDevice'] {
                        case 'timer':        this.DDLchoose(this.m['mUser']['periodTimerDevice'], this.aPeriodTimerDevice, this.gEvent.device.ddl_period)
                        case 'deviceChange': this.DDLchoose(this.m['mUser']['delayDeviceChange'], this.aDelayDeviceChange, this.gEvent.device.ddl_period)
                    }            
                }
            }
        }

        this.gEvent.device.txt_ms := this.gEvent.Add('Text', 'x+5 yp+6 +Hidden', 'ms')
        this.gEvent.device.txt_mode := this.gEvent.Add('Text', 'xs +Hidden', 'Mode:')
        this.gEvent.device.ddl_mode := this.gEvent.Add('DropDownList', 'xm+85 yp-2 w65 +Hidden', this.aModeDevice)
        this.gEvent.Tips.SetTip(this.gEvent.device.ddl_mode, this.mGuiCtrlTips['gEvent.device.ddl_mode'])

        switch addOrEdit {
            case 'add':  this.DDLchoose(this.m['mUser']['modeDevice'], this.aModeDevice, this.gEvent.device.ddl_mode)
            case 'edit': 
            {
                if mEvent.Has('device') 
                    this.DDLchoose(mEvent['mode'], this.aModeDevice, this.gEvent.device.ddl_mode)
                else
                    this.DDLchoose(this.m['mUser']['modeDevice'], this.aModeDevice, this.gEvent.device.ddl_mode)
            }
        }

        if (this.gEvent.device.ddl_monitoring.Text = 'deviceChange') {
            this.gEvent.device.txt_period.Text := 'Delay:      ≈'
            ControlSetEnabled(0, this.gEvent.device.ddl_period.hwnd)
        } else
            this.gEvent.device.txt_period.Text := 'Period:'

        this.gEvent.device.btn_deviceInfoFinder := this.gEvent.Add('Button', 'xm yp+50 +Hidden', 'DeviceInfoFinder')
        this.gEvent.device.btn_deviceInfoFinder.OnEvent('Click', this.RunTool.Bind(this, this.mTools['deviceInfoFinder']['path']))

        for _, value in ['deviceName', 'deviceId']
            this.gEvent.device.edit_%value%.OnEvent('Change', this.gEvent_RefreshLV.Bind(this))

        this.gEvent.btnWidth := 120
        this.gEvent.btnHeight := 35
        this.gEvent.btnsWidth := this.gEvent.btnWidth*4 + this.gEvent.MarginX*3
        this.gEvent.width := this.gEvent.gbWidth + this.gEvent.lvWidth + this.gEvent.MarginX*3
        this.gEvent.btnsPosX := this.gEvent.width- this.gEvent.btnsWidth - this.gEvent.MarginX
        
        this.gEvent.lvOpt := 'xm+' this.gEvent.gbWidth + this.gEvent.MarginX ' ym+' this.gEvent.MarginY*3 ' w' this.gEvent.lvWidth ' h' this.gEvent.lvHeight ' Grid Sort Hidden Section +BackgroundEAF4FB'
              
        this.gEvent.window.lv := this.gEvent.Add('ListView', this.gEvent.lvOpt, ['WinTitle', 'WinClass', 'Process Name', 'Process Path', 'ID', 'PID', 'Elevated'])
        this.gEvent.process.lv := this.gEvent.Add('ListView', this.gEvent.lvOpt, ['Process Name', 'Process Path', 'PID', 'Elevated', 'Command line'])
        this.gEvent.device.lv := this.gEvent.Add('ListView', this.gEvent.lvOpt, ['Device Name', 'Device ID'])

        ;==============================================
        this.gEvent.btn_refreshLV := this.gEvent.Add('Button', 'x' this.gEvent.btnsPosX ' ym+' this.gEvent.gbHeight + this.gEvent.gbHeightWin + this.gEvent.gbHeightNotif + this.gEvent.gbHeightMon - this.gEvent.btnHeight - this.gEvent.MarginY ' w' this.gEvent.btnWidth ' h' this.gEvent.btnHeight, 'Refresh List')
        this.gEvent.btn_refreshLV.OnEvent('Click', this.gEvent_RefreshLV.Bind(this))
        GuiButtonIcon(this.gEvent.btn_refreshLV, this.mThemeIcons[this.m['mUser']['themeName']]['reload'], 1, 's16 a0 l10')
        this.gEvent.btn_resetDefault := this.gEvent.Add('Button', 'x+m w' this.gEvent.btnWidth ' h' this.gEvent.btnHeight, 'Reset to Default')
        this.gEvent.btn_resetDefault.OnEvent('Click', this.gEvent_btn_resetDefault_Click.Bind(this))
        this.gEvent.btn_ok := this.gEvent.Add('Button', 'x+m w' this.gEvent.btnWidth ' h' this.gEvent.btnHeight ' Default', 'OK')
        this.gEvent.btn_ok.OnEvent('Click', this.gEvent_btn_ok_Click.Bind(this, eventName, addOrEdit))
        this.gEvent.btn_cancel := this.gEvent.Add('Button', 'x+m w' this.gEvent.btnWidth ' h' this.gEvent.btnHeight, 'Cancel')
        this.gEvent.btn_cancel.OnEvent('Click', this.gEvent_Close.Bind(this))
        this.gEvent.OnEvent('Close', this.gEvent_Close.Bind(this))

        for _, value in ['window', 'process', 'device'] {
            this.gEvent.%value%.lv.OnEvent('DoubleClick', this.gEvent_lv_DoubleClick.Bind(this))
            this.gEvent.%value%.lv.OnEvent('ContextMenu', this.gEvent_lv_ContextMenu.Bind(this))
            LV_GridColor(this.gEvent.%value%.lv, '0xcbcbcb')
        }

        for key, value in this.gEvent.%eventType%.OwnProps()
            this.gEvent.%eventType%.%key%.Visible := True

        this.gEvent.gb_notif.SetFont('bold')
        this.gEvent_RefreshLV()
        this.GUI_ShowOnSameDisplay(this.gEvent, WinExist(this.gMainTitle ' ahk_class AutoHotkeyGUI'))

        if this.m['mUser']['alwaysOnTop']
            WinSetAlwaysOnTop(1, this.gEventTitle ' ahk_class AutoHotkeyGUI')

        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gEvent_Close(*)
    {
        this.gMain.Opt('-Disabled') 
        this.gEvent.Destroy()
    }

    ;============================================================================================

    static gEvent_Size(guiObj, minMax, width, height)
    {
        if minMax = -1
            return

        this.MoveControls(        
            {Control:this.gEvent.window.lv, h:height - this.gEvent.btnHeight - this.gEvent.MarginY*8, w:width - this.gEvent.gbWidth - this.gEvent.MarginX*3},            
            {Control:this.gEvent.process.lv, h:height - this.gEvent.btnHeight - this.gEvent.MarginY*8, w:width - this.gEvent.gbWidth - this.gEvent.MarginX*3},
            {Control:this.gEvent.device.lv, h:height - this.gEvent.btnHeight - this.gEvent.MarginY*8, w:width - this.gEvent.gbWidth - this.gEvent.MarginX*3},                          
            {Control:this.gEvent.btn_help, x:width - 24 - this.gEvent.MarginX},
            {Control:this.gEvent.btn_refreshLV, x:width - this.gEvent.btnWidth*4 - this.gEvent.MarginX*4, y:height - this.gEvent.btnHeight - this.gEvent.MarginY*2},
            {Control:this.gEvent.btn_resetDefault, x:width - this.gEvent.btnWidth*3 - this.gEvent.MarginX*3, y:height - this.gEvent.btnHeight - this.gEvent.MarginY*2},
            {Control:this.gEvent.btn_ok, x:width - this.gEvent.btnWidth*2 - this.gEvent.MarginX*2, y:height - this.gEvent.btnHeight - this.gEvent.MarginY*2},
            {Control:this.gEvent.btn_cancel, x:width - this.gEvent.btnWidth - this.gEvent.MarginX, y:height - this.gEvent.btnHeight - this.gEvent.MarginY*2},        
        )

        DllCall('RedrawWindow', 'ptr', this.gEvent.hwnd, 'ptr', 0, 'ptr', 0, 'uint', 0x0081)
    }

    ;============================================================================================

    static gEvent_btn_ok_Click(objEventName, addOrEdit, *)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        this.gMain.Opt('+Disabled')
                
        m := Map()       
        eventType := this.gEvent_GetRadioEventType()
        m[eventType] := Map()
        m['monitoring'] := this.FormatTxt(this.gEvent.%eventType%.ddl_monitoring.Text)
        
        for _, value in ['eventName', 'function']
            m[value] := Trim(this.gEvent.edit_%value%.Text)

        for _, value in ['period', 'mode']
            m[value] := this.gEvent.%eventType%.ddl_%value%.Text  

        for _, value in ['critical', 'log', 'notifInfo', 'notifStatus']
            m[value] := this.gEvent.cb_%value%.Value 
        
        for _, value in ['soundCreated', 'soundTerminated']
            m[value] := this.gEvent.ddl_%value%.Text
        
        switch {
            case m.Has('window'): 
            {
                for _, value in ['winTitle', 'winClass', 'processName', 'processPath']
                    m['window'][value] := this.gEvent.window.edit_%value%.Text

                for _, value in ['winTitleMatchMode', 'winMinMax'] 
                    m['window'][value] := this.gEvent.window.ddl_%value%.Text

                for _, value in ['winActive', 'detectHiddenWindows'] 
                    m['window'][value] := this.gEvent.window.cb_%value%.Value                
            }

            case m.Has('process'):
            {
                for _, value in ['processName', 'processPath']
                    m['process'][value] := this.gEvent.process.edit_%value%.Text   
            }

            case m.Has('device'):
            {
                for _, value in ['deviceName', 'deviceId']
                    m['device'][value] := this.gEvent.device.edit_%value%.Text
            }
        }    

        ;==============================================
        switch {
            case m['eventName'] !== objEventName && this.mTmp['mEvents'].Has(m['eventName']):
            {
                this.MsgBox('Two Events cannot have the same name.', this.gErrorTitle, this.mIcons['gX'],, 
                this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true
            }

            case !m['eventName']: 
            {  
                this.MsgBox('An Event name is required.', this.gErrorTitle, this.mIcons['gX'],,
                this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true                     
            }
            
            case (m['eventName'] && !this.IsValidEventName(m['eventName'])):
            {
                this.MsgBox('Invalid Event name. It cannot contain special characters and must be limited to 50 characters in length.', this.gErrorTitle, this.mIcons['gX'],,
                this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true  
            }

            case !m['function']:  
            {              
                this.MsgBox('A function name is required.', this.gErrorTitle, this.mIcons['gX'],,
                this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true
            }

            case (m['function'] && !this.IsValidFunctionName(m['function'])):
            {
                this.MsgBox('Invalid function name. The function name must begin with an alphabetic character, can include letters, digits, underscores and'
                . ' must be limited to 50 characters in length.', this.gErrorTitle, this.mIcons['gX'],,
                this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true
            }

            case (m.Has('process') && (!m['process']['processName'] && !m['process']['processPath'])):
            {
                this.MsgBox('At least one of these input fields must be filled:`n"Process Name" or "Process Path".', this.gErrorTitle, this.mIcons['gX'],,
                this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true                
            }

            case (m.Has('window') && (!m['window']['winTitle'] && !m['window']['winClass'] && !m['window']['processName'] && !m['window']['processPath'])):
            {
                this.MsgBox('At least one of these input fields must be filled:`n"WinTitle", "WinClass", "Process Name" or "Process Path".', this.gErrorTitle, this.mIcons['gX'],,
                this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true                
            }

            case (m.Has('device') && (!m['device']['deviceName'] && !m['device']['deviceId'])):
            {
                this.MsgBox('At least one of these input fields must be filled:`n"Device Name" or "Device ID".', this.gErrorTitle, this.mIcons['gX'],,
                this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true                 
            }
        }

        if (IsSet(returnFlag)) {
            this.SetCritical_Off(wasNotCritical)
            return
        }

        ;==============================================
        ; Warning, if another event have the same values.
        for eventName, mEvent in this.mTmp['mEvents'] {
            if eventName !== m['eventName'] 
            && eventName !== objEventName 
            && this.GetEventType(mEvent) == eventType {
                switch {                  
                    case mEvent.Has('window'): 
                    {
                        if mEvent['window']['winTitle']    == m['window']['winTitle'] 
                        && mEvent['window']['winClass']    == m['window']['winClass']
                        && mEvent['window']['processName'] == m['window']['processName'] 
                        && mEvent['window']['processPath'] == m['window']['processPath']
                        && mEvent['window']['winActive']   =  m['window']['winActive'] 
                        && mEvent['window']['winMinMax']   =  m['window']['winMinMax'] 
                        {
                            this.MsgBox('Two events cannot have the same values.`nEvent: "' eventName '" contains the same "WinTitle", "WinClass", "Process Name", "Process Path", "WinActive" and "WinMinMax".', 
                            this.gErrorTitle, this.mIcons['gX'],, this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                            [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true                            
                            break  
                        }
                    }

                    case mEvent.Has('process'):
                    {
                        if mEvent['process']['processName'] == m['process']['processName']
                        && mEvent['process']['processPath'] == m['process']['processPath'] 
                        {
                            this.MsgBox('Two events cannot have the same values.`nEvent: "' eventName '" contains the same "Process Name" and "Process Path".',
                            this.gErrorTitle, this.mIcons['gX'],, this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                            [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true
                            break
                        }
                    } 

                    case mEvent.Has('device'):
                    {
                        if mEvent['device']['deviceName'] == m['device']['deviceName']
                        && mEvent['device']['deviceId']   == m['device']['deviceId'] 
                        {
                            this.MsgBox('Two events cannot have the same values.`nEvent: "' eventName '" contains the same "Device Name" and "Device ID".',
                            this.gErrorTitle, this.mIcons['gX'],, this.gEvent, this.m['mUser']['alwaysOnTop'] ? 1 : 0,,,
                            [{name: '*OK', callback: 'this.MsgBox_Destroy'}]), returnFlag := true                           
                            break  
                        }
                    }                         
                }
            }
        }

        if (IsSet(returnFlag)) {
            this.SetCritical_Off(wasNotCritical)
            return
        }

        ;==============================================
        strMode := this.FormatTxtMode(m)
        strOptions := this.FormatTxtOptions(m)
        strMonitoring := this.FormatTxt(m['monitoring'])
        strEventType := this.FormatTxt(eventType)
        m['lineNumber'] := this.GetEventFunctionLineNumber(m['function'])

        switch {
            case m['monitoring'] == 'timer':        period :=  m['period']
            case m['monitoring'] == 'winEvent':     period := this.m['mUser']['delayWinEvent']
            case m['monitoring'] == 'wmi':          period := this.m['mUser']['periodWMIprocess']
            case m['monitoring'] == 'deviceChange': period := this.m['mUser']['delayDeviceChange']
        }

        switch addOrEdit {
            case 'add': this.gMain.lv_event.Insert(1, 'Icon2 +Check', m['eventName'], strEventType, m['function'], m['lineNumber'], strMonitoring, period, strMode, strOptions)             
            case 'edit':
            {
                Loop this.gMain.lv_event.GetCount() {
                    if (objEventName == this.gMain.lv_event.GetText(index := A_Index, 1)) {
                        this.gMain.lv_event.Modify(index, 'Col1', m['eventName'])
                        this.gMain.lv_event.Modify(index, 'Col2', strEventType)
                        this.gMain.lv_event.Modify(index, 'Col3', m['function'])
                        this.gMain.lv_event.Modify(index, 'Col4', m['lineNumber'])
                        this.gMain.lv_event.Modify(index, 'Col5', strMonitoring)
                        this.gMain.lv_event.Modify(index, 'Col6', m['period'])
                        this.gMain.lv_event.Modify(index, 'Col7', strMode)
                        this.gMain.lv_event.Modify(index, 'Col8', strOptions)
                        break
                    }
                }
            }
        }            

        this.gMain.lv_event.Modify(0, '-Select')

        Loop this.gMain.lv_event.GetCount() {
            if (m['eventName'] == this.gMain.lv_event.GetText(A_Index, 1)) {
                this.gMain.lv_event.Modify(A_Index, 'Select')
                this.gMain.lv_event.Modify(A_Index, 'Vis')
                break
            }  
        }  
        
        this.gMain_lv_event_ModifyCol()

        ;==============================================
        
        if (addOrEdit == 'edit') {
            objEventNameState := this.mTmp['mEvents'][objEventName]['state']
            this.mTmp['mEvents'].Delete(objEventName)
        }
        
        this.mTmp['mEvents'][m['eventName']] := Map()
        this.mTmp['mEvents'][m['eventName']][eventType] := Map()
        
        switch addOrEdit {
            case 'add':  this.mTmp['mEvents'][m['eventName']]['state'] := 1
            case 'edit': this.mTmp['mEvents'][m['eventName']]['state'] := objEventNameState
        }

        for _, value in ['monitoring', 'mode', 'period', 'function', 'lineNumber', 'critical', 'log', 'notifInfo', 'notifStatus', 'soundCreated', 'soundTerminated']
            this.mTmp['mEvents'][m['eventName']][value] := m[value]  

        switch {
            case m.Has('window'):
            {
                for _, value in ['winTitle', 'winClass', 'processName', 'processPath']
                    this.mTmp['mEvents'][m['eventName']]['window'][value] := this.gEvent.window.edit_%value%.Text

                for _, value in ['winTitleMatchMode', 'winMinMax']
                    this.mTmp['mEvents'][m['eventName']]['window'][value] := this.gEvent.window.ddl_%value%.Text

                this.mTmp['mEvents'][m['eventName']]['window']['winActive'] := this.gEvent.window.cb_winActive.Value
                this.mTmp['mEvents'][m['eventName']]['window']['detectHiddenWindows'] := this.gEvent.window.cb_detectHiddenWindows.Value
            }  
            
            case m.Has('process'):
            {
                for _, value in ['processName', 'processPath']
                    this.mTmp['mEvents'][m['eventName']]['process'][value] := this.gEvent.process.edit_%value%.Text

                if !this.mTmp['mEvents'][m['eventName']]['process']['processName']
                &&  this.mTmp['mEvents'][m['eventName']]['process']['processPath']
                    this.mTmp['mEvents'][m['eventName']]['process']['processName'] := RegExReplace(this.mTmp['mEvents'][m['eventName']]['process']['processPath'], '.*\\', '')                
            }            
                
            case m.Has('device'):
            {
                for _, value in ['deviceName', 'deviceId']
                    this.mTmp['mEvents'][m['eventName']]['device'][value] := this.gEvent.device.edit_%value%.Text
            }
        }

        this.Create_gMainMenuLoadProfile()
        this.gMain_btn_apply_EnableDisable()
        this.gEvent_Close()
        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gEvent_btn_resetDefault_Click(*)
	{
        eventType := this.gEvent_GetRadioEventType()

        for _, value in ['critical', 'log', 'notifInfo', 'notifStatus']
            this.gEvent.cb_%value%.Value := this.m['mUser'][value] 

        this.DDLchoose(this.mDefault['soundCreated'], this.aSounds, this.gEvent.ddl_soundCreated)
        this.DDLchoose(this.mDefault['soundTerminated'], this.aSounds, this.gEvent.ddl_soundTerminated)

        switch eventType {
            case 'window':
            {
                this.gEvent.window.cb_detectHiddenWindows.value := this.m['mUser']['detectHiddenWindows']
                this.gEvent.window.cb_winActive.value := this.m['mUser']['winActive']
                this.DDLchoose(this.m['mUser']['winTitleMatchMode'], this.aWinTitleMatchMode, this.gEvent.window.ddl_winTitleMatchMode)
                this.DDLchoose(this.m['mUser']['winMinMax'], this.aWinMinMax, this.gEvent.window.ddl_winMinMax)
                this.DDLchoose(this.m['mUser']['monitoringWindow'], this.aMonitoringWindow, this.gEvent.window.ddl_monitoring, caseSensitive := false)
                this.DDLchoose(this.m['mUser']['modeWindow'], this.aModeWindow, this.gEvent.window.ddl_mode)

                if (this.m['mUser']['monitoringWindow'] == 'winEvent') {
                    this.gEvent.window.txt_period.Text := 'Delay: ≈'
                    this.DDLarrayChange_Choose(this.m['mUser']['delayWinEvent'], this.aDelayWinEvent, this.gEvent.window.ddl_period)
                }

                if (this.m['mUser']['monitoringWindow'] == 'timer') {
                    this.gEvent.window.txt_period.Text := 'Period:'
                    this.DDLarrayChange_Choose(this.m['mUser']['periodTimerWindow'], this.aPeriodTimerWindow, this.gEvent.window.ddl_period)
                }

                if this.m['mUser']['winActive'] = 1 || RegExMatch(this.m['mUser']['winMinMax'], '^(0|1|-1)$')
                    ControlSetEnabled(0, this.gEvent.window.ddl_mode)
                else
                    ControlSetEnabled(1, this.gEvent.window.ddl_mode)
            } 

            case 'process':
            {
                this.DDLchoose(this.m['mUser']['monitoringProcess'], this.aMonitoringProcess, this.gEvent.process.ddl_monitoring, caseSensitive := false)
                this.DDLchoose(this.m['mUser']['modeProcess'], this.aModeProcess, this.gEvent.process.ddl_mode)
                
                if (this.m['mUser']['monitoringProcess'] == 'wmi') {
                    this.DDLchoose(this.m['mUser']['periodWMIprocess'], this.aPeriodWMIprocess, this.gEvent.process.ddl_period)
                    ControlSetEnabled(0, this.gEvent.process.ddl_period)
                }

                if (this.m['mUser']['monitoringProcess'] == 'timer') {
                    this.DDLchoose(this.m['mUser']['periodTimerWindow'], this.aPeriodWMIprocess, this.gEvent.process.ddl_period)
                    ControlSetEnabled(1, this.gEvent.process.ddl_period)
                }
            }

            case 'device':
            {  
                this.DDLchoose(this.m['mUser']['monitoringDevice'], this.aMonitoringDevice, this.gEvent.device.ddl_monitoring, caseSensitive := false)
                this.DDLchoose(this.m['mUser']['modeDevice'], this.aModeDevice, this.gEvent.device.ddl_mode)
                
                if (this.m['mUser']['monitoringDevice'] == 'deviceChange') {
                    this.DDLarrayChange_Choose(this.m['mUser']['delayDeviceChange'], this.aDelayDeviceChange, this.gEvent.device.ddl_period)
                    ControlSetEnabled(0, this.gEvent.device.ddl_period)
                }

                if (this.m['mUser']['monitoringDevice'] == 'timer') {                  
                    this.DDLarrayChange_Choose(this.m['mUser']['periodTimerDevice'] , this.aPeriodTimerDevice, this.gEvent.device.ddl_period)
                    ControlSetEnabled(1, this.gEvent.device.ddl_period)
                }
            }
        }

        if eventType == 'window'
            this.gEvent_RefreshLV()
    }

    ;============================================================================================

    static gEvent_radio_eventType_Click(eventType, *)
	{
        for _, value in ['window', 'process', 'device']
            if eventType == value && this.gEvent.%value%.lv.Visible = True
                return

        aEventType := ['window', 'process', 'device']
        aEventType.RemoveAt(this.HasVal(eventType, aEventType))

        Loop aEventType.Length
            for key, value in this.gEvent.%aEventType[index := A_Index]%.OwnProps()
                this.gEvent.%aEventType[index]%.%key%.Visible := False

        for key, value in this.gEvent.%eventType%.OwnProps()
            this.gEvent.%eventType%.%key%.Visible := True

        this.gEvent_RefreshLV()
    }

    ;============================================================================================

    static gEvent_ddl_monitoring_CtrlChange(*)
    {
        eventType := this.gEvent_GetRadioEventType()
        monitoring := this.FormatTxt(this.gEvent.%eventType%.ddl_monitoring.Text)

        switch {
            case (RegExMatch(eventType, '^(window|process|device)$') && monitoring = 'timer'):
            {
                this.DDLarrayChange_Choose(this.m['mUser']['periodTimer' StrTitle(eventType)], this.aPeriodTimer%eventType%, this.gEvent.%eventType%.ddl_period)
                this.gEvent.%eventType%.txt_period.Text := 'Period:'
                ControlSetEnabled(1, this.gEvent.%eventType%.ddl_period)
            }

            case (eventType == 'process' && monitoring = 'wmi'):
            {
                this.DDLarrayChange_Choose(this.m['mUser']['periodWMIprocess'], this.aPeriodWMIprocess, this.gEvent.process.ddl_period)
                ControlSetEnabled(0, this.gEvent.process.ddl_period)
            }

            case (eventType == 'window' && monitoring = 'winEvent'):
            {
                this.DDLarrayChange_Choose(this.m['mUser']['delayWinEvent'], this.aDelayWinEvent, this.gEvent.window.ddl_period)
                this.gEvent.window.txt_period.Text := 'Delay: ≈'
            }

            case (eventType == 'device' && monitoring = 'deviceChange'):
            {
                this.DDLarrayChange_Choose(this.m['mUser']['delayDeviceChange'], this.aDelayDeviceChange, this.gEvent.device.ddl_period)
                this.gEvent.device.txt_period.Text := 'Delay:      ≈'
                ControlSetEnabled(0, this.gEvent.device.ddl_period)
            }
        }
    }    
    
    ;============================================================================================
    
    static DDL_Sound_CtrlChange(gName, DDLname, *) => SetTimer( this.Sound.Bind(this, this.mSounds[this.%gName%.ddl_%DDLname%.Text]) , -1 )

    ;============================================================================================

    static gEvent_lv_ContextMenu(ctrlObj, item, isRightClick, x, y)
    {
        MouseGetPos(,,, &mouseOverClassNN)
        if item = 0 || InStr(mouseOverClassNN, 'SysHeader')
            return

        ctxMenu := Menu()
        ctxMenu.Add('Copy row data to input Fields', this.gEvent_FillEditField.Bind(this, ctrlObj, item))
        ctxMenu.SetIcon('Copy row data to input Fields', this.mThemeIcons[this.m['mUser']['themeName']]['gMenuEdit'])       
        ctxMenu.Show(x, y)
    }

    ;============================================================================================

    static gEvent_lv_DoubleClick(ctrlObj, item, *)
    {
        if item = 0
            return

        this.gEvent_FillEditField(ctrlObj, item)
    }  
    
    ;============================================================================================

    static gEvent_FillEditField(ctrlObj, item, *)
    {
        switch this.gEvent_GetRadioEventType() {
            case 'window':
            {
                for _, value in ['winTitle', 'winClass', 'processName', 'processPath']
                    this.gEvent.window.edit_%value%.Text := ctrlObj.GetText(item, A_Index)  
            }            
            case 'process':
            {
                for _, value in ['processName', 'processPath']                   
                    this.gEvent.process.edit_%value%.Text := ctrlObj.GetText(item, A_Index)
            }
            case 'device':
            {
                for _, value in ['deviceName', 'deviceId']
                    this.gEvent.device.edit_%value%.Text := ctrlObj.GetText(item, A_Index)
            }
        }

        this.gEvent_RefreshLV()
    }

    ;============================================================================================

    static gEvent_btn_clear_Click(*)
    {
        switch this.gEvent_GetRadioEventType() {
            case 'window':
            {
                for _, value in ['winTitle', 'winClass', 'processName', 'processPath']
                    this.gEvent.window.edit_%value%.Text := ''
            }          
            case 'process':
            {
                for _, value in ['processName', 'processPath']
                    this.gEvent.process.edit_%value%.Text := ''
            }
            case 'device':
            {
                for _, value in ['deviceName', 'deviceId']
                    this.gEvent.device.edit_%value%.Text := ''
            }
        }

        this.gEvent_RefreshLV()
    }

    ;============================================================================================

    static gEvent_window_ddl_mode_Enable_Disable(*)
    {
        if this.gEvent.window.ddl_winMinMax.Text = '' && this.gEvent.window.cb_winActive.Value = 0
            ControlSetEnabled(1, this.gEvent.window.ddl_mode)
        else
            ControlSetEnabled(0, this.gEvent.window.ddl_mode)

        this.gEvent_RefreshLV()
    }

    ;============================================================================================

    static gEvent_RefreshLV(*)
    {
        m := Map()
        this.gEvent.Opt('+Disabled')
        eventType := this.gEvent_GetRadioEventType()

        switch eventType {
            case 'process':
            {
                for _, value in ['processName', 'processPath']
                    m[value] := this.gEvent.process.edit_%value%.Text

                aProcessFinder := this.ProcessFinder(m['processName'], m['processPath'])
                this.gEvent.process.lv.Delete()
                imageList := IL_Create()
                this.gEvent.process.lv.SetImageList(imageList)

                for _, proc in aProcessFinder
                    this.gEvent.process.lv.Insert(1, 'Icon' IL_Add(imageList, (iconIndex := IL_Add(imageList, proc['processPath'])) ? proc['processPath'] : this.mIcons['exe']), proc['processName'], proc['processPath'], proc['pid'], proc['elevated'], proc['cmdLine'])

                for _, col in [1, 2, 4, 5]
                    this.gEvent.process.lv.ModifyCol(col, 'AutoHdr')                

                this.gEvent.process.lv.ModifyCol(3, '75 Integer')
            }

            case 'window':
            {
                for _, value in ['winTitle', 'winClass', 'processName', 'processPath']
                    m[value] := this.gEvent.window.edit_%value%.Text

                for _, value in ['winTitleMatchMode', 'winMinMax']
                    m[value] := this.gEvent.window.ddl_%value%.Text

                m['winActive'] := this.gEvent.window.cb_winActive.Value
                m['detectHiddenWindows'] := this.gEvent.window.cb_detectHiddenWindows.Value

                aWindowFinder := this.WindowFinder(m['winTitle'], m['winClass'], m['processName'], m['processPath'], m['winTitleMatchMode'], m['detectHiddenWindows'], m['winActive'], m['winMinMax'])
                this.gEvent.window.lv.Delete()
                imageList := IL_Create()
                this.gEvent.window.lv.SetImageList(imageList)

                for index, win in aWindowFinder
                    this.gEvent.window.lv.Insert(
                        1, 'Icon' IL_Add(imageList, (iconIndex := IL_Add(imageList, win['processPath'])) ? win['processPath'] : this.mIcons['window']), 
                        win['winTitle'], win['winClass'], win['processName'], win['processPath'], win['id'], win['pid'], win['elevated']
                    )

                for _, col in [1, 2, 3, 4, 7]
                    this.gEvent.window.lv.ModifyCol(col, 'AutoHdr')

                for _, col in [5, 6]
                    this.gEvent.window.lv.ModifyCol(col, '75 Integer')
            }

            case 'device':
            {
                for _, value in ['deviceName', 'deviceId']
                    m[value] := this.gEvent.device.edit_%value%.Text

                aDeviceFinder := this.DeviceFinder(m['deviceName'], m['deviceId'], 'getDeviceIconSmall')
                this.gEvent.device.lv.Delete()
                imageList := IL_Create()
                this.gEvent.device.lv.SetImageList(imageList)

                for _, dev in aDeviceFinder
                    this.gEvent.device.lv.Insert(1, 'Icon' IL_Add(imageList, 'HICON:' dev['hIconSmall']), dev['deviceName'], dev['deviceId'])

                loop 2
                    this.gEvent.device.lv.ModifyCol(A_Index, 'AutoHdr')
            }
        }

        this.gEvent.Opt('-Disabled')
    }

    ;============================================================================================

    static gEvent_GetRadioEventType()
    {
        for _, value in ['window', 'process', 'device']
            if this.gEvent.radio_%value%.Value = 1
                return value
    }

    ;============================================================================================

    static gSettings_Show(*)
    {
        if this.GUI_IfExistReturn_Destroy(['gEvent', 'gSettings', 'gAbout', 'gProfile', 'gDelProfile', 'gDelConfirm'], 'ifExistReturn')
            return        

        wasNotCritical := this.SetCritical_On(A_IsCritical)

        this.gSettings := Gui('-MinimizeBox' , this.gSettingsTitle)
        this.gSettings.SetFont('s10')
        
        try this.gMain.Opt('+Disabled')
        try this.gSettings.Opt('+Owner' this.gMain.hwnd)        
        
        gSettingsHeight := 540
        gbWidth := 170
        gbHeightRow1 := 305
        
        this.gSettings.gb_window := this.gSettings.Add('GroupBox', 'w' gbWidth ' h' gbHeightRow1 ' cBlack', 'Window')  
        this.gSettings.Add('Text', 'xp+10 yp+25 Section ', 'Default monitoring:')
        this.gSettings.ddl_monitoringWindow := this.gSettings.Add('DropDownList', 'xs w80', this.aMonitoringWindow)
        this.gSettings.Add('Text', 'xs', 'Default period timer:')
        this.gSettings.ddl_periodTimerWindow := this.gSettings.Add('DropDownList', 'xs w65', this.aPeriodTimerWindow)
        this.gSettings.Add('Text', 'x+5 yp+7', 'ms')
        this.gSettings.Add('Text', 'xs', 'Default mode:')
        this.gSettings.ddl_modeWindow := this.gSettings.Add('DropDownList', 'xs w65', this.aModeWindow)
        this.gSettings.Add('Text', 'xs', 'Default delay ≈ :')
        this.gSettings.ddl_delayWinEvent := this.gSettings.Add('DropDownList', 'xs w65', this.aDelayWinEvent)
        this.gSettings.Add('Text', 'x+5 yp+7', 'ms')
        this.gSettings.Add('Text', 'xs', 'dwFlags WinEvent:')
        this.gSettings.ddl_dwFlagsWinEvent := this.gSettings.Add('DropDownList', 'xs w65', this.aDWflagsWinEvent)
        ;==============================================
        this.gSettings.gb_process := this.gSettings.Add('GroupBox', 'xm+' gbWidth + this.gSettings.MarginX ' ym w' gbWidth ' h' gbHeightRow1 ' cBlack', 'Process')
        this.gSettings.Add('Text', 'xp+10 yp+25 Section', 'Default monitoring:')
        this.gSettings.ddl_monitoringProcess := this.gSettings.Add('DropDownList', 'xs w65', this.aMonitoringProcess)
        this.gSettings.Add('Text', 'xs', 'Default period timer:')
        this.gSettings.ddl_periodTimerProcess := this.gSettings.Add('DropDownList', 'xs w65', this.aPeriodTimerProcess)
        this.gSettings.Add('Text', 'x+5 yp+7', 'ms')       
        this.gSettings.Add('Text', 'xs', 'Default mode:')
        this.gSettings.ddl_modeProcess := this.gSettings.Add('DropDownList', 'xs w65', this.aModeProcess)
        this.gSettings.btn_resetWMI := this.gSettings.Add('Button', 'xs yp+35 w145', 'Restart WMI Service*')
        this.gSettings.btn_resetWMI.OnEvent('Click', this.gSettings_btn_restartWMI_Click.Bind(this))
        ;==============================================
        this.gSettings.gb_device := this.gSettings.Add('GroupBox', 'xm+' gbWidth*2 + this.gSettings.MarginX*2 ' ym w' gbWidth ' h' gbHeightRow1 ' cBlack', 'Device')
        this.gSettings.Add('Text', 'xp+10 yp+25 Section', 'Default monitoring:')
        this.gSettings.ddl_monitoringDevice := this.gSettings.Add('DropDownList', 'xs w110', this.aMonitoringDevice)
        this.gSettings.Add('Text', 'xs', 'Default period timer:')
        this.gSettings.ddl_periodTimerDevice := this.gSettings.Add('DropDownList', 'xs w65', this.aPeriodTimerDevice)
        this.gSettings.Add('Text', 'x+5 yp+7', 'ms')
        this.gSettings.Add('Text', 'xs', 'Default mode:')
        this.gSettings.ddl_modeDevice := this.gSettings.Add('DropDownList', 'xs w65', this.aModeDevice)      
        this.gSettings.Add('Text', 'xs', 'Delay DeviceChange:')
        this.gSettings.ddl_delayDeviceChange := this.gSettings.Add('DropDownList', 'xs w65', this.aDelayDeviceChange)
        this.gSettings.Add('Text', 'x+5 yp+7', 'ms')
        ;==============================================        
        gbHeightRow2 := 140
        this.gSettings.gb_custom := this.gSettings.Add('GroupBox', 'xm ym+' gbHeightRow1 + this.gSettings.MarginY ' w' gbWidth ' h' gbHeightRow2 ' cBlack', 'Customization')
        this.gSettings.Add('Text', 'xp+10 yp+25 Section', 'Themes*')
        this.gSettings.ddl_themeName := this.gSettings.Add('DropDownList', 'xs w150', this.aThemeName)
        this.gSettings.Add('Text', 'xs', 'Tray menu icons size*')
        this.gSettings.ddl_trayMenuIconSize := this.gSettings.Add('DropDownList', 'xs w65', this.aTrayMenuIconSize)
        ;==============================================
        gbTrayIcoWidth := 240
        this.gSettings.gb_trayClick := this.gSettings.Add('GroupBox', 'xm+' gbWidth + this.gSettings.MarginX ' ym+' gbHeightRow1 + this.gSettings.MarginY ' w' gbTrayIcoWidth ' h' gbHeightRow2 ' cBlack', 'On tray icon')
        this.gSettings.Add('Text', 'xp+10 yp+25 Section', 'Left click:')
        this.gSettings.ddl_trayIconLeftClick := this.gSettings.Add('DropDownList', 'xs w220', this.Create_Array_aTrayIconLeftClick())
        this.gSettings.Add('Text', 'xs', 'Double left click:')
        this.gSettings.ddl_trayIconLeftDClick := this.gSettings.Add('DropDownList', 'xs w220',this.Create_Array_aTrayIconLeftClick())
        ;==============================================
        gbStartupWidth := 275
        this.gSettings.gb_startup := this.gSettings.Add('GroupBox', 'xm+' gbTrayIcoWidth + gbWidth + this.gSettings.MarginX*2 ' ym+' gbHeightRow1 + this.gSettings.MarginY ' w' gbStartupWidth ' h' gbHeightRow2 ' cBlack', 'Startup')
        this.gSettings.Add('Text', 'xp+10 yp+25 Section', 'To automatically run this script on startup,`nadd its shortcut to the startup folder.')
        this.gSettings.Add('Button', 'xs yp+38 w145', 'Open Script Folder').OnEvent('Click', (*) => Run(A_WorkingDir))
        this.gSettings.Add('Button', 'xs yp+35 w145', 'Open Startup Folder').OnEvent('Click', (*) => Run(A_Startup))

        btnWidth := 120
        btnHeight := 40
        btnsWidth := btnWidth*3 + this.gSettings.MarginX*2
        btnsPosX := gbWidth + gbTrayIcoWidth + gbStartupWidth - btnsWidth + this.gSettings.MarginX*3

        this.gSettings.btn_resetDefault := this.gSettings.Add('Button', 'x' btnsPosX ' y' gbHeightRow1 + gbHeightRow2 + this.gSettings.MarginY*4 ' w' btnWidth ' h' btnHeight, 'Reset to Default')
        this.gSettings.btn_resetDefault.OnEvent('Click', this.gSettings_btn_ResetDefault_Click.Bind(this))
        this.gSettings.btn_ok := this.gSettings.Add('Button', 'x+m w' btnWidth ' h' btnHeight ' Default', 'OK')
        this.gSettings.btn_ok.OnEvent('Click', this.gSettings_btn_ok_Click.Bind(this))
        this.gSettings.btn_cancel := this.gSettings.Add('Button', 'x+m w' btnWidth ' h' btnHeight, 'Cancel')
        this.gSettings.btn_cancel.OnEvent('Click', this.gSettings_Close.Bind(this))
        this.gSettings.OnEvent('Close', this.gSettings_Close.Bind(this))   
        this.gSettings.Add('Text', 'xm yp h15', '* Required reload')
        this.gSettings.MarginY := 15

        ;==============================================
        for _, value in ['monitoringProcess', 'monitoringWindow', 'monitoringDevice', 'periodTimerProcess', 'periodTimerWindow', 'periodTimerDevice', 
        'modeProcess', 'modeWindow', 'modeDevice', 'dwFlagsWinEvent', 'delayWinEvent', 'delayDeviceChange', 'trayMenuIconSize', 'themeName']
        {
            this.DDLchoose(this.m['mUser'][value], this.a%value%, this.gSettings.ddl_%value%, caseSensitive := false)
        }

        this.DDLchoose(this.m['mUser']['trayIconLeftClick'], this.Create_Array_aTrayIconLeftClick(), this.gSettings.ddl_trayIconLeftClick)
        this.DDLchoose(this.m['mUser']['trayIconLeftDClick'], this.Create_Array_aTrayIconLeftClick(), this.gSettings.ddl_trayIconLeftDClick)

        for _, value in ['window', 'process', 'device', 'trayClick', 'custom', 'startup']
            this.gSettings.%'gb_' value%.SetFont('bold')

        this.GUI_ShowOnSameDisplay(this.gSettings, WinExist(this.gMainTitle ' ahk_class AutoHotkeyGUI'))
        
        if this.m['mUser']['alwaysOnTop']
            WinSetAlwaysOnTop(1, this.gSettingsTitle ' ahk_class AutoHotkeyGUI')        
        
        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gSettings_Close(*)
    {  
        try this.gMain.Opt('-Disabled')
        this.gSettings.Destroy()
    }

    ;============================================================================================   
    
    static gSettings_btn_ok_Click(ctrlObj, *)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        this.gSettings.Opt('+Disabled')

        if ctrlObj.Text == 'OK'
            ControlSetEnabled(0, this.gSettings.btn_ok.hwnd)

        ;==============================================
        if (this.m['mUser']['delayDeviceChange'] != this.gSettings.ddl_delayDeviceChange.Text) 
        {
            this.m['mUser']['delayDeviceChange'] := this.gSettings.ddl_delayDeviceChange.Text

            for eventName, mEvent in this.m['mEvents']
                if mEvent.Has('device') && mEvent['monitoring'] == 'deviceChange'
                    mEvent['period'] := this.m['mUser']['delayDeviceChange']          

            if WinExist(this.gMainTitle ' ahk_class AutoHotkeyGUI') 
            {
                for eventName, mEvent in this.mTmp['mEvents']
                    if mEvent.Has('device') && mEvent['monitoring'] == 'deviceChange'
                        mEvent['period'] := this.m['mUser']['delayDeviceChange']

                Loop this.gMain.lv_event.GetCount()
                    if this.gMain.lv_event.GetText(A_Index , 5) = 'deviceChange'
                        this.gMain.lv_event.Modify(A_Index, "Col6", this.m['mUser']['delayDeviceChange'])
            }
        }

        ;==============================================
        for _, value in ['monitoringProcess', 'monitoringWindow', 'monitoringDevice']
        {
            if this.m['mUser'][value] !== (strValue := this.FormatTxt(this.gSettings.ddl_%value%.Text))
                this.m['mUser'][value] := strValue
        }

        for _, value in ['periodTimerProcess', 'periodTimerWindow', 'periodTimerDevice', 'modeProcess', 'modeWindow', 'modeDevice', 
        'dwFlagsWinEvent', 'delayWinEvent', 'delayDeviceChange', 'trayIconLeftClick', 'trayIconLeftDClick']
        {
            if this.m['mUser'][value] !== (strValue := this.gSettings.ddl_%value%.Text)
                this.m['mUser'][value] := strValue
        }

        for _, value in ['trayMenuIconSize', 'themeName'] {
            if (this.m['mUser'][value] !== (strValue := this.gSettings.ddl_%value%.Text)) {
                this.m['mUser'][value] := strValue
                reloadFlag := true
            }
        }

        this.SetMonitoringTrayClick(  (this.m['mUser']['trayIconLeftClick'] !== '- None -' || this.m['mUser']['trayIconLeftDClick'] !== '- None -')  )
        
        if IsSet(reloadFlag)
            Reload

        this.profileMatch := this.IsEventsValuesMatchProfile('m')
        this.Create_TrayMenuSelectProfile()
        this.SetIconTip()

        if WinExist(this.gMainTitle ' ahk_class AutoHotkeyGUI')
        {
            this.Create_gMainMenuLoadProfile()
            this.gMain_txt_SetTextBottom()
        }

        this.gSettings_Close()
        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================    

    static gSettings_btn_ResetDefault_Click(*)
    {
        for _, value in ['monitoringProcess', 'monitoringWindow', 'monitoringDevice', 'periodTimerProcess', 'periodTimerWindow', 'periodTimerDevice', 
        'modeProcess', 'modeWindow', 'modeDevice', 'dwFlagsWinEvent', 'delayWinEvent', 'delayDeviceChange', 'trayMenuIconSize'] 
        {
            this.DDLchoose(this.mDefault[value], this.a%value%, this.gSettings.ddl_%value%, caseSensitive := false)
        }        

        this.DDLchoose(this.mDefault['trayIconLeftClick'], this.Create_Array_aTrayIconLeftClick(), this.gSettings.ddl_trayIconLeftClick)
        this.DDLchoose(this.mDefault['trayIconLeftDClick'], this.Create_Array_aTrayIconLeftClick(), this.gSettings.ddl_trayIconLeftDClick)
    }

    ;============================================================================================

    static gSettings_btn_restartWMI_Click(*)
    {
        this.gSettings.Opt('+Disabled')
        this.TrayMenu_Enable_Disable('disable')

        try {
            RunWait('*RunAs Powershell.exe -Command "Restart-Service -Name winmgmt -Force"',, 'Hide')
        } 
        catch {
            this.TrayMenu_Enable_Disable('enable')
            this.gSettings.Opt('-Disabled')
            return
        }

        Reload
    }

    ;============================================================================================

    static gAbout_Show(*)
    {
        if this.GUI_IfExistReturn_Destroy(['gEvent', 'gSettings', 'gAbout', 'gProfile', 'gDelProfile', 'gDelConfirm'], 'ifExistReturn')
            return

        wasNotCritical := this.SetCritical_On(A_IsCritical)
        
        this.gAbout := Gui('-MinimizeBox', this.gAboutTitle)
        this.gAbout.BackColor := 'White'

        try this.gMain.Opt('+Disabled')
        try this.gAbout.Opt('+Owner' this.gMain.hwnd)

        this.gAbout.Add('Picture', 'y15 w50 h50', this.mThemeIcons[this.m['mUser']['themeName']]['gAbout'])
        this.gAbout.SetFont('s16')
        this.gAbout.Add('Text', 'x+m w350 Section', this.scriptName).SetFont(, 'Segoe UI bold')
        this.gAbout.SetFont('s10')
        this.gAbout.Add('Text', 'yp+40', 'Version: ' this.scriptVersion)
        this.gAbout.Add('Text', 'yp+25', 'Author:')
        this.gAbout.Add('Link', 'x+10', '<a href="https://github.com/XMCQCX">Martin Chartier (XMCQCX)</a>')
        this.gAbout.SetFont('s12')
        this.gAbout.Add('Text', 'xs yp+35 w125', 'Credits').SetFont('Bold')
        this.gAbout.SetFont('s10')
        this.gAbout.Add('Text', 'xs yp+40', 'Steve Gray, Chris Mallett, portions of AutoIt Team and various others.')
        this.gAbout.Add('Link', 'yp+20', '<a href="https://www.autohotkey.com">https://www.autohotkey.com</a>') 
        this.gAbout.Add('Text', 'xs yp+15', '_____________________________')
        this.gAbout.Add('Link', 'xs yp+25', '<a href="https://github.com/thqby/ahk2_lib/blob/master/JSON.ahk">JSON</a>')
        this.gAbout.Add('Text', 'x+5', 'by thqby, HotKeyIt.')
        this.gAbout.Add('Link', 'xs yp+30', '<a href="https://github.com/AHK-just-me/AHKv2_GuiCtrlTips">GuiCtrlTips</a>')
        this.gAbout.Add('Text', 'x+5', 'by just me.')
        this.gAbout.Add('Link', 'xs yp+30', '<a href="https://www.autohotkey.com/boards/viewtopic.php?f=83&t=125259">LVGridColor</a>')
        this.gAbout.Add('Text', 'x+5', 'by just me.')   
        this.gAbout.Add('Link', 'xs yp+30', '<a href="https://www.autohotkey.com/boards/viewtopic.php?f=83&t=115871">GuiButtonIcon</a>')
        this.gAbout.Add('Text', 'x+5', 'by FanaticGuru.')
        this.gAbout.Add('Link', 'xs yp+30', '<a href="https://www.autohotkey.com/boards/viewtopic.php?p=507896#p507896">DisplayObj</a>')
        this.gAbout.Add('Text', 'x+5', 'by FanaticGuru. (inspired by tidbit and Lexikos v1 code)')             
        this.gAbout.Add('Link', 'xs yp+30', '<a href="https://www.autohotkey.com/boards/viewtopic.php?t=121125&p=537515">EnumDeviceInfo</a>')
        this.gAbout.Add('Text', 'x+5', 'by teadrinker. (based on JEE_DeviceList by jeeswg)')
        this.gAbout.Add('Link', 'xs yp+30', '<a href="https://www.autohotkey.com/boards/viewtopic.php?p=526409#p526409">GetCommandLine</a>')
        this.gAbout.Add('Text', 'x+5', 'by teadrinker. (based on Sean and SKAN v1 code)')
        this.gAbout.Add('Link', 'xs yp+30', '<a href="https://www.autohotkey.com/boards/viewtopic.php?t=4365">WTSEnumProcesses</a>')
        this.gAbout.Add('Text', 'x+5', 'by SKAN.')
        this.gAbout.Add('Link', 'xs yp+30', '<a href="https://github.com/jNizM/ahk-scripts-v2/blob/main/src/ProcessThreadModule/IsProcessElevated.ahk">IsProcessElevated</a>')
        this.gAbout.Add('Text', 'x+5', 'by jNizM.')
        this.gAbout.Add('Link', 'xs yp+30', '<a href="https://github.com/Descolada/UIA-v2">MoveControls</a>')
        this.gAbout.Add('Text', 'x+5', 'by Descolada. (from UIATreeInspector.ahk)')
        this.gAbout.Add('Link', 'xs yp+30', '<a href="https://www.autohotkey.com/boards/viewtopic.php?f=6&t=29117&hilit=FrameShadow">FrameShadow</a>')
        this.gAbout.Add('Text', 'x+5', 'by Klark92.')   
        this.gAbout.Add('Text', 'xs yp+30', 'Notify is inspired by')
        this.gAbout.Add('Link', 'x+5', '<a href="https://www.the-automator.com/downloads/maestrith-notify-class-v2/">NotifyV2</a>')
        this.gAbout.Add('Text', 'x+5', '(from the-Automator.com)')
        
        gAboutWidth := 550
        btnWidth := 120
        btnHeight := 35
        btnPosX := gAboutWidth/2 - btnWidth/2

        this.gAbout.btn_close := this.gAbout.Add('Button', 'x' btnPosX ' yp+30 w' btnWidth ' h' btnHeight ' Default', 'Close')
        this.gAbout.btn_close.OnEvent('Click', this.gAbout_Close.Bind(this))
        this.gAbout.OnEvent('Close', this.gAbout_Close.Bind(this))
        this.gAbout.MarginY := 15
        
        try ControlFocus(this.gAbout.btn_close.hwnd, this.gAbout.hwnd)
        this.GUI_ShowOnSameDisplay(this.gAbout, WinExist(this.gMainTitle ' ahk_class AutoHotkeyGUI'),, 'w' gAboutWidth)

        if this.m['mUser']['alwaysOnTop']
            WinSetAlwaysOnTop(1, this.gAboutTitle ' ahk_class AutoHotkeyGUI')

        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================

    static gAbout_Close(*)
    {
        try this.gMain.Opt('-Disabled')
        this.gAbout.Destroy()
    }

    ;============================================================================================

    static SetMonitoringTrayClick(state)
    {
        if (state = 1 && this.monitoringTrayClick) || (state = 0 && !this.monitoringTrayClick)
            return

        if state = 1
            OnMessage(0x404, this.callbackNotifyIcon), this.monitoringTrayClick := true
        else
            OnMessage(0x404, this.callbackNotifyIcon, 0), this.monitoringTrayClick := false
    }

    ;============================================================================================

    static AHK_NOTIFYICON(wParam, lParam, msg, hwnd)
    {
        if (lParam = 0x201) {
            if KeyWait('LButton', 'D T0.25') ; double-click
                SetTimer(this.TrayIconClickAction.Bind(this, this.m['mUser']['trayIconLeftDClick']), -1)
            else
                SetTimer(this.TrayIconClickAction.Bind(this, this.m['mUser']['trayIconLeftClick']), -1)
        }
    }

    ;============================================================================================

    static TrayIconClickAction(action)
    {
        switch {
            case action == '- None -':
                return

            case action == 'Open Events Manager': this.gMain_Show()
            
            case RegExMatch(action, 'Toggle All Events Enable/Disable'): 
            {
                if this.monitoringWMIProcess || this.monitoringTimerProcess || this.monitoringTimerWindow || this.monitoringTimerDevice
                    monitoringFlag := true

                if IsSet(monitoringFlag) 
                    this.SetProfile('All Events Disable')
                else
                    this.SetProfile('All Events Enable')
            }

            case RegExMatch(action, '^All Events (Enable|Disable)$') || RegExMatch(action, '^Set Profile:'): 
                this.SetProfile(StrReplace(action, 'Set Profile: '))
        }
    }

    ;============================================================================================

    static Create_Array_aTrayIconLeftClick()
    {
        arr := ['- None -', 'Open Events Manager', 'Toggle All Events Enable/Disable', 'All Events Enable', 'All Events Disable']

        for profileName, mProfile in this.mProfiles
            arr.Push('Set Profile: ' profileName)        

        return arr
    }

    ;============================================================================================

    static Create_Array_Profiles()
    {
        arr := Array()
        
        for profileName, mProfile in this.mProfiles
            arr.Push(profileName)   

        return arr
    }

    ;=============================================================================================

    static Create_TrayMenuSelectProfile()
    {  
        for _, value in ['All Events Enable', 'All Events Disable']
            A_TrayMenu.SetIcon(value, this.mIcons['tMenuTrans'],, this.m['mUser']['trayMenuIconSize'])

        for profileName, mProfile in this.mProfiles {
            this.trayMenuSelectProfile.Add(profileName, this.TrayMenuProfile_ItemClick.Bind(this))
            this.trayMenuSelectProfile.SetIcon(profileName, this.mIcons['tMenuTrans'],, this.m['mUser']['trayMenuIconSize'])
        }

        if RegExMatch(this.profileMatch, '^All Events (Enable|Disable)$')
            A_TrayMenu.SetIcon(this.profileMatch, this.mThemeIcons[this.m['mUser']['themeName']]['select'],, this.m['mUser']['trayMenuIconSize'])
        else
            if this.mProfiles.Has(this.profileMatch)
                this.trayMenuSelectProfile.SetIcon(this.profileMatch, this.mThemeIcons[this.m['mUser']['themeName']]['select'],, this.m['mUser']['trayMenuIconSize'])
    }

    ;============================================================================================

    static Create_TrayMenuEvents()
    {
        this.trayMenuEvents.Delete()

        for eventName, mEvent in this.m['mEvents'] {
            switch {
                case mEvent.Has('window'): 
                {
                    itemName := (
                        (eventName ' <' this.FormatTxt(this.GetEventType(mEvent)) ' - ' this.FormatTxt(mEvent['monitoring']))
                      . (mEvent['window']['winActive'] || RegExMatch(mEvent['window']['winMinMax'], '^(0|1|-1)$') ? ' - ' : '')
                      . (mEvent['window']['winActive'] ? 'WinActive' : '')
                      . (mEvent['window']['winActive'] && RegExMatch(mEvent['window']['winMinMax'], '^(0|1|-1)$') ? ' - ' : '')
                      . (RegExMatch(mEvent['window']['winMinMax'], '^(0|1|-1)$') ? 'WinMinMax ' mEvent['window']['winMinMax'] '' : '')
                      . (!mEvent['window']['winActive'] && !RegExMatch(mEvent['window']['winMinMax'], '^(0|1|-1)$') ? ' - Mode ' mEvent['mode'] '>' : '>')
                    )
                }
                default: itemName := eventName ' <' this.FormatTxt(this.GetEventType(mEvent)) ' - ' this.FormatTxt(mEvent['monitoring']) ' - Mode ' mEvent['mode'] '>'
            }

            this.trayMenuEvents.Add(itemName, this.TrayMenuEvent_ItemClick.Bind(this, eventName))
            
            if mEvent['state'] = 1
                this.trayMenuEvents.SetIcon(itemName, this.mThemeIcons[this.m['mUser']['themeName']]['checkmark'],, this.m['mUser']['trayMenuIconSize'])
            else
                this.trayMenuEvents.SetIcon(itemName, this.mIcons['tMenuTrans'],, this.m['mUser']['trayMenuIconSize'])
        }
    }

    ;============================================================================================

    static TrayMenu_Enable_Disable(state)
    {
        for _, value in ['Settings', 'About', 'Tools', 'Edit Script', 'Open Script Folder', 'Select Profile', 'Events Manager', 'Events', 'All Events Enable', 'All Events Disable'] 
            A_TrayMenu.%state%(value)

        if state == 'enable'
            TraySetIcon(this.mThemeIcons[this.mInitValues['mUser']['themeName']]['trayMain'])
        else
            TraySetIcon(this.mThemeIcons[this.mInitValues['mUser']['themeName']]['trayLoading'])
    }

    ;============================================================================================

    static SetIconTip(*) => A_IconTip := this.scriptName ' ' this.scriptVersion '`nActive Profile:  ' (this.profileMatch ? this.profileMatch : '- None -')

    ;============================================================================================
    ; When given a single parameter, this function will copy all events from m1 to a new map m2 (Except the boundFuncTimer key). With two parameters, it will add event m1 to the existing map m2.
    static Copy_Add_mEvents(m1, m2:='')
    {
        if !IsObject(m2)
            m2 := Map()

        for eventName, mEvent in m1 {
            m2[eventName] := Map()
            for _, value in ['monitoring', 'mode', 'period', 'function', 'state', 'critical', 'log', 'notifInfo', 'notifStatus', 'soundCreated', 'soundTerminated']
                m2[eventName][value] := mEvent[value]
        }

        for eventName, mEvent in m1
            for _, value in ['process', 'window', 'device']
                if mEvent.Has(value)
                    m2[eventName][value] := mEvent[value]
        
        return m2
    }

    ;============================================================================================

    static Copy_mWinEventsStates(m)
    {
        mCopy := Map()

        for winEventName, value in m
            mCopy[winEventName] := value

        return mCopy
    }  

    ;============================================================================================

    static GUI_IfExistReturn_Destroy(aGUIs, param)
    {
        (this.HasVal('gMain', aGUIs))       ? WinExist(this.gMainTitle ' ahk_class AutoHotkeyGUI')       ? param == 'destroy' ? this.gMain.Destroy(): returnFlag := true : '' : ''
        (this.HasVal('gEvent', aGUIs))      ? WinExist(this.gEventTitle ' ahk_class AutoHotkeyGUI')      ? param == 'destroy' ? this.gEvent_Close(): returnFlag := true : '' : ''
        (this.HasVal('gSettings', aGUIs))   ? WinExist(this.gSettingsTitle ' ahk_class AutoHotkeyGUI')   ? param == 'destroy' ? this.gSettings_Close(): returnFlag := true : '' : ''
        (this.HasVal('gAbout', aGUIs))      ? WinExist(this.gAboutTitle ' ahk_class AutoHotkeyGUI')      ? param == 'destroy' ? this.gAbout_Close(): returnFlag := true : '' : ''
        (this.HasVal('gProfile', aGUIs))    ? WinExist(this.gProfileTitle  ' ahk_class AutoHotkeyGUI')   ? param == 'destroy' ? Winclose(this.gProfileTitle  ' ahk_class AutoHotkeyGUI'): returnFlag := true : '' : ''
        (this.HasVal('gDelConfirm', aGUIs)) ? WinExist(this.gDelConfirmTitle ' ahk_class AutoHotkeyGUI') ? param == 'destroy' ? Winclose(this.gDelConfirmTitle ' ahk_class AutoHotkeyGUI'): returnFlag := true : '' : ''
    
        if param == 'ifExistReturn' && IsSet(returnFlag)
            return true
    }

    ;============================================================================================

    static RunTool(filePath, *)
    {
        if FileExist(filePath)
            Run(filePath)
        else
            this.MsgBox('The system cannot find the file specified.`n"' filePath '"', this.gErrorTitle, this.mIcons['gX'], '+AlwaysOnTop -MinimizeBox -MaximizeBox -SysMenu',,,,,
            [{name:'*OK', callback:'this.MsgBox_Destroy'}])
    }

    ;============================================================================================

    static LogEventToFile(str) 
    {
        logFile := FileOpen('EventsLog.txt', 'a', 'UTF-8')
        logFile.Write('`n' FormatTime(A_Now, 'yyyy-MM-dd H:m:s') '`n' str '`n')
        logFile.Close()
    }

    ;============================================================================================
    ; get individual Event function line number
    static GetEventFunctionLineNumber(function)
    {
        for index, line in this.linesFuncFile
            if RegExMatch(line, '^(?!;\s*)(?i)' function '_(Created|Terminated)', &match) && IsSet(%match[0]%)
                return index
                
        return '-'
    }

    ;============================================================================================
    ; get all Events functions line number
    static GetEventsFunctionLineNumber(mEvents)
    {
        for index, line in this.linesFuncFile {
            if RegExMatch(line, '^(?!;\s*)(?i)(.+?)_(Created|Terminated)', &match) {
                for _, mEvent in mEvents {
                    if (mEvent['function'] = match[1] && IsSet(%match[0]%) && !mEvent.Has('lineNumber')) {                                       
                        mEvent['lineNumber'] := index
                        continue
                    }
                }
            }
        }

        for _, mEvent in mEvents  
            if !mEvent.Has('lineNumber') 
                mEvent['lineNumber'] := '-'
    }

    ;============================================================================================

    static IsValidFunctionName(str) => RegExMatch(str, 'i)^(?!\d)(\w|[^\x00-\x7f])+$') && StrLen(str) <= 50
    
    ;============================================================================================

    static IsValidEventName(str) => RegExMatch(str, 'i)^(\w|\s|[^\x00-\x7f])+$') && StrLen(str) <= 50

    ;============================================================================================

    static GetEventType(mEvent)
    {
        for _, value in ['window', 'process', 'device']
            if mEvent.Has(value)
                return value

        return 0
    }

    ;============================================================================================

    static FormatTxt(str)
    {
        static m := Map(
            'timer', 'Timer',
            'wmi', 'WMI',
            'winEvent', 'WinEvent',
            'deviceChange', 'DeviceChange',
            'window', 'Window',
            'process', 'Process',
            'device', 'Device'
        )

        if m.Has(str)
            return m[str]
    
        for key, value in m
            if value == str
                return key
    }    

    ;============================================================================================

    static FormatTxtMode(mEvent)
    {
        m := Map('mode', '')

        if (mEvent.Has('window'))
            m['mode'] := ((mEvent['window']['winActive'] ? 'WinActive - ' : '') (RegExMatch(mEvent['window']['winMinMax'], '^(0|1|-1)$') ? 'WinMinMax ' mEvent['window']['winMinMax'] : ''))

        m['mode'] := Trim(m['mode'], ' - ')

        if !m['mode']
            m['mode'] := mEvent['mode']

        return m['mode'] 
    }

    ;============================================================================================

    static FormatTxtOptions(mEvent)
    {
        m := Map('options', '')

        m['options'] := (
            (mEvent['critical'] ? 'Critical - ' : '') (mEvent['notifStatus'] ? 'Status - ' : '') (mEvent['notifInfo'] ? 'Info - ' : '') (mEvent['log'] ? 'Log - ' : '')
         .  ((mEvent['soundCreated'] !== '- None -' || mEvent['soundTerminated'] !== '- None -') ? 'Sound': '')
        )
        
        m['options'] := Trim(m['options'], ' - ')
   
        if !m['options']
            m['options'] := '-'

        return m['options']        
    }    

    ;============================================================================================

    static DDLchoose(choice, array, ctrlObj, caseSensitive := true)
    {
        for index, value in array {
            if (caseSensitive && value == choice) || (!caseSensitive && value = choice){
                ctrlObj.Choose(index)
                break
            }
        }
    }
    
    ;============================================================================================

    static DDLArrayChange_Choose(choice, array, ctrlObj)
    {
        ctrlObj.Delete()
        ctrlObj.Add(array)
        this.DDLchoose(choice, array, ctrlObj)
    }

    ;============================================================================================

    static CreateIncrementalArray(interval, min, max)
    {
        arr := Array()
        Loop ((max - min) // interval) + 1 ; loop until the maximum is reached.
            arr.Push(min + (A_Index - 1) * interval)
        return arr
    }

    ;============================================================================================

    static FindClosestNumberInArray(num, arr) 
    {
        if !IsInteger(num)
            return 0

        closest := arr[1]
        closestDiff := Abs(num - arr[1])
    
        for index, value in arr {
            diff := Abs(num - value)
            if (diff < closestDiff) {
                closest := value
                closestDiff := diff
            }
        }
        
        return Format("{:.0f}", closest)
    }

    ;=============================================================================================

    static HasVal(needle, haystack, caseSensitive := true)
    {
        for index, value in haystack
            if (caseSensitive && value == needle) || (!caseSensitive && value = needle)
                return index

        return 0
    }
    
    ;============================================================================================

    static IsArrayContainsSameValues(a1, a2) 
    {
        if a1.Length != a2.Length
            return 0

        cntValueMatch := 0
        
        for _, value in a1
            if this.HasVal(value, a2)
                cntValueMatch++

        if a1.Length = cntValueMatch
            return 1

        return 0
    }

    ;============================================================================================

    static GUI_ShowOnSameDisplay(gObj, gParentid := '', posXY := '', posWH := '')
    {
        if (gParentid) {
            WinGetClientPos(&posgParentX, &posgParentY,,, 'ahk_id ' gParentid)
            indexMon := this.MonitorGetMouseIsIn()
        }
        
        if IsSet(indexMon) && indexMon != 1
            gObj.Show((posXY ? posXY : 'x' posgParentX ' y' posgParentY) (posWH ? ' ' posWH : ''))             
        else
            gObj.Show(posXY (posWH ? ' ' posWH : ''))
    }    

    ;============================================================================================

    static MonitorGetMouseIsIn()
    {
        coordmmPrev := A_CoordModeMouse
        CoordMode('Mouse', 'Screen'), MouseGetPos(&mX, &mY)
        CoordMode('Mouse', coordmmPrev)
        
        Loop MonitorGetCount() {
            MonitorGet(A_Index, &left, &top, &right, &bottom)
            if (mX >= left) && (mX < right) && (mY >= top) && (mY < bottom)
                return A_Index
        }
        
        return 0
    }

    ;=============================================================================================
    ; MoveControls by Descolada. (from UIATreeInspector.ahk)    https://github.com/Descolada/UIA-v2
    static MoveControls(ctrls*)
    {
        for ctrl in ctrls
            ctrl.Control.Move(ctrl.HasOwnProp('x') ? ctrl.x : unset, ctrl.HasOwnProp('y') ? ctrl.y : unset, ctrl.HasOwnProp('w') ? ctrl.w : unset, ctrl.HasOwnProp('h') ? ctrl.h : unset)
    }

    ;=============================================================================================
    ; Based on DisplayObj by FanaticGuru (inspired by tidbit and Lexikos v1 code)   https://www.autohotkey.com/boards/viewtopic.php?p=507896#p507896
    static DisplayObj(obj, depth:=5, indentLevel:='')
    {
        if !RegExMatch(Type(obj), '^(Array|Map|Object)$') 
        || (Type(obj) == 'Object' && !ObjOwnPropCount(obj)) 
        || (Type(obj) == 'Array' && !obj.Length)
        || (Type(obj) == 'Map' && !obj.Count)
            return

        if Type(obj) == 'Object'
            obj := obj.OwnProps()

        for k,v in obj
        {
            list .= indentLevel '[' k ']'
            if (IsObject(v) && depth>1)
                list .= '`n' this.DisplayObj(v, depth-1, indentLevel . '    ')
            else
                list .= ' => ' v
            list .= '`n'
        }
        return RTrim(list)
    }

    ;============================================================================================

    static MsgBox(
        text       := '', 
        title      := A_ScriptName, 
        icon       := '',
        options    := '-MinimizeBox -MaximizeBox -SysMenu',
        owner      := '',
        winSetAoT  := 0,
        posXY      := '',
        sound      := '',
        aObjBtn    := [{name:'*OK', callback:'this.MsgBox_Destroy'}], 
        iconSize   := 32,
        font       := 'Segoe UI',
        fontSize   := 10,
        btnWidth   := 100,
        btnHeight  := 30) {

        if !text || !aObjBtn.Length
            return            
 
        wasNotCritical := this.SetCritical_On(A_IsCritical)
       
        static gIndex:=0
        gIndex++

        g := Gui(options, title)
        g.SetFont('s' fontSize)
        g.OnEvent('Close', this.MsgBox_Destroy.Bind(this, gIndex, owner))
        ogMarginY := g.MarginY
        g.MarginY := 15
        
        if (owner) {
            owner.Opt('+Disabled')
            g.Opt('+Owner' owner.hwnd)
        }
        
        switch {
            case FileExist(icon) || RegExMatch(icon, 'i)h(icon|bitmap).*\d+', &match):  
                try g.Add('Picture', 'w' iconSize ' h' iconSize, icon)
            
            case RegExMatch(icon, '^(icon!|icon\?|iconx|iconi)$'):           
                try g.Add('Picture', 'w' iconSize ' h' iconSize ' Icon' this.mIconsUser32[icon], A_WinDir '\system32\user32.dll')

            case RegExMatch(icon, 'i)(.+?\.(?:dll|exe))\|icon(\d+)', &match) && FileExist(match[1]):
                try g.Add('Picture', 'w' iconSize ' h' iconSize ' Icon' match[2], match[1])
        }

        ;============================================== 
        For hwnd, ctrlObj in g
            if ctrlObj.Type == 'Pic'
                gHasPic := true

        if text && this.ControlGetTextWidth(text, font, fontSize) + g.MarginX*2 >= (A_ScreenWidth/16 * 3)     
            textWidth := A_ScreenWidth/16 * 3

        g.txt := g.Add('Text', (IsSet(gHasPic) ? 'x+m' : '') (IsSet(textWidth) ? ' w' textWidth : ''), text)

        ;==============================================
        g.gIndex := gIndex
        g.owner := owner

        for _, obj in aObjBtn 
        {
            if A_Index > 1
                g.MarginY := ogMarginY

            g.btn%A_Index% := g.Add('Button', (A_Index=1 ? 'yp' : 'xp') . (InStr(obj.Name, '*') ? " Default " : ' ') . 'w' btnWidth ' h' btnHeight, RegExReplace(obj.Name, '\*'))
            
            if Regexmatch(obj.callback, '^this.(.*)$', &match)
                g.btn%A_Index%.OnEvent('Click', this.%match[1]%.Bind(this, g))
            else
                g.btn%A_Index%.OnEvent('Click', %obj.callback%.Bind(g))
        }

        ; Hidden button to avoid short gui.
        if aObjBtn.Length = 1
            g.Add('Button', 'w' btnWidth ' h' btnHeight ' +Hidden')

        this.mMsgBoxGUIs[gIndex] := g
        g.MarginX := 15
        g.MarginY := 20

        if sound
            this.Sound(sound)

        this.GUI_ShowOnSameDisplay(g, (owner ? WinExist('ahk_id ' owner.hwnd) : ''), posXY)

        if winSetAoT
            WinSetAlwaysOnTop(1, title ' ahk_class AutoHotkeyGUI')

        this.SetCritical_Off(wasNotCritical)
        return g.hwnd
    }

    ;============================================================================================

    static MsgBox_Destroy(g, *)
    { 
        wasNotCritical := this.SetCritical_On(A_IsCritical)
        
        if g.HasOwnProp('owner')
            try g.owner.Opt('-Disabled')
        
        this.mMsgBoxGUIs[g.gIndex].Destroy()
        this.mMsgBoxGUIs.Delete(g.gIndex)
        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================
    ; Notify by XMCQCX (Inspired by NotifyV2 from the-Automator.com)  https://www.the-automator.com/downloads/maestrith-notify-class-v2/
    static Notify(
        hdTxt       := '', 
        bdTxt       := '',
        icon        := '',
        options     := '+Owner -Caption +AlwaysOnTop',
        position    := 'bottomRight',
        duration    := 8000,
        callback    := '',
        sound       := '',
        iconSize    := 32,
        hdFontSize  := 15,
        hdFontColor := 'white',
        hdFont      := 'Segoe UI bold', 
        bdFontSize  := 12, 
        bdFontColor := 'white', 
        bdFont      := 'Segoe UI', 
        bgColor     := '1F1F1F',
        style       := 'round') {  

        if !hdTxt && !bdTxt && !icon
            return              
        
        wasNotCritical := this.SetCritical_On(A_IsCritical)        

        static gIndex:=0, gLastPosYbr:=0, gLastPosYbc:=0, gLastPosYbL:=0, gLastPosYtL:=10, gLastPosYtc:=10, gLastPosYtr:=10
        static paddingX := 15, paddingY := 10
        gIndex++

        g := Gui(options) 
        g.BackColor := bgColor
        g.MarginX := 15
        g.MarginY := 15
               
        switch {
            case FileExist(icon) || RegExMatch(icon, 'i)h(icon|bitmap).*\d+', &match):  
                try g.Add('Picture', 'w' iconSize ' h' iconSize, icon)
            
            case RegExMatch(icon, '^(icon!|icon\?|iconx|iconi)$'):           
                try g.Add('Picture', 'w' iconSize ' h' iconSize ' Icon' this.mIconsUser32[icon], A_WinDir '\system32\user32.dll')

            case RegExMatch(icon, 'i)(.+?\.(?:dll|exe))\|icon(\d+)', &match) && FileExist(match[1]):
                try g.Add('Picture', 'w' iconSize ' h' iconSize ' Icon' match[2], match[1])
        }

        ;============================================== 
        For hwnd, ctrlObj in g
            if ctrlObj.Type == 'Pic'
                gHasPic := true

        if hdTxt && (this.ControlGetTextWidth(hdTxt, hdFont, hdFontSize) + (IsSet(gHasPic) ? iconSize + 10 : 0) + g.MarginX*2) >= (A_ScreenWidth/16 * 13)            
            hdtxtWidth := A_ScreenWidth/100 * 60

        if bdTxt && (this.ControlGetTextWidth(bdTxt, bdFont, bdFontSize) + (IsSet(gHasPic) ? iconSize + 10 : 0) + g.MarginX*2) >= (A_ScreenWidth/16 * 13) 
            bdtxtWidth := A_ScreenWidth/100 * 60        

        ;==============================================       
        if (hdTxt) {
            g.SetFont('s' hdFontSize ' c' hdFontColor, hdFont)
            g.Add('Text', (IsSet(gHasPic) ? 'x+10' : '') (IsSet(hdtxtWidth) ? ' w' hdtxtWidth : ''), hdTxt)
        }

        if (bdTxt) {
            if hdTxt
                g.MarginY := 6            

            g.SetFont('s' bdFontSize ' c' bdFontColor, bdFont)
            g.Add('Text', (!hdTxt && IsSet(gHasPic) ? 'x+10' : '') (IsSet(bdtxtWidth) ? ' w' bdtxtWidth : ''), bdTxt)
        }

        g.MarginY := 15
        g.Show('Hide')
        WinGetPos(,, &gW, &gH, g)
        clickArea := g.Add('Text', 'x0 y0 w' gW ' h' gH ' BackgroundTrans')
        
        if callback
            clickArea.OnEvent('Click', callback), duration := 0
            
        clickArea.OnEvent('Click', this.gNotify_Destroy.Bind(this, g, 'click'))                   
        
        g.gIndex := gIndex
        g.position := position
        g.style := style        
        g.boundFuncTimer := this.gNotify_Destroy.Bind(this, g)
        SetTimer(g.boundFuncTimer, -duration)
        this.mNotifyGUIs[position][gIndex] := g
        
        Switch position {
            case 'bottomRight': 
            {
                if this.mNotifyGUIs[position].Count <= 1
                    gLastPosYbr := 0
            
                if gLastPosYbr < gH + paddingY
                    gPos := 'x' A_ScreenWidth - gW - paddingX 'y' (gLastPosYbr := A_ScreenHeight - gH - 75)
                else
                    gPos := 'x' A_ScreenWidth - gW - paddingX ' y' (gLastPosYbr := gLastPosYbr - gH - paddingY)
            }

            case 'bottomCenter':
            {
                if this.mNotifyGUIs[position].Count <= 1
                    gLastPosYbc := 0

                if gLastPosYbc < gH + paddingY 
                    gPos := 'x' A_ScreenWidth/2 - gW/2 'y' (gLastPosYbc := A_ScreenHeight - gH - 75)            
                else
                    gPos := 'x' A_ScreenWidth/2 - gW/2 'y' (gLastPosYbc := gLastPosYbc - gH - paddingY)
            }

            case 'bottomLeft': 
            {
                if this.mNotifyGUIs[position].Count <= 1
                    gLastPosYbL := 0

                if gLastPosYbL < gH + paddingY 
                    gPos := 'x' paddingX ' y' (gLastPosYbL := A_ScreenHeight - gH - 75)
                else 
                    gPos := 'x' paddingX ' y' (gLastPosYbL := gLastPosYbL - gH - paddingY)
            }             

            case 'topLeft': 
            {
                if this.mNotifyGUIs[position].Count <= 1
                    gLastPosYtL := paddingY

                if (gLastPosYtL = paddingY || (gLastPosYtL + gH + paddingY) > A_ScreenHeight)
                    gPos := 'x' paddingX ' y' paddingY, gLastPosYtL := paddingY + gH
                else
                    gPos := 'x' paddingX ' y' gLastPosYtL + paddingY, (gLastPosYtL := gLastPosYtL + paddingY + gH)
            }     
            
            case 'topCenter': 
            {
                if this.mNotifyGUIs[position].Count <= 1
                    gLastPosYtc := paddingY

                if (gLastPosYtc = paddingY || (gLastPosYtc + gH + paddingY) > A_ScreenHeight)
                    gPos := 'x' A_ScreenWidth/2 - gW/2 ' y' paddingY, gLastPosYtc := paddingY + gH
                else
                    gPos := 'x' A_ScreenWidth/2 - gW/2 ' y' gLastPosYtc + paddingY, (gLastPosYtc := gLastPosYtc + paddingY + gH)
            }  
            
            case 'topRight': 
            {
                if this.mNotifyGUIs[position].Count <= 1
                    gLastPosYtr := paddingY

                if (gLastPosYtr = paddingY || (gLastPosYtr + gH + paddingY) > A_ScreenHeight)
                    gPos := 'x' A_ScreenWidth - gW - paddingX ' y' paddingY, gLastPosYtr := paddingY + gH
                else
                    gPos := 'x' A_ScreenWidth - gW - paddingX ' y' gLastPosYtr + paddingY, (gLastPosYtr := gLastPosYtr + paddingY + gH)
            }             
        }

        if sound
            this.Sound(sound)

        Switch style {
            case 'round':
            {
                g.Show(gPos ' NoActivate Hide')
                this.FrameShadow(g.hwnd)
                DllCall('user32\AnimateWindow', 'UInt', this.mNotifyGUIs[position][gIndex].hwnd, 'Int', 25, 'UInt', AW_BLEND := 0x00080000)
            }
            
            default:
            {
                g.Opt('+Border')
                g.Show(gPos ' NoActivate')
            }
        }

        this.SetCritical_Off(wasNotCritical)
        return g.hwnd
    }

    ;============================================================================================

    static gNotify_Destroy(g, onEvent:='', *)
    {
        wasNotCritical := this.SetCritical_On(A_IsCritical)

        if onEvent == 'click'
            SetTimer(this.mNotifyGUIs[g.position][g.gIndex].boundFuncTimer, 0)

        Switch g.style {
            case 'round': DllCall('AnimateWindow', 'ptr', this.mNotifyGUIs[g.position][g.gIndex].hwnd, 'uint', 25, 'UInt', AW_HIDE := 0x00010000)
            default: 
            {
                Switch g.position {
                    case 'bottomRight', 'topRight': animate := '0x00050001' ; AW_HOR_POSITIVE
                    case 'bottomLeft', 'topLeft':   animate := '0x00050002' ; AW_HOR_NEGATIVE
                    case 'bottomCenter': animate := '0x00050004' ; AW_VER_POSITIVE
                    case 'topCenter':    animate := '0x00050008' ; AW_VER_NEGATIVE
                }
                DllCall('AnimateWindow', 'UInt', this.mNotifyGUIs[g.position][g.gIndex].hwnd, 'Int', 25, 'UInt', animate)
            }  
        }   
                
        this.mNotifyGUIs[g.position][g.gIndex].Destroy()
        this.mNotifyGUIs[g.position].Delete(g.gIndex)
        this.SetCritical_Off(wasNotCritical)
    }

    ;============================================================================================
    ; FrameShadow by Klark92  https://www.autohotkey.com/boards/viewtopic.php?f=6&t=29117&hilit=FrameShadow
    static FrameShadow(hwnd)
    {
        DllCall("dwmapi.dll\DwmIsCompositionEnabled", "int*", &dwmEnabled:=0)
        
        if !dwmEnabled {
            DllCall("user32.dll\SetClassLongPtr", "ptr", hwnd, "int", -26, "ptr", DllCall("user32.dll\GetClassLongPtr", "ptr", hwnd, "int", -26) | 0x20000)
        }
        else {
            margins := Buffer(16, 0)    
            NumPut("int", 1, "int", 1, "int", 1, "int", 1, margins)
            DllCall("dwmapi.dll\DwmSetWindowAttribute", "ptr", hwnd, "Int", 2, "Int*", 2, "Int", 4)
            DllCall("dwmapi.dll\DwmExtendFrameIntoClientArea", "ptr", hwnd, "ptr", margins)
        }
    }

    ;============================================================================================

    static ControlGetTextWidth(str:='', font:='', fontSize:='')
    {
        g := Gui()
        g.SetFont('s' fontSize, font)
        g.txt := g.Add('Text',, str)                   
        g.txt.GetPos(,, &ctrlW)        
        return ctrlW
    }   

    ;============================================================================================

    static Sound(sound)
    {
        if RegExMatch(sound, '^(soundx|soundi)$')
            sound := this.mSounds[sound]

        if FileExist(sound) || RegExMatch(sound,'^\*\-?\d+')
            Soundplay(sound)
    }

    ;============================================================================================
    
    static SetCritical_On(IsCritical)
    {
        if (!IsCritical) {
            Critical
            this.isCriticalMethodExecuting := true
            return wasNotCritical := true
        }
    }

    ;============================================================================================
    
    static SetCritical_Off(wasNotCritical)
    {
		if (wasNotCritical) {
            Critical('Off')
            this.isCriticalMethodExecuting := false
        }
    }    
}
