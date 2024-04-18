/*
    Script:    WinExeCommander.ahk
    Author:    Martin Chartier (XMCQCX)
    Github:    https://github.com/XMCQCX/WinExeCommander
    AHKForum:  https://www.autohotkey.com/boards/viewtopic.php?f=83&t=128956
*/

#Requires AutoHotkey v2.0
#SingleInstance

#Include '.\Lib\WinExeCmd.ahk'

/* Examples =================================================================

When the Calculator app is opened, set the window to be always on top.

Calculator_AlwaysOnTop_Created(mEvent) {
    
    if WinExist('ahk_id ' mEvent['id'])
        WinSetAlwaysOnTop(1, 'ahk_id ' mEvent['id'])
    else
        WinExeCmd.MsgBox('Calculator does not exist.', 'WinExeCommander', 'iconx')    
}

;===================================

When the "mspaint.exe" process is created, change its priority level to "High".

MSPaint_ProcessSetPriority_Created(mEvent) {

    if ProcessExist(mEvent['pid'])
        ProcessSetPriority('High', mEvent['pid'])
    else
        WinExeCmd.MsgBox('MSPaint does not exist.', 'WinExeCommander', 'iconx')
}
*/ 
;============================================================================
