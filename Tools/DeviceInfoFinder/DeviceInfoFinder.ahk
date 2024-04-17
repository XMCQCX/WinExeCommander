/*
    Script:    DeviceInfoFinder.ahk
    Author:    Martin Chartier (XMCQCX)
    Date:      2024-04-17
    Version:   1.0.0
    Tested on: Windows 11
    Github:    https://github.com/XMCQCX/WinExeCommander
    AHKForum:
*/

#Requires AutoHotkey v2.0
#SingleInstance

#Include '..\..\Lib\LV_GridColor.ahk'


class DeviceInfoFinder {

    static __New() 
    {
        this.scriptName    := 'DeviceInfoFinder'
        this.scriptVersion := 'v1.0.0'
        try TraySetIcon('HICON:*' LoadPicture(A_ScriptDir '\DeviceInfoFinder.ico', 'w32', &ImageType))
        ;==============================================
        this.gMain := Gui('+Resize +MinSize500x300',  this.scriptName ' ' this.scriptVersion)
        this.gMain.SetFont('s10')
        this.gMain.OnEvent('Size', this.gMain_Size.bind(this))        
        
        this.gMain.lvWidth := 1000
        this.gMain.lvHeight := 600
        this.gMain.btnWidth := 125
        this.gMain.btnHeight := 40
        this.gMain.btnsWidth := this.gMain.btnWidth*2 + this.gMain.MarginX
        this.gMain.btnsPos := this.gMain.lvWidth/2 - this.gMain.btnsWidth/2 + this.gMain.MarginX

        this.gMain.lv := this.gMain.AddListView('w' this.gMain.lvWidth ' h' this.gMain.lvHeight ' r10 Grid +BackgroundEAF4FB', ['Device Name', 'Device ID'])
        this.gMain.lv.OnEvent('ContextMenu', this.gMain_ContextMenu.bind(this))
        this.gMain.btn_findDeviceInfo := this.gMain.Add('Button', 'x' this.gMain.btnsPos ' w' this.gMain.btnWidth ' h' this.gMain.btnHeight, 'Find Device Info')
        this.gMain.btn_findDeviceInfo.OnEvent('Click', this.gMain_btn_findDeviceInfo_Click.Bind(this))
        this.gMain.btn_close:= this.gMain.Add('Button', 'x+' this.gMain.MarginX ' w' this.gMain.btnWidth ' h' this.gMain.btnHeight, 'Close')
        this.gMain.btn_close.OnEvent('Click', (*) => ExitApp())
        this.gMain.OnEvent('Close', (*) => ExitApp())        
        this.gMain.lv.ModifyCol(1, '230')
        this.gMain.lv.ModifyCol(2, '465')
        LV_GridColor(this.gMain.lv, '0xcbcbcb')
        this.gMain.Show()
    }

    ;=============================================================================================
    
    static gMain_btn_findDeviceInfo_Click(*)
    {
        this.gMain.Opt('+OwnDialogs')
        
        if MsgBox('If your device is already connected, click OK. If it`'s not, connect it and then click OK.', 'Find Device Info', 'OKCancel Iconi') == 'Cancel'
            return

        oDevices := this.DeviceFinder()

        if MsgBox('Disconnect your device and click OK.', 'Find Device Info', 'OKCancel Iconi') == 'Cancel'
            return

        oDevicesOut := this.DeviceFinder()

        oDevicesDiff := {}
        for key, value in  oDevices.OwnProps()
            if !oDevicesOut.HasOwnProp(key)
                oDevicesDiff.%key% := value
    
        if (ObjOwnPropCount(oDevicesDiff)) {
            imageList := IL_Create(10)
            this.gMain.lv.SetImageList(imageList)

            for id, value in oDevicesDiff.OwnProps() {
                IL_Add(imageList, 'HICON:' . value.hIcon)
                this.gMain.lv.Insert(1,'Icon' . A_Index, value.FriendlyName, id)
            }
            
            this.gMain.lv.ModifyCol(1,'AutoHdr')
            this.gMain.lv.ModifyCol(2,'AutoHdr')
        }
    }

    ;============================================================================================

    static gMain_ContextMenu(ctrlObj, item, isRightClick, x, y)
    {
        MouseGetPos(,,, &mouseOverClassNN)
        if item = 0 || InStr(mouseOverClassNN, 'SysHeader')
            return

        ctxMenu := Menu()
        ctxMenu.Add('Copy Selected Item', CopyToClipboard)
        ctxMenu.Add('Clear Listview', ClearListview)
        ctxMenu.Show(x, y)

        CopyToClipboard(*)
        {
            formattedTxt := ''
            
            for index, line in StrSplit( ListViewGetContent('Selected', ctrlObj.hwnd, this.scriptName ' ahk_class AutoHotkeyGUI') , '`n')
                formattedTxt .= StrSplit(line, '`t')[1] ':`n' StrSplit(line, '`t')[2] '`n`n'
                
            A_Clipboard := '', A_Clipboard := Trim(formattedTxt, '`r`n`t ')
            if !ClipWait(1)
                return
        }

        ClearListview(*) => ctrlObj.Delete()
    }

    ;=============================================================================================

    static gMain_Size(guiObj, minMax, width, height)
    {
        if minMax = -1
            return

        MoveControls(
            {Control:this.gMain.lv, w:width - this.gMain.MarginX*2, h:height - this.gMain.btnHeight - this.gMain.MarginY - this.gMain.MarginY*3},
            {Control:this.gMain.btn_findDeviceInfo, x:width/2 - this.gMain.btnsWidth/2, y:height - this.gMain.btnHeight - this.gMain.MarginY*2},
            {Control:this.gMain.btn_close, x:width/2 - this.gMain.btnsWidth/2 + this.gMain.btnWidth + this.gMain.MarginX, y:height - this.gMain.btnHeight - this.gMain.MarginY*2}
        )
        
        DllCall('RedrawWindow', 'ptr', this.gMain.hwnd, 'ptr', 0, 'ptr', 0, 'uint', 0x0081)

        ; MoveControls by Descolada. (from UIATreeInspector.ahk)  https://github.com/Descolada/UIA-v2
        MoveControls(ctrls*) {
            for ctrl in ctrls
                ctrl.Control.Move(ctrl.HasOwnProp('x') ? ctrl.x : unset, ctrl.HasOwnProp('y') ? ctrl.y : unset, ctrl.HasOwnProp('w') ? ctrl.w : unset, ctrl.HasOwnProp('h') ? ctrl.h : unset)
        }
    }

    ;============================================================================================
    ; Based on EnumDeviceInfo by teadrinker and JEE_DeviceList by jeeswg.
    ; https://www.autohotkey.com/boards/viewtopic.php?t=69380  https://www.autohotkey.com/boards/viewtopic.php?f=82&t=120154  https://www.autohotkey.com/boards/viewtopic.php?t=121125&p=537515
    static DeviceFinder()
    {
        Critical
        static flags := (DIGCF_PRESENT := 0x2) | (DIGCF_ALLCLASSES := 0x4)
             , PKEY_Device_DeviceDesc   := '{A45C254E-DF1C-4EFD-8020-67D146A850E0} 2'
             , PKEY_Device_FriendlyName := '{A45C254E-DF1C-4EFD-8020-67D146A850E0} 14'
             , PKEY_Device_InstanceId   := '{78C34FC8-104A-4ACA-9EA4-524D52996E57} 256'
             , cxSmallIcon := SysGet(SM_CXSMICON := 49), cySmallIcon := SysGet(SM_CYSMICON := 50)
    
        if !(hModule := DllCall('kernel32\LoadLibrary', 'Str','setupapi.dll', 'Ptr'))
        || !(hDevInfo := DllCall('SetupAPI\SetupDiGetClassDevs', 'Ptr', 0, 'Ptr', 0, 'Ptr', 0, 'UInt', flags))
            return   
        
        SP_DEVINFO_DATA := Buffer(size := 24 + A_PtrSize, 0)
        NumPut('UInt', size, SP_DEVINFO_DATA)
        obj := {}
    
        friendlyName := deviceDesc := instanceId := ''
        for prop in ['FriendlyName', 'DeviceDesc', 'InstanceId'] {
            DllCall('Propsys\PSPropertyKeyFromString', 'Str', PKEY_Device_%prop%, 'Ptr', %prop% := Buffer(20, 0))
        }

        while DllCall('SetupAPI\SetupDiEnumDeviceInfo', 'Ptr', hDevInfo, 'UInt', A_Index - 1, 'Ptr', SP_DEVINFO_DATA) {
            GetDevicePropFn := GetDeviceProp.Bind(hDevInfo, SP_DEVINFO_DATA)
            obj.%GetDevicePropFn(instanceId)% := {friendlyName: GetDevicePropFn(friendlyName) || GetDevicePropFn(deviceDesc), hIcon: GetDeviceIcon(hDevInfo, SP_DEVINFO_DATA, cxSmallIcon, cySmallIcon)}
        }

        DllCall('SetupAPI\SetupDiDestroyDeviceInfoList', 'Ptr', hDevInfo)
        DllCall('kernel32\FreeLibrary', 'Ptr',hModule)
        return obj
    
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
}