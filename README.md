# WinExeCommander	
WinExeCommander is an AutoHotkey script to simplify the calling of functions when windows/processes are created/terminated, devices are connected/disconnected.

![](https://i.imgur.com/fnRhH1h.png)

![](https://i.imgur.com/l2ZXYuj.png)

- [Requirement](#requirement)
- [Features](#features)
- [Supported devices](#supported-devices)
- [How to use it?](#how-to-use-it)
- [Event Parameters](#event-parameters)
	- [Window Parameters](#window-parameters)
  - [Process Parameters](#process-parameters)
  - [Device Parameters](#device-parameters)
- [Methods](#methods)
	- [SetProfile](#setprofile)
	- [SetEvent](#setevent)
  	- [SetPeriodWMI](#setperiodwmi)
	- [ProcessFinder](#processfinder)
	- [WindowFinder](#windowfinder)
	- [DeviceFinder](#devicefinder)
	- [Notify](#notify)
	- [MsgBox](#msgbox)
	- [Sound](#sound)
- [Donation](#donation)
- [Credits](#credits)

## Requirement
* AutoHotkey v2

## Features
* Execute fonctions upon:
  - Process creation/termination.
  - Window creation/termination, activated/deactivated.
  - Device connection/disconnection.

* Select from various criteria, including wintitle, winclass, process name, process path, active/maximize/hidden window and additional parameters.
* Enable or disable the monitoring of individual events using the tray menu, user interface (GUI), or method call.
* Save profiles and load them via the tray menu, user interface (GUI), or method call.
* Themes customization.

## Supported devices
USB, Bluetooth, HDMI etc...

## How to use it?

* Add an Event
  - Double-click the tray icon or right-click on it and choose "Event Manager".
  - Click the "Add Event" button.
  - Click on Window, Process, or Device button.
  - To fill the edit fields, you can manually enter data, or double-click on a listview item, or right-click on an item and select "Copy Row Data to Edit Fields".

* Edit an Event
  - Select an event and click the "Edit" button, double-click on it in the listview, or right-click and choose "Edit Event".
  
* Write a function associated with the event.
  - In the file "WinExeCommander.ahk", create a function to be called when the event is created or terminated. Append "_Created" or "_Terminated" to the event function name.

For example:
 
When the Calculator app is opened, set the window to be always on top.

Function Name: Calculator_AlwaysOnTop

	Calculator_AlwaysOnTop_Created(mEvent) {
	    
	    if WinExist('ahk_id ' mEvent['id'])
	        WinSetAlwaysOnTop(1, 'ahk_id ' mEvent['id'])
	    else
	        WinExeCmd.MsgBox('Calculator does not exist.', 'WinExeCommander', 'iconx')    
	}

When the "mspaint.exe" process is created, change its priority level to "High".

Function Name: MSPaint_ProcessSetPriority
	
	MSPaint_ProcessSetPriority_Created(mEvent) {
	
	    if ProcessExist(mEvent['pid'])
	        ProcessSetPriority('High', mEvent['pid'])
	    else
	        WinExeCmd.MsgBox('MSPaint does not exist.', 'WinExeCommander', 'iconx')
	}

* To identify a device
  - Open the "Event Manager"
  - Click "Add Event"
  - In the device section, check if the device is listed.
  - Alternatively, run "DeviceInfoFinder.ahk", found in the "Tools" tray menu, menubar, and the device section.
  
* Loading Profiles
  - The profile includes all events values, WMI period interval and window events (WinEvents) states. In the Profile Manager, you can select whether to load all event values or only the events states when loading profiles.

* Applying Changes
  - To apply modifications, make sure to click the "Apply" button after creating or modifying an event, changing the WMI period, loading a profile etc...

* Themes Creation
  - Create a "Themes" folder in the root directory.
  - Within that folder, create another folder and place 13 icons named: "about", "checkmark", "edit", "events", "exit", "folder", "loading", "main", "profile", "reload", "select", "settings", "tools".
  - To apply, select it from the dropdown menu in the settings. 

* Adding notification sounds
  - Create a "Sounds" folder in the root directory and place WAV files within that folder.
 
* Start with Windows
  - To automatically run this script on startup, add its shortcut to the Startup folder.  
 
## Event Parameters
  > - **Event Name**
  >
  > - **Function**
  >    - The name of the function to call upon event creation or termination. To call the function, append "_Created" or "_Terminated" to the function name.
  >
  > - **Critical**
  >    - Set the function's thread to critical.
  >
  > - **Status**
  >    - Display GUI with the event status in the bottom right corner.
  >  
  > - **Info**
  >    - Display GUI with the event information in the top left corner.
  >  
  > - **Log Info**
  >    - Log event information to "EventsLog.txt"
  >   
  > - **Sound**
  >    - Notification Sounds.


### Window Parameters
  > - **WinTitle**:
  >    - Case-sensitive, except when using the i) modifier in a RegEx pattern.
  >
  > - **WinTitleMatchMode**   
  >    - **1:** A window's title must start with the specified WinTitle to be a match.
  >    - **2:** A window's title can contain WinTitle anywhere inside it to be a match. (Default)
  >    - **3:** A window's title must exactly match WinTitle.
  >    - **RegEx:** Regular expression WinTitle matching.
  >	
  > - **DetectHiddenWindows** 
  >    - **Uncheck:** Hidden windows are not detected. (Default)
  >    - **Check:** Hidden windows are detected.
  > 
  > - **WinActive**   
  >    - **Uncheck:** Not monitoring for active window. (Default)
  >    - **Check:** Monitoring for active window.
  > 
  > - **WinMinMax** 
  >    - **Null:** Not monitoring WinMinMax.
  >    - **0:** The window is neither minimized nor maximized.
  >    - **1:** The window is maximized.
  >    - **-1:** The window is minimized.
  >    - **Limitation:** Only one window per event can be monitored.
  >   
  > - **Monitoring** 
  >    - **WinEvent**
  >        - SetWinEventHook. Sets an event hook function for a range of Windows events. (Default)
  >    - **Timer**
  >        - Check for the existence of the window at a specified time interval.
  >
  > - **Mode** 			
  >    - **1:** Call "Function_Created" for every window ID created. Call "Function_Terminated" for every window ID terminated.
  >    - **2:** Call "Function_Created" only for the initial window ID created. Call "Function_Terminated" only when the last window ID is terminated. (Default)	
  >    - **3:** Call "Function_Created" for every window ID created. Call "Function_Terminated" only when the last window ID is terminated.		
  >    - **4:** Call "Function_Created" for every window created. Call "Function_Terminated" for every window terminated.
  >    - **5:** Call "Function_Created" only for the initial window created. Call "Function_Terminated" only when the last window is terminated.
  >    - **6:** Call "Function_Created" for every window created. Call "Function_Terminated" only when the last window is terminated.
  >    
  >    *Some programs generate various windows, some of which can be visible or hidden. This has the potential to cause confusion when the mode is set to 1, as the event function will execute multiple times. The default mode is 2 to eliminate this potential confusion. 
  >   
  > - **Period/Delay**
  >    - Period: The interval (in milliseconds) to check for the window's existence.
  >    - Delay (WinEvent): The approximate delay (in milliseconds) for checking the existence of the window after a window event message is fired by the SetWinEventHook function.


### Process Parameters
  > - **Monitoring** 
  >    - **WMI**
  >    		- Check for the existence of the process at a specified time interval using WMI Provider Host process. (Default)
  >    - **Timer**
  >        - Check for the existence of the process at a specified time interval.
  >
  > - **Mode**
  >    - **1:** Call "Function_Created" for every process ID created. Call "Function_Terminated" for every process ID terminated.
  >    - **2:** Call "Function_Created" only for the initial process ID created. Call "Function_Terminated" only when the last process ID is terminated. (Default)
  >    - **3:** Call "Function_Created" for every process ID created. Call "Function_Terminated" only when the last process ID is terminated.
  >  
  >    *Some programs generate various processes with the same name. This has the potential to cause confusion when the mode is set to 1, as the event function will execute multiple times. The default mode is 2 to eliminate this potential confusion.
  >
  > - **Period**
  >    - Period: The interval (in milliseconds) to check for the presence of the process.
  
### Device Parameters
  > - **DeviceName**
  >    - Names of the device.
  >
  > - **DeviceID**
  >    - ID of the device.
  >
  > - **Monitoring** 
  >    - **DeviceChange** (Default)
  >        - Send message notifications when there is a change to the hardware configuration of a device or the computer.
  >    - **Timer**
  >        - Check for the existence of the device at a specified time interval.
  > 
  > - **Mode**
  >    - **1:** Call "Function_Created" for every device connected. Call "Function_Terminated" for every device disconnected.
  >    - **2:** Call "Function_Created" only for the initial device connected. Call "Function_Terminated" only when the last device is disconnected. (Default)
  >    - **3:** Call "Function_Created" for every device connected. Call "Function_Terminated" only when the last device is disconnected.
  >
  > - **Period/Delay**
  >    - Period: The interval (in milliseconds) to check for the device's existence.
  >    - Delay (DeviceChange): The approximate delay (in milliseconds) for checking the existence of the device after a device event message is fired by the DeviceChange function.

# Methods


## SetProfile

	SetProfile(Profile Name)
  
For example:

	WinExeCmd.SetProfile('All Events Disable')


## SetEvent
	SetEvent(State, Event name)

  > - **State**
  >    - **0:** Disable
  >    - **1:** Enable

For example:

	WinExeCmd.SetEvent(1, 'Calculator_WinSetAlwaysOnTop')	


## SetPeriodWMI
	SetPeriodWMI(Period)
  
For example:
  
	WinExeCmd.SetPeriodWMI(3000)


## ProcessFinder
Returns an array containing objects with all existing processes that match the specified parameters. If there are no matching processes, an empty array is returned.
        
	ProcessFinder(ProcessName, ProcessPath)   

For example:

	aObjProcessFinder := WinExeCmd.ProcessFinder('notepad.exe')


## WindowFinder
Returns an array containing objects with all existing windows that match the specified parameters. If there are no matching windows, an empty array is returned.
        
	WindowFinder(WinTitle, WinClass, ProcessName, ProcessPath, WinTitleMatchMode, DetectHiddenWindows, WinActive, WinMinMax)

  > - **WinTitleMatchMode**
  >    - **1:** A window's title must start with the specified WinTitle to be a match.
  >    - **2:** A window's title can contain WinTitle anywhere inside it to be a match. (Default)
  >    - **3:** A window's title must exactly match WinTitle.
  >    - **RegEx:** Regular expression WinTitle matching.
  > 
  > - **DetectHiddenWindows**
  >    - **0:** Hidden windows are not detected. (Default)
  >    - **1:** Hidden windows are detected.
  > 
  > - **WinActive**
  >    - **0**
  >    - **1**  
  > 
  > - **WinMinMax**
  >    - **0:** The window is neither minimized nor maximized.
  >    - **1:** The window is maximized.
  >    - **-1:** The window is minimized.  
  
For example:
        
	aObjWindowFinder := WinExeCmd.WindowFinder(, 'WordPadClass', 'wordpad.exe')


## DeviceFinder
Returns an array containing objects with the device matching the specified parameters. If there are no matching device, an empty array is returned.

	DeviceFinder(DeviceName, DeviceID)
 
  > - **DeviceName**
  >    - Name of the device.
  >
  > - **DeviceID**
  >    - ID of the device.

For example:

	aObjDeviceFinder := WinExeCmd.DeviceFinder('Kingston DataTraveler 3.0 USB Device')
  
Check if a device is connected:  
 	
	if WinExeCmd.DeviceFinder(,'USBSTOR\DISK&VEN_KINGSTON&PROD_DATATRAVELER_3.0&REV_\E0D55EA573DCF450E97C104C&0').Length
		WinExeCmd.MsgBox('The device is connected',, 'iconi')
  

## Notify
Display a notifications GUI.
        
	Notify(hdTxt, bdTxt, Icon, options, position, duration, callback, sound, iconSize, hdFontSize, hdFontColor, hdFont, bdFontSize, bdFontColor, bdFont, bgColor, style)

> - **hdTxt**
>    - Header text.
>
> - **bdTxt**
>    - Body text.
>
> - **icon**
>    - Picture controls https://www.autohotkey.com/docs/v2/lib/GuiControls.htm#Picture
>    - "icon!" , "icon?", "iconx", "iconi"
>    - Icon from dll: A_WinDir '\system32\user32.dll|Icon4'
>    - Loads picture from file: 'HICON:*' LoadPicture(A_WinDir '\System32\imageres.dll', 'Icon4 w48', &imageType)     
>  
> - **options**
>    - Default: "+Owner -Caption +AlwaysOnTop"
>
> - **position**
>    - "bottomRight", "bottomCenter", "bottomLeft", "topLeft", "topCenter", "topRight".
>    - Default: "bottomRight"
>
> - **duration**
>    - The display duration (in milliseconds) for the notification before it disappears. Set it to 0 to keep it on the screen until left-clicking on the GUI.
>    - Default: 8000
>
> - **callback**
>    - Type: Function Object
>    - A function object to call when left-clicking on the GUI.
>
> - **sound**
>    - The path of the .wav file to be played. https://www.autohotkey.com/docs/v2/lib/SoundPlay.htm
>    - WinExeCmd.mSounds['Windows Ding'], WinExeCmd.mSounds['tada'] etc...
>    - WAV files located in the "Sounds" folder at the root directory. WinExeCmd.mSounds['Insert filename']
>
> - **iconSize**
>    - Default: 32
>
> - **hdFontSize**
>    - Default: 15
>
> - **hdFontColor**
>    -  Default: 'white'
>
> - **hdFont**
>    - Default: 'Segoe UI bold'
>
> - **bdFontSize**
>    - Default: 12
>
> - **bdFontColor**
>    - Default: 'white'
>
> - **bdFont**
>    - Default: 'Segoe UI'
>
> - **bgColor**
>    - Default: '1F1F1F'
>
> - **style**
>    - Rounded or edged corners. "round" or "edge".
>    - Default: "round"

Example 1:

	WinExeCmd.Notify('The header text', 'The body text', A_WinDir '\system32\user32.dll|Icon4')

Example 2:

	WinExeCmd.Notify('WinExeCommander', 'Calculator does not exist.', 'HICON:*' LoadPicture(A_WinDir '\System32\imageres.dll', 'Icon8 w48', &imageType))


## MsgBox
Display a custom MsgBox without stopping the thread.

	MsgBox(text, title, icon, options, owner, winSetAoT, posXY, sound, objBtn, iconSize, fontSize, btnWidth, btnHeight)

> - **text**
>    - The text to display inside the GUI.
>
> - **title**
>    - The title of the GUI. 
>    - Default: A_ScriptName
>
> - **icon**
>    - Picture controls https://www.autohotkey.com/docs/v2/lib/GuiControls.htm#Picture
>    - "icon!" , "icon?", "iconx", "iconi"
>    - Icon from dll: A_WinDir '\system32\user32.dll|Icon4'
>    - Loads picture from file: 'HICON:*' LoadPicture(A_WinDir '\System32\imageres.dll', 'Icon4 w48', &imageType)     
>                
> - **options**
>    - Sets various options and styles for the appearance and behavior of the window. https://www.autohotkey.com/docs/v2/lib/Gui.htm#Opt
>    - Default: "-MinimizeBox -MaximizeBox"
>
> - **owner**
>    - Type: GUI object
>    - To make the window owned by another.
>
> - **winSetAoT**
>    - WinSetAlwaysOnTop https://www.autohotkey.com/docs/v2/lib/WinSetAlwaysOnTop.htm
>    - **0:** turns off the setting. (Default)
>    - **1:** turns on the setting.
>
> - **posXY**
>
> - **sound**
>    - The path of the .wav file to be played. https://www.autohotkey.com/docs/v2/lib/SoundPlay.htm
>    - "soundx" , "soundi"
>    - WinExeCmd.mSounds['Windows Ding'], WinExeCmd.mSounds['tada'] etc...
>    - WAV files located in the "Sounds" folder at the root directory. WinExeCmd.mSounds['Insert filename']
>
> - **aObjBtn**
>    - Type: Array of Object
>    - The button(s) of the GUI.
>    - Default: [{name:'*OK', callback:'this.MsgBox_Destroy'}]
>
> - **iconSize**
>    - Default: 32
> 
> - **fontSize**
>    - Default: 10
> 
> - **btnWidth**
>    - Default: 100
> 
> - **btnHeight**
>    - Default: 30

Example 1:

	WinExeCmd.MsgBox('The script file failed to open.', 'Error', A_WinDir '\system32\user32.dll|Icon4')

Example 2:

	WinExeCmd.MsgBox('Are you sure you want to delete all selected items?',, "icon?",,,,,,
	[{name:'*Yes', callback: 'Btn_Yes_Click'},
	 {name: 'Cancel', callback: 'this.MsgBox_Destroy'}])
	
	Btn_Yes_Click(g, *) {
	    
	    WinExeCmd.Notify(, 'You clicked the Yes Button.', 'iconi',,,,, 'soundi')
	    WinExeCmd.MsgBox_Destroy(g)
	}


## Sound
Plays a sound.

	Sound(sound)

> - **sound**
>    - SoundPlay. https://www.autohotkey.com/docs/v2/lib/SoundPlay.htm
>    - "soundx" , "soundi"
>    - WinExeCmd.mSounds['Windows Ding'], WinExeCmd.mSounds['tada'] etc... see Example 2.
>    - WAV files located in the "Sounds" folder at the root directory. WinExeCmd.mSounds['Insert filename']

Example 1:

	WinExeCmd.Sound('soundx')

Example 2:
  
	WinExeCmd.Sound(WinExeCmd.mSounds['Windows Ding'])


## Donation
  - If you found this script useful and would like to donate. It would be greatly appreciated. Thank you!
    https://www.paypal.com/paypalme/martinchartier/10


## License
  - MIT

## Credits
* **AutoHotkey**
  - Authors: Chris Mallett and Steve Gray (Lexikos), with portions by AutoIt Team and various AHK community members.
  - License: GNU General public license
  - Info and source code at: https://autohotkey.com/
  
* **JSON.ahk by thqby, HotKeyIt.**
  - https://github.com/thqby/ahk2_lib/blob/master/JSON.ahk
  - https://github.com/HotKeyIt/Yaml

* **GuiCtrlTips by just me.**
  - https://github.com/AHK-just-me/AHKv2_GuiCtrlTips  

* **LV_GridColor by just me.**
  - https://www.autohotkey.com/boards/viewtopic.php?f=83&t=125259 

* **GuiButtonIcon by FanaticGuru.**	
  - https://www.autohotkey.com/boards/viewtopic.php?f=83&t=115871
  
* **DisplayObj by FanaticGuru. (inspired by tidbit and Lexikos v1 code)**
  - https://www.autohotkey.com/boards/viewtopic.php?p=507896#p507896
  
* **EnumDeviceInfo by teadrinker. (based on JEE_DeviceList by jeeswg)**
  - https://www.autohotkey.com/boards/viewtopic.php?t=121125&p=537515  
  
* **GetCommandLine by teadrinker. (based on Sean and SKAN v1 code)**
  - https://www.autohotkey.com/boards/viewtopic.php?p=526409#p526409
  - https://www.autohotkey.com/board/topic/15214-getcommandline/
	
* **WTSEnumProcesses by SKAN.**
  - https://www.autohotkey.com/boards/viewtopic.php?t=4365
 
* **IsProcessElevated by jNizM**    
  - https://github.com/jNizM/ahk-scripts-v2/blob/main/src/ProcessThreadModule/IsProcessElevated.ahk

* **MoveControls by Descolada. (from UIATreeInspector.ahk)**
  - https://github.com/Descolada/UIA-v2

* **FrameShadow by Klark92.**
  - https://www.autohotkey.com/boards/viewtopic.php?f=6&t=29117&hilit=FrameShadow

* **Notify is inspired by NotifyV2 (from the-Automator.com)**
  - https://www.the-automator.com/downloads/maestrith-notify-class-v2/ 
