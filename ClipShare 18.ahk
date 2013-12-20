;***********************Clipshare version .99********************************.
; place this program in a folder which you have shared with multiple computers.  Press Windows C to copy and Windows V to paste.  
; Plain text can be copied/pasted from most computers
; One file at a time can be copied/pasted when using Windows Explorer
; Joe@SPSSGod.Com
;*******************************************************.
#SingleInstance force
#Noenv
#MaxThreads, 3
DetectHiddenWindows, On
SetBatchLines -1
Menu, Tray, Icon, Shell32.dll, 113 ;Blue
SetWorkingDir %A_ScriptDir% 

IniRead, Sound_Notify, %A_ScriptDir%\ClipShare.ini, %A_ComputerName%, Sound
IniRead, Mouse_Notify, %A_ScriptDir%\ClipShare.ini, %A_ComputerName%, Mouse
;***********************if either setting isn't there set both to on********************************.
if (Sound_Notify="ERROR") or ( Mouse_Notify="ERROR"){
  IniWrite, 1, %A_ScriptDir%\ClipShare.ini, %A_ComputerName%, Sound
  IniWrite, 1, %A_ScriptDir%\ClipShare.ini, %A_ComputerName%, Mouse
  IniRead, Sound_Notify, %A_ScriptDir%\ClipShare.ini, %A_ComputerName%, Sound
  IniRead, Mouse_Notify, %A_ScriptDir%\ClipShare.ini, %A_ComputerName%, Mouse
  }

FileGetTime, TimeClip, %A_WorkingDir%\FileName.ini
previousTimeClip := TimeClip ,previousTimeFile := TimeFile

SetTimer, check_time, 800 ;check file status once per second
GoSub Notification

;***********************tray items********************************.
;~ Menu, Tray, Add, Sound
Menu, Tray, Add
Menu, Tray, Add, Help
Menu, Tray, Add, About this program, About
Menu, Tray, Add
menu, tray, Add, Sound Notification, Sound_Notify
menu, tray, Add, Mouse Notification, Mouse_Notify
Menu, Tray, Add, Exit
Menu, Tray, NoStandard ;removes defualt options
if (Sound_Notify=1) ;determin if sound was on from preferences and change toggle if needed
      menu, tray,Check, Sound Notification
if (Mouse_Notify=1)
      menu, tray,Check, Mouse Notification
return

Sound_Notify:
menu, tray, ToggleCheck, Sound Notification
Sound_Notify :=!Sound_Notify
IniWrite, %Sound_Notify%, %A_ScriptDir%\ClipShare.ini, %A_ComputerName%, Sound
return

Mouse_Notify:
Menu, Tray, ToggleCheck, Mouse Notification
Mouse_Notify:=!Mouse_Notify
IniWrite, %Mouse_Notify%, %A_ScriptDir%\ClipShare.ini, %A_ComputerName%, Mouse
return

Help:
Gui, Help:Destroy
Gui, Help:Add, Text,x10 y10, This program allows computers that share a mutual folder to share the clipboard.
Gui, Help:Add, Text,x10 y+15, After launching the script on both computers FROM THE SAME DIRECTORY, Copy/paste as you would normally however use the Windows Key INSTEAD of the Control Key.
Gui, Help:Add, Text,x10 y+15, To Copy hold down the windows key and press C (Alternatively you can press Alt and F1)
Gui, Help:Add, Text,x10 y+15, To Paste hold down the windows key and press V (Alternatively you can press Alt and F2)

Gui, Help:Add, Text,x10 y+15, You can copy Text or, using Windows Explorer, you can copy ONE file.
Gui, Help:Font,CBlue Underline
Gui, Help:Add,Text,y+15 GDemo, Video demo
hCurs:=DllCall("LoadCursor","UInt",NULL,"Int",32649,"UInt") ;IDC_HAND
onMessage(0x200, "MsgHandler")
Gui, Help:Show,, Help
return

Demo:
Run,http://youtu.be/N-ku9XCSEs8
  gosub GuiClose
Return

About:
Gui,About:Destroy
Gui,About:Font,Bold
Gui,About:Add,Text,x10 y10,ClipShare v 1.0
Gui,About:Font
Gui,About:Add,Text, y+5,Direct questions to 
Gui,About:Font,CBlue Underline
Gui,About:Add,Text, y+5 Ge-mail, e-mail Joe Glines
Gui,About:Add,Text,y+5 GWebsite,LinkedIn
hCurs:=DllCall("LoadCursor","UInt",NULL,"Int",32649,"UInt") ;IDC_HAND
onMessage(0x200, "MsgHandler")
 Gui,About:Font
 Gui,About:Show,, About
return

e-mail:
Subj:=A_ScriptName " AHK=" A_AhkVersion " OS=" A_OSVersion
  Run,Mailto:Joe@working-smarter-not-harder.com?Subject=%Subj%
   gosub GuiClose
Return

Website:
Run,http://www.linkedin.com/in/joeglines/
  gosub GuiClose
Return

GuiClose:
  Gui,About:Destroy
  OnMessage(0x200,"")
  DllCall("DestroyCursor","Uint",hCur)
Return

Exit:
ExitApp
Return

;**************Compare time stamps of files*****************************************.
check_time: 
FileGetTime, TimeClip, %A_WorkingDir%\FileName.ini
;~ MsgBox % previousTimeFile
If (TimeClip > previousTimeClip){
previousTimeClip := TimeClip
option1:=!option1 ; toggle
If (option1=1)
  Menu, Tray, Icon, Shell32.dll, 28 ;Red 
If (Option1=0)
  Menu, Tray, Icon, Shell32.dll, 113 ;Blue
;~ FormatTime, PrettyTime, PrettyTime, dddd MMMM d, yyyy hh:mm:ss tt
FormatTime, PrettyTime, PrettyTime, MM/dd/yy @ h:mm

Gosub Notification
return
}
Return
;~ Browser_Back::Edit
;~ Browser_Forward::Reload
;*******************************************************.
;***********************Copy********************************.
;*******************************************************.
!F1::
#c:: ;Windows C to copy whaterver and save it to correct location 
WinActivate  ; Uses the last found window.
FileDelete, %A_WorkingDir%\tmp.clip
clipboard :="" ; Empty the clipboard
  Send, ^c
    ClipWait, 2
      if ErrorLevel
        return
PrettyTime:=TimeFile

FormatTime, PrettyTime, PrettyTime, MM/dd/yy @ h:mm
IniWrite, %A_ComputerName%, %A_WorkingDir%\FileName.ini, Source, ComputerName
IniWrite, %A_UserName%, %A_WorkingDir%\FileName.ini, Source, A_UserName
IniWrite, %PrettyTime%, %A_WorkingDir%\FileName.ini, Source, Time

;***********************Determine what is in clipboard********************************.
IfWinActive ahk_class CabinetWClass ;ExploreWClass
{   
SplitPath, clipboard, OutFileName
IniWrite, %OutFileName%, %A_WorkingDir%\FileName.ini, Source, Filename
IniWrite, File, %A_WorkingDir%\FileName.ini, Type, Copied

Sort, clipboard  ; This also converts to text (full path and name of each file).
sleep, 100
Filecopy , %Clipboard%, %A_WorkingDir%\tmp.clip , 1
} 
else ;***********************copy if it was Not a file********************************.
{ 
FileAppend, %ClipBoard%, %A_WorkingDir%\tmp.clip     
IniWrite, Text, %A_WorkingDir%\FileName.ini, Type, Copied
}
return

;*******************************************************.
;*******************************************************.
!F2::
#v:: ;Windows V to paste
;*******************************************************.
IfWinActive ahk_class CabinetWClass ;ExploreWClass ;Detect if Explorer is active window- if so paste as file
{
IniRead, Filename, %A_WorkingDir%\FileName.ini, Source, Filename ;read file name from ini
CurrentPath := Explorer_GetPath() . "\" ;Get current Explorer Path
ifexist ,%Currentpath%%filename%  ;See if file already exists and prompt to overwrite
  MsgBox, 4, , %Currentpath%%filename% already exist`n`nDo you want to overwrite the file?
    ifmsgbox,no
      return    ; exit if user does not want to overwrite md5 file
FileRecycle, %Currentpath%%filename%
  Filecopy , %A_WorkingDir%\tmp.Clip, %CurrentPath%\%Filename%, 1
}
Else ;****************Paste if not file***************************************.
{ 
  clipboard :="" ; Empty the clipboard
sleep, 100
    ;~ FileRead, Clipboard, *c %A_WorkingDir%\tmp.clip ;this is old way reading in as clipboard all
    FileRead, Clipboard, %A_WorkingDir%\tmp.clip ;just text
              ClipWait, 2
          sleep, 100
  SendInput, ^v
  sleep, 100
}
return

;***********************Information about clip last saved********************************.
Notification: 
IniRead, Source_ComputerName, %A_WorkingDir%\FileName.ini, Source, ComputerName ;read file name from ini
IniRead, Source_UserName, %A_WorkingDir%\FileName.ini, Source, A_UserName ;read file name from ini
IniRead, PrettyTime, %A_WorkingDir%\FileName.ini, Source, Time
IniRead, Type, %A_WorkingDir%\FileName.ini, Type, Copied

If (Type="Text"){
FileRead, NewText, %A_WorkingDir%\tmp.clip
}Else{
IniRead, NewText, %A_WorkingDir%\FileName.ini, Source, Filename ;read file name from ini
}
;***********************Create tips showing what was copied********************************.
Message:=Source_UserName . " on " . Source_ComputerName . " :: " . PrettyTime . "`n"
Full_Tip:=Message . NewText

StringReplace, Short_Tip, Full_Tip, `r`n,`n, All
StringReplace, Short_Tip, Short_Tip, %A_Tab%, %A_Space% , All
if (StrLen(Short_Tip)>123)
  Short_Tip:=SubStr(Short_Tip,1,123) . "..."

menu , tray, tip, %Short_Tip%
If (sound_Notify=1)
  SoundPlay, %A_WinDir%\Media\Windows Ding.wav
If (Mouse_Notify=1){
ToolTip, %Full_Tip%
sleep, 2500
ToolTip
}
return