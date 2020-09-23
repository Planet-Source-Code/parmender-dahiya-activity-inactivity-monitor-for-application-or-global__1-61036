<div align="center">

## Activity/Inactivity Monitor for Application or Global


</div>

### Description

How to get the inactivity timeout for an application or globally? Here is the answer. Application timeout means a particular application is running but no mouse or keyboard activity is there for that application for a particluar time. On PSC I found only for Global but this code has both application and global.I have made for the application also. This code can be directly incorporated into your application in minutes.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Parmender Dahiya](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/parmender-dahiya.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/parmender-dahiya-activity-inactivity-monitor-for-application-or-global__1-61036/archive/master.zip)

### API Declarations

```
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
```


### Source Code

```
'************************************************************
'HOW TO USE THIS CODE
' 1. ADD a Timer Control with name ActivityTimer on form with 100 millisecond interval
' 2. ADD a Label with name label1
' 3. COPY and PASTE this code into your code in the form where the above controls are added
'************************************************************
Option Explicit
' API Definitions
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
 x As Long
 y As Long
End Type
Dim GlobalTimeout As Boolean
' this is the time you want to set for inactivity
' if you want to define 60 seconds then TimeOutTiem constant must be
' multiplied by 10 because timer control on form is 100 millisecond
Private Const TimeoutTime = 50
Private Sub ActivityTimer_Timer()
Dim KeyPressed As Boolean, MouseMoved As Boolean, tmpPos As POINTAPI
Dim i As Integer, Ret As Long
Static MyCount As Long, OtherCount As Long
Static MousePos As POINTAPI
KeyPressed = False
MouseMoved = False
GlobalTimeout = False
' find if any of the keyboard is pressed
' if a key is pressed on keyboard then value returned by
' GetAsyncKeyStatewill be -32767
For i = 0 To 255
 If GetAsyncKeyState(i) = -32767 Then
  KeyPressed = True
  Exit For
 End If
Next
' get the mouse position
Ret = GetCursorPos(tmpPos)
' compare with previous mouse positions
If tmpPos.x <> MousePos.x Or tmpPos.y <> MousePos.y Then
 MouseMoved = True
 MousePos.x = tmpPos.x
 MousePos.y = tmpPos.y
End If
' if NO keyboard or mouse activity then increment the count
' if activity then two things two be checked
' 1. If it is on any other application or on our application
' 2. If we have to check for Global or Local (Variable GlobalTimeout)
If KeyPressed = True Or MouseMoved = True Then
 If GlobalTimeout = False Then ' if checked for our application
  If GetForegroundWindow() = Me.hwnd Then
   MyCount = 0
  Else
   MyCount = MyCount + 1
  End If
 Else  ' if checked globally
  MyCount = 0
 End If
Else  ' if no keyboard or mouse activity
 MyCount = MyCount + 1
End If
Label1.Caption = MyCount
' if count has exceeded the limit
If MyCount >= TimeoutTime Then
 MsgBox "Timeout Occurred after " & MyCount / 10 & " seconds", vbOKOnly + vbCritical, "No Avtivity"
 End  ' TAKE ACTION HERE AFTER THE TIMEOUT
End If
End Sub
```

