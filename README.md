<div align="center">

## TWTCIPIAO \- Three ways to check if program is already open


</div>

### Description

This code shows 3 ways to check if program is already open. First way doesn't use any API (it uses App.PrevInstance property). Second way to find out if program is already open is to use FindWindow function. Third and the best way is to create mutex object.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Filip Wielewski](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/filip-wielewski.md)
**Level**          |Beginner
**User Rating**    |4.8 (43 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/filip-wielewski-twtcipiao-three-ways-to-check-if-program-is-already-open__1-48018/archive/master.zip)

### API Declarations

```
Option Explicit
'API function for 2nd way
Public Declare Function FindWindow Lib "user32" _
Alias "FindWindowA" (ByVal lpClassName _
As String, ByVal lpWindowName As String) As Long
'API functions for 3rd way
Public Declare Function CreateMutex Lib _
"kernel32" Alias "CreateMutexA" _
(lpMutexAttributes As Any, ByVal bInitialOwner _
As Long, ByVal lpName As String) As Long
Public Declare Function ReleaseMutex _
Lib "kernel32" (ByVal hMutex As Long) As Long
Public Declare Function CloseHandle _
Lib "kernel32" (ByVal hObject As Long) As Long
```


### Source Code

```
'3WTCIPIAO - 3 ways to check if program
'is already open
'By Filip Wielewski
'Sorry for my english
'==================1==================
'First way doesn't use any API.
'For example:
Option Explicit
Private Sub Form_Load()
 'If this program is already open
 'then end
 If App.PrevInstance = True Then End
End Sub
'App.PrevInstance may be True or
'False. If it is True, that means this
'program is already open.
'!!! - if there are already open the
'same programs but exe files' paths
'are different then App.PrevInstance
'retrieves False. Use App.PrevInstance
' only if paths of your programs
'are the same.
'==================2==================
'Second way to find out if program
'is already open is to use FindWindow
'function. For example:
Option Explicit
Private Sub Form_Initialize()
 'Find window that has the same
 'caption as Form1
 If FindWindow(vbNullString, _
 Form1.Caption) <> 0 Then End
End Sub
'When function retrieves 0 that means
'there is no window with caption like
'in Form1 (if window is found then
'function retrieves handle to it).
'But if there exists Explorer's
'window which caption is the same
'like Form1.Caption then function
'does also retrieve a handle to that
'window!
'To avoid that situation you have
'to know name of class of your Form
'(window). For example:
Option Explicit
Private Sub Form_Initialize()
 'Find window that has the same
 'caption like Form1
 If FindWindow("ThunderFormDC", _
 Form1.Caption) <> 0 Then End
End Sub
'!!! - Use FindWindow function in
'Form_Initialize(), not in Form_Load()
'because when you use FindWindow
'function in Form_Load() then program
'will find itself (in Form_Initialize()
'form isn't loaded yet so you can
'use FindWindow function to check if
'program is already open).
'==================3==================
'Third and the best way is to create
'mutex object. For example:
Option Explicit
Dim lonMutex As Long 'It will store a
           'handle to
           'mutex object.
Private Sub Form_Load()
 Const ERROR_ALREADY_EXISTS = 183&
 'Is this application already open?
 '(If it is open then end program)
 lonMutex = CreateMutex(ByVal 0&, _
 1, App.Title)
 If (Err.LastDllError = 183&) Then
  'free memory
  ReleaseMutex lonMutex
  CloseHandle lonMutex
  End
 End If
End Sub
Private Sub Form_Unload(Cancel As _
Integer)
 'free memory
 ReleaseMutex lonMutex
 CloseHandle lonMutex
End Sub
'CreateMutex function creates mutex
'object which represents our
'application in memory. If Err object
'returns error ERROR_ALREADY_EXISTS
'that means mutex object representing
'our application already exists.
'In this case free memory
'destroying mutex object and closing
'handle to it.
```

