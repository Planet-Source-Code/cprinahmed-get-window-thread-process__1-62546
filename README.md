<div align="center">

## get window thread process


</div>

### Description

use this project to get the process identifier of any window

rather than makeing a call to EnumProcesses,it also learn how to

get the classname,titlename,Foregroundwindow and thread

identifier of the window.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2005-09-12 20:40:20
**By**             |[cprinahmed](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cprinahmed.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[get\_window1932209122005\.zip](https://github.com/Planet-Source-Code/cprinahmed-get-window-thread-process__1-62546/archive/master.zip)

### API Declarations

```
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
```





