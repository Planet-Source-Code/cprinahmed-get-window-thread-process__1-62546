use this project to get the process identifier of any window

rather than makeing a call to EnumProcesses,it also learn how to

get the  classname,titlename,Foregroundwindow and thread 

identifier of the window. 


Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long


email us for any comments or questions.
     
support@cpringold.atspace.com

www.cpringold.atspace.com