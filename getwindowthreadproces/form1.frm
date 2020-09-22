VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Focus the window in which you want to get it's info"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11456
      _Version        =   393216
      Cols            =   6
      AllowUserResizing=   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetForegroundWindow Lib "User32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Activate()
   SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub GetWindowInfo(i)

Dim ForeHwnd As Long
Dim ForeThread As Long
Dim ForeProcess As Long
Dim ClassName As String
Dim WindowTitle As String

ForeHwnd = GetForegroundWindow

ForeThread = GetWindowThreadProcessId(ForeHwnd, ForeProcess)

GetClassTitle ForeHwnd, ClassName, WindowTitle


MSFlexGrid1.TextMatrix(i, 1) = ForeHwnd
MSFlexGrid1.TextMatrix(i, 2) = WindowTitle
MSFlexGrid1.TextMatrix(i, 3) = ForeThread
MSFlexGrid1.TextMatrix(i, 4) = ForeProcess
MSFlexGrid1.TextMatrix(i, 5) = ClassName




End Sub

Private Sub Form_Load()
Dim str As String
  str = Len("Process identifier")
  MSFlexGrid1.TextMatrix(0, 0) = "Index"
  MSFlexGrid1.TextMatrix(0, 1) = "Hwnd"
  MSFlexGrid1.ColWidth(2) = 1300
  MSFlexGrid1.TextMatrix(0, 2) = "WindowTitle"
  MSFlexGrid1.ColWidth(3) = 1300
  MSFlexGrid1.TextMatrix(0, 3) = "Thread identifier"
  MSFlexGrid1.ColWidth(4) = 1300
  MSFlexGrid1.TextMatrix(0, 4) = "Process identifier"
  MSFlexGrid1.TextMatrix(0, 5) = "ClassName"
End Sub

Private Sub Timer1_Timer()
Static i
i = i + 1

If i = 24 Then MSFlexGrid1.Clear: i = 1: Form_Load

MSFlexGrid1.AddItem i, i
GetWindowInfo i
End Sub

Private Sub GetClassTitle(WHwnd As Long, TitleName As String, ClassName As String)
     
    Dim res As Long
    
    TitleName = String(100, Chr$(0))
    
    GetWindowText WHwnd, TitleName, 100
    
    TitleName = Left$(TitleName, InStr(TitleName, Chr$(0)) - 1)
    
    ClassName = Space(256)
    
    res = GetClassName(WHwnd, ClassName, 256)
    'res is the char number so we can trim the string
    ClassName = Left$(ClassName, res)

End Sub

