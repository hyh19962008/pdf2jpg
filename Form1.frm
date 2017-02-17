VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CovRetar"
   ClientHeight    =   4308
   ClientLeft      =   120
   ClientTop       =   768
   ClientWidth     =   7536
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4308
   ScaleWidth      =   7536
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   4128
      TabIndex        =   7
      Text            =   "200"
      Top             =   1440
      Width           =   780
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "Form1.frx":000C
      Left            =   1824
      List            =   "Form1.frx":002E
      TabIndex        =   5
      Top             =   1440
      Width           =   1260
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Start"
      Height          =   396
      Left            =   3264
      TabIndex        =   4
      Top             =   2496
      Width           =   1260
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   396
      Left            =   0
      TabIndex        =   3
      Top             =   3912
      Width           =   7536
      _ExtentX        =   13293
      _ExtentY        =   699
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10584
            MinWidth        =   10584
            Text            =   "Holding"
            TextSave        =   "Holding"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   300
      Left            =   6240
      TabIndex        =   2
      Top             =   960
      Width           =   1068
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   1824
      TabIndex        =   1
      Top             =   960
      Width           =   4236
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   864
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Ñ¡ÔñPDFÄ¿Â¼"
      Filter          =   "*.pdf|*.pdf"
      InitDir         =   "c:\"
   End
   Begin VB.Label Label4 
      Caption         =   "DPI:"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3456
      TabIndex        =   8
      Top             =   1440
      Width           =   492
   End
   Begin VB.Label Label3 
      Caption         =   "Output Format£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   192
      TabIndex        =   6
      Top             =   1440
      Width           =   1740
   End
   Begin VB.Label Label2 
      Caption         =   "Input Dirtory£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   192
      TabIndex        =   0
      Top             =   960
      Width           =   1836
   End
   Begin VB.Menu Menu1 
      Caption         =   "Register"
   End
   Begin VB.Menu Menu2 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time As Single
Dim f As Integer
Dim pb
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub Combo1_LostFocus()
f = Combo1.ListIndex
End Sub

Private Sub Command2_Click()
cd1.ShowOpen
Text1.Text = Left(cd1.Filename, Len(cd1.Filename) - Len(cd1.FileTitle))
End Sub

Private Sub Command5_Click()
Me.Cls
If Combo1.Text = "" Then
    MsgBox ("Select a Format")
    Exit Sub
End If
Dim pb As New DebenuPDFLibraryAX1016.PDFLibrary
Call pb.UnlockKey("j87ig3k84fb9eq9dy34z7u66y")  'Register the proprietary pdf dll

Dim dib As Long
Dim PdfFile As String
st = Shell("cmd /c DIR *.pdf /b >list.txt")     'Get all filenames in the diretory
Sleep 2400
Open Text1.Text & "list.txt" For Input As #1
time = Timer

Do While Not EOF(1)
    Line Input #1, PdfFile
Loop
Close
Open Text1.Text & "list.txt" For Input As #1
    
Do While Not EOF(1)
    Line Input #1, PdfFile
    Call pb.LoadFromFile(PdfFile, "")
    For i = 1 To pb.PageCount()
        SB1.Panels(1).Text = "Processing " & PdfFile & " " & "Page" & i & " of " & pb.PageCount & "..."
        Call pb.RenderPageToFile(Int(Text2.Text), i, f, PdfFile & i & Combo1.Text)
        DoEvents
    Next i
Loop

Close
Kill (Text1.Text & "list.txt")
SB1.Panels(1).Text = "Time Elasped£º" & Format(Timer - time, "0.00") & "s"
End Sub

Private Sub Form_Load()
cd1.InitDir = App.Path
App.Title = ""
End Sub

Private Sub Menu1_Click()
Dim sysdir$, dirlen%
sysdir = Space(50)
dirlen = GetSystemDirectory(sysdir, 50)
sysdir = Left(sysdir, dirlen)
tdir = Dir(sysdir & "\pdf2parts.dll")
If tdir = "" Then
    On Error GoTo ERRmsg
    Call FileCopy(App.Path & "\pdf2parts.dll", sysdir & "\pdf2parts.dll")
    Shell App.Path & "\regsvr32.exe /s " & sysdir & "\pdf2parts.dll"
End If
ERRmsg:
MsgBox "Error:" & vbCrLf & vbCrLf & "Run as Administrator"
End Sub

Private Sub Menu2_Click()
MsgBox "         CovRetar v1.1  2017.1" & vbCrLf & vbCrLf _
, vbOKOnly, "About"
End Sub
