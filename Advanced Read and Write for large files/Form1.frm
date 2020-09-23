VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Read and Write"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4095
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2400
      Width           =   3975
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   767
      _Version        =   393216
      Min             =   1111111
      Max             =   11500000
      SelStart        =   8000000
      TickStyle       =   3
      Value           =   8000000
      TextPosition    =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read/Write"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Current/Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Read Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   1455
   End
   Begin VB.Line Line5 
      X1              =   -120
      X2              =   4320
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Input:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4080
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[0 mins.]Time Left"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[0 mins.]Total Time"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Buffer Size:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[0]Bytes/Sec."
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[0/0]Bytes"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File I/O Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   120
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim sFile As String
Dim sName As String
  Slider1.Enabled = False
  ProgressBar1.Value = 0
  sFile = GetFileString(CommonDialog1, "Open File to be copied.", True)
  Text1.Text = sFile
    If sFile = "" Then Exit Sub
  sName = Split(sFile, "\")(UBound(Split(sFile, "\")))
  Text2.Text = App.Path & "\Copy_of_" & sName
     Timer1.Enabled = True
       'call our read/write sub
       ' sFile = File to copy and path
       ' nsFile = Copy of file and path (i just have it default and
       '          save to the same location as the app.)
       FileCopy sFile, App.Path & "\Copy_of_" & sName, Slider1.Value
  Timer1.Enabled = False
  ProgressBar1.Value = ProgressBar1.Max
  Label2.Caption = "[" & TotalBytes & "/" & TotalBytes & "]Bytes"
  Slider1.Enabled = True
      MsgBox "Success! The copied file's location is: (" & Text2.Text & ")", vbOKOnly + vbInformation, "Read/Write Complete"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Slider1_Scroll()
ReDim bBytes(Slider1.Value) As Byte
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'this is just our timer to show the calculated read/write rate, total time
'it will take, time left, and current read/written out of total size
Dim sSec  As String
Dim sMin  As String
Dim sBuff As String
  Label2.Caption = "[" & CurrentBytes & "/" & TotalBytes & "]Bytes"
  ProgressBar1.Max = TotalBytes
  ProgressBar1.Value = CurrentBytes
  Label3.Caption = "[" & CurrentRead & "]Byes/Sec."
  sSec = Int(TotalBytes / CurrentRead)
  sMin = Int(sSec / 60)
    If sMin = 0 Then
      Label5.Caption = "[" & sSec & " Second(s)]Total Time"
    Else
      Label5.Caption = "[" & sMin & " Minute(s)]Total Time"
    End If
  sBuff = (TotalBytes - CurrentBytes)
  sSec = Int(sBuff / CurrentRead)
  sMin = Int(sSec / 60)
    If sMin = 0 Then
      Label6.Caption = "[" & sSec & " Second(s)]Time Left"
    Else
      Label6.Caption = "[" & sMin & " Minute(s)]Time Left"
    End If
  CurrentRead = 0
End Sub
