VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Mobile Connection"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WTP Status"
      Height          =   1635
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   8535
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "WTP Response Transmitter IP Address: 000.000.000.000"
         Top             =   1200
         Width           =   8295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "No Mobile Connection Response Received"
         Top             =   840
         Width           =   8295
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Server Offline"
         Top             =   360
         Width           =   8295
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   8400
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command Execution"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton chameleonButton1 
         Caption         =   "Browse For File"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   8295
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   8295
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   8400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File To Execute Upon Mobile Connection Response:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   8295
      End
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   2640
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   8400
   End
   Begin MSWinsockLib.Winsock Sender 
      Index           =   0
      Left            =   2160
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Listener 
      Left            =   1680
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim FreeConn(10) As Boolean


Private Sub chameleonButton1_Click()

With CMD
    .FileName = ""
    .DialogTitle = "Select A File To Execute Upon WTP Response"
    .Filter = "All Files (*.*)|*.*"
    .ShowOpen
End With

If CMD.FileName = "" Then Exit Sub

Text2.Text = CMD.FileName

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Integer
    
For i = 0 To 10

        If i <> 0 Then Load Sender(i)
        FreeConn(i) = True
Next
    
Listener.Listen

Text3.Text = "Server Online"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

Listener.Close
Unload Me

End Sub

Private Sub listener_Close()

    Listener.Close
    Listener.Listen
    
End Sub

Private Sub listener_ConnectionRequest(ByVal requestID As Long)

Dim i As Integer

    For i = 0 To 10
        If FreeConn(i) = True Then
        
            FreeConn(i) = False
            
            Sender(i).Accept requestID
            Sender(i).Close
            
            Text4.Text = "Mobile Connection Response Transmitter IP Address: " & Sender(i).RemoteHostIP

            Text1.Text = "Mobile Connection Response Received OK!"
        
            FreeConn(i) = True
            
            Exit For
        End If
    Next
End Sub

Function LaunchApp(ByVal URL As String) As Long

On Error Resume Next

Dim strFile As String
strFile = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)

End Function

Private Sub Timer1_Timer()

If Text1.Text = "Mobile Connection Response Received OK!" Then
LaunchApp (Text2.Text)
Text1.Text = "Mobile Connection Was Successful"
End If

End Sub
