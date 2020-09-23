VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Portscan 0.1 by Thomas Raben"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Control"
      Height          =   2295
      Left            =   2820
      TabIndex        =   14
      Top             =   0
      Width           =   1515
      Begin VB.CommandButton cmdDie 
         Caption         =   "Exit"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1275
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   660
      Top             =   3840
   End
   Begin MSWinsockLib.Winsock sckScan 
      Left            =   180
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Port Info"
      Height          =   2715
      Left            =   0
      TabIndex        =   7
      Top             =   2340
      Width           =   4335
      Begin VB.PictureBox ProgressBG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   3375
         Begin VB.PictureBox Progress 
            BackColor       =   &H80000002&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   0
            ScaleHeight     =   21
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   225
            TabIndex        =   21
            Top             =   0
            Width           =   3375
            Begin VB.Label lblProgress 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "Label3"
               ForeColor       =   &H80000009&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   22
               Top             =   0
               Width           =   3375
            End
         End
         Begin VB.Label lblProgress 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            Caption         =   "Label3"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   3375
         End
      End
      Begin VB.Timer tmrConnected 
         Left            =   1140
         Top             =   1860
      End
      Begin VB.ListBox lstResult 
         Height          =   1980
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label lblScan 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Scanning:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "IP && Range"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.CheckBox Check1 
         Caption         =   "Reconize Services."
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         TabIndex        =   16
         Text            =   "1000"
         Top             =   1500
         Width           =   675
      End
      Begin VB.TextBox txtTime 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Text            =   "100"
         Top             =   1080
         Width           =   675
      End
      Begin VB.TextBox txtTo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1980
         TabIndex        =   5
         Text            =   "200"
         Top             =   660
         Width           =   675
      End
      Begin VB.TextBox txtFrom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Text            =   "1"
         Top             =   660
         Width           =   675
      End
      Begin VB.TextBox txtIP 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Data Time:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ms."
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   17
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ms."
         Height          =   255
         Index           =   0
         Left            =   1620
         TabIndex        =   10
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Timeout:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "to:"
         Height          =   255
         Index           =   2
         Left            =   1620
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Range:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "IP Adress:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Port As Integer
Dim Scanning As Boolean

Dim RemoteOS As String

Dim PortType As Integer

Private Sub Check1_Click()
    If Me.Check1.Value = 0 Then
        Me.txtData.Enabled = False
    Else
        Me.txtData.Enabled = True
    End If
    
End Sub

Private Sub cmdDie_Click()
    MsgBox "If you like this program, then please vote. Thanks.", vbInformation, "Dont forget 2 vote"
    End
    
End Sub

Private Sub cmdScan_Click()
    If Scanning = False Then
        Me.ProgressBG.Visible = True
        
        PortType = 0
        Me.lstResult.Clear
        Me.lstResult.AddItem "Port:      Service:"
        Me.lstResult.AddItem "--------------------------------------------------------"
        Me.txtFrom.Enabled = False
        Me.txtIP.Enabled = False
        Me.txtTime.Enabled = False
        Me.txtTo.Enabled = False
        Me.txtData.Enabled = False
        Me.Check1.Enabled = False
        
        Me.cmdScan.Caption = "Cancel"
        Scanning = True
        Port = Me.txtFrom.Text + 1
        BeginScan
    Else
        Me.ProgressBG.Visible = False
        
        Me.txtFrom.Enabled = True
        Me.txtIP.Enabled = True
        Me.txtTime.Enabled = True
        Me.txtTo.Enabled = True
        If Me.Check1.Value = 1 Then Me.txtData.Enabled = True
        Me.Check1.Enabled = True
        
        Me.cmdScan.Caption = "Scan"
        Scanning = False
        Me.TimeOut.Enabled = False
        Me.tmrConnected.Enabled = False
        Me.sckScan.Close
        Port = Me.txtTo.Text
        Me.lblScan.Caption = ""
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "If you like this program, then please vote. Thanks.", vbInformation, "Dont forget 2 vote"

End Sub

'we have a connection...
Private Sub sckScan_Connect()
    Me.TimeOut.Enabled = False
    If Me.Check1.Value = 1 Then
        Me.sckScan.SendData "GET fake.html" & vbCrLf & "USER fake" & vbCrLf & "FINGER fake" & vbCrLf
        Me.tmrConnected.Interval = Me.txtData.Text
        Me.tmrConnected.Enabled = True
    Else
        Me.lstResult.AddItem Format(Me.sckScan.RemotePort, "#00000") & " - Not Checked"
        Port = Port + 1
        BeginScan
    End If
    
End Sub

Private Sub sckScan_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    Me.sckScan.GetData Data, vbString
    
    Debug.Print "PORT: " & Port
    Debug.Print "data: " & Data
    Debug.Print "---------------------------------------------------------------------------------------------"
    
    If InStr(1, Data, vbLf) <> 0 Then
        ResolvePort Data
        Me.sckScan.Close
        Port = Port + 1
        BeginScan
    End If
End Sub

Private Sub BeginScan()

    If Port <= Me.txtTo Then
        Me.sckScan.Close
        Me.sckScan.RemoteHost = Me.txtIP.Text
        Me.sckScan.RemotePort = Port
        Me.lblScan.Caption = "Port: " & Port & " @ " & Me.sckScan.RemoteHost
        Me.sckScan.Connect
        Me.TimeOut.Interval = Me.txtTime.Text
        Me.TimeOut.Enabled = True
        On Error Resume Next
        Me.Progress.Width = (Port - Me.txtFrom.Text) / (Me.txtTo.Text - Me.txtFrom.Text) * Me.ProgressBG.ScaleWidth
        Me.lblProgress(0).Caption = Int((Port - Me.txtFrom.Text) / (Me.txtTo.Text - Me.txtFrom.Text) * 100) & " %" '
        Me.lblProgress(1).Caption = Int((Port - Me.txtFrom.Text) / (Me.txtTo.Text - Me.txtFrom.Text) * 100) & " %"
    Else
        Call cmdScan_Click
    End If
    
End Sub

Private Sub TimeOut_Timer()
    Me.TimeOut.Enabled = False
    Port = Port + 1
    BeginScan
    
End Sub


Private Sub ResolvePort(Data As String)
    Dim MyType As String
    
    '**************
    '* FTP SERVER *
    '**************
    
    If InStr(1, Data, "FTP") > 0 Then
         MyType = "FTP Deamon"
         
        'Serv-U server...
        If InStr(1, Data, "Serv-U") > 0 Then
            MyType = MyType & " (Serv-U)"
        End If
        
        
    '***************
    '* HTTP DEAMON *
    '***************
    ElseIf InStr(1, UCase(Data), "HTTP") > 0 Or InStr(1, UCase(Data), "HTML") > 0 Then
        'check for version.
        MyType = "HTTP Deamon"
        
        'Microsoft server...
        If InStr(1, Data, "Microsoft") > 0 Then
            MyType = MyType & " (Microsoft)"
        'Apache
        ElseIf InStr(1, Data, "Apache") > 0 Then
            MyType = MyType & " (Apache)"
        End If

        
        
        
    '***************
    '* MAIL DEAMON *
    '***************
    ElseIf InStr(1, UCase(Data), "MAIL") > 0 Then
        'check for version.
        MyType = "MAIL Deamon"
        
        'Microsoft server...
        If InStr(1, Data, "Microsoft") > 0 Then
            MyType = MyType & " (Microsoft)"
        End If
    
    
    '***************
    '* IMAP DEAMON *
    '***************
    ElseIf InStr(1, UCase(Data), "IMAP") > 0 Then
        'check for version.
        MyType = "IMAP Deamon"
        
        'Microsoft server...
        If InStr(1, Data, "Microsoft") > 0 Then
            MyType = MyType & " (Microsoft)"
        End If

        
    '***************
    '* NNTP DEAMON *
    '***************
    ElseIf InStr(1, UCase(Data), "NNTP") > 0 Then
        'check for version.
        MyType = "NNTP Deamon"
        
        'Microsoft server...
        If InStr(1, Data, "Microsoft") > 0 Then
            MyType = MyType & " (Microsoft)"
        End If

        
    '**************
    '* IRC DEAMON *
    '**************
    ElseIf InStr(1, UCase(Data), "NOTICE AUTH") > 0 Then
        'check for version.
        MyType = "IRC Deamon"
    ElseIf InStr(1, Data, "ERROR: Your host is trying to (re)connect too fast") > 0 Then
        'check for version.
        MyType = "IRC Deamon"
        
    '***************
    '* PING DEAMON *
    '***************
    ElseIf Mid(Data, 1, Len("GET fake.html")) = "GET fake.html" Then
        'check for version.
        MyType = "PING Deamon"
        
    'Not know service.
    Else
        MyType = "?"
    End If
    
    
    'Add the info to the list...
    Me.lstResult.AddItem Format(Me.sckScan.RemotePort, "#00000") & " - " & MyType
End Sub

'OK THE PORT DOESN'T SEND ANY DATA TO US
Private Sub tmrConnected_Timer()
    Me.lstResult.AddItem Format(Me.sckScan.RemotePort, "#00000") & " - ?"
    Me.tmrConnected.Enabled = False
    Port = Port + 1
    BeginScan
End Sub
