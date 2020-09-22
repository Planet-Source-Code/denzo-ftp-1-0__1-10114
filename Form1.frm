VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "DenzoFtpClient"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   3960
      TabIndex        =   21
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4560
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   2040
      TabIndex        =   17
      Top             =   3480
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2040
      TabIndex        =   16
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Use Proxy"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   2760
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Text            =   "ftp://arnold.c64.org/pub/games/z/"
      Top             =   240
      Width           =   3495
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   4920
      Width           =   5415
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3840
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Server"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Anonymous connection"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Text            =   "pub/games/z/"
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "Password@praxis.it"
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Text            =   "ftp"
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Total="
      Height          =   255
      Left            =   3720
      TabIndex        =   20
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Server"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Remote Directory"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewInet As New FtpClient
Dim CancelOperation As Boolean

Private Sub Command1_Click()
    NewInet.CancelOperation = False
    NewInet.OpenUrl Text1.Text, Text2.Text, Text3.Text, Check2.Value
    a$ = NewInet.GetValues
    NewInet.InetExecute "CD " + Text4.Text
    NewInet.InetExecute "DIR"
    NewInet.GetValuesInList List2, List1
    Command1.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = True

End Sub

Private Sub Command2_Click()
    Dim TotalDL As Long
    NewInet.CancelOperation = False
    InPath$ = Dir1.Path
    ChDir$ InPath$
    GetDirectory InPath$ + "\", TotalDL
    EndLoop = List1.ListCount - 1
    For i = 0 To EndLoop
        On Error GoTo ServerError
        NewDir$ = Mid$(List1.List(i), 1, Len(List1.List(i)) - 1)
        If NewDir$ <> "." And NewDir$ <> ".." Then
            ChDir$ InPath$
            On Error Resume Next
            MkDir$ NewDir$
            ChDir$ InPath$ + "\" + NewDir$
            On Error GoTo 0
            NewInet.InetExecute "CD " + List1.List(i)
            NewInet.InetExecute "DIR"
            NewInet.GetValuesInList List2, List1
            GetDirectory InPath$ + "\" + NewDir$ + "\", TotalDL
            NewInet.InetExecute "CDUP"
            NewInet.InetExecute "DIR"
            NewInet.GetValuesInList List2, List1

        End If
        
    Next i
ServerError:
    Exit Sub
End Sub

Private Sub Command3_Click()
    NewInet.CancelOperation = False
    NewInet.InetExecute "CLOSE"  ' Close the connection.
    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = False

End Sub

Private Sub Command4_Click()
    NewInet.CancelOperation = True
    MsgBox "Cancel Operation"
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    NewInet.SetInetComponent Inet1, Text6
End Sub

Private Sub List1_DblClick()
    NewInet.InetExecute "CD " + List1.List(List1.ListIndex)
    NewInet.InetExecute "DIR"
    NewInet.GetValuesInList List2, List1
End Sub

Sub GetDirectory(SavePath$, TotalDL As Long)
    For o = 0 To List2.ListCount - 2
        List2.ListIndex = o
        NewInet.InetExecute "SIZE " + List2.List(o)
        FileSize& = Val(NewInet.GetValues)
        Label5 = "Loading: " & List2.List(o) & " " & FileSize&
        DoEvents
        TotalDL = TotalDL + FileSize&
        If FileSize& > 0 Then
            Label6 = "Total= " + Format(TotalDL, "###.###.###.###")
            'NewInet.GetFileSize = CLng(a$)
            If Dir$(SavePath$ + List2.List(o)) = List2.List(o) Then
                If FileLen(SavePath$ + List2.List(o)) <> FileSize& Then
                    NewInet.InetExecute "GET " + List2.List(o) + " " + SavePath$ + List2.List(o)
                End If
            Else
                NewInet.InetExecute "GET " + List2.List(o) + " " + SavePath$ + List2.List(o)
            End If
        End If
    Next o
End Sub

Private Sub List2_Click()

End Sub
