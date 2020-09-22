VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FtpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private LocalInet As Inet
Private FileList As ListBox
Private ResText As TextBox
Private ResTextOut$
Private FormOut As Form
Public CancelOperation As Boolean
Public GetFileSize As Long
Public Function GetValues$()
    Dim vtData As Variant
    FormOut.MousePointer = 11
    Do While LocalInet.StillExecuting
        DoEvents
    Loop
    vtData = LocalInet.GetChunk(1024, icString)
    a$ = ""
    Do While LenB(vtData) > 0
        a$ = a$ & vtData
        vtData = LocalInet.GetChunk(1024, icString)
    Loop
    GetValues$ = a$
    FormOut.MousePointer = 1

End Function


Public Sub GetValuesInList(ResFileList As ListBox, ResDirList As ListBox)
    a$ = GetValues$
    ResFileList.Clear
    ResDirList.Clear
    Do While InStr(a$, Chr$(13)) > 0
        b$ = Mid$(a$, 1, InStr(a$, Chr$(13)) - 1)
        If Right$(b$, 1) = "/" Then
            ResDirList.AddItem b$
        Else
            ResFileList.AddItem b$
        End If
        a$ = Mid$(a$, InStr(a$, Chr$(13)) + 2)
    Loop
End Sub


Public Sub OpenUrl(Url$, Username$, Password$, Proxy)
    FormOut.MousePointer = 11
    With LocalInet
        If Proxy Then .AccessType = icUseDefault
        .Url = Url$
        .Username = Username$
        .Password = Password$
    End With
    WriteMessage "Opening " + Url$
    WriteMessage "User " + Username$
    WriteMessage "Password " + Password$
    FormOut.MousePointer = 1
End Sub


Public Sub SetInetComponent(DefInet As Inet, OutText As TextBox)
    Set LocalInet = DefInet
    Set ResText = OutText
    Set FormOut = OutText.Parent
End Sub




Public Sub InetExecute(ByVal InetCommand$)
    If CancelOperation Then Exit Sub

    On Error GoTo ExecuteError
    FormOut.MousePointer = 11
    Do While LocalInet.StillExecuting
        If CancelOperation Then
            FormOut.MousePointer = 1
            Exit Sub
        End If
        DoEvents
    Loop
    WriteMessage InetCommand$
    LocalInet.Execute , InetCommand$
    FormOut.MousePointer = 1
    Exit Sub
ExecuteError:
    FormOut.MousePointer = 1
    MsgBox Error$
    Exit Sub
End Sub
Public Sub SetResultList(InetList As ListBox)
    Set FileList = InetList
End Sub

Private Sub WriteMessage(Message$)
    ResTextOut$ = ResTextOut$ + Chr$(13) + Chr$(10) + "Executing: " + Message$
    ResText.Text = ResTextOut$
    If Len(ResText.Text) > 1000 Then ResText.Text = Mid$(ResText.Text, 500)
    ResText.SelStart = Len(ResText.Text)
End Sub

