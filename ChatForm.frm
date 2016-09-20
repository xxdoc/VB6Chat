VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5040
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7000
   End
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IntegralHeight  =   0   'False
      Left            =   5040
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtIn 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox richOut 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"ChatForm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mNickName As String
Dim mCmd As String

Public Sub AddText(msg As String)
    Dim i As Long
    i = Len(richOut.Text)
    richOut.SelStart = i
    richOut.SelText = msg
    richOut.SelStart = i
    richOut.SelLength = Len(richOut.Text) - i
End Sub

Public Sub ShowMessage(msg As String)
    AddText msg & vbCrLf
    richOut.SelStart = richOut.SelStart + richOut.SelLength
End Sub

Public Sub ShowInformation(msg As String)
    AddText msg & vbCrLf
    richOut.SelColor = RGB(0, 128, 0)
    richOut.SelStart = richOut.SelStart + richOut.SelLength
    richOut.SelColor = RGB(0, 0, 0)
End Sub

Private Sub Form_Load()
    mCmd = ""
    Winsock1.Connect "127.0.0.1"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    richOut.Move 0, 0, ScaleWidth * 0.8, ScaleHeight - txtIn.Height
    txtIn.Move 0, ScaleHeight - txtIn.Height, ScaleWidth
    lstUsers.Move richOut.Width, 0, ScaleWidth - richOut.Width, richOut.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close
    Do While Winsock1.State <> 0
        DoEvents
    Loop
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
    Dim s As String
    If Winsock1.State <> sckConnected Then Exit Sub
    
    If KeyAscii = 13 And txtIn.Text <> "" Then
        KeyAscii = 0
        s = txtIn.Text
        Winsock1.SendData s
        txtIn.Text = ""
'
'        If Left(s, 1) = "/" Then
'            s = LCase(Right(s, Len(s) - 1))
'            Select Case s
'                Case "quit"
'                    ShowInformation "Ο χρήστης " & mNickName & " αποχώρησε από το κανάλι."
'            End Select
'        Else
'            ShowMessage "<" & mNickName & "> " & s
'        End If
    End If
End Sub

Private Sub Winsock1_Connect()
Winsock1.SendData "/connect " & mNickName
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim k() As String
    Dim i As Integer
    Dim msg As String

    Winsock1.GetData msg, , bytesTotal
    mCmd = mCmd & msg
    If InStr(1, mCmd, Chr(0)) > 0 Then
        k = Split(mCmd, Chr(0))
        mCmd = k(UBound(k))
        For i = LBound(k) To UBound(k) - 1
            ParseCmd k(i)
        Next
    End If
End Sub

Private Sub ParseCmd(MsgString As String)
    Dim msg As String
    Dim cmd As String
    Dim i As Integer
    Dim g() As String
    msg = MsgString
    If Left(msg, 1) = "/" Then
        msg = Right(msg, Len(msg) - 1)
        i = InStr(1, msg, " ")
        If i <= 0 Then
            cmd = msg
        Else
            cmd = Left(msg, i - 1)
            msg = Mid(msg, i + 1, Len(msg) - i)
        End If
        cmd = Trim(LCase(cmd))
        msg = Trim(msg)
        
        Select Case cmd
            Case "connect"
                Select Case msg
                    Case "dupnick"
                        mNickName = InputBox("Enter User Name: ")
                        Winsock1.SendData "/connect " & mNickName
                    Case "invalidnick"
                        mNickName = InputBox("Enter User Name: ")
                        Winsock1.SendData "/connect " & mNickName
                End Select
            Case "nick"
                Select Case msg
                    Case "dupnick"
                        mNickName = InputBox("Enter User Name: ")
                        Winsock1.SendData "/nick " & mNickName
                    Case "invalid"
                        mNickName = InputBox("Enter User Name: ")
                        Winsock1.SendData "/nick " & mNickName
                    Case "disconnected"
                        MsgBox "You are not connected."
                End Select
            Case "userlist"
                g = Split(msg, Chr(13))
                For i = LBound(g) To UBound(g)
                    lstUsers.AddItem g(i)
                Next
            Case "userjoined"
                lstUsers.AddItem msg
            Case "userleft"
                For i = 0 To lstUsers.ListCount - 1
                    If lstUsers.List(i) = msg Then
                        lstUsers.RemoveItem i
                        Exit For
                    End If
                Next
            Case Else
                ShowInformation "Cannot understand command: " & cmd
        End Select
    Else
        ShowMessage msg
    End If

End Sub
