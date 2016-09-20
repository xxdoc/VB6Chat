VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock wskPrimary 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7000
   End
   Begin MSWinsockLib.Winsock wskChild 
      Index           =   1
      Left            =   600
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7000
   End
   Begin MSWinsockLib.Winsock wskChild 
      Index           =   2
      Left            =   1080
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7000
   End
   Begin MSWinsockLib.Winsock wskChild 
      Index           =   3
      Left            =   1560
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7000
   End
   Begin MSWinsockLib.Winsock wskChild 
      Index           =   4
      Left            =   2040
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7000
   End
   Begin MSWinsockLib.Winsock wskChild 
      Index           =   5
      Left            =   2520
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7000
   End
   Begin MSWinsockLib.Winsock wskChild 
      Index           =   0
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim u As clsUsers
Dim uI As Collection

Private Sub Form_Load()
    Set u = New clsUsers
    Set uI = New Collection
    wskPrimary.Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set u = Nothing
    Set uI = Nothing
End Sub

Private Sub BroadCast(Msg As String)
    Dim w As Winsock
    For Each w In wskChild
        If w.State = sckConnected Then
            w.SendData Msg & Chr(0)
            DoEvents
        End If
    Next
End Sub

Private Sub Tell(Index As Integer, Msg As String)
    If wskChild(Index).State = sckConnected Then
        wskChild(Index).SendData Msg & Chr(0)
        DoEvents
    End If
End Sub

Private Sub wskChild_Close(Index As Integer)
    If UserExists(Index) Then
        BroadCast "/userleft " & uI(CStr(Index))
        BroadCast "User " & uI(CStr(Index)) & " has left"
    
        u.Remove uI(CStr(Index))
        uI.Remove CStr(Index)
    End If
End Sub

Private Function UserExists(Index As Integer) As Boolean
    On Error GoTo ErrTrap
    Dim s As String
    UserExists = True
    s = uI(CStr(Index))
    Exit Function
ErrTrap:
    UserExists = False
End Function

Private Sub wskChild_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Msg As String
    Dim cmd As String
    Dim i As Integer
    Dim Key As String
    Dim p As clsUser
    Dim s As String

    wskChild(Index).GetData Msg, , bytesTotal
    Key = CStr(Index)
    If UserExists(Index) Then
        Set p = u(uI(Key))
    Else
        Set p = Nothing
    End If
    
    If Left(Msg, 1) = "/" Then
        Msg = Right(Msg, Len(Msg) - 1)
        i = InStr(1, Msg, " ")
        If i <= 0 Then
            cmd = Msg
        Else
            cmd = Left(Msg, i - 1)
            Msg = Mid(Msg, i + 1, Len(Msg) - i)
        End If
        cmd = Trim(LCase(cmd))
        Msg = Trim(Msg)
        
    
        Select Case cmd
            Case "connect"
                If Msg <> "" Then
                    If u.UserExists(Msg) Then
                        Tell Index, "/connect dupnick"
                    Else
                        s = ""
                        For Each p In u
                            s = s & p.NickName & Chr(13)
                        Next
                        If s <> "" Then
                            Tell Index, "/userlist " & Left(s, Len(s) - 1)
                        End If
                        DoEvents
                        u.Add Msg
                        uI.Add Msg, Key
                        BroadCast "/userjoined " & Msg
                        BroadCast "User " & Msg & " has Connected"
                    End If
                Else
                    Tell Index, "/connect invalidnick"
                End If
            Case "nick"
                If Not p Is Nothing Then
                    If Msg <> "" Then
                        If u.UserExists(Msg) Then
                            If Not u(Msg) Is p Then
                                Tell Index, "/nick dupnick"
                            End If
                        Else
                            BroadCast "User " & p.NickName & " is known as " & Msg
                            p.NickName = Msg
                        End If
                    Else
                        Tell Index, "/nick invalid"
                    End If
                Else
                    Tell Index, "/nick disconnected"
                End If
            Case "quit"
                
                wskChild(Index).Close
            Case Else
                Tell Index, "Unknown Command : " & cmd
        End Select
        
        Set p = Nothing
    Else
        BroadCast "<" & p.NickName & ">" & Msg
    End If
End Sub


Private Sub wskPrimary_ConnectionRequest(ByVal requestID As Long)
    Dim i As Long
    Text1.SelText = "Connection Request" & vbCrLf
    For i = wskChild.LBound To wskChild.UBound
        If wskChild(i).State = sckClosing Then
            wskChild(i).Close
        End If
        If wskChild(i).State = sckClosed Then
            wskChild(i).Accept requestID
            Text1.SelText = "Connection Accepted at " & i & vbCrLf
            Exit For
        End If
    Next
End Sub

