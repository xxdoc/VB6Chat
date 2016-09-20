VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mNickName As String
'local variable(s) to hold property value(s)
Private mvarTag As Variant 'local copy
Public Property Let Tag(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tag = 5
    mvarTag = vData
End Property


Public Property Set Tag(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Tag = Form1
    Set mvarTag = vData
End Property


Public Property Get Tag() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    If IsObject(mvarTag) Then
        Set Tag = mvarTag
    Else
        Tag = mvarTag
    End If
End Property




Public Property Get NickName() As String
Attribute NickName.VB_UserMemId = 0
    NickName = mNickName
End Property

Public Property Let NickName(Value As String)
    mNickName = Value
End Property

Public Sub Speak(Message As String)
    
End Sub

