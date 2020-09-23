VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UMA - By FutureX!"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9915
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear Msg Board"
      Height          =   375
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Maker"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Version"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Activate UMA"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "Message Board"
      ForeColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   0
      TabIndex        =   25
      Top             =   5280
      Width           =   9855
      Begin VB.ListBox lstMsgBoard 
         ForeColor       =   &H00000000&
         Height          =   1815
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Password"
      ForeColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   5160
      TabIndex        =   23
      Top             =   3240
      Width           =   4695
      Begin VB.TextBox txtpass 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Text            =   "Pass (You Need To Change This)"
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Away Message"
      ForeColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   0
      TabIndex        =   22
      Top             =   3240
      Width           =   5175
      Begin VB.TextBox txtAway 
         ForeColor       =   &H00000000&
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "Form1.frx":0CCA
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.ComboBox cmbType 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      ItemData        =   "Form1.frx":0D34
      Left            =   6000
      List            =   "Form1.frx":0D4A
      TabIndex        =   20
      Text            =   "Version"
      Top             =   2760
      Width           =   3615
   End
   Begin VB.ComboBox cmbType 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      ItemData        =   "Form1.frx":0DBD
      Left            =   6000
      List            =   "Form1.frx":0DD3
      TabIndex        =   19
      Text            =   "Change Name & Status (Password needed)"
      Top             =   2400
      Width           =   3615
   End
   Begin VB.ComboBox cmbType 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      ItemData        =   "Form1.frx":0E46
      Left            =   6000
      List            =   "Form1.frx":0E5C
      TabIndex        =   18
      Text            =   "Add to Message Board"
      Top             =   2040
      Width           =   3615
   End
   Begin VB.ComboBox cmbType 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      ItemData        =   "Form1.frx":0ECF
      Left            =   6000
      List            =   "Form1.frx":0EE5
      TabIndex        =   17
      Text            =   "Read Message Board"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.ComboBox cmbType 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      ItemData        =   "Form1.frx":0F58
      Left            =   6000
      List            =   "Form1.frx":0F6E
      TabIndex        =   16
      Text            =   "Time"
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox txtmessage 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   14
      Text            =   "Version"
      Top             =   2760
      Width           =   4815
   End
   Begin VB.TextBox txtmessage 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   13
      Text            =   "Change Name"
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox txtmessage 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   12
      Text            =   "Add to Message"
      Top             =   2040
      Width           =   4815
   End
   Begin VB.TextBox txtmessage 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Text            =   "Read Messages"
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox txtmessage 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   10
      Text            =   "Time"
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txtkey 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Text            =   "6"
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtkey 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Text            =   "5"
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtkey 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Text            =   "4"
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox txtkey 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Text            =   "3"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtkey 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Text            =   "2"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Menu"
      ForeColor       =   &H00E0E0E0&
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9855
      Begin VB.ComboBox cmbType 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         ItemData        =   "Form1.frx":0FE1
         Left            =   6000
         List            =   "Form1.frx":0FF7
         TabIndex        =   15
         Text            =   "Away Message"
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtmessage 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Text            =   "Away Message"
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtkey 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Text            =   "1"
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Type To Send!"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   21
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "Message to Send"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2850
         TabIndex        =   8
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "Key to Press"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Caption         =   "FutureX - FutureXIsHere@hotmail.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   33
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Ultimate Away Messenger! (UMA)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents MsgrObj As MsgrObject
Attribute MsgrObj.VB_VarHelpID = -1
Dim msg As Boolean
Dim pass As Boolean
Dim pass2 As Boolean
Dim activate As Boolean

Private Sub Command1_Click()
lstMsgBoard.Clear
End Sub


Private Sub Command2_Click()
Select Case Command2.Caption
Case "Activate UMA"
    activate = True
    Command2.Caption = "Deactivate UMA"
Case "Deactivate UMA"
    activate = False
    Command2.Caption = "Activate UMA"
End Select
End Sub

Private Sub Command3_Click()
MsgBox "Version is 1.1!", vbInformation
End Sub

Private Sub Command4_Click()
MsgBox ("Hi People. I made this using Visual Basic 6! If you want Visual Basic 6 download it using Morpheus!" & vbCrLf & "I Made this becasue i needed a Away Messenger, with tools that i can edit!" & vbCrLf & "FutureX (FutureXIsHere@hotmail.com), Thanks for using it!")
End Sub

Private Sub Command5_Click()
lstMsgBoard.Clear
End Sub

Private Sub Form_Load()
Set MsgrObj = New MsgrObject
msg = False
pass = False
pass2 = False
activate = False
Index = 0
End Sub

Private Sub msgrobj_OnTextReceived(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal pSourceUser As Messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)

If activate = False Then Exit Sub
If pass = True Then
    If bstrMsgText = txtpass Then
        pass2 = True
        pass = False
        pSourceUser.SendText bstrMsgHeader, "Enter the new Name!", MMSGTYPE_NO_RESULT
        Exit Sub
     End If
  
    
End If
If pass2 = True Then
    MsgrObj.Services(0).FriendlyName = bstrMsgText
    pass2 = False
    Exit Sub
End If
If msg = True Then
    
     lstMsgBoard.AddItem bstrMsgText & pSourceUser
     
     msg = False
     Exit Sub
     Exit Sub
End If

If bstrMsgText = txtkey(0).Text Then
    Select Case cmbType(0).Text
    Case "Away Message"
        pSourceUser.SendText bstrMsgHeader, "Away Message:" & txtAway, MMSGTYPE_NO_RESULT
    Case "Time"
        pSourceUser.SendText bstrMsgHeader, "The Time is " & Time, MMSGTYPE_NO_RESULT
    Case "Read Message Board"
        A = 0
        Do
        A = A + 1
        If lstMsgBoard.ListCount = A Then Exit Sub
        pSourceUser.SendText bstrMsgHeader, lstMsgBoard.List(A), MMSGTYPE_NO_RESULT
        Loop
    Case "Add to Message Board"
        msg = True
        pSourceUser.SendText bstrMsgHeader, "Please Type in your message now!", MMSGTYPE_NO_RESULT
    Case "Change Name & Status (Password needed)"
        pass = True
        pSourceUser.SendText bstrMsgHeader, "Please enter your password now!", MMSGTYPE_NO_RESULT
    Case "Version"
        pSourceUser.SendText bstrMsgHeader, "The Version of this program is 1.0!", MMSGTYPE_NO_RESULT
    End Select
End If
If bstrMsgText = txtkey(1).Text Then
    Select Case cmbType(1).Text
    Case "Away Message"
        pSourceUser.SendText bstrMsgHeader, "Away Message:" & txtAway, MMSGTYPE_NO_RESULT
    Case "Time"
        pSourceUser.SendText bstrMsgHeader, "The Time is " & Time, MMSGTYPE_NO_RESULT
    Case "Read Message Board"
        A = 0
        Do
        A = A + 1
        If lstMsgBoard.ListCount = A Then Exit Sub
        pSourceUser.SendText bstrMsgHeader, lstMsgBoard.List(A), MMSGTYPE_NO_RESULT
        Loop
    Case "Add to Message Board"
        msg = True
        pSourceUser.SendText bstrMsgHeader, "Please Type in your message now!", MMSGTYPE_NO_RESULT
    Case "Change Name & Status (Password needed)"
        pass = True
        pSourceUser.SendText bstrMsgHeader, "Please enter your password now!", MMSGTYPE_NO_RESULT
    Case "Version"
        pSourceUser.SendText bstrMsgHeader, "The Version of this program is 1.0!", MMSGTYPE_NO_RESULT
    End Select
End If
If bstrMsgText = txtkey(2).Text Then
    Select Case cmbType(2).Text
    Case "Away Message"
        pSourceUser.SendText bstrMsgHeader, "Away Message:" & txtAway, MMSGTYPE_NO_RESULT
    Case "Time"
        pSourceUser.SendText bstrMsgHeader, "The Time is " & Time, MMSGTYPE_NO_RESULT
    Case "Read Message Board"
        On Error Resume Next
        A = 0
        Do
        A = A + 1
        If lstMsgBoard.ListCount = A Then Exit Sub
        pSourceUser.SendText bstrMsgHeader, lstMsgBoard.List(A), MMSGTYPE_NO_RESULT
        Loop

    Case "Add to Message Board"
        msg = True
        pSourceUser.SendText bstrMsgHeader, "Please Type in your message now!", MMSGTYPE_NO_RESULT
    Case "Change Name & Status (Password needed)"
        pass = True
        pSourceUser.SendText bstrMsgHeader, "Please enter your password now!", MMSGTYPE_NO_RESULT
    Case "Version"
        pSourceUser.SendText bstrMsgHeader, "The Version of this program is 1.0!", MMSGTYPE_NO_RESULT
    End Select
End If
If bstrMsgText = txtkey(3).Text Then
    Select Case cmbType(3).Text
    Case "Away Message"
        pSourceUser.SendText bstrMsgHeader, "Away Message:" & txtAway, MMSGTYPE_NO_RESULT
    Case "Time"
        pSourceUser.SendText bstrMsgHeader, "The Time is " & Time, MMSGTYPE_NO_RESULT
    Case "Read Message Board"
        A = 0
        Do
        A = A + 1
        If lstMsgBoard.ListCount = A Then Exit Sub
        pSourceUser.SendText bstrMsgHeader, lstMsgBoard.List(A), MMSGTYPE_NO_RESULT
        Loop
    Case "Add to Message Board"
        msg = True
        pSourceUser.SendText bstrMsgHeader, "Please Type in your message now!", MMSGTYPE_NO_RESULT
    Case "Change Name & Status (Password needed)"
        pass = True
        pSourceUser.SendText bstrMsgHeader, "Please enter your password now!", MMSGTYPE_NO_RESULT
    Case "Version"
        pSourceUser.SendText bstrMsgHeader, "The Version of this program is 1.0!", MMSGTYPE_NO_RESULT
    End Select
End If
If bstrMsgText = txtkey(4).Text Then
    Select Case cmbType(4).SelText
    Case "Away Message"
        pSourceUser.SendText bstrMsgHeader, "Away Message:" & txtAway, MMSGTYPE_NO_RESULT
    Case "Time"
        pSourceUser.SendText bstrMsgHeader, "The Time is " & Time, MMSGTYPE_NO_RESULT
    Case "Read Message Board"
        A = 0
        Do
        A = A + 1
        If lstMsgBoard.ListCount = A Then Exit Sub
        pSourceUser.SendText bstrMsgHeader, lstMsgBoard.List(A), MMSGTYPE_NO_RESULT
        Loop
    Case "Add to Message Board"
        msg = True
        pSourceUser.SendText bstrMsgHeader, "Please Type in your message now!", MMSGTYPE_NO_RESULT
    Case "Change Name & Status (Password needed)"
        pass = True
        pSourceUser.SendText bstrMsgHeader, "Please enter your password now!", MMSGTYPE_NO_RESULT
    Case "Version"
        pSourceUser.SendText bstrMsgHeader, "The Version of this program is 1.0!", MMSGTYPE_NO_RESULT
    End Select
End If
If bstrMsgText = txtkey(5).Text Then
    Select Case cmbType(5).SelText
    Case "Away Message"
        pSourceUser.SendText bstrMsgHeader, "Away Message:" & txtAway, MMSGTYPE_NO_RESULT
    Case "Time"
        pSourceUser.SendText bstrMsgHeader, "The Time is " & Time, MMSGTYPE_NO_RESULT
    Case "Read Message Board"
        A = 0
        Do
        A = A + 1
        If lstMsgBoard.ListCount = A Then Exit Sub
        pSourceUser.SendText bstrMsgHeader, lstMsgBoard.List(A), MMSGTYPE_NO_RESULT
        Loop
    Case "Add to Message Board"
        msg = True
        pSourceUser.SendText bstrMsgHeader, "Please Type in your message now!", MMSGTYPE_NO_RESULT
    Case "Change Name & Status (Password needed)"
        pass = True
        pSourceUser.SendText bstrMsgHeader, "Please enter your password now!", MMSGTYPE_NO_RESULT
    Case "Version"
        pSourceUser.SendText bstrMsgHeader, "The Version of this program is 1.1!", MMSGTYPE_NO_RESULT
    End Select
End If
Menu = "Welcome to the Ultimate Away Messager (UAM)" & vbCrLf
Menu = Menu & "Press " & txtkey(0) & " for " & txtmessage(0) & vbCrLf
Menu = Menu & "Press " & txtkey(1) & " for " & txtmessage(1) & vbCrLf
Menu = Menu & "Press " & txtkey(2) & " for " & txtmessage(2) & vbCrLf
Menu = Menu & "Press " & txtkey(3) & " for " & txtmessage(3) & vbCrLf
Menu = Menu & "Press " & txtkey(4) & " for " & txtmessage(4) & vbCrLf
Menu = Menu & "Press " & txtkey(5) & " for " & txtmessage(5) & vbCrLf
pSourceUser.SendText bstrMsgHeader, Menu, MMSGTYPE_NO_RESULT

End Sub

