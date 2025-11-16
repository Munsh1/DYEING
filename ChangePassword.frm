VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form ChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4740
   Begin LVbuttons.LaVolpeButton Changed 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Change Password"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "ChangePassword.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.TextBox New_Password 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1140
      Width           =   3000
   End
   Begin VB.TextBox Old_Password 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   735
      Width           =   3000
   End
   Begin VB.TextBox User_Name 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   345
      Width           =   3000
   End
   Begin VB.Label Label3 
      Caption         =   "New Password"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1140
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Old Password"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   735
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   345
      Width           =   1095
   End
End
Attribute VB_Name = "ChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub User_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim(Me.User_Name.Text)) > 0 Then
        Me.Old_Password.SetFocus
    End If
End Sub
Private Sub Old_Password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim(Me.Old_Password.Text)) > 0 Then
        Me.New_Password.SetFocus
    End If
End Sub
Private Sub New_Password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim(Me.New_Password.Text)) > 0 Then
        Me.Changed.SetFocus
    End If
End Sub
Private Sub Changed_Click()
    DBConn
    Dim rstList  As New ADODB.Recordset
    Dim sql As String
    
    sql = " Select * FROM [User] where [UserId] = 0"
    rstList.Open sql, cnDatabase, 1, adLockReadOnly
    If rstList(1).value = Me.User_Name.Text And rstList(2).value = Me.Old_Password.Text Then
        sql = "Update [User] set UserPassword = '" & Me.New_Password.Text & "' where [UserId] = 0"
        cnDatabase.Execute sql
        Set rstList = Nothing
        MsgBox ("Password Changed!")
    End If
End Sub
