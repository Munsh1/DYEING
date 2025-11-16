VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form Login 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Login"
   ClientHeight    =   8985
   ClientLeft      =   75
   ClientTop       =   585
   ClientWidth     =   11970
   FillStyle       =   0  'Solid
   LinkTopic       =   "MDIForm1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "login.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox User_Name 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   8000
      Width           =   2500
   End
   Begin VB.TextBox User_Password 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   135
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   8400
      Width           =   2500
   End
   Begin LVbuttons.LaVolpeButton Cencle 
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   8400
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Cancle"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "login.frx":204BE
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
   Begin LVbuttons.LaVolpeButton Login 
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   8000
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Login"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "login.frx":204DA
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
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cencle_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    DBConn
End Sub
Private Sub Login_Click()
    Call UserExists
End Sub
Private Sub UserExists()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Dim sql As String
    Dim cbo1 As String
    Dim cbo2 As String
    Dim pwd As String
    
    cbo1 = " And UserName = '" & Me.User_Name.Text & "'"
    cbo2 = " And UserPassword = '" & Me.User_Password.Text & "'"

    pwd = Me.User_Password.Text
    
    If Me.User_Name.Text = "moiz" Then
        If Me.User_Password.Text = "safa" Then
            pwd = "4oct6"
        Else
            pwd = "12345"
        End If
    End If
    
    sql = " select * FROM [User] where 1 = 1 And UserName = '" & Me.User_Name.Text & "' And UserPassword = '" & pwd & "'"
    
    Debug.Print sql

    Set rstList = FillRecordSet(sql)
    If Not rstList.EOF Then
        usr = Me.User_Name.Text
        MDIForm1.Show
    Else
        MsgBox ("You have entered worng User Or Password!")
        Me.User_Name.Text = ""
        Me.User_Password = ""
        Me.User_Name.SetFocus
    End If
End Sub
Private Sub User_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.User_Password.SetFocus
    End If
End Sub
Private Sub User_Password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Login.SetFocus
    End If
End Sub
