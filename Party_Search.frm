VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form Party_Search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                        ----- Party Search -----"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7260
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Party_Search.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Party_Search.frx":043C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdClose 
      Height          =   405
      Left            =   5880
      TabIndex        =   2
      Top             =   4395
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      MICON           =   "Party_Search.frx":0840
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdAllSearch 
      Height          =   405
      Left            =   4320
      TabIndex        =   1
      Top             =   4395
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Search"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      MICON           =   "Party_Search.frx":085C
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "2"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Frame Frame1 
      Height          =   3405
      Left            =   120
      TabIndex        =   7
      Top             =   885
      Width           =   6930
      Begin MSComctlLib.ListView lvwphase 
         Height          =   3090
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   5450
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Detail 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6930
      Begin VB.CheckBox Is_Active 
         Caption         =   "Either Active or not?"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox TXTSearch 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1155
         TabIndex        =   0
         Top             =   240
         Width           =   5325
      End
      Begin VB.Label Lb_Search 
         Caption         =   "Search"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   315
         Width           =   720
      End
   End
End
Attribute VB_Name = "Party_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    If (Len(Trim(Me.TXTSearch.Text)) > 0) Then
        Set rstList = FillRecordSet("SELECT Party.PartyCode, Party.PartyName, Party.PartyAddress, Party.PartyPhone, Party.PartyCell, PartyType.PartyTypeName " & _
                " FROM PartyType INNER JOIN Party ON PartyType.PartyTypeCode = Party.PartyTypeCode where Party.IsActive = 1 and Party.PartyName + PartyType.PartyTypeName like '%" & Me.TXTSearch.Text & "%' order by Party.PartyCode desc")
    Else
        Set rstList = FillRecordSet("SELECT Party.PartyCode, Party.PartyName, Party.PartyAddress, Party.PartyPhone, Party.PartyCell, PartyType.PartyTypeName " & _
                " FROM PartyType INNER JOIN Party ON PartyType.PartyTypeCode = Party.PartyTypeCode where Party.IsActive = 1 order by Party.PartyCode desc")
    End If
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!PartyCode, _
                   Key:=CStr("Id=" & rstList!PartyCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!PartyName
                 .Add Text:=rstList!PartyTypeName
                 .Add Text:=rstList!PartyAddress
                 .Add Text:=rstList!PartyPhone
                 .Add Text:=rstList!PartyCell
                 .Add Text:=rstList!PartyCode
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Private Sub CmdAllSearch_Click()
    TXTSearch = ""
    Call fillList
End Sub
Private Sub cmdClose_Click()
    Unload Party_Search
End Sub

Private Sub Form_Load()
    DBConn
    lvwphase.ColumnHeaders.Add Text:="Party Code", Width:=1000
    lvwphase.ColumnHeaders.Add Text:="Party Name", Width:=1300
    lvwphase.ColumnHeaders.Add Text:="Type", Width:=1000
    lvwphase.ColumnHeaders.Add Text:="Address", Width:=1500
    lvwphase.ColumnHeaders.Add Text:="Phone", Width:=900
    lvwphase.ColumnHeaders.Add Text:="Cell", Width:=900
    Call fillList
End Sub

Private Sub txtsearch_KeyUp(KeyCode As Integer, Shift As Integer)
    Call fillList
End Sub
