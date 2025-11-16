VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Delivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                    -----  Delivery -----"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12750
   FillStyle       =   3  'Vertical Line
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7666.072
   ScaleMode       =   0  'User
   ScaleWidth      =   12750
   Begin VB.Frame Frame3 
      Height          =   3855
      Left            =   120
      TabIndex        =   39
      Top             =   2640
      Width           =   9255
      Begin MSComctlLib.ListView lvwphase 
         Height          =   3480
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   6138
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
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detail Block"
      Height          =   1335
      Left            =   120
      TabIndex        =   34
      Top             =   1200
      Width           =   9255
      Begin VB.ComboBox f_DeliveryParty 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Text            =   "f_DeliveryParty"
         Top             =   840
         Width           =   7935
      End
      Begin VB.TextBox f_Cone 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6960
         TabIndex        =   6
         Top             =   435
         Width           =   855
      End
      Begin VB.TextBox Rates 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7980
         TabIndex        =   7
         Top             =   435
         Width           =   1095
      End
      Begin VB.TextBox Qty 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6000
         TabIndex        =   5
         Top             =   435
         Width           =   855
      End
      Begin VB.ComboBox Item 
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   58
         Top             =   435
         Width           =   3015
      End
      Begin VB.ComboBox Item_Type 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   57
         Top             =   435
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Delivery Party"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   870
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Cones"
         Height          =   255
         Left            =   7080
         TabIndex        =   45
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Unit-Price"
         Height          =   375
         Left            =   8160
         TabIndex        =   38
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   6120
         TabIndex        =   37
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Item"
         Height          =   255
         Left            =   4080
         TabIndex        =   36
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Item Type"
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   165
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Master Block"
      Height          =   1095
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   9255
      Begin VB.ComboBox f_ProcessCode 
         Height          =   315
         Left            =   4380
         TabIndex        =   3
         Top             =   600
         Width           =   1665
      End
      Begin VB.TextBox f_Color 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6180
         TabIndex        =   56
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Challan 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7920
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox Party 
         Height          =   315
         Left            =   1350
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DeliveryDate 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16449539
         CurrentDate     =   38216
      End
      Begin VB.Label Label11 
         Caption         =   "Process Code"
         Height          =   255
         Left            =   4500
         TabIndex        =   55
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Color"
         Height          =   255
         Left            =   6600
         TabIndex        =   41
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Challan No."
         Height          =   255
         Left            =   7965
         TabIndex        =   33
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Party"
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   9480
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.Frame Frame10 
         Height          =   700
         Left            =   120
         TabIndex        =   59
         Top             =   5800
         Width           =   2895
         Begin VB.TextBox srSER2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   62
            Top             =   320
            Width           =   1215
         End
         Begin VB.TextBox srSER1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   61
            Top             =   320
            Width           =   1215
         End
         Begin VB.CheckBox SERChk 
            Caption         =   "Serial"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame Frame19 
         Height          =   700
         Left            =   120
         TabIndex        =   51
         Top             =   5100
         Width           =   2895
         Begin VB.TextBox srPC1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   54
            Top             =   320
            Width           =   1215
         End
         Begin VB.TextBox srPC2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   53
            Top             =   320
            Width           =   1215
         End
         Begin VB.CheckBox PCChk 
            Caption         =   "PC Code"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame Frame9 
         Height          =   700
         Left            =   120
         TabIndex        =   47
         Top             =   4400
         Width           =   2895
         Begin VB.TextBox SrChallan2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   50
            Top             =   320
            Width           =   1215
         End
         Begin VB.TextBox SrChallan 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   49
            Top             =   320
            Width           =   1215
         End
         Begin VB.CheckBox ChChk 
            Caption         =   "Challan"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame Frame18 
         Height          =   735
         Left            =   120
         TabIndex        =   42
         Top             =   3660
         Width           =   2895
         Begin VB.CheckBox ClChk 
            Caption         =   "Color"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox SrColor 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   43
            Top             =   320
            Width           =   2655
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1095
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2895
         Begin MSComCtl2.DTPicker SrDate2 
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   680
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   16449537
            CurrentDate     =   38313
         End
         Begin VB.CheckBox dtChk 
            Caption         =   "Date"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   0
            Width           =   735
         End
         Begin MSComCtl2.DTPicker SrDate 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   280
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   16449537
            CurrentDate     =   38224
         End
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   1380
         Width           =   2895
         Begin VB.CheckBox PtChk 
            Caption         =   "Party"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   0
            Width           =   735
         End
         Begin VB.ComboBox SrParty 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   20
            Text            =   "SrParty"
            Top             =   280
            Width           =   2655
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   2150
         Width           =   2895
         Begin VB.CheckBox ImTChk 
            Caption         =   "Item Type"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   1095
         End
         Begin VB.ComboBox SrItemType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Text            =   "SrItemType"
            Top             =   280
            Width           =   2655
         End
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   120
         TabIndex        =   26
         Top             =   2900
         Width           =   2895
         Begin VB.CheckBox ImChk 
            Caption         =   "Item"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   0
            Width           =   735
         End
         Begin VB.ComboBox SrItem 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Text            =   "SrItem"
            Top             =   280
            Width           =   2655
         End
      End
      Begin LVbuttons.LaVolpeButton Cmdhide 
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   6645
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Hide"
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
         COLTYPE         =   2
         BCOL            =   12632256
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Delivery.frx":0000
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "7"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Delivery.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Delivery.frx":0284
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Delivery.frx":06DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Delivery.frx":0AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Delivery.frx":0F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Delivery.frx":134C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Delivery.frx":1788
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton CmdAllSearch 
      Height          =   405
      Left            =   6975
      TabIndex        =   13
      Top             =   6600
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
      MICON           =   "Delivery.frx":181C
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "5"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdClose 
      Height          =   405
      Left            =   8235
      TabIndex        =   14
      Top             =   6600
      Width           =   1095
      _ExtentX        =   1931
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
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Delivery.frx":1838
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "6"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdDel 
      Height          =   405
      Left            =   5715
      TabIndex        =   12
      Top             =   6600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Delete"
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
      MICON           =   "Delivery.frx":1854
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "4"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdNew 
      Height          =   405
      Left            =   4680
      TabIndex        =   11
      Top             =   6600
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Add"
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
      MICON           =   "Delivery.frx":1870
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "3"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdSave 
      Height          =   405
      Left            =   3615
      TabIndex        =   9
      Top             =   6600
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Save"
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
      MICON           =   "Delivery.frx":188C
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
   Begin LVbuttons.LaVolpeButton CMDdetail 
      Height          =   405
      Left            =   2475
      TabIndex        =   10
      Top             =   6600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Detail"
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
      MICON           =   "Delivery.frx":18A8
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
   Begin VB.Label lblCaption 
      Caption         =   "Label12"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   40
      Top             =   6720
      Width           =   1215
   End
End
Attribute VB_Name = "Delivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_ListID As String
Dim d_ListID As String
Dim PreQty As Double
Private Sub Challan_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Qty.SetFocus
    End If
End Sub
Private Sub ChChk_Click()
    If ChChk.value = Checked Then
        Me.SrChallan.Enabled = True
    Else
        Me.SrChallan.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub ClChk_Click()
    If ClChk.value = Checked Then
        Me.SrColor.Enabled = True
    Else
        Me.SrColor.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub CmdAllSearch_Click()
        Delivery.Left = 200
        Delivery.Top = 500
        Delivery.Width = 12840
        
        Call SrfillList
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub CmdDel_Click()
     If (Len(Trim(m_ListID)) > 0 And Len(Trim(d_ListID)) > 0) Then
        Dim strAns As String
        Dim AvbQty As Integer
        Dim rstGetQty As New ADODB.Recordset
        
        strAns = MsgBox("Do you want to delete this record...?", vbYesNo + vbInformation)
        If strAns = vbYes Then
            cnDatabase.Execute "update DeliveryDetail set IsActive = 0 where DeliveryDetailCode =" & d_ListID
            Call fillList
            MsgBox ("Record deleted succesfully..."), vbInformation
            Me.cmdSave.Enabled = False
            Call addNewMaster
        End If
        m_ListID = ""
        m_AddMode = True
        Me.CmdNew.SetFocus
        End If
End Sub
Private Sub CMDdetail_Click()
        isdetail = True
        Call addNewDetail
        Me.Item_Type.SetFocus
End Sub
Private Sub Cmdhide_Click()
        Delivery.Width = 9540
        Delivery.Left = 2000
        Delivery.Top = 500
        Me.SrItem.ListIndex = -1
        Me.SrItemType.ListIndex = -1
        Me.SrParty.ListIndex = -1
        Call fillList
End Sub
Private Sub CmdNew_Click()
    Call addNewMaster
    Me.Party.SetFocus
End Sub
Private Sub cmdSave_Click()
If Len(Trim(Party)) > 0 And Len(Trim(Challan)) > 0 And Len(Trim(Item_Type)) > 0 And Len(Trim(Item)) > 0 And Len(Trim(Qty)) > 0 And Len(Trim(Rates)) > 0 Then
            Call setValMaster
            Call setValDetail
            Call fillList
            Call addNewMaster
            Call addNewDetail
            Me.Party.SetFocus
            MsgBox ("Record saved successfully"), vbInformation
Else
    MsgBox "Provide data in all Fields"
End If
End Sub
Private Sub f_Color_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Challan.SetFocus
    End If
End Sub
Private Sub f_Cone_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Me.Rates.SetFocus
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
       KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_DeliveryParty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        If Len(Trim(Challan)) > 0 And Len(Trim(Qty)) > 0 Then
            Me.cmdSave.SetFocus
        End If
    End If
End Sub
Private Sub f_ProcessCode_Click()
    
    If (Me.f_ProcessCode.ListIndex > -1) Then
        i = Me.f_ProcessCode.ItemData(Me.f_ProcessCode.ListIndex)
        FillCombo "Select ItemTypeCode, ItemTypeName from vwRpt_Delivery where ProcessCode = " & i, Item_Type, "ItemTypeName", "ItemTypeCode"
        FillCombo "Select ItemCode, ItemName from vwRpt_Delivery where ProcessCode = " & i, Item, "ItemName", "ItemCode"
        Dim rstGetVal As New ADODB.Recordset
        Set rstGetVal = FillRecordSet("select NewColor, ItemTypeCode, ItemCode, Quantity, Cone from vwRpt_Delivery where ProcessCode =" & Me.f_ProcessCode.ItemData(Me.f_ProcessCode.ListIndex) & " and PartyCode = " & Me.Party.ItemData(Me.Party.ListIndex))
        If Not (rstGetVal.EOF) Then
            Me.f_Color.Text = rstGetVal("NewColor")
            Call selectValueInCombo(Me.Item_Type, rstGetVal("ItemTypeCode"))
            Call selectValueInCombo(Me.Item, rstGetVal("ItemCode"))
            Me.Qty.Text = rstGetVal("Quantity")
            Me.f_Cone.Text = rstGetVal("Cone")
        End If
        rstGetVal.Close
        Set rstGetVal = Nothing
        '        i = Me.Item_Type.ItemData(Me.Item_Type.ListIndex)
        '        a = Me.Party.ItemData(Me.Party.ListIndex)
        '        B = Me.f_ProcessCode.ItemData(Me.f_ProcessCode.ListIndex)
        '        FillCombo "Select Distinct i.ItemCode, i.ItemName from Item i inner join vwRpt_Delivery v on i.ItemCode = v.ItemCode where i.ItemTypeCode = " & i & " and PartyCode = " & a & " and ProcessCode = " & B & " order by 2", Item, "ItemName", "ItemCode"
        '    Else
        '        Me.Item.Clear
    End If
End Sub
Private Sub f_ProcessCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Challan.SetFocus
    End If
End Sub

Private Sub Form_Load()
    isdetail = True
    m_AddMode = True
    DBConn
    DeliveryDate = Date
    SrDate = Date
    SrDate2 = Date
    Call addNewMaster
   ' FillCombo "Select Distinct t.ItemTypeCode, ItemTypeName from ItemType t inner join vwDelivery v on v.ItemTypeCode = t.ItemTypeCode where IsActive = 1 order by 2", Item_Type, "ItemTypeName", "ItemTypeCode"
    FillCombo "Select Distinct p.PartyCode, PartyName from Party p inner join vwDelivery v on p.PartyCode = v.PartyCode where IsActive = 1 order by 2", Party, "PartyName", "PartyCode"
    FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", f_DeliveryParty, "PartyName", "PartyCode"
    
    FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", SrItemType, "ItemTypeName", "ItemTypeCode"
    FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", SrParty, "PartyName", "PartyCode"

    lvwphase.ColumnHeaders.Add Text:="Detail Code", Width:=0
    lvwphase.ColumnHeaders.Add Text:="Master Code", Width:=0
    lvwphase.ColumnHeaders.Add Text:="PC #", Width:=800
    lvwphase.ColumnHeaders.Add Text:="Serial #", Width:=800
    lvwphase.ColumnHeaders.Add Text:="Party Name", Width:=1800
    'lvwphase.ColumnHeaders.Add Text:="PC #", Width:=800
    lvwphase.ColumnHeaders.Add Text:="Color", Width:=1250
    lvwphase.ColumnHeaders.Add Text:="Challan", Width:=800
    lvwphase.ColumnHeaders.Add Text:="Item Name", Width:=1800
    lvwphase.ColumnHeaders.Add Text:="Quantity", Width:=800
    lvwphase.ColumnHeaders.Add Text:="Cone", Width:=700
'   lvwphase.ColumnHeaders.Add Text:="Rates", Width:=800
    
    Call fillList
    lblCaption.Caption = "Add Master"
End Sub

Private Sub Item_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Qty.SetFocus
    End If
End Sub
'Private Sub Item_Type_Click()
'    If (Me.Item_Type.ListIndex > -1) Then
'        i = Me.Item_Type.ItemData(Me.Item_Type.ListIndex)
'        a = Me.Party.ItemData(Me.Party.ListIndex)
'        FillCombo "Select Distinct ProcessCode, ProcessCode from vwRpt_Delivery where ItemTypeCode = " & i & " and PartyCode = " & a, f_ProcessCode, "ProcessCode", "ProcessCode"
'    Else
'        Me.f_ProcessCode.Clear
'        Me.Item.Clear
'    End If
'End Sub
Public Sub addNewDetail()
    d_ListID = ""
    Me.Item_Type.ListIndex = -1
    Me.Item.ListIndex = -1
    Me.Qty.Text = ""
    Me.Rates.Text = ""
    Me.f_Cone.Text = ""
    Me.f_DeliveryParty.ListIndex = -1
    lblCaption.Caption = "Add Detail"
End Sub
Private Sub getValMaster()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("Select DeliveryMasterCode, DeliveryMasterDate, PartyCode, ChallanCode From DeliveryMaster Where DeliveryMasterCode = " & m_ListID)
    If Not (rstGetVal.EOF) Then
        Me.DeliveryDate.value = IIf(IsNull(rstGetVal("DeliveryMasterDate")), Now, rstGetVal("DeliveryMasterDate"))
        Call selectValueInCombo(Me.Party, rstGetVal("PartyCode"))
        Me.Challan.Text = rstGetVal("ChallanCode")
        Me.f_Color.Text = rstGetVal("Color")
   End If
   rstGetVal.Close
   Set rstGetVal = Nothing
End Sub
Private Sub getValDetail()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("Select ItemTypeCode, ItemCode, Quantity, Rate From DeliveryDetail Where DeliveryDetailCode = " & d_ListID)
    If Not (rstGetVal.EOF) Then
    Debug.Print rstGetVal("ItemTypeCode")
        Call selectValueInCombo(Me.f_ProcessCode, rstGetVal("ProcessCode"))
        Call selectValueInCombo(Me.Item_Type, rstGetVal("ItemTypeCode"))
        Call selectValueInCombo(Me.Item, rstGetVal("ItemCode"))
        Me.Qty.Text = rstGetVal("Quantity")
        Me.f_Cone.Text = IIf(IsNull(rstGetVal("Cone")), 0, rstGetVal("Cone"))
        Me.Rates.Text = rstGetVal("Rate")
        Call selectValueInCombo(Me.f_DeliveryParty, rstGetVal("DeliveryParty"))
   End If
   rstGetVal.Close
   Set rstGetVal = Nothing
End Sub
Public Sub setValMaster()
    Dim rstSave As New ADODB.Recordset
    If (Len(Trim(m_ListID)) = 0) Then
        Set rstSave = FillRecordSet("select * from DeliveryMaster Where 1 = 2")
        rstSave.AddNew
        m_ListID = ValAutoNumber("DeliveryMaster", "DeliveryMasterCode")
        rstSave("DeliveryMasterCode") = m_ListID
    Else
       Set rstSave = FillRecordSet("select * from DeliveryMaster where DeliveryMasterCode =" & m_ListID)
    End If
    
    rstSave("DeliveryMasterDate") = Me.DeliveryDate.value
    rstSave("PartyCode") = Me.Party.ItemData(Party.ListIndex)
    rstSave("ChallanCode") = Me.Challan.Text
    rstSave("Color") = IIf(IsNull(Me.f_Color.Text), "", Me.f_Color.Text)
    rstSave("IsActive") = "1"
    
    rstSave.Update
    rstSave.Close
    Set rstSave = Nothing
End Sub
Public Sub setValDetail()
    Dim rstSave As New ADODB.Recordset
    If (Len(Trim(d_ListID)) = 0) Then
        Set rstSave = FillRecordSet("select * from DeliveryDetail Where 1 = 2")
        rstSave.AddNew
        d_ListID = ValAutoNumber("DeliveryDetail", "DeliveryDetailCode")
        rstSave("DeliveryDetailCode") = d_ListID
    Else
       Set rstSave = FillRecordSet("select * from DeliveryDetail where DeliveryDetailCode =" & d_ListID)
    End If
    
    rstSave("DeliveryMasterCode") = m_ListID
    rstSave("ProcessCode") = f_ProcessCode.ItemData(f_ProcessCode.ListIndex)
    rstSave("ItemTypeCode") = Item_Type.ItemData(Item_Type.ListIndex)
    rstSave("ItemCode") = Me.Item.ItemData(Item.ListIndex)
    rstSave("Quantity") = IIf(IsNull(Me.Qty.Text), 0, Me.Qty.Text)
    rstSave("Cone") = IIf(IsNull(Me.f_Cone.Text), 0, Me.f_Cone.Text)
    rstSave("Rate") = Me.Rates.Text
    If (Me.f_DeliveryParty.ListIndex > -1) Then
        rstSave("DeliveryParty") = Me.f_DeliveryParty.ItemData(f_DeliveryParty.ListIndex)
    End If
    rstSave("IsActive") = "1"
    
    rstSave.Update
    rstSave.Close
    Set rstSave = Nothing
    Call addNewDetail
End Sub
Public Sub addNewMaster()
    m_ListID = ""
    Me.DeliveryDate.value = Now
    Me.Party.ListIndex = -1
    Me.f_Color.Text = ""
    Me.Challan.Text = ""
    
    d_ListID = ""
    Me.Item_Type.ListIndex = -1
    Me.Item.ListIndex = -1
    Me.f_ProcessCode.ListIndex = -1
    Me.Qty.Text = ""
    Me.f_Cone.Text = ""
    Me.Rates.Text = ""
    
    lblCaption.Caption = "Add Master"
End Sub
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Set rstList = FillRecordSet("SELECT top 100 ProcessCode, (select SerialNo from Process where ProcessCode = DeliveryDetail.ProcessCode) as SerialNo, DeliveryDetailCode, ChallanCode, Color, DeliveryMaster.DeliveryMasterCode, PartyName, DeliveryMaster.PartyCode, DeliveryDetail.ItemTypeCode, ItemTypeName, DeliveryDetail.ItemCode, ItemName, Quantity, DeliveryDetail.Cone, Rate " & _
                                "FROM (Party INNER JOIN (DeliveryMaster INNER JOIN DeliveryDetail ON DeliveryMaster.DeliveryMasterCode = DeliveryDetail.DeliveryMasterCode) ON Party.PartyCode = DeliveryMaster.PartyCode) INNER JOIN (ItemType INNER JOIN Item ON ItemType.ItemTypeCode = Item.ItemTypeCode) ON (DeliveryDetail.ItemCode = Item.ItemCode) AND (DeliveryDetail.ItemTypeCode = ItemType.ItemTypeCode) " & _
                                "Where DeliveryDetail.IsActive = 1 order by DeliveryMaster.DeliveryMasterCode desc, DeliveryDetailCode desc")
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!DeliveryDetailCode, _
                   Key:=CStr("Id=" & rstList!DeliveryDetailCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!DeliveryMasterCode
                 .Add Text:=rstList!ProcessCode
                 .Add Text:=rstList!SerialNo
                 .Add Text:=rstList!PartyName
'                 .Add Text:=rstList!ProcessCode
                 .Add Text:=rstList!Color
                 .Add Text:=rstList!ChallanCode
                 .Add Text:=rstList!ItemName
                 .Add Text:=rstList!Quantity
                 .Add Text:=rstList!Cone
'                 .Add Text:=rstList!Rate
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Private Sub Item_Type_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Item.SetFocus
    End If
End Sub

Private Sub lvwphase_Click()
'    isdetail = False
'    cmdSave.Enabled = True
'    CmdDel.Enabled = True
'    m_AddMode = False
    d_ListID = Me.lvwphase.SelectedItem.Text
    m_ListID = Me.lvwphase.ListItems.Item(Me.lvwphase.SelectedItem.Index).ListSubItems(1).Text
'
'    Call getValMaster
'    Call getValDetail
End Sub
Private Sub lvwphase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        isdetail = False
'        cmdSave.Enabled = True
'        CmdDel.Enabled = True
'        m_AddMode = False
        d_ListID = Me.lvwphase.SelectedItem.Text
        m_ListID = Me.lvwphase.ListItems.Item(Me.lvwphase.SelectedItem.Index).ListSubItems(1).Text
'
'        Call getValMaster
'        Call getValDetail
    End If
End Sub
Private Sub Party_Click()
    If (Me.Party.ListIndex > -1) Then
        i = Me.Party.ItemData(Me.Party.ListIndex)
   '     FillCombo "Select Distinct i.ItemTypeCode, i.ItemTypeName from ItemType i inner join vwDelivery v on i.ItemTypeCode = v.ItemTypeCode where IsActive = 1 and v.PartyCode = " & i & " order by 2", Item_Type, "ItemTypeName", "ItemTypeCode"
        FillCombo "Select Distinct ProcessCode, ProcessCode from vwRpt_Delivery v where v.PartyCode = " & i & " order by 2", f_ProcessCode, "ProcessCode", "ProcessCode"
        Me.Item_Type.Clear
        Me.Item.Clear
    End If
End Sub

Private Sub Party_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Me.f_ProcessCode.SetFocus
    End If
End Sub

Private Sub PCChk_Click()
    If PCChk.value = Checked Then
        Me.srPC1.Enabled = True
        Me.srPC2.Enabled = True
    Else
        Me.srPC1.Enabled = False
        Me.srPC2.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub SERChk_Click()
    If SERChk.value = Checked Then
        Me.srSER1.Enabled = True
        Me.srSER2.Enabled = True
    Else
        Me.srSER1.Enabled = False
        Me.srSER2.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub Qty_LostFocus()
    If (Item_Type.ListIndex <> -1 And Item.ListIndex <> -1) Then
        Dim AvbQty As Double
        Dim rstGetQty As New ADODB.Recordset
        Set rstGetQty = FillRecordSet("Select Qty from vwAvailableQty where ItemCode = " & Me.Item.ItemData(Item.ListIndex) & " and ItemTypeCode = " & Me.Item_Type.ItemData(Item_Type.ListIndex))
        AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Qty"))) Then
                AvbQty = CDbl(rstGetQty("Qty"))
            End If
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        If (Len(Trim(Me.Qty.Text)) > 0) Then
            If (IsNull(Me.Qty.Text)) Then
                MsgBox "Quantity must be greater then zero"
                Me.cmdSave.Enabled = False
            ElseIf (IIf(m_AddMode = False, (CLng(AvbQty) + CLng(PreQty - Me.Qty.Text)), CLng(AvbQty)) < IIf(m_AddMode = False, Abs(CLng(Me.Qty.Text) - PreQty), CLng(Me.Qty.Text))) Then
                MsgBox "Quantity not Available !"
                Me.cmdSave.Enabled = False
            ElseIf CLng(Me.Qty.Text) = 0 Then
                MsgBox "Quantity must be greater then zero"
                Me.cmdSave.Enabled = False
            Else
                Me.cmdSave.Enabled = True
                Me.f_Cone.SetFocus
            End If
        End If
    End If
End Sub
Private Sub Qty_GotFocus()
If Len(Trim(Qty)) > 0 Then
    PreQty = Me.Qty.Text
Else
    PreQty = 0
End If
End Sub
Private Sub Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         f_Cone.SetFocus
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub Rates_KeyPress(KeyAscii As Integer)
    Dim numVar As Integer
    Call EnableSave
    If KeyAscii = 13 Then
       Me.f_DeliveryParty.SetFocus
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub EnableSave()
    If Len(Trim(Challan)) > 0 And Len(Trim(Qty)) > 0 Then
        Me.cmdSave.Enabled = True
        Me.CmdDel.Enabled = True
    Else
        Me.cmdSave.Enabled = False
        Me.CmdDel.Enabled = False
    End If
End Sub

Private Sub SrChallan2_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub SrChallan_keyup(KeyCode As Integer, Shift As Integer)
    SrChallan2 = SrChallan
    Call SrfillList
End Sub

Private Sub SrColor_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub

Private Sub SrItemType_Click()
    If Me.SrItemType.ListIndex > -1 Then
        i = Me.SrItemType.ItemData(Me.SrItemType.ListIndex)
        FillCombo "Select Distinct i.ItemCode, i.ItemName from Item i inner join vwAvailableQty v on i.ItemCode = v.ItemCode where IsActive = 1 and i.ItemTypeCode = " & i, SrItem, "ItemName", "ItemCode"
    Else
        Me.SrItem.Clear
    End If
    Call SrfillList
End Sub
Private Sub SrfillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Dim sql As String
    Dim cbo1 As String
    Dim cbo2 As String
    Dim cbo3 As String
    Dim cbo4 As String
    Dim cbo5 As String
    Dim cbo6 As String
    Dim cbo7 As String
    
    If dtChk.value = Checked Then
        'srdt = " And (DeliveryMasterDate between #" & Me.SrDate.value & " #  and #" & Me.SrDate2.value + 1 & " #)"
        srdt = " And (DeliveryMasterDate between Convert(datetime, '" & Me.SrDate.value - 1 & "')  and Convert(datetime, '" & Me.SrDate2.value + 1 & "'))"
    Else
        srdt = ""
    End If
    
    If PtChk.value = Checked And Me.SrParty.ListIndex > -1 Then
        cbo1 = " And DeliveryMaster.partycode = " & Me.SrParty.ItemData(Me.SrParty.ListIndex)
    Else
        cbo1 = ""
    End If
    
    If ImTChk.value = Checked And Me.SrItemType.ListIndex > -1 Then
        cbo2 = " And DeliveryDetail.ItemTypeCode = " & Me.SrItemType.ItemData(Me.SrItemType.ListIndex)
    Else
        cbo2 = ""
    End If
    
    If ImChk.value = Checked And Me.SrItem.ListIndex > -1 Then
        cbo3 = " And DeliveryDetail.ItemCode = " & Me.SrItem.ItemData(Me.SrItem.ListIndex)
    Else
        cbo3 = ""
    End If
    
    If ClChk.value = Checked And Len(Trim(Me.SrColor)) > 0 Then
        cbo4 = " And DeliveryMaster.Color like '%" & Me.SrColor & "%'"
    Else
        cbo4 = ""
    End If
    
    If ChChk.value = Checked And Len(Trim(Me.SrChallan)) > 0 And Len(Trim(Me.SrChallan2)) > 0 Then
        cbo5 = " And (DeliveryMaster.ChallanCode between '" & Me.SrChallan & "' and '" & Me.SrChallan2 & "')"
    Else
        cbo5 = ""
    End If
    
    If PCChk.value = Checked And Len(Trim(Me.srPC1)) > 0 And Len(Trim(Me.srPC2)) > 0 Then
        cbo6 = " And (ProcessCode between " & Me.srPC1 & " and " & Me.srPC2 & " )"
    Else
        cbo6 = ""
    End If
    
    If SERChk.value = Checked And Len(Trim(Me.srSER1)) > 0 And Len(Trim(Me.srSER2)) > 0 Then
        cbo7 = " And (ProcessCode in (Select ProcessCode from Process where SerialNo between '" & Me.srSER1 & "' and '" & Me.srSER2 & "' ))"
    Else
        cbo7 = ""
    End If
    
    sql = " SELECT top 100 ProcessCode, (select SerialNo from Process where ProcessCode = DeliveryDetail.ProcessCode) as SerialNo, DeliveryDetailCode, ChallanCode, Color, DeliveryMaster.DeliveryMasterCode, PartyName, DeliveryMaster.PartyCode, DeliveryDetail.ItemTypeCode, ItemTypeName, DeliveryDetail.ItemCode, ItemName, Quantity, Cone, Rate " & _
          " FROM (Party INNER JOIN (DeliveryMaster INNER JOIN DeliveryDetail ON DeliveryMaster.DeliveryMasterCode = DeliveryDetail.DeliveryMasterCode) ON Party.PartyCode = DeliveryMaster.PartyCode) INNER JOIN (ItemType INNER JOIN Item ON ItemType.ItemTypeCode = Item.ItemTypeCode) ON (DeliveryDetail.ItemCode = Item.ItemCode) AND (DeliveryDetail.ItemTypeCode = ItemType.ItemTypeCode) " & _
          " Where DeliveryDetail.IsActive = 1" & _
          srdt & _
          cbo1 & _
          cbo2 & _
          cbo3 & _
          cbo4 & _
          cbo5 & _
          cbo6 & _
          cbo7 & _
          " Order by DeliveryMaster.DeliveryMasterCode desc, DeliveryDetailCode desc"
                                
    Debug.Print sql
    Set rstList = FillRecordSet(sql)
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!DeliveryDetailCode, _
                   Key:=CStr("Id=" & rstList!DeliveryDetailCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!DeliveryMasterCode
                 .Add Text:=rstList!ProcessCode
                 .Add Text:=rstList!SerialNo
                 .Add Text:=rstList!PartyName
'                 .Add Text:=rstList!ProcessCode
                 .Add Text:=rstList!Color
                 .Add Text:=rstList!ChallanCode
                 .Add Text:=rstList!ItemName
                 .Add Text:=rstList!Quantity
                 .Add Text:=rstList!Cone
'                 .Add Text:=rstList!Rate
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Private Sub SrParty_Click()
    Call SrfillList
End Sub
Private Sub SrItem_Click()
    Call SrfillList
End Sub
Private Sub SrDate_Change()
    Call SrfillList
End Sub
Private Sub PtChk_Click()
    If PtChk.value = Checked Then
        Me.SrParty.Enabled = True
    Else
        Me.SrParty.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub ImChk_Click()
    If ImChk.value = Checked Then
        Me.SrItem.Enabled = True
    Else
        Me.SrItem.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub ImTChk_Click()
    If ImTChk.value = Checked Then
        Me.SrItemType.Enabled = True
    Else
        Me.SrItemType.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub dtChk_Click()
    If dtChk.value = Checked Then
        Me.SrDate.Enabled = True
        Me.SrDate2.Enabled = True
    Else
        Me.SrDate.Enabled = False
        Me.SrDate2.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub SrPC1_keyup(KeyCode As Integer, Shift As Integer)
    srPC2 = srPC1
    Call SrfillList
End Sub
Private Sub SrPC2_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub SrSER1_keyup(KeyCode As Integer, Shift As Integer)
    srSER2 = srSER1
    Call SrfillList
End Sub
Private Sub SrSER2_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
