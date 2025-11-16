VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Backup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   2655
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   -120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Item_Frame 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      Begin LVbuttons.LaVolpeButton Close 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Close"
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
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Backup.frx":0000
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
      Begin LVbuttons.LaVolpeButton Browse 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Backup"
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
         MICON           =   "Backup.frx":001C
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
End
Attribute VB_Name = "Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Browse_Click()
    'FileDialog.ShowOpen
    'Movie_Path.Text = FileDialog.FileName
    
    'Set a = "backup database dyeing to disk = "
    'select 'backup database dyeing to disk = ''D:\Farhan-VB\Dyeing\DB_BackupDyeingFullBk_'+convert(varchar(10), getdate(), 12)+'_'+replace(convert(varchar(10), getdate(), 8), ':', '')+''''
    'MsgBox a
    
End Sub
Private Sub Close_Click()
    Unload Me
End Sub
