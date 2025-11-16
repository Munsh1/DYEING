VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Dyeing System"
   ClientHeight    =   5265
   ClientLeft      =   555
   ClientTop       =   735
   ClientWidth     =   6435
   LinkTopic       =   "MDIForm1"
   Picture         =   "Dyeing.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dyeing.frx":15012
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dyeing.frx":15466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dyeing.frx":158BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dyeing.frx":15BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dyeing.frx":1602A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dyeing.frx":1647E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dyeing.frx":168D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dyeing.frx":16D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dyeing.frx":1717A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   1429
      ButtonWidth     =   1826
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Party"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search Party"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Item"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search Item"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "&Transaction"
      Begin VB.Menu SideMenuTrans 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Transactions|Font:Arial|BOLD|Fsize:10|Fcolor:16778215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnuReceive 
         Caption         =   "{IMG:6}Receive"
      End
      Begin VB.Menu MnuProcess 
         Caption         =   "{IMG:6}Polister Form"
      End
      Begin VB.Menu MnuSale 
         Caption         =   "{IMG:6}Sale"
      End
      Begin VB.Menu MnuDelivery 
         Caption         =   "{IMG:6}Delivery"
      End
      Begin VB.Menu mnuReceiveOrder 
         Caption         =   "{IMG:6}ReceiveOrder"
      End
      Begin VB.Menu MenuSep 
         Caption         =   "-"
      End
      Begin VB.Menu BackupMenu 
         Caption         =   "{IMG:6}Backup"
      End
      Begin VB.Menu menuDyeing 
         Caption         =   "{IMG:6}Cotton Dyeing"
      End
      Begin VB.Menu MenuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReceiveReturn 
         Caption         =   "{IMG:6}Receive Return"
      End
      Begin VB.Menu mnuRecipe 
         Caption         =   "{IMG:6}Recipe"
      End
      Begin VB.Menu mnuoldprocess 
         Caption         =   "{IMG:6}Old Process"
      End
      Begin VB.Menu MenuSep30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "{IMG:6}Change Password"
      End
   End
   Begin VB.Menu MnuSetup 
      Caption         =   "&Party"
      Begin VB.Menu SideMenuParty 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Party|Font:Arial|BOLD|Fsize:10|Fcolor:16778215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu MnuAddParty 
         Caption         =   "{IMG:7}Add Party"
      End
      Begin VB.Menu MnuPartySearch 
         Caption         =   "{IMG:5}Party Search"
      End
      Begin VB.Menu MnuAddPartyType 
         Caption         =   "{IMG:7}Add Party Type"
      End
   End
   Begin VB.Menu MnueItem 
      Caption         =   "&Item"
      Begin VB.Menu SideMenuItem 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Item|Font:Arial|BOLD|Fsize:10|Fcolor:16778215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnuAddItem 
         Caption         =   "{IMG:8}Add Item"
      End
      Begin VB.Menu MnuItemSearch 
         Caption         =   "{IMG:5}Item Search"
      End
      Begin VB.Menu mnuAddItemType 
         Caption         =   "{IMG:8}Add Item Type"
      End
   End
   Begin VB.Menu mnuRpt 
      Caption         =   "&Reports"
      Begin VB.Menu SideMenuRep 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Reports|Font:Arial|BOLD|Fsize:10|Fcolor:16778215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu MenRepAvailableProducts 
         Caption         =   "{IMG:9}Available Products"
      End
      Begin VB.Menu MnuRepProcess 
         Caption         =   "{IMG:9}Polister"
      End
      Begin VB.Menu MnuPolisterCosting 
         Caption         =   "{IMG:9}Polister Form Costing"
      End
      Begin VB.Menu MenSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParty 
         Caption         =   "{IMG:9}All Party"
      End
      Begin VB.Menu MnuItem 
         Caption         =   "{IMG:9}All Item"
      End
      Begin VB.Menu MenSep2 
         Caption         =   "-"
      End
      Begin VB.Menu muitemAct 
         Caption         =   "{IMG:9}Item Activity [ Date wise ]"
      End
      Begin VB.Menu mnuDptAct 
         Caption         =   "{IMG:9}Item Activity [ Party wise ]"
      End
      Begin VB.Menu mnuReceiveQuantity 
         Caption         =   "{IMG:9}Receive Quantity"
      End
      Begin VB.Menu MenSep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDeliv 
         Caption         =   "{IMG:9}Pending"
      End
      Begin VB.Menu MnuAllPending 
         Caption         =   "{IMG:9}All Pending"
      End
      Begin VB.Menu MenSep4 
         Caption         =   "-"
      End
      Begin VB.Menu DeliveryMnu 
         Caption         =   "{IMG:9}Delivery"
         Begin VB.Menu MnuDeliveryParty 
            Caption         =   "{IMG:9}Party Delivery"
         End
         Begin VB.Menu MnuItemTypeDelivery 
            Caption         =   "{IMG:9}Item Type Delivery"
         End
         Begin VB.Menu MnuItemDelivery 
            Caption         =   "{IMG:9}Item Delivery"
         End
      End
      Begin VB.Menu MnuProduction 
         Caption         =   "{IMG:9}Production"
         Begin VB.Menu MnuAllProduction 
            Caption         =   "{IMG:9}All Production"
         End
         Begin VB.Menu MnuTypeProduction 
            Caption         =   "{IMG:9}Item Type Production"
         End
         Begin VB.Menu MnuItemProduction 
            Caption         =   "{IMG:9}Item Production"
         End
         Begin VB.Menu MnuMachineProduction 
            Caption         =   "{IMG:9}Machine Production"
         End
      End
   End
   Begin VB.Menu MnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    RepReceiveDailyActivity.Show
End Sub
Private Sub HalfBleachMenu_Click()
        HalfBleach.Show
        HalfBleach.Height = 8300
        HalfBleach.Width = 8000
        HalfBleach.Left = 1700
End Sub
Private Sub BackupMenu_Click()
    Backup.Show
    Backup.Left = 3500
    Backup.Top = 2700
End Sub
Private Sub MDIForm_Load()
   SetMenus hwnd, ImageList1
   Unload Login
End Sub
Private Sub MenRepAvailableProducts_Click()
    RepAvailableProducts.Show
    RepAvailableProducts.Left = 3500
    RepAvailableProducts.Top = 2700

End Sub
Private Sub menuDyeing_Click()
        CottonDyeing.Show
        CottonDyeing.Top = 0
        CottonDyeing.Height = 8700
        CottonDyeing.Width = 8000
        CottonDyeing.Left = 1700
End Sub
Private Sub Mnu_CottonReDying_Click()
        CottonReDyeing.Show
        CottonReDyeing.Top = 0
        CottonReDyeing.Height = 8700
        CottonReDyeing.Width = 8000
        CottonReDyeing.Left = 1700
End Sub

Private Sub mnuAddItem_Click()
        Item.Show
        Item.Left = 2000
        Item.Top = 1000
End Sub
Private Sub mnuAddItemType_Click()
        ItemType.Show
        ItemType.Left = 2000
        ItemType.Top = 1000
End Sub
Private Sub MnuAddParty_Click()
        Party.Show
        Party.Left = 2000
        Party.Top = 1000
End Sub
Private Sub MnuAddPartyType_Click()
        PartyType.Show
        PartyType.Left = 2000
        PartyType.Top = 1000
End Sub
Private Sub MnuAllPending_Click()
        RepAllPending.Show
        RepAllPending.Left = 2000
        RepAllPending.Top = 1000
End Sub

Private Sub MnuAllProduction_Click()
    If usr = "admin" Then
        MsgBox ("No Privilege!")
    Else
        RepAllProduction.Show
        RepAllProduction.Left = 2000
        RepAllProduction.Top = 1000
    End If
End Sub

Private Sub mnuChangePassword_Click()
        ChangePassword.Show
        ChangePassword.Left = 2000
        ChangePassword.Top = 1000
End Sub

Private Sub mnuDeliv_Click()
        RepDelivery.Show
        RepDelivery.Left = 2000
        RepDelivery.Top = 1000
End Sub

Private Sub MnuDelivery_Click()
        Delivery.Show
        Delivery.Height = 7640
        Delivery.Width = 9540
        Delivery.Left = 2000
        Delivery.Top = 500
End Sub

Private Sub MnuDeliveryParty_Click()
        RepDeliveryParty.Show
        RepDeliveryParty.Left = 2000
        RepDeliveryParty.Top = 1000
End Sub

Private Sub mnuDptAct_Click()
        RepPartyItem.Show
        RepPartyItem.Left = 2000
        RepPartyItem.Top = 1000
End Sub

Private Sub MnuExit_Click()
    Unload Me
End Sub

Private Sub MnuItem_Click()
        RepAllItem.Show
        RepAllItem.Left = 2000
        RepAllItem.Top = 1000
End Sub

Private Sub MnuItemDelivery_Click()
        RepDeliveryPartyI.Show
        RepDeliveryPartyI.Left = 2000
        RepDeliveryPartyI.Top = 1000
End Sub
Private Sub MnuItemProduction_Click()
    If usr = "admin" Then
        MsgBox ("No Privilege!")
    Else
        RepItemProduction.Show
        RepItemProduction.Left = 2000
        RepItemProduction.Top = 1000
    End If
End Sub
Private Sub MnuItemSearch_Click()
        Item_Search.Show
        Item_Search.Left = 2000
        Item_Search.Top = 1000
End Sub
Private Sub MnuItemTypeDelivery_Click()
        RepDeliveryPartyIT.Show
        RepDeliveryPartyIT.Left = 2000
        RepDeliveryPartyIT.Top = 1000
End Sub
Private Sub MnuMachineProduction_Click()
    If usr = "admin" Then
        MsgBox ("No Privilege!")
    Else
        RepMachineProduction.Show
        RepMachineProduction.Left = 2000
        RepMachineProduction.Top = 1000
    End If
End Sub
Private Sub mnuoldprocess_Click()
        old_Process.Show
        old_Process.Height = 8700
        old_Process.Width = 8900
        old_Process.Left = 1700
End Sub
Private Sub mnuParty_Click()
        Rpt_AllParty.Show
        Rpt_AllParty.Left = 2000
        Rpt_AllParty.Top = 1000
End Sub
Private Sub MnuPartySearch_Click()
        Party_Search.Show
        Party_Search.Left = 2000
        Party_Search.Top = 1000
End Sub
Private Sub MnuPolisterCosting_Click()
    If usr = "admin" Then
        MsgBox ("No Privilege!")
    Else
        RepPolisterCost.Show
        RepPolisterCost.Left = 2000
        RepPolisterCost.Top = 1000
    End If
End Sub
Private Sub MnuProcess_Click()
        Process.Show
        Process.Height = 10800
        Process.Width = 10900
        Process.Left = 1700
End Sub
Private Sub mnuReceive_Click()
        Receivings.Show
        Receivings.Height = 6400
        Receivings.Width = 8250
        Receivings.Left = 2000
        Receivings.Top = 500
End Sub
Private Sub mnuReceiveOrder_Click()
        Receiving_Order.Show
        Receiving_Order.Height = 6400
        Receiving_Order.Width = 8250
        Receiving_Order.Left = 2000
        Receiving_Order.Top = 500
End Sub
Private Sub mnuReceiveQuantity_Click()
    RepReceiveQuantity.Show
    RepReceiveQuantity.Left = 3500
    RepReceiveQuantity.Top = 2700
End Sub
Private Sub mnuReceiveReturn_Click()
        Receive_Return.Show
        Receive_Return.Width = 8250
        Receive_Return.Height = 6400
        Receive_Return.Left = 2000
        Receive_Return.Top = 500
End Sub
Private Sub mnuRecipe_Click()
        Recipe.Show
        Recipe.Width = 8250
        Recipe.Height = 6400
        Recipe.Left = 2000
        Recipe.Top = 500
End Sub
Private Sub MnuRepProcess_Click()
    RepProcess.Show
    RepProcess.Left = 3500
    RepProcess.Top = 2700
End Sub
Private Sub MnuSale_Click()
        Sale.Show
        Sale.Width = 8250
        Sale.Height = 6400
        Sale.Left = 2000
        Sale.Top = 500
End Sub
Private Sub MnuTypeProduction_Click()
    If usr = "admin" Then
        MsgBox ("No Privilege!")
    Else
        RepProduction.Show
        RepProduction.Left = 2000
        RepProduction.Top = 1000
    End If
End Sub
Private Sub muitemAct_Click()
        RepDateItem.Show
        RepDateItem.Left = 2000
        RepDateItem.Top = 1000
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        Party.Show
        Party.Left = 2000
        Party.Top = 1000
    ElseIf Button.Index = 2 Then
        Party_Search.Show
        Party_Search.Left = 2000
        Party_Search.Top = 1000
    ElseIf Button.Index = 4 Then
        Item.Show
        Item.Left = 2000
        Item.Top = 1000
    ElseIf Button.Index = 5 Then
        Item_Search.Show
        Item_Search.Left = 2000
        Item_Search.Top = 1000
    End If
End Sub
