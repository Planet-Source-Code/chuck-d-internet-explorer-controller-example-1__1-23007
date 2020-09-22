VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Explorer Controller Example 1"
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList HotIcon 
      Left            =   600
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":055C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1014
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ColdIcon 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1570
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2028
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2584
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   1138
      ButtonWidth     =   1376
      ButtonHeight    =   1085
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ColdIcon"
      HotImageList    =   "HotIcon"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Favorites"
            Object.Tag             =   "Show Favorites Bar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "History"
            Object.Tag             =   "Show History Bar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Newest Code at PSC"
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const HWND_TOPMOST = -1                               'Needed for Topmost window (Form1)
Const HWND_NOTOPMOST = -2                             'Needed for Topmost window (Form1)
Const SWP_NOSIZE = &H1                                'Needed for Topmost window (Form1)
Const SWP_NOMOVE = &H2                                'Needed for Topmost window (Form1)
Const SWP_NOACTIVATE = &H10                           'Needed for Topmost window (Form1)
Const SWP_SHOWWINDOW = &H40                           'Needed for Topmost window (Form1)
Const FavoritesBar = "{EFA24E61-B078-11D0-89E4-00C04FC9E26E}"    'Needed to show Favorites Bar In Internet Explorer
Const HistoryBar = "{EFA24E62-B078-11D0-89E4-00C04FC9E26E}"    'Needed to show History Bar In Internet Explorer
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)    'Needed for Topmost window (Form1)
Dim Browser As New InternetExplorer                   'Create a new instance of Internet Explorer

Private Sub Command1_Click()
    'Have Internet Explorer navigate to PSC-VB-Newest Code URL, Then Load the URL into the same window (Same because PSC doesn't use FRAMES)
    Browser.Navigate "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWId=1", "", "_self"
End Sub

Private Sub Form_Activate()
    'Set (Form1) as Topmost window
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Load()
    Browser.ToolBar = False                           'No toolbar shown for Internet Explorer
    Browser.StatusBar = True                          'StatusBar shown for Internet Explorer
    Browser.Visible = True                            'Internet Explorer is Visible
    Browser.Navigate "http://www.planet-source-code.com/vb/default.asp?lngWId=1"    'navigate to PSC-VB URL
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Tag
        Case "Show Favorites Bar"                     'Show Favorites Bar in Internet Explorer
            Browser.ShowBrowserBar FavoritesBar, True
            Button.Tag = "Hide Favorites Bar"
            Toolbar1.Buttons(2).Tag = "Show History Bar"
        Case "Show History Bar"                       'Show History Bar in Internet Explorer
            Browser.ShowBrowserBar HistoryBar, True
            Button.Tag = "Hide History Bar"
            Toolbar1.Buttons(1).Tag = "Show Favorites Bar"
        Case "Hide Favorites Bar"                     'Hide Favorites Bar in Internet Explorer
            Browser.ShowBrowserBar FavoritesBar, False
            Button.Tag = "Show Favorites Bar"
        Case "Hide History Bar"                       'Hide History Bar in Internet Explorer
            Browser.ShowBrowserBar HistoryBar, False
            Button.Tag = "Show History Bar"
        Case "Back"                                   'Browser Back button
            Browser.GoBack
        Case "Forward"                                'Browser Forward button
            Browser.GoForward
    End Select
End Sub
