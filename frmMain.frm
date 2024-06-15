VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "MonitorPolBasicas"
   ClientHeight    =   3225
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
      Top             =   720
   End
   Begin MSComctlLib.ImageList imlIconsBig 
      Left            =   5640
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   7
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "Ejecutando"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0060
            Key             =   "Informacion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":059E
            Key             =   "Avanzar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":061C
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0868
            Key             =   "Libro"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0924
            Key             =   "Guardando"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D00
            Key             =   "Eliminar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E4E
            Key             =   "Bloqueo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F07
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1259
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15AB
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIconsSmall 
      Left            =   5640
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18FD
            Key             =   "Ejecutando"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":195D
            Key             =   "Informacion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E9B
            Key             =   "Avanzar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F19
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2165
            Key             =   "Libro"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2221
            Key             =   "Guardando"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25FD
            Key             =   "Eliminar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":274B
            Key             =   "Bloqueo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2804
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B56
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EA8
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   7
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6585
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   6585
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   200
         Left            =   2110
         TabIndex        =   8
         Top             =   50
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   2078
         TabIndex        =   4
         Tag             =   " Vista Lista:"
         Top             =   12
         Width           =   3216
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Tag             =   " Vista �rbol:"
         Top             =   12
         Width           =   2016
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2955
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5980
            Text            =   "Estado"
            TextSave        =   "Estado"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "26/11/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "12:17 p.m."
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31FA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":330C
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":341E
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3530
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3642
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3754
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3866
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3978
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A8A
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B9C
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CAE
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F80
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42D2
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4624
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Actualizar"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Atr�s"
            Object.ToolTipText     =   "Atr�s"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Adelante"
            Object.ToolTipText     =   "Adelante"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Propiedades"
            Object.ToolTipText     =   "Propiedades"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ver iconos grandes"
            Object.ToolTipText     =   "Ver iconos grandes"
            ImageKey        =   "View Large Icons"
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ver iconos peque�os"
            Object.ToolTipText     =   "Ver iconos peque�os"
            ImageKey        =   "View Small Icons"
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ver lista"
            Object.ToolTipText     =   "Ver lista"
            ImageKey        =   "View List"
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ver detalles"
            Object.ToolTipText     =   "Ver detalles"
            ImageKey        =   "View Details"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Iniciar"
            Object.ToolTipText     =   "Iniciar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Pausar"
            Object.ToolTipText     =   "Pausar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Detener"
            Object.ToolTipText     =   "Detener"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4800
      Left            =   2040
      TabIndex        =   5
      Top             =   705
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   8467
      SortKey         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlIconsBig"
      SmallIcons      =   "imlIconsSmall"
      ColHdrIcons     =   "imlIconsSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "NoSolicitud"
         Text            =   "NoSolicitud"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Estado"
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "TipoId"
         Text            =   "Tipo Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "NumId"
         Text            =   "Numero Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Rol"
         Text            =   "Rol"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Linea"
         Text            =   "Tipo Solicitud(Linea)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Categoria"
         Text            =   "Categoria"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4800
      Left            =   0
      TabIndex        =   6
      Top             =   705
      Width           =   2016
      _ExtentX        =   3545
      _ExtentY        =   8467
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Image imgSplitter 
      Height          =   4788
      Left            =   1965
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileFind 
         Caption         =   "B&uscar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "En&viar a"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIniciar 
         Caption         =   "Iniciar"
      End
      Begin VB.Menu mnuPausar 
         Caption         =   "Pausar"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDetener 
         Caption         =   "Detener"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Eliminar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "&Propiedades"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Cerrar"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Ver"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Barra de herramientas"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "B&arra de estado"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Iconos g&randes"
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Iconos pe&que�os"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Li&sta"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Det&alles"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Renovar"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Opciones..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Acerca de "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  
Dim mbMoving As Boolean
Dim blDetener As Boolean
Const sglSplitLimit = 500

' Variables de Acceso a Datos
Dim arsPrimary As ADODB.Recordset

Private Sub Form_Activate()
    lvListView.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
    Select Case lvListView.View
        Case lvwIcon
            tbToolBar.Buttons(LISTVIEW_MODE0).Value = tbrPressed
        Case lvwSmallIcon
            tbToolBar.Buttons(LISTVIEW_MODE1).Value = tbrPressed
        Case lvwList
            tbToolBar.Buttons(LISTVIEW_MODE2).Value = tbrPressed
        Case lvwReport
            tbToolBar.Buttons(LISTVIEW_MODE3).Value = tbrPressed
    End Select

End Sub

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    tvTreeView.Nodes.Add , , "kRoot", "Procesos"
    fCargarProcesos
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Cierra la Conexi�n con la BD
    mDatos.CerrarCx
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    SaveSetting App.Title, "Settings", "ViewMode", lvListView.View
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    If Source = imgSplitter Then
        SizeControls X
    End If
End Sub


Sub SizeControls(X As Single)
    On Error Resume Next
    

    'establecer el ancho
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    tvTreeView.Width = X
    imgSplitter.Left = X
    lvListView.Left = X + 40
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = lvListView.Left + 20
    ProgressBar1.Left = lblTitle(1).Left + 30
    lblTitle(1).Width = lvListView.Width - 40
    ProgressBar1.Width = lblTitle(1).Width - 60


    'establecer la coordenada superior
  

    If tbToolBar.Visible Then
        tvTreeView.Top = tbToolBar.Height + picTitles.Height
    Else
        tvTreeView.Top = picTitles.Height
    End If

  lvListView.Top = tvTreeView.Top
    

    'establecer el alto
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If
    

    lvListView.Height = tvTreeView.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
End Sub

Private Sub lvListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvListView.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub mnuDetener_Click()
    Timer1.Enabled = False
    mnuIniciar.Enabled = True
    mnuDetener.Enabled = False
    tbToolBar.Buttons(17).Enabled = True
    tbToolBar.Buttons(19).Enabled = False
    blDetener = True
End Sub

Private Sub mnuIniciar_Click()
    Timer1.Enabled = True
    mnuIniciar.Enabled = False
    mnuDetener.Enabled = True
    tbToolBar.Buttons(17).Enabled = False
    tbToolBar.Buttons(19).Enabled = True
    blDetener = False
End Sub

Private Sub mnuPausar_Click()
    MsgBox ""
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Nuevo"
            mnuFileNew_Click
        Case "Deshacer"
            'TareasPendientes: Agregar c�digo de bot�n 'Deshacer'.
            MsgBox "Agregar c�digo de bot�n 'Deshacer'."
        Case "Atr�s"
            'TareasPendientes: Agregar c�digo de bot�n 'Atr�s'.
            MsgBox "Agregar c�digo de bot�n 'Atr�s'."
        Case "Adelante"
            'TareasPendientes: Agregar c�digo de bot�n 'Adelante'.
            MsgBox "Agregar c�digo de bot�n 'Adelante'."
        Case "Eliminar"
            mnuFileDelete_Click
        Case "Propiedades"
            mnuFileProperties_Click
        Case "Ver iconos grandes"
            lvListView.View = lvwIcon
        Case "Ver iconos peque�os"
            lvListView.View = lvwSmallIcon
        Case "Ver lista"
            lvListView.View = lvwList
        Case "Ver detalles"
            lvListView.View = lvwReport
        Case "Iniciar"
            mnuIniciar_Click
        Case "Pausar"
            mnuPausar_Click
        Case "Detener"
            mnuDetener_Click
        Case "Actualizar"
            mnuViewRefresh_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    ' Detiene si esta corriendo
    mnuDetener_Click
    ' Limpia la Vista
    lvListView.ListItems.Clear
    ' Trae los nuevos procesos(Estado: No Procesado)
    fCargarProcesos False
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuFileClose_Click()
    'descargar el formulario
    Unload Me

End Sub

Private Sub mnuFileProperties_Click()
    'TareasPendientes: Agregar c�digo 'mnuFileProperties_Click'.
    MsgBox "Agregar c�digo 'mnuFileProperties_Click'."
End Sub

Private Sub mnuFileDelete_Click()
    'TareasPendientes: Agregar c�digo 'mnuFileDelete_Click'.
    MsgBox "Agregar c�digo 'mnuFileDelete_Click'."
End Sub

Private Sub mnuFileNew_Click()
    'TareasPendientes: Agregar c�digo 'mnuFileNew_Click'.
    MsgBox "Agregar c�digo 'mnuFileNew_Click'."
End Sub

Private Sub mnuFileSendTo_Click()
    'TareasPendientes: Agregar c�digo 'mnuFileSendTo_Click'.
    MsgBox "Agregar c�digo 'mnuFileSendTo_Click'."
End Sub


Private Sub mnuFileFind_Click()
    'TareasPendientes: Agregar c�digo 'mnuFileFind_Click'.
    MsgBox "Agregar c�digo 'mnuFileFind_Click'."
End Sub

Private Sub fCargarProcesos(Optional blCompleto As Boolean)
    On Error GoTo Err_label
    Dim strSQL As String
    Dim lvItem As ListItem
    
    ' Obtiene los Items a Procesar
    strSQL = "SELECT IPS.NoSolicitud, IPS.TipoId, IPS.NumId, IPS.CodRolP, S.CodTipoS, S.CodCategoria "
    strSQL = strSQL & " FROM ITR_PERSONA_SOL IPS"
    strSQL = strSQL & " INNER JOIN DB_Solicitud S ON (IPS.NoSolicitud = S.NoSolicitud) "
    strSQL = strSQL & " WHERE IPS.CodEstado = '2' "
    If blCompleto Then
        strSQL = strSQL & " AND IPS.CodValor IN ('1', '8') "
    Else
        strSQL = strSQL & " AND IPS.CodValor = '1' "
    End If
    strSQL = strSQL & " UNION "
    strSQL = strSQL & " SELECT IES.NoSolicitud, '2' AS TipoId, IES.NIT, IES.CodRolE, S.CodTipoS, S.CodCategoria "
    strSQL = strSQL & " FROM ITR_EMPRESA_SOL IES"
    strSQL = strSQL & " INNER JOIN DB_Solicitud S ON (IES.NoSolicitud = S.NoSolicitud) "
    strSQL = strSQL & " WHERE IES.CodEstado = '2' "
    If blCompleto Then
        strSQL = strSQL & " AND IES.CodValor IN ('1', '8') "
    Else
        strSQL = strSQL & " AND IES.CodValor = '1' "
    End If
    
    Set arsPrimary = mDatos.adcConexion.Execute(strSQL)
    
    ' Sube los Items a Procesar al ListView
    Do While Not arsPrimary.EOF
        If Not fBuscarItem("k" & arsPrimary.Fields("NoSolicitud") & "@" & arsPrimary.Fields("TipoId") & "@" & arsPrimary.Fields("NumId") & "@" & arsPrimary.Fields("CodRolP")) Then
            Set lvItem = lvListView.ListItems.Add(, "k" & arsPrimary.Fields("NoSolicitud") & "@" & arsPrimary.Fields("TipoId") & "@" & arsPrimary.Fields("NumId") & "@" & arsPrimary.Fields("CodRolP"), arsPrimary.Fields("NoSolicitud"), 2, 2)
            lvItem.SubItems(1) = ""
            lvItem.SubItems(2) = arsPrimary.Fields("TipoId")
            lvItem.SubItems(3) = arsPrimary.Fields("NumId")
            lvItem.SubItems(4) = arsPrimary.Fields("CodRolP")
            lvItem.SubItems(5) = arsPrimary.Fields("CodTipoS")
            lvItem.SubItems(6) = arsPrimary.Fields("CodCategoria")
        End If
        
        arsPrimary.MoveNext
    Loop
    
    ' Barra de Progreso
    ProgressBar1.Max = lvListView.ListItems.Count
    ProgressBar1.Value = 0
    
    Exit Sub
Err_label:
    MsgBox Err.Description & "(" & Err.Number & ")"
End Sub

' Procesa las solicitudes
Private Function fProcesar(strTipId As String, strNumId As String, strNoSolicitud As String, strRol As String, strTipoSol As String, strLineaSol As String)
    On Error GoTo Err_label
    Dim strSQL As String
    
    fProcesar = ""
    
    strSQL = "EXEC SP_POLITICAS_BASICAS '" & strNoSolicitud & "',  " & strTipoSol & ", " & strLineaSol & ""

    'MsgBox strSQL
    
    
    
    Exit Function
Err_label:
    fProcesar = Err.Description & "(" & Err.Number & ")"
End Function

' Determina si un Item se encuentra en la coleccion de Items por la KEY
Private Function fBuscarItem(strItemID As String) As Boolean
    On Error GoTo Err_label
    Dim Temp
    fBuscarItem = True
    Temp = lvListView.ListItems(strItemID).Bold
    
    Exit Function
Err_label:
    fBuscarItem = False
End Function

Private Sub Timer1_Timer()
    On Error GoTo Err_label
    Dim i As Long
    Dim strResultado As String
    Timer1.Enabled = False
    
    ' Procesar Items
    For i = lvListView.ListItems.Count To 1 Step -1
        DoEvents
        If blDetener Then '  Si se seleccion� detener
            Exit Sub
        End If
        lvListView.ListItems(i).SubItems(1) = "Procesando" ' Estado
        
        ' Envia a procesar el Item con los siguientes par�metros:
        'lvItem.Text = NoSolicitud
        'lvItem.SubItems(2) = TipoId
        'lvItem.SubItems(3) = NumId
        'lvItem.SubItems(4) = CodRolP
        'lvItem.SubItems(5) = CodTipoS
        'lvItem.SubItems(6) = CodCategoria ' Linea
        strResultado = fProcesar(lvListView.ListItems(i).SubItems(2), _
        lvListView.ListItems(i).SubItems(3), _
        lvListView.ListItems(i).Text, _
        lvListView.ListItems(i).SubItems(4), _
        lvListView.ListItems(i).SubItems(5), _
        lvListView.ListItems(i).SubItems(6))
        
        ' Progress Bar - Progreso de procesamientos
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        If strResultado = "" Then
            lvListView.ListItems(i).SubItems(1) = "OK" ' Estado
            lvListView.ListItems.Remove i
        Else
            lvListView.ListItems(i).SubItems(1) = strResultado ' Estado
        End If
    Next i
    
    ' Trae los nuevos procesos(Estado: No Procesado)
    fCargarProcesos False
    
    Timer1.Enabled = True
    
    Exit Sub
Err_label:
    MsgBox Err.Description & "(" & Err.Number & ")"
End Sub
