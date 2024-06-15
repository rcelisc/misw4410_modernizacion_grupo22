VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Opciones"
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Tag             =   "Aceptar"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "Cancelar"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Tag             =   "&Aplicar"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Height          =   3705
         Left            =   90
         TabIndex        =   9
         Top             =   -30
         Width           =   5520
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Height          =   3705
         Left            =   90
         TabIndex        =   7
         Top             =   -30
         Width           =   5520
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   210
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample1 
         Height          =   3705
         Left            =   90
         TabIndex        =   4
         Top             =   -30
         Width           =   5520
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1920
            TabIndex        =   13
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1920
            TabIndex        =   11
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label2 
            Caption         =   "Cadena de Conexion"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Cadena de Conexion"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   1695
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Avanzado"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Base de Datos"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    'ToDo: Add 'cmdApply_Click' code.
    MsgBox "Aplicar código va aquí para establecer opciones sin cerrar el cuadro de diálogo"
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    'Pendiente: Agregar código para 'cmdOK_Click'.
    MsgBox "Aquí se coloca código para establecer opciones y cerrar el cuadro de diálogo"
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'controlar ctrl+tab para mover a la siguiente ficha
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'última ficha, por lo que hay que volver a la primera ficha
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'incrementar la ficha
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'última ficha, por lo que hay que volver a la primera ficha
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
        Else
            'incrementa la ficha
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If
End Sub


Private Sub tbsOptions_Click()
    

    Dim i As Integer
    'mostrar y habilitar los controles seleccionados de la ficha
    'y ocultar y deshabilitar los demás
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    

End Sub

