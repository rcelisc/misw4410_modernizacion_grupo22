VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de MonitorPolBasicas"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Acerca de MonitorPolBasicas"
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Tag             =   "Aceptar"
      Top             =   2625
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "Info. del &sistema..."
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Tag             =   "Info. del &sistema..."
      Top             =   3075
      Width           =   1452
   End
   Begin VB.Label lblDescription 
      Caption         =   "Descripción de la aplicación"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   6
      Tag             =   "Descripción de la aplicación"
      Top             =   1125
      Width           =   4092
   End
   Begin VB.Label lblTitle 
      Caption         =   "Título de la aplicación"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Tag             =   "Título de la aplicación"
      Top             =   240
      Width           =   4092
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225
      X2              =   5657
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Tag             =   "Versión"
      Top             =   780
      Width           =   4092
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Advertencia: ..."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   3
      Tag             =   "Advertencia: ..."
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Opciones de seguridad de clave del Registro...
Const KEY_ALL_ACCESS = &H2003F
                                          

' Tipos ROOT de claves del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' cadena terminada en valor nulo Unicode
Const REG_DWORD = 4                      ' número de 32 bits


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub



Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOK_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Intentar obtener el nombre y la ruta del programa en el Registro...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Intentar obtener sólo la ruta del programa en el Registro...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Validar la existencia de versión conocida de 32 bits de archivo
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Error: no se encuentra el archivo...
                Else
                        GoTo SysInfoErr
                End If
        ' Error: no se encuentra la entrada del Registro...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "La información del sistema no está disponible en este momento", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Contador de bucle
        Dim rc As Long                                          ' Código de retorno
        Dim hKey As Long                                        ' Controlador a una clave de Registro abierta
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Tipo de datos de una clave del Registro
        Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave del Registro
        Dim KeyValSize As Long                                  ' Tamaño de variable de clave del Registro
        '------------------------------------------------------------
        ' Abrir RegKey bajo KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir la clave del Registro
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar error...
        

        tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
        KeyValSize = 1024                                       ' Marcar tamaño de variable
        

        '------------------------------------------------------------
        ' Obtener valor de clave del Registro...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determinar el tipo de valor de clave para conversión...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Buscar tipos de datos...
        Case REG_SZ                                             ' Tipo de datos String de clave del Registro
                KeyVal = tmpVal                                     ' Copiar valor String
        Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
                For i = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor carácter a carácter
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a String
        End Select
        

        GetKeyValue = True                                      ' Operación realizada correctamente
        rc = RegCloseKey(hKey)                                  ' Cerrar clave del Registro
        Exit Function                                           ' Salir
        

GetKeyError:    ' Limpiar después de que se produzca un error...
        KeyVal = ""                                             ' Establecer el valor de retonor a la cadena vacía
        GetKeyValue = False                                     ' La operación no se ha realizado correctamente
        rc = RegCloseKey(hKey)                                  ' Cerrar clave del Registro
End Function

