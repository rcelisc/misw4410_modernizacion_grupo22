Attribute VB_Name = "Module1"
Option Explicit

Global Const LISTVIEW_MODE0 = "Ver iconos grandes"
Global Const LISTVIEW_MODE1 = "Ver iconos pequeños"
Global Const LISTVIEW_MODE2 = "Ver lista"
Global Const LISTVIEW_MODE3 = "Ver detalles"
Public fMainForm As frmMain


' Acceso a .ini ***********************************************
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
' Obtiene desde .ini
Function GetFromINI(sSection As String, sKey As String, sDefault As String, sIniFile As String)
    Dim sBuffer As String, lRet As Long
    sBuffer = String$(255, 0)
    lRet = GetPrivateProfileString(sSection, sKey, "", sBuffer, Len(sBuffer), sIniFile)
    If lRet = 0 Then
        If sDefault <> "" Then AddToINI sSection, sKey, sDefault, sIniFile
        GetFromINI = sDefault
    Else
        GetFromINI = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    End If
End Function
' Guarda en .ini
Function AddToINI(sSection As String, sKey As String, sValue As String, sIniFile As String) As Boolean
    Dim lRet As Long
    lRet = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
    AddToINI = (lRet)
End Function
' FIN Acceso a .ini ***********************************************


Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    
    ' Abrir Conexión BD
    If Not mDatos.AbrirCx Then
        MsgBox "No se ha podido abrir la Conexión con la Base de datos", vbCritical, "Error de Acceso a Datos - " & App.Title
        End
    End If
    
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash
    
    
    fMainForm.Show
End Sub

