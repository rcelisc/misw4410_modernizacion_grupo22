Attribute VB_Name = "mDatos"
Option Explicit

' Variables Públicas
Public adcConexion As ADODB.Connection

Public Function AbrirCx() As Boolean
    On Error GoTo Err_label
    AbrirCx = False
    Set mDatos.adcConexion = New ADODB.Connection

    mDatos.adcConexion.Open "DSN=ODBCDesarrollo", "sa", "sa123456"
    
    AbrirCx = True
    
    Exit Function
Err_label:
    MsgBox Err.Description & "(" & Err.Number & ")"
End Function

Public Function CerrarCx() As Boolean
    On Error Resume Next
    CerrarCx = False
    
    mDatos.adcConexion.Close
    Set mDatos.adcConexion = Nothing
    
    CerrarCx = True
    
    Exit Function
Err_label:
    MsgBox Err.Description & "(" & Err.Number & ")"
End Function

