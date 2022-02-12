'---------------------------------------------------------------------------------------
' Module      : mGenGUID
' Fecha       : 05/02/2009 18:10
' Autor       : XcryptOR
' Proposito   : Generar un número de identificación unico
' Creditos    : Creditos a trilithium, Autor del code original en Delphi
'---------------------------------------------------------------------------------------
 
Option Explicit
 
Private Type GUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(7)        As Byte
End Type
 
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
pDest As Any, _
pSource As Any, _
ByVal dwLength As Long)
 
Private Declare Function StringFromCLSID Lib "ole32" ( _
pclsid As GUID, _
lpsz As Long) As Long
 
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
 
Public Function GetGUID() As String
    Dim udtGUID     As GUID
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = GUIDToStr(udtGUID)
    End If
End Function
 
Private Function GUIDToStr(ID As GUID) As String
    Dim strRet      As String
    Dim ptrSource   As Long
    Dim lngRet      As Long
 
    strRet = Space(38)
    lngRet = StringFromCLSID(ID, ptrSource)
    If lngRet = 0 Then
        CopyMemory ByVal StrPtr(strRet), ByVal ptrSource, 76
        GUIDToStr = strRet
    End If
End Function
