Attribute VB_Name = "Mod_BasDat"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'-------------------------------------------------------------------------------
'Modificaciones realizadas el
'
'-------------------------------------------------------------------------------
'1.) -- 05/02/2011 -- ABV
'Se agregó la funcionalidad de agregar los nuevos campos para el requerimiento del Reajuste al 2%,
'por lo cual se modificaron:
'Formularios de mantención y consulta de Pólizas
'Reportes
'Rutinas de cálculo, las cuales deben considerar un reajuste de trimestres fijos

