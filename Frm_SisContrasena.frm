VERSION 5.00
Begin VB.Form Frm_SisContrasena 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Contrase�a."
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "Frm_SisContrasena.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4800
   Begin VB.Frame Fra_Datos 
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4575
      Begin VB.Frame Frame1 
         Caption         =   "Reglas de Password"
         Height          =   1575
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   4335
         Begin VB.Label lbl_2 
            Caption         =   "2.- No podra utilizar los 6 ultimos password ingresados."
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "! # $ % & / ( ) = ? � ' � { } ^ ` [ ] * \ - + . , ; : _ "
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   4095
         End
         Begin VB.Label lbl_3 
            Caption         =   "3.- Minimo 1 caracter alfanum�ricos tales como:"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   4095
         End
         Begin VB.Label lbl_1 
            Caption         =   "1.- El tama�o m�nimo de caracteres es "
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.TextBox Txt_Actual 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Txt_Repetir 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Txt_Contrase�a 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Frm_SisContrasena.frx":0442
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Lbl_Etiqueta 
         Caption         =   "Ingrese Contrase�a Actual"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Lbl_Etiqueta 
         Caption         =   "Repita Nueva Contrase�a  "
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Lbl_Etiqueta 
         Caption         =   "Ingrese Nueva Contrase�a  "
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.Frame Fra_Botones 
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   4575
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   2440
         Picture         =   "Frm_SisContrasena.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmd_grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1320
         Picture         =   "Frm_SisContrasena.frx":097E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm_SisContrasena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vlConexion  As ADODB.Connection
Dim vlRecord    As ADODB.Recordset

Dim vlOperacion As String
Dim vlPassword  As String
Dim vlSw        As Boolean

'------------------------------------------------------------
'Permite Modificar la Contrase�a del Usuario Autorizado
'------------------------------------------------------------
Function flRegistrar()
Dim iContrase�a As String

On Error GoTo Err_Registrar

    'Validar el registro del Rut Usuario Autorizado
    If (Trim(vgLogin) = "") Then
        MsgBox "Error en el registro del Usuario autorizado para el Sistema. No es posible modificaci�n de Password de Usuario.", vbCritical, "Error de Usuario"
        cmd_salir.SetFocus
        Exit Function
    End If
    
    'Validar el ingreso de Contrase�a Actual
    If (Trim(Txt_Actual) = "") Then
        MsgBox "La Contrase�a Actual no ha sido ingresada para realizar el Cambio.", vbCritical, "Error de Dato"
        Txt_Actual.SetFocus
        Exit Function
    End If
    
    'Validar la Contrase�a y su Repetici�n
    If (Trim(Txt_Contrase�a) <> "") And (Txt_Repetir <> "") Then
        Txt_Contrase�a = Trim(Txt_Contrase�a)
        Txt_Repetir = Trim(Txt_Repetir)
        
         'RRR 18/01/2012'
        If fIaplicavalidacion(vgUsuario, Txt_Contrase�a, Txt_Repetir) = 0 Then
            vgValorAr = 1
            Exit Function
        End If
        'RRR '
        
'        If (Len(Txt_Contrase�a) < 6) Then
'            MsgBox "La contrase�a debe contar como m�nimo de 6 caracteres, y como m�ximo 10.", vbInformation, "Error de Dato"
'            Txt_Contrase�a = ""
'            Txt_Contrase�a.SetFocus
'            Exit Function
'        End If
'        If (Txt_Contrase�a <> Txt_Repetir) Then
'            MsgBox "Las Contrase�as registradas son distintas, vuelva a registrarlas.", vbExclamation, "Error de Contrase�a"
'            Txt_Contrase�a = ""
'            Txt_Repetir = ""
'            Txt_Contrase�a.SetFocus
'            Exit Function
'        End If
    Else
        MsgBox "Debe ingresar la Contrase�a y su correspondiente repetici�n para ser registrado.", vbCritical, "Error de Contrase�a"
        Txt_Contrase�a = ""
        Txt_Contrase�a.SetFocus
        Exit Function
    End If

    'Encriptar Contrase�a
    vlPassword = fgEncPassword(Txt_Contrase�a)
    iContrase�a = fgEncPassword(Txt_Actual)
    If (vlPassword = "") Then
        MsgBox "Error en la transformaci�n de la Password o Contrase�a por el Sistema.", vbCritical, "Error de Transformaci�n"
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    'Verificar existencia del Login de Usuario para la Actualizaci�n
'    vgQuery = "SELECT cod_usuario FROM MA_tmae_usuario WHERE "
'    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' and "
'    vgQuery = vgQuery & "cod_usuario = '" & vgLogin & "' and "
'    vgQuery = vgQuery & "gls_password = '" & iContrase�a & "'"
    
    vgQuery = " SELECT A.cod_usuario FROM MA_TMAE_USUPASSWORD A"
    vgQuery = vgQuery & " JOIN MA_TMAE_USUMODULO B ON A.COD_USUARIO=B.COD_USUARIO"
    vgQuery = vgQuery & " WHERE cod_sistema = '" & vgTipoSistema & "' and A.cod_usuario = '" & vgLogin & "'"
    vgQuery = vgQuery & " AND NRO_USUPASS=(SELECT MAX(NRO_USUPASS) FROM MA_TMAE_USUPASSWORD WHERE COD_USUARIO=A.COD_USUARIO)"
    vgQuery = vgQuery & " and gls_password = '" & iContrase�a & "'"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If (vgRs.EOF) Then
        vlOperacion = "E"
    Else
        vlOperacion = "A"
    End If
    vgRs.Close
    
    If (vlOperacion = "A") Then
        'Actualiza Password
        vgRes = MsgBox("� Est� seguro que desea Modificar la Contrase�a ?", 4 + 32 + 256, "Operaci�n de Actualizaci�n")
        If vgRes <> 6 Then
            cmd_salir.SetFocus
            Screen.MousePointer = 0
            Exit Function
        End If
        
        vgContrase�a = vlPassword
        
        'RRR se guarda primero en tabla maestra matriz
        '18/01/2012
        Dim c As Integer
        
        vgSql = "select count(*) as num from MA_TMAE_USUPASSWORD where cod_usuario='" & vgUsuario & "'"
        Set vgRs = vgConexionBD.Execute(vgSql)
        
        c = CInt(vgRs!num) + 1
         
         ''INSERTA NUEVA CONTRASE�A AL HISTORICO LA TABLA USUARIOMATRIZ
        vgSql = " insert into MA_TMAE_USUPASSWORD(cod_usuario, nro_usupass, fec_inipass, fec_finpass, fec_antPass, gls_password,"
        vgSql = vgSql & " gls_passwordconf, ind_Segu, cod_usucrea,fec_crea,hor_crea)"
        vgSql = vgSql & " values('" & vgUsuario & "'," & c & ",'" & FechaIni & "', '" & FechaFin & "', '" & fechaant & "', '" & vgContrase�a & "',"
        vgSql = vgSql & " '" & vgContrase�a & "', '1', '" & vgUsuario & "' ,'" & Format(Date, "yyyymmdd") & "','" & Format(Time, "hhmmss") & "')"
        vgConexionBD.Execute (vgSql)
        
'        ''ACTUALIZA LA TABLA USUARIO
'        vgQuery = "UPDATE MA_tmae_usuario SET "
'        vgQuery = vgQuery & "gls_password = '" & vlPassword & "' "
'        vgQuery = vgQuery & "WHERE "
'        vgQuery = vgQuery & "cod_usuario = '" & vgLogin & "' "
'        vgConexionBD.Execute (vgQuery)
        
        'RRR
        
'        vgQuery = "UPDATE MA_tmae_usuario SET "
'        vgQuery = vgQuery & "gls_password = '" & vlPassword & "' "
'        vgQuery = vgQuery & "WHERE "
'        vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' and "
'        vgQuery = vgQuery & "cod_usuario = '" & vgLogin & "' "
'        vgConexionBD.Execute (vgQuery)
    Else
        MsgBox "El Usuario y Contrase�a no corresponden a los registrados en la Base de Datos.", vbCritical, "Operaci�n Cancelada"
    End If

    'Limpia los Casilleros de Informaci�n
    Txt_Actual = ""
    Txt_Contrase�a = ""
    Txt_Repetir = ""
    cmd_salir.SetFocus
    
    Screen.MousePointer = 0
    vgValorAr = 0
Exit Function
Err_Registrar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "�ERROR!..."
    End Select
End Function

Private Sub Cmd_Grabar_Click()
On Error GoTo Err_Grabar
    
    'Permite registrar los datos de los Usuarios
    Call flRegistrar
    
Exit Sub
Err_Grabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "�ERROR!..."
    End Select
End Sub

Private Sub Cmd_Salir_Click()
On Error GoTo Err_Descargar

   Screen.MousePointer = 11
    If vgValorAr = 1 Then
        End
    Else
        Unload Me
    End If

    Screen.MousePointer = 0

Exit Sub
Err_Descargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "�ERROR!..."
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Frm_SisContrasena.Left = 0
    Frm_SisContrasena.Top = 0

    
    vgSql = "SELECT * FROM MA_TMAE_ADMINCUENTAS WHERE "
    vgSql = vgSql & "cod_cliente = '1' "
    Set vgRs = vgConexionBD.Execute(vgSql)
    
    If Not vgRs.EOF Then
        tammin = vgRs!ntama�omin
        cantclvant = vgRs!ncanclvant
        canmincaralf = vgRs!ncaracmin
        freccambio = vgRs!nfrecuencia
        canantclv = vgRs!ncanclvant
        'balfanum = vgRs!balfanum
    End If

    lbl_1.Caption = "1.- El tama�o m�nimo de caracteres es de " & CStr(tammin) & " caracteres."
    lbl_2.Caption = "2.- No podra utilizar los " & CStr(cantclvant) & " ultimos password ingresados."
    lbl_3.Caption = "3.- Minimo " & CStr(canmincaralf) & " caracter alfanum�ricos tales como:"

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "�ERROR!..."
    End Select
End Sub

Private Sub Txt_Actual_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Txt_Actual) <> "") Then
        'If (Len(Txt_Actual) < 6) Then
        '    MsgBox "La contrase�a Actual debe contar como m�nimo de 6 caracteres, y como m�ximo 10.", vbInformation, "Error de Dato"
        '    Txt_Contrase�a.SetFocus
        'Else
            Txt_Actual = (Trim(Txt_Actual))
            Txt_Contrase�a.SetFocus
        'End If
    End If
End If
End Sub

Private Sub Txt_Contrase�a_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
'    If (Trim(Txt_Contrase�a) <> "") And (Txt_Repetir <> "") Then
'        Txt_Contrase�a = Trim(Txt_Contrase�a)
'        Txt_Repetir = Trim(Txt_Repetir)
'        If (Len(Txt_Contrase�a) < 6) Then
'            MsgBox "La contrase�a debe contar como m�nimo de 6 caracteres, y como m�ximo 10.", vbInformation, "Error de Dato"
'            Txt_Contrase�a.SetFocus
'        End If
'        If (Txt_Contrase�a <> Txt_Repetir) Then
'            MsgBox "Las Contrase�as registradas son distintas, vuelva a registrarlas.", vbExclamation, "Error de Contrase�a"
'            Txt_Contrase�a = ""
'            Txt_Repetir = ""
'        Else
'            Cmd_Grabar.SetFocus
'        End If
'    Else
        If (Trim(Txt_Contrase�a) <> "") Then
            If (Len(Txt_Contrase�a) < 6) Then
                MsgBox "La contrase�a debe contar como m�nimo de 6 caracteres, y como m�ximo 10.", vbInformation, "Error de Dato"
                Txt_Contrase�a = ""
                Txt_Contrase�a.SetFocus
            Else
                Txt_Contrase�a = (Trim(Txt_Contrase�a))
                Txt_Repetir.SetFocus
            End If
        End If
'    End If
End If
End Sub

Private Sub Txt_Repetir_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Txt_Contrase�a) <> "") And (Txt_Repetir <> "") Then
        Txt_Contrase�a = Trim(Txt_Contrase�a)
        Txt_Repetir = Trim(Txt_Repetir)
        If (Len(Txt_Contrase�a) < 6) Then
            MsgBox "La contrase�a debe contar como m�nimo de 6 caracteres, y como m�ximo 10.", vbInformation, "Error de Dato"
            Txt_Contrase�a = ""
            Txt_Contrase�a.SetFocus
        End If
        If (Txt_Contrase�a <> Txt_Repetir) Then
            MsgBox "Las Contrase�as registradas son distintas, vuelva a registrarlas.", vbExclamation, "Error de Contrase�a"
            Txt_Contrase�a = ""
            Txt_Repetir = ""
            Txt_Contrase�a.SetFocus
        Else
            cmd_grabar.SetFocus
        End If
    Else
        MsgBox "Debe ingresar primeramente la Contrase�a, y posteriormente repetirla.", vbExclamation, "Error de Contrase�a"
        Txt_Contrase�a = ""
        Txt_Repetir = ""
        Txt_Contrase�a.SetFocus
    End If
End If
End Sub
