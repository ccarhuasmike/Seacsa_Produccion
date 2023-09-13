VERSION 5.00
Begin VB.Form Frm_Password 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Sistema de Producción."
   ClientHeight    =   1890
   ClientLeft      =   1050
   ClientTop       =   1470
   ClientWidth     =   5025
   Icon            =   "Frm_Password.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5025
   Begin VB.Frame Fra_Datos 
      Height          =   1815
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         Begin VB.Frame Fra_Datos 
            Height          =   855
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   855
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               Height          =   480
               Left            =   120
               Picture         =   "Frm_Password.frx":0442
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.TextBox TxtPass 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2880
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox TxtLogin 
            Height          =   285
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   2
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "Frm_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vlContador As Integer
'Dim vlRegistro  As RECORSET
Dim vlRegistro  As ADODB.Recordset
'Dim vlRgA As ADODB.Recordset
Dim vlCuenta As Integer
Dim vlIntentos As Integer
Dim fechaSis As String
Dim FechaFin As String
Dim fechaant As String
Dim vlIsApl As Integer
Dim vlBlquea As Integer
Private Sub btn_volver_Click()

    If (vgNivel = 0) Then
        Exit Sub
    End If
        
        'Determinar el Nivel de Acceso al Sistema
        vgSql = "SELECT * FROM MA_TPAR_NIVEL WHERE "
        vgSql = vgSql & "cod_nivel = " & vgNivel & " "
        Set vlRegistro = vgConexionBD.Execute(vgSql)
        If Not vlRegistro.EOF Then
'                    If (vlRegistro!num_menu_1 <> "0") Then
            Frm_Menu.Mnu_AdmSistema.Enabled = True
'                    End If
            If (vlRegistro!num_menu_1_1 = "0") Then
                Frm_Menu.Mnu_SisUsuarios.Enabled = False
            Else
                Frm_Menu.Mnu_SisUsuarios.Enabled = True
            End If
            
            If (vlRegistro!num_menu_1_2 = "0") Then
                Frm_Menu.Mnu_SisContrasena.Enabled = False
            Else
                Frm_Menu.Mnu_SisContrasena.Enabled = True
            End If
            
            If (vlRegistro!num_menu_1_3 = "0") Then
                Frm_Menu.Mnu_SisNivel.Enabled = False
            Else
                Frm_Menu.Mnu_SisNivel.Enabled = True
            End If
            'JTC 11/10/2007
            If (vlRegistro!num_menu_1_4 = "0") Then
                Frm_Menu.Mnu_SisSucursal.Enabled = False
            Else
                Frm_Menu.Mnu_SisSucursal.Enabled = True
            End If
  
            If (vlRegistro!num_menu_2 <> "0") Then
                Frm_Menu.Mnu_AdmParametros.Enabled = True
            Else
                Frm_Menu.Mnu_AdmParametros.Enabled = False
            End If
            If (vlRegistro!num_menu_2_1 = "0") Then
                Frm_Menu.Mnu_AdmApoderado.Enabled = False
            Else
                Frm_Menu.Mnu_AdmApoderado.Enabled = True
            End If
                                
            Frm_Menu.Mnu_Acerca.Enabled = True
        Else
        
            Frm_Menu.Mnu_AdmSistema.Enabled = True
            Frm_Menu.Mnu_AdmParametros.Enabled = False
            Frm_Menu.Mnu_ProcGeneracion.Enabled = False
            Frm_Menu.Mnu_ProcConsulta.Enabled = False
'            Frm_Menu.Mnu_Notebook.Enabled = False
            Frm_Menu.Mnu_Acerca.Enabled = True
            
            'vgPertenece = "L"
            'MDIPrincipal.Toolbar1.Enabled = False
        End If
        vlRegistro.Close
        Frm_Password.Hide
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

'    vgDsn = ""
'    vgNombreBaseDatos = ""
'    vgMensaje = ""
'    vgRutaArchivo = ""
'
'    'Valida Si Existe Archivo de AdmBasDat.Inicio
'    vgRutaArchivo = App.Path & "\VidaTrad.ini"
'    If Not fgExiste(vgRutaArchivo) Then
'        MsgBox "No existe el Archivo de Parámetros Nestle.ini para ejecutar la Aplicación.", vbCritical, "Ejecución Cancelada"
'        End
'    End If
'
'    lpFileName = vgRutaArchivo
'    lpAppName = "Conexion"
'    lpDefault = ""
'    lpReturnString = Space$(128)
'    Size = Len(lpReturnString)
'    lpKeyName = ""
'
''    'Valida Si Existe Nombre de la entrada para definir al Servidor
''    vgNombreServidor = fgGetPrivateIni(lpAppName, "Servidor", lpFileName)
''    If (vgNombreServidor = "DESCONOCIDO") Then
''        vgMensaje = "La Entrada 'Servidor', no está definida en el Archivo AdmBD.ini" & vbCrLf
''    End If
'
'    'Valida Si Existe Nombre de la entrada para definir Base de Datos SisSin
'    vgNombreBaseDatos = fgGetPrivateIni(lpAppName, "BasedeDatos", lpFileName)
'    If (vgNombreBaseDatos = "DESCONOCIDO") Then
'        vgMensaje = "La Entrada 'Base de Datos', no está definida en el Archivo AdmBasDat.ini" & vbCrLf
'    End If
'
''    'Valida Si Existe Nombre de la entrada para definir Usuario de SisSin
''    vgNombreUsuario = fgGetPrivateIni(lpAppName, "Usuario", lpFileName)
''    If (vgNombreUsuario = "DESCONOCIDO") Then
''        vgMensaje = "La Entrada 'Usuario', no está definida en el Archivo AdmBasDat.ini" & vbCrLf
''    End If
'
''    'Valida Si Existe Nombre de la entrada para definir Password de SisSin
''    vgPassWord = fgGetPrivateIni(lpAppName, "Password", lpFileName)
''    'vgPassWord = ""
''    If (vgPassWord = "DESCONOCIDO") Then
''        vgMensaje = "La Entrada 'PassWord', no está definida en el Archivo AdmBasDat.ini" & vbCrLf
''    End If
'
'    'Valida Si Existe Nombre de la entrada para definir DSN de SisSin
'    vgDsn = fgGetPrivateIni(lpAppName, "DSN", lpFileName)
'    If (vgDsn = "DESCONOCIDO") Then
'        vgMensaje = "La Entrada 'DSN', no está definida en el Archivo AdmBasDat.ini" & vbCrLf
'    End If
'
'    If (vgMensaje <> "") Then
'        MsgBox "Status de los Datos de Inicio" & vbCrLf & vbCrLf & vgMensaje & vbCrLf & vbCrLf & "Proceso Cancelado." & vbCrLf & "Se deben Ingresar todos los datos Básicos."
'        Exit Sub
'        End
'    End If
'
'    vgRutaBasedeDatos = LeeArchivoIni("Conexion", "Ruta", "", App.Path & "\VidaTradBD.Ini")
'    vgRutaBasedeDatos = vgRutaBasedeDatos & LeeArchivoIni("Conexion", "BasedeDatos", "", App.Path & "\VidaTradBD.Ini")
'    AbrirBaseAccess (vgRutaBasedeDatos)
    
'''**********************Borrar******************
'TxtLogin.Text = "Paranab"
'TxtPass.Text = "solange"

TxtLogin.Text = "CLAROSA"
TxtPass.Text = "Protecta%6"
''**********************Borrar******************
    
    Call Center(Frm_Password)
    
    'Inicialización de Variables
    vgFechaSistema = ""
    vgFechaCalculo = ""
    vgLogin = ""
    'vgRut = ""
    'vgCliente = "1"
    'vgNombreCliente = "Servicios Actuariales S.A."
    'vgPassword = ""
    vgContraseña = ""
    vlContador = 0
    
    'Cmd_Cancelar.Visible = False
    'Cmd_Aceptar.Left = 1320
    'Cmd_Salir.Left = 2700
    

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Descargar
Dim x%

    x% = MsgBox("¿ Está Seguro que desea Salir del Sistema ?", 32 + 4, "Salir")
    If x% = 6 Then
        End
    Else
        Cancel = 1
    End If

Exit Sub
Err_Descargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub TxtLogin_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
   If (Trim(TxtLogin.Text) = "") Then
      MsgBox "Debe ingresar su Login o Nombre de Usuario", vbInformation, "Advertencia"
   Else
      TxtPass.SetFocus
   End If
End If
End Sub

Private Sub TxtLogin_LostFocus()
TxtLogin = UCase(TxtLogin)
TxtPass.SelStart = 0
TxtPass.SelLength = Len(TxtPass)
End Sub

Private Sub TxtPass_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Acceso



If (KeyAscii = 13) Then
  If (Trim(TxtPass.Text) = "") Then
        MsgBox "Debe Ingresar su Clave o Password de Acceso.", vbInformation, "..Advertencia.."
        
    Else
        TxtLogin = UCase(Trim(TxtLogin))
        TxtPass = UCase(Trim(TxtPass))
        
'----------------------------------------------
        If (TxtLogin <> "") And (TxtPass <> "") Then
            vgNivel = 0
            vgLogin = TxtLogin
            'vgPassword = fgEncPassword(TxtPass)
            vgContraseña = fgEncPassword(TxtPass)
            'vgContraseña = vgPassword
                
'I--- ABV 05/07/2007 ---
            vgNivelIndicadorVer = "N"
            vgNivelIndicadorBoton = "N"
'F--- ABV 05/07/2007 ---
            'Call AbrirBaseDeDatos(vgRutaBasedeDatos)
                
            'Abrir la Conexión a la Base de Datos
            If Not fgConexionBaseDatos(vgConexionBD) Then
                MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
                Exit Sub
            End If
                           
            vgSql = "select * from MA_TMAE_USUMODULO where cod_usuario = '" & vgLogin & "'  and cod_sistema = '" & vgTipoSistema & "'"
            Set vgRs = vgConexionBD.Execute(vgSql)
                         
            If vgRs.EOF Then
               MsgBox "El usuario no tiene permisos para acceder a este sistema.", vbCritical, "Usuario no admitido"
               TxtLogin = ""
               TxtPass = ""
               TxtLogin.SetFocus
               Exit Sub
            End If
            vgRs.Close
                     
'            vgSql = "SELECT cod_usuario,cod_nivel FROM MA_TMAE_USUARIO WHERE "
'            vgSql = vgSql & "cod_sistema = '" & vgTipoSistema & "' AND "
'            vgSql = vgSql & "cod_usuario = '" & vgLogin & "' AND "
'            vgSql = vgSql & "GLS_PASsword = '" & vgContraseña & "'"
            
            vgSql = " select  A.COD_USUARIO,cod_nivel FROM MA_TMAE_USUMATRIZ A"
            vgSql = vgSql & " join MA_TMAE_USUPASSWORD B ON A.COD_USUARIO=B.COD_USUARIO"
            vgSql = vgSql & " join MA_TMAE_USUMODULO C ON A.COD_USUARIO=C.COD_USUARIO"
            vgSql = vgSql & " WHERE A.cod_usuario = '" & vgLogin & "' AND NRO_USUPASS=(select max(nro_usupass) from MA_TMAE_USUPASSWORD WHERE cod_usuario = A.COD_USUARIO)"
            vgSql = vgSql & " AND cod_sistema = '" & vgTipoSistema & "' AND cod_estado='A'" ' " AND GLS_PASsword = '" & vgContraseña & "' "
            
            Set vgRs = vgConexionBD.Execute(vgSql)
            If Not vgRs.EOF Then
                'vgNombreUsuario = vgRs!gls_nomusu & " "
                'vgNombreUsuario = vgusuarioLogin & vgRs!gls_patusu & " "
                'vgNombreUsuario = vgusuarioLogin & vgRs!gls_matusu
                vgRutUsuario = vgRs!COD_USUARIO
                vgUsuario = vgRs!COD_USUARIO
                vgNivel = vgRs!cod_nivel
            Else
                MsgBox "Usuario o password incorrectos o el Usuario se encuentra cesado. Comuniquese con el admistrador del sistema de Rentas Vitalicias.", vbCritical, "Error de Password"
                If usuariopa <> "" Then
                    If TxtLogin <> usuariopa Then vlContador = 0
                End If
                vlContador = vlContador + 1
                vgLogin = ""
                vgContraseña = ""
                'vgPassword = ""
                If vlContador >= vgIntentos Then
                    'RRR 18/01/2012
                    MsgBox "El usuario será bloqueado por exceder el número de intentos permitidos.", vbCritical, "Operación de Cierre"
                    vgSql = "update MA_TMAE_USUMATRIZ set isblock=1 where cod_usuario='" & Trim(TxtLogin) & "'"
                    vgConexionBD.Execute (vgSql)
                    Call FgGuardaLog("Bloqueo de usuario de sistema de Producciòn", vgUsuario, "1501")
                    vlContador = 0
                    'RRR
                End If
                TxtLogin.SetFocus
            End If
            vgRs.Close
                
            'Call CerrarBaseDeDatos(vgConexionBD)

            '------ Activar después de Actualización de Menú -----------
            '     Frm_Password.Hide
            ''I------ ABV 21/01/2004 -----
'
                ' RRR 18/01/2012
            
            Dim fechasisDate, fechafindate, fechaantdate As Date
                
            vgSql = " SELECT a.fec_finpass, a.fec_antpass, a.ind_segu, b.isblock FROM MA_TMAE_USUPASSWORD a"
            vgSql = vgSql & " join MA_TMAE_USUMATRIZ b on a.cod_usuario=b.cod_usuario"
            vgSql = vgSql & " WHERE a.COD_USUARIO='" & vgUsuario & "'"
            vgSql = vgSql & " AND NRO_USUPASS =(SELECT MAX(NRO_USUPASS) FROM MA_TMAE_USUPASSWORD WHERE COD_USUARIO='" & vgUsuario & "')"
                        
            Set vgRs = vgConexionBD.Execute(vgSql)
                
                If Not vgRs.EOF Then
                vlIsApl = vgRs!ind_segu
                
                    If vlIsApl <> 0 Then
                        vlBlquea = vgRs!isblock
                        
                        If vlBlquea = 1 Then
                            MsgBox "Usuario se encuentra bloqueado. Consulte con el Administrador.", vbCritical, "PASSWORD"
                            Exit Sub
                        End If
                        
                        
                        FechaFin = Mid(vgRs!fec_finpass, 1, 4) & "/" & Mid(vgRs!fec_finpass, 5, 2) & "/" & Mid(vgRs!fec_finpass, 7, 2)
                        fechaant = Mid(vgRs!fec_antpass, 1, 4) & "/" & Mid(vgRs!fec_antpass, 5, 2) & "/" & Mid(vgRs!fec_antpass, 7, 2)
                        fechaSis = Mid(CStr(Now), 1, 2) & "/" & Mid(CStr(Now), 4, 2) & "/" & Mid(CStr(Now), 7, 4)
                        
                        fechasisDate = CDate(fechaSis)
                        fechafindate = CDate(FechaFin)
                        fechaantdate = CDate(fechaant)
                        
                        Dim vlFaltan As Long
                        
                        vlFaltan = DateDiff("d", Now, fechafindate)
                        
                        If vgChkdiaant = 1 Then
                          Select Case vlFaltan
                            Case Is > 0
                                If vgDiasFaltan >= vlFaltan Then
                                     vgRes = MsgBox("Su contraseña esta por caducar en " & CStr(vlFaltan) & " dias , ¿Desea Crear una nueva?", 4 + 32 + 256, "Operación de Actualización")
                                    If vgRes <> 6 Then
                                        Screen.MousePointer = 0
                                        GoTo continua
                                    End If
                                    vgValorAr = 0
                                    Frm_SisContrasena.Show
                                End If
                            Case Is < 0
                                vgRes = MsgBox(" La contraseña ha caducado, ¿Desea Crear una nueva?", 4 + 32 + 256, "Operación de Actualización")
                                If vgRes <> 6 Then
                                    Screen.MousePointer = 0
                                    'vgSql = "update MA_TMAE_USUMATRIZ set isblock=1 where cod_usuario='" & Trim(TxtLogin) & "'"
                                    'vgConexionBD.Execute (vgSql)
                                    End
                                End If
                                vgValorAr = 1
                                Frm_SisContrasena.Show
                            Case Is = 0
                                vgRes = MsgBox("Su contraseña ha caducado el día de hoy, ¿Desea Crear una nueva?", 4 + 32 + 256, "Operación de Actualización")
                                If vgRes <> 6 Then
                                    Screen.MousePointer = 0
                                    GoTo continua
                                End If
                                vgValorAr = 0
                                Frm_SisContrasena.Show
                            End Select
                        End If
                        If ValidaClave(TxtPass) <> 99 Then
                            vgRes = MsgBox("Su contraseña no es valida. ¿Desea Crear una nueva?", 4 + 32 + 256, "Operación de Actualización")
                            If vgRes <> 6 Then
                                Screen.MousePointer = 0
                                TxtPass.SetFocus
                                'GoTo continua
                                Exit Sub
                            End If
                            vgValorAr = 0
                            Frm_SisContrasena.Show
                        End If
                    End If
                End If
continua:
            ' RRR
            
            'JEVC CORPTEC 24/07/2017
            Call fgLogIn_Pro
            
            If (vgNivel <> 0) Then

                'vgFechaSistema = Format(Date, "dd/mm/yyyy")
                'vgFechaCalculo = Format(Date, "dd/mm/yyyy")

                ''Determina si el Acceso a un mantenedor está Denegado
                'vgPertenece = "N" 'Permite el acceso en ToolBar
                ''vgPertenece = "L" 'No Permite el acceso en ToolBar

                'Call AbrirBaseDeDatos(vgRutaBasedeDatos)

                'Abrir la Conexión a la Base de Datos
                'If Not AbrirBaseDeDatos(vgConexionBD) Then
                '    MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
                '    Exit Sub
                'End If

                'Determinar el Nivel de Acceso al Sistema
                vgSql = "SELECT "
                vgSql = vgSql & "num_menu_1,"
                vgSql = vgSql & "num_menu_1_1,"
                vgSql = vgSql & "num_menu_1_2,"
                vgSql = vgSql & "num_menu_1_3,"
                vgSql = vgSql & "num_menu_1_4,"

                vgSql = vgSql & "num_menu_2,"
                vgSql = vgSql & "num_menu_2_1,"

                vgSql = vgSql & "num_menu_3,"
                vgSql = vgSql & "num_menu_3_1,"
                
                vgSql = vgSql & "num_menu_3_2,"
                vgSql = vgSql & "num_menu_3_2_1,"
                vgSql = vgSql & "num_menu_3_2_2,"
                vgSql = vgSql & "num_menu_3_2_3,"
                vgSql = vgSql & "num_menu_3_2_4,"
                
                vgSql = vgSql & "num_menu_3_3,"
                vgSql = vgSql & "num_menu_3_3_1,"
                vgSql = vgSql & "num_menu_3_3_2,"
                
                vgSql = vgSql & "num_menu_3_4, "
                vgSql = vgSql & "num_menu_3_4_1, "
                vgSql = vgSql & "num_menu_3_4_2, "
                vgSql = vgSql & "num_menu_3_4_3, "
                
                vgSql = vgSql & "num_menu_3_5, "
                vgSql = vgSql & "num_menu_3_5_1, "
                vgSql = vgSql & "num_menu_3_5_2, "
                
                vgSql = vgSql & "num_menu_4 "

'I--- ABV 05/07/2007 ---
                vgSql = vgSql & ",ind_ver,ind_boton "
'F--- ABV 05/07/2007 ---
                vgSql = vgSql & "FROM MA_TPAR_NIVEL WHERE "
                vgSql = vgSql & "cod_sistema = '" & vgTipoSistema & "' AND "
                vgSql = vgSql & "cod_nivel = " & vgNivel & " "
                Set vlRegistro = vgConexionBD.Execute(vgSql)
                If Not vlRegistro.EOF Then
                    If (vlRegistro!num_menu_1 <> "0") Then
                        Frm_Menu.Mnu_AdmSistema.Enabled = True
                    End If
                    If (vlRegistro!num_menu_1_1 = "0") Then
                        Frm_Menu.Mnu_SisUsuarios.Enabled = False
                    Else
                        Frm_Menu.Mnu_SisUsuarios.Enabled = True
                    End If
                    If (vlRegistro!num_menu_1_2 = "0") Then
                        Frm_Menu.Mnu_SisContrasena.Enabled = False
                    Else
                        Frm_Menu.Mnu_SisContrasena.Enabled = True
                    End If
                    If (vlRegistro!num_menu_1_3 = "0") Then
                        Frm_Menu.Mnu_SisNivel.Enabled = False
                    Else
                        Frm_Menu.Mnu_SisNivel.Enabled = True
                    End If
                    If (vlRegistro!num_menu_1_4 = "0") Then
                        Frm_Menu.Mnu_SisSucursal.Enabled = False
                    Else
                        Frm_Menu.Mnu_SisSucursal.Enabled = True
                    End If
    
                    If (vlRegistro!num_menu_2 <> "0") Then
                        Frm_Menu.Mnu_AdmParametros.Enabled = True
                    Else
                        Frm_Menu.Mnu_AdmParametros.Enabled = False
                    End If
                    If (vlRegistro!num_menu_2_1 = "0") Then
                        Frm_Menu.Mnu_AdmApoderado.Enabled = False
                    Else
                        Frm_Menu.Mnu_AdmApoderado.Enabled = True
                    End If
                    
                    If (vlRegistro!num_menu_3 <> "0") Then
                        Frm_Menu.Mnu_ProcGeneracion.Enabled = True
                    Else
                        Frm_Menu.Mnu_ProcGeneracion.Enabled = False
                    End If
                    
                    If (vlRegistro!num_menu_3_1 = "0") Then
                        Frm_Menu.Mnu_ProcPoliza.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcPoliza.Enabled = True
                    End If
                    
                    If (vlRegistro!num_menu_3_2 = "0") Then
                        Frm_Menu.Mnu_ProcProduccion.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcProduccion.Enabled = True
                    End If
                    
                    If (vlRegistro!num_menu_3_2_1 = "0") Then
                        Frm_Menu.Mnu_ProcProPrima.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcProPrima.Enabled = True
                    End If
                    If (vlRegistro!num_menu_3_2_2 = "0") Then
                        Frm_Menu.Mnu_ProcProInformes.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcProInformes.Enabled = True
                    End If
                    If (vlRegistro!num_menu_3_2_3 = "0") Then
                        Frm_Menu.Mnu_ProcProArchivo.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcProArchivo.Enabled = True
                    End If
                    If (vlRegistro!num_menu_3_2_4 = "0") Then
                        Frm_Menu.Mnu_ProcProConsulta.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcProConsulta.Enabled = True
                    End If
        
                    If (vlRegistro!num_menu_3_3 = "0") Then
                        Frm_Menu.Mnu_ProcTraspaso.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcTraspaso.Enabled = True
                    End If
                    If (vlRegistro!num_menu_3_3_1 = "0") Then
                        Frm_Menu.Mnu_ProcTraPago.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcTraPago.Enabled = True
                    End If
                    If (vlRegistro!num_menu_3_3_2 = "0") Then
                        Frm_Menu.Mnu_ProcTraConsulta.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcTraConsulta.Enabled = True
                    End If
                    
                    If (vlRegistro!num_menu_3_4 = "0") Then
                        Frm_Menu.Mnu_ProcInforme.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcInforme.Enabled = True
                    End If
                    If (vlRegistro!num_menu_3_4_1 = "0") Then
                        Frm_Menu.Mnu_ProcInformeSBS.Enabled = False
                    Else
                         Frm_Menu.Mnu_ProcInformeSBS.Enabled = True
                    End If
                    If (vlRegistro!num_menu_3_4_2 = "0") Then
                        Frm_Menu.Mnu_ProcInformeInt.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcInformeInt.Enabled = True
                    End If
                    If (vlRegistro!num_menu_3_4_3 = "0") Then
                        Frm_Menu.Mnu_ProcInformeRec.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcInformeRec.Enabled = True
                    End If
                    
                    If (vlRegistro!num_menu_3_5 = "0") Then
                        Frm_Menu.Mnu_ArchContable.Enabled = False
                    Else
                        Frm_Menu.Mnu_ArchContable.Enabled = True
                    End If
                    If (vlRegistro!num_menu_3_5_1 = "0") Then
                        Frm_Menu.Mnu_ArchConPriUni.Enabled = False
                    Else
                         Frm_Menu.Mnu_ArchConPriUni.Enabled = True
                    End If
                    If (vlRegistro!num_menu_3_5_2 = "0") Then
                        Frm_Menu.Mnu_ArchConPriPag.Enabled = False
                    Else
                        Frm_Menu.Mnu_ArchConPriPag.Enabled = True
                    End If
                   
                    If (vlRegistro!num_menu_4 = "0") Then
                        Frm_Menu.Mnu_ProcConsulta.Enabled = False
                    Else
                        Frm_Menu.Mnu_ProcConsulta.Enabled = True
                    End If

                    Frm_Menu.Mnu_Acerca.Enabled = True

'I--- ABV 05/07/2007 ---
                    If Not IsNull(vlRegistro!ind_ver) Then
                        If (vlRegistro!ind_ver = "N") Then
                            vgNivelIndicadorVer = "N"
                        Else
                            vgNivelIndicadorVer = "S"
                        End If
                    End If
                    If Not IsNull(vlRegistro!ind_boton) Then
                        If (vlRegistro!ind_boton = "N") Then
                            vgNivelIndicadorBoton = "N"
                        Else
                            vgNivelIndicadorBoton = "S"
                        End If
                    End If
'F--- ABV 05/07/2007 ---

                    'vgPertenece = "L"
                    'MDIPrincipal.Toolbar1.Enabled = False
                    Frm_Menu.Mnu_Salir.Enabled = True
                End If
                vlRegistro.Close

                'Call CerrarBaseDeDatos

                'Cerrar la Conexión
                'Call CerrarBaseDeDatos(vgConexionBD)

                Call fgApoderado

                'Buscar Monedas de Conversión
                Call fgBuscarMonedaOfiTran(vgMonedaCodOfi, vgMonedaCodTran)
                Call FgGuardaLog("Logueo al sistema de Producciòn", vgUsuario, "1500")
                Frm_Password.Hide
                'vgPertenece = "N" 'Funciona en forma Normal

                ''VGFechaTope = Mid(VGFechaTope, 1, 2) + "/" + Mid(VGFechaTope, 4, 2) + "/" + Trim(Str((Val(Mid(VGFechaTope, 7, 4)) + 100)))
                ''FrmIngresoFechas.MkSistema = vgFechaSistema
                ''FrmIngresoFechas.MkCalculo = vgFechaCalculo
                ''FrmIngresoFechas.MkTope = vgFechaTope
                'FrmIngresoFechas.Show
            End If
            
        End If

'----------------------------------------------
        'Cmd_Aceptar.SetFocus
    End If
End If

Exit Sub
Err_Acceso:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
Public Function ValidaClave(password As String) As Integer

 Dim i, l, n, a As Integer
 Dim car As String
 Dim aplicavalidacion As Integer

 ValidaClave = 99

 vgSql = "SELECT * FROM MA_TMAE_ADMINCUENTAS WHERE "
 vgSql = vgSql & "cod_cliente = '1' "
 Set vgRs = vgConexionBD.Execute(vgSql)

    If Not vgRs.EOF Then
        tammin = vgRs!ntamañomin
        cantclvant = vgRs!ncanclvant
        canmincaralf = vgRs!ncaracmin
        freccambio = vgRs!nfrecuencia
        canantclv = vgRs!ncanclvant
        balfanum = vgRs!balfanum
    End If
    
    
    
    For i = 1 To Len(password)
    
        car = Mid(password, i, 1)
    
        If VLetras(Asc(car)) <> 0 Then l = l + 1
        If Numeros(Asc(car)) <> 0 Then n = n + 1
        If VAlfanumerico(Asc(car)) <> 0 Then a = a + 1
        
    Next
    
    If Len(password) < tammin Then
        'MsgBox "Password debe ser minimo de " & CStr(tammin) & " caracteres ", vbCritical, "Error de Datos"
        'strMensaje = "Password debe ser minimo de " & CStr(tammin) & " caracteres "
        ValidaClave = 1
        Exit Function
    End If
    
    If balfanum = 1 Then
        If a < canmincaralf Then
            'MsgBox "La clave debe contener como minimo " & canmincaralf & " caracteres alfanumericos.", vbCritical, "Error de Datos"
            ValidaClave = 2
            Exit Function
        End If
    Else
        If a > 0 Then
            'MsgBox "La clave no debe contener caracteres alfanumericos.", vbCritical, "Error de Datos"
            ValidaClave = 3
            Exit Function
        End If
    End If

    'ValidaClave = 1
End Function

Private Sub TxtPass_LostFocus()
TxtLogin.SelStart = 0
TxtLogin.SelLength = Len(TxtLogin)
'TxtLogin.SelText = TxtLogin.SelLength
End Sub
