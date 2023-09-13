VERSION 5.00
Begin VB.Form Frm_SisNivel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Niveles de Acceso."
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "Frm_SisNivel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7845
   Begin VB.Frame Fra_Nivel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   7575
      Begin VB.TextBox Txt_Descripcion 
         Height          =   285
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
      Begin VB.ComboBox Cmb_Nivel 
         Height          =   315
         ItemData        =   "Frm_SisNivel.frx":0442
         Left            =   1920
         List            =   "Frm_SisNivel.frx":0444
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel de Acceso  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Fra_Menu 
      Caption         =   "  Menú de Acceso  "
      ForeColor       =   &H00800000&
      Height          =   4935
      Left            =   120
      TabIndex        =   26
      Top             =   840
      Width           =   7575
      Begin VB.CheckBox Chk_ContPriPag 
         Caption         =   "Primeros Pagos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   35
         Top             =   4440
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_ContPriUni 
         Caption         =   "Prima Unica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   34
         Top             =   4200
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_Contables 
         Caption         =   "Archivos Contables"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   3960
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Chk_CasosRecal 
         Caption         =   "Casos Recalculados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   32
         Top             =   3720
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_ConPriRecep 
         Caption         =   "Consulta de Primas Recepcionadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   31
         Top             =   2040
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_ArchConf 
         Caption         =   "Archivo de Confirmación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   30
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_Sucursal 
         Caption         =   "Sucursal"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Chk_InfInternos 
         Caption         =   "Informes Internos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   19
         Top             =   3480
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_InfSBS 
         Caption         =   "Informes SBS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   18
         Top             =   3240
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_Informes 
         Caption         =   "Informes Generales"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Chk_ConVistaTasas 
         Caption         =   "Activar Vista de Tasas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   600
         TabIndex        =   21
         Top             =   3000
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_PolBotonRecal 
         Caption         =   "Activar Botón Recalcular y Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   10
         Top             =   840
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox Chk_ConTraspPP 
         Caption         =   "Consulta de Traspasos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   16
         Top             =   2760
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_TraspPP 
         Caption         =   "Traspaso "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   15
         Top             =   2520
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_InfPrimas 
         Caption         =   "Informe de Póliza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   13
         Top             =   1560
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_RecepPrimas 
         Caption         =   "Recepción de Primas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   12
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Chk_ConsPol 
         Caption         =   "Consulta de Pólizas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2760
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Chk_TraPagPen 
         Caption         =   "Traspaso a Pago Pensiones"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Chk_Produccion 
         Caption         =   "Producción"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox Chk_Niveles 
         Caption         =   "Nivel de Acceso"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Chk_Contrasena 
         Caption         =   "Contraseña"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Chk_Usuarios 
         Caption         =   "Usuarios "
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Chk_Analista 
         Caption         =   "Mantenedor de Analista"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox Chk_Calculo 
         Caption         =   "Procesos de Cálculo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Chk_Parametros 
         Caption         =   "Administración de Parámetros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Chk_Sistema 
         Caption         =   "Adm. de Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Chk_Polizas 
         Caption         =   "Pre-Pólizas"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Left            =   120
      TabIndex        =   25
      Top             =   5880
      Width           =   7575
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_SisNivel.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmd_limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3420
         Picture         =   "Frm_SisNivel.frx":0540
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmd_grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2220
         Picture         =   "Frm_SisNivel.frx":0BFA
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm_SisNivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim vlConexion  As ADODB.Connection
'Dim vlRecord    As ADODB.Recordset
Dim vlRecord    As ADODB.Recordset

Dim vlVerificar       As String
Dim vlOperacion       As String
Dim vlNivel           As Integer
Dim vlSw              As Boolean

Dim vlGlsUsuarioCrea As Variant
Dim vlFecCrea As Variant
Dim vlHorCrea As Variant
Dim vlGlsUsuarioModi As Variant
Dim vlFecModi As Variant
Dim vlHorModi As Variant

'--------------------------------------------------------
'                   F U N C I O N E S
'--------------------------------------------------------
'--------------------------------------------------------
'Permite grabar nivel en la base de datos
'--------------------------------------------------------
Function flRegistrarNivel()
On Error GoTo Err_Registrar
    
    Screen.MousePointer = 11
    vlNivel = CLng(Cmb_Nivel)

    'Verificar existencia de Código Nivel para el Ingreso/Actualización
    vgQuery = ""
    vgQuery = "SELECT cod_nivel FROM ma_tpar_nivel WHERE "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' AND "
    vgQuery = vgQuery & "cod_nivel = " & vlNivel & ""
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If (vgRs.EOF) Then
        vlOperacion = "I"
    Else
        vlOperacion = "A"
    End If
    vgRs.Close
    
    If (vlOperacion = "I") Then
        'Ingresar
        vlGlsUsuarioCrea = vgUsuario
        vlFecCrea = Format(Date, "yyyymmdd")
        vlHorCrea = Format(Time, "hhmmss")
        
        vgQuery = ""
        vgQuery = "INSERT INTO ma_tpar_nivel ("
        vgQuery = vgQuery & "cod_sistema,"
        vgQuery = vgQuery & "cod_nivel,"
        vgQuery = vgQuery & "num_menu_1,"
        vgQuery = vgQuery & "num_menu_1_1,num_menu_1_2,"
        vgQuery = vgQuery & "num_menu_1_3,"
        vgQuery = vgQuery & "num_menu_1_4,"
        
        vgQuery = vgQuery & "num_menu_2,"
        vgQuery = vgQuery & "num_menu_2_1,"
        
        vgQuery = vgQuery & "num_menu_3,"
        vgQuery = vgQuery & "num_menu_3_1,"
        vgQuery = vgQuery & "num_menu_3_2,"
        vgQuery = vgQuery & "num_menu_3_2_1,"
        vgQuery = vgQuery & "num_menu_3_2_2,"
        vgQuery = vgQuery & "num_menu_3_2_3,"
        vgQuery = vgQuery & "num_menu_3_2_4,"
        vgQuery = vgQuery & "num_menu_3_3,"
        vgQuery = vgQuery & "num_menu_3_3_1,"
        vgQuery = vgQuery & "num_menu_3_3_2,"
        vgQuery = vgQuery & "num_menu_3_4,"
        vgQuery = vgQuery & "num_menu_3_4_1,"
        vgQuery = vgQuery & "num_menu_3_4_2,"
        vgQuery = vgQuery & "num_menu_3_4_3,"
        vgQuery = vgQuery & "num_menu_3_5,"
        vgQuery = vgQuery & "num_menu_3_5_1,"
        vgQuery = vgQuery & "num_menu_3_5_2,"
        vgQuery = vgQuery & "num_menu_4,"
        vgQuery = vgQuery & "cod_usuariocrea,"
        vgQuery = vgQuery & "fec_crea,"
        vgQuery = vgQuery & "hor_crea "
        If (Txt_Descripcion <> "") Then vgQuery = vgQuery & ",gls_nivel "
        vgQuery = vgQuery & ",ind_boton,ind_ver "
        
        vgQuery = vgQuery & ") VALUES ("
        vgQuery = vgQuery & "'" & vgTipoSistema & "',"
        vgQuery = vgQuery & " " & vlNivel & ", "
        vgQuery = vgQuery & " " & Chk_Sistema.Value & ", "
        vgQuery = vgQuery & " " & Chk_Usuarios.Value & ", "
        vgQuery = vgQuery & " " & Chk_Contrasena.Value & ", "
        vgQuery = vgQuery & " " & Chk_Niveles.Value & ", "
        vgQuery = vgQuery & " " & Chk_Sucursal.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_Parametros.Value & ", "
        vgQuery = vgQuery & " " & Chk_Analista.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_Calculo.Value & ", "
        vgQuery = vgQuery & " " & Chk_Polizas.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_Produccion.Value & ", "
        vgQuery = vgQuery & " " & Chk_RecepPrimas.Value & ", "
        vgQuery = vgQuery & " " & Chk_InfPrimas.Value & ", "
        vgQuery = vgQuery & " " & Chk_ArchConf.Value & ", "
        vgQuery = vgQuery & " " & Chk_ConPriRecep.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_TraPagPen.Value & ", "
        vgQuery = vgQuery & " " & Chk_TraspPP.Value & ", "
        vgQuery = vgQuery & " " & Chk_ConTraspPP.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_InfSBS.Value & ", "
        vgQuery = vgQuery & " " & Chk_InfInternos.Value & ", "
        vgQuery = vgQuery & " " & Chk_CasosRecal.Value & ", "
        vgQuery = vgQuery & " " & Chk_InfInternos.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_Contables.Value & ", "
        vgQuery = vgQuery & " " & Chk_ContPriUni.Value & ", "
        vgQuery = vgQuery & " " & Chk_ContPriPag.Value & ", "
        
        vgQuery = vgQuery & " " & Chk_ConsPol.Value & ", "
        vgQuery = vgQuery & "'" & vlGlsUsuarioCrea & "', "
        vgQuery = vgQuery & "'" & vlFecCrea & "', "
        vgQuery = vgQuery & "'" & vlHorCrea & "' "
        If (Txt_Descripcion <> "") Then vgQuery = vgQuery & ",'" & Txt_Descripcion & "'"
        If (Chk_PolBotonRecal.Value = 1) Then
            vgQuery = vgQuery & ",'S' "
        Else
            vgQuery = vgQuery & ",'N' "
        End If
        If (Chk_ConVistaTasas.Value = 1) Then
            vgQuery = vgQuery & ",'S' "
        Else
            vgQuery = vgQuery & ",'N' "
        End If
        vgQuery = vgQuery & " " & ") "
        vgConexionBD.Execute (vgQuery)
        
        MsgBox "El registro de Datos fue realizado Correctamente", vbInformation, "Información"
'        Call Cmd_Limpiar_Click
    Else
        If (vlOperacion = "A") Then
            'Actualizar
            vgRes = MsgBox("¿ Está seguro que desea Modificar los Niveles de Acceso ?", 4 + 32 + 256, "Operación de Actualización")
            If vgRes <> 6 Then
                Screen.MousePointer = 0
                Exit Function
            End If
            
            vlGlsUsuarioModi = vgUsuario
            vlFecModi = Format(Date, "yyyymmdd")
            vlHorModi = Format(Time, "hhmmss")
            
            vgQuery = "UPDATE ma_tpar_nivel SET "
            vgQuery = vgQuery & "num_menu_1 =" & Chk_Sistema.Value & ", "
            vgQuery = vgQuery & "num_menu_1_1 =" & Chk_Usuarios.Value & ", "
            vgQuery = vgQuery & "num_menu_1_2 =" & Chk_Contrasena.Value & ", "
            vgQuery = vgQuery & "num_menu_1_3 =" & Chk_Niveles.Value & ", "
            vgQuery = vgQuery & "num_menu_1_4 =" & Chk_Sucursal.Value & ", "
  
            vgQuery = vgQuery & "num_menu_2 =" & Chk_Parametros.Value & ", "
            vgQuery = vgQuery & "num_menu_2_1 =" & Chk_Analista.Value & ", "
            
            vgQuery = vgQuery & "num_menu_3 =" & Chk_Calculo.Value & ", "
            vgQuery = vgQuery & "num_menu_3_1 =" & Chk_Polizas.Value & ", "
            vgQuery = vgQuery & "num_menu_3_2 =" & Chk_Produccion.Value & ", "
            vgQuery = vgQuery & "num_menu_3_2_1 =" & Chk_RecepPrimas.Value & ", "
            vgQuery = vgQuery & "num_menu_3_2_2 =" & Chk_InfPrimas.Value & ", "
            vgQuery = vgQuery & "num_menu_3_2_3 =" & Chk_ArchConf.Value & ", "
            vgQuery = vgQuery & "num_menu_3_2_4 =" & Chk_ConPriRecep.Value & ", "
            vgQuery = vgQuery & "num_menu_3_3 =" & Chk_TraPagPen.Value & ", "
            vgQuery = vgQuery & "num_menu_3_3_1 =" & Chk_TraspPP.Value & ", "
            vgQuery = vgQuery & "num_menu_3_3_2 =" & Chk_ConTraspPP.Value & ", "
            vgQuery = vgQuery & "num_menu_3_4 =" & Chk_Informes.Value & ", "
            vgQuery = vgQuery & "num_menu_3_4_1 =" & Chk_InfSBS.Value & ", "
            vgQuery = vgQuery & "num_menu_3_4_2 =" & Chk_InfInternos.Value & ", "
            vgQuery = vgQuery & "num_menu_3_4_3 =" & Chk_CasosRecal.Value & ", "
            
            vgQuery = vgQuery & "num_menu_3_5 =" & Chk_Contables.Value & ", "
            vgQuery = vgQuery & "num_menu_3_5_1 =" & Chk_ContPriUni.Value & ", "
            vgQuery = vgQuery & "num_menu_3_5_2 =" & Chk_ContPriPag.Value & ", "
            
            vgQuery = vgQuery & "num_menu_4 =" & Chk_ConsPol.Value & ", "
            
            vgQuery = vgQuery & "cod_usuariomodi = '" & vlGlsUsuarioModi & "', "
            vgQuery = vgQuery & "fec_modi = '" & vlFecModi & "', "
            vgQuery = vgQuery & "hor_modi = '" & vlHorModi & "' "
            If (Txt_Descripcion <> "") Then
                vgQuery = vgQuery & ",gls_nivel = '" & Txt_Descripcion & "' "
            Else
                vgQuery = vgQuery & ",gls_nivel = Null "
            End If
            
            If (Chk_PolBotonRecal.Value = 1) Then
                vgQuery = vgQuery & ",ind_boton = 'S' "
            Else
                vgQuery = vgQuery & ",ind_boton = 'N' "
            End If
        
            If (Chk_ConVistaTasas.Value = 1) Then
                vgQuery = vgQuery & ",ind_ver = 'S' "
            Else
                vgQuery = vgQuery & ",ind_ver = 'N' "
            End If
            
            vgQuery = vgQuery & "WHERE "
            vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' AND "
            vgQuery = vgQuery & "cod_nivel = " & vlNivel & ""
            vgConexionBD.Execute (vgQuery)
            
            MsgBox "La Actualización de Datos fue realizado Correctamente", vbInformation, "Información"
'            Call Cmd_Limpiar_Click
            
        End If
    End If
    
    If (vlOperacion = "I") Then
        fgComboNivel Cmb_Nivel
    End If

    Screen.MousePointer = 0
    
Exit Function
Err_Registrar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'--------------------------------------------------------
'Permite mostrar datos niveles
'--------------------------------------------------------
Function flMostrarDatos(iNivel As Integer)
On Error GoTo Err_mostrar
    
    'Consulta por nivel
    vgQuery = "SELECT  "
    vgQuery = vgQuery & "num_menu_1,"
    vgQuery = vgQuery & "num_menu_1_1,num_menu_1_2,"
    vgQuery = vgQuery & "num_menu_1_3,"
    vgQuery = vgQuery & "num_menu_1_4,"
    
    vgQuery = vgQuery & "num_menu_2,"
    vgQuery = vgQuery & "num_menu_2_1,"
    
    vgQuery = vgQuery & "num_menu_3,"
    vgQuery = vgQuery & "num_menu_3_1,num_menu_3_2,"
    vgQuery = vgQuery & "num_menu_3_2_1,"
    vgQuery = vgQuery & "num_menu_3_2_2,"
    vgQuery = vgQuery & "num_menu_3_2_3,"
    vgQuery = vgQuery & "num_menu_3_2_4,"
    vgQuery = vgQuery & "num_menu_3_3,"
    vgQuery = vgQuery & "num_menu_3_3_1,"
    vgQuery = vgQuery & "num_menu_3_3_2,"
    vgQuery = vgQuery & "num_menu_3_4, "
    vgQuery = vgQuery & "num_menu_3_4_1, "
    vgQuery = vgQuery & "num_menu_3_4_2, "
    vgQuery = vgQuery & "num_menu_3_4_3, "
    vgQuery = vgQuery & "num_menu_3_5,"
    vgQuery = vgQuery & "num_menu_3_5_1,"
    vgQuery = vgQuery & "num_menu_3_5_2,"
    vgQuery = vgQuery & "num_menu_4 "
    vgQuery = vgQuery & ",gls_nivel "
    vgQuery = vgQuery & ",ind_boton,ind_ver "
    vgQuery = vgQuery & "FROM ma_tpar_nivel WHERE "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' AND "
    vgQuery = vgQuery & "cod_nivel = " & iNivel
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        Chk_Sistema.Value = vgRs!num_menu_1
        Chk_Usuarios.Value = vgRs!num_menu_1_1
        Chk_Contrasena.Value = vgRs!num_menu_1_2
        Chk_Niveles.Value = vgRs!num_menu_1_3
        Chk_Sucursal.Value = vgRs!num_menu_1_4
        
        Chk_Parametros.Value = vgRs!num_menu_2
        Chk_Analista.Value = vgRs!num_menu_2_1
        
        Chk_Calculo.Value = vgRs!num_menu_3
        Chk_Polizas.Value = vgRs!num_menu_3_1
        If (vgRs!ind_boton = "S") Then
            Chk_PolBotonRecal.Value = 1
        Else
            Chk_PolBotonRecal.Value = 0
        End If

        Chk_Produccion.Value = vgRs!num_menu_3_2
        Chk_RecepPrimas.Value = vgRs!num_menu_3_2_1
        Chk_InfPrimas.Value = vgRs!num_menu_3_2_2
        Chk_ArchConf.Value = vgRs!num_menu_3_2_3
        Chk_ConPriRecep.Value = vgRs!num_menu_3_2_4
   
        Chk_TraPagPen.Value = vgRs!num_menu_3_3
        Chk_TraspPP.Value = vgRs!num_menu_3_3_1
        Chk_ConTraspPP.Value = vgRs!num_menu_3_3_2
        
        Chk_Informes.Value = vgRs!num_menu_3_4
        Chk_InfSBS.Value = vgRs!num_menu_3_4_1
        Chk_InfInternos.Value = vgRs!num_menu_3_4_2
        Chk_CasosRecal.Value = vgRs!num_menu_3_4_3
        
        Chk_Contables.Value = vgRs!num_menu_3_5
        Chk_ContPriUni.Value = vgRs!num_menu_3_5_1
        Chk_ContPriPag.Value = vgRs!num_menu_3_5_2
        
        Chk_ConsPol.Value = vgRs!num_menu_4
        
        If (vgRs!ind_ver = "S") Then
            Chk_ConVistaTasas.Value = 1
        Else
            Chk_ConVistaTasas.Value = 0
        End If
        
        If Not IsNull(vgRs!gls_nivel) Then
            Txt_Descripcion = Trim(vgRs!gls_nivel)
        Else
            Txt_Descripcion = ""
        End If
    
    Else
        Txt_Descripcion.Text = ""
        Chk_Sistema.Value = 0
        Chk_Parametros.Value = 0
        Chk_Calculo.Value = 0
        Chk_ConsPol.Value = 0
        
    End If
    vgRs.Close
    Screen.MousePointer = 0

Exit Function
Err_mostrar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function



Private Sub Chk_Calculo_Click()

    If Chk_Calculo.Value = 1 Then
    
       Chk_Polizas.Value = 1
       Chk_Produccion.Value = 1
       Chk_TraPagPen.Value = 1
       Chk_Informes.Value = 1
       Chk_Contables.Value = 1
       
       Chk_Polizas.Enabled = True
       Chk_Produccion.Enabled = True
       Chk_TraPagPen.Enabled = True
       Chk_Informes.Enabled = True
       Chk_Contables.Enabled = True
    
    Else
    
        Chk_Polizas.Value = 0
        Chk_Produccion.Value = 0
        Chk_TraPagPen.Value = 0
        Chk_Informes.Value = 0
        Chk_Contables.Value = 0
        
        Chk_Polizas.Enabled = False
        Chk_Produccion.Enabled = False
        Chk_TraPagPen.Enabled = False
        Chk_Informes.Enabled = False
        Chk_Contables.Enabled = False
    
    End If
    

End Sub

Private Sub Chk_ConsPol_Click()
    If Chk_ConsPol.Value = 1 Then
    
        Chk_ConVistaTasas.Value = 1
        Chk_ConVistaTasas.Enabled = True
    
    Else
    
        Chk_ConVistaTasas.Value = 0
        Chk_ConVistaTasas.Enabled = False
    
    End If
End Sub

Private Sub Chk_Contables_Click()

    If Chk_Contables.Value = 1 Then
    
       Chk_ContPriUni.Value = 1
       Chk_ContPriPag.Value = 1
       
       Chk_ContPriUni.Enabled = True
       Chk_ContPriPag.Enabled = True
    
    Else
    
        Chk_ContPriUni.Value = 0
        Chk_ContPriPag.Value = 0
       
        Chk_ContPriUni.Enabled = False
        Chk_ContPriPag.Enabled = False
    
    End If
    
End Sub

Private Sub Chk_Informes_Click()

    If Chk_Informes.Value = 1 Then
    
       Chk_InfSBS.Value = 1
       Chk_InfInternos.Value = 1
       Chk_CasosRecal.Value = 1
       
       Chk_InfSBS.Enabled = True
       Chk_InfInternos.Enabled = True
       Chk_CasosRecal.Enabled = True
    
    Else
    
        Chk_InfSBS.Value = 0
        Chk_InfInternos.Value = 0
        Chk_CasosRecal.Value = 0
       
        Chk_InfSBS.Enabled = False
        Chk_InfInternos.Enabled = False
        Chk_CasosRecal.Enabled = False
    
    End If

End Sub

Private Sub Chk_Polizas_Click()
    If Chk_Polizas.Value = 1 Then
    
       Chk_PolBotonRecal.Value = 1
       Chk_PolBotonRecal.Enabled = True
    
    Else
    
        Chk_PolBotonRecal.Value = 0
        Chk_PolBotonRecal.Enabled = False
    
    End If
End Sub

Private Sub Chk_Produccion_Click()
    If Chk_Produccion.Value = 1 Then
    
       Chk_RecepPrimas.Value = 1
       Chk_InfPrimas.Value = 1
       Chk_ArchConf.Value = 1
       Chk_ConPriRecep.Value = 1
             
       Chk_RecepPrimas.Enabled = True
       Chk_InfPrimas.Enabled = True
       Chk_ArchConf.Enabled = True
       Chk_ConPriRecep.Enabled = True
    
    Else
    
       Chk_RecepPrimas.Value = 0
       Chk_InfPrimas.Value = 0
       Chk_ArchConf.Value = 0
       Chk_ConPriRecep.Value = 0
             
       Chk_RecepPrimas.Enabled = False
       Chk_InfPrimas.Enabled = False
       Chk_ArchConf.Enabled = False
       Chk_ConPriRecep.Enabled = False
    
    End If
End Sub

Private Sub Chk_Sistema_Click()

    If Chk_Sistema.Value = 1 Then
        Chk_Usuarios.Value = 1
        Chk_Contrasena.Value = 1
        Chk_Niveles.Value = 1
        Chk_Sucursal.Value = 1
        
        Chk_Usuarios.Enabled = True
        Chk_Contrasena.Enabled = True
        Chk_Niveles.Enabled = True
        Chk_Sucursal.Enabled = True
        
       
    Else
        Chk_Usuarios.Value = 0
        Chk_Contrasena.Value = 0
        Chk_Niveles.Value = 0
        Chk_Sucursal.Value = 0
        
        Chk_Usuarios.Enabled = False
        Chk_Contrasena.Enabled = False
        Chk_Niveles.Enabled = False
        Chk_Sucursal.Enabled = False
       
    End If
End Sub

Private Sub Chk_Parametros_Click()

    If Chk_Parametros.Value = 1 Then
        Chk_Analista.Value = 1
        
        Chk_Analista.Enabled = True
    Else
        Chk_Analista.Value = 0
        
        Chk_Analista.Enabled = False
    End If
End Sub

Private Sub Chk_TraPagPen_Click()

    If Chk_TraPagPen.Value = 1 Then
    
       Chk_TraspPP.Value = 1
       Chk_ConTraspPP.Value = 1
       
       Chk_TraspPP.Enabled = True
       Chk_ConTraspPP.Enabled = True
    
    Else
    
        Chk_TraspPP.Value = 0
        Chk_ConTraspPP.Value = 0
       
        Chk_TraspPP.Enabled = False
        Chk_ConTraspPP.Enabled = False
    
    End If

End Sub

Private Sub Cmb_Nivel_Change()
If Not IsNumeric(Cmb_Nivel) Then
    Cmb_Nivel = ""
End If
If (Cmb_Nivel) = "" Then
    Chk_Sistema.Value = 0
    Chk_Parametros.Value = 0
    Chk_Calculo.Value = 0
End If
End Sub

Private Sub Cmb_Nivel_Click()
On Error GoTo Err_Nivel

    If IsNumeric(Cmb_Nivel) Then
        vlNivel = CLng(Cmb_Nivel)
        flMostrarDatos vlNivel
    Else
        Txt_Descripcion = ""
        Chk_Sistema.Value = 0
        Chk_Parametros.Value = 0
        Chk_Calculo.Value = 0
    End If

Exit Sub
Err_Nivel:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmb_Nivel_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Cmb_Nivel) = "") Then
        MsgBox "Debe Ingresar un Valor para el Nivel a Registrar.", vbInformation, "Error de Datos"
    Else
        Txt_Descripcion.SetFocus
    End If
Else
    If Cmb_Nivel <> "" Then
        'Validar que no sobrepase los 20 caracteres
        If Len(Cmb_Nivel) > 2 And KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End If
End Sub

Private Sub Cmb_Nivel_LostFocus()
    Cmb_Nivel_Click
End Sub

Private Sub Cmd_Grabar_Click()
On Error GoTo Err_Grabar
    Cmb_Nivel = Format(Cmb_Nivel, "#0")
    
    'Validar ingreso de Nivel
    If (Not IsNumeric(Cmb_Nivel)) Then
        MsgBox "Debe ingresar el Nivel a registrar.", vbCritical, "Error de Datos"
        Cmb_Nivel.SetFocus
        Exit Sub
    End If
    
    'Validar rangos de Nivel
    If CLng(Cmb_Nivel) < 0 Then
        MsgBox "Debe ingresar un valor mayor a 0 (Cero) para el Nivel a registrar.", vbCritical, "Error de Datos"
        Cmb_Nivel.SetFocus
        Exit Sub
    End If
    If CLng(Cmb_Nivel) > 100 Then
        MsgBox "Debe ingresar un valor menor a 100 (Cien) para el Nivel a registrar.", vbCritical, "Error de Datos"
        Cmb_Nivel.SetFocus
        Exit Sub
    End If
    
    Txt_Descripcion = UCase(Trim(Txt_Descripcion))
    
    Call flRegistrarNivel

Exit Sub
Err_Grabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

'    Txt_Descripcion = ""
    Chk_Sistema.Value = 0
    Chk_Parametros.Value = 0
    Chk_Calculo.Value = 0
    Chk_ConsPol.Value = 0

'    If (Cmb_Nivel.ListCount <> 0) Then
'        Cmb_Nivel.ListIndex = 0
'    End If

Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Salir_Click()
On Error GoTo Err_Salir
    
    Screen.MousePointer = 11
    Unload Me
    Screen.MousePointer = 0

Exit Sub
Err_Salir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Frm_SisNivel.Left = 0
    Frm_SisNivel.Top = 0

    fgComboNivel Cmb_Nivel
    
    If IsNumeric(Cmb_Nivel) Then
        vlNivel = CLng(Cmb_Nivel)
        flMostrarDatos vlNivel
    End If

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Descripcion_GotFocus()
    Txt_Descripcion.SelStart = 0
    Txt_Descripcion.SelLength = Len(Txt_Descripcion)
End Sub

Private Sub Txt_Descripcion_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Txt_Descripcion) <> "") Then
        Txt_Descripcion = UCase(Trim(Txt_Descripcion))
        Chk_Sistema.SetFocus
    End If
End If
End Sub

Private Sub txt_descripcion_LostFocus()
    If (Trim(Txt_Descripcion) <> "") Then
        Txt_Descripcion = UCase(Trim(Txt_Descripcion))
    End If
End Sub
