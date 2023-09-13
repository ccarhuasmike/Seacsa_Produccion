VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm Frm_Menu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Producción de Seguros Previsionales."
   ClientHeight    =   7575
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   12075
   Icon            =   "Frm_Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar Stb_Barra 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1667
            MinWidth        =   1658
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3704
            MinWidth        =   3704
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "11/09/2023"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "11:57 a.m."
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnu_AdmSistema 
      Caption         =   "Administración del Sistema"
      Begin VB.Menu Mnu_SisUsuarios 
         Caption         =   "&Usuarios"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_SisContrasena 
         Caption         =   "&Contraseña"
      End
      Begin VB.Menu Mnu_SisNivel 
         Caption         =   "&Nivel de Acceso"
      End
      Begin VB.Menu Mnu_SisSucursal 
         Caption         =   "&Sucursal"
      End
      Begin VB.Menu Mnu_Separar8 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_AdmParametros 
      Caption         =   "Administración de Parámetros"
      Begin VB.Menu Mnu_AdmApoderado 
         Caption         =   "Mantenedor de &Analista"
      End
   End
   Begin VB.Menu Mnu_ProcGeneracion 
      Caption         =   "Procesos de Cálculo"
      Begin VB.Menu Mnu_ProcPoliza 
         Caption         =   "&Pre-Pólizas"
         Begin VB.Menu Mnu_ProcPolIngreso 
            Caption         =   "Mantenedor Pre-Pólizas"
         End
         Begin VB.Menu Mnu_ProcPolAFP 
            Caption         =   "Antecedentes de Prima AFP"
         End
      End
      Begin VB.Menu Mnu_ProcProduccion 
         Caption         =   "P&roducción"
         Begin VB.Menu Mnu_ProcProPrima 
            Caption         =   "Recepción de Primas"
         End
         Begin VB.Menu Mnu_ProcProInformes 
            Caption         =   "Informe de Póliza"
         End
         Begin VB.Menu mnu_kitImpresion 
            Caption         =   "Kit de Impresión"
         End
         Begin VB.Menu Mnu_ProcProArchivo 
            Caption         =   "Archivo de Confirmación"
         End
         Begin VB.Menu Mnu_Separar7 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_ProcProConsulta 
            Caption         =   "Consulta de Primas Recepcionadas"
         End
         Begin VB.Menu Mnu_Separar11 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_ProcImprimirMas 
            Caption         =   "Impresion Masiva de Poliza"
         End
         Begin VB.Menu Mnu_ProcRegCargos 
            Caption         =   "Registro de Cargos Varios"
         End
      End
      Begin VB.Menu Mnu_Separar12 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_ProcTraspaso 
         Caption         =   "Tra&spaso a Pago Pensiones"
         Begin VB.Menu Mnu_ProcTraPago 
            Caption         =   "Traspaso"
         End
         Begin VB.Menu Mnu_Separar10 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_ProcTraConsulta 
            Caption         =   "Consulta de Traspasos"
         End
      End
      Begin VB.Menu Mnu_Separar9 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_ProcInforme 
         Caption         =   "Informes Generales"
         Begin VB.Menu Mnu_ProcInformeSBS 
            Caption         =   "Infomes SBS"
         End
         Begin VB.Menu Mnu_ProcInformeInt 
            Caption         =   "Informes Internos"
         End
         Begin VB.Menu Mnu_Separar6 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_ProcInformeRec 
            Caption         =   "Casos Recalculados"
         End
      End
      Begin VB.Menu Mnu_Separar1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_ArchContable 
         Caption         =   "Archivos Contables"
         Begin VB.Menu Mnu_ArchConPriUni 
            Caption         =   "Prima Unica"
         End
         Begin VB.Menu Mnu_ArchConPriPag 
            Caption         =   "Primeros Pagos"
         End
      End
   End
   Begin VB.Menu Mnu_ProcConsulta 
      Caption         =   "C&onsulta de Pólizas"
   End
   Begin VB.Menu mnu_report 
      Caption         =   "Reportes"
      Begin VB.Menu mnu_replog 
         Caption         =   "Reporte de Log"
      End
   End
   Begin VB.Menu Mnu_Acerca 
      Caption         =   "Acerca de ..."
   End
   Begin VB.Menu mnu_gestor 
      Caption         =   "Gestor Cliente"
   End
   Begin VB.Menu Mnu_Salir 
      Caption         =   "S&alir"
   End
End
Attribute VB_Name = "Frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()

    If (vgNombreCortoCompania <> "") Then
        Frm_Menu.Caption = vgNombreCortoCompania & " - " & Frm_Menu.Caption
    End If
    
    Stb_Barra.Panels.Item(1).Text = "Sistema :  " & vgNombreSubSistema & "  " 'vgNombreCortoCompania
    'Stb_Barra.Panels.Item(2).Text = "BD : " & UCase(vgRutaBasedeDatos) & Space(2)
    Stb_Barra.Panels.Item(2).Text = "BD : " & vgNombreBaseDatos & Space(2)
    'Stb_Barra.Panels.Item(4).Text = "Cliente : " & UCase(vgnombrecliente)

On Error GoTo Siguiente
    
    Me.Picture = LoadPicture(App.Path & "\ModuloProduccion.bmp")

Siguiente:
    If Err.Number <> 0 Then
        MsgBox "No se encontró el archivo: " & App.Path & "\Logo.bmp" & Chr(13) & "Se continuará con la carga del Sistema", vbExclamation
    End If
End Sub

'Private Sub MDIForm_Unload(Cancel As Integer)
'Dim X%
'    X% = MsgBox("¿ Está Seguro que desea Salir del Sistema?", 32 + 4, "Salir")
'    If X% = 6 Then
'        'JEVC CORPTEC 24/07/2017
'        Call fgLogOut
'        Call FgGuardaLog("Salida de sistema de Producciòn", vgUsuario, "1501")
'        End
'    Else
'        Cancel = 1
'    End If
'End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Or UnloadMode = 1 Then
        Dim x%
        x% = MsgBox("¿ Está Seguro que desea Salir del Sistema?", 32 + 4, "Salir")
        If x% = 6 Then
            'JEVC CORPTEC 24/07/2017
            Call fgLogOut_Pro
            'Call FgGuardaLog("Salida de sistema de Producciòn", vgUsuario, "1501")
            End
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub Mnu_Acerca_Click()
    Screen.MousePointer = 11
    Frm_AcercaDe.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_AdmApoderado_Click()
    Screen.MousePointer = 11
    Frm_SisApoderado.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ArchConPriPag_Click()
    Screen.MousePointer = 11
    ''Frm_ArchContable.Show
    Frm_ContableArch.flInicio ("PriPag")
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ArchConPriUni_Click()
    Screen.MousePointer = 11
    ''Frm_ArchContable.Show
    Frm_ContableArch.flInicio ("PriUni")
    Screen.MousePointer = 0
End Sub

Private Sub mnu_gestor_Click()

'FrmGestorCliente.Show
 Frm_StockFirmas.Show
'FrmGeneraJson.Show

End Sub

Private Sub mnu_kitImpresion_Click()
    Screen.MousePointer = 11
    Frm_KitImpresion.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcConsulta_Click()
    Screen.MousePointer = 11
    Frm_CalConsulta.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcImprimirMas_Click()
    Screen.MousePointer = 11
    Frm_ImpresionMasiva.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcInformeInt_Click()
    Screen.MousePointer = 11
    Frm_CalInfInt.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcInformeRec_Click()
    Screen.MousePointer = 11
    Frm_CalInfRecal.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcInformeSBS_Click()
    Screen.MousePointer = 11
    Frm_CalInfSBS.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcPolAFP_Click()
    Screen.MousePointer = 11
    Frm_CalPrimaAFP.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcPolIngreso_Click()
    Screen.MousePointer = 11
    Frm_CalPoliza.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcProArchivo_Click()
    Screen.MousePointer = 11
    Frm_CalPrimaArc.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcProConsulta_Click()
    Screen.MousePointer = 11
    Frm_CalPriConsulta.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcProInformes_Click()
    Screen.MousePointer = 11
    Frm_CalPrimaInf.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcProPrima_Click()
    Screen.MousePointer = 11
    Frm_CalPrima.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcRegCargos_Click()
    Screen.MousePointer = 11
    Frm_RecalculoPol.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcTraConsulta_Click()
    Screen.MousePointer = 11
    Frm_CalTraConsulta.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_ProcTraPago_Click()
    Screen.MousePointer = 11
    Frm_CalTraspaso.Show
    Screen.MousePointer = 0
End Sub

Private Sub mnu_replog_Click()
   Screen.MousePointer = 11
    frm_reportelog.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_Salir_Click()
Dim x%
    x% = MsgBox("¿ Está Seguro que desea Salir del Sistema?", 32 + 4, "Salir")
    If x% = 6 Then
        'JEVC CORPTEC 24/07/2017
        Call fgLogOut_Pro
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub Mnu_SisContrasena_Click()
    Screen.MousePointer = 11
    vgValorAr = 0
    Frm_SisContrasena.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_SisNivel_Click()
    Screen.MousePointer = 11
    Frm_SisNivel.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_SisSucursal_Click()
    Screen.MousePointer = 11
    Frm_SisSucursal.Show
    Screen.MousePointer = 0
End Sub

Private Sub Mnu_SisUsuarios_Click()
    Screen.MousePointer = 11
    Frm_SisUsuario.Show
    Screen.MousePointer = 0
End Sub
