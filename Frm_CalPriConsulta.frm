VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_CalPriConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Primas Recepcionadas."
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9255
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   9015
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   675
         Left            =   3600
         Picture         =   "Frm_CalPriConsulta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4680
         Picture         =   "Frm_CalPriConsulta.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "Imprimir"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_CalPriConsulta.frx":0AFC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5880
         Picture         =   "Frm_CalPriConsulta.frx":11B6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   720
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      BackColor       =   14745599
      AllowUserResizing=   1
   End
   Begin VB.Frame Fra_RangoFec 
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.OptionButton opt_todasRan 
         Caption         =   "Todas las Polizas del Rango"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   1900
         Width           =   4695
      End
      Begin VB.OptionButton Opt_Rezagadas 
         Caption         =   "Primas de Pre-Pólizas"
         Height          =   255
         Left            =   3360
         TabIndex        =   16
         Top             =   1400
         Width           =   2655
      End
      Begin VB.OptionButton Opt_TodasInf 
         Caption         =   "Todas las informadas."
         Height          =   195
         Left            =   3360
         TabIndex        =   15
         Top             =   1680
         Width           =   3735
      End
      Begin VB.OptionButton Opt_Traspasada 
         Caption         =   "Primas de Pólizas Traspasadas a Pago de Pensiones."
         Height          =   195
         Left            =   3360
         TabIndex        =   14
         Top             =   1040
         Width           =   4455
      End
      Begin VB.OptionButton Opt_NoTraspasada 
         Caption         =   "Primas de Pólizas No Traspasadas a Pago de Pensiones."
         Height          =   195
         Left            =   3360
         TabIndex        =   13
         Top             =   820
         Value           =   -1  'True
         Width           =   4935
      End
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   6960
         Picture         =   "Frm_CalPriConsulta.frx":12B0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Efectuar Consulta"
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   3360
         X2              =   7920
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   17
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Desea Consultar por :"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Recepción de Primas"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   11
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rango de Fechas   :"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4560
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "Frm_CalPriConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlRegistro     As ADODB.Recordset
Dim vlFechaDesde   As String
Dim vlFechaHasta   As String
Dim vlFecTraspaso  As String
Dim vlAnno         As String
Dim vlMes          As String
Dim vlDia          As String
Dim vlOpcion       As String
Dim vlCodMoneda    As String
Dim objRep As New ClsReporte
Private Sub Cmd_Buscar_Click()
On Error GoTo Err_Buscar

    If (Txt_Desde) <> "" Then
        If flValidaFecha(Txt_Desde) = True Then
           If (Txt_Hasta) <> "" Then
               If flValidaFecha(Txt_Hasta) = True Then
                  Call flConsultaRecepcion
               End If
           Else
               MsgBox "Debe Ingresar Fecha Hasta", vbCritical, "Error de Datos"
               Txt_Hasta.SetFocus
           End If
        End If
    Else
       MsgBox "Debe Ingresar Fecha Desde", vbCritical, "Error de Datos"
       Txt_Desde.SetFocus
    End If

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Imprimir_Click()
Dim vlArchivo As String
Dim rs As ADODB.Recordset
Dim LNGa As Long
Err.Clear
On Error GoTo Errores1
   
    'Validar el Ingreso del Rango de Fechas
    If Txt_Desde = "" Then
        MsgBox "Debe ingresar el Rango de Inicio de la Consulta de Primas Recaudadas.", vbCritical, "Error de Datos"
        Txt_Desde.SetFocus
        Exit Sub
    Else
        vlFechaIni = Trim(Txt_Desde)
        If (flValidaFecha(vlFechaIni) = False) Then
            Txt_Desde = ""
            Txt_Desde.SetFocus
            Exit Sub
        End If
    End If
    
    If Txt_Hasta = "" Then
       MsgBox "Debe ingresar el Rango de Inicio de la Consulta de Primas Recaudadas.", vbCritical, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    Else
        vlFechaTer = Trim(Txt_Hasta)
        If (flValidaFecha(vlFechaTer) = False) Then
            Txt_Hasta = ""
            Txt_Hasta.SetFocus
            Exit Sub
        End If
    End If
   
    Screen.MousePointer = 11
   
    If Opt_Rezagadas.Value = True Then
          
        vgSql = ""
        vgSql = "SELECT p.num_poliza,p.cod_tippension,p.cod_afp,p.mto_priuni,"
        vgSql = vgSql & " p.num_idencor,i.gls_tipoidencor,to_date(p.fec_solicitud, 'YYYYMMDD')"
        vgSql = vgSql & " as fec_solicitud,to_date(p.fec_vigencia, 'YYYYMMDD') as fec_vigencia,"
        vgSql = vgSql & " i.gls_tipoidencor as gls_tipidenben, x.num_idenben, x.gls_nomben || ' ' || x.gls_nomsegben || ' ' || x.gls_patben || ' ' || x.gls_matben as nombres,"
        vgSql = vgSql & " t.gls_elemento as gls_pension,a.gls_elemento as gls_afp,"
        vgSql = vgSql & " to_date(q.fec_acepta, 'YYYYMMDD') as fec_incorp_poliza, p.cod_cuspp, mto_pricia"
        vgSql = vgSql & " FROM pd_tmae_oripoliza p,ma_tpar_tabcod t,ma_tpar_tabcod a, pd_tmae_polprirecaux r,"
        vgSql = vgSql & " ma_tpar_tipoiden i, pt_tmae_detcotizacion q, pd_tmae_oripolben x, ma_tpar_tipoiden y"
        vgSql = vgSql & " WHERE"
        vgSql = vgSql & " p.cod_tippension = t.cod_elemento AND x.num_poliza = p.num_poliza AND x.num_orden = 1 AND"
        vgSql = vgSql & " t.cod_tabla = '" & vgCodTabla_TipPen & "' AND"
        vgSql = vgSql & " p.cod_afp = a.cod_elemento AND r.num_poliza(+) = p.num_poliza AND"
        vgSql = vgSql & " a.cod_tabla = '" & vgCodTabla_AFP & "' AND"
        vgSql = vgSql & " q.num_cot = p.num_cot AND"
        vgSql = vgSql & " p.fec_vigencia >= '" & vlFechaDesde & "' AND"
        vgSql = vgSql & " p.fec_vigencia <= '" & vlFechaHasta & "' AND"
        vgSql = vgSql & " p.cod_tipoidencor = i.cod_tipoiden AND"
        vgSql = vgSql & " x.cod_tipoidenben = y.cod_tipoiden"
        vgSql = vgSql & " order by 1"
        
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open vgSql, vgConexionBD, adOpenForwardOnly, adLockReadOnly
        
        
        LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\pd_rpt_CalPriRezagados.rpt"), ".RPT", ".TTX"), 1)
        
        If objRep.CargaReporte(strRpt & "", "pd_rpt_CalPriRezagados.rpt", "Carta de Bienvenida", rs, True, _
                                ArrFormulas("NombreCompaniaCorto", vgNombreCortoCompania), _
                                ArrFormulas("Nombre", vgNombreApoderado), _
                                ArrFormulas("Cargo", vgCargoApoderado), _
                                ArrFormulas("Sucursal", "Surquillo"), _
                                ArrFormulas("fec_ini", Txt_Desde.Text), _
                                ArrFormulas("fec_fin", Txt_Hasta.Text), _
                                ArrFormulas("Titulo", "Listado de Todas las Polizas Informadas por Rango de Fechas")) = False Then
                
            MsgBox "No se pudo abrir el reporte", vbInformation
            Exit Sub
        End If
        Exit Sub
        
    ElseIf opt_todasRan = True Then
    
        vgSql = ""
        vgSql = "select a.num_poliza, b.gls_nomben || ' ' || b.gls_nomsegben || ' ' || b.gls_patben || ' ' || b.gls_matben as nombres,"
        vgSql = vgSql & " e.gls_elemento as TipoPen,"
        vgSql = vgSql & " c.mto_pRItotal, mto_prirec as primaTranf,"
        vgSql = vgSql & " a.prc_corcomreal as Comision, d.gls_nomcor || ' ' || d.gls_patcor || ' ' || d.gls_matcor as Asesor, to_date(fec_traspaso, 'YYYYMMDD') as FEC_TRAS,"
        vgSql = vgSql & " substr(fec_traspaso,1,6) as MesAno, cod_cuspp"
        vgSql = vgSql & " from pd_tmae_poliza a"
        vgSql = vgSql & " join pd_tmae_polben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
        vgSql = vgSql & " join pd_tmae_polprirec c on c.num_poliza=a.num_poliza"
        vgSql = vgSql & " join pt_tmae_corredor d on a.num_idencor=d.num_idencor"
        vgSql = vgSql & " join ma_tpar_tabcod e on a.cod_tippension=e.cod_elemento and e.cod_tabla='TP'"
        vgSql = vgSql & " where a.num_endoso=(select max(num_endoso) from pd_tmae_poliza where num_poliza=a.num_poliza)"
        vgSql = vgSql & " and cod_par=99 and fec_traspaso between '" & vlFechaDesde & "' and '" & vlFechaHasta & "'"
        vgSql = vgSql & " Union All"
        vgSql = vgSql & " select a.num_poliza, b.gls_nomben || ' ' || b.gls_nomsegben || ' ' || b.gls_patben || ' ' || b.gls_matben as nombres,"
        vgSql = vgSql & " e.gls_elemento as TipoPen,"
        vgSql = vgSql & " c.mto_prirec, mto_pricia as primaTranf,"
        vgSql = vgSql & " a.prc_corcomreal as Comision, d.gls_nomcor || ' ' || d.gls_patcor || ' ' || d.gls_matcor as Asesor, to_date(fec_traspaso, 'YYYYMMDD') as FEC_TRAS,"
        vgSql = vgSql & " substr(fec_traspaso,1,6) as MesAno, cod_cuspp"
        vgSql = vgSql & " from pd_tmae_oripoliza a"
        vgSql = vgSql & " join pd_tmae_oripolben b on a.num_poliza=b.num_poliza"
        vgSql = vgSql & " join pd_tmae_polprirecaux c on c.num_poliza=a.num_poliza"
        vgSql = vgSql & " join pt_tmae_corredor d on a.num_idencor=d.num_idencor"
        vgSql = vgSql & " join ma_tpar_tabcod e on a.cod_tippension=e.cod_elemento and e.cod_tabla='TP'"
        vgSql = vgSql & " where cod_par=99 and fec_traspaso between  '" & vlFechaDesde & "' and '" & vlFechaHasta & "'"
        vgSql = vgSql & " order by 1"
        
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open vgSql, vgConexionBD, adOpenForwardOnly, adLockReadOnly
        
        'Dim LNGa As Long
        LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\pd_rpt_CalPriRangos.rpt"), ".RPT", ".TTX"), 1)
        
        If objRep.CargaReporte(strRpt & "", "PD_Rpt_CalPriRangos.rpt", "Carta de Bienvenida", rs, True, _
                                ArrFormulas("NombreCompaniaCorto", vgNombreCortoCompania), _
                                ArrFormulas("Nombre", vgNombreApoderado), _
                                ArrFormulas("Cargo", vgCargoApoderado), _
                                ArrFormulas("Sucursal", "Surquillo"), _
                                ArrFormulas("fec_ini", Txt_Desde.Text), _
                                ArrFormulas("fec_fin", Txt_Hasta.Text), _
                                ArrFormulas("Titulo", "Listado de Todas las Polizas por Rango de Fechas")) = False Then
                
            MsgBox "No se pudo abrir el reporte", vbInformation
            Exit Sub
        End If
        Exit Sub
    Else
     
        vgSql = ""
        vgSql = "select a.num_poliza, b.gls_nomben || ' ' || b.gls_nomsegben || ' ' || b.gls_patben || ' ' || b.gls_matben as nombres,"
        vgSql = vgSql & " e.gls_elemento as TipoPen,"
        vgSql = vgSql & " c.mto_pRItotal, mto_prirec as primaTranf,"
        vgSql = vgSql & " a.prc_corcomreal as Comision, d.gls_nomcor || ' ' || d.gls_patcor || ' ' || d.gls_matcor as Asesor, to_date(fec_traspaso, 'YYYYMMDD') as FEC_TRAS,"
        vgSql = vgSql & " substr(fec_traspaso,1,6) as MesAno, cod_cuspp,"
        vgSql = vgSql & " f.gls_elemento as afp, to_date(g.fec_acepta, 'YYYYMMDD') as fec_incorp_poliza"
        vgSql = vgSql & " from pd_tmae_poliza a"
        vgSql = vgSql & " join pd_tmae_polben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
        vgSql = vgSql & " join pd_tmae_polprirec c on c.num_poliza=a.num_poliza"
        vgSql = vgSql & " join pt_tmae_corredor d on a.num_idencor=d.num_idencor"
        vgSql = vgSql & " join ma_tpar_tabcod e on a.cod_tippension=e.cod_elemento and e.cod_tabla='TP'"
        vgSql = vgSql & " join ma_tpar_tabcod f on a.cod_afp=f.cod_elemento and f.cod_tabla='AF'"
        vgSql = vgSql & " join pt_tmae_detcotizacion g on a.num_cot=g.num_cot"
        vgSql = vgSql & " where a.num_endoso=(select max(num_endoso) from pd_tmae_poliza where num_poliza=a.num_poliza)"
        vgSql = vgSql & " and cod_par=99 and fec_traspaso between '" & vlFechaDesde & "' and '" & vlFechaHasta & "'"
        vgSql = vgSql & " order by 1"
        
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open vgSql, vgConexionBD, adOpenForwardOnly, adLockReadOnly
        
        
        LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\pd_rpt_CalPriRangos.rpt"), ".RPT", ".TTX"), 1)
        
        If objRep.CargaReporte(strRpt & "", "PD_Rpt_CalPriRangos.rpt", "Carta de Bienvenida", rs, True, _
                                ArrFormulas("NombreCompaniaCorto", vgNombreCortoCompania), _
                                ArrFormulas("Nombre", vgNombreApoderado), _
                                ArrFormulas("Cargo", vgCargoApoderado), _
                                ArrFormulas("Sucursal", "Surquillo"), _
                                ArrFormulas("fec_ini", Txt_Desde.Text), _
                                ArrFormulas("fec_fin", Txt_Hasta.Text), _
                                ArrFormulas("Titulo", "Listado de Todas las Polizas Informadas por Rango de Fechas")) = False Then
                
            MsgBox "No se pudo abrir el reporte", vbInformation
            Exit Sub
        End If
        Exit Sub
    End If
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

    If Opt_Rezagadas = True Then
        flLmpGrRez
    Else
        flLmpGrilla
    End If
    Txt_Desde = ""
    Txt_Hasta = ""
    Txt_Desde.SetFocus

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
'*****INI GCP 27/01/2022 Exportacion de Primas ********************
Private Sub cmdExportar_Click()

On Error GoTo Err_Cargar

    If Me.Opt_Traspasada.Value = True Then
          'Validar el Ingreso del Rango de Fechas
    If Txt_Desde = "" Then
        MsgBox "Debe ingresar el Rango de Inicio de la Consulta de Primas Recaudadas.", vbCritical, "Error de Datos"
        Txt_Desde.SetFocus
        Exit Sub
    Else
        vlFechaIni = Trim(Txt_Desde)
        If (flValidaFecha(vlFechaIni) = False) Then
            Txt_Desde = ""
            Txt_Desde.SetFocus
            Exit Sub
        End If
    End If
    
    If Txt_Hasta = "" Then
       MsgBox "Debe ingresar el Rango de Inicio de la Consulta de Primas Recaudadas.", vbCritical, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    Else
        vlFechaTer = Trim(Txt_Hasta)
        If (flValidaFecha(vlFechaTer) = False) Then
            Txt_Hasta = ""
            Txt_Hasta.SetFocus
            Exit Sub
        End If
    End If
   
   Screen.MousePointer = 11
   Exportar_polizasTraspasadas vlFechaDesde, vlFechaHasta
    
        
    Else
        MsgBox "No implementado para la opción seleccionada", vbInformation, "Consulta de Primas Recepcionadas"
        
    End If
    
     
    Screen.MousePointer = 0
    
    Exit Sub
    
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
    
    
End Sub
Private Sub Exportar_polizasTraspasadas(ByVal pfechaIni As String, ByVal pfechaFin As String)
    
   
        Dim objCmd As ADODB.Command
        Dim rs As ADODB.Recordset
        Dim conn As ADODB.Connection
            
        Set rs = New ADODB.Recordset
        Set conn = New ADODB.Connection
        
        Dim Texto As String
        
        
        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Set objCmd = New ADODB.Command
        
        conn.Provider = "OraOLEDB.Oracle"
        conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
        conn.CursorLocation = adUseClient
        conn.Open
  
                       
        Set objCmd = CreateObject("ADODB.Command")
        Set objCmd.ActiveConnection = conn
        
        objCmd.CommandText = "SP_getPrimasRecepcionadas"
        objCmd.CommandType = adCmdStoredProc
        
         
        
        Set param1 = objCmd.CreateParameter("pfecha_ini", adVarChar, adParamInput, 8, pfechaIni)
        objCmd.Parameters.Append param1
                
        Set param2 = objCmd.CreateParameter("pfecha_fin", adVarChar, adParamInput, 8, pfechaFin)
        objCmd.Parameters.Append param2
       
     
        Set rs = objCmd.Execute
        
        Exportar rs, Me.Txt_Desde, Me.Txt_Hasta
                       
                   
        conn.Close
        Set rs = Nothing
        Set conn = Nothing

End Sub

    Private Sub Exportar(ByRef rs As ADODB.Recordset, ByVal FechaIni As String, ByVal FechaFin As String)

    Dim Obj_Excel As Object
    Dim Obj_Libro As Object
    Dim Obj_Hoja As Object
    Dim vFila As Integer
    Dim vColumna As Integer
    
    Set Obj_Excel = CreateObject("Excel.Application")
    
    Set Obj_Libro = Obj_Excel.Workbooks.Add
    Set Obj_Hoja = Obj_Libro.Worksheets.Add
    
    vFila = 0
    vColumna = 0
    
    Dim Titulos(19) As String
 
    Titulos(0) = "Poliza"
    Titulos(1) = "Tipo Documento"
    Titulos(2) = "Número Documento"
    Titulos(3) = "Nombre Titular"
    Titulos(4) = "AFP"
    Titulos(5) = "Tipo Prestación"
    Titulos(6) = "Fec. Incor."
    Titulos(7) = "Prima Total"
    Titulos(8) = "Prima Transferida"
    Titulos(9) = "Comisión"
    Titulos(10) = "NroDoc Asesor"
    Titulos(11) = "AsesorJV"
    Titulos(12) = "NroDoc Supervisor"
    Titulos(13) = "Supervisor"
    Titulos(14) = "NroDoc Jefe"
    Titulos(15) = "Jefe"
    Titulos(16) = "Traspaso"
    Titulos(17) = "CUSPP"
    Titulos(18) = "Mes Dif."
    Titulos(19) = "Fec. 1er Pago"
    
    vTotaCampos = 19
      
    vFila = 2
    Obj_Hoja.Cells(vFila, 1) = "PROTECTA"
    Obj_Hoja.Cells(vFila, 19) = "FECHA"
    Obj_Hoja.Cells(vFila, 20).NumberFormat = "@"
    Obj_Hoja.Cells(vFila, 20) = Format(Now, "dd/MM/yyyy")
    Obj_Hoja.rows(vFila).Font.Bold = True
    vFila = 5

    Obj_Hoja.Cells(vFila, 1) = "Listado de Todas las Polizas Informadas por Rango de Fechas"
    Obj_Hoja.Range("A5:T5").MergeCells = True
    Obj_Hoja.Cells(vFila, 1).HorizontalAlignment = xlCenter
    Obj_Hoja.rows(vFila).Font.Bold = True
    
     vFila = 6
     Obj_Hoja.Cells(vFila, 1) = "(Desde " & FechaIni & " hasta " & FechaFin & ")"
     Obj_Hoja.rows(vFila).Font.Bold = True
     Obj_Hoja.Range("A6:T6").MergeCells = True
     Obj_Hoja.Cells(vFila, 1).HorizontalAlignment = xlCenter
     
    
     Obj_Hoja.Columns("A:G").NumberFormat = "@"
     Obj_Hoja.Columns("H:I").NumberFormat = "######,##0.00"
     Obj_Hoja.Columns("J:S").NumberFormat = "@"
     Obj_Hoja.Columns("T").NumberFormat = "@"
     
   
    vFila = 12
    
    For i = 0 To UBound(Titulos)
    
          Obj_Hoja.Cells(vFila, i + 1) = Titulos(i)
    
    Next
  
     Obj_Hoja.rows(3).Font.Bold = True
     Obj_Hoja.Range("A12:T12").Borders.LineStyle = xlContinuous
     Obj_Hoja.Range("A12:T12").VerticalAlignment = xlVAlignCenter
     Obj_Hoja.Range("A12:T12").Font.Bold = True

          
      Do While Not rs.EOF
               vFila = vFila + 1
              For vColumna = 0 To vTotaCampos
                Obj_Hoja.Cells(vFila, vColumna + 1) = rs.Fields(vColumna).Value
                
              Next
            
            Me.Refresh
            
        rs.MoveNext
        
      Loop
      

    Obj_Hoja.Columns("A:T").AutoFit
      
    Obj_Excel.Visible = True
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing
    
   
End Sub


'*****FIN GCP 27/01/2022 Exportacion de Primas ********************

Private Sub Form_Load()
On Error GoTo Err_Cargar

   Frm_CalPriConsulta.Top = 0
   Frm_CalPriConsulta.Left = 0
   flLmpGrilla

    Call fgCargarTablaMoneda(vgCodTabla_TipMon, egTablaMoneda(), vgNumeroTotalTablasMoneda)

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Opt_NoTraspasada_Click()
    flLmpGrilla
End Sub

Private Sub Opt_Recepcionadas_Click()
    flLmpGrilla
End Sub

Private Sub Opt_Rezagadas_Click()
    flLmpGrRez
End Sub

Private Sub Opt_TodasInf_Click()
    flLmpGrilla
End Sub

Private Sub Opt_Traspasada_Click()
    flLmpGrilla
End Sub

Private Sub Txt_Desde_GotFocus()
    Txt_Desde.SelStart = 0
    Txt_Desde.SelLength = Len(Txt_Desde)
End Sub

Private Sub Txt_Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And (Trim(Txt_Desde <> "")) Then
       If flValidaFecha(Txt_Desde) = True Then
          Txt_Hasta.SetFocus
       End If
    End If
End Sub

Private Sub Txt_Desde_LostFocus()
    If Txt_Desde = "" Then
       Exit Sub
    End If
    If Not IsDate(Txt_Desde) Then
       Txt_Desde = ""
       Exit Sub
    End If
    If Txt_Desde <> "" Then
       vlFechaDesde = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
       Txt_Desde = DateSerial(Mid((vlFechaDesde), 1, 4), Mid((vlFechaDesde), 5, 2), Mid((vlFechaDesde), 7, 2))
    End If
End Sub

Private Sub Txt_Hasta_GotFocus()
    Txt_Hasta.SelStart = 0
    Txt_Hasta.SelLength = Len(Txt_Hasta)
End Sub

Private Sub Txt_Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(Txt_Hasta <> "") Then
       If flValidaFecha(Txt_Hasta) = True Then
          Cmd_Buscar.SetFocus
       End If
    End If
End Sub

Private Sub Txt_Hasta_LostFocus()
    If Txt_Hasta = "" Then
       Exit Sub
    End If
    If Not IsDate(Txt_Hasta) Then
       Txt_Hasta = ""
       Exit Sub
    End If
    If Txt_Hasta <> "" Then
       vlFechaHasta = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
       Txt_Hasta = DateSerial(Mid((vlFechaHasta), 1, 4), Mid((vlFechaHasta), 5, 2), Mid((vlFechaHasta), 7, 2))
    End If
End Sub

Function flValidaFecha(iFecha)
On Error GoTo Err_valfecha

      flValidaFecha = False
     
     'valida que la fecha este correcta
      If Trim(iFecha <> "") Then
         If Not IsDate(iFecha) Then
                MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Dato Incorrecto"
                Exit Function
         End If
         
         If Txt_Hasta <> "" Then
            If (CDate(Format(Txt_Hasta, "dd/mm/yyyy")) < Format(Txt_Desde, "dd/mm/yyyy")) Then
                 MsgBox "Rango Fecha Hasta es menor al Rango de Fecha Desde", vbCritical, "Dato Incorrecto"
                 Exit Function
            End If
         End If
     
         If (Year(iFecha) < 1900) Then
             MsgBox "La Fecha ingresada es menor a la mínima que se puede ingresar (1900).", vbCritical, "Dato Incorrecto"
             Exit Function
         End If
     
        flValidaFecha = True
     
     End If

Exit Function
Err_valfecha:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Function

Function flConsultaRecepcion()
On Error GoTo Err_Consulta
Dim Mensaje As String
    
    If Opt_Rezagadas = True Then
       Call flRezagadas
    ElseIf opt_todasRan = True Then
       Call flTodosRango
    Else
        If Opt_TodasInf = True Then
            vlOpcion = ""
            Call flTodasInformadas
        Else
            If Opt_NoTraspasada = True Then
                vlOpcion = "N"
            Else
                If Opt_Traspasada = True Then
                    vlOpcion = "S"
                    Call ActualizaAsesores(vlFechaDesde, vlFechaHasta)
                    
                End If
            End If
            Call flTraspasadaSINO
        End If
    End If
                       
Exit Function
Err_Consulta:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Function
Function flLmpGrRango()

    Msf_Grilla.Clear
    Msf_Grilla.rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0
    
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Nº de Póliza"
    Msf_Grilla.ColWidth(0) = 1100
    Msf_Grilla.ColAlignment(0) = 1
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Tip.Pensión"
    Msf_Grilla.ColWidth(1) = 1700
    Msf_Grilla.ColAlignment(1) = 1
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Nombres"
    Msf_Grilla.ColWidth(2) = 2000
    Msf_Grilla.ColAlignment(2) = 1
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Prima Unica"
    Msf_Grilla.ColWidth(3) = 1200
    Msf_Grilla.ColAlignment(3) = 3
    
    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Prima Trans."
    Msf_Grilla.ColWidth(4) = 1200
    Msf_Grilla.ColAlignment(4) = 1

    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "CUSPP"
    Msf_Grilla.ColWidth(5) = 1400
    Msf_Grilla.ColAlignment(5) = 1
    
    Msf_Grilla.Col = 6
    Msf_Grilla.Text = "Comisiòn %"
    Msf_Grilla.ColWidth(6) = 1000
    Msf_Grilla.ColAlignment(6) = 3

    Msf_Grilla.Col = 7
    Msf_Grilla.Text = "Traspaso"
    Msf_Grilla.ColWidth(7) = 1000
    Msf_Grilla.ColAlignment(7) = 3
    
    Msf_Grilla.Col = 8
    Msf_Grilla.Text = "Nombre Asesor"
    Msf_Grilla.ColWidth(8) = 2000
    Msf_Grilla.ColAlignment(8) = 3
    
    Msf_Grilla.Col = 9
    Msf_Grilla.ColWidth(9) = 0
    
    Msf_Grilla.Col = 10
    Msf_Grilla.ColWidth(10) = 0

End Function
Function flTodosRango()
On Error GoTo Err_TraRan
     
    vgSql = ""
    vgSql = "select a.num_poliza, b.gls_nomben || ' ' || b.gls_nomsegben || ' ' || b.gls_patben || ' ' || b.gls_matben as nombres,"
    vgSql = vgSql & " e.gls_elemento as TipoPen,"
    vgSql = vgSql & " c.mto_pRItotal, mto_prirec as primaTranf,"
    vgSql = vgSql & " a.prc_corcomreal as Comision, d.gls_nomcor || ' ' || d.gls_patcor || ' ' || d.gls_matcor as Asesor, to_date(fec_traspaso, 'YYYYMMDD') as FEC_TRAS,"
    vgSql = vgSql & " substr(fec_traspaso,1,6) as MesAno, cod_cuspp"
    vgSql = vgSql & " from pd_tmae_poliza a"
    vgSql = vgSql & " join pd_tmae_polben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    vgSql = vgSql & " join pd_tmae_polprirec c on c.num_poliza=a.num_poliza"
    vgSql = vgSql & " join pt_tmae_corredor d on a.num_idencor=d.num_idencor"
    vgSql = vgSql & " join ma_tpar_tabcod e on a.cod_tippension=e.cod_elemento and e.cod_tabla='TP'"
    vgSql = vgSql & " where a.num_endoso=(select max(num_endoso) from pd_tmae_poliza where num_poliza=a.num_poliza)"
    vgSql = vgSql & " and cod_par=99 and fec_traspaso between '" & vlFechaDesde & "' and '" & vlFechaHasta & "'"
    vgSql = vgSql & " Union All"
    vgSql = vgSql & " select a.num_poliza, b.gls_nomben || ' ' || b.gls_nomsegben || ' ' || b.gls_patben || ' ' || b.gls_matben as nombres,"
    vgSql = vgSql & " e.gls_elemento as TipoPen,"
    vgSql = vgSql & " c.mto_prirec, mto_pricia as primaTranf,"
    vgSql = vgSql & " a.prc_corcomreal as Comision, d.gls_nomcor || ' ' || d.gls_patcor || ' ' || d.gls_matcor as Asesor, to_date(fec_traspaso, 'YYYYMMDD') as FEC_TRAS,"
    vgSql = vgSql & " substr(fec_traspaso,1,6) as MesAno, cod_cuspp"
    vgSql = vgSql & " from pd_tmae_oripoliza a"
    vgSql = vgSql & " join pd_tmae_oripolben b on a.num_poliza=b.num_poliza"
    vgSql = vgSql & " join pd_tmae_polprirecaux c on c.num_poliza=a.num_poliza"
    vgSql = vgSql & " join pt_tmae_corredor d on a.num_idencor=d.num_idencor"
    vgSql = vgSql & " join ma_tpar_tabcod e on a.cod_tippension=e.cod_elemento and e.cod_tabla='TP'"
    vgSql = vgSql & " where cod_par=99 and fec_traspaso between  '" & vlFechaDesde & "' and '" & vlFechaHasta & "'"
    vgSql = vgSql & " order by 1"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
        flLmpGrRango
        While Not vlRegistro.EOF
            'If (vlRegistro!fec_traspaso) >= (vlFechaDesde) And _
            '   (vlRegistro!fec_traspaso) <= (vlFechaHasta) And _
            '   (vlRegistro!cod_trapagopen) = vlOpcion Then
                                  
                vlFecTraspaso = (vlRegistro!FEC_TRAS)
                'vlAnno = Mid(vlFecTraspaso, 1, 4)
                'vlMes = Mid(vlFecTraspaso, 5, 2)
                'vlDia = Mid(vlFecTraspaso, 7, 2)
                'vlFecTraspaso = DateSerial((vlAnno), (vlMes), (vlDia))
'I--- ABV 05/02/2011 ---
'                vlCodMoneda = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vlRegistro!Cod_Moneda)
                'If IsNull(vlRegistro!Cod_Moneda) Then
                '    vlCodMoneda = vlRegistro!Cod_Monedapol
                'Else
                    'vlCodMoneda = vlRegistro!Cod_Moneda
                'End If
'F--- ABV 05/02/2011 ---
            
                Msf_Grilla.AddItem (vlRegistro!Num_Poliza) & vbTab & (vlRegistro!TipoPen) & vbTab & (vlRegistro!nombres) & vbTab & _
                                   (Format(vlRegistro!mto_pRItotal, "#,#0.00")) & vbTab & _
                                   (Format(vlRegistro!primaTranf, "#,#0.00")) & vbTab & (vlRegistro!Cod_Cuspp) & vbTab & _
                                   (Format(vlRegistro!Comision, "#,#0.00")) & vbTab & _
                                   (vlFecTraspaso) & vbTab & (vlRegistro!Asesor)
              'Else
              '    MsgBox "No Existe Información Para el Rango de Fecha a Consultar", vbInformation, "Información"
              '    Txt_Desde.SetFocus
            'End If
            vlRegistro.MoveNext
        Wend
    Else
        MsgBox "No Existe Información Para el Rango de Fecha a Consultar", vbInformation, "Información"
        Txt_Desde.SetFocus
    End If
    Screen.MousePointer = 0
    vlRegistro.Close

Exit Function
Err_TraRan:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


Function flTraspasadaSINO()
On Error GoTo Err_TraSINO
     
    vgSql = ""
    vgSql = "SELECT r.num_poliza,P.COD_CUSPP,r.fec_traspaso,r.fec_vigencia,r.mto_priinf,"
    vgSql = vgSql & " r.mto_pensioninf,r.mto_prirecpesos,r.mto_prirec,"
    vgSql = vgSql & " r.prc_facvar,r.mto_pension,p.num_poliza," ' P.COD_MONEDA," 'I--- ABV 05/02/2011 ---
    vgSql = vgSql & " p.cod_trapagopen,r.mto_pensiongarinf,r.mto_pensiongar"
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & ",p.cod_moneda as cod_monedapol "
    vgSql = vgSql & ",mtr.cod_scomp as cod_moneda "
'F--- ABV 05/02/2011 ---
    vgSql = vgSql & " FROM pd_tmae_polprirec r, pd_tmae_poliza p "
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & ",ma_tpar_monedatiporeaju mtr "
'F--- ABV 05/02/2011 ---
    vgSql = vgSql & " WHERE "
    vgSql = vgSql & " r.num_poliza = p.num_poliza AND "
    vgSql = vgSql & " p.num_endoso = 1 "
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & "AND p.cod_tipreajuste = mtr.cod_tipreajuste(+) AND "
    vgSql = vgSql & "p.cod_moneda = mtr.cod_moneda(+) "
'F--- ABV 05/02/2011 ---
    vgSql = vgSql & " ORDER BY r.num_poliza"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
        flLmpGrilla
        While Not vlRegistro.EOF
            If (vlRegistro!fec_traspaso) >= (vlFechaDesde) And _
               (vlRegistro!fec_traspaso) <= (vlFechaHasta) And _
               (vlRegistro!cod_trapagopen) = vlOpcion Then
                                  
                vlFecTraspaso = (vlRegistro!fec_traspaso)
                vlAnno = Mid(vlFecTraspaso, 1, 4)
                vlMes = Mid(vlFecTraspaso, 5, 2)
                vlDia = Mid(vlFecTraspaso, 7, 2)
                vlFecTraspaso = DateSerial((vlAnno), (vlMes), (vlDia))
'I--- ABV 05/02/2011 ---
'                vlCodMoneda = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vlRegistro!Cod_Moneda)
                If IsNull(vlRegistro!Cod_Moneda) Then
                    vlCodMoneda = vlRegistro!Cod_Monedapol
                Else
                    vlCodMoneda = vlRegistro!Cod_Moneda
                End If
'F--- ABV 05/02/2011 ---
            
                Msf_Grilla.AddItem (vlRegistro!Num_Poliza) & vbTab & (vlRegistro!Cod_Cuspp) & vbTab & (vlFecTraspaso) & vbTab & _
                                   (Format(vlRegistro!MTO_PRIINF, "#,#0.00")) & vbTab & _
                                   (Format(vlRegistro!MTO_PRIREC, "#,#0.00")) & vbTab & _
                                   (Format(vlRegistro!PRC_FACVAR, "#0.00")) & vbTab & _
                                   (vlCodMoneda) & vbTab & (Format(vlRegistro!MTO_PENSIONINF, "#,#0.00")) & vbTab & _
                                   (Format(vlRegistro!Mto_Pension, "#,#0.00")) & vbTab & _
                                   (Format(vlRegistro!mto_pensiongarinf, "#,#0.00")) & vbTab & _
                                   (Format(vlRegistro!Mto_PensionGar, "#,#0.00"))
              'Else
              '    MsgBox "No Existe Información Para el Rango de Fecha a Consultar", vbInformation, "Información"
              '    Txt_Desde.SetFocus
            End If
            vlRegistro.MoveNext
        Wend
    Else
        MsgBox "No Existe Información Para el Rango de Fecha a Consultar", vbInformation, "Información"
        Txt_Desde.SetFocus
    End If
    Screen.MousePointer = 0
    vlRegistro.Close

Exit Function
Err_TraSINO:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flTodasInformadas()
On Error GoTo Err_Inf
     
    vgSql = ""
    vgSql = "select a.num_poliza, b.gls_nomben || ' ' || b.gls_nomsegben || ' ' || b.gls_patben || ' ' || b.gls_matben as nombres,"
    vgSql = vgSql & " e.gls_elemento as TipoPen,"
    vgSql = vgSql & " c.mto_pRItotal, mto_prirec as primaTranf,"
    vgSql = vgSql & " a.prc_corcomreal as Comision, d.gls_nomcor || ' ' || d.gls_patcor || ' ' || d.gls_matcor as Asesor, to_date(fec_traspaso, 'YYYYMMDD') as FEC_TRAS,"
    vgSql = vgSql & " substr(fec_traspaso,1,6) as MesAno, cod_cuspp"
    vgSql = vgSql & " from pd_tmae_poliza a"
    vgSql = vgSql & " join pd_tmae_polben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    vgSql = vgSql & " join pd_tmae_polprirec c on c.num_poliza=a.num_poliza"
    vgSql = vgSql & " join pt_tmae_corredor d on a.num_idencor=d.num_idencor"
    vgSql = vgSql & " join ma_tpar_tabcod e on a.cod_tippension=e.cod_elemento and e.cod_tabla='TP'"
    vgSql = vgSql & " where a.num_endoso=(select max(num_endoso) from pd_tmae_poliza where num_poliza=a.num_poliza)"
    vgSql = vgSql & " and cod_par=99 and fec_traspaso between '" & vlFechaDesde & "' and '" & vlFechaHasta & "'"
    vgSql = vgSql & " order by 1"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
        flLmpGrRango
        While Not vlRegistro.EOF
            vlFecTraspaso = (vlRegistro!FEC_TRAS)

            Msf_Grilla.AddItem (vlRegistro!Num_Poliza) & vbTab & (vlRegistro!TipoPen) & vbTab & (vlRegistro!nombres) & vbTab & _
                                   (Format(vlRegistro!mto_pRItotal, "#,#0.00")) & vbTab & _
                                   (Format(vlRegistro!primaTranf, "#,#0.00")) & vbTab & (vlRegistro!Cod_Cuspp) & vbTab & _
                                   (Format(vlRegistro!Comision, "#,#0.00")) & vbTab & _
                                   (vlFecTraspaso) & vbTab & (vlRegistro!Asesor)
              'Else
              '    MsgBox "No Existe Información Para el Rango de Fecha a Consultar", vbInformation, "Información"
              '    Txt_Desde.SetFocus
            'End If
            vlRegistro.MoveNext
        Wend
    Else
        MsgBox "No Existe Información Para el Rango de Fecha a Consultar", vbInformation, "Información"
        Txt_Desde.SetFocus
    End If
    Screen.MousePointer = 0
    vlRegistro.Close

Exit Function
Err_Inf:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLmpGrilla()

    Msf_Grilla.Clear
    Msf_Grilla.rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0
    
    Msf_Grilla.Col = 0
    Msf_Grilla.ColWidth(0) = 1200
    Msf_Grilla.ColAlignment(0) = 1
    Msf_Grilla.Text = "Nº de Póliza"
    
    Msf_Grilla.Col = 1
    Msf_Grilla.ColWidth(1) = 1600
    Msf_Grilla.ColAlignment(1) = 3
    Msf_Grilla.Text = "CUSPP"
    
    Msf_Grilla.Col = 2
    Msf_Grilla.ColWidth(2) = 1300
    Msf_Grilla.ColAlignment(2) = 3
    Msf_Grilla.Text = "Fecha Traspaso"
    
    Msf_Grilla.Col = 3
    Msf_Grilla.ColWidth(3) = 1200
    Msf_Grilla.ColAlignment(3) = 3
    Msf_Grilla.Text = "Prima Inf."
    
    Msf_Grilla.Col = 4
    Msf_Grilla.ColWidth(4) = 1200
    Msf_Grilla.ColAlignment(4) = 3
    Msf_Grilla.Text = "Prima Rec."
    
    Msf_Grilla.Col = 5
    Msf_Grilla.ColWidth(5) = 1000
    Msf_Grilla.ColAlignment(5) = 3
    Msf_Grilla.Text = "Factor Var."

    Msf_Grilla.Col = 6
    Msf_Grilla.ColWidth(6) = 1000
    Msf_Grilla.ColAlignment(6) = 3
    Msf_Grilla.Text = "Moneda"
    
    Msf_Grilla.Col = 7
    Msf_Grilla.ColWidth(7) = 1000
    Msf_Grilla.ColAlignment(7) = 3
    Msf_Grilla.Text = "Pensión Inf."

    Msf_Grilla.Col = 8
    Msf_Grilla.ColWidth(8) = 1100
    Msf_Grilla.ColAlignment(8) = 3
    Msf_Grilla.Text = "Pensión Def."
    
    Msf_Grilla.Col = 9
    Msf_Grilla.ColWidth(9) = 0 '1200
    'Msf_Grilla.ColAlignment(9) = 3
    Msf_Grilla.Text = "Pens.Gar.Inf."
    
    Msf_Grilla.Col = 10
    Msf_Grilla.ColWidth(10) = 0 '1200
    'Msf_Grilla.ColAlignment(10) = 3
    Msf_Grilla.Text = "Pens.Gar.Def."

End Function

Function flRezagadas()
On Error GoTo Err_Rez
     
     vgSql = ""
     vgSql = "SELECT p.num_poliza,p.cod_tippension,p.cod_afp,p.mto_priuni,"
     vgSql = vgSql & "p.num_idencor,i.gls_tipoidencor,p.fec_solicitud,p.fec_vigencia,"
     vgSql = vgSql & "t.gls_elemento as gls_pension,a.gls_elemento as gls_afp "
     vgSql = vgSql & " FROM pd_tmae_oripoliza p,ma_tpar_tabcod t,ma_tpar_tabcod a, ma_tpar_tipoiden i "
     vgSql = vgSql & " WHERE "
     vgSql = vgSql & " p.cod_tippension = t.cod_elemento AND "
     vgSql = vgSql & " t.cod_tabla = '" & vgCodTabla_TipPen & "' AND "
     vgSql = vgSql & " p.cod_afp = a.cod_elemento AND "
     vgSql = vgSql & " a.cod_tabla = '" & vgCodTabla_AFP & "' AND "
     vgSql = vgSql & " p.fec_vigencia >= '" & vlFechaDesde & "' AND "
     vgSql = vgSql & " p.fec_vigencia <= '" & vlFechaHasta & "' AND "
     vgSql = vgSql & " p.cod_tipoidencor = i.cod_tipoiden"
     Set vlRegistro = vgConexionBD.Execute(vgSql)
     If Not vlRegistro.EOF Then
        flLmpGrRez
        While Not vlRegistro.EOF
            If (vlRegistro!Fec_Vigencia) >= (vlFechaDesde) And _
               (vlRegistro!Fec_Vigencia) <= (vlFechaHasta) Then

                vlFecCot = (vlRegistro!Fec_Solicitud)
                vlAnno = Mid(vlFecCot, 1, 4)
                vlMes = Mid(vlFecCot, 5, 2)
                vlDia = Mid(vlFecCot, 7, 2)
                vlFecCot = DateSerial((vlAnno), (vlMes), (vlDia))
              
                vlVigencia = (vlRegistro!Fec_Vigencia)
                vlAnno = Mid(vlVigencia, 1, 4)
                vlMes = Mid(vlVigencia, 5, 2)
                vlDia = Mid(vlVigencia, 7, 2)
                vlVigencia = DateSerial((vlAnno), (vlMes), (vlDia))
                      
                vlDiasRetraso = DateDiff("d", CDate(vlFecCot), CDate(vlVigencia))
                      
                Msf_Grilla.AddItem (vlRegistro!Num_Poliza) & vbTab & (vlRegistro!Gls_Pension) & vbTab & _
                                   (vlRegistro!GLS_AFP) & vbTab & _
                                   (Format(vlRegistro!MTO_PRIUNI, "#,#0.00")) & vbTab & _
                                   ((vlRegistro!GLS_TIPOIDENCOR) & vbTab & (vlRegistro!Num_IdenCor)) & vbTab & _
                                   (vlFecCot) & vbTab & (vlVigencia) & vbTab & (vlDiasRetraso)

            End If
            vlRegistro.MoveNext
        Wend
    Else
        MsgBox "No Existe Información Para el Rango de Fecha a Consultar", vbInformation, "Información"
        Txt_Desde.SetFocus
    End If
    Screen.MousePointer = 0
    vlRegistro.Close

Exit Function
Err_Rez:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLmpGrRez()

    Msf_Grilla.Clear
    Msf_Grilla.rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0
    
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Nº de Póliza"
    Msf_Grilla.ColWidth(0) = 1100
    Msf_Grilla.ColAlignment(0) = 1
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Tip.Pensión"
    Msf_Grilla.ColWidth(1) = 1700
    Msf_Grilla.ColAlignment(1) = 1
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "AFP"
    Msf_Grilla.ColWidth(2) = 1700
    Msf_Grilla.ColAlignment(2) = 1
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Prima Unica"
    Msf_Grilla.ColWidth(3) = 1200
    Msf_Grilla.ColAlignment(3) = 3
    
    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Tipo Ident. Inter."
    Msf_Grilla.ColWidth(4) = 1400
    Msf_Grilla.ColAlignment(4) = 1

    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "Nº Ident. Inter."
    Msf_Grilla.ColWidth(5) = 1400
    Msf_Grilla.ColAlignment(5) = 1
    
    Msf_Grilla.Col = 6
    Msf_Grilla.Text = "Fec.Cotiz."
    Msf_Grilla.ColWidth(6) = 1000
    Msf_Grilla.ColAlignment(6) = 3

    Msf_Grilla.Col = 7
    Msf_Grilla.Text = "Fec.Cierre"
    Msf_Grilla.ColWidth(7) = 1000
    Msf_Grilla.ColAlignment(7) = 3
    
    Msf_Grilla.Col = 8
    Msf_Grilla.Text = "Dias Retraso"
    Msf_Grilla.ColWidth(8) = 1200
    Msf_Grilla.ColAlignment(8) = 3
    
    Msf_Grilla.Col = 9
    Msf_Grilla.ColWidth(9) = 0
    
    Msf_Grilla.Col = 10
    Msf_Grilla.ColWidth(10) = 0

End Function
Function ActualizaAsesores(ByVal pfecha_ini As String, ByVal pfecha_fin As String) As String
          
    
                Dim objCmd As ADODB.Command
                Dim rs As ADODB.Recordset
                Dim conn As ADODB.Connection
                    
                Set rs = New ADODB.Recordset
                Set conn = New ADODB.Connection
                
                Dim Texto As String
                       
                Set conn = New ADODB.Connection
                Set rs = New ADODB.Recordset
                Set objCmd = New ADODB.Command
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
          
                               
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "SP_CorreccionAsesores"
                objCmd.CommandType = adCmdStoredProc
                
'                pFecha_ini varchar2,
'                pFecha_Fin varchar2,
'                p_outNumError out number,
'                p_outMsgError out VARCHAR2
                          
                Set param1 = objCmd.CreateParameter("pFecha_ini", adVarChar, adParamInput, 8, pfecha_ini)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("pFecha_Fin", adInteger, adParamInput, 8, pfecha_fin)
                objCmd.Parameters.Append param2
                
                    
                Set param3 = objCmd.CreateParameter("p_outNumError", adDouble, adParamOutput)
                objCmd.Parameters.Append param3
                
                Set param4 = objCmd.CreateParameter("p_outMsgError", adVarChar, adParamOutput, 200)
                objCmd.Parameters.Append param4
                
                                       
                Set rs = objCmd.Execute
                
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    Mensaje = objCmd.Parameters.Item("p_outMsgError").Value
                Else
                    Mensaje = ""
                End If
                
                ActualizaAsesores = p_outMsgError
                   
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        
End Function
