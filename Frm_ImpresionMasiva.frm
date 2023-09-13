VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ImpresionMasiva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Masiva de Documentos"
   ClientHeight    =   5295
   ClientLeft      =   5190
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8865
   Begin VB.CommandButton Cmd_Imprimir 
      Caption         =   "&AFP"
      Height          =   1035
      Index           =   2
      Left            =   10800
      Picture         =   "Frm_ImpresionMasiva.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Imprimir Reporte"
      Top             =   7080
      Width           =   885
   End
   Begin VB.Frame Frame5 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1185
      Left            =   180
      TabIndex        =   41
      Top             =   0
      Width           =   8565
      Begin MSComCtl2.DTPicker DTPDesde 
         Height          =   345
         Left            =   3390
         TabIndex        =   47
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   74645505
         CurrentDate     =   43573
      End
      Begin VB.CommandButton Cmd_BuscarCotizaciones 
         Caption         =   "&Buscar"
         Height          =   675
         Left            =   7770
         Picture         =   "Frm_ImpresionMasiva.frx":53E2
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Buscar Todas las Póliza"
         Top             =   450
         Width           =   720
      End
      Begin VB.OptionButton OptHoy 
         Caption         =   "Hoy"
         Height          =   255
         Left            =   180
         TabIndex        =   43
         Top             =   300
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton OptRangoFechas 
         Caption         =   "Por Rango de Fechas"
         Height          =   405
         Left            =   180
         TabIndex        =   42
         Top             =   690
         Width           =   2505
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   345
         Left            =   5100
         TabIndex        =   48
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   74645505
         CurrentDate     =   43573
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   46
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   3390
         TabIndex        =   45
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.Frame Fra_AntGral 
      Caption         =   "  Antecedentes Generales  "
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
      Height          =   3255
      Left            =   480
      TabIndex        =   7
      Top             =   5850
      Width           =   8655
      Begin VB.Label Lbl_ReajusteDescripcion 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Lbl_ReajusteValorMen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   39
         Top             =   915
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Lbl_ReajusteValor 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   38
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   29
         Left            =   240
         TabIndex        =   37
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Lbl_ReajusteTipo 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6120
         TabIndex        =   36
         Top             =   915
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Reajuste Trim."
         Height          =   255
         Index           =   28
         Left            =   6000
         TabIndex        =   35
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Lbl_NumLiquidacion 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6600
         TabIndex        =   34
         Top             =   2805
         Width           =   1575
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Boleta de Venta"
         Height          =   255
         Index           =   9
         Left            =   4680
         TabIndex        =   33
         Top             =   2805
         Width           =   1815
      End
      Begin VB.Label Lbl_FechaRec 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   2805
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Traspaso Prima"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   31
         Top             =   2805
         Width           =   1815
      End
      Begin VB.Label Lbl_Diferidos 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   30
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Lbl_Meses 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   29
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label Lbl_PensionDef 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6600
         TabIndex        =   28
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Lbl_PrimaDef 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Lbl_NumIdent 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4920
         TabIndex        =   26
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Lbl_TipoIdent 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Lbl_Modalidad 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   1755
         Width           =   3615
      End
      Begin VB.Label Lbl_TipoRenta 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   1470
         Width           =   3615
      End
      Begin VB.Label Lbl_TipoPension 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   1185
         Width           =   6135
      End
      Begin VB.Label Lbl_NomAfiliado 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   630
         Width           =   6135
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Ident."
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Meses Garant."
         Height          =   255
         Index           =   6
         Left            =   6000
         TabIndex        =   19
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Años Diferidos"
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   18
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Modalidad"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   1755
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Renta"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Pensión"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1185
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   630
         Width           =   1215
      End
      Begin VB.Line Lin_Separar 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   8160
         Y1              =   2385
         Y2              =   2385
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Prima Definitiva"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Pensión Definitiva"
         Height          =   195
         Index           =   8
         Left            =   4680
         TabIndex        =   12
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "CUSPP"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   11
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label Lbl_CUSPP 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   915
         Width           =   2535
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   9
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   8
         Top             =   2520
         Width           =   375
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   4020
      Width           =   8655
      Begin VB.CommandButton cmdBoleta 
         Caption         =   "&Boleta"
         Height          =   1035
         Left            =   3960
         Picture         =   "Frm_ImpresionMasiva.frx":54E4
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Prima"
         Height          =   1035
         Index           =   5
         Left            =   3000
         Picture         =   "Frm_ImpresionMasiva.frx":A8C6
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Imprimir Reporte"
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton cmdReporteBienvenida 
         Caption         =   "&Bienvenida"
         Height          =   1035
         Left            =   120
         Picture         =   "Frm_ImpresionMasiva.frx":FCA8
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton cmd_afp 
         Caption         =   "&AFP"
         Height          =   1035
         Left            =   2040
         Picture         =   "Frm_ImpresionMasiva.frx":1508A
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton cmdPoliza 
         Caption         =   "&Póliza"
         Height          =   1035
         Left            =   1080
         Picture         =   "Frm_ImpresionMasiva.frx":1A46C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   1035
         Left            =   4920
         Picture         =   "Frm_ImpresionMasiva.frx":1F84E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir del Formulario"
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.CommandButton Cmd_Imprimir 
      Caption         =   "&Bienvenida"
      Height          =   675
      Index           =   0
      Left            =   9750
      Picture         =   "Frm_ImpresionMasiva.frx":1F948
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Reporte"
      Top             =   6480
      Width           =   960
   End
   Begin VB.TextBox Txt_Poliza 
      Height          =   285
      Left            =   9420
      MaxLength       =   10
      TabIndex        =   2
      Top             =   6030
      Width           =   1545
   End
   Begin VB.TextBox Txt_Endoso 
      Height          =   285
      Left            =   10980
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "1"
      Top             =   6030
      Width           =   465
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_GriAseg 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1260
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4789
      _Version        =   393216
      Cols            =   5
      BackColor       =   14745599
      GridColor       =   0
      AllowBigSelection=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport Rpt_Reporte 
      Left            =   9810
      Top             =   7290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Frm_ImpresionMasiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit
Dim vlCorredor As String

Const ciMonedaPensionNew As String = 0
Const ciMonedaPrimaNew As String = 1

Const ciImprimirBienvenida As Integer = 0
Const ciImprimirPoliza As Integer = 1
Const ciImprimirAFP As Integer = 2
Const ciImprimirVariacion As Integer = 3
Const ciImprimirPrimerPago As Integer = 4
Const ciImprimirPrima As Integer = 5
Const ciImprimirFactura As Integer = 6

Dim vlafp As String
Dim vlCobertura As String
Dim vlFecNacTitular As String, vlFecNacConyuge As String
Dim vlMonedaPension As String
Dim vlTipoBoleta As String
Dim vlRepresentante As String, vlDocum As String
Dim objRep As New ClsReporte
Dim vlCodTipReajusteScomp As String 'I--- ABV 05/02/2011 ---
Dim vlTipoRenta As String

Private Sub cmd_afp_Click()

Dim vlCodPar As String
Dim vlCodDerPen As String
Dim vlArchivo As String
Dim vlNombreSucursal As String, vlNombreTipoPension As String
Dim vlMonto As Double, vlMoneda As String
Dim objRep As New ClsReporte
Dim strQuery, vlFecTras As String
Dim RS As ADODB.Recordset
Dim LNGa As Long
Dim cadenaPol As String

For i = 1 To Msf_GriAseg.Rows - 1

    npoliza = Msf_GriAseg.TextMatrix(i, 0)
    nendoso = Msf_GriAseg.TextMatrix(i, 1)
    
    Txt_Poliza.Text = npoliza
    Txt_Endoso.Text = nendoso
    
    cadenaPol = cadenaPol & ",'" & Txt_Poliza.Text & "'"
Next
cadenaPol = Mid(cadenaPol, 2, Len(cadenaPol) - 1)


    On Error GoTo Errores1
   
  
    vlCodPar = cgCodParentescoCau ' "99" 'Causante
    vlCodDerPen = "10" 'Sin Derecho a Pension
    
    Screen.MousePointer = vbHourglass

    vlArchivo = strRpt & "PD_Rpt_PolizaAFP.rpt"
    
    If Not fgExiste(vlArchivo) Then     ', vbNormal
        MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
    vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_TipoPension)
    
   'Dim objRep As New ClsReporte
            'Dim strQuery, vlFecTras As String
            'Dim RS As ADODB.Recordset
            Set RS = New ADODB.Recordset
            vgQuery = "select a.num_poliza, gls_nomben, gls_nomsegben, gls_patben, gls_matben, c.cod_scomp, d.gls_elemento as TipoAFP, e.gls_elemento as TipoPension, f.gls_elemento as TipoMoneda,"
            vgQuery = vgQuery & " dist.Gls_Comuna Gls_Direccion, tc.gls_nomcontacto, tc.gls_dircontacto, t.fec_traspaso"
            vgQuery = vgQuery & " from pd_tmae_poliza a"
            vgQuery = vgQuery & " join pd_tmae_polben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
            vgQuery = vgQuery & " join ma_tpar_monedatiporeaju c on a.cod_moneda=c.cod_moneda and a.cod_tipreajuste=c.cod_tipreajuste"
            vgQuery = vgQuery & " join ma_tpar_tabcod d on a.cod_afp=d.cod_elemento and d.cod_tabla='AF'"
            vgQuery = vgQuery & " join ma_tpar_tabcod e on a.cod_tippension=e.cod_elemento and e.cod_tabla='TP'"
            vgQuery = vgQuery & " join ma_tpar_tabcod f on a.cod_moneda=f.cod_elemento and f.cod_tabla='TM'"
            vgQuery = vgQuery & " left join ma_tpar_tabcontacto tc on a.cod_afp=tc.cod_elemento and tc.cod_tabla='AF'"
            vgQuery = vgQuery & " join MA_TPAR_COMUNA dist on tc.cod_direccion=dist.cod_direccion"
            vgQuery = vgQuery & " join pd_tmae_polprirec t on a.num_poliza=t.num_poliza"
            vgQuery = vgQuery & " where a.num_endoso=(select max(num_endoso) from pd_tmae_poliza where num_poliza=a.num_poliza)"
            vgQuery = vgQuery & " and cod_par='99'"
            vgQuery = vgQuery & " and a.num_poliza in (" & cadenaPol & ")"
            vgQuery = vgQuery & " order by 1"
            RS.Open vgQuery, vgConexionBD, adOpenForwardOnly, adLockReadOnly
        
            If RS.EOF Then
                MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
                Exit Sub
            End If
            'Dim vlNombreComuna As String, vlJefeBeneficios As String, vlDireccion As String
            'Call flObtieneDatosContacto(vlafp, vlNombreComuna, vlJefeBeneficios, vlDireccion)
        
            'Dim LNGa As Long
            LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaAFP.rpt"), ".RPT", ".TTX"), 1)
            vgPalabra = fgObtenerNombre_TextoCompuesto(Lbl_NumLiquidacion)
                
            vlNombreSucursal = "Surquillo"
                
            If objRep.CargaReporte(strRpt & "", "PD_Rpt_PolizaAFP.rpt", "Carta AFP", RS, True, _
                                    ArrFormulas("Nombre", vgNombreApoderado), _
                                    ArrFormulas("Cargo", vgCargoApoderado), _
                                    ArrFormulas("Direccion", vlDireccion), _
                                    ArrFormulas("Distrito", vlNombreComuna), _
                                    ArrFormulas("ContactoAFP", vlJefeBeneficios), _
                                    ArrFormulas("Sucursal", vlNombreSucursal), _
                                    ArrFormulas("Fecha", Lbl_FechaRec)) = False Then
                                    
                MsgBox "No se pudo abrir el reporte", vbInformation
                'Exit Sub
            End If
            Exit Sub
            
    Screen.MousePointer = 0
    Exit Sub

Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If




End Sub

Private Sub Cmd_BuscarCotizaciones_Click()
    Call flCargaCarpBenef("0000000078")
End Sub
Private Sub Cmd_Buscar_Click()
On Error GoTo Err_Buscar
    
    Txt_Poliza = UCase(Trim(Txt_Poliza))
    Txt_Poliza = Format(Txt_Poliza, "0000000000")
    If (Txt_Poliza) <> "" Then
        If flBuscaAntecedentes = True Then
            Cmd_Imprimir(ciImprimirBienvenida).SetFocus
        End If
    Else
        MsgBox "Debe ingresar el Número de la Póliza a Consultar.", vbCritical, "Falta Información"
        Txt_Poliza.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = 0

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Sub
Function flBuscaAntecedentes()
'Dim vlCodPa As String
Dim vlRegistro As ADODB.Recordset
Dim vlDif As Double
Dim vlNomSeg As String
On Error GoTo Err_buscaAnt
    
    vgSql = ""
'I--- ABV 05/02/2011 ---
'    vlCodTp = "TP"
'    vlCodTr = "TR"
'    vlCodAl = "AL"
'    vlCodPa = "99"
'F--- ABV 05/02/2011 ---
    
    flBuscaAntecedentes = False
    vgSql = ""
    vgSql = "SELECT  p.num_poliza,p.cod_tippension,p.num_idenafi,a.gls_tipoidencor,p.cod_cuspp,"
    vgSql = vgSql & "p.cod_tipren,p.num_mesdif,p.cod_modalidad,p.num_mesgar,"
    vgSql = vgSql & "p.mto_priuni,p.mto_pension,p.mto_pensiongar,"
    vgSql = vgSql & "t.gls_elemento as gls_pension,"
    vgSql = vgSql & "r.gls_elemento as gls_renta,"
    vgSql = vgSql & "m.gls_elemento as gls_modalidad,"
    vgSql = vgSql & "be.gls_nomben,be.gls_patben,be.gls_matben, p.cod_afp,"
    vgSql = vgSql & "p.cod_moneda, p.mto_valmoneda, cod_liquidacion, pr.fec_traspaso, "
    vgSql = vgSql & "p.cod_cobercon, b.gls_cobercon,p.cod_dercre, p.cod_dergra, be.fec_nacben "
    vgSql = vgSql & ",p.cod_tipoidencor,p.num_idencor " '09/10/2007
    vgSql = vgSql & ",cod_renvit "
    vgSql = vgSql & ",be.gls_nomsegben " 'MC - 24/01/2008
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & ",p.cod_tipreajuste,p.mto_valreajustetri,p.mto_valreajustemen,"
    vgSql = vgSql & "tr.gls_elemento as gls_tipreajuste "
    vgSql = vgSql & ",mtr.cod_scomp as cod_montipreaju,mtr.gls_descripcion as gls_montipreaju "
'F--- ABV 05/02/2011 ---
    vgSql = vgSql & "FROM "
    vgSql = vgSql & "pd_tmae_poliza p, pd_tmae_polprirec pr, ma_tpar_tabcod t, ma_tpar_tabcod r, "
    vgSql = vgSql & "ma_tpar_tabcod m, pd_tmae_polben be, ma_tpar_tipoiden a, ma_tpar_cobercon b "
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & ",ma_tpar_tabcod tr "
    vgSql = vgSql & ",ma_tpar_monedatiporeaju mtr "
'F--- ABV 05/02/2011 ---
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "p.num_poliza = '" & Trim(Txt_Poliza) & "' AND "
    vgSql = vgSql & "p.num_poliza = pr.num_poliza AND "
    vgSql = vgSql & "p.num_poliza = be.num_poliza AND "
    vgSql = vgSql & "p.num_endoso = be.num_endoso AND "
    vgSql = vgSql & "be.cod_par = '" & cgCodParentescoCau & "' AND "
    vgSql = vgSql & "t.cod_tabla = '" & vgCodTabla_TipPen & "' AND "
    vgSql = vgSql & "t.cod_elemento = p.cod_tippension AND "
    vgSql = vgSql & "r.cod_tabla = '" & vgCodTabla_TipRen & "' AND "
    vgSql = vgSql & "r.cod_elemento = p.cod_tipren AND "
    vgSql = vgSql & "m.cod_tabla = '" & vgCodTabla_AltPen & "' AND "
    vgSql = vgSql & "m.cod_elemento = p.cod_modalidad AND "
    vgSql = vgSql & "p.cod_tipoidenafi = a.cod_tipoiden AND "
    vgSql = vgSql & "p.cod_cobercon = b.cod_cobercon "
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & "AND p.cod_tipreajuste = tr.cod_elemento(+) AND "
    vgSql = vgSql & "tr.cod_tabla = '" & vgCodTabla_TipReajuste & "' "
    vgSql = vgSql & "AND p.cod_tipreajuste = mtr.cod_tipreajuste(+) AND "
    vgSql = vgSql & "p.cod_moneda = mtr.cod_moneda(+) and p.num_endoso='" & Trim(Txt_Endoso) & "'"  '(select max(num_endoso) from pd_tmae_poliza where num_poliza=p.num_poliza )"
'F--- ABV 05/02/2011 ---
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro.EOF) Then
        'Fra_Poliza.Enabled = False
        vlFecNacTitular = DateSerial(Mid(vlRegistro!Fec_NacBen, 1, 4), Mid(vlRegistro!Fec_NacBen, 5, 2), Mid(vlRegistro!Fec_NacBen, 7, 2))
        
        
        If vlRegistro!Cod_TipRen = "6" Then 'ESCALONADA mvg 09/11/2016
            vlTipoRenta = vlRegistro!Gls_Renta
            If Not IsNull(vlRegistro!Gls_Modalidad) Then
                vlCobertura = "CON PERIODO " & vlRegistro!Gls_Modalidad
            Else
                vlCobertura = "SIN PERIODO " & vlRegistro!Gls_Modalidad
            End If
            If vlRegistro!Cod_CoberCon <> 0 Then
                If Not IsNull(vlRegistro!GLS_COBERCON) Then
                    vlCobertura = vlCobertura & " CON " & vlRegistro!GLS_COBERCON
                End If
            End If
            If vlRegistro!Cod_DerCre = "S" Then
                vlCobertura = vlCobertura & " CON D.CRECER"
            End If
            
            If vlRegistro!Cod_DerGra = "S" Then
                vlCobertura = vlCobertura & " Y CON GRATIFICACIÓN"
            End If
        Else
            vlCobertura = vlRegistro!Gls_Renta
            If vlRegistro!Cod_Modalidad = 1 Then
                If Not IsNull(vlRegistro!Gls_Modalidad) Then
                    vlCobertura = vlCobertura & " " & vlRegistro!Gls_Modalidad
                End If
            Else
                If Not IsNull(vlRegistro!Gls_Modalidad) Then
                    vlCobertura = vlCobertura & " P. " & vlRegistro!Gls_Modalidad
                End If
            End If
            If vlRegistro!Cod_CoberCon <> 0 Then
                If Not IsNull(vlRegistro!GLS_COBERCON) Then
                    vlCobertura = vlCobertura & " CON " & vlRegistro!GLS_COBERCON
                End If
            End If
            If vlRegistro!Cod_DerCre = "S" Then
                vlCobertura = vlCobertura & " CON D.CRECER"
            End If
            
            If vlRegistro!Cod_DerGra = "S" Then
                vlCobertura = vlCobertura & " Y CON GRATIFICACIÓN"
            End If
        End If
        'I - MC 24/01/2008
''        If Not IsNull(vlRegistro!Gls_MatBen) Then
''            Lbl_NomAfiliado = Trim(vlRegistro!Gls_NomBen) + " " + Trim(vlRegistro!Gls_PatBen) + " " + Trim(vlRegistro!Gls_MatBen)
''        Else
''            Lbl_NomAfiliado = Trim(vlRegistro!Gls_NomBen) + " " + Trim(vlRegistro!Gls_PatBen) + " "
''        End If
        vlNomSeg = IIf(IsNull(vlRegistro!Gls_NomSegBen), "", Trim(vlRegistro!Gls_NomSegBen))
        Lbl_NomAfiliado = fgFormarNombreCompleto(Trim(vlRegistro!Gls_NomBen), Trim(vlNomSeg), Trim(vlRegistro!Gls_PatBen), IIf(IsNull(vlRegistro!Gls_MatBen), "", Trim(vlRegistro!Gls_MatBen)))
        'F - MC 24/01/2008
        
        vlCorredor = "SEGUROS DIRECTOS"
        
        Lbl_TipoPension = Trim(vlRegistro!Cod_TipPension) + " - " + Trim(vlRegistro!Gls_Pension)
        Lbl_TipoRenta = Trim(vlRegistro!Cod_TipRen) + " - " + Trim(vlRegistro!Gls_Renta)
        Lbl_Modalidad = Trim(vlRegistro!Cod_Modalidad) + " - " + Trim(vlRegistro!Gls_Modalidad)
        Lbl_NumIdent = Trim(vlRegistro!num_idenafi)
        Lbl_TipoIdent = Trim(vlRegistro!gls_Tipoidencor)
    
        'Lbl_NomAfiliado = Trim(vlRegistro!Gls_NomBen) + " " + Trim(vlRegistro!Gls_PatBen) + " " + Trim(vlRegistro!Gls_MatBen)
        Lbl_CUSPP = Trim(vlRegistro!Cod_Cuspp)
               
        vlDif = (vlRegistro!Num_MesDif)
        Lbl_Diferidos = ((vlDif) / 12)
        Lbl_Meses = (vlRegistro!Num_MesGar)
        
'I--- ABV 05/02/2011 ---
        Lbl_ReajusteTipo = Trim(vlRegistro!Cod_TipReajuste) + " - " + Trim(vlRegistro!Gls_TipReajuste)
        Lbl_ReajusteValor = Format(vlRegistro!Mto_ValReajusteTri, "#0.00000000")
        Lbl_ReajusteValorMen = Format(vlRegistro!Mto_ValReajusteMen, "#0.00000000")
        Lbl_ReajusteDescripcion = Trim(vlRegistro!cod_montipreaju) + " - " + Trim(vlRegistro!gls_montipreaju)
        vlCodTipReajusteScomp = IIf(IsNull(vlRegistro!cod_montipreaju), "", vlRegistro!cod_montipreaju)
'F--- ABV 05/02/2011 ---
        
        Lbl_PrimaDef = Format((vlRegistro!mto_priuni), "#,#0.00")
        Lbl_PensionDef = Format((vlRegistro!Mto_Pension), "#,#0.00")
        'vlMtoPenGar = Format((vlRegistro!Mto_PensionGar), "#,#0.00")
        
        vlMonedaPension = Trim(vlRegistro!Cod_Moneda)
        Lbl_Moneda(ciMonedaPensionNew) = vlMonedaPension
        
        Lbl_FechaRec = DateSerial(Mid(vlRegistro!fec_traspaso, 1, 4), Mid(vlRegistro!fec_traspaso, 5, 2), Mid(vlRegistro!fec_traspaso, 7, 2))
        Lbl_NumLiquidacion = Format(vlRegistro!cod_renvit, "000") & " - " & Format(vlRegistro!cod_liquidacion, "0000000")
        vlafp = vlRegistro!Cod_AFP
        Call flObtieneFecNacConyuge(Trim(Txt_Poliza), vlFecNacConyuge) 'Fecha Nacimiento Conyuge
        'Cmd_Poliza.Enabled = False
        flBuscaAntecedentes = True

    Else
        MsgBox "Nº de Póliza No Existe", vbCritical, "Verificar Información"
        Txt_Poliza = ""
        Txt_Poliza.SetFocus
    End If
    vlRegistro.Close
    
Exit Function
Err_buscaAnt:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Cmd_Imprimir_Click(Index As Integer)
Dim vlCodPar As String
Dim vlCodDerPen As String
Dim vlArchivo As String
Dim vlNombreSucursal As String, vlNombreTipoPension As String
Dim vlMonto As Double, vlMoneda As String
Dim objRep As New ClsReporte
Dim strQuery, vlFecTras As String
Dim RS As ADODB.Recordset
Dim LNGa As Long
Dim cadenaPol As String
For i = 1 To Msf_GriAseg.Rows - 1

    npoliza = Msf_GriAseg.TextMatrix(i, 0)
    nendoso = Msf_GriAseg.TextMatrix(i, 1)
    
    Txt_Poliza.Text = npoliza
    Txt_Endoso.Text = nendoso
    'MsgBox (npoliza & nendoso)
    
    cadenaPol = cadenaPol & ",'" & Txt_Poliza.Text & "'"
    'cadenaPol = Mid(cadenaPol, 1, 1)
'    If flBuscaAntecedentes = True Then
'            Cmd_Imprimir(ciImprimirBienvenida).SetFocus
'            CmdConstancia_Click
'
'    End If
    
    
Next
cadenaPol = Mid(cadenaPol, 2, Len(cadenaPol) - 1)
'cadenaPol = Mid(cadenaPol, 2, Len(cadenaPol) - 1)

    On Error GoTo Errores1
   
'    'Validar el Ingreso de la Póliza
'    If Txt_Poliza = "" Then
'        MsgBox "Debe ingresar Póliza a Consultar.", vbCritical, "Error de Datos"
'        Txt_Poliza.SetFocus
'        Exit Sub
'    End If
'
'    'Valida que exista la Póliza
'    If Trim(Lbl_TipoIdent) = "" Then
'        MsgBox "Debe Buscar Datos de la Póliza", vbCritical, "Error de Datos"
'        Cmd_Buscar.SetFocus
'        Exit Sub
'    End If
    
    vlCodPar = cgCodParentescoCau ' "99" 'Causante
    vlCodDerPen = "10" 'Sin Derecho a Pension
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case ciImprimirBienvenida
            vlArchivo = strRpt & "PD_Rpt_PolizaBien.rpt"
        Case ciImprimirPoliza
            vlArchivo = strRpt & "PD_Rpt_PolizaDef.rpt"
        Case ciImprimirAFP
            vlArchivo = strRpt & "PD_Rpt_PolizaAFP.rpt"
        Case ciImprimirVariacion
            vlArchivo = strRpt & "PD_Rpt_PolizaVar.rpt"
        Case ciImprimirPrimerPago
            vlArchivo = strRpt & "PD_Rpt_LiquidacionRV.rpt"
        Case ciImprimirPrima
            vlArchivo = strRpt & "PD_Rpt_PolizaPrima.rpt"
        Case ciImprimirFactura
            vlArchivo = strRpt & "PD_Rpt_PolizaFactura.rpt"
    End Select
    
    If Not fgExiste(vlArchivo) Then     ', vbNormal
        MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
    vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_TipoPension)
    
  
    Select Case Index
        Case ciImprimirBienvenida

        Case ciImprimirPoliza

        Case ciImprimirAFP
            
        Case ciImprimirVariacion
            
        Case ciImprimirPrimerPago
         
        Case ciImprimirPrima

            vlFecTras = "99991231"
            'vlFecTras = Mid(vlFecTras, 7, 4) & Mid(vlFecTras, 4, 2) & Mid(vlFecTras, 1, 2)
        
            Set RS = New ADODB.Recordset
            strQuery = "select a.num_poliza, a.num_endoso,a.gls_direccion, mto_priuni, gls_nacionalidad, fec_inipencia,"
            strQuery = strQuery & " b.Fec_FallBen , gls_comuna, gls_provincia, gls_region, cod_tippension, a.fec_crea, r.fec_traspaso"
            strQuery = strQuery & " from pd_tmae_poliza a"
            strQuery = strQuery & " join pd_tmae_polben b"
            strQuery = strQuery & " on a.num_poliza=b.num_poliza"
            strQuery = strQuery & " join MA_TPAR_COMUNA c"
            strQuery = strQuery & " on a.COD_DIRECCION=c.COD_DIRECCION"
            strQuery = strQuery & " join MA_TPAR_PROVINCIA d"
            strQuery = strQuery & " on c.cod_provincia=d.cod_provincia"
            strQuery = strQuery & " join MA_TPAR_REGION e"
            strQuery = strQuery & " on c.cod_region=e.cod_region"
            strQuery = strQuery & " join MA_TPAR_TABCOD f"
            strQuery = strQuery & " on a.cod_moneda=f.cod_elemento"
            strQuery = strQuery & " join PD_TMAE_POLPRIREC r"
            strQuery = strQuery & " on a.num_poliza=r.num_poliza"
            strQuery = strQuery & " Where b.Fec_FallBen Is Null and a.num_poliza in (" & cadenaPol & ") and a.num_endoso=(select max(num_endoso) from pd_tmae_poliza where num_poliza=a.num_poliza)"
            strQuery = strQuery & " group by a.num_poliza, a.num_endoso,a.gls_direccion, mto_priuni, gls_nacionalidad, fec_inipencia,"
            strQuery = strQuery & " b.Fec_FallBen , gls_comuna, gls_provincia, gls_region, cod_tippension, a.fec_crea, r.fec_traspaso"
            strQuery = strQuery & " order by 2"
            
            RS.Open strQuery, vgConexionBD, adOpenForwardOnly, adLockReadOnly
        
            If RS.EOF Then
                MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
                Exit Sub
            End If
            
            'Call pBuscaRepresentante(Lbl_TipoPension)
        
        
        
        'Dim LNGa As Long
        LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaPrima.rpt"), ".RPT", ".TTX"), 1)
        vgPalabra = fgObtenerNombre_TextoCompuesto(Lbl_NumLiquidacion)
            
        If objRep.CargaReporte(strRpt & "", "PD_Rpt_PolizaPrima.rpt", "Liquidación de Prima", RS, True, _
                                ArrFormulas("Contratante", Lbl_NomAfiliado), _
                                ArrFormulas("Asegurado", Lbl_NomAfiliado), _
                                ArrFormulas("Concatenar", vlCobertura), _
                                ArrFormulas("NombreCompania", UCase(vgNombreCompania)), _
                                ArrFormulas("NroLiquidacion", vgPalabra), _
                                ArrFormulas("Sucursal", vlNombreSucursal), _
                                ArrFormulas("NroBoleta", Lbl_NumLiquidacion), _
                                ArrFormulas("NomRepresentante", vlRepresentante), _
                                ArrFormulas("Fec_trasp", vlFecTras)) = False Then
                                
            MsgBox "No se pudo abrir el reporte", vbInformation
            'Exit Sub
        End If
        Exit Sub
        Case ciImprimirFactura
            
    End Select
    Rpt_Reporte.SelectionFormula = vgQuery
    Rpt_Reporte.Destination = crptToWindow
    Rpt_Reporte.WindowState = crptMaximized
    Rpt_Reporte.Action = 1
   
    Screen.MousePointer = 0
    Exit Sub

Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If


End Sub

Private Sub Cmd_Salir_Click()
Unload Me
End Sub


'Private Sub impMasivos(numPol As String, numEnd As String)
'
'Dim vlCodPar As String
'Dim vlCodDerPen As String
'Dim vlNombreSucursal, vlNombreTipoPension As String
'Dim RS As ADODB.Recordset
'Dim vlFecTras As String
'Dim vlTipRen As String
'Dim NomReporte As String
' 'Validar el Ingreso de la Póliza
'    If Txt_Poliza = "" Then
'        MsgBox "Debe ingresar Póliza a Consultar.", vbCritical, "Error de Datos"
'        Txt_Poliza.SetFocus
'        Exit Sub
'    End If
'
'    'Valida que exista la Póliza
'    If Trim(Lbl_TipoIdent) = "" Then
'        MsgBox "Debe Buscar Datos de la Póliza", vbCritical, "Error de Datos"
'        Cmd_Buscar.SetFocus
'        Exit Sub
'    End If
'
'    vlCodPar = cgCodParentescoCau ' "99" 'Causante
'    vlCodDerPen = "10" 'Sin Derecho a Pension
'    vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
'    vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_TipoPension)
'    Call pBuscaRepresentante(Lbl_TipoPension)
'    vlFecTras = Lbl_FechaRec.Caption
'    vlFecTras = Mid(vlFecTras, 7, 4) & Mid(vlFecTras, 4, 2) & Mid(vlFecTras, 1, 2)
'    vlTipRen = Mid(Lbl_TipoRenta.Caption, 1, 1)
'
'    If vlTipRen <> "6" Then
'        NomReporte = "PD_Rpt_PolizaDef.rpt"
'    Else
'        NomReporte = "PD_Rpt_PolizaDefEsc.rpt"
'    End If
'
'On Error GoTo mierror
'
'    Set RS = New ADODB.Recordset
'    RS.CursorLocation = adUseClient
'    RS.Open "PD_LISTA_POLIZA.LISTAR('" & Txt_Poliza.Text & "', '" & Txt_Endoso & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
'
'    Dim LNGa As Long
'    LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\" & NomReporte), ".RPT", ".TTX"), 1)
'
'    If objRep.CargaReporte(strRpt & "", NomReporte, "Póliza", RS, True, _
'                            ArrFormulas("NombreAfi", Lbl_NomAfiliado.Caption), _
'                            ArrFormulas("TipoPension", vlNombreTipoPension), _
'                            ArrFormulas("MesGar", Lbl_Meses.Caption), _
'                            ArrFormulas("NombreCompania", vgNombreCompania), _
'                            ArrFormulas("Concatenar", vlCobertura), _
'                            ArrFormulas("Sucursal", "Surquillo"), _
'                            ArrFormulas("RepresentanteNom", vlRepresentante), _
'                            ArrFormulas("RepresentanteDoc", vlDocum), _
'                            ArrFormulas("CodTipPen", Left(Trim(Lbl_TipoPension), 2)), _
'                            ArrFormulas("TipoDocTit", Trim(Lbl_TipoIdent.Caption)), _
'                            ArrFormulas("NumDocTit", Trim(Lbl_NumIdent.Caption)), _
'                            ArrFormulas("fec_trasp", vlFecTras), _
'                            ArrFormulas("TipoRenta", vlTipoRenta)) = False Then
'
'
'        MsgBox "No se pudo abrir el reporte", vbInformation
'        Exit Sub
'    End If
'
'    If vlTipRen = "6" Then
'        Set RS = New ADODB.Recordset
'        RS.CursorLocation = adUseClient
'        RS.Open "PD_LISTA_POLIZA.LISTAR('" & Txt_Poliza.Text & "', '" & Txt_Endoso & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
'
'        'Dim LNGa As Long
'        LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaConsta.rpt"), ".RPT", ".TTX"), 1)
'
'        If objRep.CargaReporte(strRpt & "", "PD_Rpt_PolizaDefEscRes.rpt", "Póliza Resumen", RS, True, _
'                                ArrFormulas("NombreAfi", Lbl_NomAfiliado.Caption)) = False Then
'
'
'            MsgBox "No se pudo abrir el reporte", vbInformation
'            Exit Sub
'        End If
'        'Exit Sub
'    End If
'
'    '''******************************************CONDTSNCIA DE POLIZA*********************************************************************
'
'    Dim strTipoPension As String
'    strTipoPension = Mid(Lbl_TipoPension.Caption, 1, 2)
'    If strTipoPension >= "08" Then
'        Exit Sub
'    End If
'
'
'
'    If CInt(Lbl_Diferidos.Caption) > 0 Then
'        NomReporte = "PD_Rpt_PolizaConstaDif.rpt"
'    Else
'        NomReporte = "PD_Rpt_PolizaConstaDif.rpt"
'    End If
'
'
'    Set RS = New ADODB.Recordset
'    RS.CursorLocation = adUseClient
'    RS.Open "PD_LISTA_POLIZA.LISTAR('" & Txt_Poliza.Text & "', '" & Txt_Endoso & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
'
'    'Dim LNGa As Long
'    LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaConsta.rpt"), ".RPT", ".TTX"), 1)
'
'    If objRep.CargaReporte(strRpt & "", NomReporte, "Póliza Constancia", RS, True, _
'                            ArrFormulas("NombreAfi", Lbl_NomAfiliado.Caption), _
'                            ArrFormulas("TipoPension", vlNombreTipoPension), _
'                            ArrFormulas("MesGar", Lbl_Meses.Caption), _
'                            ArrFormulas("NombreCompania", vgNombreCompania), _
'                            ArrFormulas("Concatenar", vlCobertura), _
'                            ArrFormulas("Sucursal", "Surquillo"), _
'                            ArrFormulas("RepresentanteNom", vlRepresentante), _
'                            ArrFormulas("RepresentanteDoc", vlDocum), _
'                            ArrFormulas("CodTipPen", Left(Trim(Lbl_TipoPension), 2)), _
'                            ArrFormulas("TipoDocTit", Trim(Lbl_TipoIdent.Caption)), _
'                            ArrFormulas("NumDocTit", Trim(Lbl_NumIdent.Caption)), _
'                            ArrFormulas("fec_trasp", vlFecTras)) = False Then
'
'
'        MsgBox "No se pudo abrir el reporte", vbInformation
'        Exit Sub
'    End If
'
'
'Exit Sub
'mierror:
'    MsgBox "No pudo cargar el reporte " & Err.Description, vbInformation
'
'
'End Sub


Private Sub cmdBoleta_Click()
Dim vlCodPar As String
Dim vlCodDerPen As String
Dim vlArchivo As String
Dim vlNombreSucursal As String, vlNombreTipoPension As String
Dim vlMonto As Double, vlMoneda As String
Dim RS As ADODB.Recordset

Dim cadenaPol As String
For i = 1 To Msf_GriAseg.Rows - 1

    npoliza = Msf_GriAseg.TextMatrix(i, 0)
    nendoso = Msf_GriAseg.TextMatrix(i, 1)
    
    Txt_Poliza.Text = npoliza
    Txt_Endoso.Text = nendoso
    
    cadenaPol = cadenaPol & "," & CStr(CInt(Txt_Poliza.Text)) & ""
Next
cadenaPol = Mid(cadenaPol, 2, Len(cadenaPol) - 1)
'cadenaPol = Mid(cadenaPol, 2, Len(cadenaPol) - 1)

    On Error GoTo mierror
   
'    'Validar el Ingreso de la Póliza
'    If Txt_Poliza = "" Then
'        MsgBox "Debe ingresar Póliza a Consultar.", vbCritical, "Error de Datos"
'        Txt_Poliza.SetFocus
'        Exit Sub
'    End If
'
'    'Valida que exista la Póliza
'    If Trim(Lbl_TipoIdent) = "" Then
'        MsgBox "Debe Buscar Datos de la Póliza", vbCritical, "Error de Datos"
'        Cmd_Buscar.SetFocus
'        Exit Sub
'    End If
    
    vlCodPar = cgCodParentescoCau ' "99" 'Causante
    vlCodDerPen = "10" 'Sin Derecho a Pension
    
    Screen.MousePointer = vbHourglass
    vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
    'vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_TipoPension)
    'vlMonto = CDbl(Lbl_PrimaDef)
    vlMoneda = cgCodTipMonedaUF

    
    

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "PD_LISTA_POLIZA_FACTURA.LISTAR_MASIVO(" & cadenaPol & ")", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaBoletaMas.rpt"), ".RPT", ".TTX"), 1)
    'Lbl_NomAfiliado
    If objRep.CargaReporte(strRpt & "", "PD_Rpt_PolizaFacturaMas.rpt", "Boleta de Venta", RS, True, _
                            ArrFormulas("IdentificacionEmpresa", vgNumIdenCompania), _
                            ArrFormulas("Corredor", vlCorredor), _
                            ArrFormulas("Asegurado", ""), _
                            ArrFormulas("Concatenar", vlCobertura), _
                            ArrFormulas("NroLiquidacion", "000"), _
                            ArrFormulas("Sucursal", vlNombreSucursal), _
                            ArrFormulas("MontoPalabras", ""), _
                            ArrFormulas("IdentificacionAfiliado", "")) = False Then
            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
Exit Sub
mierror:
    MsgBox "No pudo cargar el reporte " & Err.Description, vbInformation
End Sub

Private Sub cmdPoliza_Click()


Dim vlCodPar As String
Dim vlCodDerPen As String
Dim vlNombreSucursal As String
Dim vlNombreTipoPension As String
Dim RS As ADODB.Recordset
Dim vlFecTras As String
Dim vlTipRen As String
Dim NomReporte As String

Dim cadenaPol As String
For i = 1 To Msf_GriAseg.Rows - 1

    npoliza = Msf_GriAseg.TextMatrix(i, 0)
    nendoso = Msf_GriAseg.TextMatrix(i, 1)
    
    Txt_Poliza.Text = npoliza
    Txt_Endoso.Text = nendoso
    'MsgBox (npoliza & nendoso)
    
    cadenaPol = cadenaPol & ",''" & Txt_Poliza.Text & "''"
    'cadenaPol = Mid(cadenaPol, 1, 1)
'    If flBuscaAntecedentes = True Then
'            Cmd_Imprimir(ciImprimirBienvenida).SetFocus
'            CmdConstancia_Click
'
'    End If
    
    
Next
cadenaPol = Mid(cadenaPol, 2, Len(cadenaPol) - 2)
cadenaPol = Mid(cadenaPol, 2, Len(cadenaPol) - 1)
'cadenaPol = "'" & cadenaPol
''MsgBox (cadenaPol)
 'Validar el Ingreso de la Póliza
'    If Txt_Poliza = "" Then
'        MsgBox "Debe ingresar Póliza a Consultar.", vbCritical, "Error de Datos"
'        Txt_Poliza.SetFocus
'        Exit Sub
'    End If

'    'Valida que exista la Póliza
'    If Trim(Lbl_TipoIdent) = "" Then
'        MsgBox "Debe Buscar Datos de la Póliza", vbCritical, "Error de Datos"
'        Cmd_Buscar.SetFocus
'        Exit Sub
'    End If

    vlCodPar = cgCodParentescoCau ' "99" 'Causante
    vlCodDerPen = "10" 'Sin Derecho a Pension
    vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
    vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_TipoPension)
    Call pBuscaRepresentante(Lbl_TipoPension)
    vlFecTras = Lbl_FechaRec.Caption
    vlFecTras = Mid(vlFecTras, 7, 4) & Mid(vlFecTras, 4, 2) & Mid(vlFecTras, 1, 2)
    vlTipRen = Mid(Lbl_TipoRenta.Caption, 1, 1)

    If vlTipRen <> "6" Then
        NomReporte = "PD_Rpt_PolizaDef.rpt"
    Else
        NomReporte = "PD_Rpt_PolizaDefEsc.rpt"
    End If

On Error GoTo mierror

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "PD_LISTA_POLIZA.LISTAR_MASIVO(" & cadenaPol & ")", vgConexionBD, adOpenForwardOnly, adLockReadOnly

    Dim LNGa As Long
    LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\" & NomReporte), ".RPT", ".TTX"), 1)

    If objRep.CargaReporte(strRpt & "", NomReporte, "Póliza", RS, True, _
                            ArrFormulas("NombreAfi", Lbl_NomAfiliado.Caption), _
                            ArrFormulas("TipoPension", vlNombreTipoPension), _
                            ArrFormulas("MesGar", Lbl_Meses.Caption), _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("Concatenar", vlCobertura), _
                            ArrFormulas("Sucursal", "Surquillo"), _
                            ArrFormulas("RepresentanteNom", vlRepresentante), _
                            ArrFormulas("RepresentanteDoc", vlDocum), _
                            ArrFormulas("CodTipPen", Left(Trim(Lbl_TipoPension), 2)), _
                            ArrFormulas("TipoDocTit", Trim(Lbl_TipoIdent.Caption)), _
                            ArrFormulas("NumDocTit", Trim(Lbl_NumIdent.Caption)), _
                            ArrFormulas("fec_trasp", vlFecTras), _
                            ArrFormulas("TipoRenta", vlTipoRenta)) = False Then


        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If

'    If vlTipRen = "6" Then
'        Set RS = New ADODB.Recordset
'        RS.CursorLocation = adUseClient
'        RS.Open "PD_LISTA_POLIZA.LISTAR('" & Txt_Poliza.Text & "', '" & Txt_Endoso & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
'
'        'Dim LNGa As Long
'        LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaConsta.rpt"), ".RPT", ".TTX"), 1)
'
'        If objRep.CargaReporte(strRpt & "", "PD_Rpt_PolizaDefEscRes.rpt", "Póliza Resumen", RS, True, _
'                                ArrFormulas("NombreAfi", Lbl_NomAfiliado.Caption)) = False Then
'
'
'            MsgBox "No se pudo abrir el reporte", vbInformation
'            Exit Sub
'        End If
'        'Exit Sub
'    End If
'
    '''******************************************CONDTSNCIA DE POLIZA*********************************************************************

'    Dim strTipoPension As String
'    strTipoPension = Mid(Lbl_TipoPension.Caption, 1, 2)
'    If strTipoPension >= "08" Then
'        Exit Sub
'    End If



    'If CInt(Lbl_Diferidos.Caption) > 0 Then
        NomReporte = "PD_Rpt_PolizaConstaDif.rpt"
    'Else
    '    NomReporte = "PD_Rpt_PolizaConstaDif.rpt"
    'End If


    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "PD_LISTA_POLIZA.LISTAR_MASIVO(" & cadenaPol & ")", vgConexionBD, adOpenForwardOnly, adLockReadOnly

    'Dim LNGa As Long
    LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaConsta.rpt"), ".RPT", ".TTX"), 1)

    If objRep.CargaReporte(strRpt & "", NomReporte, "Póliza Constancia", RS, True, _
                            ArrFormulas("NombreAfi", Lbl_NomAfiliado.Caption), _
                            ArrFormulas("TipoPension", vlNombreTipoPension), _
                            ArrFormulas("MesGar", Lbl_Meses.Caption), _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("Concatenar", vlCobertura), _
                            ArrFormulas("Sucursal", "Surquillo"), _
                            ArrFormulas("RepresentanteNom", vlRepresentante), _
                            ArrFormulas("RepresentanteDoc", vlDocum), _
                            ArrFormulas("CodTipPen", Left(Trim(Lbl_TipoPension), 2)), _
                            ArrFormulas("TipoDocTit", Trim(Lbl_TipoIdent.Caption)), _
                            ArrFormulas("NumDocTit", Trim(Lbl_NumIdent.Caption)), _
                            ArrFormulas("fec_trasp", vlFecTras)) = False Then


        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If


Exit Sub
mierror:
    MsgBox "No pudo cargar el reporte " & Err.Description, vbInformation

End Sub

Private Sub cmdReporteBienvenida_Click()
Dim vlCodPar As String
Dim vlCodDerPen As String
Dim vlNombreSucursal As String
Dim vlFecTras As String
Dim RS As ADODB.Recordset
Dim cadenaPol As String

For i = 1 To Msf_GriAseg.Rows - 1

    npoliza = Msf_GriAseg.TextMatrix(i, 0)
    nendoso = Msf_GriAseg.TextMatrix(i, 1)
    
    Txt_Poliza.Text = npoliza
    Txt_Endoso.Text = nendoso
    'MsgBox (npoliza & nendoso)
    
    cadenaPol = cadenaPol & ",''" & Txt_Poliza.Text & "''"
    'cadenaPol = Mid(cadenaPol, 1, 1)
'    If flBuscaAntecedentes = True Then
'            Cmd_Imprimir(ciImprimirBienvenida).SetFocus
'            CmdConstancia_Click
'
'    End If
    
    
Next
cadenaPol = Mid(cadenaPol, 2, Len(cadenaPol) - 2)
cadenaPol = Mid(cadenaPol, 2, Len(cadenaPol) - 1)

'    'Validar el Ingreso de la Póliza
'    If Txt_Poliza = "" Then
'        MsgBox "Debe ingresar Póliza a Consultar.", vbCritical, "Error de Datos"
'        Txt_Poliza.SetFocus
'        Exit Sub
'    End If
'
'    'Valida que exista la Póliza
'    If Trim(Lbl_TipoIdent) = "" Then
'        MsgBox "Debe Buscar Datos de la Póliza", vbCritical, "Error de Datos"
'        Cmd_Buscar.SetFocus
'        Exit Sub
'    End If
    
    vlCodPar = cgCodParentescoCau ' "99" 'Causante
    vlCodDerPen = "10" 'Sin Derecho a Pension
    vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
    'vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_TipoPension)
    vlFecTras = Lbl_FechaRec.Caption
    
    vlFecTras = "99991231" 'Mid(vlFecTras, 7, 4) & Mid(vlFecTras, 4, 2) & Mid(vlFecTras, 1, 2)
    
On Error GoTo mierror

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "PP_LISTA_BIENVENIDA.LISTAR_MASIVO(" & cadenaPol & ")", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaBien.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PD_Rpt_PolizaBien.rpt", "Carta de Bienvenida", RS, True, _
                            ArrFormulas("NombreCompaniaCorto", vgNombreCortoCompania), _
                            ArrFormulas("Nombre", vgNombreApoderado), _
                            ArrFormulas("Cargo", vgCargoApoderado), _
                            ArrFormulas("Sucursal", "Surquillo"), _
                            ArrFormulas("fec_trasp", vlFecTras)) = False Then
            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
Exit Sub
mierror:
    MsgBox "No pudo cargar el reporte " & Err.Description, vbInformation
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    DTPDesde.Value = Date - 1
    DTPHasta.Value = Date
    
        
    Msf_GriAseg.Clear
    Msf_GriAseg.Cols = 3
    Msf_GriAseg.Rows = 1
    Msf_GriAseg.FixedCols = 0
    Msf_GriAseg.Row = 0
    
    Msf_GriAseg.Col = 0
    Msf_GriAseg.ColWidth(0) = 1100
    Msf_GriAseg.Text = "Nº POLIZA"
    
    Msf_GriAseg.Col = 1
    Msf_GriAseg.ColWidth(1) = 1100
    Msf_GriAseg.Text = "Nº ENDOSO"
    
    Msf_GriAseg.Col = 2
    Msf_GriAseg.ColWidth(2) = 1500
    Msf_GriAseg.Text = "FECHA EMISION"
    
End Sub

Private Sub OptHoy_Click()
    DTPDesde.Enabled = False
    DTPHasta.Enabled = False
End Sub

Private Sub OptRangoFechas_Click()
    DTPDesde.Enabled = True
    DTPHasta.Enabled = True
End Sub
'-------------------------------------------
'CARGA INFORMACION EN LA GRILLA
'-------------------------------------------
Function flCargaCarpBenef(iNumPol As String)
Dim vlFechaEmis As String

On Error GoTo Err_CargaBen
    Dim vlRut As String
    If OptHoy.Value = True Then
      vlSql = ""
      vlSql = "SELECT MAx(A.NUM_ENDOSO) NUM_ENDOSO,A.NUM_POLIZA, B.FEC_TRASPASO FEC_EMISION "
      vlSql = vlSql & "FROM PD_TMAE_POLIZA A JOIN PD_TMAE_POLPRIREC B ON A.NUM_POLIZA=B.NUM_POLIZA  "
      vlSql = vlSql & "WHERE B.FEC_TRASPASO = TO_CHAR(sysdate,'YYYYMMDD') group by  A.NUM_POLIZA,B.FEC_TRASPASO order by FEC_EMISION,NUM_POLIZA asc"
    End If
    
    If OptRangoFechas.Value = True Then
      vlSql = ""
      vlSql = "SELECT MAx(A.NUM_ENDOSO) NUM_ENDOSO,A.NUM_POLIZA, B.FEC_TRASPASO FEC_EMISION "
      vlSql = vlSql & "FROM PD_TMAE_POLIZA A JOIN PD_TMAE_POLPRIREC B ON A.NUM_POLIZA=B.NUM_POLIZA  "
      vlSql = vlSql & "WHERE B.FEC_TRASPASO >= TO_CHAR(TO_DATE('" & DTPDesde & "','dd/mm/yyyy'), 'YYYYMMDD') AND B.FEC_TRASPASO <= TO_CHAR(TO_DATE('" & DTPHasta & "','dd/mm/yyyy'), 'YYYYMMDD') group by  A.NUM_POLIZA,B.FEC_TRASPASO order by FEC_EMISION,NUM_POLIZA asc"
    End If
        
'    vlSql = ""
'    vlSql = "SELECT MAx(A.NUM_ENDOSO) NUM_ENDOSO,A.NUM_POLIZA, B.FEC_TRASPASO FEC_EMISION "
'    vlSql = vlSql & "FROM PD_TMAE_POLIZA A JOIN PD_TMAE_POLPRIREC B ON A.NUM_POLIZA=B.NUM_POLIZA  "
'    vlSql = vlSql & "WHERE B.FEC_TRASPASO >= TO_CHAR(TO_DATE('" & DTPDesde & "','dd/mm/yyyy'), 'YYYYMMDD') AND B.FEC_TRASPASO <= TO_CHAR(TO_DATE('" & DTPHasta & "','dd/mm/yyyy'), 'YYYYMMDD') group by  A.NUM_POLIZA,B.FEC_TRASPASO order by FEC_EMISION,NUM_POLIZA asc"
'
'    vlSql = ""
'    vlSql = "SELECT MAx(A.NUM_ENDOSO) NUM_ENDOSO,A.NUM_POLIZA, B.FEC_TRASPASO FEC_EMISION "
'    vlSql = vlSql & "FROM PD_TMAE_POLIZA A JOIN PD_TMAE_POLPRIREC B ON A.NUM_POLIZA=B.NUM_POLIZA  "
'    vlSql = vlSql & "WHERE B.FEC_TRASPASO = TO_CHAR(sysdate,'YYYYMMDD') group by  A.NUM_POLIZA,B.FEC_TRASPASO order by FEC_EMISION,NUM_POLIZA asc"
'
    
    Set vgRs = vgConexionBD.Execute(vlSql)
    Msf_GriAseg.Rows = 1
    
    While Not vgRs.EOF
        
        If Not IsNull(vgRs!Fec_Emision) Then
            vlFechaEmis = DateSerial(Mid(vgRs!Fec_Emision, 1, 4), Mid(vgRs!Fec_Emision, 5, 2), Mid(vgRs!Fec_Emision, 7, 2))
        Else
            vlFechaEmis = ""
        End If
                
        Msf_GriAseg.AddItem Trim(vgRs!Num_Poliza) & vbTab _
                            & vgRs!Num_Endoso & vbTab _
                            & vlFechaEmis & vbTab
        vgRs.MoveNext
    Wend
    vgRs.Close
 
Exit Function
Err_CargaBen:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flObtieneFecNacConyuge(iPoliza As String, iFecNacConyuge As String)
    
    Dim vlCodPar As String
    Dim vlRegistro As ADODB.Recordset
    vlCodPar = "'10', '11'" 'Parentesco Conyuge
    
    vgSql = "SELECT a.fec_nacben"
    vgSql = vgSql & " FROM pd_tmae_polben a"
    vgSql = vgSql & " WHERE a.num_poliza = '" & iPoliza & "'"
    vgSql = vgSql & " AND a.cod_par IN (" & vlCodPar & ")"
    vgSql = vgSql & " AND (a.cod_sexo = 'F' OR a.cod_sitinv <> 'N')" 'Conyuge Mujer o Esposo Inválido
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro.EOF) Then
        iFecNacConyuge = DateSerial(Mid(vlRegistro!Fec_NacBen, 1, 4), Mid(vlRegistro!Fec_NacBen, 5, 2), Mid(vlRegistro!Fec_NacBen, 7, 2))
    End If
    
End Function

Private Sub CmdConstancia_Click()
Dim vlCodPar As String
Dim vlCodDerPen As String
Dim vlNombreSucursal, vlNombreTipoPension As String
Dim RS As ADODB.Recordset
Dim vlFecTras As String
 'Validar el Ingreso de la Póliza
    If Txt_Poliza = "" Then
        MsgBox "Debe ingresar Póliza a Consultar.", vbCritical, "Error de Datos"
        Txt_Poliza.SetFocus
        Exit Sub
    End If
    
    'Valida que exista la Póliza
    If Trim(Lbl_TipoIdent) = "" Then
        MsgBox "Debe Buscar Datos de la Póliza", vbCritical, "Error de Datos"
        Cmd_Buscar.SetFocus
        Exit Sub
    End If
    
    vlCodPar = cgCodParentescoCau ' "99" 'Causante
    vlCodDerPen = "10" 'Sin Derecho a Pension
    vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
    vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_TipoPension)
    Call pBuscaRepresentante(Lbl_TipoPension)
    vlFecTras = Lbl_FechaRec.Caption
    vlFecTras = Mid(vlFecTras, 7, 4) & Mid(vlFecTras, 4, 2) & Mid(vlFecTras, 1, 2)
    
    
On Error GoTo mierror


    '''******************************************CONDTSNCIA DE POLIZA*********************************************************************
    
    If Mid(Lbl_TipoPension.Caption, 1, 2) = "08" Then
        Exit Sub
    End If
    
    Dim NomReporte As String
    
    If CInt(Lbl_Diferidos.Caption) > 0 Then
        NomReporte = "PD_Rpt_PolizaConstaDif.rpt"
    Else
        NomReporte = "PD_Rpt_PolizaConstaDif.rpt"
    End If
    
    
      Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "PD_LISTA_POLIZA.LISTAR('" & Txt_Poliza.Text & "', '" & Txt_Endoso & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaConsta.rpt"), ".RPT", ".TTX"), 1)

    If objRep.CargaReporte(strRpt & "", NomReporte, "Póliza Constancia", RS, True, _
                            ArrFormulas("NombreAfi", Lbl_NomAfiliado.Caption), _
                            ArrFormulas("TipoPension", vlNombreTipoPension), _
                            ArrFormulas("MesGar", Lbl_Meses.Caption), _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("Concatenar", vlCobertura), _
                            ArrFormulas("Sucursal", "Surquillo"), _
                            ArrFormulas("RepresentanteNom", vlRepresentante), _
                            ArrFormulas("RepresentanteDoc", vlDocum), _
                            ArrFormulas("CodTipPen", Left(Trim(Lbl_TipoPension), 2)), _
                            ArrFormulas("TipoDocTit", Trim(Lbl_TipoIdent.Caption)), _
                            ArrFormulas("NumDocTit", Trim(Lbl_NumIdent.Caption)), _
                            ArrFormulas("fec_trasp", vlFecTras)) = False Then
            
            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
    
Exit Sub
mierror:
    MsgBox "No pudo cargar el reporte " & Err.Description, vbInformation
    
End Sub
Private Sub pBuscaRepresentante(TP As String)
On Error GoTo Err_Cargarep
Dim vlSql As String


TP = Mid(TP, 1, 2)

If TP = "08" Then

    vlSql = ""
    vlSql = "SELECT * FROM pd_tmae_polrep a, ma_tpar_tipoiden b WHERE "
    vlSql = vlSql & "num_poliza = '" & Txt_Poliza & "' and a.cod_tipoidenrep = b.cod_tipoiden"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlRepresentante = IIf(Not IsNull(vgRs!Gls_NombresRep), vgRs!Gls_NombresRep, "")
        vlRepresentante = vlRepresentante & " " & IIf(Not IsNull(vgRs!Gls_ApepatRep), vgRs!Gls_ApepatRep, "")
        vlRepresentante = vlRepresentante & " " & IIf(Not IsNull(vgRs!Gls_ApematRep), vgRs!Gls_ApematRep, "")
        vlDocum = IIf(Not IsNull(vgRs!gls_Tipoidencor), vgRs!gls_Tipoidencor, "") & " " & IIf(Not IsNull(vgRs!NUM_IDENREP), vgRs!NUM_IDENREP, "")
    End If
    vgRs.Close
Else
    vlSql = ""
    vlSql = "SELECT * FROM pd_tmae_polben WHERE "
    vlSql = vlSql & "num_poliza = '" & Txt_Poliza & "' and cod_par='99'"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlRepresentante = IIf(Not IsNull(vgRs!Gls_NomBen), vgRs!Gls_NomBen, "")
        vlRepresentante = vlRepresentante & " " & IIf(Not IsNull(vgRs!Gls_NomSegBen), vgRs!Gls_NomSegBen, "")
        vlRepresentante = vlRepresentante & " " & IIf(Not IsNull(vgRs!Gls_PatBen), vgRs!Gls_PatBen, "")
        vlRepresentante = vlRepresentante & " " & IIf(Not IsNull(vgRs!Gls_MatBen), vgRs!Gls_MatBen, "")
        'vlDocum = IIf(Not IsNull(vgRs!gls_Tipoidencor), vgRs!gls_Tipoidencor, "") & " " & IIf(Not IsNull(vgRs!Num_Idenrep), vgRs!Num_Idenrep, "")
    End If
    vgRs.Close
End If
 
Exit Sub
Err_Cargarep:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub
Function flObtieneDatosContacto(iAfp As String, iNombreComuna As String, iApoderado As String, iDireccion As String)
    Dim vlRegistro As ADODB.Recordset
    vgSql = "SELECT a.cod_direccion, a.gls_nomcontacto, a.gls_dircontacto"
    vgSql = vgSql & " FROM ma_tpar_tabcontacto a"
    vgSql = vgSql & " WHERE a.cod_tabla = '" & vgCodTabla_AFP & "'"
    vgSql = vgSql & " AND a.cod_elemento = '" & iAfp & "'"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro.EOF) Then
        Call fgBuscarNombreComunaProvinciaRegion(vlRegistro!Cod_Direccion)
        iNombreComuna = vgNombreComuna
        iApoderado = vlRegistro!gls_nomcontacto
        iDireccion = vlRegistro!gls_dircontacto
    Else
        MsgBox "No se encontraron Datos de Contacto AFP", vbCritical, "Error de Datos"
    End If
    
End Function
Function flVerificaPrimerPago(iPoliza As String, oMontoPago As Double, oMoneda As String) As Boolean
    Dim vlRegistro As ADODB.Recordset
    flVerificaPrimerPago = False
    vgSql = "SELECT SUM(a.mto_liqpagar) AS monto, b.gls_elemento"
    vgSql = vgSql & " FROM pd_tmae_liqpagopen a, ma_tpar_tabcod b"
    vgSql = vgSql & " WHERE a.num_poliza = '" & iPoliza & "'"
    vgSql = vgSql & " AND b.cod_tabla = 'TM'"
    vgSql = vgSql & " AND b.cod_elemento = a.cod_moneda"
    vgSql = vgSql & " GROUP BY b.gls_elemento"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If (vlRegistro.EOF) Then
        MsgBox "No se encontraron Primeros Pagos para la Póliza: '" & iPoliza & "'", vbCritical, "Inexistencia de Datos"
        Exit Function
    Else
        oMontoPago = vlRegistro!monto
        oMoneda = vlRegistro!gls_elemento
    End If
    flVerificaPrimerPago = True
End Function



