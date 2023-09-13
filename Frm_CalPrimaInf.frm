VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_CalPrimaInf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes Definitivos de Pólizas."
   ClientHeight    =   5760
   ClientLeft      =   675
   ClientTop       =   2535
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8895
   Begin VB.CommandButton Cmd_Imprimir 
      Caption         =   "&Póliza"
      Height          =   990
      Index           =   1
      Left            =   2880
      Picture         =   "Frm_CalPrimaInf.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Imprimir Reporte"
      Top             =   6420
      Width           =   825
   End
   Begin VB.CommandButton Cmd_Imprimir 
      Caption         =   "&Bienvenida"
      Height          =   675
      Index           =   0
      Left            =   4005
      Picture         =   "Frm_CalPrimaInf.frx":53E2
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Imprimir Reporte"
      Top             =   6345
      Width           =   960
   End
   Begin VB.CommandButton Cmd_Imprimir 
      Caption         =   "&Primer Pago"
      Height          =   675
      Index           =   4
      Left            =   5280
      Picture         =   "Frm_CalPrimaInf.frx":5A9C
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Imprimir Reporte"
      Top             =   6315
      Width           =   1065
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   8655
      Begin VB.CommandButton CmdConstancia 
         Caption         =   "&Constancia"
         Height          =   1035
         Left            =   6150
         Picture         =   "Frm_CalPrimaInf.frx":6156
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Imprimir Constancia"
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdReportPrimerPago 
         Caption         =   "Pri. Pago"
         Height          =   1035
         Left            =   3600
         Picture         =   "Frm_CalPrimaInf.frx":B538
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdBoleta 
         Caption         =   "&Boleta"
         Height          =   1035
         Left            =   4440
         Picture         =   "Frm_CalPrimaInf.frx":1091A
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdPoliza 
         Caption         =   "&Póliza"
         Height          =   1035
         Left            =   970
         Picture         =   "Frm_CalPrimaInf.frx":15CFC
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdReporteBienvenida 
         Caption         =   "&Bienvenida"
         Height          =   1035
         Left            =   0
         Picture         =   "Frm_CalPrimaInf.frx":1B0DE
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Prima"
         Height          =   1035
         Index           =   5
         Left            =   2720
         Picture         =   "Frm_CalPrimaInf.frx":204C0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir Reporte"
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Variación"
         Height          =   1035
         Index           =   3
         Left            =   5280
         Picture         =   "Frm_CalPrimaInf.frx":258A2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir Reporte"
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&AFP"
         Height          =   1035
         Index           =   2
         Left            =   1830
         Picture         =   "Frm_CalPrimaInf.frx":2AC84
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprimir Reporte"
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   1035
         Left            =   6960
         Picture         =   "Frm_CalPrimaInf.frx":30066
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpiar Formulario"
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   1035
         Left            =   7800
         Picture         =   "Frm_CalPrimaInf.frx":304A8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir del Formulario"
         Top             =   180
         Width           =   735
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
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   8655
      Begin VB.Label Lbl_ReajusteDescripcion 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   52
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Lbl_ReajusteValorMen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   51
         Top             =   915
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Lbl_ReajusteValor 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   47
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   29
         Left            =   240
         TabIndex        =   50
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Lbl_ReajusteTipo 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6120
         TabIndex        =   49
         Top             =   915
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Reajuste Trim."
         Height          =   255
         Index           =   28
         Left            =   6000
         TabIndex        =   48
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Lbl_NumLiquidacion 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6600
         TabIndex        =   39
         Top             =   2805
         Width           =   1575
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Boleta de Venta"
         Height          =   255
         Index           =   9
         Left            =   4680
         TabIndex        =   38
         Top             =   2805
         Width           =   1815
      End
      Begin VB.Label Lbl_FechaRec 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   2805
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Traspaso Prima"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   36
         Top             =   2805
         Width           =   1815
      End
      Begin VB.Label Lbl_Diferidos 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   20
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Lbl_Meses 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   19
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label Lbl_PensionDef 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6600
         TabIndex        =   15
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Lbl_PrimaDef 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Lbl_NumIdent 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4920
         TabIndex        =   35
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Lbl_TipoIdent 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Lbl_Modalidad 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   1755
         Width           =   3615
      End
      Begin VB.Label Lbl_TipoRenta 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   1470
         Width           =   3615
      End
      Begin VB.Label Lbl_TipoPension 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   31
         Top             =   1185
         Width           =   6135
      End
      Begin VB.Label Lbl_NomAfiliado 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   630
         Width           =   6135
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Ident."
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Meses Garant."
         Height          =   255
         Index           =   6
         Left            =   6000
         TabIndex        =   28
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Años Diferidos"
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   27
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Modalidad"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Top             =   1755
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Renta"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Pensión"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   1185
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Pensión Definitiva"
         Height          =   195
         Index           =   8
         Left            =   4680
         TabIndex        =   21
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "CUSPP"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   18
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label Lbl_CUSPP 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   915
         Width           =   2535
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   16
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   14
         Top             =   2520
         Width           =   375
      End
   End
   Begin VB.CommandButton Cmd_Poliza 
      Caption         =   "&Poliza"
      Height          =   675
      Left            =   7680
      Picture         =   "Frm_CalPrimaInf.frx":305A2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Buscar Datos de la Póliza"
      Top             =   240
      Width           =   840
   End
   Begin VB.Frame Fra_Poliza 
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7365
      Begin VB.TextBox Txt_Endoso 
         Height          =   285
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   53
         Text            =   "1"
         Top             =   360
         Width           =   585
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   6840
         Top             =   255
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   5400
         Picture         =   "Frm_CalPrimaInf.frx":30C5C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Buscar Póliza"
         Top             =   310
         Width           =   615
      End
      Begin VB.TextBox Txt_Poliza 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Num. Endoso"
         Height          =   195
         Index           =   11
         Left            =   3720
         TabIndex        =   54
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "  Póliza  "
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
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Frm_CalPrimaInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

Function flRecibe(vlNumPoliza As String, vlNumEndoso As Integer)
    If vlNumPoliza <> "" Then
        Txt_Poliza = vlNumPoliza
        Cmd_Buscar_Click
        Cmd_Imprimir(ciImprimirBienvenida).SetFocus
        Exit Function
    End If
End Function

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
        Lbl_TipoIdent = Trim(vlRegistro!GLS_TIPOIDENCOR)
    
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
        
        Lbl_PrimaDef = Format((vlRegistro!MTO_PRIUNI), "#,#0.00")
        Lbl_PensionDef = Format((vlRegistro!Mto_Pension), "#,#0.00")
        'vlMtoPenGar = Format((vlRegistro!Mto_PensionGar), "#,#0.00")
        
        vlMonedaPension = Trim(vlRegistro!Cod_Moneda)
        Lbl_Moneda(ciMonedaPensionNew) = vlMonedaPension
        
        Lbl_FechaRec = DateSerial(Mid(vlRegistro!fec_traspaso, 1, 4), Mid(vlRegistro!fec_traspaso, 5, 2), Mid(vlRegistro!fec_traspaso, 7, 2))
        Lbl_NumLiquidacion = Format(vlRegistro!cod_renvit, "000") & " - " & Format(vlRegistro!cod_liquidacion, "0000000")
        vlafp = vlRegistro!Cod_AFP
        Call flObtieneFecNacConyuge(Trim(Txt_Poliza), vlFecNacConyuge) 'Fecha Nacimiento Conyuge
        Cmd_Poliza.Enabled = False
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


    On Error GoTo Errores1
   
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
    
    vgQuery = "{PD_TMAE_POLIZA.NUM_POLIZA} = '" & Txt_Poliza & "'"
    
    Rpt_Reporte.Reset
    Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\rpt_Areas.rpt"
    Rpt_Reporte.Connect = vgRutaDataBase
    Rpt_Reporte.Formulas(0) = ""
    Rpt_Reporte.Formulas(1) = ""
    Rpt_Reporte.Formulas(2) = ""
    Rpt_Reporte.Formulas(3) = ""
    Rpt_Reporte.Formulas(4) = ""
    Rpt_Reporte.Formulas(5) = ""
    Rpt_Reporte.Formulas(6) = ""
    Rpt_Reporte.Formulas(7) = ""
    Rpt_Reporte.Formulas(8) = ""
    Rpt_Reporte.Formulas(9) = ""
    Rpt_Reporte.Formulas(10) = ""
    Select Case Index
        Case ciImprimirBienvenida
'Reporte de Bienvenida
'----------------------------------------------------
            vgQuery = vgQuery & " AND (({PD_TMAE_POLIZA.COD_TIPPENSION} = '" & clCodTipPensionSob & "' AND {PD_TMAE_POLBEN.COD_DERPEN} <> '" & Trim(vlCodDerPen) & "')"  'Sobrevivencia con Derecho a Pension
            vgQuery = vgQuery & " OR ({PD_TMAE_POLIZA.COD_TIPPENSION} <> '" & clCodTipPensionSob & "' AND {PD_TMAE_POLBEN.COD_PAR} = '" & Trim(vlCodPar) & "'))" 'o Solo los Causantes de Invalidez o Vejez
            Rpt_Reporte.Formulas(0) = "NombreCompaniaCorto = '" & vgNombreCortoCompania & "'"
            Rpt_Reporte.Formulas(1) = "Nombre= '" & vgNombreApoderado & "'"
            Rpt_Reporte.Formulas(2) = "Cargo= '" & vgCargoApoderado & "'"
            Rpt_Reporte.Formulas(3) = "Sucursal = '" & vlNombreSucursal & "'"
            Rpt_Reporte.WindowTitle = "Carta de Bienvenida"
            
        Case ciImprimirPoliza
            'vgQuery = vgQuery & " AND {PD_TMAE_AFILIADO.COD_PAR} = '" & Trim(vlCodPar) & "'" 'Afiliado
            Rpt_Reporte.Formulas(0) = "NombreAfi = '" & Lbl_NomAfiliado & "'"
            Rpt_Reporte.Formulas(1) = "TipoPension = '" & vlNombreTipoPension & "'"
            Rpt_Reporte.Formulas(2) = "MesGar = '" & Lbl_Meses & "'"
            Rpt_Reporte.Formulas(3) = "NombreCompania = '" & UCase(vgNombreCompania) & "'"
            Rpt_Reporte.Formulas(4) = "Concatenar = '" & vlCobertura & "'"
            Rpt_Reporte.Formulas(5) = "Sucursal = '" & vlNombreSucursal & "'"
            'RVF 20090914
            Call pBuscaRepresentante(vlNombreTipoPension)
            Rpt_Reporte.Formulas(6) = "RepresentanteNom = '" & vlRepresentante & "'"
            Rpt_Reporte.Formulas(7) = "RepresentanteDoc = '" & vlDocum & "'"
            Rpt_Reporte.Formulas(8) = "CodTipPen = '" & Left(Trim(Lbl_TipoPension), 2) & "'"
            '*****
            'RVF 20100121
            Rpt_Reporte.Formulas(9) = "TipoDocTit = '" & Trim(Lbl_TipoIdent.Caption) & "'"
            Rpt_Reporte.Formulas(10) = "NumDocTit = '" & Trim(Lbl_NumIdent.Caption) & "'"
            '*****
            Rpt_Reporte.WindowTitle = "Póliza"
        Case ciImprimirAFP
            
            'Dim objRep As New ClsReporte
            'Dim strQuery, vlFecTras As String
            'Dim RS As ADODB.Recordset
            Set RS = New ADODB.Recordset
            vgQuery = "select a.num_poliza, gls_nomben, gls_nomsegben, gls_patben, gls_matben, c.cod_scomp,"
            vgQuery = vgQuery & " d.gls_elemento as TipoAFP, e.gls_elemento as TipoPension, f.gls_elemento as TipoMoneda"
            vgQuery = vgQuery & " from pd_tmae_poliza a"
            vgQuery = vgQuery & " join pd_tmae_polben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
            vgQuery = vgQuery & " join ma_tpar_monedatiporeaju c on a.cod_moneda=c.cod_moneda and a.cod_tipreajuste=c.cod_tipreajuste"
            vgQuery = vgQuery & " join ma_tpar_tabcod d on a.cod_afp=d.cod_elemento and d.cod_tabla='AF'"
            vgQuery = vgQuery & " join ma_tpar_tabcod e on a.cod_tippension=e.cod_elemento and e.cod_tabla='TP'"
            vgQuery = vgQuery & " join ma_tpar_tabcod f on a.cod_moneda=f.cod_elemento and f.cod_tabla='TM'"
            vgQuery = vgQuery & " where a.num_endoso=" & CInt(Txt_Endoso) & ""  '(select max(num_endoso) from pd_tmae_poliza where num_poliza=a.num_poliza)"
            vgQuery = vgQuery & " and cod_par='" & Trim(vlCodPar) & "' and a.num_poliza='" & Txt_Poliza & "' "
            'vgQuery = vgQuery & " order by 1"
            RS.Open vgQuery, vgConexionBD, adOpenForwardOnly, adLockReadOnly
        
            If RS.EOF Then
                MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
                Exit Sub
            End If
            Dim vlNombreComuna As String, vlJefeBeneficios As String, vlDireccion As String
            Call flObtieneDatosContacto(vlafp, vlNombreComuna, vlJefeBeneficios, vlDireccion)
        
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
        
'            vgQuery = vgQuery & " AND {PD_TMAE_POLBEN.COD_PAR} = '" & Trim(vlCodPar) & "'"
'            vgQuery = vgQuery & " AND {PD_TMAE_POLIZA.NUM_ENDOSO}=1" 'Solo los Causantes
'            Dim vlNombreComuna As String, vlJefeBeneficios As String, vlDireccion As String
'            Call flObtieneDatosContacto(vlafp, vlNombreComuna, vlJefeBeneficios, vlDireccion)
'            Rpt_Reporte.Formulas(0) = "Nombre= '" & vgNombreApoderado & "'"
'            Rpt_Reporte.Formulas(1) = "Cargo= '" & vgCargoApoderado & "'"
'            Rpt_Reporte.Formulas(2) = "Direccion = '" & vlDireccion & "'"
'            Rpt_Reporte.Formulas(3) = "Distrito = '" & vlNombreComuna & "'"
'            Rpt_Reporte.Formulas(4) = "ContactoAFP = '" & vlJefeBeneficios & "'"
'            Rpt_Reporte.Formulas(5) = "Sucursal = 'Surquillo'"
'            Rpt_Reporte.Formulas(6) = "Fecha = '" & Lbl_FechaRec & "'" 'MC - 18/03/2008
'            Rpt_Reporte.WindowTitle = "Carta AFP"
            
        Case ciImprimirVariacion
'Reporte de Variacion de Primas/Pensión
'----------------------------------------------------
            Rpt_Reporte.Formulas(0) = "TipoRenta = '" & Lbl_TipoRenta & "'"
            Rpt_Reporte.Formulas(1) = "FecNacTitular = '" & vlFecNacTitular & "'"
            Rpt_Reporte.Formulas(2) = "FecNacConyuge = '" & vlFecNacConyuge & "'"
            Rpt_Reporte.Formulas(3) = "NombreCompaniaCorto = '" & Mid(vgNombreCortoCompania, 1, 1) & LCase(Mid(vgNombreCortoCompania, 2)) & "'"
'I--- ABV 05/02/2011 ---
'            Rpt_Reporte.Formulas(4) = "MonedaPension = '" & vlMonedaPension & "'"
            Rpt_Reporte.Formulas(4) = "MonedaPension = '" & vlCodTipReajusteScomp & "'"
'F--- ABV 05/02/2011 ---
            Rpt_Reporte.Formulas(5) = "Concatenar = '" & vlCobertura & "'"
            Rpt_Reporte.WindowTitle = "Variación Pensión"
            
        Case ciImprimirPrimerPago
            If flVerificaPrimerPago(Txt_Poliza, vlMonto, vlMoneda) Then
                
                'Rpt_Reporte.Formulas(0) = "MontoPalabras = '" & fgConvierteNumeroLetras(vlMonto, vlMoneda) & "'"
                Rpt_Reporte.Formulas(0) = ""
                Rpt_Reporte.Formulas(1) = ""
                Rpt_Reporte.Formulas(2) = ""
                Rpt_Reporte.Formulas(3) = "NombreCompania = '" & UCase(vgNombreCompania) & "'"
                Rpt_Reporte.Formulas(4) = "rutcliente = '" & vgNumIdenCliente & "'"
                Rpt_Reporte.Formulas(5) = ""
                Rpt_Reporte.WindowTitle = "Liquidación Primer Pago"
            Else
                Screen.MousePointer = 0
                Exit Sub
            End If
        Case ciImprimirPrima
            'Dim objRep As New ClsReporte
            'Dim strQuery, vlFecTras As String
            'Dim RS As ADODB.Recordset
            
            vlFecTras = Lbl_FechaRec.Caption
            vlFecTras = Mid(vlFecTras, 7, 4) & Mid(vlFecTras, 4, 2) & Mid(vlFecTras, 1, 2)
        
            Set RS = New ADODB.Recordset
            strQuery = "select a.num_poliza, a.num_endoso,a.gls_direccion, mto_priuni, gls_nacionalidad, fec_inipencia,"
            strQuery = strQuery & " b.Fec_FallBen , gls_comuna, gls_provincia, gls_region, cod_tippension, a.fec_crea"
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
            strQuery = strQuery & " Where b.Fec_FallBen Is Null and a.num_poliza='" & Txt_Poliza & "' and a.num_endoso='" & Txt_Endoso & "' " '(select max(num_endoso) from pd_tmae_poliza where num_poliza=a.num_poliza)"
            strQuery = strQuery & " group by a.num_poliza, a.num_endoso,a.gls_direccion, mto_priuni, gls_nacionalidad, fec_inipencia,"
            strQuery = strQuery & " b.Fec_FallBen , gls_comuna, gls_provincia, gls_region, cod_tippension, a.fec_crea"
            strQuery = strQuery & " order by 2"
            
            RS.Open strQuery, vgConexionBD, adOpenForwardOnly, adLockReadOnly
        
            If RS.EOF Then
                MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
                Exit Sub
            End If
            
            Call pBuscaRepresentante(Lbl_TipoPension)
        
        
        
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
        
        ',
        '                        ArrFormulas("Asegurado", Lbl_NomAfiliado), _
        '                        ArrFormulas("Concatenar", vlCobertura), _
        '                        ArrFormulas("NombreCompania", UCase(vgNombreCompania)), _
        '                        ArrFormulas("NroLiquidacion", vgPalabra), _
        '                        ArrFormulas("Sucursal", vlNombreSucursal), _
        '                        ArrFormulas("NroBoleta", Lbl_NumLiquidacion)
        
        'GoTo Imprime
        
            'vgPalabra = fgObtenerNombre_TextoCompuesto(Lbl_NumLiquidacion)
            'If IsNumeric(vgPalabra) Then vgPalabra = CDbl(vgPalabra)
            'Rpt_Reporte.Formulas(0) = "Contratante = '" & Lbl_NomAfiliado & "'"
            'Rpt_Reporte.Formulas(1) = "Asegurado = '" & Lbl_NomAfiliado & "'"
            'Rpt_Reporte.Formulas(2) = "Concatenar = '" & vlCobertura & "'"
            'Rpt_Reporte.Formulas(3) = "NombreCompania = '" & UCase(vgNombreCompania) & "'"
'            Rpt_Reporte.Formulas(4) = "NroLiquidacion = '" & Lbl_NumLiquidacion & "'"
            'Rpt_Reporte.Formulas(4) = "NroLiquidacion = '" & vgPalabra & "'"
            'Rpt_Reporte.Formulas(5) = "Sucursal = '" & vlNombreSucursal & "'"
            'Rpt_Reporte.Formulas(6) = "NroBoleta = '" & Lbl_NumLiquidacion & "'"
            'Rpt_Reporte.WindowTitle = "Liquidación de Prima"
        Case ciImprimirFactura
            vlMonto = CDbl(Lbl_PrimaDef)
            vlMoneda = cgCodTipMonedaUF
            'For vgI = 1 To 3
                Rpt_Reporte.Formulas(0) = "IdentificacionEmpresa = '" & vgNumIdenCompania & "'"
                Rpt_Reporte.Formulas(1) = "Corredor='" & vlCorredor & "'"
                Rpt_Reporte.Formulas(2) = "Asegurado = '" & Lbl_NomAfiliado & "'"
                Rpt_Reporte.Formulas(3) = "Concatenar = '" & vlCobertura & "'"
                Rpt_Reporte.Formulas(4) = "NroLiquidacion = '" & Lbl_NumLiquidacion & "'"
                '" & vlIdenCor & "' + " " & vlIdenCor & "'"
                ' Lbl_NomAfiliado = Trim(vlRegistro!Gls_NomBen) + " " + Trim(vlRegistro!Gls_PatBen) + " " + Trim(vlRegistro!Gls_MatBen)
                'Rpt_Reporte.Formulas(3) = "NombreCompania = '" & UCase(vgNombreCompania) & "'"
                'Rpt_Reporte.Formulas(2) = "Copia1='ADQUIRIENTE'"
                'Rpt_Reporte.Formulas(3) = "Copia2='SUNAT'"
                'Rpt_Reporte.Formulas(6) = "Copia3='EMISOR'"
                Rpt_Reporte.Formulas(5) = "Sucursal = '" & vlNombreSucursal & "'"
                'Rpt_Reporte.Formulas(6) = "MontoPalabras = '" & vlNombreSucursal & "'"
                Rpt_Reporte.Formulas(6) = "MontoPalabras = '" & fgConvierteNumeroLetras(vlMonto, vlMoneda) & "'"
                Rpt_Reporte.Formulas(7) = "IdentificacionAfiliado = '" & Lbl_TipoIdent & " - " & Lbl_NumIdent & "'"
                
'                Select Case vgI
'                    Case 1:
'                        Rpt_Reporte.Formulas(6) = "Copia = 'Cliente'"
'                        Rpt_Reporte.WindowTitle = "Boleta de Venta - Cliente"
'                    Case 2:
'                        Rpt_Reporte.Formulas(6) = "Copia = 'Corredor'"
'                        Rpt_Reporte.WindowTitle = "Boleta de Venta - Corredor"
'                    Case 3:
'                        Rpt_Reporte.Formulas(6) = "Copia = 'Compañía'"
'                        Rpt_Reporte.WindowTitle = "Boleta de Venta - Compañía"
'                End Select
            
                Rpt_Reporte.WindowTitle = "Boleta de Venta "
                Rpt_Reporte.SelectionFormula = vgQuery
                Rpt_Reporte.Destination = crptToWindow
                Rpt_Reporte.WindowState = crptMaximized
                Rpt_Reporte.Action = 1
            'Next vgI
            
            Screen.MousePointer = 0
            
            Exit Sub
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
Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

    Txt_Poliza = ""
    Lbl_TipoIdent = ""
    Lbl_NumIdent = ""
    Lbl_NomAfiliado = ""
    Lbl_CUSPP = ""
    Lbl_TipoPension = ""
    Lbl_TipoRenta = ""
    Lbl_Diferidos = ""
    Lbl_Modalidad = ""
    Lbl_Meses = ""
    Txt_Endoso = ""
'I--- ABV 05/02/2011 ---
    Lbl_ReajusteTipo = ""
    Lbl_ReajusteValor = ""
    Lbl_ReajusteValorMen = ""
    Lbl_ReajusteDescripcion = ""
'F--- ABV 05/02/2011 ---
    Lbl_PrimaDef = ""
    Lbl_FechaRec = ""
    Lbl_PensionDef = ""
    Lbl_NumLiquidacion = ""
    
    Fra_Poliza.Enabled = True
    Cmd_Poliza.Enabled = True
    Txt_Poliza.SetFocus

Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Sub

Private Sub Cmd_Poliza_Click()
On Error GoTo Err_BuscarPoliza

    Frm_BuscarPolEnd.flInicio ("Frm_CalPrimaInf")
    
Exit Sub
Err_BuscarPoliza:
Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Salir_Click()
On Error GoTo Err_CmdSalir
    
    Screen.MousePointer = 11
    Unload Me
    Screen.MousePointer = 0
    
Exit Sub
Err_CmdSalir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


Private Sub cmdBoleta_Click()
Dim vlCodPar As String
Dim vlCodDerPen As String
Dim vlArchivo As String
Dim vlNombreSucursal As String, vlNombreTipoPension As String
Dim vlMonto As Double, vlMoneda As String
Dim RS As ADODB.Recordset

    On Error GoTo mierror
   
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
    
    Screen.MousePointer = vbHourglass
    vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
    vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_TipoPension)
    vlMonto = CDbl(Lbl_PrimaDef)
    vlMoneda = cgCodTipMonedaUF

    
    

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "PD_LISTA_POLIZA_FACTURA.LISTAR('" & Txt_Poliza.Text & "', '" & Txt_Endoso & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaBoleta.rpt"), ".RPT", ".TTX"), 1)
    'Lbl_NomAfiliado
    If objRep.CargaReporte(strRpt & "", "PD_Rpt_PolizaFactura.rpt", "Boleta de Venta", RS, True, _
                            ArrFormulas("IdentificacionEmpresa", vgNumIdenCompania), _
                            ArrFormulas("Corredor", vlCorredor), _
                            ArrFormulas("Asegurado", Lbl_NomAfiliado.Caption), _
                            ArrFormulas("Concatenar", vlCobertura), _
                            ArrFormulas("NroLiquidacion", Lbl_NumLiquidacion.Caption), _
                            ArrFormulas("Sucursal", vlNombreSucursal), _
                            ArrFormulas("MontoPalabras", fgConvierteNumeroLetras(vlMonto, vlMoneda)), _
                            ArrFormulas("IdentificacionAfiliado", Lbl_TipoIdent.Caption & " - " & Lbl_NumIdent.Caption)) = False Then
            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
Exit Sub
mierror:
    MsgBox "No pudo cargar el reporte " & Err.Description, vbInformation
End Sub


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

Private Sub cmdPoliza_Click()
Dim vlCodPar As String
Dim vlCodDerPen As String
Dim vlNombreSucursal, vlNombreTipoPension As String
Dim RS As ADODB.Recordset
Dim vlFecTras As String
Dim vlTipRen As String
Dim NomReporte As String
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
    vlTipRen = Mid(Lbl_TipoRenta.Caption, 1, 1)
    
    If vlTipRen <> "6" Then
        NomReporte = "PD_Rpt_PolizaDef.rpt"
    Else
        NomReporte = "PD_Rpt_PolizaDefEsc.rpt"
    End If
    
On Error GoTo mierror

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "PD_LISTA_POLIZA.LISTAR('" & Txt_Poliza.Text & "', '" & Txt_Endoso & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
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
    
    If vlTipRen = "6" Then
        Set RS = New ADODB.Recordset
        RS.CursorLocation = adUseClient
        RS.Open "PD_LISTA_POLIZA.LISTAR('" & Txt_Poliza.Text & "', '" & Txt_Endoso & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
        
        'Dim LNGa As Long
        LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaConsta.rpt"), ".RPT", ".TTX"), 1)
        
        If objRep.CargaReporte(strRpt & "", "PD_Rpt_PolizaDefEscRes.rpt", "Póliza Resumen", RS, True, _
                                ArrFormulas("NombreAfi", Lbl_NomAfiliado.Caption)) = False Then
                
                
            MsgBox "No se pudo abrir el reporte", vbInformation
            Exit Sub
        End If
        'Exit Sub
    End If
    
    '''******************************************CONDTSNCIA DE POLIZA*********************************************************************
    
    Dim strTipoPension As String
    strTipoPension = Mid(Lbl_TipoPension.Caption, 1, 2)
    If strTipoPension >= "08" Then
        Exit Sub
    End If
    
    
    
    If CInt(Lbl_Diferidos.Caption) > 0 Then
        NomReporte = "PD_Rpt_PolizaConstaDif.rpt"
    Else
        NomReporte = "PD_Rpt_PolizaConstaDif.rpt"
    End If
    
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "PD_LISTA_POLIZA.LISTAR('" & Txt_Poliza.Text & "', '" & Txt_Endoso & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
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
    'vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_TipoPension)
    vlFecTras = Lbl_FechaRec.Caption
    
    vlFecTras = Mid(vlFecTras, 7, 4) & Mid(vlFecTras, 4, 2) & Mid(vlFecTras, 1, 2)
    
On Error GoTo mierror

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "PP_LISTA_BIENVENIDA.LISTAR('" & Txt_Poliza.Text & "','" & clCodTipPensionSob & "','" & Trim(vlCodDerPen) & "','" & Trim(vlCodPar) & "', '" & Trim(Txt_Endoso) & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
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
Function fgConvierteNumeroLetras(iMonto As Double, Optional iMoneda As String) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlDecimales As Double
Dim vlEntero As Double
vlEntero = Fix(iMonto)
vlDecimales = Format((iMonto - vlEntero) * 100, "#0.00")
vlMontoPalabras = fgConvierteMillones(vlEntero)
If iMonto > 2 Then
    If Mid(vlMontoPalabras, 1, 3) = "uno" Then
        vlMontoPalabras = Mid(vlMontoPalabras, 1, 2) + Mid(vlMontoPalabras, 4)
    End If
End If
If vlDecimales > 0 Then
    vlMontoPalabras = vlMontoPalabras + " con " & vlDecimales & "/100"
End If
If Not IsMissing(iMoneda) Then
    vlMontoPalabras = vlMontoPalabras + " " & iMoneda
End If
fgConvierteNumeroLetras = UCase(vlMontoPalabras)

End Function
Function fgConvierteMillones(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlFraccion As Double
Dim vlCantidad As Double
vlFraccion = iMonto Mod 1000000
vlCantidad = Int(iMonto / 1000000) 'Parte Entera
Select Case iMonto
    Case 1000000
        vlMontoPalabras = "un millón"
    Case Is > 999999
        vlMontoPalabras = Trim(fgConvierteMiles(vlCantidad) + " millones " + fgConvierteMiles(vlFraccion))
    Case Else
        vlMontoPalabras = fgConvierteMiles(iMonto)
End Select
fgConvierteMillones = vlMontoPalabras
End Function
Function fgConvierteMiles(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlFraccion As Double
Dim vlCantidad As Double
vlFraccion = iMonto Mod 1000
vlCantidad = Int(iMonto / 1000) 'Parte Entera
Select Case iMonto
    Case 1000
        vlMontoPalabras = "mil"
    Case Is > 999
        vlMontoPalabras = Trim(fgConvierteCentenas(vlCantidad) + " mil " + fgConvierteCentenas(vlFraccion))
    Case Else
        vlMontoPalabras = fgConvierteCentenas(iMonto)
End Select
fgConvierteMiles = vlMontoPalabras
End Function
Function fgConvierteCentenas(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlDecena As Double
vlDecena = iMonto Mod 100
Select Case iMonto
    Case Is > 900
        vlMontoPalabras = "novecientos " + fgConvierteDecenas(vlDecena)
    Case 900
        vlMontoPalabras = "novecientos"
    Case Is > 800
        vlMontoPalabras = "ochocientos " + fgConvierteDecenas(vlDecena)
    Case 800
        vlMontoPalabras = "ochocientos"
    Case Is > 700
        vlMontoPalabras = "setecientos " + fgConvierteDecenas(vlDecena)
    Case 700
        vlMontoPalabras = "setecientos"
    Case Is > 600
        vlMontoPalabras = "seiscientos " + fgConvierteDecenas(vlDecena)
    Case 600
        vlMontoPalabras = "seiscientos"
    Case Is > 500
        vlMontoPalabras = "quinientos " + fgConvierteDecenas(vlDecena)
    Case 500
        vlMontoPalabras = "quinientos"
    Case Is > 400
        vlMontoPalabras = "cuatrocientos " + fgConvierteDecenas(vlDecena)
    Case 400
        vlMontoPalabras = "cuatrocientos"
    Case Is > 300
        vlMontoPalabras = "trescientos " + fgConvierteDecenas(vlDecena)
    Case 300
        vlMontoPalabras = "trescientos"
    Case Is > 200
        vlMontoPalabras = "doscientos " + fgConvierteDecenas(vlDecena)
    Case 200
        vlMontoPalabras = "doscientos"
    Case Is > 100
        vlMontoPalabras = "ciento " + fgConvierteDecenas(vlDecena)
    Case 100
        vlMontoPalabras = "cien"
    Case Is < 100
        vlMontoPalabras = fgConvierteDecenas(iMonto)
End Select
fgConvierteCentenas = vlMontoPalabras
End Function
Function fgConvierteDecenas(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlDigito As Double
vlDigito = iMonto Mod 10
Select Case iMonto
    Case Is > 90
        vlMontoPalabras = "noventa y " + fgConvierteDigito(vlDigito)
    Case 90
        vlMontoPalabras = "noventa"
    Case Is > 80
        vlMontoPalabras = "ochenta y " + fgConvierteDigito(vlDigito)
    Case 80
        vlMontoPalabras = "ochenta"
    Case Is > 70
        vlMontoPalabras = "setenta y " + fgConvierteDigito(vlDigito)
    Case 70
        vlMontoPalabras = "setenta"
    Case Is > 60
        vlMontoPalabras = "sesenta y " + fgConvierteDigito(vlDigito)
    Case 60
        vlMontoPalabras = "sesenta"
    Case Is > 50
        vlMontoPalabras = "cincuenta y " + fgConvierteDigito(vlDigito)
    Case 50
        vlMontoPalabras = "cincuenta"
    Case Is > 40
        vlMontoPalabras = "cuarenta y " + fgConvierteDigito(vlDigito)
    Case 40
        vlMontoPalabras = "cuarenta"
    Case Is > 30
        vlMontoPalabras = "treinta y " + fgConvierteDigito(vlDigito)
    Case 30
        vlMontoPalabras = "treinta"
    Case Is > 20
        vlMontoPalabras = "veinti" + fgConvierteDigito(vlDigito)
    Case Is <= 20
        vlMontoPalabras = fgConvierteDigito(iMonto)
End Select
fgConvierteDecenas = vlMontoPalabras
End Function

Function fgConvierteDigito(iDigito As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Select Case iDigito
    Case 1
        vlMontoPalabras = "uno"
    Case 2
        vlMontoPalabras = "dos"
    Case 3
        vlMontoPalabras = "tres"
    Case 4
        vlMontoPalabras = "cuatro"
    Case 5
        vlMontoPalabras = "cinco"
    Case 6
        vlMontoPalabras = "seis"
    Case 7
        vlMontoPalabras = "siete"
    Case 8
        vlMontoPalabras = "ocho"
    Case 9
        vlMontoPalabras = "nueve"
    Case 10
        vlMontoPalabras = "diez"
    Case 11
        vlMontoPalabras = "once"
    Case 12
        vlMontoPalabras = "doce"
    Case 13
        vlMontoPalabras = "trece"
    Case 14
        vlMontoPalabras = "catorce"
    Case 15
        vlMontoPalabras = "quince"
    Case 16
        vlMontoPalabras = "dieciseis"
    Case 17
        vlMontoPalabras = "diecisiete"
    Case 18
        vlMontoPalabras = "dieciocho"
    Case 19
        vlMontoPalabras = "diecinueve"
    Case 20
        vlMontoPalabras = "veinte"
End Select
fgConvierteDigito = vlMontoPalabras
End Function
Private Sub cmdReportPrimerPago_Click()

Dim RS As ADODB.Recordset

On Error GoTo mierror

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "PD_LISTA_PRIMER_PAGO.LISTAR('" & Txt_Poliza.Text & "','" & Txt_Endoso & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    'Dim LNGa As Long
    'LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_LiquidacionRV_PP.rpt"), ".RPT", ".TTX"), 1)
    
        
    'If objRep.CargaReporte(strRpt & "", "PD_Rpt_LiquidacionRV.rpt", "Liquidación Primer Pago", RS, True, _
                            ArrFormulas("NombreCompania", UCase(vgNombreCompania)), _
                            ArrFormulas("rutcliente", vgNumIdenCliente)) = False Then
                            
    '    MsgBox "No se pudo abrir el reporte", vbInformation
    '    Exit Sub
    'End If
    If Not RS.EOF Then
        Call p_crear_rs
        sName_Reporte = "PD_Rpt_LiquidacionRV.rpt"
        'XX = fgConvierteNumeroLetras
        Do Until RS.EOF
            GoSub p_asiga_valores
            RS.MoveNext
        Loop
        frm_plantilla.Show
    End If
    
Exit Sub

p_asiga_valores:
'***************
    objRsRpt.AddNew
    objRsRpt.Fields("NUM_IDENAFI").Value = Trim(RS!num_idenafi)
    objRsRpt.Fields("GLS_DIRECCION").Value = Trim(RS!Gls_Direccion)
    objRsRpt.Fields("NUM_PERPAGO").Value = Trim(RS!Num_PerPago)
    objRsRpt.Fields("NUM_POLIZA").Value = Trim(RS!Num_Poliza)
    objRsRpt.Fields("NUM_ORDEN").Value = Trim(RS!Num_Orden)
    objRsRpt.Fields("FEC_PAGO").Value = Trim(RS!Fec_Pago)
    objRsRpt.Fields("COD_VIAPAGO").Value = Trim(RS!Cod_ViaPago)
    objRsRpt.Fields("NUM_IDENRECEPTOR").Value = Trim(RS!Num_IdenReceptor)
    objRsRpt.Fields("GLS_NOMRECEPTOR").Value = Trim(RS!Gls_NomReceptor)
    objRsRpt.Fields("GLS_NOMSEGRECEPTOR").Value = IIf(IsNull(Trim(RS!Gls_NomSegReceptor)), "", Trim(RS!Gls_NomSegReceptor))
    objRsRpt.Fields("GLS_PATRECEPTOR").Value = Trim(RS!Gls_PatReceptor)
    objRsRpt.Fields("GLS_MATRECEPTOR").Value = Trim(RS!Gls_MatReceptor)
    objRsRpt.Fields("COD_TIPRECEPTOR").Value = Trim(RS!Cod_TipReceptor)
    objRsRpt.Fields("MTO_HABER").Value = Trim(RS!Mto_Haber)
    objRsRpt.Fields("MTO_DESCUENTO").Value = Trim(RS!Mto_Descuento)
    objRsRpt.Fields("MTO_LIQPAGAR").Value = Trim(RS!Mto_LiqPagar)
    objRsRpt.Fields("COD_MONEDA").Value = Trim(RS!Cod_Moneda)
    objRsRpt.Fields("NUM_IDENBEN").Value = Trim(RS!Num_IdenBen)
    objRsRpt.Fields("GLS_NOMBEN").Value = Trim(RS!Gls_NomBen)
    objRsRpt.Fields("GLS_NOMSEGBEN").Value = IIf(IsNull(Trim(RS!Gls_NomSegBen)), "", Trim(RS!Gls_NomSegBen))
    objRsRpt.Fields("GLS_PATBEN").Value = Trim(RS!Gls_PatBen)
    objRsRpt.Fields("GLS_MATBEN").Value = Trim(RS!Gls_MatBen)
    objRsRpt.Fields("TIPDOC").Value = Trim(RS!TipDoc)
    objRsRpt.Fields("GLS_SUCURSAL").Value = Trim(RS!gls_sucursal)
    objRsRpt.Fields("GLS_COMUNA").Value = Trim(RS!gls_comuna)
    objRsRpt.Fields("GLS_PROVINCIA").Value = Trim(RS!gls_provincia)
    objRsRpt.Fields("GLS_REGION").Value = Trim(RS!gls_region)
    objRsRpt.Fields("NOM_ASEGURADO").Value = Trim(RS!NOM_ASEGURADO)
    objRsRpt.Fields("DOC_ASEGURADO").Value = Trim(RS!DOC_ASEGURADO)
    objRsRpt.Fields("TIPO_DOC_ASEGURADO").Value = Trim(RS!TIPO_DOC_ASEGURADO)
    objRsRpt.Fields("TIPO_PENSION").Value = Trim(RS!TIPO_PENSION)
    objRsRpt.Fields("TIPO_PAGO").Value = Trim(RS!TIPO_PAGO)
    If IsNull(RS!AFP) Then
        objRsRpt.Fields("AFP").Value = Trim(RS!TIPO_PAGO)
    Else
        objRsRpt.Fields("AFP").Value = Trim(RS!AFP)
    End If
    objRsRpt.Fields("TUTOR").Value = Trim(RS!TUTOR)
    objRsRpt.Fields("COD_SCOMP").Value = Trim(RS!cod_scomp)
    objRsRpt.Fields("DESC_MONEDA_AJUSTE").Value = Trim(RS!DESC_MONEDA_AJUSTE)
    If IsNull(RS!GLS_VEJEZ) Then
        objRsRpt.Fields("GLS_VEJEZ").Value = ""
    Else
        objRsRpt.Fields("GLS_VEJEZ").Value = Trim(RS!GLS_VEJEZ)
    End If
    objRsRpt.Fields("LETRAMTO").Value = fgConvierteNumeroLetras(RS!LETRAMTO, IIf(Trim(RS!Cod_Moneda) = "NS", "NUEVOS SOLES", "DOLARES"))
    objRsRpt.Update
    
    Return
    
mierror:
    MsgBox "No pudo cargar el reporte " & Err.Description, vbInformation
    
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Me.Top = 0
    Me.Left = 0
    Lbl_Moneda(ciMonedaPrimaNew) = cgCodTipMonedaUF
    Lbl_Moneda(ciMonedaPensionNew) = cgCodTipMonedaUF
    vgUsuarioSuc = "0000"
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Endoso_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Poliza
   
   If KeyAscii = 13 Then
        Cmd_Buscar.SetFocus
   End If

Exit Sub
Err_Poliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Poliza_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Poliza
   
   If KeyAscii = 13 Then
      If Trim(Txt_Poliza) > "" Then
         Txt_Poliza = Trim(UCase(Txt_Poliza))
         Txt_Poliza = Format(Txt_Poliza, "0000000000")
         'RRR
         
         Dim vlRegistro As New ADODB.Recordset
        
           vgSql = "SELECT MAX(NUM_ENDOSO) AS NUM_ENDOSO FROM PD_TMAE_POLIZA WHERE NUM_POLIZA='" & Txt_Poliza & "'"
        
           Set vlRegistro = vgConexionBD.Execute(vgSql)
           If Not (vlRegistro.EOF) Then
               Txt_Endoso = vlRegistro!Num_Endoso
           End If
         
         
         Txt_Endoso.SetFocus
         'Cmd_Buscar.SetFocus
      End If
   End If

Exit Sub
Err_Poliza:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Poliza_LostFocus()
    If Txt_Poliza > "" Then
        Txt_Poliza = Trim(UCase(Txt_Poliza))
        Txt_Poliza = Format(Txt_Poliza, "0000000000")
    End If
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
        vlDocum = IIf(Not IsNull(vgRs!GLS_TIPOIDENCOR), vgRs!GLS_TIPOIDENCOR, "") & " " & IIf(Not IsNull(vgRs!Num_idenrep), vgRs!Num_idenrep, "")
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



Private Sub p_crear_rs()
   
   Set objRsRpt = New ADODB.Recordset
   
   objRsRpt.Fields.Append "NUM_IDENAFI", adVarChar, 16
   objRsRpt.Fields.Append "GLS_DIRECCION", adVarChar, 120
   objRsRpt.Fields.Append "NUM_PERPAGO", adVarChar, 6
   objRsRpt.Fields.Append "NUM_POLIZA", adVarChar, 10
   objRsRpt.Fields.Append "NUM_ORDEN", adInteger
   objRsRpt.Fields.Append "FEC_PAGO", adVarChar, 10
   objRsRpt.Fields.Append "COD_VIAPAGO", adVarChar, 5
   objRsRpt.Fields.Append "NUM_IDENRECEPTOR", adVarChar, 16
   objRsRpt.Fields.Append "GLS_NOMRECEPTOR", adVarChar, 25
   objRsRpt.Fields.Append "GLS_NOMSEGRECEPTOR", adVarChar, 25
   objRsRpt.Fields.Append "GLS_PATRECEPTOR", adVarChar, 20
   objRsRpt.Fields.Append "GLS_MATRECEPTOR", adVarChar, 20
   objRsRpt.Fields.Append "COD_TIPRECEPTOR", adVarChar, 1
   objRsRpt.Fields.Append "MTO_HABER", adDouble
   objRsRpt.Fields.Append "MTO_DESCUENTO", adDouble
   objRsRpt.Fields.Append "MTO_LIQPAGAR", adDouble
   objRsRpt.Fields.Append "COD_MONEDA", adVarChar, 2
   objRsRpt.Fields.Append "NUM_IDENBEN", adVarChar, 16
   objRsRpt.Fields.Append "GLS_NOMBEN", adVarChar, 20
   objRsRpt.Fields.Append "GLS_NOMSEGBEN", adVarChar, 20
   objRsRpt.Fields.Append "GLS_PATBEN", adVarChar, 20
   objRsRpt.Fields.Append "GLS_MATBEN", adVarChar, 20
   objRsRpt.Fields.Append "TIPDOC", adVarChar, 10
   objRsRpt.Fields.Append "GLS_SUCURSAL", adVarChar, 50
   objRsRpt.Fields.Append "GLS_COMUNA", adVarChar, 50
   objRsRpt.Fields.Append "GLS_PROVINCIA", adVarChar, 50
   objRsRpt.Fields.Append "GLS_REGION", adVarChar, 50
   objRsRpt.Fields.Append "NOM_ASEGURADO", adVarChar, 83
   objRsRpt.Fields.Append "DOC_ASEGURADO", adVarChar, 16
   objRsRpt.Fields.Append "TIPO_DOC_ASEGURADO", adVarChar, 10
   objRsRpt.Fields.Append "TIPO_PENSION", adVarChar, 50
   objRsRpt.Fields.Append "TIPO_PAGO", adVarChar, 50
   objRsRpt.Fields.Append "AFP", adVarChar, 50
   objRsRpt.Fields.Append "TUTOR", adVarChar, 88
   objRsRpt.Fields.Append "COD_SCOMP", adVarChar, 10
   objRsRpt.Fields.Append "DESC_MONEDA_AJUSTE", adVarChar, 50
   objRsRpt.Fields.Append "GLS_VEJEZ", adVarChar, 100
   objRsRpt.Fields.Append "LETRAMTO", adVarChar, 255
         
   objRsRpt.Open
  

End Sub


