VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_CalPrimaAFP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Solicitud de Traspaso a AFP."
   ClientHeight    =   7440
   ClientLeft      =   600
   ClientTop       =   1140
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10425
   Begin VB.Frame Fra_Dif 
      Caption         =   "  Diferencias  "
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
      Height          =   1065
      Left            =   3720
      TabIndex        =   58
      Top             =   4080
      Width           =   4575
      Begin VB.CommandButton Cmd_Calcular 
         Caption         =   "&Calcular"
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_CalPrimaAFP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Cálcular Pensión"
         Top             =   200
         Width           =   720
      End
      Begin VB.Label Lbl_FactVarRta 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   65
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Vigencia Póliza"
         Height          =   255
         Index           =   12
         Left            =   1920
         TabIndex        =   64
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Factor de Variación Renta"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   63
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label Lbl_FecVigPoliza 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   62
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Cambio"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   61
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Lbl_TipoCambio 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   60
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Fra_AntGral 
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
      Height          =   2925
      Left            =   120
      TabIndex        =   30
      Top             =   1080
      Width           =   8175
      Begin VB.Label Lbl_ReajusteDescripcion 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   74
         Top             =   1875
         Width           =   3615
      End
      Begin VB.Label Lbl_ReajusteValorMen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6600
         TabIndex        =   73
         Top             =   795
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Lbl_Diferidos 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6600
         TabIndex        =   36
         Top             =   1335
         Width           =   1095
      End
      Begin VB.Label Lbl_Meses 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6600
         TabIndex        =   35
         Top             =   1605
         Width           =   1095
      End
      Begin VB.Label Lbl_ReajusteValor 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6600
         TabIndex        =   72
         Top             =   1875
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Reajuste Trim."
         Height          =   255
         Index           =   15
         Left            =   5520
         TabIndex        =   71
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Lbl_ReajusteTipo 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5400
         TabIndex        =   70
         Top             =   795
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   69
         Top             =   1875
         Width           =   1215
      End
      Begin VB.Label Lbl_SumPenInf 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   66
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Lbl_FecAceptacion 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   57
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Aceptación"
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   56
         Top             =   2580
         Width           =   1575
      End
      Begin VB.Label Lbl_FecDevengue 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   55
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Devengue"
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   54
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Lbl_NumIdent 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4560
         TabIndex        =   53
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Lbl_TipoIdent 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   52
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Lbl_Modalidad 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   51
         Top             =   1605
         Width           =   3615
      End
      Begin VB.Label Lbl_TipoRenta 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   50
         Top             =   1335
         Width           =   3615
      End
      Begin VB.Label Lbl_TipoPension 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   49
         Top             =   1065
         Width           =   5895
      End
      Begin VB.Label Lbl_NomAfiliado 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   48
         Top             =   510
         Width           =   5895
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Identificación"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Modalidad"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   46
         Top             =   1605
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Renta"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   45
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Pensión"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   44
         Top             =   1065
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   43
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "CUSPP"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   42
         Top             =   795
         Width           =   1095
      End
      Begin VB.Label Lbl_CUSPP 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   41
         Top             =   795
         Width           =   2535
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Meses Garant."
         Height          =   255
         Index           =   6
         Left            =   5520
         TabIndex        =   40
         Top             =   1605
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Años Diferidos"
         Height          =   255
         Index           =   5
         Left            =   5520
         TabIndex        =   39
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Line Lin_Separar 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   7680
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Prima Cotizada"
         Height          =   195
         Index           =   7
         Left            =   4080
         TabIndex        =   38
         Top             =   2310
         Width           =   1050
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Pensión Cotizada"
         Height          =   195
         Index           =   8
         Left            =   4080
         TabIndex        =   37
         Top             =   2580
         Width           =   1230
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   1
         Left            =   5565
         TabIndex        =   34
         Top             =   2580
         Width           =   375
      End
      Begin VB.Label Lbl_PensionInf 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6000
         TabIndex        =   33
         Top             =   2580
         Width           =   1455
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   3
         Left            =   5565
         TabIndex        =   32
         Top             =   2310
         Width           =   375
      End
      Begin VB.Label Lbl_PrimaInf 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6000
         TabIndex        =   31
         Top             =   2310
         Width           =   1455
      End
   End
   Begin VB.Frame Fra_AntRec 
      Caption         =   "  Antecedentes Recepcionados  "
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
      Height          =   1065
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Width           =   3495
      Begin VB.TextBox Txt_FecTraspaso 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   26
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox Txt_MtoPrimaRec 
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Información AFP"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   29
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Prima Informada"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   2
         Left            =   1515
         TabIndex        =   27
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame Fra_Definitiva 
      Caption         =   "  Definición de Pensión y Primas  "
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
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   8175
      Begin VB.Label Lbl_PrimaDefAFP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5880
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Lbl_PrimaDefCia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3480
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Lbl_SumPenDefAFP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6840
         TabIndex        =   68
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Lbl_SumPenDef 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4680
         TabIndex        =   67
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   6
         Left            =   5280
         TabIndex        =   23
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   22
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Lbl_MtoPenDefUfAFP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5880
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "AFP"
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
         Index           =   24
         Left            =   5880
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Lbl_MtoPenDefUf 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3480
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   15
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Mto Pensión Definitiva"
         Height          =   255
         Index           =   11
         Left            =   720
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Prima Definitiva"
         Height          =   255
         Index           =   23
         Left            =   720
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Compañía"
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
         Index           =   21
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   6240
      Width           =   8175
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   3000
         Picture         =   "Frm_CalPrimaAFP.frx":04A2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar Datos"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4080
         Picture         =   "Frm_CalPrimaAFP.frx":0B5C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5280
         Picture         =   "Frm_CalPrimaAFP.frx":1216
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
   Begin VB.Frame Fra_poliza 
      Caption         =   "  Selección de Póliza  "
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
      TabIndex        =   3
      Top             =   120
      Width           =   8175
      Begin VB.TextBox Txt_Poliza 
         Height          =   285
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   4920
         Picture         =   "Frm_CalPrimaAFP.frx":1310
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Efectuar Busqueda de la Póliza"
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº de Póliza                 :"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Fra_lista 
      Caption         =   "  Pólizas  "
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
      Height          =   7245
      Left            =   8400
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.ListBox Lst_Poliza 
         BackColor       =   &H00E0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6780
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Pólizas                 :"
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
         Height          =   255
         Index           =   22
         Left            =   600
         TabIndex        =   2
         Top             =   2400
         Width           =   735
      End
   End
End
Attribute VB_Name = "Frm_CalPrimaAFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'DECLARACION DE VARIABLES
Dim vlRegistro    As ADODB.Recordset, vlRegistro1 As ADODB.Recordset
Dim vlRegistro2   As ADODB.Recordset, vlRegistro3 As ADODB.Recordset
Dim vlReg         As ADODB.Recordset, vlDif       As Integer
Dim vlPasa        As Boolean, Gls_Pension   As String, Gls_Renta   As String
Dim Gls_Modalidad As String, iFecha         As String, vlFTraspaso As String
Dim vlAnno        As String, vlMes          As String, vlDia       As String
Dim vlCodPar       As String, vlFBRec     As String
Dim vlFBComp      As String, vlFBExon       As String, vlBonRec    As String
Dim vlBonComp     As String, vlBonExon      As String, vlFVigencia As String
Dim vlMtoPenGar   As String, vlMtoPenGarUf  As String, vlFecIniPP  As String
Dim vlFactorDef   As Double, vlFactorDefGar As Double, vlCtaInd    As Double
Dim vlSumPenInf   As Double, vlSumPenDefAFP As Double, vlSumPenDef As Double

Dim vlMtoTotalBono As Double

Const vlEndoso = "1"
Const vlCodTraspaso = "N"
'Const vlCodMoneda = "UF"
Const ciMonedaPensionNew = 0
Const ciMonedaPensionAnt = 1
Const ciMonedaPrimaNew = 2
Const ciMonedaPrimaAnt = 3
Const ciMonedaPrimaDef = 4
Const ciMonedaPrimaDefAFP = 5
Const ciMonedaPensionDefAFP = 6

Dim vlNumDias As Long

Dim vlFecAcepta As String, vlFecDev As String
Dim vlTipoPen As String, vlTipoRen As String, vlMesDif As Long
Dim vlMonedaPension As String
Dim vlTipoCambioAcep As Double ' A la Fecha de Aceptacion
Dim vlTipoCambioRec As Double ' A la Fecha de Recepciòn de la Prima
Dim vlMtoVarFon As Double, vlMtoVarTC As Double, vlMtoVarFonTC As Double 'Para Insertar
Dim vlCalculoPrimerPago As Boolean 'Para verificar si se calculó exitosamente el Monto del Primer Pago
Dim vlFecTraspasoPrimas As String
Dim vl_tipoRenta As String

Function flGrabaPensionActualizada(i As Long) As Boolean
On Error GoTo Errores
flGrabaPensionActualizada = False
If stLiquidacion(i).Fac_Ajuste <> 1 Then
    'Graba Pension Total Actualizada
    'Sql = "INSERT INTO PD_TMAE_PENSIONACT "
    'Sql = Sql & "(NUM_POLIZA, FEC_DESDE, MTO_PENSION) VALUES ('"
    'Sql = Sql & stLiquidacion(i).Num_Poliza & "','" & stLiquidacion(i).Fec_IniPago & "',"
    'Sql = Sql & Str(stLiquidacion(i).Mto_PensionTotal) & ")"
    
    Sql = "INSERT INTO PD_TMAE_PENSIONACT "
    Sql = Sql & "(NUM_POLIZA, FEC_DESDE, MTO_PENSION, MTO_PENSIONGAR) VALUES ('"
    Sql = Sql & stLiquidacion(i).Num_Poliza & "','" & stLiquidacion(i).Fec_IniPago & "',"
    Sql = Sql & Str(stLiquidacion(i).Mto_PensionTotal) & ","
    Sql = Sql & Str(stLiquidacion(i).Mto_pensiongarTotal) & ")"
    
    
    vgConectarBD.Execute (Sql)
End If
Errores:

If Err.Number <> 0 Then
    If Err.Number = -2147217873 Then 'Error de PK, no se toma como error
        flGrabaPensionActualizada = True
    End If
Else
    flGrabaPensionActualizada = True
End If
End Function

Private Function flObtieneNumLiquidacion(iConexion As ADODB.Connection, vlNumLiq As String, oNumRenVit As String) As Boolean
Dim vlSql As String
'genera nuevo numero de Liquidacion
Dim vlNewNumLiq As Long
Dim vlNewNumRenVit  As String
    
    flObtieneNumLiquidacion = False
    vgSql = "SELECT num_liquidacion, cod_renvit "
    vgSql = vgSql & " FROM pd_tmae_gennumliq WHERE num_liquidacion = "
    vgSql = vgSql & " (SELECT MAX(num_liquidacion) FROM pd_tmae_gennumliq)"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        vlNewNumLiq = CInt((vgRs!num_liquidacion)) + 1
        vlNewNumRenVit = Trim(vgRs!cod_renvit)
    Else
        vlNewNumLiq = 1
        vlNewNumRenVit = "002"
    End If
    
    If vlNewNumLiq = 1 Then
        vlSql = "INSERT INTO pd_tmae_gennumliq "
        vlSql = vlSql & "(num_liquidacion,cod_renvit, "
        vlSql = vlSql & " cod_usuariocrea, fec_crea, hor_crea) "
        vlSql = vlSql & " VALUES ('"
        vlSql = vlSql & Trim(Str(vlNewNumLiq)) & "','"
        vlSql = vlSql & Trim(Str(vlNewNumRenVit)) & "','"
        vlSql = vlSql & vgUsuario & "','"
        vlSql = vlSql & Format(Date, "yyyymmdd") & "','"
        vlSql = vlSql & Format(Time, "hhmmss") & "')"
    Else
        vlSql = "UPDATE pd_tmae_gennumliq SET "
        vlSql = vlSql & "num_liquidacion = '" & Trim(Str(vlNewNumLiq)) & "',"
        vlSql = vlSql & "cod_renvit = '" & Trim(Str(vlNewNumRenVit)) & "',"
        vlSql = vlSql & "cod_usuariomodi = '" & vgUsuario & "',"
        vlSql = vlSql & "fec_modi = '" & Format(Date, "yyyymmdd") & "',"
        vlSql = vlSql & "hor_modi = '" & Format(Time, "hhmmss") & "'"
    End If
    iConexion.Execute (vlSql)
    
    vlNumLiq = vlNewNumLiq
    oNumRenVit = vlNewNumRenVit
    
    flObtieneNumLiquidacion = True
End Function


Private Function flValidaPeriodoPagoRegimen(iPeriodo As String, oProxPeriodo As String) As Boolean
    'Valida que el Periodo de Pago en Régimen no se encuentre Cerrado
    
    flValidaPeriodoPagoRegimen = False
    
    vgSql = "SELECT a.cod_estadoreg"
    vgSql = vgSql & " FROM pp_tmae_propagopen a"
    vgSql = vgSql & " WHERE a.num_perpago = '" & iPeriodo & "'"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        If vgRs!cod_estadoreg <> "C" Then
            flValidaPeriodoPagoRegimen = True
        Else
            flValidaPeriodoPagoRegimen = False
        End If
    Else
        flValidaPeriodoPagoRegimen = True
    End If
    
    If Not flValidaPeriodoPagoRegimen Then
        'Obtener el Siguiente Periodo de Pagos en Régimen
        vgSql = "SELECT min(a.num_perpago) as num_perpago FROM pp_tmae_propagopen a"
        vgSql = vgSql & " WHERE a.cod_estadoreg <> 'C'"
        Set vgRs = vgConexionBD.Execute(vgSql)
        If Not vgRs.EOF Then
            If Not IsNull(vgRs!Num_PerPago) Then
                oProxPeriodo = Mid(vgRs!Num_PerPago, 5, 2) & "/" & Mid(vgRs!Num_PerPago, 1, 4)
            Else
                oProxPeriodo = ""
            End If
        End If
    End If
End Function


Private Sub Cmd_Buscar_Click()
On Error GoTo Err_Buscar
    
    Txt_Poliza = UCase(Trim(Txt_Poliza))
    Txt_Poliza = Format(Txt_Poliza, "0000000000")
    If (Txt_Poliza) <> "" Then
        If flBuscaAntecedentes = True Then
            Txt_FecTraspaso.SetFocus
        End If
    Else
        MsgBox "Debe ingresar el Número de la Póliza a Registrar la Prima Recaudada.", vbCritical, "Falta Información"
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

Private Sub Cmd_Calcular_Click()
Dim vlFactVarRta As Double
Dim vlPension As Double
Dim vlExistenPensiones As Boolean
Dim vlProxPeriodo As String

Dim vlMtoPrimaDefAFP As Double, vlMtoPrimaDefCia As Double
Dim vlMtoPensionDefAFP As Double

Dim vlMto_PriUniDif As Double, vlPrc_RentaTmp As Double, vlMto_ValPrePenTmp As Double, vlMto_RentaTmpAfp As Double
Dim vlMto_CtaIndAfp As Double, vlMto_ResMat As Double, vlMto_PenAnual As Double, vlMto_RMPension As Double
Dim vlMto_RMGtoSep As Double, vlMto_RMGtoSepRV As Double, PURD As Double
Dim vlMto_PriUniSim As Double
Dim vlPrc_Tasa_Afp As Double, vlNumMesDif1 As Long, Vpptem As Double
Dim vlFecDevCalculo As String
Dim vlMtoSumPensionDef As Double
Dim vlMtoSumPensionDefAFP As Double
Dim vlCodTipoPension As String
Dim vlCodTipReajuste As String, vlMtoValReajusteTri As Double, vlMtoValReajusteMen As Double 'I--- ABV 05/02/2011 ---

' var RRR 13/01/2012
Dim Mesdif As Long
Dim vppfactor As Double
Dim ival As Double
Dim mescon As Long
Dim FecDev As Long
Dim Fasolp As Integer
Dim Fmsolp As Integer
Dim Fdsolp As Integer
Dim Fechap As Long
'Dim facgratif(1332) As doble
Dim DerGratificacion As String
Dim fecha1 As Date
'Dim vlFecAcepta As Date
Dim mesconTmp As Double
Dim mesdiftmp As Double
Dim vlPrc_RentaTmpSob As Double

On Error GoTo Err_Cal

    'Valida que se hayan cargado los Valores Originales de la Póliza
    If (Lbl_PrimaInf) = "" Or (Lbl_PensionInf) = "" Then
        MsgBox "No se encuentran especificados los valores Originales de la Póliza.", vbCritical, "Error de Datos"
        Cmd_Limpiar.SetFocus
        Exit Sub
    End If
   
'I--- ABV 05/02/2011 ---
    'Valida que se hayan cargado los nuevos valores para el Reajuste
    If (Lbl_ReajusteTipo = "") Or (Lbl_ReajusteValor = "") Or (Lbl_ReajusteValorMen = "") Then
        MsgBox "No se encuentran especificados los valores de Tipo de Reajuste de la Póliza.", vbCritical, "Error de Datos"
        Cmd_Limpiar.SetFocus
        Exit Sub
    End If
'F--- ABV 05/02/2011 ---

    'Valida el Ingreso de la Póliza
    If Fra_poliza.Enabled = False Then
        'Valida el ingreso de la Fecha de Traspaso
        If Txt_FecTraspaso = "" Then
           MsgBox "Debe ingresar la Fecha de Traspaso de la Prima.", vbCritical, "Error de Datos"
           Txt_FecTraspaso.SetFocus
           Exit Sub
        End If
        vlPasa = True
        If (flValidaFecha(Txt_FecTraspaso) = False) Then
           Txt_FecTraspaso = ""
           Txt_FecTraspaso.SetFocus
           Exit Sub
        End If
    Else
        Txt_Poliza.SetFocus
        Exit Sub
    End If
    
    'Valida el ingreso de la Prima Recaudada
    If Txt_MtoPrimaRec = "" Then
       MsgBox "Debe Ingresar un valor para Monto Prima Recibida o Recaudada.", vbCritical, "Error de Datos"
       Txt_MtoPrimaRec.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_MtoPrimaRec.Text) = 0 Then
       MsgBox "Debe Ingresar un Valor para Monto Prima Recibida o Recaudada.", vbCritical, "Error de Datos"
       Txt_MtoPrimaRec.SetFocus
       Exit Sub
    End If
    
'    'Valida que la Fecha de Traspaso no sea menor a la Fecha de Aceptación
'    If (Format(Lbl_FecAceptacion, "yyyymmdd") > Format(Txt_FecTraspaso, "yyyymmdd")) Then
'        MsgBox "La Fecha de Aceptación no puede ser superior a la Fecha de Traspaso de la Prima.", vbCritical, "Error de Datos"
'        Txt_FecTraspaso.SetFocus
'        Exit Sub
'    End If
    
    'Código de Tipo de Pensión
    vlCodTipoPension = fgObtenerCodigo_TextoCompuesto(Lbl_TipoPension)
    
    'Fecha Vigencia Póliza, es el primer dia del mes en que estoy haciendo el Traspaso de la prima
    vlFVigencia = Format(CDate(Trim(Txt_FecTraspaso)), "yyyymmdd")
    vlAnno = Mid(vlFVigencia, 1, 4)
    vlMes = Mid(vlFVigencia, 5, 2)
    vlDia = "01"
    
    Fechap = vlAnno * 12 + vlMes
    
'I--- ABV 05/02/2011 ---
    vlCodTipReajuste = fgObtenerCodigo_TextoCompuesto(Lbl_ReajusteTipo)
    vlMtoValReajusteTri = CDbl(Lbl_ReajusteValor)
    vlMtoValReajusteMen = CDbl(Lbl_ReajusteValorMen)
'F--- ABV 05/02/2011 ---
    
    Lbl_FecVigPoliza = DateSerial(vlAnno, vlMes, vlDia)
    Lbl_MtoPenDefUf = ""
    Lbl_MtoPenDefUfAFP = ""
    Lbl_PrimaDefCia = ""
    Lbl_PrimaDefAFP = ""
    
    Lbl_FactVarRta = ""

'I--- ABV 07/11/2007 ---
    Lbl_SumPenDef = ""
    Lbl_SumPenDefAFP = ""
    vlMtoSumPensionDef = 0
    vlMtoSumPensionDefAFP = 0
'F--- ABV 07/11/2007 ---

    vlFactVarRta = 0
    vlFactorDef = 0
    vlFactorDefGar = 0
    vlMtoPenGarUf = ""
    
    vlMtoPrimaDefAFP = 0
    vlMtoPrimaDefCia = 0
    vlMtoPensionDefAFP = 0

   'Obtiene el Valor de la Moneda a la Fecha de Recepción
    If Not fgObtieneConversion(vlFVigencia, vlMonedaPension, vlTipoCambioRec) Then
        Screen.MousePointer = vbNormal
        MsgBox "No se encuentra registrado el Valor de la Moneda [" & vlMonedaPension & "] para la Fecha de Traspaso de la Prima.", vbCritical, "Advertencia"
        Exit Sub
    End If
    
'    If Not fgObtieneConversion(vlFecAcepta, vlMonedaPension, vlTipoCambioAcep) = True Then
'        Screen.MousePointer = vbNormal
'        MsgBox "No se encuentra registrado el Valor de la Moneda [" & vlMonedaPension & "] para la Fecha de Aceptación de la Póliza.", vbCritical, "Advertencia"
'        Exit Sub
'    End If
    
    Lbl_TipoCambio = vlTipoCambioRec
    
    'Obtiene Variación de la renta
'DAJ    vlFactVarRta = flObtieneFactor(CDbl(Txt_MtoPrimaRec), CDbl(Lbl_PrimaInf), vlTipoCambioRec, vlTipoCambioAcep)
    If (vlMonedaPension = "NS") Then
        vlFactVarRta = Format(CDbl(Txt_MtoPrimaRec) / CDbl(Lbl_PrimaInf), "##0.00000000")
    Else
        vlFactVarRta = Format((CDbl(Txt_MtoPrimaRec) / vlTipoCambioRec) / (CDbl(Lbl_PrimaInf) / vlTipoCambioAcep), "##0.00000000")
    End If
    
    'Lleva a Porcentaje el Factor calculado
'DAJ    Lbl_FactVarRta = Format((vlFactVarRta * 100), "##0.00000000")
    Lbl_FactVarRta = Format((vlFactVarRta), "##0.00000000")
               
    'factor definitivo pension
    vlFactorDef = Format((CDbl(Lbl_PensionInf) * vlFactVarRta), "#,##0.00")
      
    'Aumento o Disminución de la renta o Pensión en UF
'DAJ    vlPension = Format((CDbl(Lbl_PensionInf) + (vlFactorDef)), "#,##0.00")
    vlPension = Format((CDbl(Lbl_PensionInf) * (vlFactVarRta)), "#,##0.00")
    Lbl_MtoPenDefUf = vlPension

'I--- ABV 07/11/2007 ---
    vlMtoSumPensionDef = Format((CDbl(Lbl_SumPenInf) * (vlFactVarRta)), "#,##0.00")
'F--- ABV 07/11/2007 ---
    
'I--- ABV 14/10/2007 ---
'vl_tipoRenta = "5"
'Lbl_Diferidos = 0

    If vl_tipoRenta = "6" Then
            vgSql = ""
            vgSql = "SELECT num_poliza,cod_tippension,"
            vgSql = vgSql & "cod_tipren,num_mesdif,mto_priuni,mto_pension,"
            vgSql = vgSql & "cod_moneda,mto_valmoneda "
            vgSql = vgSql & ",mto_priunidif,prc_rentatmp,mto_valprepentmp,mto_rentatmpafp,mto_ctaindafp "
            vgSql = vgSql & ",mto_resmat,mto_penanual,mto_rmpension,mto_rmgtosep,mto_rmgtoseprv "
            vgSql = vgSql & ",mto_priunisim,prc_tasarprt,num_mesdif,fec_dev,fec_acepta,cod_dergra, case when num_mesdif>0 then round(mto_sumpension/mto_rentatmpafp,8) else 0 end as prc_rentaTMPSob "
            vgSql = vgSql & "FROM "
            vgSql = vgSql & "pd_tmae_oripoliza "
            vgSql = vgSql & "WHERE "
            vgSql = vgSql & "num_poliza = '" & Trim(Txt_Poliza) & "' "
            Set vlRegistro = vgConexionBD.Execute(vgSql)
            If Not (vlRegistro.EOF) Then
                vlMto_ValPrePenTmp = vlRegistro!Mto_ValPrePenTmp
                vlMto_RentaTmpAfp = vlRegistro!Mto_RentaTMPAFP
                vlMto_CtaIndAfp = vlRegistro!Mto_CtaIndAFP
            End If
            
            If vl_tipoRenta = "5" Or vl_tipoRenta = "3" Then
              vlMtoPensionDefAFP = Format(CDbl(vlMto_RentaTmpAfp) * vlFactVarRta, "#0.00")
              vlMtoPrimaDefAFP = CDbl(Txt_MtoPrimaRec) * (Valor_ParRC("RCPRC") / 100)
              vlMtoPrimaDefCia = CDbl(Txt_MtoPrimaRec) - vlMtoPrimaDefAFP
            Else
                vlMtoPensionDefAFP = 0
                vlMtoPrimaDefAFP = 0
                vlMtoPrimaDefCia = CDbl(Txt_MtoPrimaRec)
        'I--- ABV 07/11/2007 ---
                vlMtoSumPensionDefAFP = 0
        'F--- ABV 07/11/2007 ---
            End If
    Else
        If (CLng(Lbl_Diferidos) <> 0) Then
        
            vlMto_PriUniSim = 0
            vlMto_PriUniDif = 0
            vlPrc_RentaTmp = 0
            vlMto_ValPrePenTmp = 0
            vlMto_RentaTmpAfp = 0
            vlMto_CtaIndAfp = 0
            vlMto_ResMat = 0
            vlMto_PenAnual = 0
            vlMto_RMPension = 0
            vlMto_RMGtoSep = 0
            vlMto_RMGtoSepRV = 0
            
            vgSql = "SELECT num_poliza,cod_tippension,"
            vgSql = vgSql & "cod_tipren,num_mesdif,mto_priuni,mto_pension,"
            vgSql = vgSql & "cod_moneda,mto_valmoneda "
            vgSql = vgSql & ",mto_priunidif,prc_rentatmp,mto_valprepentmp,mto_rentatmpafp,mto_ctaindafp "
            vgSql = vgSql & ",mto_resmat,mto_penanual,mto_rmpension,mto_rmgtosep,mto_rmgtoseprv "
            vgSql = vgSql & ",mto_priunisim,prc_tasarprt,num_mesdif,fec_dev,fec_acepta,cod_dergra, case when num_mesdif>0 then round(mto_sumpension/mto_rentatmpafp,8) else 0 end as prc_rentaTMPSob "
            vgSql = vgSql & "FROM "
            vgSql = vgSql & "pd_tmae_oripoliza "
            vgSql = vgSql & "WHERE "
            vgSql = vgSql & "num_poliza = '" & Trim(Txt_Poliza) & "' "
            Set vlRegistro = vgConexionBD.Execute(vgSql)
            If Not (vlRegistro.EOF) Then
    
                vlMto_PriUniSim = vlRegistro!Mto_PriUniSim
                vlMto_PriUniDif = vlRegistro!Mto_PriUniDif
                
                If vlRegistro!Cod_TipPension = "08" Then
                    vlPrc_RentaTmp = vlRegistro!prc_rentaTMPSob * 100
                Else
                    vlPrc_RentaTmp = vlRegistro!Prc_RentaTMP
                End If
                vlMto_ValPrePenTmp = vlRegistro!Mto_ValPrePenTmp
                vlMto_RentaTmpAfp = vlRegistro!Mto_RentaTMPAFP
                vlMto_CtaIndAfp = vlRegistro!Mto_CtaIndAFP
                vlMto_ResMat = vlRegistro!Mto_ResMat
                vlMto_PenAnual = vlRegistro!Mto_PenAnual
                vlMto_RMPension = vlRegistro!Mto_RMPension
                vlMto_RMGtoSep = vlRegistro!Mto_RMGtoSep
                vlMto_RMGtoSepRV = vlRegistro!Mto_RMGtoSepRV
                vlNumMesDif1 = vlRegistro!Num_MesDif
                vlFecAcepta = vlRegistro!Fec_Acepta
                DerGratificacion = vlRegistro!Cod_DerGra
                vlPrc_RentaTmpSob = vlRegistro!prc_rentaTMPSob
                
                'I - MC 26/05/2008
                vlFecDevCalculo = DateAdd("m", vlNumMesDif1, DateSerial(Mid(vlRegistro!Fec_Dev, 1, 4), Mid(vlRegistro!Fec_Dev, 5, 2), Mid(vlRegistro!Fec_Dev, 7, 2)))
                If (Format(vlFecDevCalculo, "yyyymmdd") < Format(Date, "yyyymmdd")) Then
                    'Cuando la fecha de devengue + los meses diferidos es menor a la fecha de cálculo utiliza este valor definido por PAC
                    vlPrc_Tasa_Afp = 0.0000000001
                Else
                    vlPrc_Tasa_Afp = vlRegistro!Prc_TasaRPRT / 100
                End If
                'F - MC 26/05/2008
                
                ''  RRR 13/01/2012
                FecDev = vlRegistro!Fec_Dev
                Fasolp = CLng(Mid(vlRegistro!Fec_Dev, 1, 4))
                Fmsolp = CLng(Mid(vlRegistro!Fec_Dev, 5, 2))
                ''  RRR 13/01/2012
            End If
            vlRegistro.Close
          
            'If (vl_tipoRenta = "5" Or vl_tipoRenta = "3") Then
              
            'Else
                ''  RRR 13/01/2012
                mescon = Fechap - ((Fasolp * 12) + Fmsolp)
                Mesdif = vlNumMesDif1
                ''  RRR 13/01/2012
        
            
                'Valor Presente de la Temporal
                Vpptem = 0
                Vpptem = ((1 - 1 / ((1 + vlPrc_Tasa_Afp) ^ (vlNumMesDif1 / 12))) / vlPrc_Tasa_Afp) * (1 + vlPrc_Tasa_Afp) * 12
                Vpptem = Format(Vpptem, "#,#0.00000000")
                
                '' cambio RRR 12/01/2012
        
        
                               ' If Mesdif < mescon Then
                               '     mescon = Mesdif
                               ' Else
                               '     mescon = mescon
                               ' End If
                            
                               'RRR 21/05/2012 cambio que corresponde a la distribucion correcta de la prima para los casos con grati....washin!!
                                If DerGratificacion = "S" Then
                                     Mesdif = Mesdif + ((Mesdif / 12) * 2)
                                End If
                               'RRR
        
                                mesdiftmp = Mesdif - mescon
        
                                If mesdiftmp < 0 Then mesdiftmp = 0
        
                                ival = (((1 + vlPrc_Tasa_Afp) ^ (1 / 12)) - 1)
        
                                mesconTmp = CDbl(IIf(mescon > Mesdif, Mesdif, mescon))
        
                                vppfactor = ival / ((ival * mesconTmp) + (1 - (1 + ival) ^ -(mesdiftmp)) * (1 + ival))
        
                                Vpptem = 1 / vppfactor
                                'Sald_sim = CDbl(Format(CDbl(rete_sim / vppfactor), "#,#0.00"))
                '' cambio RRR
                
                
                'vlMtoPensionDefAFP = Format(CDbl(Txt_MtoPrimaRec) * (1 / (vlMto_ValPrePenTmp + vlPrc_RentaTmp * vlMto_PriUniSim)), "#0.00")
                vlMtoPensionDefAFP = Format(vlPension * (1 / (vlPrc_RentaTmp / 100)) * vlTipoCambioRec, "#0.00")
                vlMtoPrimaDefAFP = Format(vlMtoPensionDefAFP / vppfactor, "#0.000") 'vlMto_ValPrePenTmp , "#0.00")
                'vlMtoPrimaDefAFP = Format(vlMtoPensionDefAFP * Vpptem, "#0.000") 'vlMto_ValPrePenTmp , "#0.00")
                
        'I--- DA 07/11/2007 ---
                vlMtoSumPensionDefAFP = Format(vlMtoSumPensionDef * (1 / (vlPrc_RentaTmp / 100)) * vlTipoCambioRec, "#0.00")
                If (vlCodTipoPension = clCodTipPensionSob) Then
                    vlMtoPrimaDefAFP = Format(vlMtoSumPensionDefAFP * Vpptem, "#0.000") 'vlMto_ValPrePenTmp , "#0.00")
                End If
        'F--- DA 07/11/2007 ---
                
                vlMtoPrimaDefCia = CDbl(Txt_MtoPrimaRec) - vlMtoPrimaDefAFP
                'vlMtoPensionDefAFP = Format(((vlMtoPrimaDefCia - vlMto_RMGtoSepRV) / vlMto_PriUniSim) / vlTipoCambioRec, "#0.00")
            'End If
            
    
        Else
            vgSql = ""
            vgSql = "SELECT num_poliza,cod_tippension,"
            vgSql = vgSql & "cod_tipren,num_mesdif,mto_priuni,mto_pension,"
            vgSql = vgSql & "cod_moneda,mto_valmoneda "
            vgSql = vgSql & ",mto_priunidif,prc_rentatmp,mto_valprepentmp,mto_rentatmpafp,mto_ctaindafp "
            vgSql = vgSql & ",mto_resmat,mto_penanual,mto_rmpension,mto_rmgtosep,mto_rmgtoseprv "
            vgSql = vgSql & ",mto_priunisim,prc_tasarprt,num_mesdif,fec_dev,fec_acepta,cod_dergra, case when num_mesdif>0 then round(mto_sumpension/mto_rentatmpafp,8) else 0 end as prc_rentaTMPSob "
            vgSql = vgSql & "FROM "
            vgSql = vgSql & "pd_tmae_oripoliza "
            vgSql = vgSql & "WHERE "
            vgSql = vgSql & "num_poliza = '" & Trim(Txt_Poliza) & "' "
            Set vlRegistro = vgConexionBD.Execute(vgSql)
            If Not (vlRegistro.EOF) Then
                vlMto_ValPrePenTmp = vlRegistro!Mto_ValPrePenTmp
                vlMto_RentaTmpAfp = vlRegistro!Mto_RentaTMPAFP
                vlMto_CtaIndAfp = vlRegistro!Mto_CtaIndAFP
            End If
            
            If vl_tipoRenta = "5" Or vl_tipoRenta = "3" Then
              vlMtoPensionDefAFP = Format(CDbl(vlMto_RentaTmpAfp) * vlFactVarRta, "#0.00")
              vlMtoPrimaDefAFP = CDbl(Txt_MtoPrimaRec) * (Valor_ParRC("RCPRC") / 100)
              vlMtoPrimaDefCia = CDbl(Txt_MtoPrimaRec) - vlMtoPrimaDefAFP
            Else
                vlMtoPensionDefAFP = 0
                vlMtoPrimaDefAFP = 0
                vlMtoPrimaDefCia = CDbl(Txt_MtoPrimaRec)
        'I--- ABV 07/11/2007 ---
                vlMtoSumPensionDefAFP = 0
        'F--- ABV 07/11/2007 ---
            End If
        End If
    End If
    
    
    Lbl_PrimaDefCia = Format(vlMtoPrimaDefCia, "#,#0.00")
    Lbl_PrimaDefAFP = Format(vlMtoPrimaDefAFP, "#,#0.00")
    Lbl_MtoPenDefUfAFP = Format(vlMtoPensionDefAFP, "#,#0.00")
'F--- ABV 14/10/2007 ---

'I--- ABV 07/11/2007 ---
    Lbl_SumPenDef = Format(vlMtoSumPensionDef, "#,#0.00")
    Lbl_SumPenDefAFP = Format(vlMtoSumPensionDefAFP, "#,#0.00")
'F--- ABV 07/11/2007 ---

    If CDbl(vlMtoPenGar) <> 0 Then
'DAJ         vlFactorDefGar = Format((CDbl(vlMtoPenGar) * vlFactVarRta), "#,##0.00")
         vlFactorDefGar = vlFactVarRta
        'Aumento o Disminución de la renta garantizada
'DAJ         vlMtoPenGarUf = Format((CDbl(vlMtoPenGar) + (vlFactorDefGar)), "#,##0.00")
         vlMtoPenGarUf = Format((CDbl(vlMtoPenGar) * (vlFactorDefGar)), "#,##0.00")
'I--- ABV 21/08/2007 ---
    Else
        vlFactorDefGar = 0
        vlMtoPenGarUf = 0
'F--- ABV 21/08/2007 ---
    End If
    
    Cmd_Grabar.SetFocus
    
    Screen.MousePointer = 0
    
Exit Sub
Err_Cal:
    Screen.MousePointer = 0
    
    Lbl_FecVigPoliza = ""
    Lbl_FactVarRta = ""
    Lbl_TipoCambio = ""
    Lbl_MtoPenDefUf = ""
    Lbl_PrimaDefCia = ""
    Lbl_MtoPenDefUfAFP = ""
    Lbl_PrimaDefAFP = ""
    
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

 Private Function Valor_ParRC(ByVal valor As String) As Double

        'Dim ds As New DataSet
        Dim douMonto As Double

        vgSql = "SELECT mto_elemento FROM MA_TPAR_TABCOD WHERE COD_TABLA='VRC' AND COD_ELEMENTO='" & valor & "'"
        Set vlRegistro = vgConexionBD.Execute(vgSql)
        If Not (vlRegistro.EOF) Then
            Valor_ParRC = CDbl(vlRegistro!mto_elemento)
        Else
            Valor_ParRC = 0
        End If
        
    End Function

Private Sub Cmd_Grabar_Click()
Dim vlOp As String
Dim vlResp As Long
On Error GoTo Err_Grabar
    
    'Valida el ingreso del Número de Póliza
    Txt_Poliza = Format(Trim(Txt_Poliza), "0000000000")
    If (Txt_Poliza = "") Then
        MsgBox "Debe seleccionar la Póliza a registrar la Prima.", vbCritical, "Error de Datos"
        If Fra_poliza.Enabled = True Then Txt_Poliza.SetFocus
        If Fra_poliza.Enabled = False Then Cmd_Salir.SetFocus
        Exit Sub
    End If
    
    'Valida el Ingreso de la Póliza
    If Fra_poliza.Enabled = False Then
        'Valida el ingreso de la Fecha de Traspaso
        If Txt_FecTraspaso = "" Then
           MsgBox "Debe ingresar la Fecha de Traspaso de la Prima.", vbCritical, "Error de Datos"
           Txt_FecTraspaso.SetFocus
           Exit Sub
        End If
        vlPasa = True
        If (flValidaFecha(Txt_FecTraspaso) = False) Then
           Txt_FecTraspaso = ""
           Txt_FecTraspaso.SetFocus
           Exit Sub
        End If
    Else
        Txt_Poliza.SetFocus
        Exit Sub
    End If
    
    'Valida el ingreso de la Prima Recaudada
    If Txt_MtoPrimaRec = "" Then
       MsgBox "Debe ingresar el Monto de la Prima Recibida o Recaudada.", vbCritical, "Error de Datos"
       Txt_MtoPrimaRec.SetFocus
       Exit Sub
    End If
    If CDbl(Txt_MtoPrimaRec) = 0 Then
       MsgBox "Debe Ingresar un Valor Mayor que 0 para Monto Prima Recibida o Recaudada.", vbCritical, "Error de Datos"
       Txt_MtoPrimaRec.SetFocus
       Exit Sub
    End If
    
    'Validar que se haya realizado el Proceso de Cálculo de Diferencias
    'Valida el Cálculo de la Fecha de Vigencia de la Póliza
    If Lbl_FecVigPoliza = "" Then
       MsgBox "Debe calcular la Fecha de Inicio de Vigencia de la Póliza.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    'Valida el Cálculo del Factor
    If Lbl_FactVarRta = "" Then
       MsgBox "Debe calcular el Factor de Variación de la Prima Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    'Valida el Cálculo del Monto de la Pensión Nueva en UF
    If Lbl_MtoPenDefUf = "" Then
       MsgBox "Debe calcular la Pensión a recibir, de acuerdo a la Prima Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    'Valida el cálculo de la Pensión de la AFP
    If Lbl_MtoPenDefUfAFP = "" Then
       MsgBox "Debe calcular la Pensión a recibir por la AFP, de acuerdo a la Prima Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    'Valida el cálculo de la Prima de la AFP
    If Lbl_PrimaDefAFP = "" Then
       MsgBox "Debe calcular la Prima a recibir por la AFP, de acuerdo a la Prima Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    'Valida el cálculo de la Prima de la Cía
    If Lbl_PrimaDefCia = "" Then
       MsgBox "Debe calcular la Prima a recibir por la Cía, de acuerdo a la Prima Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
'I--- 07/11/2007 ---
    'Valida el cálculo de la Suma Pensión (<> 0 para Casos de Sobrevivencia)
    If Lbl_SumPenInf = "" Then
       MsgBox "Debe calcular la Pensión a recibir por la Cía, de acuerdo a la Prima Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    'Valida el cálculo de la Suma Pensión Definitiva (<> 0 para Casos de Sobrevivencia)
    If Lbl_SumPenDef = "" Then
       MsgBox "Debe calcular la Pensión Definitiva a recibir por la Cía, de acuerdo a la Prima Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    'Valida el cálculo de la Suma Pensión Definitiva AFP (<> 0 para Casos de Sobrevivencia y con Periodo Diferido)
    If Lbl_SumPenDefAFP = "" Then
       MsgBox "Debe calcular la Pensión Definitiva a recibir por la Cía, de acuerdo a la Prima Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
'F--- 07/11/2007 ---
    
    Screen.MousePointer = 11
    
    Txt_Poliza = Trim(Txt_Poliza)
    vlFTraspaso = Format(CDate(Trim(Txt_FecTraspaso)), "yyyymmdd")
    vlFVigencia = Format(CDate(Trim(Lbl_FecVigPoliza)), "yyyymmdd")
    vlOp = ""
    
   'Verifica la existencia de la póliza
    vgSql = ""
    vgSql = "SELECT num_poliza FROM pd_tmae_polprirecaux WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_Poliza) & "' "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If vlRegistro.EOF Then
        vlOp = "I"
    End If
    vlRegistro.Close
    
    If (vlOp = "I") Then
        vlResp = MsgBox(" ¿ Está seguro que desea ingresar los Datos ?", 4 + 32 + 256, "Proceso de Ingreso de Datos")
        If vlResp <> 6 Then
            Cmd_Salir.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        vgPalabra = "N"
         
        Call flGrabaRecepcion
        
        If vgPalabra = "S" Then
            Cmd_Limpiar_Click
            Fra_lista.Enabled = True
            Call flCargaPolizas
            MsgBox " Los Datos se han Actualizado Satisfactoriamente", vbInformation, "Proceso de Actualización"
        End If
    End If
    Screen.MousePointer = 0
    
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

    Lbl_NumIdent = ""
    Lbl_TipoIdent = ""
    Lbl_CUSPP = ""
    Lbl_NomAfiliado = ""
    Lbl_TipoPension = ""
    Lbl_TipoRenta = ""
    Lbl_Modalidad = ""
    Lbl_Diferidos = ""
    Lbl_Meses = ""
'I--- ABV 05/02/2011 ---
    Lbl_ReajusteTipo = ""
    Lbl_ReajusteValor = ""
    Lbl_ReajusteValorMen = ""
    Lbl_ReajusteDescripcion = ""
'F--- ABV 05/02/2011 ---
    Lbl_FecDevengue = ""
    Lbl_FecAceptacion = ""
    Lbl_PrimaInf = ""
    Lbl_PensionInf = ""
    Txt_FecTraspaso = ""
    Txt_MtoPrimaRec = ""
    Lbl_FecVigPoliza = ""
    Lbl_FactVarRta = ""
    Lbl_MtoPenDefUf = ""
    Lbl_MtoPenDefUfAFP = ""
    Lbl_PrimaDefCia = ""
    Lbl_PrimaDefAFP = ""
    
'I--- ABV 07/11/2007 ---
    Lbl_SumPenInf = ""
    Lbl_SumPenDef = ""
    Lbl_SumPenDefAFP = ""
    Lbl_SumPenInf.Visible = False
    Lbl_SumPenDef.Visible = False
    Lbl_SumPenDefAFP.Visible = False
'F--- ABV 07/11/2007 ---
    
    Txt_Poliza = ""
    
    Fra_AntRec.Enabled = False
    Fra_Dif.Enabled = False
    'SSTab1.Enabled = False
    Fra_poliza.Enabled = True
    Fra_lista.Enabled = True
    'SSTab1.Tab = 0
    
    Txt_Poliza.SetFocus
    
Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
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

    Me.Top = 0
    Me.Left = 0
    
    'Actualizar los Datos en la Lista
    flCargaPolizas
    Lbl_Moneda(ciMonedaPrimaAnt) = vgMonedaCodOfi  'cgCodTipMonedaUF
    Lbl_Moneda(ciMonedaPrimaDef) = vgMonedaCodOfi
    Lbl_Moneda(ciMonedaPrimaNew) = vgMonedaCodOfi  'cgCodTipMonedaUF
    Lbl_Moneda(ciMonedaPrimaDefAFP) = vgMonedaCodOfi
    Lbl_Moneda(ciMonedaPensionDefAFP) = vgMonedaCodOfi
    
    Fra_AntRec.Enabled = False
    Fra_Dif.Enabled = False
  
'I--- ABV 07/11/2007 ---
    Lbl_SumPenInf = ""
    Lbl_SumPenDef = ""
    Lbl_SumPenDefAFP = ""
    Lbl_SumPenInf.Visible = False
    Lbl_SumPenDef.Visible = False
    Lbl_SumPenDefAFP.Visible = False
'F--- ABV 07/11/2007 ---
    

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


Private Sub Lst_Poliza_Click()
On Error GoTo Err_Click

    Screen.MousePointer = 11
    
    'valida que exista informacion en la lista
    If (Lst_Poliza.ListCount > 0) And (Lst_Poliza.Text <> "") Then
        Txt_Poliza = Lst_Poliza.Text
        Cmd_Buscar.SetFocus
    End If
          
    Screen.MousePointer = 0

Exit Sub
Err_Click:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_FecTraspaso_Change()
    Lbl_FecVigPoliza = ""
    Lbl_FactVarRta = ""
    Lbl_TipoCambio = ""
    Lbl_MtoPenDefUf = ""
    Lbl_MtoPenDefUfAFP = ""
    Lbl_PrimaDefCia = ""
    Lbl_PrimaDefAFP = ""
    Lbl_SumPenDef = ""
    Lbl_SumPenDefAFP = ""
End Sub

Private Sub Txt_FecTraspaso_GotFocus()

    Txt_FecTraspaso.SelStart = 0
    Txt_FecTraspaso.SelLength = Len(Txt_FecTraspaso)
    
End Sub

Private Sub Txt_FecTraspaso_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (Trim(Txt_FecTraspaso.Text) <> "") Then
       If Not (Txt_FecTraspaso = "0") Then
          vlPasa = True
          If flValidaFecha(Txt_FecTraspaso) = True Then
             Txt_MtoPrimaRec.SetFocus
          End If
       End If
    End If
    
End Sub

Private Sub Txt_FecTraspaso_LostFocus()

    If (Trim(Txt_FecTraspaso) = "") Then
       Txt_FecTraspaso = ""
       Exit Sub
    End If
    If Not IsDate(Txt_FecTraspaso.Text) Then
       Txt_FecTraspaso = ""
       Exit Sub
    End If
    If (CDate(Txt_FecTraspaso) > CDate(Date)) Then
       Txt_FecTraspaso = ""
       Exit Sub
    End If
    If (Year(Txt_FecTraspaso) < 1900) Then
       Txt_FecTraspaso = ""
       Exit Sub
    End If
    Txt_FecTraspaso.Text = Format(CDate(Trim(Txt_FecTraspaso)), "yyyymmdd")
    Txt_FecTraspaso.Text = DateSerial(Mid((Txt_FecTraspaso.Text), 1, 4), Mid((Txt_FecTraspaso.Text), 5, 2), Mid((Txt_FecTraspaso.Text), 7, 2))
    vlFecTraspasoPrimas = Format(CDate(Trim(Txt_FecTraspaso)), "yyyymmdd")
End Sub

Private Sub Txt_MtoPrimaRec_Change()

    'If Not IsNumeric(Txt_MtoPrimaRec) Then
    '    Txt_MtoPrimaRec = ""
        Lbl_FecVigPoliza = ""
        Lbl_FactVarRta = ""
        Lbl_TipoCambio = ""
        Lbl_MtoPenDefUf = ""
        Lbl_MtoPenDefUfAFP = ""
        Lbl_PrimaDefCia = ""
        Lbl_PrimaDefAFP = ""
    'End If

End Sub

Private Sub Txt_MtoPrimaRec_GotFocus()

    Txt_MtoPrimaRec.SelStart = 0
    Txt_MtoPrimaRec.SelLength = Len(Txt_MtoPrimaRec)
    
End Sub

Private Sub Txt_MtoPrimaRec_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Txt_MtoPrimaRec.Text) Then
           MsgBox "Debe Ingresar un Valor para Monto Prima Recibida.", vbCritical, "Error de Datos"
           Txt_MtoPrimaRec.SetFocus
           Exit Sub
        End If
        If CDbl(Txt_MtoPrimaRec.Text) = 0 Then
           MsgBox "Debe Ingresar un Valor Mayor que Cero para Monto Prima Recibida o Recaudada.", vbCritical, "Error de Datos"
           Txt_MtoPrimaRec.SetFocus
           Exit Sub
        End If
        If IsNumeric(Txt_MtoPrimaRec) Then
            Txt_MtoPrimaRec = Format(Txt_MtoPrimaRec, "#,##0.00")
            Cmd_Calcular.SetFocus
        End If
    End If
    
End Sub

Private Sub Txt_MtoPrimaRec_LostFocus()

    If (Txt_MtoPrimaRec.Text) = "" Then
       Exit Sub
    End If
    If CDbl(Txt_MtoPrimaRec.Text) = 0 Then
       Exit Sub
    End If
    If IsNumeric(Txt_MtoPrimaRec) Then
        Txt_MtoPrimaRec = Format(Txt_MtoPrimaRec, "#,##0.00")
    End If
    
End Sub

Private Sub Txt_Poliza_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Poliza
   
   If KeyAscii = 13 Then
      If Trim(Txt_Poliza) > "" Then
         Txt_Poliza = Trim(UCase(Txt_Poliza))
         Txt_Poliza = Format(Txt_Poliza, "0000000000")
         Cmd_Buscar.SetFocus
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

Function flCargaPolizas()
On Error GoTo Err_Act

    Lst_Poliza.Clear
    
    vgSql = ""
    vgSql = "SELECT DISTINCT num_poliza FROM pd_tmae_oripoliza "
    vgSql = vgSql & "where num_poliza not in (select num_poliza from pd_tmae_polprirecaux) "
    vgSql = vgSql & "ORDER BY num_poliza "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    While Not vlRegistro.EOF
        Lst_Poliza.AddItem Trim(vlRegistro!Num_Poliza)
        vlRegistro.MoveNext
    Wend
    vlRegistro.Close
     
Exit Function
Err_Act:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flBuscaAntecedentes()
'Dim vlCodPa As String
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
    vgSql = vgSql & "p.mto_bono,p.mto_priuni,p.mto_pension,p.mto_pensiongar,"
    vgSql = vgSql & "t.gls_elemento as gls_pension,"
    vgSql = vgSql & "r.gls_elemento as gls_renta,"
    vgSql = vgSql & "m.gls_elemento as gls_modalidad,"
    vgSql = vgSql & "be.gls_nomben,be.gls_patben,be.gls_matben, "
    vgSql = vgSql & "p.cod_moneda, p.fec_acepta, p.fec_dev, p.mto_valmoneda "
    vgSql = vgSql & ",p.mto_priunidif,p.prc_rentatmp "
'I--- ABV 07/11/2007 ---
    vgSql = vgSql & ",p.mto_sumpension "
'F--- ABV 07/11/2007 ---
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & ",p.cod_tipreajuste,p.mto_valreajustetri,p.mto_valreajustemen,"
    vgSql = vgSql & "tr.gls_elemento as gls_tipreajuste "
    vgSql = vgSql & ",mtr.cod_scomp as cod_montipreaju,mtr.gls_descripcion as gls_montipreaju "
'F--- ABV 05/02/2011 ---
    vgSql = vgSql & "FROM "
    vgSql = vgSql & "pd_tmae_oripoliza p, ma_tpar_tabcod t, ma_tpar_tabcod r, "
    vgSql = vgSql & "ma_tpar_tabcod m, pd_tmae_oripolben be, ma_tpar_tipoiden a "
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & ",ma_tpar_tabcod tr, ma_tpar_monedatiporeaju mtr "
'F--- ABV 05/02/2011 ---
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "p.num_poliza = '" & Trim(Txt_Poliza) & "' AND "
    vgSql = vgSql & "p.num_poliza = be.num_poliza AND "
    vgSql = vgSql & "be.cod_par = '" & cgCodParentescoCau & "' AND "
    vgSql = vgSql & "t.cod_tabla = '" & vgCodTabla_TipPen & "' AND "
    vgSql = vgSql & "t.cod_elemento = p.cod_tippension AND "
    vgSql = vgSql & "r.cod_tabla = '" & vgCodTabla_TipRen & "' AND "
    vgSql = vgSql & "r.cod_elemento = p.cod_tipren AND "
    vgSql = vgSql & "m.cod_tabla = '" & vgCodTabla_AltPen & "' AND "
    vgSql = vgSql & "m.cod_elemento = p.cod_modalidad AND "
    vgSql = vgSql & "p.cod_tipoidenafi = a.cod_tipoiden "
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & "AND p.cod_tipreajuste = tr.cod_elemento(+) AND "
    vgSql = vgSql & "tr.cod_tabla = '" & vgCodTabla_TipReajuste & "' "
    vgSql = vgSql & "AND p.cod_tipreajuste = mtr.cod_tipreajuste(+) AND "
    vgSql = vgSql & "p.cod_moneda = mtr.cod_moneda(+) "
'F--- ABV 05/02/2011 ---
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro.EOF) Then
        'SSTab1.Enabled = True
        Fra_poliza.Enabled = False
        Fra_AntRec.Enabled = True
        Fra_Dif.Enabled = True
        Fra_lista.Enabled = False
            
        Lbl_TipoPension = Trim(vlRegistro!Cod_TipPension) + " - " + Trim(vlRegistro!Gls_Pension)
        Lbl_TipoRenta = Trim(vlRegistro!Cod_TipRen) + " - " + Trim(vlRegistro!Gls_Renta)
        vl_tipoRenta = Trim(vlRegistro!Cod_TipRen)
        Lbl_Modalidad = Trim(vlRegistro!Cod_Modalidad) + " - " + Trim(vlRegistro!Gls_Modalidad)
        Lbl_NumIdent = Trim(vlRegistro!num_idenafi)
        Lbl_TipoIdent = Trim(vlRegistro!gls_Tipoidencor)
        Lbl_NomAfiliado = Trim(vlRegistro!Gls_NomBen) + " " + Trim(vlRegistro!Gls_PatBen) + " " + IIf(IsNull(vlRegistro!Gls_MatBen), "", Trim(vlRegistro!Gls_MatBen))
        Lbl_CUSPP = Trim(vlRegistro!Cod_Cuspp)
        
        vlDif = (vlRegistro!Num_MesDif)
        Lbl_Diferidos = ((vlDif) / 12)
        Lbl_Meses = (vlRegistro!Num_MesGar)
        
'I--- ABV 05/02/2011 ---
        Lbl_ReajusteTipo = Trim(vlRegistro!Cod_TipReajuste) + " - " + Trim(vlRegistro!Gls_TipReajuste)
        Lbl_ReajusteValor = Format(vlRegistro!Mto_ValReajusteTri, "#0.00000000")
        Lbl_ReajusteValorMen = Format(vlRegistro!Mto_ValReajusteMen, "#0.00000000")
        Lbl_ReajusteDescripcion = Trim(vlRegistro!cod_montipreaju) + " - " + Trim(vlRegistro!gls_montipreaju)
'F--- ABV 05/02/2011 ---

        Lbl_FecDevengue = DateSerial(Mid(vlRegistro!Fec_Dev, 1, 4), Mid(vlRegistro!Fec_Dev, 5, 2), Mid(vlRegistro!Fec_Dev, 7, 2))
        Lbl_FecAceptacion = DateSerial(Mid(vlRegistro!Fec_Acepta, 1, 4), Mid(vlRegistro!Fec_Acepta, 5, 2), Mid(vlRegistro!Fec_Acepta, 7, 2))
'I--- ABV 21/08/2007 --- Debe ser sobre la Prima a Recibir por la Cía.
        Lbl_PrimaInf = Format((vlRegistro!mto_priuni), "#,#0.00")
'        If (vlDif > 0) Then
'            If (vlRegistro!Cod_Moneda = vgMonedaCodOfi) Then
'                Lbl_PrimaInf = Format((vlRegistro!Mto_PriUniDif), "#,#0.00")
'            Else
'                If (vlRegistro!Mto_ValMoneda > 0) Then
'                    Lbl_PrimaInf = Format((vlRegistro!Mto_PriUniDif * vlRegistro!Mto_ValMoneda), "#,#0.00")
'                Else
'                    Lbl_PrimaInf = Format((vlRegistro!Mto_PriUniDif), "#,#0.00")
'                End If
'            End If
'        Else
'            Lbl_PrimaInf = Format((vlRegistro!mto_priuni), "#,#0.00")
'        End If
'F--- ABV 21/08/2007 ---
        Lbl_PensionInf = Format((vlRegistro!Mto_Pension), "#,#0.00")
        vlMtoPenGar = Format((vlRegistro!Mto_SumPension), "#,#0.00")
        
'I--- ABV 07/11/2007 ---
        Lbl_SumPenInf = Format((vlRegistro!Mto_SumPension), "#,#0.00")
        If (vlRegistro!Cod_TipPension = clCodTipPensionSob) Then
            Lbl_SumPenInf.Visible = True
            Lbl_SumPenDef.Visible = True
            Lbl_SumPenDefAFP.Visible = True

            Lbl_PensionInf.Enabled = False
            Lbl_MtoPenDefUf.Enabled = False
            Lbl_MtoPenDefUfAFP.Enabled = False

            Lbl_SumPenInf.Top = 2580 '2340
            Lbl_SumPenInf.Left = 6000
            Lbl_SumPenDef.Top = 480
            Lbl_SumPenDef.Left = 3480
            Lbl_SumPenDefAFP.Top = 480
            Lbl_SumPenDefAFP.Left = 5880
        Else
            Lbl_SumPenInf.Visible = False
            Lbl_SumPenDef.Visible = False
            Lbl_SumPenDefAFP.Visible = False

            Lbl_PensionInf.Enabled = True
            Lbl_MtoPenDefUf.Enabled = True
            Lbl_MtoPenDefUfAFP.Enabled = True
        End If
'F--- ABV 07/11/2007 ---
        
        vlMonedaPension = Trim(vlRegistro!Cod_Moneda)
        Lbl_Moneda(ciMonedaPensionAnt) = vlMonedaPension
        Lbl_Moneda(ciMonedaPensionNew) = Lbl_Moneda(ciMonedaPensionAnt)
        
        vlFecAcepta = vlRegistro!Fec_Acepta
        vlFecDev = vlRegistro!Fec_Dev
        vlTipoPen = vlRegistro!Cod_TipPension
        vlTipoRen = vlRegistro!Cod_TipRen
        vlMesDif = vlRegistro!Num_MesDif
''        If vlMesDif > 0 Then 'Diferida
''            Txt_FecPriPag.Enabled = False
''        Else
''            Txt_FecPriPag.Enabled = True
''        End If
        vlNumDias = fgObtieneDiasPrimerPagoEst(vlTipoPen)
        vlTipoCambioAcep = vlRegistro!Mto_ValMoneda
'        Lbl_PorcRentaTmp = Format(vlRegistro!Prc_RentaTMP, "#,#0.00")
'        Lbl_PriUniDif = Format(vlRegistro!Mto_PriUniDif, "#,#0.00")
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

Function flValidaFecha(iFecha)
On Error GoTo Err_valfecha

    flValidaFecha = False
    If (Trim(iFecha) = "") Then
       MsgBox "Debe Ingresar Fecha de Traspaso", vbCritical, "Error de Datos"
       iFecha.SetFocus
       Exit Function
    End If
    If Not IsDate(iFecha.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       iFecha.SetFocus
       Exit Function
    End If
    If (CDate(iFecha) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       iFecha.SetFocus
       Exit Function
    End If
    If (Year(iFecha) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       iFecha.SetFocus
       Exit Function
    End If
    flValidaFecha = True

Exit Function
Err_valfecha:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(2) & Err.Description & " ]"
    End Select
End Function

Function flGrabaRecepcion()
    Dim vlNumLiquidacion As String
    Dim vlNumFactura As String, vlNumRenVit As String
    
    On Error GoTo Err_Grabar
        
    If Not fgConexionBaseDatos(vgConectarBD) Then
        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
        Exit Function
    End If
    
    'Comenzar la Transacción
    vgConectarBD.BeginTrans
    
'I--- ABV 14/10/2007 ---
''El Número de Liquidación es el mismo de la Boleta o Factura de Pago
'    If Not flObtieneNumLiquidacion(vgConectarBD, vlNumLiquidacion, vlNumRenVit) Then
'        MsgBox "Error al Obtener Número de Liquidación.", vbCritical, "Error de Datos"
'        Exit Function
'    End If
    
''    If Not flObtieneNumFactura(vgConectarBD, vlNumFactura, vlNumRenVit) Then
''        MsgBox "Error al Obtener Número de Factura.", vbCritical, "Error de Datos"
''        Exit Function
''    End If

    vlNumLiquidacion = 0
    vlNumRenVit = "002"
    vlNumFactura = vlNumLiquidacion
'F--- ABV 14/10/2007 ---
    vlMtoVarFon = 0
    vlMtoVarTC = 0
    vlMtoVarFonTC = 0

    'Inserta los Datos en la Tabla de pd_tmae_polprirecaux
    Sql = "INSERT INTO pd_tmae_polprirecaux ("
    Sql = Sql & "num_poliza,fec_traspaso,fec_vigencia,mto_priinf,"
    Sql = Sql & "mto_pensioninf,mto_pensiongarinf,mto_prirecpesos,"
    Sql = Sql & "mto_prirec,prc_facvar,mto_pension,mto_pensiongar,"
    Sql = Sql & "cod_usuariocrea,fec_crea,hor_crea,"
    Sql = Sql & "cod_monedapriinf,cod_monedapeninf,mto_valmonedarec,mto_valmonedainf,"
    Sql = Sql & "mto_penvarfon,mto_penvartc,mto_penvarfontc,cod_liquidacion"
    Sql = Sql & ",cod_factura,cod_renvit "
    Sql = Sql & ",mto_priafp,mto_pensionafp,mto_pricia "
'I--- ABV 07/11/2007 ---
    Sql = Sql & ",mto_sumpensioninf,mto_sumpension,mto_sumpensionafp "
'F--- ABV 07/11/2007 ---
    Sql = Sql & ") VALUES ("
    Sql = Sql & "'" & Trim(Txt_Poliza) & "',"
    Sql = Sql & "'" & Trim(vlFTraspaso) & "',"
    Sql = Sql & "'" & Trim(vlFVigencia) & "',"
    Sql = Sql & " " & Str(Lbl_PrimaInf) & ","
    Sql = Sql & " " & Str(Lbl_PensionInf) & ","
    Sql = Sql & " " & Str(vlMtoPenGar) & ","
    Sql = Sql & " " & Str(Txt_MtoPrimaRec) & ","
    Sql = Sql & " " & Str(Txt_MtoPrimaRec) & ","
    Sql = Sql & " " & Str(Lbl_FactVarRta) & ","
    Sql = Sql & " " & Str(Lbl_MtoPenDefUf) & ","
    If (vlMtoPenGarUf) = "" Then
        Sql = Sql & " " & Str(Format("0", "#0.00")) & ","
    Else
        Sql = Sql & " " & Str(vlMtoPenGarUf) & ","
    End If
    Sql = Sql & "'" & (vgUsuario) & "',"
    Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
    Sql = Sql & "'" & Format(Time, "hhmmss") & "',"
    Sql = Sql & "'" & Lbl_Moneda(ciMonedaPrimaAnt) & "',"
    Sql = Sql & "'" & Lbl_Moneda(ciMonedaPensionAnt) & "',"
    Sql = Sql & " " & Str(vlTipoCambioRec) & ","
    Sql = Sql & " " & Str(vlTipoCambioAcep) & ","
    Sql = Sql & " " & Str(vlMtoVarFon) & ","
    Sql = Sql & " " & Str(vlMtoVarTC) & ","
    Sql = Sql & " " & Str(vlMtoVarFonTC) & ","
    Sql = Sql & " " & Str(vlNumLiquidacion)
    Sql = Sql & ",'" & Str(vlNumFactura) & "'"
    Sql = Sql & ",'" & (vlNumRenVit) & "'"
    Sql = Sql & "," & Str(Lbl_PrimaDefAFP) & " "
    Sql = Sql & "," & Str(Lbl_MtoPenDefUfAFP) & " "
    Sql = Sql & "," & Str(Lbl_PrimaDefCia) & " "
'I--- ABV 07/11/2007 ---
    Sql = Sql & "," & Str(Lbl_SumPenInf) & " "
    Sql = Sql & "," & Str(Lbl_SumPenDef) & " "
    Sql = Sql & "," & Str(Lbl_SumPenDefAFP) & " "
'F--- ABV 07/11/2007 ---
    Sql = Sql & ")"
    vgConectarBD.Execute (Sql)
    
    'Ejecutar la Transacción
    vgConectarBD.CommitTrans
    
    'Cerrar la Transacción
    vgConectarBD.Close
    
    vgPalabra = "S"
        
Exit Function
Err_Grabar:
    
    'Deshacer la Transacción
    vgConectarBD.RollbackTrans
    'Cerrar la Transacción
    vgConectarBD.Close
    
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flEliminarPolizaOri()

    vgQuery = ""
    vgQuery = "DELETE FROM pd_tmae_oripoliza WHERE "
    vgQuery = vgQuery & "num_poliza = '" & Trim(Txt_Poliza) & "'"
    vgConectarBD.Execute (vgQuery)
     
End Function

Function flEliminarBeneficiariosOri()
            
    vgQuery = ""
    vgQuery = "DELETE FROM pd_tmae_oripolben WHERE "
    vgQuery = vgQuery & "num_poliza = '" & Trim(Txt_Poliza) & "'"
    vgConectarBD.Execute (vgQuery)
            
End Function

Private Function flObtieneFactor(iPrimaRec As Double, iPrimaInf As Double, iCambioRec As Double, iCambioInf As Double) As Double

    Dim vlPrimaInf As Double
    Dim vlPrimaRec As Double
    Dim vlFactVarRta As Double
    
    flObtieneFactor = 0
    
     vlPrimaInf = Format((iPrimaInf / iCambioInf), "##0.00")
     vlPrimaRec = Format((iPrimaRec / iCambioRec), "##0.00")
     vlFactVarRta = Format((vlPrimaRec - vlPrimaInf), "#,#0.00")
     
     If vlPrimaInf > 0 Then
         vlFactVarRta = Format(vlFactVarRta / vlPrimaInf, "#0.00000000")
     End If
     
     flObtieneFactor = vlFactVarRta
End Function

Function flEliminarBonosOri()
            
    vgQuery = ""
    vgQuery = "DELETE FROM pd_tmae_oripolbon WHERE "
    vgQuery = vgQuery & "num_poliza = '" & Trim(Txt_Poliza) & "'"
    vgConectarBD.Execute (vgQuery)
            
End Function

Private Function flLlenaGrillaVariacion(Msf_Grilla As MSFlexGrid, iPrimaRec As Double, iPrimaInf As Double, iCambioRec As Double, iCambioInf As Double, iPensionInf As Double, iMtoVar As Double) As Boolean
    Dim vlFactVarRta As Double
    Dim vlFactorDef As Double
    Dim vlPension As Double
    Dim vlVarFondo As Double, vlVarCambio As Double, vlVarPension As Double 'Usados en Calculo
    
    flLlenaGrillaVariacion = False
    
    'Llena Grillas de Variación
    Msf_Grilla.Row = 1
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = Format(iPrimaInf, "###,##0.00")
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = Format(iCambioInf, "###,##0.00")
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = Format(iPensionInf, "###,##0.00")
    
    Msf_Grilla.Row = 2
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = Format(iPrimaRec, "###,##0.00")
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = Format(iCambioRec, "###,##0.00")
    Msf_Grilla.Col = 3
    vlFactVarRta = flObtieneFactor(iPrimaRec, iPrimaInf, iCambioRec, iCambioInf)
    vlFactorDef = Format((iPensionInf * vlFactVarRta), "#,##0.00")
    vlPension = Format((iPensionInf + (vlFactorDef)), "#,#0.00")
    Msf_Grilla.Text = Format(vlPension, "###,##0.00")
        
    vlVarFondo = ((iPrimaRec / iPrimaInf) - 1) * 100
    vlVarCambio = ((iCambioRec / iCambioInf) - 1) * 100
    vlVarPension = ((vlPension / iPensionInf) - 1) * 100
    iMtoVar = vlPension
    
    Msf_Grilla.Row = 3
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = Format(vlVarFondo, "###,##0.00") & " %"
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = Format(vlVarCambio, "###,##0.00") & " %"
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = Format(vlVarPension, "###,##0.00") & " %"
    
    flLlenaGrillaVariacion = True
End Function

Private Function flObtieneNumFactura(iConexion As ADODB.Connection, oNumFactura As String, oNumRenVit As String) As Boolean
    Dim vlSql As String
    'genera nuevo numero de Factura
    Dim vlNewNumFactura As Long
    Dim vlNewNumRenVit  As String
    
    flObtieneNumFactura = False
    
    vgSql = "SELECT num_factura, cod_renvit "
    vgSql = vgSql & "FROM pd_tmae_gennumfac WHERE num_factura = "
    vgSql = vgSql & " (SELECT MAX(num_factura) FROM pd_tmae_gennumfac)"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        vlNewNumFactura = CInt((vgRs!num_factura)) + 1
        vlNewNumRenVit = Trim(vgRs!cod_renvit)
    Else
        vlNewNumFactura = 1
        vlNewNumRenVit = "001"
    End If
    
    If vlNewNumFactura = 1 Then
        vlSql = "INSERT INTO pd_tmae_gennumfac "
        vlSql = vlSql & "(num_factura,cod_renvit,"
        vlSql = vlSql & "cod_usuariocrea, fec_crea, hor_crea) "
        vlSql = vlSql & " VALUES ('"
        vlSql = vlSql & Trim(Str(vlNewNumFactura)) & "','"
        vlSql = vlSql & Trim(Str(vlNewNumRenVit)) & "','"
        vlSql = vlSql & vgUsuario & "','"
        vlSql = vlSql & Format(Date, "yyyymmdd") & "','"
        vlSql = vlSql & Format(Time, "hhmmss") & "')"
    Else
        vlSql = "UPDATE pd_tmae_gennumfac SET "
        vlSql = vlSql & "num_factura = '" & Trim(Str(vlNewNumFactura)) & "',"
        vlSql = vlSql & "cod_renvit = '" & Trim(Str(vlNewNumRenVit)) & "',"
        vlSql = vlSql & "cod_usuariomodi = '" & vgUsuario & "',"
        vlSql = vlSql & "fec_modi = '" & Format(Date, "yyyymmdd") & "',"
        vlSql = vlSql & "hor_modi = '" & Format(Time, "hhmmss") & "'"
    End If
    iConexion.Execute (vlSql)
    
    oNumFactura = vlNewNumFactura
    oNumRenVit = vlNewNumRenVit
    
    flObtieneNumFactura = True
End Function
