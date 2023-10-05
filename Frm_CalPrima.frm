VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_CalPrima 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Ingreso de Primas."
   ClientHeight    =   8400
   ClientLeft      =   3225
   ClientTop       =   555
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11040
   Begin TabDlg.SSTab SSTab2 
      Height          =   6255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11033
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Antecedentes Generales"
      TabPicture(0)   =   "Frm_CalPrima.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_AntGral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_Definitiva"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Fra_Dif"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Fra_AntRec"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Primer Pago"
      TabPicture(1)   =   "Frm_CalPrima.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Variación"
      TabPicture(2)   =   "Frm_CalPrima.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Fra_AntIni"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).ControlCount=   3
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
         Height          =   1275
         Left            =   3840
         TabIndex        =   46
         Top             =   4080
         Width           =   4815
         Begin VB.TextBox Txt_FecTraspaso 
            Height          =   285
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   7
            Top             =   285
            Width           =   1215
         End
         Begin VB.TextBox Txt_FecPriPag 
            Height          =   285
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha Traspaso Prima"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   50
            Top             =   285
            Width           =   1815
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha Primer Pago"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   48
            Top             =   885
            Width           =   1695
         End
         Begin VB.Label Lbl_FecTopePriPag 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2160
            TabIndex        =   47
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha Tope Primer Pago"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Fra_Dif 
         Caption         =   " Resultado del Cálculo  "
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
         Height          =   795
         Left            =   3840
         TabIndex        =   70
         Top             =   5350
         Width           =   4815
         Begin VB.CommandButton Cmd_Calcular 
            Caption         =   "&Calcular"
            Height          =   675
            Left            =   3840
            Picture         =   "Frm_CalPrima.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Cálcular Pensión"
            Top             =   110
            Width           =   720
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha Vigencia Póliza"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Lbl_FecVigPoliza 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2160
            TabIndex        =   71
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Fra_Definitiva 
         Caption         =   "  Distribución de la Prima Traspasada  "
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
         Height          =   2055
         Left            =   120
         TabIndex        =   57
         Top             =   4080
         Width           =   3615
         Begin VB.Label Lbl_SumPenDefAFP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   360
            TabIndex        =   87
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Lbl_SumPenDef 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   360
            TabIndex        =   86
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Prima Definitiva"
            Height          =   255
            Index           =   27
            Left            =   240
            TabIndex        =   84
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Pensión Definitiva"
            Height          =   255
            Index           =   26
            Left            =   240
            TabIndex        =   83
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "(TM)"
            Height          =   255
            Index           =   5
            Left            =   1635
            TabIndex        =   68
            Top             =   1695
            Width           =   375
         End
         Begin VB.Label Lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "(TM)"
            Height          =   255
            Index           =   6
            Left            =   1635
            TabIndex        =   69
            Top             =   1380
            Width           =   375
         End
         Begin VB.Label Lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "(TM)"
            Height          =   255
            Index           =   4
            Left            =   1640
            TabIndex        =   65
            Top             =   860
            Width           =   375
         End
         Begin VB.Label Lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "(TM)"
            Height          =   255
            Index           =   0
            Left            =   1640
            TabIndex        =   61
            Top             =   540
            Width           =   375
         End
         Begin VB.Label Lbl_PrimaDefAFP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   67
            Top             =   1695
            Width           =   1215
         End
         Begin VB.Label Lbl_PrimaDefCia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   66
            Top             =   860
            Width           =   1215
         End
         Begin VB.Label Lbl_MtoPenDefUfAFP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   64
            Top             =   1380
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
            Left            =   2040
            TabIndex        =   63
            Top             =   1125
            Width           =   1215
         End
         Begin VB.Label Lbl_MtoPenDefUf 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   62
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Pensión Definitiva"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   60
            Top             =   540
            Width           =   1335
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Prima Definitiva"
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   59
            Top             =   855
            Width           =   1815
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
            Left            =   2040
            TabIndex        =   58
            Top             =   280
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "  Resumen Primer Pago  "
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
         Height          =   2175
         Left            =   -74880
         TabIndex        =   51
         Top             =   480
         Width           =   8535
         Begin MSFlexGridLib.MSFlexGrid MSF_Resumen 
            Height          =   1695
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            BackColor       =   14745599
            FormatString    =   "Tipo Ident.|Num. Ident           |Parentesco|Fecha Inicio | Fecha Término| Monto Pensión | Descuento Salud | Líquido a Pagar"
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "  Detalle Primer Pago  "
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
         Height          =   3375
         Left            =   -74880
         TabIndex        =   44
         Top             =   2760
         Width           =   8535
         Begin MSFlexGridLib.MSFlexGrid MSF_Detalle 
            Height          =   3015
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   5318
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            BackColor       =   14745599
            FormatString    =   "Tipo Ident.|Num. Ident           |Parentesco|Fecha Inicio | Fecha Término| Monto Pensión | Descuento Salud | Líquido a Pagar"
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Variación Fondo y Tipo de Cambio"
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
         Height          =   1455
         Left            =   -74520
         TabIndex        =   29
         Top             =   4320
         Width           =   7815
         Begin MSFlexGridLib.MSFlexGrid Msf_VarFondoTC 
            Height          =   1080
            Left            =   120
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   1905
            _Version        =   393216
            BackColor       =   14745599
            GridColor       =   0
            AllowBigSelection=   0   'False
         End
      End
      Begin VB.Frame Fra_AntIni 
         Caption         =   "Variación Tipo de Cambio"
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
         Height          =   1455
         Left            =   -74520
         TabIndex        =   27
         Top             =   2520
         Width           =   7815
         Begin MSFlexGridLib.MSFlexGrid Msf_VarTipCambio 
            Height          =   1080
            Left            =   120
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   1905
            _Version        =   393216
            BackColor       =   14745599
            GridColor       =   0
            AllowBigSelection=   0   'False
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Variación Fondo"
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
         Height          =   1455
         Left            =   -74520
         TabIndex        =   25
         Top             =   720
         Width           =   7815
         Begin MSFlexGridLib.MSFlexGrid Msf_VarFondo 
            Height          =   1080
            Left            =   120
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   1905
            _Version        =   393216
            BackColor       =   14745599
            GridColor       =   0
            AllowBigSelection=   0   'False
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
         Height          =   3615
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   8535
         Begin VB.Label lbl_TR 
            Height          =   255
            Left            =   5640
            TabIndex        =   95
            Top             =   1320
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Lbl_ReajusteDescripcion 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   94
            Top             =   1875
            Width           =   3615
         End
         Begin VB.Label Lbl_ReajusteValorMen 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   7320
            TabIndex        =   93
            Top             =   795
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Lbl_ReajusteValor 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   7320
            TabIndex        =   89
            Top             =   1875
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Moneda"
            Height          =   255
            Index           =   29
            Left            =   240
            TabIndex        =   92
            Top             =   1875
            Width           =   1215
         End
         Begin VB.Label Lbl_ReajusteTipo 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6000
            TabIndex        =   91
            Top             =   795
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Reajuste Trim."
            Height          =   255
            Index           =   28
            Left            =   6240
            TabIndex        =   90
            Top             =   1875
            Width           =   1095
         End
         Begin VB.Label Lbl_SumPenInf 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3480
            TabIndex        =   85
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Lbl_Meses 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   7320
            TabIndex        =   19
            Top             =   1605
            Width           =   1095
         End
         Begin VB.Label Lbl_Diferidos 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   7320
            TabIndex        =   20
            Top             =   1335
            Width           =   1095
         End
         Begin VB.Label Lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "(TM)"
            Height          =   255
            Index           =   2
            Left            =   1635
            TabIndex        =   82
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Prima Traspasada"
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   81
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label Lbl_MtoPrimaRec 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   80
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo de Cambio"
            Height          =   255
            Index           =   17
            Left            =   5040
            TabIndex        =   79
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Factor Variación Rta."
            Height          =   255
            Index           =   9
            Left            =   5040
            TabIndex        =   78
            Top             =   2955
            Width           =   1575
         End
         Begin VB.Label Lbl_FactVarRta 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6960
            TabIndex        =   77
            Top             =   2955
            Width           =   1215
         End
         Begin VB.Label Lbl_TipoCambio 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6960
            TabIndex        =   76
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Lbl_FecSolicitud 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   74
            Top             =   2955
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha Información AFP"
            Height          =   255
            Index           =   25
            Left            =   240
            TabIndex        =   75
            Top             =   2955
            Width           =   1935
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   120
            X2              =   8280
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Label Lbl_FecAceptacion 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   55
            Top             =   2550
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha Aceptación"
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   56
            Top             =   2550
            Width           =   1575
         End
         Begin VB.Label Lbl_FecDevengue 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   54
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Fecha Devengue"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   53
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Lbl_NumIdent 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4800
            TabIndex        =   43
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Lbl_TipoIdent 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   42
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Lbl_Modalidad 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   41
            Top             =   1605
            Width           =   3615
         End
         Begin VB.Label Lbl_TipoRenta 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   40
            Top             =   1335
            Width           =   3615
         End
         Begin VB.Label Lbl_TipoPension 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   39
            Top             =   1065
            Width           =   6375
         End
         Begin VB.Label Lbl_NomAfiliado 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   38
            Top             =   510
            Width           =   6375
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Nº Identificación"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Modalidad"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   36
            Top             =   1605
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo de Renta"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   35
            Top             =   1335
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo de Pensión"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   34
            Top             =   1065
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   33
            Top             =   510
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "CUSPP"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   32
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label Lbl_CUSPP 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   31
            Top             =   795
            Width           =   2535
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Meses Garant."
            Height          =   255
            Index           =   6
            Left            =   6240
            TabIndex        =   24
            Top             =   1605
            Width           =   1095
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Años Diferidos"
            Height          =   255
            Index           =   5
            Left            =   6240
            TabIndex        =   23
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Line Lin_Separar 
            BorderColor     =   &H00808080&
            X1              =   240
            X2              =   8400
            Y1              =   2200
            Y2              =   2200
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Prima Cotizada"
            Height          =   195
            Index           =   7
            Left            =   5040
            TabIndex        =   22
            Top             =   2280
            Width           =   1050
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Pensión Cotizada"
            Height          =   195
            Index           =   8
            Left            =   5040
            TabIndex        =   21
            Top             =   2550
            Width           =   1230
         End
         Begin VB.Label Lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "(TM)"
            Height          =   255
            Index           =   1
            Left            =   6525
            TabIndex        =   18
            Top             =   2550
            Width           =   375
         End
         Begin VB.Label Lbl_PensionInf 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6960
            TabIndex        =   17
            Top             =   2550
            Width           =   1455
         End
         Begin VB.Label Lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "(TM)"
            Height          =   255
            Index           =   3
            Left            =   6525
            TabIndex        =   16
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Lbl_PrimaInf 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6960
            TabIndex        =   15
            Top             =   2280
            Width           =   1455
         End
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
      Height          =   8325
      Left            =   9000
      TabIndex        =   5
      Top             =   0
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
         Height          =   7830
         Left            =   120
         TabIndex        =   6
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
         TabIndex        =   9
         Top             =   2400
         Width           =   735
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
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   4920
         Picture         =   "Frm_CalPrima.frx":04F6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Efectuar Busqueda de la Póliza"
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox Txt_Poliza 
         Height          =   285
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº de Póliza                 :"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   7200
      Width           =   8775
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&AFP"
         Height          =   675
         Index           =   0
         Left            =   2280
         Picture         =   "Frm_CalPrima.frx":05F8
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5775
         Picture         =   "Frm_CalPrima.frx":0CB2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4680
         Picture         =   "Frm_CalPrima.frx":0DAC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   3600
         Picture         =   "Frm_CalPrima.frx":1466
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar Datos"
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
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   3285
         X2              =   3285
         Y1              =   240
         Y2              =   960
      End
   End
End
Attribute VB_Name = "Frm_CalPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'DECLARACION DE VARIABLES
Dim vlRegistro    As ADODB.Recordset, vlRegistro1 As ADODB.Recordset
Dim vlRegistro2   As ADODB.Recordset, vlRegistro3 As ADODB.Recordset, vlRegistro4 As ADODB.Recordset
Dim vlReg         As ADODB.Recordset, vlDif       As Integer
Dim vlPasa        As Boolean, Gls_Pension   As String, Gls_Renta   As String
Dim Gls_Modalidad As String, iFecha         As String, vlFTraspaso As String
Dim vlAnno        As String, vlMes          As String, vlDia       As String
Dim vlCodTp        As String, vlCodTr     As String
Dim vlCodAl       As String, vlCodPar       As String, vlFBRec     As String
Dim vlFBComp      As String, vlFBExon       As String, vlBonRec    As String
Dim vlBonComp     As String, vlBonExon      As String, vlFVigencia As String
Dim vlMtoPenGar   As String, vlMtoPenGarUf  As String, vlFecIniPP  As String
Dim vlFactorDef   As Double, vlFactorDefGar As Double, vlCtaInd    As Double
Dim vlCtaIndMod    As Double
Dim vlMtoTotalBono As Double
Dim vlafp As String
Dim vlNombreComuna As String
Dim vlJefeBeneficios As String
Dim vlCodTipReajuste As String 'hqr 13/01/2011
Dim vlMtoValReajusteTri As Double 'hqr 13/01/2011
Dim vlMtoValReajusteMen As Double

Dim vlNomBen, vlApeBenPat, vlApeBenMat As String


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

Dim vlFecAcepta As String, vlFecDev As String, vlFecVigencia As String
Dim vlTipoPen As String, vlTipoRen As String, vlMesDif As Long
Dim vlMonedaPension As String
Dim vlTipoCambioAcep As Double ' A la Fecha de Aceptacion
Dim vlTipoCambioRec As Double ' A la Fecha de Recepciòn de la Prima
Dim vlMtoVarFon As Double, vlMtoVarTC As Double, vlMtoVarFonTC As Double 'Para Insertar
Dim vlCalculoPrimerPago As Boolean 'Para verificar si se calculó exitosamente el Monto del Primer Pago
Dim vlFecTraspasoPrimas As String

Function flDespliegaPrimerPago()
Dim i As Long
Dim vlNumIdent As String, vlCodTipoIdent As Long, vlNumOrden As Long
Dim vlFilaResumen As Long
Dim vlMontoPension As Double, vlMontoDescuento As Double, vlLiquido As Double
Dim vlValidaPension As String

vlNumIdent = ""
vlCodTipoIdent = 0
vlNumOrden = 0
vlFilaResumen = 0
vlMontoPension = 0
vlMontoDescuento = 0
vlLiquido = 0
vlValidaPension = ""
Call CorreccionMontosCero

MSF_Detalle.rows = 1 'Detalle por Beneficiario
MSF_Resumen.rows = 1 'Resumen por Beneficiario
    For i = 0 To UBound(stLiquidacion) - 1
        'Resumen
        If stLiquidacion(i).Cod_TipoIdenPensionado <> vlCodTipoIdent Or stLiquidacion(i).Num_IdenPensionado <> vlNumIdent Or stLiquidacion(i).Num_Orden <> vlNumOrden Then
            'Nuevo Registro
            vlFilaResumen = vlFilaResumen + 1
            vlMontoPension = 0
            vlMontoDescuento = 0
            vlLiquido = 0
            vlCodTipoIdent = stLiquidacion(i).Cod_TipoIdenPensionado
            vlNumIdent = stLiquidacion(i).Num_IdenPensionado
            vlNumOrden = stLiquidacion(i).Num_Orden
            MSF_Resumen.rows = vlFilaResumen + 1
            MSF_Resumen.Row = vlFilaResumen
            MSF_Resumen.Col = 0
            MSF_Resumen.Text = stLiquidacion(i).Gls_TipoIdentCor
            MSF_Resumen.Col = 1
            MSF_Resumen.Text = stLiquidacion(i).Num_IdenPensionado
            MSF_Resumen.Col = 2
            MSF_Resumen.Text = stLiquidacion(i).Cod_Parentesco
            MSF_Resumen.Col = 3
            MSF_Resumen.Text = DateSerial(Mid(stLiquidacion(i).Fec_IniPago, 1, 4), Mid(stLiquidacion(i).Fec_IniPago, 5, 2), Mid(stLiquidacion(i).Fec_IniPago, 7, 2))
            MSF_Resumen.Col = 4
            MSF_Resumen.Text = DateSerial(Mid(stLiquidacion(i).Fec_TerPago, 1, 4), Mid(stLiquidacion(i).Fec_TerPago, 5, 2), Mid(stLiquidacion(i).Fec_TerPago, 7, 2))
            vlMontoPension = stLiquidacion(i).Mto_Pension
            vlMontoDescuento = stLiquidacion(i).Mto_Salud
            vlLiquido = stLiquidacion(i).Mto_LiqPagar
            MSF_Resumen.Col = 5
            MSF_Resumen.Text = Format(vlMontoPension, "###,##0.00")
            MSF_Resumen.Col = 6
            MSF_Resumen.Text = Format(vlMontoDescuento, "###,##0.00")
            MSF_Resumen.Col = 7
            MSF_Resumen.Text = Format(vlLiquido, "###,##0.00")
        Else
            'El mismo registro anterior
            MSF_Resumen.Col = 4
            MSF_Resumen.Text = DateSerial(Mid(stLiquidacion(i).Fec_TerPago, 1, 4), Mid(stLiquidacion(i).Fec_TerPago, 5, 2), Mid(stLiquidacion(i).Fec_TerPago, 7, 2))
            vlMontoPension = vlMontoPension + stLiquidacion(i).Mto_Pension
            vlMontoDescuento = vlMontoDescuento + stLiquidacion(i).Mto_Salud
            vlLiquido = vlLiquido + stLiquidacion(i).Mto_LiqPagar
            MSF_Resumen.Col = 5
            MSF_Resumen.Text = Format(vlMontoPension, "###,##0.00")
            MSF_Resumen.Col = 6
            MSF_Resumen.Text = Format(vlMontoDescuento, "###,##0.00")
            MSF_Resumen.Col = 7
            MSF_Resumen.Text = Format(vlLiquido, "###,##0.00")
        End If
        
        'Detalle
        MSF_Detalle.rows = i + 2
        MSF_Detalle.Row = i + 1
        MSF_Detalle.Col = 0
        MSF_Detalle.Text = stLiquidacion(i).Gls_TipoIdentCor
        MSF_Detalle.Col = 1
        MSF_Detalle.Text = stLiquidacion(i).Num_IdenPensionado
        MSF_Detalle.Col = 2
        MSF_Detalle.Text = stLiquidacion(i).Cod_Parentesco
        MSF_Detalle.Col = 3
        MSF_Detalle.Text = DateSerial(Mid(stLiquidacion(i).Fec_IniPago, 1, 4), Mid(stLiquidacion(i).Fec_IniPago, 5, 2), Mid(stLiquidacion(i).Fec_IniPago, 7, 2))
        MSF_Detalle.Col = 4
        MSF_Detalle.Text = DateSerial(Mid(stLiquidacion(i).Fec_TerPago, 1, 4), Mid(stLiquidacion(i).Fec_TerPago, 5, 2), Mid(stLiquidacion(i).Fec_TerPago, 7, 2))
        MSF_Detalle.Col = 5
        MSF_Detalle.Text = Format(stLiquidacion(i).Mto_Pension, "###,##0.00")
        MSF_Detalle.Col = 6
        MSF_Detalle.Text = Format(stLiquidacion(i).Mto_Salud, "###,##0.00")
        MSF_Detalle.Col = 7
        MSF_Detalle.Text = Format(stLiquidacion(i).Mto_LiqPagar, "###,##0.00")
        
        If Format(stLiquidacion(i).Mto_Pension, "###,##0.00") = "0.00" Then
            vlValidaPension = Format(stLiquidacion(i).Mto_Pension, "###,##0.00")
        End If
    Next i

        If vlValidaPension = "0.00" Then
            MsgBox "Se ha generado registros con pensión en cero, Por favor Revise la Pestaña de Primer Pago", vbExclamation, "Primer Pago"
        End If
        
End Function

Function flGrabaPensionActualizada(i As Long) As Boolean
On Error GoTo Errores
flGrabaPensionActualizada = False
If stLiquidacion(i).Fac_Ajuste <> 1 Then
    'Graba Pension Total Actualizada
    Sql = "INSERT INTO PD_TMAE_PENSIONACT "
    Sql = Sql & "(NUM_POLIZA, FEC_DESDE, MTO_PENSION, MTO_PENSIONGAR, PRC_FATORAJUS) VALUES ('"
    Sql = Sql & stLiquidacion(i).Num_Poliza & "','" & stLiquidacion(i).Fec_IniPago & "',"
    Sql = Sql & str(stLiquidacion(i).Mto_PensionTotal) & ","
    Sql = Sql & str(stLiquidacion(i).Mto_pensiongarTotal) & ","
    Sql = Sql & str(stLiquidacion(i).prc_factorAjus) & ")"
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
        vlSql = vlSql & Trim(str(vlNewNumLiq)) & "','"
        vlSql = vlSql & Trim(str(vlNewNumRenVit)) & "','"
        vlSql = vlSql & vgUsuario & "','"
        vlSql = vlSql & Format(Date, "yyyymmdd") & "','"
        vlSql = vlSql & Format(Time, "hhmmss") & "')"
    Else
        vlSql = "UPDATE pd_tmae_gennumliq SET "
        vlSql = vlSql & "num_liquidacion = '" & Trim(str(vlNewNumLiq)) & "',"
        vlSql = vlSql & "cod_renvit = '" & Trim(str(vlNewNumRenVit)) & "',"
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
Dim vlPrcPen As Double
Dim vlExistenPensiones As Boolean
Dim vlProxPeriodo As String

On Error GoTo Err_Cal

    'Valida que se hayan cargado los Valores Originales de la Póliza
    If (Lbl_PrimaInf) = "" Or (Lbl_PensionInf) = "" Then
        MsgBox "No se encuentran especificados los valores Originales de la Póliza.", vbCritical, "Error de Datos"
        Cmd_Limpiar.SetFocus
        Exit Sub
    End If
   
    'Valida el Ingreso de la Póliza
    If Fra_Poliza.Enabled = False Then
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
    If Lbl_MtoPrimaRec = "" Then
       MsgBox "Debe Ingresar un valor para Monto Prima Recibida o Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    If CDbl(Lbl_MtoPrimaRec) = 0 Then
       MsgBox "Debe Ingresar un Valor para Monto Prima Recibida o Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    
    'Valida el ingreso de la Fecha de Primer Pago
    If Txt_FecPriPag = "" Then
       MsgBox "Debe ingresar la Fecha del Primer Pago.", vbCritical, "Error de Datos"
       Txt_FecPriPag.SetFocus
       Exit Sub
    End If

    If Not IsDate(Txt_FecPriPag.Text) Then
       MsgBox "La Fecha de Primer Pago No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_FecPriPag.SetFocus
       Exit Sub
    End If
    If vlMesDif = 0 Then
        If (CDate(Txt_FecPriPag) < CDate(Date)) Then
           MsgBox "La Fecha Ingresada es Menor a la Fecha Actual", vbCritical, "Error de Datos"
           Txt_FecPriPag.SetFocus
           Exit Sub
        End If
        If CDate(Txt_FecPriPag) > CDate(DateAdd("d", clNumeroMaximoDiasPrimerPago, Date)) Then
            MsgBox "Fecha de Primer Pago es Mayor a la Fecha Actual por más de '" & clNumeroMaximoDiasPrimerPago & "' días", vbCritical, "Error de Datos"
            Txt_FecPriPag.SetFocus
            Exit Sub
        End If
    Else
        If Trim(Lbl_FecTopePriPag) <> Trim(Txt_FecPriPag) Then 'Se modificó la Fecha de Primer Pago
        
            If lbl_TR <> "6" Then
                If (CDate(Txt_FecPriPag) < CDate(Date)) Then
                   MsgBox "La Fecha de Primer Pago Ingresada es Menor a la Fecha Actual", vbCritical, "Error de Datos"
                   Txt_FecPriPag.SetFocus
                   Exit Sub
                End If
                If CDate(Txt_FecPriPag) > CDate(DateAdd("d", clNumeroMaximoDiasPrimerPago, Date)) Then
                    MsgBox "Fecha de Primer Pago es Mayor a la Fecha Actual por más de '" & clNumeroMaximoDiasPrimerPago & "' días", vbCritical, "Error de Datos"
                    Txt_FecPriPag.SetFocus
                    Exit Sub
                End If
            End If
           
            
'            'Validar que el siguiente periodo no se encuentre 'Cerrado'
'            If Not flValidaPeriodoPagoRegimen(Format(DateAdd("m", 1, Txt_FecPriPag), "yyyymm"), vlProxPeriodo) Then
'                MsgBox "Fecha de Primer Pago no válida." & IIf(vlProxPeriodo = "", "", Chr(13) & "Proximo Pago Recurrente se realizará en el Periodo '" & vlProxPeriodo & "'"), vbCritical, "Error de Datos"
'                Txt_FecPriPag.SetFocus
'                Exit Sub
'            End If
        Else '= Se pagará el primer pago Diferido como Pago en Régimen
            If (Format(Txt_FecPriPag, "yyyymm") < Format(Date, "yyyymm")) Then
               MsgBox "La Fecha de Primer Pago Ingresada no debe ser Menor al Mes Actual", vbCritical, "Error de Datos"
               Txt_FecPriPag.SetFocus
               Exit Sub
            End If
            If Format(Lbl_FecTopePriPag, "yyyymmdd") > Format(Txt_FecPriPag, "yyyymmdd") Then 'Se modificó la Fecha de Primer Pago
               MsgBox "La Fecha de Tope del Primer Pago definida no debe ser Mayor a la Fecha de Primer Pago (o Inicio del Pago Diferido)", vbCritical, "Error de Datos"
               Txt_FecPriPag.SetFocus
               Exit Sub
            End If
            'Validar Fecha de Primer Pago (El mes será un pago en régimen, validar que no esté cerrado)
            If Not flValidaPeriodoPagoRegimen(Format(Txt_FecPriPag, "yyyymm"), vlProxPeriodo) Then
                MsgBox "Fecha de Primer Pago se encuentra definida para un Periodo de Pago de Pensiones ya 'Cerrado'. " & IIf(vlProxPeriodo = "", "", Chr(13) & "Próximo Pago Recurrente se realizará en el Periodo '" & vlProxPeriodo & "'"), vbCritical, "Error de Datos"
                Txt_FecPriPag.SetFocus
                Exit Sub
            End If
            
        End If
    End If
    
    If (Year(Txt_FecPriPag) < 1900) Then
       MsgBox "La Fecha de Primer Pago es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_FecPriPag.SetFocus
       Exit Sub
    End If
    
    'Validar que la Fecha de Traspaso no sea inferior a la Fecha de Solicitud
    If (Format(Lbl_FecSolicitud, "yyyymmdd") > Format(Txt_FecTraspaso, "yyyymmdd")) Then
        MsgBox "La Fecha de Traspaso es menor que la Fecha de Solicitud a la AFP.", vbCritical, "Error de Datos"
        Txt_FecTraspaso.SetFocus
        Exit Sub
    End If
    
    'Validar existencia de Valor de Tipo de Cambio
    If Not IsNumeric(Lbl_TipoCambio) Then
        MsgBox "No existe valor de Tipo de Cambio para la Prima Solicitada.", vbCritical, "Inexistencia de Datos"
        Cmd_Salir.SetFocus
        Exit Sub
    End If
    'Validar existencia de Factor de Cálculo
    If Not IsNumeric(Lbl_FactVarRta) Then
        MsgBox "No existe valor de Factor de Tipo de Renta para la Prima Solicitada.", vbCritical, "Inexistencia de Datos"
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

'I--- ABV 05/02/2011 ---
    'Valida que se hayan cargado los nuevos valores para el Reajuste
    If (Lbl_ReajusteTipo = "") Or (Lbl_ReajusteValor = "") Or (Lbl_ReajusteValorMen = "") Then
        MsgBox "No se encuentran especificados los valores de Tipo de Reajuste de la Póliza.", vbCritical, "Error de Datos"
        Cmd_Salir.SetFocus
        Exit Sub
    End If
'F--- ABV 05/02/2011 ---
    
    'Fecha Vigencia Póliza, es el primer dia del mes en que estoy haciendo el Traspaso de la prima
    vlFVigencia = Format(CDate(Trim(Txt_FecTraspaso)), "yyyymmdd")
    vlAnno = Mid(vlFVigencia, 1, 4)
    vlMes = Mid(vlFVigencia, 5, 2)
    vlDia = "01"
    
    Lbl_FecVigPoliza = DateSerial(vlAnno, vlMes, vlDia)
'    Lbl_MtoPenDefUf = ""
'    Lbl_FactVarRta = ""

    vlFactVarRta = 0
    vlFactorDef = 0
    vlFactorDefGar = 0
    vlMtoPenGarUf = ""
    
'I--- ABV 05/02/2011 ---
    vlCodTipReajuste = fgObtenerCodigo_TextoCompuesto(Lbl_ReajusteTipo) 'hqr 13/01/2011
    vlMtoValReajusteTri = CDbl(Lbl_ReajusteValor) 'hqr 13/01/2011
    vlMtoValReajusteMen = CDbl(Lbl_ReajusteValorMen) 'ABV 16/02/2011
'F--- ABV 05/02/2011 ---
    
   'Obtiene el Valor de la Moneda a la Fecha de Recepción
'    If Not fgObtieneConversion(vlFVigencia, vlMonedaPension, vlTipoCambioRec) Then
'        Screen.MousePointer = vbNormal
'        MsgBox "No se encuentra registrado el Valor de la Moneda [" & vlMonedaPension & "] para la Fecha de Traspaso de la Prima.", vbCritical, "Advertencia"
'        Exit Sub
'    End If
'
''    If Not fgObtieneConversion(vlFecAcepta, vlMonedaPension, vlTipoCambioAcep) = True Then
''        Screen.MousePointer = vbNormal
''        MsgBox "No se encuentra registrado el Valor de la Moneda [" & vlMonedaPension & "] para la Fecha de Aceptación de la Póliza.", vbCritical, "Advertencia"
''        Exit Sub
''    End If
'
'    Lbl_TipoCambio = vlTipoCambioRec
    vlTipoCambioRec = CDbl(Lbl_TipoCambio)
    
    'Obtiene Variación de la renta
'    vlFactVarRta = flObtieneFactor(CDbl(Lbl_MtoPrimaRec), CDbl(Lbl_PrimaInf), vlTipoCambioRec, vlTipoCambioAcep)
    vlFactVarRta = CDbl(Lbl_FactVarRta)
        
    'Lleva a Porcentaje el Factor calculado
'    Lbl_FactVarRta = Format((vlFactVarRta * 100), "##0.00")
               
    'factor definitivo pension
    vlFactorDef = Format((CDbl(Lbl_PensionInf) * vlFactVarRta), "#,##0.00")
      
    'Aumento o Disminución de la renta o Pensión en UF
'    vlPension = Format((CDbl(Lbl_PensionInf) + (vlFactorDef)), "#,##0.00")
'    Lbl_MtoPenDefUf = vlPension

'I--- ABV 07/11/2007 ---
'    vlPension = CDbl(Lbl_MtoPenDefUf)
    Dim ctaFam As Integer

    vgSql = "select count(*) cuenta from pd_tmae_oripolben where num_poliza='" & Trim(Txt_Poliza) & "' and cod_derpen=99"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        ctaFam = vgRs!cuenta
    End If

    Dim vlPensionRef As Double

    If (fgObtenerCodigo_TextoCompuesto(Lbl_TipoPension) <> clCodTipPensionSob) Then
        vlPension = CDbl(Lbl_MtoPenDefUf)
    Else
        vlPensionRef = CDbl(Lbl_MtoPenDefUf)
        vlPension = CDbl(Lbl_SumPenDef)
    End If
'F--- ABV 07/11/2007 ---
    
    If CDbl(vlMtoPenGar) <> 0 Then
'         vlFactorDefGar = Format((CDbl(vlMtoPenGar) * vlFactVarRta), "#,##0.00")
         vlFactorDefGar = vlFactorDef 'Format((CDbl(vlMtoPenGar) * vlFactVarRta), "#,##0.00")
        'Aumento o Disminución de la renta garantizada
'         vlMtoPenGarUf = Format((CDbl(vlMtoPenGar) + (vlFactorDefGar)), "#,##0.00")
         
'I--- ABV 07/11/2007 ---
'        vlMtoPenGarUf = Format((CDbl(Lbl_PensionInf) * (vlFactorDefGar)), "#,##0.00")

    vlMtoPenGarUf = Format(vlPension, "#,##0.00")
    If vlCodTp = "08" Then
        'vlNumBeneficiarios = 0
        'vlNumConceptos = 0
        vlPrcPen = 0
        vgSql = "SELECT a.num_orden, a.prc_pension "
        vgSql = vgSql & "FROM pd_tmae_oripolben a "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "a.num_poliza = '" & Trim(Txt_Poliza.Text) & "' "
        vgSql = vgSql & "AND a.cod_derpen = 99 " 'campo aun no existe
        vgSql = vgSql & "ORDER BY a.num_orden"
        Set vgRs = vgConexionBD.Execute(vgSql)
        If Not vgRs.EOF Then
            Do While Not (vgRs.EOF)
                vlPrcPen = vlPrcPen + (vgRs!Prc_Pension / 100)
                vgRs.MoveNext
            Loop
            vlMtoPenGarUf = Format(vlPensionRef * vlPrcPen, "#,##0.00")
        End If
    Else
        vlMtoPenGarUf = Format(vlPension, "#,##0.00")
    End If


    
'        If ctaFam = 1 Then
'            vlMtoPenGarUf = Format(CDbl(Lbl_MtoPenDefUf), "#,##0.00")
'        End If
        
'F--- ABV 07/11/2007 ---

'I--- ABV 21/08/2007 ---
    Else
        vlFactorDefGar = 0
        vlMtoPenGarUf = 0
'F--- ABV 21/08/2007 ---
    End If
    
    '1 Variacion Fondos
    vlMtoVarFon = 0
'I--- ABV 07/11/2007 ---
    Call flLlenaGrillaVariacion(Msf_VarFondo, CDbl(Lbl_MtoPrimaRec), CDbl(Lbl_PrimaInf), vlTipoCambioAcep, vlTipoCambioAcep, CDbl(Lbl_PensionInf), vlMtoVarFon)
'    If (fgObtenerCodigo_TextoCompuesto(Lbl_TipoPension) <> clCodTipPensionSob) Then
'        Call flLlenaGrillaVariacion(Msf_VarFondo, CDbl(Lbl_MtoPrimaRec), CDbl(Lbl_PrimaInf), vlTipoCambioAcep, vlTipoCambioAcep, CDbl(Lbl_PensionInf), vlMtoVarFon)
'    Else
'        Call flLlenaGrillaVariacion(Msf_VarFondo, CDbl(Lbl_MtoPrimaRec), CDbl(Lbl_PrimaInf), vlTipoCambioAcep, vlTipoCambioAcep, CDbl(Lbl_SumPenInf), vlMtoVarFon)
'    End If
'F--- ABV 07/11/2007 ---
        
    '2 Variación Tipo de Cambio
    vlMtoVarTC = 0
'I--- ABV 07/11/2007 ---
    Call flLlenaGrillaVariacion(Msf_VarTipCambio, CDbl(Lbl_PrimaInf), CDbl(Lbl_PrimaInf), vlTipoCambioRec, vlTipoCambioAcep, CDbl(Lbl_PensionInf), vlMtoVarTC)
'    If (fgObtenerCodigo_TextoCompuesto(Lbl_TipoPension) <> clCodTipPensionSob) Then
'        Call flLlenaGrillaVariacion(Msf_VarTipCambio, CDbl(Lbl_PrimaInf), CDbl(Lbl_PrimaInf), vlTipoCambioRec, vlTipoCambioAcep, CDbl(Lbl_PensionInf), vlMtoVarTC)
'    Else
'        Call flLlenaGrillaVariacion(Msf_VarTipCambio, CDbl(Lbl_PrimaInf), CDbl(Lbl_PrimaInf), vlTipoCambioRec, vlTipoCambioAcep, CDbl(Lbl_SumPenInf), vlMtoVarTC)
'    End If
'F--- ABV 07/11/2007 ---
    
    '3 Variación Fondo y Tipo de Cambio
    vlMtoVarFonTC = 0
'I--- ABV 07/11/2007 ---
    Call flLlenaGrillaVariacion(Msf_VarFondoTC, CDbl(Lbl_MtoPrimaRec), CDbl(Lbl_PrimaInf), vlTipoCambioRec, vlTipoCambioAcep, CDbl(Lbl_PensionInf), vlMtoVarFonTC)
'    If (fgObtenerCodigo_TextoCompuesto(Lbl_TipoPension) <> clCodTipPensionSob) Then
'        Call flLlenaGrillaVariacion(Msf_VarFondoTC, CDbl(Lbl_MtoPrimaRec), CDbl(Lbl_PrimaInf), vlTipoCambioRec, vlTipoCambioAcep, CDbl(Lbl_PensionInf), vlMtoVarFonTC)
'    Else
'        Call flLlenaGrillaVariacion(Msf_VarFondoTC, CDbl(Lbl_MtoPrimaRec), CDbl(Lbl_PrimaInf), vlTipoCambioRec, vlTipoCambioAcep, CDbl(Lbl_SumPenInf), vlMtoVarFonTC)
'    End If
'F--- ABV 07/11/2007 ---
    
    '4 Calcula Monto de la Pension por Beneficiario
    vlExistenPensiones = False
    vlCalculoPrimerPago = False
    
    If Format(Lbl_FecTopePriPag, "yyyymmdd") <> Format(Txt_FecPriPag, "yyyymmdd") Or vlMesDif = 0 Or lbl_TR = "6" Then 'Se modificó la Fecha de Primer Pago
        If Not fgCalcularPrimerPago(Txt_Poliza, Format(Txt_FecPriPag, "YYYYMMDD"), vlFecDev, "30001231", Lbl_MtoPenDefUf, CDbl(vlMtoPenGarUf), Lbl_Moneda(ciMonedaPensionNew), vlExistenPensiones, vlCodTipReajuste, vlMtoValReajusteTri, vlMtoValReajusteMen, Format(Lbl_FecVigPoliza, "yyyymmdd"), lbl_TR) Then 'HQR 12/01/2011 Pendiente ABV validar los dos campos agregados
            MsgBox "Error al Calcular Primeros Pagos", vbCritical
            Exit Sub
        End If
    End If
    vlCalculoPrimerPago = True
    '5 Muestra datos en la grilla
    If vlExistenPensiones Then
        Call flDespliegaPrimerPago
    Else
        MSF_Detalle.rows = 1 'Detalle por Beneficiario
        MSF_Resumen.rows = 1 'Resumen por Beneficiario
        MSF_Detalle.rows = 2 'Detalle por Beneficiario
        MSF_Resumen.rows = 2 'Resumen por Beneficiario
        MsgBox "No se genera Primer Pago de Pensiones", vbInformation
    End If
    
    Cmd_Grabar.SetFocus
    
    Screen.MousePointer = 0
    
Exit Sub
Err_Cal:
    
    Lbl_FecVigPoliza = ""
    
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Grabar_Click()
Dim vlOp As String
Dim vlResp As Long
On Error GoTo Err_Grabar
    
    'Valida el ingreso del Número de Póliza
    Txt_Poliza = Format(Trim(Txt_Poliza), "0000000000")
    If (Txt_Poliza = "") Then
        MsgBox "Debe seleccionar la Póliza a registrar la Prima.", vbCritical, "Error de Datos"
        If Fra_Poliza.Enabled = True Then Txt_Poliza.SetFocus
        If Fra_Poliza.Enabled = False Then Cmd_Salir.SetFocus
        Exit Sub
    End If
    
    'Valida el Ingreso de la Póliza
    If Fra_Poliza.Enabled = False Then
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
    If Lbl_MtoPrimaRec = "" Then
       MsgBox "Debe ingresar el Monto de la Prima Recibida o Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    If CDbl(Lbl_MtoPrimaRec) = 0 Then
       MsgBox "Debe Ingresar un Valor Mayor que 0 para Monto Prima Recibida o Recaudada.", vbCritical, "Error de Datos"
       Cmd_Salir.SetFocus
       Exit Sub
    End If
    
    'Valida el ingreso de la Fecha de Primer Pago
    If Txt_FecPriPag = "" Then
       MsgBox "Debe ingresar la Fecha del Primer Pago.", vbCritical, "Error de Datos"
       Txt_FecPriPag.SetFocus
       Exit Sub
    End If

    If Not IsDate(Txt_FecPriPag.Text) Then
       MsgBox "La Fecha de Primer Pago No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_FecPriPag.SetFocus
       Exit Sub
    End If
    
    If vlMesDif = 0 Then 'Si no es Diferida, se valida la fecha
        If (CDate(Txt_FecPriPag) < CDate(Date)) Then
           MsgBox "La Fecha Ingresada es Menor a la Fecha Actual", vbCritical, "Error de Datos"
           Txt_FecPriPag.SetFocus
           Exit Sub
        End If
    End If
    
    If (Year(Txt_FecPriPag) < 1900) Then
       MsgBox "La Fecha de Primer Pago es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_FecPriPag.SetFocus
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
    
    'Calcular el Primer Pago
    If Not (vlCalculoPrimerPago) Then
       MsgBox "Debe realizar Cálculo del Monto del Primer Pago", vbCritical, "Error de Cálculo"
       Cmd_Calcular.SetFocus
       Exit Sub
    End If

    Screen.MousePointer = 11
    
    Txt_Poliza = Trim(Txt_Poliza)
    vlFTraspaso = Format(CDate(Trim(Txt_FecTraspaso)), "yyyymmdd")
    vlFVigencia = Format(CDate(Trim(Lbl_FecVigPoliza)), "yyyymmdd")
    vlOp = ""
    
   'Verifica la existencia de la póliza
    vgSql = ""
    vgSql = "SELECT num_poliza FROM pd_tmae_polprirec WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_Poliza) & "' "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If vlRegistro.EOF Then
        vlOp = "I"
    End If
    vlRegistro.Close
    
    Dim strTipoPension As String
    strTipoPension = Mid(Lbl_TipoPension.Caption, 1, 2)
    Dim strTipoRenta As String
    strTipoRenta = Mid(Lbl_TipoRenta.Caption, 1, 1)
    'mvg
    If (strTipoPension = "04" Or strTipoPension = "05") And strTipoRenta = "6" Then
        vlResp = MsgBox(" ¿ Este póliza cuenta con exoneración del Descuento Essalud ?", 4 + 32 + 256, "Proceso de Ingreso de Datos")
        If vlResp <> 6 Then
            strID = "N"
        Else
            strID = "S"
        End If
    Else
        strID = "N"
    End If
    
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

Private Sub Cmd_Imprimir_Click(index As Integer)
Dim vlCodPar As String
Dim vlCodDerPen As String
Dim vlArchivo As String
Dim vlNombreSucursal As String, vlNombreTipoPension As String
Dim rs As ADODB.Recordset
Dim objRep As New ClsReporte
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
      
    If Not fgExiste(vlArchivo) Then     ', vbNormal
        MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
'    If vlRegistro.EOF Then
'       vlRegistro.Close
'       MsgBox "No existen el Archivo Seleccionado en la BD.", vbInformation, "Inexistencia de Datos"
'       Screen.MousePointer = 0
'       Exit Sub
'    End If
'    vlRegistro.Close
    
    vlCodPar = cgCodParentescoCau ' "99" 'Causante
     
     vlArchivo = strRpt & "PD_Rpt_IngPrimaAfp.rpt"   '\Reportes
     If Not fgExiste(vlArchivo) Then     ', vbNormal
         MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
         Screen.MousePointer = 0
         Exit Sub
     End If
     
    Set rs = New ADODB.Recordset
     
    vgQuery = "select d.gls_elemento as AFP, e.gls_elemento as Moneda, b.gls_nomben, b.gls_nomsegben, b.gls_patben, b.gls_matben,"
    vgQuery = vgQuery & " Cod_TipPension , c.mto_priafp, c.mto_pensionafp, c.mto_pricia, c.mto_sumpensioninf, c.Mto_SumPension, c.mto_sumpensionafp, c.mto_prirec, c.mto_pension, e.COD_SCOMP"
    vgQuery = vgQuery & " from pd_tmae_oripoliza a"
    vgQuery = vgQuery & " join pd_tmae_oripolben b on a.num_poliza=b.num_poliza"
    vgQuery = vgQuery & " join pd_tmae_polprirecaux c on a.num_poliza=c.num_poliza"
    vgQuery = vgQuery & " join ma_tpar_tabcod d on a.cod_afp=d.cod_elemento and d.cod_tabla='AF'"
    vgQuery = vgQuery & " join ma_tpar_tabcod e on a.cod_moneda=e.cod_elemento and e.cod_tabla='TM'"
    vgQuery = vgQuery & " where c.NUM_POLIZA = '" & Txt_Poliza & "' AND b.COD_PAR = '" & Trim(vlCodPar) & "'"
    rs.Open vgQuery, vgConexionBD, adOpenForwardOnly, adLockReadOnly
        
    If rs.EOF Then
        MsgBox "No hay datos para mostrar", vbExclamation, "Pago de pensiones"
        Exit Sub
    End If
    Dim vlNombreComuna As String, vlJefeBeneficios As String, vlDireccion As String
    Call flObtieneDatosContacto(vlafp, vlNombreComuna, vlJefeBeneficios, vlDireccion)
    
    'Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PD_Rpt_IngPrimaAfp.rpt"), ".RPT", ".TTX"), 1)
                
    If objRep.CargaReporte(strRpt & "", "PD_Rpt_IngPrimaAfp.rpt", "Recepción de Prima.", rs, True, _
                                    ArrFormulas("Nombre", vgNombreApoderado), _
                                    ArrFormulas("Cargo", vgCargoApoderado), _
                                    ArrFormulas("Distrito", vlNombreComuna), _
                                    ArrFormulas("NombreContacto", vlJefeBeneficios), _
                                    ArrFormulas("Direccion", vlDireccion)) = False Then
                                    
            MsgBox "No se pudo abrir el reporte", vbInformation
            'Exit Sub
    End If
   
'    vgQuery = "{PD_TMAE_POLPRIRECAUX.NUM_POLIZA} = '" & Txt_Poliza & "'"
'    vgQuery = vgQuery & " AND {PD_TMAE_ORIPOLBEN.COD_PAR} = '" & Trim(vlCodPar) & "'" 'Solo los Causantes
   

'    Rpt_Reporte.Reset
'    Rpt_Reporte.ReportFileName = vlArchivo     'App.Path & "\rpt_Areas.rpt"
'    Rpt_Reporte.Connect = vgRutaDataBase
'    Rpt_Reporte.Formulas(0) = "Nombre= '" & vgNombreApoderado & "'"
'    Rpt_Reporte.Formulas(1) = "Cargo= '" & vgCargoApoderado & "'"
'    Rpt_Reporte.Formulas(2) = "Distrito = '" & vlNombreComuna & "'"
'    Rpt_Reporte.Formulas(3) = "NombreContacto = '" & vlJefeBeneficios & "'"
'    Rpt_Reporte.Formulas(4) = "Direccion = '" & vlDireccion & "'"
'    Rpt_Reporte.WindowTitle = "Recepción de Prima."
'    Rpt_Reporte.SelectionFormula = vgQuery
'    Rpt_Reporte.Destination = crptToWindow
'    Rpt_Reporte.WindowState = crptMaximized
'    Rpt_Reporte.Action = 1
'
'    Screen.MousePointer = 0
    
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar

    Lbl_NumIdent = ""
    Lbl_TipoIdent = ""
    Lbl_Cuspp = ""
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
    Lbl_MtoPrimaRec = ""
    Lbl_FecVigPoliza = ""
    Lbl_FactVarRta = ""
    Lbl_FecSolicitud = ""
    Lbl_MtoPenDefUf = ""
    Lbl_MtoPenDefUfAFP = ""
    Lbl_PrimaDefCia = ""
    Lbl_PrimaDefAFP = ""
    Lbl_TipoCambio = ""
    
    Txt_Poliza = ""
    Lbl_FecTopePriPag = ""
    Txt_FecPriPag = ""
    
    Fra_AntRec.Enabled = False
    Fra_Dif.Enabled = False
    Fra_Definitiva.Enabled = False
    
    'SSTab1.Enabled = False
    Fra_Poliza.Enabled = True
    Fra_lista.Enabled = True
    'SSTab1.Tab = 0
    Call flInicializaGrillaVarFondo
    Call flInicializaGrillaVarTipoCambio
    Call flInicializaGrillaVarFondoTC
    MSF_Resumen.rows = 1
    MSF_Detalle.rows = 1
    MSF_Resumen.rows = 2
    MSF_Detalle.rows = 2
    
'I--- ABV 07/11/2007 ---
    Lbl_SumPenInf = ""
    Lbl_SumPenDef = ""
    Lbl_SumPenDefAFP = ""
    Lbl_SumPenInf.Visible = False
    Lbl_SumPenDef.Visible = False
    Lbl_SumPenDefAFP.Visible = False
'F--- ABV 07/11/2007 ---
    
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

    SSTab2.Tab = 0
    
    Frm_CalPrima.Top = 0
    Frm_CalPrima.Left = 0
    
    Call flInicializaGrillaVarFondo
    
    Call flInicializaGrillaVarTipoCambio
    
    Call flInicializaGrillaVarFondoTC
    
    'Actualizar los Datos en la Lista
    flCargaPolizas
    Lbl_Moneda(ciMonedaPrimaAnt) = vgMonedaCodOfi  'cgCodTipMonedaUF
    Lbl_Moneda(ciMonedaPrimaDef) = vgMonedaCodOfi
    Lbl_Moneda(ciMonedaPrimaNew) = vgMonedaCodOfi  'cgCodTipMonedaUF
    Lbl_Moneda(ciMonedaPrimaDefAFP) = vgMonedaCodOfi
    Lbl_Moneda(ciMonedaPensionDefAFP) = vgMonedaCodOfi
    
    '''SSTab1.Enabled = False
    Fra_AntRec.Enabled = False
    Fra_Dif.Enabled = False
    Fra_Definitiva.Enabled = False
  
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

Private Sub Txt_FecPriPag_Change()
    Lbl_FecVigPoliza = ""
End Sub

Private Sub Txt_FecPriPag_GotFocus()
    Txt_FecPriPag.SelStart = 0
    Txt_FecPriPag.SelLength = Len(Txt_FecPriPag)
End Sub

Private Sub Txt_FecPriPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And (Trim(Txt_FecPriPag <> "")) Then
       If Not (Txt_FecPriPag = "0") Then
          vlPasa = True
          'If flValidaFecha(Txt_FecPriPag) = True Then
            If Not IsDate(Txt_FecPriPag.Text) Then
               MsgBox "La Fecha de Primer Pago No es una Fecha Válida.", vbCritical, "Error de Datos"
               Txt_FecPriPag.SetFocus
               Exit Sub
            End If

            If vlMesDif = 0 Then
                If (CDate(Txt_FecPriPag) < CDate(Date)) Then
                   MsgBox "La Fecha de Primer Pago Ingresada es Menor a la Fecha Actual", vbCritical, "Error de Datos"
                   Txt_FecPriPag.SetFocus
                   Exit Sub
                End If
                If CDate(Txt_FecPriPag) > CDate(DateAdd("d", clNumeroMaximoDiasPrimerPago, Date)) Then
                    MsgBox "Fecha de Primer Pago es Mayor a la Fecha Actual por más de '" & clNumeroMaximoDiasPrimerPago & "' días", vbCritical, "Error de Datos"
                    Txt_FecPriPag.SetFocus
                    Exit Sub
                End If
            Else
                If Trim(CDate(Lbl_FecTopePriPag)) <> Trim(CDate(Txt_FecPriPag)) Then 'Se modificó la Fecha de Primer Pago
                    If lbl_TR <> "6" Then
                         If (Format(Txt_FecPriPag, "yyyymm") < Format(Date, "yyyymm")) Then
                           MsgBox "La Fecha de Primer Pago Ingresada no debe ser Menor al Mes Actual", vbCritical, "Error de Datos"
                           Txt_FecPriPag.SetFocus
                           Exit Sub
                        End If
                        If Format(Lbl_FecTopePriPag, "yyyymmdd") > Format(Txt_FecPriPag, "yyyymmdd") Then 'Se modificó la Fecha de Primer Pago
                           MsgBox "La Fecha de Tope del Primer Pago definida no debe ser Mayor a la Fecha de Primer Pago (o Inicio del Pago Diferido)", vbCritical, "Error de Datos"
                           Txt_FecPriPag.SetFocus
                           Exit Sub
                        End If
                    End If

                    If CDate(Txt_FecPriPag) > CDate(DateAdd("d", clNumeroMaximoDiasPrimerPago, Date)) Then
                        MsgBox "Fecha de Primer Pago es Mayor a la Fecha Actual por más de '" & clNumeroMaximoDiasPrimerPago & "' días", vbCritical, "Error de Datos"
                        Txt_FecPriPag.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            If (Year(Txt_FecPriPag) < 1900) Then
               MsgBox "La Fecha de Primer Pago es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
               Txt_FecPriPag.SetFocus
               Exit Sub
            End If
             Cmd_Calcular.SetFocus
          'End If
       End If
    End If
End Sub


Private Sub Txt_FecPriPag_LostFocus()
    If (Trim(Txt_FecPriPag) = "") Then
       Txt_FecPriPag = ""
       Exit Sub
    End If
    If Not IsDate(Txt_FecPriPag.Text) Then
       Txt_FecPriPag = ""
       Exit Sub
    End If
    If vlMesDif = 0 Then
        If (CDate(Txt_FecPriPag) < CDate(Date)) Then
           Txt_FecPriPag = ""
           Exit Sub
        End If
        If CDate(Txt_FecPriPag) > CDate(DateAdd("d", clNumeroMaximoDiasPrimerPago, Date)) Then
            Txt_FecPriPag = ""
            Exit Sub
        End If
    Else
        If Trim(CDate(Lbl_FecTopePriPag)) <> Trim(CDate(Txt_FecPriPag)) Then 'Se modificó la Fecha de Primer Pago
            If lbl_TR <> "6" Then
                If (CDate(Txt_FecPriPag) < CDate(Date)) Then
                   Txt_FecPriPag = ""
                   Exit Sub
                End If
                If Format(Lbl_FecTopePriPag, "yyyymmdd") > Format(Txt_FecPriPag, "yyyymmdd") Then 'Se modificó la Fecha de Primer Pago
                    Txt_FecPriPag = ""
                    Exit Sub
                End If
            End If
            
            
            If CDate(Txt_FecPriPag) > CDate(DateAdd("d", clNumeroMaximoDiasPrimerPago, Date)) Then
                Txt_FecPriPag = ""
                Exit Sub
            End If
        Else
            If (Format(Txt_FecPriPag, "yyyymm") < Format(Date, "yyyymm")) Then
                Txt_FecPriPag = ""
                Exit Sub
            End If
        End If
    End If
    If (Year(Txt_FecPriPag) < 1900) Then
       Txt_FecPriPag = ""
       Exit Sub
    End If
    Txt_FecPriPag.Text = Format(CDate(Trim(Txt_FecPriPag)), "yyyymmdd")
    Txt_FecPriPag.Text = DateSerial(Mid((Txt_FecPriPag.Text), 1, 4), Mid((Txt_FecPriPag.Text), 5, 2), Mid((Txt_FecPriPag.Text), 7, 2))
End Sub

Private Sub Txt_FecTraspaso_Change()
    Lbl_FecVigPoliza = ""
End Sub

Private Sub Txt_FecTraspaso_GotFocus()

    Txt_FecTraspaso.SelStart = 0
    Txt_FecTraspaso.SelLength = Len(Txt_FecTraspaso)
    
End Sub

Private Sub Txt_FecTraspaso_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (Trim(Txt_FecTraspaso) <> "") Then
       If Not (Txt_FecTraspaso = "0") Then
          vlPasa = True
          If flValidaFecha(Txt_FecTraspaso) = True Then
             Cmd_Calcular.SetFocus
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
'I--- ABV 20/09/20007 ---
    'Lbl_FecTopePriPag = fgCalcularFechaPrimerPagoEst(vlFecAcepta, vlFecDev, Txt_FecTraspaso, vlTipoPen, vlTipoRen, vlMesDif, vlNumDias)
    Lbl_FecTopePriPag = fgCalcularFechaPrimerPagoEst(vlFecAcepta, vlFecDev, vlFecTraspasoPrimas, vlTipoPen, vlTipoRen, vlMesDif, vlNumDias)
'F--- ABV 20/09/20007 ---
    If Trim(Txt_FecPriPag) = "" Then
       Txt_FecPriPag = Lbl_FecTopePriPag 'Fecha Dada por Defecto
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
    vgSql = "SELECT DISTINCT num_poliza FROM pd_tmae_polprirecaux " 'pd_tmae_oripoliza "
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
Dim vlRegPrima As ADODB.Recordset
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
    vgSql = vgSql & ",p.mto_priunidif,p.prc_rentatmp,p.cod_afp,be.cod_par,"
'I--- ABV 07/11/2007 ---
    vgSql = vgSql & "p.mto_sumpension "
'F--- ABV 07/11/2007 ---
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & ",p.cod_tipreajuste,p.mto_valreajustetri,p.mto_valreajustemen,p.fec_vigencia, " 'hqr 13/01/2011
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
        Fra_Poliza.Enabled = False
        Fra_AntRec.Enabled = True
        Fra_Dif.Enabled = True
        Fra_Definitiva.Enabled = True
        Fra_lista.Enabled = False
            
        Lbl_TipoPension = Trim(vlRegistro!Cod_TipPension) + " - " + Trim(vlRegistro!Gls_Pension)
        Lbl_TipoRenta = Trim(vlRegistro!Cod_TipRen) + " - " + Trim(vlRegistro!Gls_Renta)
        lbl_TR = Trim(vlRegistro!Cod_TipRen)
        vlCodTp = Trim(vlRegistro!Cod_TipPension) 'RRR 22012019
        Lbl_Modalidad = Trim(vlRegistro!Cod_Modalidad) + " - " + Trim(vlRegistro!Gls_Modalidad)
        Lbl_NumIdent = Trim(vlRegistro!num_idenafi)
        Lbl_TipoIdent = Trim(vlRegistro!GLS_TIPOIDENCOR)
        Lbl_NomAfiliado = Trim(vlRegistro!Gls_NomBen) + " " + Trim(vlRegistro!Gls_PatBen) + " " + IIf(IsNull(vlRegistro!Gls_MatBen), "", Trim(vlRegistro!Gls_MatBen))
        Lbl_Cuspp = Trim(vlRegistro!Cod_Cuspp)
        vlCodPar = Trim(vlRegistro!Cod_Par)
        vlafp = Trim(vlRegistro!Cod_AFP)
        
        vlNomBen = Trim(vlRegistro!Gls_NomBen)
        vlApeBenPat = Trim(vlRegistro!Gls_PatBen)
        vlApeBenMat = IIf(IsNull(vlRegistro!Gls_MatBen), "", Trim(vlRegistro!Gls_MatBen))
        
        
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
        Lbl_PrimaInf = Format((vlRegistro!MTO_PRIUNI), "#,#0.00")
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
        If Lbl_Meses = "0" Then
            vlMtoPenGar = 0
        Else
            vlMtoPenGar = Format((vlRegistro!Mto_PensionGar), "#,#0.00")
        End If
        
        
'I--- ABV 07/11/2007 ---
        Lbl_SumPenInf = Format((vlRegistro!Mto_SumPension), "#,#0.00")
        If (vlRegistro!Cod_TipPension = clCodTipPensionSob) Then
            Lbl_SumPenInf.Visible = True
            Lbl_SumPenDef.Visible = True
            Lbl_SumPenDefAFP.Visible = True
            
            Lbl_PensionInf.Visible = False
            Lbl_MtoPenDefUf.Visible = False
            Lbl_MtoPenDefUfAFP.Visible = False
        
            Lbl_SumPenInf.Top = 2550 '2340
            Lbl_SumPenInf.Left = 6960
            Lbl_SumPenDef.Top = 540
            Lbl_SumPenDef.Left = 2040
            Lbl_SumPenDefAFP.Top = 1380
            Lbl_SumPenDefAFP.Left = 2040
        Else
            Lbl_SumPenInf.Visible = False
            Lbl_SumPenDef.Visible = False
            Lbl_SumPenDefAFP.Visible = False
            
            Lbl_PensionInf.Visible = True
            Lbl_MtoPenDefUf.Visible = True
            Lbl_MtoPenDefUfAFP.Visible = True
        End If
'F--- ABV 07/11/2007 ---
        
        vlMonedaPension = Trim(vlRegistro!Cod_Moneda)
        Lbl_Moneda(ciMonedaPensionAnt) = vlMonedaPension
        Lbl_Moneda(ciMonedaPensionNew) = Lbl_Moneda(ciMonedaPensionAnt)
        
        vlFecAcepta = vlRegistro!Fec_Acepta
        vlFecDev = vlRegistro!Fec_Dev
        vlFecVigencia = vlRegistro!Fec_Vigencia '15/02/2011
        vlTipoPen = vlRegistro!Cod_TipPension
        vlTipoRen = vlRegistro!Cod_TipRen
        vlMesDif = vlRegistro!Num_MesDif
''        If vlMesDif > 0 Then 'Diferida
''            Txt_FecPriPag.Enabled = False
''        Else
''            Txt_FecPriPag.Enabled = True
''        End If
        vlNumDias = fgObtieneDiasPrimerPagoEst(vlTipoPen)
        Lbl_FecTopePriPag = fgCalcularFechaPrimerPagoEst(vlFecAcepta, vlFecDev, Txt_FecTraspaso, vlTipoPen, vlTipoRen, vlMesDif, vlNumDias)
        If Trim(Txt_FecPriPag) = "" Then
           Txt_FecPriPag = Lbl_FecTopePriPag 'Fecha Dada por Defecto
        End If
        vlTipoCambioAcep = vlRegistro!Mto_ValMoneda
'        Lbl_PorcRentaTmp = Format(vlRegistro!Prc_RentaTMP, "#,#0.00")
'        Lbl_PriUniDif = Format(vlRegistro!Mto_PriUniDif, "#,#0.00")

'I--- ABV 15/10/2007 ---
        'Obtener los datos del Registro de la Solicitud de Primas
        vgSql = "SELECT Mto_Pension,Mto_PriCia,Mto_PensionAFP,Mto_PriAFP,Fec_Traspaso,"
        vgSql = vgSql & "Prc_FacVar,Mto_ValMonedaRec,Mto_PriRec "
'I--- ABV 07/11/2007 ---
        vgSql = vgSql & ",mto_sumpensioninf,mto_sumpension,mto_sumpensionafp "
'F--- ABV 07/11/2007 ---
        vgSql = vgSql & "FROM "
        vgSql = vgSql & "pd_tmae_polprirecaux p "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "p.num_poliza = '" & Trim(Txt_Poliza) & "' "
        Set vlRegPrima = vgConexionBD.Execute(vgSql)
        If Not (vlRegPrima.EOF) Then
            Lbl_MtoPenDefUf = Format(vlRegPrima!Mto_Pension, "#,#0.00")
            Lbl_PrimaDefCia = Format(vlRegPrima!mto_pricia, "#,#0.00")
            Lbl_MtoPenDefUfAFP = Format(vlRegPrima!mto_pensionafp, "#,#0.00")
            Lbl_PrimaDefAFP = Format(vlRegPrima!mto_priafp, "#,#0.00")
            Lbl_FecSolicitud = DateSerial(Mid(vlRegPrima!fec_traspaso, 1, 4), Mid(vlRegPrima!fec_traspaso, 5, 2), Mid(vlRegPrima!fec_traspaso, 7, 2))
            Lbl_FactVarRta = Format(vlRegPrima!PRC_FACVAR, "#,#0.00000000")
            Lbl_TipoCambio = Format(vlRegPrima!Mto_ValMonedaRec, "#,#0.000")
            Lbl_MtoPrimaRec = Format(vlRegPrima!MTO_PRIREC, "#,#0.00")
'I--- ABV 07/11/2007 ---
            Lbl_SumPenDef = Format(vlRegPrima!Mto_SumPension, "#,#0.00")
            Lbl_SumPenDefAFP = Format(vlRegPrima!Mto_SumPensionAfp, "#,#0.00")
'F--- ABV 07/11/2007 ---
        Else
            Lbl_MtoPenDefUf = ""
            Lbl_PrimaDefCia = ""
            Lbl_MtoPenDefUfAFP = ""
            Lbl_PrimaDefAFP = ""
            Lbl_FecSolicitud = ""
            Lbl_FactVarRta = ""
            Lbl_TipoCambio = ""
            Lbl_MtoPrimaRec = ""
'I--- ABV 07/11/2007 ---
            Lbl_SumPenDef = ""
            Lbl_SumPenDefAFP = ""
'F--- ABV 07/11/2007 ---
        End If
        vlRegPrima.Close
'F--- ABV 15/10/2007 ---

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
'El Número de Liquidación es el mismo de la Boleta o Factura de Pago
    If Not flObtieneNumLiquidacion(vgConectarBD, vlNumLiquidacion, vlNumRenVit) Then
        MsgBox "Error al Obtener Número de Liquidación.", vbCritical, "Error de Datos"
        Exit Function
    End If
    
'    If Not flObtieneNumFactura(vgConectarBD, vlNumFactura, vlNumRenVit) Then
'        MsgBox "Error al Obtener Número de Factura.", vbCritical, "Error de Datos"
'        Exit Function
'    End If
    vlNumFactura = vlNumLiquidacion
'F--- ABV 14/10/2007 ---

    'Inserta los Datos en la Tabla de pd_TMAE_POLPRIREC
    Sql = "INSERT INTO pd_tmae_polprirec ("
    Sql = Sql & "num_poliza,fec_traspaso,fec_vigencia,mto_priinf,"
    Sql = Sql & "mto_pensioninf,mto_pensiongarinf,mto_prirecpesos,"
    Sql = Sql & "mto_prirec,prc_facvar,mto_pension,mto_pensiongar,"
    Sql = Sql & "cod_usuariocrea,fec_crea,hor_crea,"
    Sql = Sql & "cod_monedapriinf,cod_monedapeninf,mto_valmonedarec,mto_valmonedainf,"
    Sql = Sql & "mto_penvarfon,mto_penvartc,mto_penvarfontc,cod_liquidacion"
    Sql = Sql & ",cod_factura,cod_renvit "
    Sql = Sql & ",mto_priafp,mto_pensionafp,mto_pricia "
    Sql = Sql & ",fec_solafp,mto_pritotal "
'I--- ABV 07/11/2007 ---
    Sql = Sql & ",mto_sumpensioninf,mto_sumpension,mto_sumpensionafp "
'F--- ABV 07/11/2007 ---
    Sql = Sql & ") VALUES ("
    Sql = Sql & "'" & Trim(Txt_Poliza) & "',"
    Sql = Sql & "'" & Trim(vlFTraspaso) & "',"
    Sql = Sql & "'" & Trim(vlFVigencia) & "',"
    Sql = Sql & " " & str(Lbl_PrimaInf) & ","
    Sql = Sql & " " & str(Lbl_PensionInf) & ","
    Sql = Sql & " " & str(vlMtoPenGar) & ","
    Sql = Sql & " " & str(Lbl_PrimaDefCia) & ","
    Sql = Sql & " " & str(Lbl_PrimaDefCia) & ","
    Sql = Sql & " " & str(Lbl_FactVarRta) & ","
    Sql = Sql & " " & str(Lbl_MtoPenDefUf) & ","
    If (vlMtoPenGarUf) = "" Then
        Sql = Sql & " " & str(Format("0", "#0.00")) & ","
    Else
        Sql = Sql & " " & str(vlMtoPenGarUf) & ","
    End If
    Sql = Sql & "'" & (vgUsuario) & "',"
    Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
    Sql = Sql & "'" & Format(Time, "hhmmss") & "',"
    Sql = Sql & "'" & Lbl_Moneda(ciMonedaPrimaAnt) & "',"
    Sql = Sql & "'" & Lbl_Moneda(ciMonedaPensionAnt) & "',"
    Sql = Sql & " " & str(vlTipoCambioRec) & ","
    Sql = Sql & " " & str(vlTipoCambioAcep) & ","
    Sql = Sql & " " & str(vlMtoVarFon) & ","
    Sql = Sql & " " & str(vlMtoVarTC) & ","
    Sql = Sql & " " & str(vlMtoVarFonTC) & ","
    Sql = Sql & " " & str(vlNumLiquidacion)
    Sql = Sql & ",'" & str(vlNumFactura) & "'"
    Sql = Sql & ",'" & (vlNumRenVit) & "'"
    Sql = Sql & "," & str(Lbl_PrimaDefAFP) & " "
    Sql = Sql & "," & str(Lbl_MtoPenDefUfAFP) & " "
    Sql = Sql & "," & str(Lbl_PrimaDefCia) & " "
    Sql = Sql & ",'" & Format(Lbl_FecSolicitud, "yyyymmdd") & "',"
    Sql = Sql & " " & str(Lbl_MtoPrimaRec) & " "
'I--- ABV 07/11/2007 ---
    Sql = Sql & "," & str(Lbl_SumPenInf) & " "
    Sql = Sql & "," & str(Lbl_SumPenDef) & " "
    Sql = Sql & "," & str(Lbl_SumPenDefAFP) & " "
'F--- ABV 07/11/2007 ---
    Sql = Sql & ")"
    vgConectarBD.Execute (Sql)
    
    'Eliminar registro de Tabla de Recepción de Primas Auxiliar
    Sql = "DELETE FROM pd_tmae_polprirecaux "
    Sql = Sql & "WHERE num_poliza = '" & Trim(Txt_Poliza) & "'"
    vgConectarBD.Execute (Sql)
    
    'Traspasar Información
    Call flTraspasarInformacion
    
    'Grabar el Primer Pago
    Call flGrabarPrimerPago
    
    'Crea el Usuario y Password para la pagina Web de Consultas. 20/11/2014
    Call flGrabarUsuarioPassword
    
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
    Return
End Function

Function flGrabarUsuarioPassword() As Boolean

Dim vlEstado As String
Dim vlPassword As String

vlEstado = "A"

'INSERTA EL USUARIO

Dim CodUser As String

CodUser = "R" & Mid(Txt_Poliza, 6, 5)


 vgSql = " insert into MA_TMAE_USUMATRIZ (cod_usuario, cod_tipoidenusu, nro_idenusu, gls_nombres, gls_apepat, gls_apemat,"
 vgSql = vgSql & " cod_usucrea,fec_crea,hor_crea, ccod_sucursal,isblock, cod_estado, cod_tipuser"
 vgSql = vgSql & ")"
 vgSql = vgSql & " values('" & CodUser & "',1 ,'" & Lbl_NumIdent & "', '" & UCase(vlNomBen) & "', '" & UCase(vlApeBenPat) & "', '" & UCase(vlApeBenMat) & "',"
 vgSql = vgSql & " '" & vgUsuario & "','" & Format(Date, "yyyymmdd") & "','" & Format(Time, "hhmmss") & "', '0001','0','" & vlEstado & "'"
 vgSql = vgSql & ", '2'"
 vgSql = vgSql & ")"
 Set vgRs = vgConexionBD.Execute(vgSql)
               
'INSERTA LA CONTRASEÑA

 vlPassword = fgEncPassword(Lbl_NumIdent)

 vgSql = " insert into MA_TMAE_USUPASSWORD(cod_usuario, nro_usupass, fec_inipass, fec_finpass, fec_antPass, gls_password,"
 vgSql = vgSql & " gls_passwordconf, ind_Segu, cod_usucrea,fec_crea,hor_crea)"
 vgSql = vgSql & " values('" & CodUser & "','1','20141201', '99991231', '99991231', '" & vlPassword & "',"
 vgSql = vgSql & " '" & vlPassword & "', '0', '" & vgUsuario & "' ,'" & Format(Date, "yyyymmdd") & "','" & Format(Time, "hhmmss") & "')"
 vgConexionBD.Execute (vgSql)

'INSERTA EL MODULO DE ACCESOS

 vgSql = " insert into MA_TMAE_USUMODULO(cod_usuario, cod_sistema, cod_nivel, cod_usucrea, fec_crea, hor_crea)"
 vgSql = vgSql & " values('" & CodUser & "', 'PW', '23', '" & vgUsuario & "','" & Format(Date, "yyyymmdd") & "','" & Format(Time, "hhmmss") & "')"
 vgConexionBD.Execute (vgSql)

End Function


Function flTraspasarInformacion()
    Dim Sql2 As String
    Dim mtoPensiongar, nummesgar As Double
    Dim codTipopen, codModal As String
    
    vgSql = ""
    vgSql = "SELECT * FROM pd_tmae_oripoliza WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_Poliza) & "'"
    Set vlRegistro1 = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro1.EOF) Then
        Call flGrabarPolizaRec
        codTipopen = vlRegistro1!Cod_TipPension
        codModal = vlRegistro1!Cod_Modalidad
        'RRR CAMBIO PARA ERROR EN POLIZAS DE SOBREVIVENCIA 20181109
        mtoPensiongar = vlMtoPenGarUf 'vlRegistro1!Mto_Pension '
        If codTipopen = "08" Then
            mtoPensiongar = vlMtoPenGarUf ' vlRegistro1!Mto_PensionGar
        End If
        
        '''
        nummesgar = vlRegistro1!Num_MesGar
        
    End If
    vlRegistro1.Close
    
    'RVF 20090914
    vgSql = ""
    vgSql = "SELECT * FROM pd_tmae_oripolrep WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_Poliza) & "'"
    Set vlRegistro4 = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro4.EOF) Then
        Call pGrabarPolizaRep
    End If
    vlRegistro4.Close
    
''    vgSql = ""
''    vgSql = "SELECT 1 FROM pd_tmae_oripolbon WHERE "
''    vgSql = vgSql & "num_poliza = '" & Trim(Txt_Poliza) & "'"
''    Set vlRegistro2 = vgConexionBD.Execute(vgSql)
''    If Not (vlRegistro2.EOF) Then
''        Call flGrabarBonosRec
''    End If
''    vlRegistro2.Close
    
    Dim sumaPor As Double
    Dim prcPension, mtoPension As Double
    
    sumaPor = 0
    prcPension = 0
    mtoPension = 0
    
    vgSql = ""
    vgSql = "SELECT * FROM pd_tmae_oripolben WHERE "
    vgSql = vgSql & "num_poliza = '" & Trim(Txt_Poliza) & "'"
    Set vlRegistro3 = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro3.EOF) Then
        While Not vlRegistro3.EOF
              Call flGrabarBeneficiariosRec
              sumaPor = sumaPor + vlRegistro3!Prc_Pension
              vlRegistro3.MoveNext
        Wend
    End If
    vlRegistro3.Close
    'RRR
    'Actualiza los % de pension garantizado
    If nummesgar > 0 Then
        vgSql = ""
        vgSql = "SELECT * FROM pd_tmae_oripolben WHERE "
        vgSql = vgSql & "num_poliza = '" & Trim(Txt_Poliza) & "'"
        Set vlRegistro3 = vgConexionBD.Execute(vgSql)
        If Not (vlRegistro3.EOF) Then
            While Not vlRegistro3.EOF
             If codTipopen <> "08" Then
                 If vlRegistro3!Num_Orden <> "1" Then
                     prcPension = 0
                     mtoPension = 0
                 Else
                     prcPension = 100
                     mtoPension = mtoPensiongar
                 End If
                 Sql2 = ""
                 Sql2 = Sql2 & " update pd_tmae_polben set prc_pensiongar=" & prcPension & ", mto_pensiongar=" & mtoPension & ""
                 Sql2 = Sql2 & " where num_poliza='" & vlRegistro3!Num_Poliza & "' and num_orden=" & vlRegistro3!Num_Orden & ""
                 vgConectarBD.Execute (Sql2)
             Else
                  prcPension = Format((CDbl(vlRegistro3!Prc_Pension) / sumaPor) * 100, "###0.00")
                  mtoPension = Format(mtoPensiongar * (((CDbl(vlRegistro3!Prc_Pension) / sumaPor) * 100) / 100), "###0.00")
                  Sql2 = ""
                  Sql2 = Sql2 & " update pd_tmae_polben set prc_pensiongar=" & prcPension & ", mto_pensiongar=" & mtoPension & ""
                  Sql2 = Sql2 & " where num_poliza='" & vlRegistro3!Num_Poliza & "' and num_orden=" & vlRegistro3!Num_Orden & ""
                  vgConectarBD.Execute (Sql2)
              End If
            vlRegistro3.MoveNext
            Wend
        End If
        vlRegistro3.Close
    End If
    
    'transpasa pd_tmae_oritutor a pd_tmae_poltutor
    Dim vlFecModi, vlHorModi As String
    vlFecModi = Format(ObtenerFechaServer, "yyyymmdd")
    vlHorModi = Format(Time, "hhmmss")
    
    Sql2 = ""
    Sql2 = Sql2 & "Insert into PD_TMAE_poltutor(NUM_POLIZA, NUM_ENDOSO,NUM_ORDEN, NUM_IDENTUT, COD_TIPOIDENTUT, GLS_NOMTUT, GLS_NOMSEGTUT,"
    Sql2 = Sql2 & " GLS_PATTUT, GLS_MATTUT, GLS_DIRTUT, COD_DIRECCION, GLS_FONOTUT, NUM_MESPODNOT, FEC_INIPODNOT, FEC_TERPODNOT,"
    Sql2 = Sql2 & " COD_VIAPAGO, COD_TIPCUENTA, COD_BANCO, NUM_CUENTA, COD_SUCURSAL, FEC_EFECTO, FEC_RECCIA, COD_USUARIOCREA, FEC_CREA,"
    Sql2 = Sql2 & " HOR_CREA, COD_USUARIOMODI, FEC_MODI, HOR_MODI,GLS_CORREOTUT)"
    Sql2 = Sql2 & " select NUM_POLIZA," & vlEndoso & ", NUM_ORDEN, NUM_IDENTUT, COD_TIPOIDENTUT, GLS_NOMTUT, GLS_NOMSEGTUT, GLS_PATTUT, GLS_MATTUT,"
    Sql2 = Sql2 & " GLS_DIRTUT, COD_DIRECCION, GLS_FONOTUT, NUM_MESPODNOT, FEC_INIPODNOT, FEC_TERPODNOT, COD_VIAPAGO, COD_TIPCUENTA,"
    Sql2 = Sql2 & " COD_BANCO, NUM_CUENTA, COD_SUCURSAL, FEC_EFECTO, FEC_RECCIA, '" & vgUsuario & "', '" & vlFecModi & "', '" & vlHorModi & "','" & vgUsuario & "',"
    Sql2 = Sql2 & " FEC_MODI, HOR_MODI, GLS_CORREOTUT from PD_TMAE_oritutor where NUM_POLIZA='" & Trim(Txt_Poliza) & "'"
    vgConectarBD.Execute (Sql2)
    'Eliminar Datos de Tablas Originales
'    Call flEliminarBonosOri
    Call flEliminarBeneficiariosOri
    'RVF 20090914
    'MARCO ELIMINA ORITUTOR
    vgQuery = ""
    vgQuery = "DELETE FROM PD_TMAE_oritutor WHERE "
    vgQuery = vgQuery & "num_poliza = '" & Trim(Txt_Poliza) & "'"
    vgConectarBD.Execute (vgQuery)
    
    Call pEliminarRepresentanteOri
    Call flEliminarPolizaOri
    
    'PD_TMAE_oritutor where NUM_POLIZA='" & Trim(Txt_Poliza) & "'
End Function

Private Sub pEliminarRepresentanteOri()
    vgQuery = ""
    vgQuery = "DELETE FROM pd_tmae_oripolrep WHERE "
    vgQuery = vgQuery & "num_poliza = '" & Trim(Txt_Poliza) & "'"
    vgConectarBD.Execute (vgQuery)

End Sub

Function flGrabarPrimerPago() As Boolean
    On Error GoTo Errores
    Dim vlMontoPension As Double
    Dim vlPoliza As String
    Dim vlOrden As String
    Dim vlNumIdenReceptor As String
    Dim vlCodTipoIdenReceptor As Long
    Dim vlTipReceptor As String
    Dim i As Long
    flGrabarPrimerPago = False
    
    'Limpia Variables
    vlMontoPension = 0
    vlPoliza = ""
    vlOrden = ""
    vlNumIdenReceptor = ""
    vlCodTipoIdenReceptor = -1
    vlTipReceptor = ""

    For i = 0 To UBound(stLiquidacion) - 1
        If vlPoliza <> stLiquidacion(i).Num_Poliza Or vlOrden <> stLiquidacion(i).Num_Orden Or vlNumIdenReceptor <> stLiquidacion(i).Num_IdenReceptor Or vlCodTipoIdenReceptor <> stLiquidacion(i).Cod_TipoIdenReceptor Or vlTipReceptor <> stLiquidacion(i).Cod_TipReceptor Then
            vlMontoPension = 0
            vlPoliza = stLiquidacion(i).Num_Poliza
            vlOrden = stLiquidacion(i).Num_Orden
            vlNumIdenReceptor = stLiquidacion(i).Num_IdenReceptor
            vlCodTipoIdenReceptor = stLiquidacion(i).Cod_TipoIdenReceptor
            vlTipReceptor = stLiquidacion(i).Cod_TipReceptor
        End If
        'Graba Encabezado Liquidación
        vlMontoPension = vlMontoPension + stLiquidacion(i).Mto_LiqPagar
        Sql = "INSERT INTO pd_tmae_liqpagopen "
        Sql = Sql & "(num_perpago, num_poliza, num_orden, cod_tipopago, "
        Sql = Sql & "gls_direccion, cod_direccion, cod_tippension, fec_pago, "
        Sql = Sql & "cod_viapago, cod_banco, cod_tipcuenta, num_cuenta, "
        Sql = Sql & "cod_sucursal, cod_inssalud, num_idenreceptor, cod_tipoidenreceptor, "
        Sql = Sql & "gls_nomreceptor, gls_nomsegreceptor, gls_patreceptor, gls_matreceptor, "
        Sql = Sql & "cod_tipreceptor, mto_haber, mto_descuento, mto_liqpagar, "
        Sql = Sql & "mto_baseimp, mto_basetri, mto_pension, "
        Sql = Sql & "mto_plansalud, cod_moneda, cod_modsalud, gls_montopension"
        Sql = Sql & ") VALUES ("
        Sql = Sql & "'" & stLiquidacion(i).Num_PerPago & "',"
        Sql = Sql & "'" & stLiquidacion(i).Num_Poliza & "',"
        Sql = Sql & stLiquidacion(i).Num_Orden & ","
        Sql = Sql & "'" & stLiquidacion(i).Cod_TipoPago & "',"
        If stLiquidacion(i).Gls_Direccion <> "NULL" Then
            Sql = Sql & "'" & stLiquidacion(i).Gls_Direccion & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        If stLiquidacion(i).Cod_Direccion <> "NULL" Then
            Sql = Sql & stLiquidacion(i).Cod_Direccion & ","
        Else
            Sql = Sql & "NULL,"
        End If
        Sql = Sql & "'" & stLiquidacion(i).Cod_TipPension & "',"
        Sql = Sql & "'" & stLiquidacion(i).Fec_Pago & "',"
        If stLiquidacion(i).Cod_ViaPago <> "NULL" Then
            Sql = Sql & "'" & stLiquidacion(i).Cod_ViaPago & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        If stLiquidacion(i).Cod_Banco <> "NULL" Then
            Sql = Sql & "'" & stLiquidacion(i).Cod_Banco & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        If stLiquidacion(i).Cod_TipCuenta <> "NULL" Then
            Sql = Sql & "'" & stLiquidacion(i).Cod_TipCuenta & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        If stLiquidacion(i).Num_Cuenta <> "NULL" Then
            Sql = Sql & "'" & stLiquidacion(i).Num_Cuenta & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        If stLiquidacion(i).Cod_Sucursal <> "NULL" Then
            Sql = Sql & "'" & stLiquidacion(i).Cod_Sucursal & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        If stLiquidacion(i).Cod_InsSalud <> "NULL" Then
            Sql = Sql & "'" & stLiquidacion(i).Cod_InsSalud & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        Sql = Sql & "'" & stLiquidacion(i).Num_IdenReceptor & "',"
        Sql = Sql & "'" & stLiquidacion(i).Cod_TipoIdenReceptor & "',"
        Sql = Sql & "'" & stLiquidacion(i).Gls_NomReceptor & "',"
        If stLiquidacion(i).Gls_NomSegReceptor <> "NULL" Then
            Sql = Sql & "'" & stLiquidacion(i).Gls_NomSegReceptor & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        Sql = Sql & "'" & stLiquidacion(i).Gls_PatReceptor & "',"
        If stLiquidacion(i).Gls_MatReceptor <> "NULL" Then
            Sql = Sql & "'" & stLiquidacion(i).Gls_MatReceptor & "',"
        Else
            Sql = Sql & "NULL,"
        End If
        Sql = Sql & "'" & stLiquidacion(i).Cod_TipReceptor & "',"
        Sql = Sql & "" & str(stLiquidacion(i).Mto_Haber) & ","
        Sql = Sql & "" & str(stLiquidacion(i).Mto_Descuento) & ","
        Sql = Sql & "" & str(stLiquidacion(i).Mto_LiqPagar) & ","
        Sql = Sql & "" & str(stLiquidacion(i).Mto_BaseImp) & ","
        Sql = Sql & "" & str(stLiquidacion(i).Mto_BaseTri) & ","
        Sql = Sql & "" & str(stLiquidacion(i).Mto_Pension) & ","
        Sql = Sql & "" & str(stLiquidacion(i).Mto_Salud) & ","
        Sql = Sql & "'" & (stLiquidacion(i).Cod_Moneda) & "',"
        Sql = Sql & "'" & clModSaludDefecto & "',"
        Sql = Sql & "'" & fgConvierteNumeroLetras(vlMontoPension, stLiquidacion(i).Gls_Moneda) & "')"
        
        vgConectarBD.Execute (Sql)
                
        If Not flGrabaPensionActualizada(i) Then
            Exit Function
        End If
    Next i
    
    'Graba Detalle Liquidación
    For i = 0 To UBound(stDetPension) - 1
        Sql = "INSERT INTO pd_tmae_pagopen "
        Sql = Sql & "(num_perpago, num_poliza, num_orden, cod_conhabdes, "
        Sql = Sql & "fec_inipago, fec_terpago, mto_conhabdes, num_idenreceptor,"
        Sql = Sql & "cod_tipoidenreceptor,cod_tipreceptor) VALUES ("
        Sql = Sql & "'" & stDetPension(i).Num_PerPago & "',"
        Sql = Sql & "'" & stDetPension(i).Num_Poliza & "',"
        Sql = Sql & stDetPension(i).Num_Orden & ","
        Sql = Sql & "'" & stDetPension(i).Cod_ConHabDes & "',"
        Sql = Sql & "'" & stDetPension(i).Fec_IniPago & "',"
        Sql = Sql & "'" & stDetPension(i).Fec_TerPago & "',"
        Sql = Sql & str(stDetPension(i).Mto_ConHabDes) & ","
        Sql = Sql & "'" & stDetPension(i).Num_IdenReceptor & "',"
        Sql = Sql & stDetPension(i).Cod_TipoIdenReceptor & ","
        Sql = Sql & "'" & stDetPension(i).Cod_TipReceptor & "')"
        vgConectarBD.Execute (Sql)
    Next i
    flGrabarPrimerPago = True
    Exit Function
Errores:
End Function

Function flGrabarPolizaRec()
    Dim vlVigPoliza As String
    Dim vlAño As Long, vlMes As Long, vlDia As Long
    
    vlVigPoliza = 0
     
    Sql = ""
    Sql = "INSERT INTO pd_tmae_poliza ("
    Sql = Sql & "num_poliza,num_endoso,num_cot,num_correlativo," 'cod_solofe,"
    'Sql = Sql & "cod_ideofe,cod_secofe,"
    Sql = Sql & "num_operacion," 'agregado
    Sql = Sql & "num_archivo,cod_trapagopen,fec_trapagopen,"
    Sql = Sql & "cod_afp,cod_isapre,cod_tippension,cod_vejez,cod_estcivil,"
    Sql = Sql & "cod_cuspp, cod_tipoidenafi,num_idenafi, " 'agregado
    Sql = Sql & "gls_direccion,cod_direccion,gls_fono,gls_correo,"
    Sql = Sql & "cod_viapago,cod_tipcuenta,cod_banco,num_cuenta,cod_sucursal,"
    Sql = Sql & "fec_solicitud,fec_ingvigencia,fec_vigencia,fec_dev,"
    Sql = Sql & "fec_acepta,fec_pripago," 'agregado
    Sql = Sql & "cod_monedafon,mto_monedafon,"
    Sql = Sql & "mto_priunifon, mto_ctaindfon,mto_bonofon,mto_apoadi," 'agregado
    Sql = Sql & "mto_priuni, mto_ctaind, mto_bono,prc_tasarprt," 'agregado
    Sql = Sql & "cod_tipoidencor,num_idencor,prc_corcom,mto_corcom,"
    Sql = Sql & "prc_corcomreal," 'agregado
    'Sql = Sql & "cod_corafeiva,"
    Sql = Sql & "num_annojub,num_cargas,"
    Sql = Sql & "ind_cob,cod_bensocial,cod_moneda,mto_valmoneda,mto_priunimod," 'agregado
    Sql = Sql & "mto_ctaindmod,mto_bonomod," 'agregado
    Sql = Sql & "cod_tipren,num_mesdif,"
    Sql = Sql & "cod_modalidad,num_mesgar,"
    'Sql = Sql & "prc_rentaafp,prc_rentaafpori,prc_rentatmp,"
    Sql = Sql & "cod_cobercon," 'agregado
    Sql = Sql & "mto_facpenella,prc_facpenella,"
    Sql = Sql & "cod_dercre, cod_dergra, prc_rentaafp," 'agregado
    Sql = Sql & "prc_rentaafpori,prc_rentatmp," 'agregado
    Sql = Sql & "mto_cuomor,"
    'Sql = Sql & "mto_ctaind,mto_bono,mto_priuni,"
    Sql = Sql & "prc_tasace,prc_tasavta,"
    Sql = Sql & "prc_tasatir,prc_tasapergar,mto_cnu,mto_priunisim,mto_priunidif,"
    Sql = Sql & "mto_pension,mto_pensiongar,mto_ctaindafp,mto_rentatmpafp,"
    Sql = Sql & "mto_resmat,mto_valprepentmp,mto_percon,prc_percon,"
    'Sql = Sql & "cod_eld,mto_eld,cod_tipeld,mto_eldofeuf,"
    Sql = Sql & "mto_sumpension,mto_penanual,mto_rmpension,mto_rmgtosep," 'agregado
    Sql = Sql & "cod_tipcot,cod_estcot,"
    Sql = Sql & "cod_usuario , cod_sucursalusu,"
    Sql = Sql & "fec_inipagopen,cod_usuariocrea,fec_crea,hor_crea,"
    Sql = Sql & "cod_succorredor,fec_finperdif,fec_finpergar," 'agregado
    Sql = Sql & "gls_nacionalidad, ind_recalculo," 'agregado
    Sql = Sql & "fec_emision, fec_inipencia, mto_rmgtoseprv, fec_calculo"
'I--- ABV 14/10/2007 ---
    Sql = Sql & ",mto_apoadifon,mto_apoadimod "
    'RVF 20090914
    Sql = Sql & ",cod_tipvia,gls_nomvia,gls_numdmc,gls_intdmc,cod_tipzon,gls_nomzon,gls_referencia "
'F--- ABV 14/10/2007 ---
'I--- ABV 05/02/2011 ---
'MVG 20170904 AGREGO ind_BolElec
    Sql = Sql & ",cod_tipreajuste,mto_valreajustetri,mto_valreajustemen, fec_devsol, ind_bendes , ind_BolElec, cod_nacionalidad "
'F--- ABV 05/02/2011 ---
'GCP-FRACTAL 08042019
     Sql = Sql & ", NUM_CUENTA_CCI, COD_MONCTA"
     Sql = Sql & ",num_idensup , num_idenjef,gls_telben2 "
    Sql = Sql & ") VALUES ("
    Sql = Sql & "'" & (vlRegistro1!Num_Poliza) & "',"
    Sql = Sql & " " & (vlEndoso) & ","
    Sql = Sql & "'" & (vlRegistro1!Num_Cot) & "',"
    Sql = Sql & "" & (vlRegistro1!Num_Correlativo) & ","
    Sql = Sql & "" & (vlRegistro1!Num_Operacion) & ","
    Sql = Sql & "" & (vlRegistro1!Num_Archivo) & ","
    Sql = Sql & "'" & (vlCodTraspaso) & "',"
    Sql = Sql & "NULL" & ","
    Sql = Sql & "'" & (vlRegistro1!Cod_AFP) & "',"
    Sql = Sql & "'" & (vlRegistro1!cod_isapre) & "',"
    Sql = Sql & "'" & (vlRegistro1!Cod_TipPension) & "',"
    Sql = Sql & "'" & (vlRegistro1!cod_vejez) & "',"
    Sql = Sql & "'" & (vlRegistro1!cod_estcivil) & "',"
    Sql = Sql & "'" & (vlRegistro1!Cod_Cuspp) & "',"
    Sql = Sql & " " & (vlRegistro1!cod_tipoidenafi) & ","
    Sql = Sql & "'" & (vlRegistro1!num_idenafi) & "',"
    Sql = Sql & "'" & (vlRegistro1!Gls_Direccion) & "',"
    Sql = Sql & " " & (vlRegistro1!Cod_Direccion) & ","
    If IsNull(vlRegistro1!GLS_FONO) Then
        Sql = Sql & "NULL" & ","
    Else
        Sql = Sql & "'" & (vlRegistro1!GLS_FONO) & "',"
    End If
    If IsNull(vlRegistro1!GLS_CORREO) Then
        Sql = Sql & "NULL" & ","
    Else
        Sql = Sql & "'" & (vlRegistro1!GLS_CORREO) & "',"
    End If
    Sql = Sql & "'" & (vlRegistro1!Cod_ViaPago) & "',"
    Sql = Sql & "'" & (vlRegistro1!Cod_TipCuenta) & "',"
    Sql = Sql & "'" & (vlRegistro1!Cod_Banco) & "',"
    If IsNull(vlRegistro1!Num_Cuenta) Then
        Sql = Sql & "NULL" & ","
    Else
        Sql = Sql & "'" & (vlRegistro1!Num_Cuenta) & "',"
    End If
    Sql = Sql & "'" & (vlRegistro1!Cod_Sucursal) & "',"
    Sql = Sql & "'" & (vlRegistro1!Fec_Solicitud) & "',"
    Sql = Sql & "'" & (vlRegistro1!Fec_Vigencia) & "',"
    Sql = Sql & "'" & (vlFVigencia) & "',"
    Sql = Sql & "'" & (vlRegistro1!Fec_Dev) & "',"
    Sql = Sql & "'" & (vlRegistro1!Fec_Acepta) & "',"
    Sql = Sql & "'" & Format(Txt_FecPriPag, "yyyymmdd") & "',"
    Sql = Sql & "'" & (vlRegistro1!Cod_MonedaFon) & "',"
    Sql = Sql & "" & str(vlRegistro1!Mto_MonedaFon) & ","
    
'    vlCtaInd = Format(CDbl(lbl_mtoprimarec) - (vlRegistro1!Mto_Bono), "#0.00")
'I--- ABV 07/11/2007 ---
'    vlCtaInd = Format(CDbl(Lbl_PrimaDefCia) - (vlRegistro1!Mto_Bono + vlRegistro1!Mto_ApoAdi), "#0.00")
    vlCtaInd = Format(CDbl(Lbl_PrimaDefCia) - (0 + 0), "#0.00")
'F--- ABV 07/11/2007 ---
    
'I--- ABV 21/08/2007 ---
''    Sql = Sql & "" & Str(vlRegistro1!mto_priunifon) & ","
''    Sql = Sql & "" & Str(vlRegistro1!Mto_CtaIndFon) & ","
'    Sql = Sql & "" & Str(CDbl(lbl_mtoprimarec) / vlRegistro1!Mto_MonedaFon) & ","
'    Sql = Sql & "" & Str(vlCtaInd / vlRegistro1!Mto_MonedaFon) & ","
    Sql = Sql & " " & str(CDbl(Lbl_PrimaDefCia) / vlRegistro1!Mto_MonedaFon) & ","
    Sql = Sql & " " & str(vlCtaInd / vlRegistro1!Mto_MonedaFon) & ","
'F--- ABV 21/08/2007 ---
'I--- ABV 07/11/2007 ---
'    Sql = Sql & "" & Str(vlRegistro1!Mto_BonoFon) & ","
'    Sql = Sql & "" & Str(vlRegistro1!Mto_ApoAdi) & ","
    Sql = Sql & "0,"
    Sql = Sql & "0,"
'F--- ABV 07/11/2007 ---
    
'    Sql = Sql & " " & Str(lbl_mtoprimarec) & ","
    Sql = Sql & " " & str(Lbl_PrimaDefCia) & ","
    
    Sql = Sql & " " & str(vlCtaInd) & ","
    Sql = Sql & " " & str(vlRegistro1!Mto_Bono) & ","
    Sql = Sql & " " & str(vlRegistro1!Prc_TasaRPRT) & ","
    Sql = Sql & " " & str(vlRegistro1!cod_tipoidencor) & ","
    Sql = Sql & "'" & (vlRegistro1!Num_IdenCor) & "',"
    Sql = Sql & " " & str(vlRegistro1!Prc_CorCom) & ","
    'Sql = Sql & " " & Str(vlRegistro1!mto_corcom) & ","
    
'    Sql = Sql & " " & Str((vlRegistro1!Prc_CorCom / 100) * CDbl(lbl_mtoprimarec)) & "," 'Nuevo Monto de la Comisión
    Sql = Sql & " " & str(Format((vlRegistro1!Prc_CorCom / 100) * CDbl(Lbl_PrimaDefCia), "#0.00")) & "," 'Nuevo Monto de la Comisión
    
    Sql = Sql & " " & str(vlRegistro1!Prc_CorComReal) & ","
    'Sql = Sql & "'" & (vlRegistro1!cod_corafeiva) & "',"
    Sql = Sql & " " & (vlRegistro1!Num_AnnoJub) & ","
    Sql = Sql & " " & (vlRegistro1!Num_Cargas) & ","
    Sql = Sql & "'" & (vlRegistro1!Ind_Cob) & "',"
    Sql = Sql & "'" & (vlRegistro1!Cod_BenSocial) & "',"
    Sql = Sql & "'" & (vlRegistro1!Cod_Moneda) & "',"
    Sql = Sql & " " & str(vlTipoCambioRec) & "," 'Sql = Sql & "" & Str(vlRegistro1!mto_valmoneda) & ","
    
'    Sql = Sql & " " & Str(Format(CDbl(lbl_mtoprimarec) / vlTipoCambioRec, "##0.00")) & "," 'Sql = Sql & "" & Str(vlRegistro1!mto_priunimod) & ","
    Sql = Sql & " " & str(Format(CDbl(Lbl_PrimaDefCia) / vlTipoCambioRec, "##0.00")) & "," 'Sql = Sql & "" & Str(vlRegistro1!mto_priunimod) & ","
    
    Sql = Sql & " " & str(Format(vlCtaInd / vlTipoCambioRec, "##0.00")) & ","
'I--- ABV 07/11/2007 ---
'    Sql = Sql & " " & Str(Format(vlRegistro1!Mto_Bono / vlTipoCambioRec, "##0.00")) & ","
    Sql = Sql & " 0,"
'F--- ABV 07/11/2007 ---
    Sql = Sql & "'" & (vlRegistro1!Cod_TipRen) & "',"
    Sql = Sql & " " & (vlRegistro1!Num_MesDif) & ","
    Sql = Sql & "'" & (vlRegistro1!Cod_Modalidad) & "',"
    Sql = Sql & " " & (vlRegistro1!Num_MesGar) & ","
    Sql = Sql & "'" & Trim(vlRegistro1!Cod_CoberCon) & "',"
    Sql = Sql & " " & str(vlRegistro1!Mto_FacPenElla) & ","
    Sql = Sql & " " & str(vlRegistro1!Prc_FacPenElla) & ","
    Sql = Sql & "'" & Trim(vlRegistro1!Cod_DerCre) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!Cod_DerGra) & "',"
    Sql = Sql & " " & str(vlRegistro1!Prc_RentaAFP) & ","
    Sql = Sql & " " & str(vlRegistro1!prc_rentaafpori) & ","
    Sql = Sql & " " & str(vlRegistro1!Prc_RentaTMP) & ","
    Sql = Sql & " " & str(vlRegistro1!Mto_CuoMor) & ","
    Sql = Sql & " " & str(vlRegistro1!Prc_TasaCe) & ","
    Sql = Sql & " " & str(vlRegistro1!Prc_TasaVta) & ","
    Sql = Sql & " " & str(vlRegistro1!Prc_TasaTir) & ","
    Sql = Sql & " " & str(vlRegistro1!prc_tasapergar) & ","
    Sql = Sql & " " & str(vlRegistro1!Mto_CNU) & ","
    If IsNull(vlRegistro1!Mto_PriUniSim) Then
        Sql = Sql & " 0,"
    Else
        Sql = Sql & " " & str(vlRegistro1!Mto_PriUniSim) & ","
    End If
    If IsNull(vlRegistro1!Mto_PriUniDif) Then
        Sql = Sql & " 0,"
    Else
        Sql = Sql & " " & str(vlRegistro1!Mto_PriUniDif) & ","
    End If
    Sql = Sql & " " & str(Lbl_MtoPenDefUf) & ","
    '' CAMBIO PARA LOS MONTOS GARANTIZADOS RRR20181109
    If (vlMtoPenGarUf) = "" Then
        Sql = Sql & " " & str(Format("0", "#0.00")) & ","
    Else
     Sql = Sql & " " & str(vlMtoPenGarUf) & ","
    End If
    
    
    If IsNull(vlRegistro1!Mto_CtaIndAFP) Then
        Sql = Sql & " 0,"
    Else
        Sql = Sql & " " & str(vlRegistro1!Mto_CtaIndAFP) & ","
    End If
    If IsNull(vlRegistro1!Mto_RentaTMPAFP) Then
        Sql = Sql & " 0,"
    Else
        Sql = Sql & " " & str(vlRegistro1!Mto_RentaTMPAFP) & ","
    End If
    If IsNull(vlRegistro1!Mto_ResMat) Then
        Sql = Sql & " 0,"
    Else
        Sql = Sql & " " & str(vlRegistro1!Mto_ResMat) & ","
    End If
    If IsNull(vlRegistro1!Mto_ValPrePenTmp) Then
        Sql = Sql & " 0,"
    Else
        Sql = Sql & " " & str(vlRegistro1!Mto_ValPrePenTmp) & ","
    End If
    Sql = Sql & " " & str(vlRegistro1!Mto_PerCon) & ","
    Sql = Sql & " " & str(vlRegistro1!Prc_PerCon) & ","
    Sql = Sql & " " & str(vlRegistro1!Mto_SumPension) & ","
    Sql = Sql & " " & str(vlRegistro1!Mto_PenAnual) & ","
    Sql = Sql & " " & str(vlRegistro1!Mto_RMPension) & ","
    Sql = Sql & " " & str(vlRegistro1!Mto_RMGtoSep) & ","
    Sql = Sql & "'" & (vlRegistro1!cod_tipcot) & "',"
    Sql = Sql & "'" & (vlRegistro1!Cod_EstCot) & "',"
    Sql = Sql & "'" & (vlRegistro1!COD_USUARIO) & "',"
    If IsNull(vlRegistro1!cod_sucursalusu) Then
        Sql = Sql & " NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro1!cod_sucursalusu) & "',"
    End If
    
    If (vlRegistro1!Num_MesDif) <> 0 Then
         'vlVigPoliza = Format(CDate(Lbl_FecVigPoliza), "yyyymmdd")
         'hqr 01/09/2007 se corrige para que si se hace primer pago,
         'quede la fecha de inicio de los pagos en régimen
         If Trim(Lbl_FecTopePriPag) <> Trim(Txt_FecPriPag) Then
            'Se deja la misma fecha que para los pagos normales
            vlFecIniPP = Format(Txt_FecPriPag, "yyyymmdd")
            vlFecIniPP = DateSerial(Mid(vlFecIniPP, 1, 4), Mid(vlFecIniPP, 5, 2) + 1, 1)
            Sql = Sql & "'" & Format(vlFecIniPP, "yyyymmdd") & "',"
         Else
         'fin hqr 01/09/2007
            vlVigPoliza = vlRegistro1!Fec_Dev 'Desde Fecha de Devengue
            vlAño = Mid(vlVigPoliza, 1, 4)
            vlMes = Mid(vlVigPoliza, 5, 2)
            vlMes = CInt(vlMes) + (vlRegistro1!Num_MesDif)
            'vlDia = Mid(vlVigPoliza, 7, 2)
            vlDia = 1
            vlFecIniPP = DateSerial(vlAño, vlMes, vlDia)
            Sql = Sql & "'" & Format(vlFecIniPP, "yyyymmdd") & "',"
         End If 'hqr 01/09/2007
    Else
         'Sql = Sql & "'" & Trim(vlRegistro1!fec_dev) & "',"
        vlFecIniPP = Format(Txt_FecPriPag, "yyyymmdd")
        vlFecIniPP = DateSerial(Mid(vlFecIniPP, 1, 4), Mid(vlFecIniPP, 5, 2) + 1, 1)
        'Fecha Inicio de Pago de Pensiones en Régimen (un mes después de primeros pagos)
        Sql = Sql & "'" & Format(vlFecIniPP, "yyyymmdd") & "',"
    End If
    
    Sql = Sql & "'" & (vgUsuario) & "',"
    Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
    Sql = Sql & "'" & Format(Time, "hhmmss") & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!cod_succorredor) & "',"
    If IsNull(vlRegistro1!fec_finperdif) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & vlRegistro1!fec_finperdif & "',"
    End If
    If IsNull(vlRegistro1!fec_finpergar) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & Trim(vlRegistro1!fec_finpergar) & "',"
    End If
    Sql = Sql & "'" & Trim(vlRegistro1!gls_nacionalidad) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!ind_recalculo) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!Fec_Emision) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!fec_inipencia) & "',"
    Sql = Sql & " " & str(vlRegistro1!Mto_RMGtoSepRV) & ","
    Sql = Sql & "'" & Trim(vlRegistro1!Fec_Calculo) & "',"
'I--- ABV 07/11/2007 ---
''I--- ABV 14/10/2007 ---
'    Sql = Sql & "," & Str(Format(vlRegistro1!Mto_ApoAdiFon / vlRegistro1!Mto_MonedaFon, "#0.00")) & ","
'    Sql = Sql & " " & Str(Format((vlRegistro1!Mto_ApoAdi / vlTipoCambioRec), "#0.00")) & " "
    Sql = Sql & " 0,"
    Sql = Sql & " 0,"
''    Sql = Sql & " " & Str(vlRegistro1!Mto_ApoAdiMod) & " "
''F--- ABV 14/10/2007 ---
'F--- ABV 07/11/2007 ---
    
    'RVF 20090914
    Sql = Sql & "'" & Trim(vlRegistro1!cod_tipvia) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!gls_nomvia) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!gls_numdmc) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!gls_intdmc) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!cod_tipzon) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!gls_nomzon) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!gls_referencia) & "'"

'I--- ABV 05/02/2011 ---
    Sql = Sql & ",'" & vlRegistro1!Cod_TipReajuste & "',"
    Sql = Sql & " " & str(vlRegistro1!Mto_ValReajusteTri) & ","
    Sql = Sql & " " & str(vlRegistro1!Mto_ValReajusteMen) & ", "
'F--- ABV 05/02/2011 ---
    Sql = Sql & " " & str(vlRegistro1!fec_devsol) & ", " 'RRR 18/9/13
    Sql = Sql & " '" & strID & "','" & strBolElec & "'," 'mvg 20170904
    '-- Begin : Modify by : ricardo.huerta
    Sql = Sql & "'" & Trim(vlRegistro1!cod_nacionalidad) & "', "
    '-- End    :  Modify by : ricardo.huerta
    'GCP-FRACTAL 08042019
    Sql = Sql & "'" & Trim(vlRegistro1!num_cuenta_cci) & "',"
    Sql = Sql & "'" & Trim(vlRegistro1!COD_MONCTA) & "', "
    Sql = Sql & "'" & Trim(vlRegistro1!num_idensup) & "', "
    Sql = Sql & "'" & Trim(vlRegistro1!num_idenjef) & "', "
    Sql = Sql & "'" & Trim(vlRegistro1!gls_fono2) & "' "
    Sql = Sql & ")"
    vgConectarBD.Execute (Sql)

End Function

Function flGrabarBonosRec()

    Sql = ""
    Sql = "INSERT INTO pd_tmae_polbon ("
    Sql = Sql & "num_poliza,num_endoso,cod_tipobono,mto_valnom,"
    Sql = Sql & "fec_emi,fec_ven,prc_tasaint,mto_bonoact,"
    Sql = Sql & "mto_bonoactuf,mto_compra,cod_afeley,num_edadcob,"
    Sql = Sql & "cod_usuariocrea,fec_crea,hor_crea"
    Sql = Sql & ") VALUES ("
    Sql = Sql & "'" & (vlRegistro2!Num_Poliza) & "',"
    Sql = Sql & " " & (vlEndoso) & ","
    Sql = Sql & "'" & (vlRegistro2!cod_tipobono) & "',"
    Sql = Sql & " " & str(vlRegistro2!mto_valnom) & ","
    Sql = Sql & "'" & (vlRegistro2!fec_emi) & "',"
    Sql = Sql & "'" & (vlRegistro2!fec_ven) & "',"
    Sql = Sql & " " & str(vlRegistro2!prc_tasaint) & ","
    Sql = Sql & " " & str(vlRegistro2!mto_bonoact) & ","
    Sql = Sql & " " & str(vlRegistro2!mto_bonoactuf) & ","
    Sql = Sql & " " & str(vlRegistro2!mto_compra) & ","
    Sql = Sql & "'" & (vlRegistro2!cod_afeley) & "',"
    Sql = Sql & " " & str(vlRegistro2!num_edadcob) & ","
    Sql = Sql & "'" & (vgUsuario) & "',"
    Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
    Sql = Sql & "'" & Format(Time, "hhmmss") & "'"
    Sql = Sql & " )"
    vgConectarBD.Execute (Sql)

End Function

Private Sub pGrabarPolizaRep()
    Sql = ""
    
    Sql = "INSERT INTO pd_tmae_polrep "
    Sql = Sql & "(NUM_POLIZA,NUM_ENDOSO,COD_TIPOIDENREP,NUM_IDENREP,GLS_NOMBRESREP,GLS_APEPATREP,"
    Sql = Sql & "GLS_APEMATREP,Cod_UsuarioCrea,Fec_Crea,Hor_Crea, gls_fono,gls_fono2,gls_correo, cod_sexo "
    Sql = Sql & ") VALUES ("
    Sql = Sql & "'" & (vlRegistro4!Num_Poliza) & "',"
    Sql = Sql & " " & (vlEndoso) & ","
    Sql = Sql & " " & (vlRegistro4!cod_tipoidenRep) & ","
    Sql = Sql & "'" & (vlRegistro4!Num_idenrep) & "',"
    Sql = Sql & "'" & (vlRegistro4!Gls_NombresRep) & "',"
    Sql = Sql & "'" & (vlRegistro4!Gls_ApepatRep) & "',"
    Sql = Sql & "'" & (vlRegistro4!Gls_ApematRep) & "',"
    Sql = Sql & "'" & (vgUsuario) & "',"
    Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
    Sql = Sql & "'" & Format(Time, "hhmmss") & "',"
    Sql = Sql & "'" & (vlRegistro4!GLS_TELREP1) & "',"
    Sql = Sql & "'" & (vlRegistro4!GLS_TELREP2) & "',"
    Sql = Sql & "'" & (vlRegistro4!gls_correorep) & "',"
    Sql = Sql & "'" & (vlRegistro4!Cod_Sexo) & "'"
    Sql = Sql & ")"
    vgConectarBD.Execute (Sql)

End Sub

Function flGrabarBeneficiariosRec()
    Dim vlPrcPenBen As Double, vlMtoPenDefUF As Double
    
    Sql = ""
    Sql = "INSERT INTO pd_tmae_polben "
    Sql = Sql & "(num_poliza,num_endoso,num_orden,cod_par,cod_grufam,cod_sexo,"
    Sql = Sql & "cod_sitinv,fec_invben,cod_cauinv,cod_derpen,cod_dercre,"
    Sql = Sql & "cod_tipoidenben,num_idenben,gls_nomben,gls_nomsegben, gls_patben,gls_matben,fec_nacben,"
    Sql = Sql & "fec_fallben,fec_nachm,prc_pension,prc_pensionleg,"
    Sql = Sql & "prc_pensionrep,mto_pension,mto_pensiongar,"
    'mvg 20170904
    Sql = Sql & "cod_usuariocrea,fec_crea,hor_crea,cod_estpension,prc_pensiongar, cod_tipcta , cod_monbco , Cod_Banco , num_ctabco,ind_bolElec, cod_nacionalidad, "
    Sql = Sql & "NUM_CUENTA_CCI, gls_fono, gls_fono2, GLS_CORREO, "
    'Inicio SMCCB 22/09/2023
    Sql = Sql & "COD_MODTIPOCUENTA_MANC , COD_TIPODOC_MANC, NUM_DOC_MANC, NOMBRE_MANC, APELLIDO_MANC"
    'Fin SMCCB 22/09/2023
    Sql = Sql & ") VALUES ("
    Sql = Sql & "'" & (vlRegistro3!Num_Poliza) & "',"
    Sql = Sql & " " & (vlEndoso) & ","
    Sql = Sql & " " & (vlRegistro3!Num_Orden) & ","
    Sql = Sql & "'" & (vlRegistro3!Cod_Par) & "',"
    Sql = Sql & "'" & (vlRegistro3!Cod_GruFam) & "',"
    Sql = Sql & "'" & (vlRegistro3!Cod_Sexo) & "',"
    Sql = Sql & "'" & (vlRegistro3!Cod_SitInv) & "',"
    If IsNull(vlRegistro3!Fec_InvBen) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro3!Fec_InvBen) & "',"
    End If
    If IsNull(vlRegistro3!Cod_CauInv) Then
        Sql = Sql & "'0',"
    Else
        Sql = Sql & "'" & (vlRegistro3!Cod_CauInv) & "',"
    End If
    If IsNull(vlRegistro3!Cod_DerPen) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro3!Cod_DerPen) & "',"
    End If
    If IsNull(vlRegistro3!Cod_DerCre) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro3!Cod_DerCre) & "',"
    End If
    If IsNull(vlRegistro3!Cod_TipoIdenBen) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & " " & (vlRegistro3!Cod_TipoIdenBen) & ","
    End If
    
    If IsNull(vlRegistro3!Num_IdenBen) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro3!Num_IdenBen) & "',"
    End If
    If IsNull(vlRegistro3!Gls_NomBen) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro3!Gls_NomBen) & "',"
    End If
    If IsNull(vlRegistro3!Gls_NomSegBen) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro3!Gls_NomSegBen) & "',"
    End If
    If IsNull(vlRegistro3!Gls_PatBen) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro3!Gls_PatBen) & "',"
    End If
    If IsNull(vlRegistro3!Gls_MatBen) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro3!Gls_MatBen) & "',"
    End If
    Sql = Sql & "'" & (vlRegistro3!Fec_NacBen) & "',"
    If IsNull(vlRegistro3!Fec_FallBen) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro3!Fec_FallBen) & "',"
    End If
    If IsNull(vlRegistro3!Fec_NacHM) Then
        Sql = Sql & "NULL,"
    Else
        Sql = Sql & "'" & (vlRegistro3!Fec_NacHM) & "',"
    End If
    Sql = Sql & " " & str(vlRegistro3!Prc_Pension) & ","
    Sql = Sql & " " & str(vlRegistro3!Prc_PensionLeg) & ","
    Sql = Sql & " " & str(vlRegistro3!Prc_PensionRep) & ","
    'Definir el Monto de Pensión a recibir por cada Pensionado
    vlPrcPenBen = Format((vlRegistro3!Prc_Pension), "#0.00")
    vlMtoPenDefUF = Format((CDbl(Lbl_MtoPenDefUf) * CDbl(vlPrcPenBen) / 100), "#0.00")
    Sql = Sql & " " & str(vlMtoPenDefUF) & ","
    If vlMtoPenGarUf = "0" Then
       Sql = Sql & " " & str(Format("0", "#0.00")) & ","
    Else
'I--- ABV 14/02/2011 ---
'      vlMtoPenGar = Format((CDbl(vlMtoPenGarUf) * CDbl(vlPrcPenBen) / 100), "#0.00")
'Obtener el monto de la Pensión Garantizada desde la Pensión de Referencia Actualizada
      vlMtoPenGar = Format((CDbl(Lbl_MtoPenDefUf) * CDbl(vlPrcPenBen) / 100), "#0.00")
'F--- ABV 14/02/2011 ---
      Sql = Sql & " " & str(vlMtoPenGar) & ","
    End If
    Sql = Sql & "'" & (vgUsuario) & "',"
    Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
    Sql = Sql & "'" & Format(Time, "hhmmss") & "',"
    Sql = Sql & "'" & vlRegistro3!Cod_EstPension & "'"
    Sql = Sql & "," & str(vlRegistro3!Prc_PensionGar) & " "
    'RRR 26/12/2013
    Sql = Sql & ",'" & vlRegistro3!cod_tipcta & "' "
    Sql = Sql & ",'" & vlRegistro3!cod_monbco & "' "
    Sql = Sql & ",'" & vlRegistro3!Cod_Banco & "' "
    Sql = Sql & ",'" & vlRegistro3!num_ctabco & "' "
    'mvg 20170904
    Sql = Sql & ",'" & strBolElec & "' "
    '-- Begin : Modify by : ricardo.huerta
    Sql = Sql & ",'" & vlRegistro3!cod_nacionalidad & "' "
    '-- End    :  Modify by : ricardo.huerta
    'RRR
    'GCP-FRACTAL 04042019
    Sql = Sql & ",'" & vlRegistro3!num_cuenta_cci & "' "
    Sql = Sql & ",'" & vlRegistro3!GLS_FONO & "' "
    Sql = Sql & ",'" & vlRegistro3!gls_fono2 & "' "
    Sql = Sql & ",'" & vlRegistro3!Gls_CorreoBen & "' "
    'Inicio SMCCB 22/09/2023
    Sql = Sql & ",'" & vlRegistro3!COD_MODTIPOCUENTA_MANC & "' "
    Sql = Sql & ",'" & vlRegistro3!COD_TIPODOC_MANC & "' "
    Sql = Sql & ",'" & vlRegistro3!NUM_DOC_MANC & "' "
    Sql = Sql & ",'" & vlRegistro3!NOMBRE_MANC & "' "
    Sql = Sql & ",'" & vlRegistro3!APELLIDO_MANC & "' "
    'Fin SMCCB 22/09/2023
    Sql = Sql & ")"
    vgConectarBD.Execute (Sql)

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

Function flInicializaGrillaVarTipoCambio()
    
    Msf_VarTipCambio.Clear
    Msf_VarTipCambio.Cols = 4
    Msf_VarTipCambio.rows = 4
    Msf_VarTipCambio.RowHeight(0) = 250
    Msf_VarTipCambio.Row = 0
    
    Msf_VarTipCambio.Col = 1
    Msf_VarTipCambio.Text = "Fondo (TM)"
    Msf_VarTipCambio.ColWidth(1) = 1900
    
    Msf_VarTipCambio.Col = 2
    Msf_VarTipCambio.Text = "Tipo Cambio (TM)"
    Msf_VarTipCambio.ColWidth(2) = 1900
    
    Msf_VarTipCambio.Col = 3
    Msf_VarTipCambio.Text = "Pensión"
    Msf_VarTipCambio.ColWidth(3) = 1900
        
    Msf_VarTipCambio.Row = 1
    Msf_VarTipCambio.Col = 0
    Msf_VarTipCambio.Text = "Cotización Inicial"
    Msf_VarTipCambio.ColWidth(0) = 1700
    
    Msf_VarTipCambio.Row = 2
    Msf_VarTipCambio.Col = 0
    Msf_VarTipCambio.Text = "Var. Tipo Cambio"
    
    Msf_VarTipCambio.Row = 3
    Msf_VarTipCambio.Col = 0
    Msf_VarTipCambio.Text = "Variación (%)"
        
End Function

'''Function flCargaGrillaBono(iNumPoliza As String)
'''
'''On Error GoTo Err_flCargaGrillaBono
'''
'''    vlMtoTotalBono = 0
'''
'''    vgSql = ""
'''    vgSql = "SELECT  b.cod_tipobono,b.mto_valnom,b.fec_ven "
'''    vgSql = vgSql & "FROM pd_tmae_oripolbon b "
'''    vgSql = vgSql & "WHERE b.num_poliza = '" & Trim(iNumPoliza) & "' "
'''    vgSql = vgSql & "ORDER BY cod_tipobono "
'''    Set vgRs = vgConexionBD.Execute(vgSql)
'''    If Not vgRs.EOF Then
'''       Call flInicializaGrillaVarFondo
'''
'''       While Not vgRs.EOF
'''            vlMtoTotalBono = vlMtoTotalBono + (vgRs!mto_valnom)
'''
'''          Msf_GrillaBono.AddItem Trim(vgRs!cod_tipobono) & vbTab _
'''          & (Format((vgRs!mto_valnom), "###,###,##0.00")) & vbTab _
'''          & (DateSerial(Mid((vgRs!fec_ven), 1, 4), Mid((vgRs!fec_ven), 5, 2), Mid((vgRs!fec_ven), 7, 2)))
'''
'''          vgRs.MoveNext
'''       Wend
'''    End If
'''    vgRs.Close
'''
'''Exit Function
'''Err_flCargaGrillaBono:
'''    Screen.MousePointer = 0
'''    Select Case Err
'''        Case Else
'''        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'''    End Select
'''
'''End Function


Function flInicializaGrillaVarFondo()
    
    Msf_VarFondo.Clear
    Msf_VarFondo.Cols = 4
    Msf_VarFondo.rows = 4
    Msf_VarFondo.RowHeight(0) = 250
    Msf_VarFondo.Row = 0
    
    Msf_VarFondo.Col = 1
    Msf_VarFondo.Text = "Fondo (TM)"
    Msf_VarFondo.ColWidth(1) = 1900
    
    Msf_VarFondo.Col = 2
    Msf_VarFondo.Text = "Tipo Cambio (TM)"
    Msf_VarFondo.ColWidth(2) = 1900
    
    Msf_VarFondo.Col = 3
    Msf_VarFondo.Text = "Pensión"
    Msf_VarFondo.ColWidth(3) = 1900
        
    'Msf_VarFondo.RowHeight(0) = 1500
    Msf_VarFondo.Row = 1
    Msf_VarFondo.Col = 0
    Msf_VarFondo.Text = "Cotización Inicial"
    Msf_VarFondo.ColWidth(0) = 1700
    
    Msf_VarFondo.Row = 2
    Msf_VarFondo.Col = 0
    Msf_VarFondo.Text = "Var. Fondo"
    
    Msf_VarFondo.Row = 3
    Msf_VarFondo.Col = 0
    Msf_VarFondo.Text = "Variación (%)"
    
End Function


Function flInicializaGrillaVarFondoTC()
    
    Msf_VarFondoTC.Clear
    Msf_VarFondoTC.Cols = 4
    Msf_VarFondoTC.rows = 4
    Msf_VarFondoTC.RowHeight(0) = 250
    Msf_VarFondoTC.Row = 0
    
    Msf_VarFondoTC.Col = 1
    Msf_VarFondoTC.Text = "Fondo (TM)"
    Msf_VarFondoTC.ColWidth(1) = 1900
    
    Msf_VarFondoTC.Col = 2
    Msf_VarFondoTC.Text = "Tipo Cambio (TM)"
    Msf_VarFondoTC.ColWidth(2) = 1900
    
    Msf_VarFondoTC.Col = 3
    Msf_VarFondoTC.Text = "Pensión"
    Msf_VarFondoTC.ColWidth(3) = 1900
        
    Msf_VarFondoTC.Row = 1
    Msf_VarFondoTC.Col = 0
    Msf_VarFondoTC.Text = "Cotización Inicial"
    Msf_VarFondoTC.ColWidth(0) = 1700
    
    Msf_VarFondoTC.Row = 2
    Msf_VarFondoTC.Col = 0
    Msf_VarFondoTC.Text = "Var. Fondo y T.C."
    
    Msf_VarFondoTC.Row = 3
    Msf_VarFondoTC.Col = 0
    Msf_VarFondoTC.Text = "Variación (%)"
        
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
    
    If vlFactorDef = 0 Then
        vlFactorDef = 1
    End If
     
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
        vlSql = vlSql & Trim(str(vlNewNumFactura)) & "','"
        vlSql = vlSql & Trim(str(vlNewNumRenVit)) & "','"
        vlSql = vlSql & vgUsuario & "','"
        vlSql = vlSql & Format(Date, "yyyymmdd") & "','"
        vlSql = vlSql & Format(Time, "hhmmss") & "')"
    Else
        vlSql = "UPDATE pd_tmae_gennumfac SET "
        vlSql = vlSql & "num_factura = '" & Trim(str(vlNewNumFactura)) & "',"
        vlSql = vlSql & "cod_renvit = '" & Trim(str(vlNewNumRenVit)) & "',"
        vlSql = vlSql & "cod_usuariomodi = '" & vgUsuario & "',"
        vlSql = vlSql & "fec_modi = '" & Format(Date, "yyyymmdd") & "',"
        vlSql = vlSql & "hor_modi = '" & Format(Time, "hhmmss") & "'"
    End If
    iConexion.Execute (vlSql)
    
    oNumFactura = vlNewNumFactura
    oNumRenVit = vlNewNumRenVit
    
    flObtieneNumFactura = True
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
'Ini GCP 15092022 Correccion Montos en Cero
Private Sub CorreccionMontosCero()
Dim i As Integer
       For i = 0 To UBound(stLiquidacion) - 1
             If Format(stLiquidacion(i).Mto_Pension, "###,##0.00") = "0.00" Then
                            
                    If (stLiquidacion(i - 1).Num_Orden = stLiquidacion(i).Num_Orden) Then
                        stLiquidacion(i).Mto_Haber = stLiquidacion(i - 1).Mto_Haber
                        stLiquidacion(i).Mto_Descuento = stLiquidacion(i - 1).Mto_Descuento
                        stLiquidacion(i).Mto_LiqPagar = stLiquidacion(i - 1).Mto_LiqPagar
                        stLiquidacion(i).Mto_BaseImp = stLiquidacion(i - 1).Mto_BaseImp
                        stLiquidacion(i).Mto_BaseTri = stLiquidacion(i - 1).Mto_BaseTri
                        stLiquidacion(i).Mto_Pension = stLiquidacion(i - 1).Mto_Pension
                        stLiquidacion(i).Mto_Salud = stLiquidacion(i - 1).Mto_Salud
                        
                    
                    End If
  
             End If
  
       Next i
   

End Sub
'Fin GCP 15092022 Correccion Montos en Cero
