VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_CalConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Póliza"
   ClientHeight    =   8175
   ClientLeft      =   1425
   ClientTop       =   1365
   ClientWidth     =   10650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10650
   Begin VB.CommandButton cmdPlanPol 
      Caption         =   "Plantilla de polizas"
      Height          =   615
      Left            =   9000
      TabIndex        =   210
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   118
      Top             =   6960
      Width           =   10455
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6240
         Picture         =   "Frm_CalConsulta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4800
         Picture         =   "Frm_CalConsulta.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3360
         Picture         =   "Frm_CalConsulta.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Poliza 
         Left            =   720
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Fra_Cabeza 
      Height          =   735
      Left            =   120
      TabIndex        =   71
      Top             =   840
      Width           =   10455
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Nº Operación"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   135
         Top             =   263
         Width           =   975
      End
      Begin VB.Label Lbl_SolOfe 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4920
         TabIndex        =   68
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Lbl_NumCot 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1920
         TabIndex        =   67
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label Lbl_FecVig 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7920
         TabIndex        =   69
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Nº de Cotizacion"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   78
         Top             =   263
         Width           =   1335
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Fec. de Emisión"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   77
         Top             =   263
         Width           =   1335
      End
   End
   Begin VB.CommandButton Cmd_Poliza 
      Caption         =   "&Poliza"
      Height          =   675
      Left            =   7440
      Picture         =   "Frm_CalConsulta.frx":0E6E
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Buscar Datos de la Póliza"
      Top             =   120
      Width           =   840
   End
   Begin TabDlg.SSTab SSTab_Poliza 
      Height          =   5295
      Left            =   120
      TabIndex        =   79
      Top             =   1680
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9340
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Datos del Afiliado"
      TabPicture(0)   =   "Frm_CalConsulta.frx":1528
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_Afiliado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos de Cálculo"
      TabPicture(1)   =   "Frm_CalConsulta.frx":1544
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Fra_Calculo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Datos de Asegurados"
      TabPicture(2)   =   "Frm_CalConsulta.frx":1560
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Fra_Benef"
      Tab(2).ControlCount=   1
      Begin VB.Frame Fra_Benef 
         Height          =   4845
         Left            =   -74880
         TabIndex        =   144
         Top             =   360
         Width           =   10215
         Begin VB.Frame Fra_DatosBenef 
            Caption         =   "Beneficiario"
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
            Height          =   2955
            Left            =   120
            TabIndex        =   145
            Top             =   1800
            Width           =   9975
            Begin VB.Label Lbl_Moneda 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   2
               Left            =   7680
               TabIndex        =   169
               Top             =   2550
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label Lbl_Moneda 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   1
               Left            =   2460
               TabIndex        =   168
               Top             =   2590
               Width           =   375
            End
            Begin VB.Label Lbl_DerPen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6480
               TabIndex        =   16
               Top             =   1970
               Width           =   3135
            End
            Begin VB.Label Lbl_CauInvBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6480
               TabIndex        =   15
               Top             =   1670
               Width           =   3135
            End
            Begin VB.Label Lbl_FecInvBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6480
               TabIndex        =   14
               Top             =   1370
               Width           =   1095
            End
            Begin VB.Label Lbl_SitInvBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6480
               TabIndex        =   13
               Top             =   1080
               Width           =   3135
            End
            Begin VB.Label Lbl_SexoBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6480
               TabIndex        =   12
               Top             =   770
               Width           =   3135
            End
            Begin VB.Label Lbl_GruFam 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6480
               TabIndex        =   11
               Top             =   470
               Width           =   3135
            End
            Begin VB.Label Lbl_Par 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6480
               TabIndex        =   10
               Top             =   170
               Width           =   3135
            End
            Begin VB.Label Lbl_FecNacBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   6
               Top             =   1965
               Width           =   1095
            End
            Begin VB.Label Lbl_ApMatBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   5
               Top             =   1670
               Width           =   3375
            End
            Begin VB.Label Lbl_ApPatBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   4
               Top             =   1380
               Width           =   3375
            End
            Begin VB.Label Lbl_NomBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   2
               Top             =   800
               Width           =   3375
            End
            Begin VB.Label Lbl_PenGarBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6480
               TabIndex        =   9
               Top             =   2550
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Lbl_FecFallBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   7
               Top             =   2265
               Width           =   1095
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Ident."
               Height          =   195
               Index           =   15
               Left            =   120
               TabIndex        =   164
               Top             =   195
               Width           =   765
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Pension Garan."
               Height          =   195
               Index           =   14
               Left            =   5280
               TabIndex        =   163
               Top             =   2550
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Lbl_NumOrden 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   7680
               TabIndex        =   162
               Top             =   2280
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Grupo Familiar"
               Height          =   255
               Index           =   9
               Left            =   5280
               TabIndex        =   161
               Top             =   490
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Dº a Pensión"
               Height          =   255
               Index           =   8
               Left            =   5280
               TabIndex        =   160
               Top             =   1990
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Fec. Nac."
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   159
               Top             =   1995
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Sit. Invalidez"
               Height          =   255
               Index           =   6
               Left            =   5280
               TabIndex        =   158
               Top             =   1090
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Parentesco"
               Height          =   255
               Index           =   5
               Left            =   5280
               TabIndex        =   157
               Top             =   190
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Sexo"
               Height          =   255
               Index           =   4
               Left            =   5280
               TabIndex        =   156
               Top             =   790
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Ap. Materno"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   155
               Top             =   1670
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Ap. Paterno"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   154
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "1er. Nombre"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   153
               Top             =   800
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Porcentaje"
               Height          =   255
               Index           =   7
               Left            =   5280
               TabIndex        =   152
               Top             =   2295
               Width           =   855
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Fec. Fallec."
               Height          =   195
               Index           =   10
               Left            =   120
               TabIndex        =   151
               Top             =   2265
               Width           =   825
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Pensión"
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   150
               Top             =   2550
               Width           =   855
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Fec. Invalidez"
               Height          =   255
               Index           =   12
               Left            =   5280
               TabIndex        =   149
               Top             =   1390
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Causal Invalidez"
               Height          =   255
               Index           =   13
               Left            =   5280
               TabIndex        =   148
               Top             =   1690
               Width           =   1215
            End
            Begin VB.Label Lbl_RutBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   0
               Top             =   165
               Width           =   2775
            End
            Begin VB.Label Lbl_DgvBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   1
               Top             =   480
               Width           =   2775
            End
            Begin VB.Label Lbl_Porcentaje 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6480
               TabIndex        =   17
               Top             =   2265
               Width           =   1095
            End
            Begin VB.Label Lbl_PensionBen 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   8
               Top             =   2550
               Width           =   1095
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Número Ident."
               Height          =   195
               Index           =   17
               Left            =   120
               TabIndex        =   147
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Lbl_NomBenSeg 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   3
               Top             =   1080
               Width           =   3375
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "2do. Nombre"
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   146
               Top             =   1080
               Width           =   915
            End
         End
         Begin MSFlexGridLib.MSFlexGrid Msf_GriAseg 
            Height          =   1455
            Left            =   120
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   2566
            _Version        =   393216
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
      End
      Begin VB.Frame Fra_Afiliado 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   73
         Top             =   360
         Width           =   10215
         Begin VB.CheckBox chkExoneradoEssalud 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5415
            TabIndex        =   215
            Top             =   4470
            Width           =   255
         End
         Begin VB.Frame Fra_PagPension 
            Caption         =   "Forma de Pago de Pensión"
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
            Height          =   1815
            Left            =   5400
            TabIndex        =   92
            Top             =   2640
            Width           =   4455
            Begin VB.Label Lbl_NumCta 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   960
               TabIndex        =   64
               Top             =   1425
               Width           =   3255
            End
            Begin VB.Label Lbl_Bco 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   960
               TabIndex        =   63
               Top             =   1125
               Width           =   3255
            End
            Begin VB.Label Lbl_TipCta 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   960
               TabIndex        =   62
               Top             =   825
               Width           =   3255
            End
            Begin VB.Label Lbl_Suc 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   960
               TabIndex        =   61
               Top             =   525
               Width           =   3255
            End
            Begin VB.Label Lbl_ViaPago 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   960
               TabIndex        =   60
               Top             =   225
               Width           =   3255
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Sucursal"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   97
               Top             =   540
               Width           =   810
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "N°Cuenta"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   96
               Top             =   1440
               Width           =   795
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Banco"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   95
               Top             =   1140
               Width           =   825
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Vía Pago"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   94
               Top             =   240
               Width           =   810
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Tipo Cta."
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   93
               Top             =   840
               Width           =   825
            End
         End
         Begin VB.Label Label3 
            Caption         =   "Exonerado del descuento de ESSALUD"
            Height          =   225
            Left            =   5700
            TabIndex        =   216
            Top             =   4500
            Width           =   4080
         End
         Begin VB.Label lbl_asesor 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   214
            Top             =   4440
            Width           =   3660
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Asesor"
            Height          =   195
            Index           =   25
            Left            =   165
            TabIndex        =   213
            Top             =   4450
            Width           =   480
         End
         Begin VB.Label lbl_teleben2 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8760
            TabIndex        =   209
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono 2"
            Height          =   255
            Left            =   7800
            TabIndex        =   208
            Top             =   1395
            Width           =   855
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Representante"
            Height          =   195
            Index           =   24
            Left            =   165
            TabIndex        =   203
            Top             =   4155
            Width           =   1050
         End
         Begin VB.Label Lbl_Representante 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   202
            Top             =   4110
            Width           =   3660
         End
         Begin VB.Label Lbl_Nacionalidad 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   195
            Top             =   1950
            Width           =   1815
         End
         Begin VB.Label Lbl_NumLiq 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   194
            Top             =   2235
            Width           =   3255
         End
         Begin VB.Label Lbl_Correo 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   193
            Top             =   1665
            Width           =   3255
         End
         Begin VB.Label Lbl_Fono 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   192
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Nº Liquidación"
            Height          =   255
            Index           =   23
            Left            =   5400
            TabIndex        =   191
            Top             =   2235
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Nacionalidad"
            Height          =   255
            Index           =   22
            Left            =   5400
            TabIndex        =   190
            Top             =   1950
            Width           =   1215
         End
         Begin VB.Label Lbl_Provincia 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   189
            Top             =   790
            Width           =   3255
         End
         Begin VB.Label Lbl_Region 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   188
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label Lbl_Dir 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   187
            Top             =   200
            Width           =   3255
         End
         Begin VB.Label Lbl_Comuna 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   186
            Top             =   500
            Width           =   3255
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "1er. Nombre"
            Height          =   195
            Index           =   21
            Left            =   170
            TabIndex        =   137
            Top             =   835
            Width           =   870
         End
         Begin VB.Label Lbl_NomAfiSeg 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   48
            Top             =   1080
            Width           =   3660
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Nº Identificación"
            Height          =   195
            Index           =   20
            Left            =   170
            TabIndex        =   136
            Top             =   545
            Width           =   1170
         End
         Begin VB.Label Lbl_Asegurados 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4680
            TabIndex        =   45
            Top             =   200
            Width           =   375
         End
         Begin VB.Label Lbl_Salud 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   59
            Top             =   3770
            Width           =   3660
         End
         Begin VB.Label Lbl_Afp 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3435
            TabIndex        =   58
            Top             =   3465
            Width           =   1785
         End
         Begin VB.Label Lbl_EstCivil 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   57
            Top             =   3465
            Width           =   1350
         End
         Begin VB.Label Lbl_CauInv 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   56
            Top             =   3165
            Width           =   3660
         End
         Begin VB.Label Lbl_FecInv 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   55
            Top             =   2865
            Width           =   1095
         End
         Begin VB.Label Lbl_TipPen 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   54
            Top             =   2565
            Width           =   3660
         End
         Begin VB.Label Lbl_FecFall 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3840
            TabIndex        =   53
            Top             =   2265
            Width           =   1095
         End
         Begin VB.Label Lbl_FecNac 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   52
            Top             =   2265
            Width           =   1095
         End
         Begin VB.Label Lbl_SexoAfi 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   51
            Top             =   1965
            Width           =   3660
         End
         Begin VB.Label Lbl_ApPatAfi 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   49
            Top             =   1380
            Width           =   3660
         End
         Begin VB.Label Lbl_ApMatAfi 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   50
            Top             =   1665
            Width           =   3660
         End
         Begin VB.Label Lbl_NomAfi 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   47
            Top             =   795
            Width           =   3660
         End
         Begin VB.Label Lbl_DgvAfi 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   46
            Top             =   500
            Width           =   2295
         End
         Begin VB.Label Lbl_RutAfi 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   44
            Top             =   200
            Width           =   2295
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Identificación"
            Height          =   195
            Index           =   0
            Left            =   170
            TabIndex        =   117
            Top             =   245
            Width           =   1305
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "2do. Nombre"
            Height          =   195
            Index           =   1
            Left            =   170
            TabIndex        =   116
            Top             =   1125
            Width           =   915
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Index           =   12
            Left            =   5400
            TabIndex        =   115
            Top             =   1125
            Width           =   480
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Provincia"
            Height          =   195
            Index           =   11
            Left            =   5400
            TabIndex        =   114
            Top             =   835
            Width           =   660
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio"
            Height          =   195
            Index           =   10
            Left            =   5400
            TabIndex        =   113
            Top             =   245
            Width           =   630
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   9
            Left            =   5400
            TabIndex        =   112
            Top             =   1395
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Est. Civil"
            Height          =   195
            Index           =   8
            Left            =   170
            TabIndex        =   111
            Top             =   3510
            Width           =   600
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Nac."
            Height          =   195
            Index           =   7
            Left            =   170
            TabIndex        =   110
            Top             =   2310
            Width           =   840
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Sexo"
            Height          =   195
            Index           =   4
            Left            =   170
            TabIndex        =   109
            Top             =   2010
            Width           =   360
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Ap. Materno"
            Height          =   195
            Index           =   3
            Left            =   170
            TabIndex        =   108
            Top             =   1710
            Width           =   870
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Ap. Paterno"
            Height          =   195
            Index           =   2
            Left            =   170
            TabIndex        =   107
            Top             =   1425
            Width           =   840
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Inst. Salud"
            Height          =   195
            Index           =   6
            Left            =   170
            TabIndex        =   106
            Top             =   3820
            Width           =   750
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Nº Benef."
            Height          =   255
            Index           =   13
            Left            =   3960
            TabIndex        =   105
            Top             =   210
            Width           =   735
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "AFP"
            Height          =   195
            Index           =   5
            Left            =   3050
            TabIndex        =   104
            Top             =   3510
            Width           =   300
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Index           =   15
            Left            =   5400
            TabIndex        =   103
            Top             =   545
            Width           =   1005
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Correo"
            Height          =   255
            Index           =   14
            Left            =   5400
            TabIndex        =   102
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Fecha Fallec."
            Height          =   255
            Index           =   16
            Left            =   2760
            TabIndex        =   101
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Pensión"
            Height          =   195
            Index           =   17
            Left            =   170
            TabIndex        =   100
            Top             =   2610
            Width           =   930
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Fec. Invalidez"
            Height          =   195
            Index           =   18
            Left            =   170
            TabIndex        =   99
            Top             =   2910
            Width           =   990
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Causal Invalidez"
            Height          =   195
            Index           =   19
            Left            =   170
            TabIndex        =   98
            Top             =   3210
            Width           =   1155
         End
      End
      Begin VB.Frame Fra_Calculo 
         Height          =   4815
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Width           =   10215
         Begin VB.Frame Fra_SumaBono 
            Height          =   615
            Left            =   120
            TabIndex        =   89
            Top             =   4080
            Width           =   9975
            Begin VB.Label Lbl_SumaBono 
               Alignment       =   2  'Center
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   4680
               TabIndex        =   199
               Top             =   180
               Width           =   255
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "AA."
               Height          =   195
               Index           =   5
               Left            =   4965
               TabIndex        =   198
               Top             =   210
               Width           =   255
            End
            Begin VB.Label Lbl_MonedaFon 
               Alignment       =   1  'Right Justify
               Caption         =   "(TM)"
               Height          =   255
               Index           =   3
               Left            =   5235
               TabIndex        =   197
               Top             =   210
               Width           =   375
            End
            Begin VB.Label Lbl_ApoAdi 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5640
               TabIndex        =   196
               Top             =   210
               Width           =   1455
            End
            Begin VB.Label Lbl_CtaInd 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   720
               TabIndex        =   41
               Top             =   210
               Width           =   1455
            End
            Begin VB.Label Lbl_BonoAct 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3240
               TabIndex        =   42
               Top             =   210
               Width           =   1455
            End
            Begin VB.Label Lbl_PriUnica 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   8400
               TabIndex        =   43
               Top             =   210
               Width           =   1455
            End
            Begin VB.Label Lbl_MonedaFon 
               Alignment       =   1  'Right Justify
               Caption         =   "(TM)"
               Height          =   255
               Index           =   2
               Left            =   7995
               TabIndex        =   174
               Top             =   210
               Width           =   375
            End
            Begin VB.Label Lbl_MonedaFon 
               Alignment       =   1  'Right Justify
               Caption         =   "(TM)"
               Height          =   255
               Index           =   1
               Left            =   2835
               TabIndex        =   173
               Top             =   225
               Width           =   375
            End
            Begin VB.Label Lbl_MonedaFon 
               Alignment       =   1  'Right Justify
               Caption         =   "(TM)"
               Height          =   255
               Index           =   0
               Left            =   315
               TabIndex        =   172
               Top             =   225
               Width           =   375
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "BR."
               Height          =   195
               Index           =   0
               Left            =   2565
               TabIndex        =   143
               Top             =   225
               Width           =   270
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "CI."
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   142
               Top             =   225
               Width           =   195
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "P. Unica"
               Height          =   195
               Index           =   2
               Left            =   7360
               TabIndex        =   141
               Top             =   210
               Width           =   615
            End
            Begin VB.Label Lbl_SumaBono 
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   7115
               TabIndex        =   91
               Top             =   210
               Width           =   255
            End
            Begin VB.Label Lbl_SumaBono 
               Alignment       =   2  'Center
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   2280
               TabIndex        =   90
               Top             =   180
               Width           =   255
            End
         End
         Begin VB.Frame Fra_DatCal 
            Height          =   3975
            Left            =   240
            TabIndex        =   81
            Top             =   240
            Width           =   9975
            Begin VB.Label lblTC 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   9120
               TabIndex        =   212
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "TC"
               Height          =   255
               Left            =   8040
               TabIndex        =   211
               Top             =   525
               Width           =   375
            End
            Begin VB.Label Lbl_IndCob 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   9120
               TabIndex        =   28
               Top             =   185
               Width           =   735
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Reajuste Trimestral"
               Height          =   195
               Index           =   27
               Left            =   240
               TabIndex        =   207
               Top             =   1680
               Width           =   1350
            End
            Begin VB.Label Lbl_ReajusteTipo 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3360
               TabIndex        =   206
               Top             =   1680
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Lbl_ReajusteValor 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   205
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Lbl_ReajusteMoneda 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   204
               Top             =   1380
               Width           =   2775
            End
            Begin VB.Label Lbl_FecTraPrima 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   201
               Top             =   3480
               Width           =   975
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Fecha Traspaso Prima"
               Height          =   255
               Index           =   25
               Left            =   4800
               TabIndex        =   200
               Top             =   3480
               Width           =   1815
            End
            Begin VB.Label Lbl_FecIniPago 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   19
               Top             =   780
               Width           =   2775
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Derecho Crecer"
               Height          =   195
               Index           =   2
               Left            =   4800
               TabIndex        =   185
               Top             =   525
               Width           =   1125
            End
            Begin VB.Label Lbl_DerCre 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   184
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Lbl_FecIncorpora 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   183
               Top             =   480
               Width           =   2775
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Fec. Incorp. a la Póliza"
               Height          =   255
               Index           =   38
               Left            =   240
               TabIndex        =   182
               Top             =   495
               Width           =   1695
            End
            Begin VB.Label Lbl_MtoPenGar 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   181
               Top             =   2280
               Width           =   1215
            End
            Begin VB.Label Lbl_MtoPrimaUniSim 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Left            =   8520
               TabIndex        =   180
               Top             =   2230
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Ind. Cobertura"
               Height          =   255
               Index           =   24
               Left            =   8040
               TabIndex        =   179
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Fec. Devengue"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   178
               Top             =   200
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Cobertura Cónyuge"
               Height          =   195
               Index           =   9
               Left            =   4800
               TabIndex        =   177
               Top             =   230
               Width           =   1365
            End
            Begin VB.Label Lbl_FecDev 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   18
               Top             =   185
               Width           =   975
            End
            Begin VB.Label Lbl_FacPenElla 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   29
               Top             =   185
               Width           =   975
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   0
               Left            =   7650
               TabIndex        =   176
               Top             =   200
               Width           =   165
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Fec. Inicio 1er. Pago"
               Height          =   255
               Index           =   39
               Left            =   240
               TabIndex        =   175
               Top             =   795
               Width           =   1695
            End
            Begin VB.Label Lbl_BenSocial 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   9015
               TabIndex        =   171
               Top             =   1080
               Width           =   720
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Ben. Social"
               Height          =   255
               Index           =   23
               Left            =   8160
               TabIndex        =   170
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Lbl_DerGra 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   30
               Top             =   780
               Width           =   975
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Gratificación"
               Height          =   195
               Index           =   19
               Left            =   4800
               TabIndex        =   167
               Top             =   825
               Width           =   885
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Moneda"
               Height          =   195
               Index           =   18
               Left            =   240
               TabIndex        =   166
               Top             =   1380
               Width           =   945
            End
            Begin VB.Label Lbl_Moneda 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   0
               Left            =   3360
               TabIndex        =   21
               Top             =   120
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Ident. Intermediario"
               Height          =   195
               Index           =   35
               Left            =   4800
               TabIndex        =   140
               Top             =   1380
               Width           =   1710
            End
            Begin VB.Label Lbl_Nombre 
               AutoSize        =   -1  'True
               Caption         =   "Nº Ident. Intermediario"
               Height          =   195
               Index           =   5
               Left            =   4800
               TabIndex        =   139
               Top             =   1680
               Width           =   1575
            End
            Begin VB.Label Lbl_CUSPP 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   20
               Top             =   1080
               Width           =   2775
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "CUSPP"
               Height          =   195
               Index           =   33
               Left            =   240
               TabIndex        =   138
               Top             =   1080
               Width           =   540
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Años Dif."
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   134
               Top             =   2280
               Width           =   735
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Meses Gar."
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   133
               Top             =   2880
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Modalidad"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   132
               Top             =   2580
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tipo de Renta"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   131
               Top             =   1980
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Prc. Renta Temporal"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   130
               Top             =   3480
               Width           =   1575
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Rentabilidad AFP"
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   129
               Top             =   3180
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   20
               Left            =   3000
               TabIndex        =   128
               Top             =   3180
               Width           =   255
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   21
               Left            =   3000
               TabIndex        =   127
               Top             =   3480
               Width           =   255
            End
            Begin VB.Label Lbl_TipoRenta 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   22
               Top             =   1980
               Width           =   2775
            End
            Begin VB.Label Lbl_AnnosDif 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   23
               Top             =   2280
               Width           =   735
            End
            Begin VB.Label Lbl_Alter 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   24
               Top             =   2580
               Width           =   2775
            End
            Begin VB.Label Lbl_MesesGar 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   25
               Top             =   2880
               Width           =   735
            End
            Begin VB.Label Lbl_RentaAFP 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   26
               Top             =   3180
               Width           =   975
            End
            Begin VB.Label Lbl_PrcRtaTmp 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   27
               Top             =   3480
               Width           =   975
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Comisión"
               Height          =   255
               Index           =   16
               Left            =   4800
               TabIndex        =   126
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   22
               Left            =   7650
               TabIndex        =   125
               Top             =   1080
               Width           =   165
            End
            Begin VB.Label Lbl_RutCorr 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   32
               Top             =   1380
               Width           =   3120
            End
            Begin VB.Label Lbl_DgvCorr 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   33
               Top             =   1680
               Width           =   3120
            End
            Begin VB.Label Lbl_MtoPrimaUniDif 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Left            =   8520
               TabIndex        =   39
               Top             =   2810
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Lbl_ComInt 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   31
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Mto. Prima Uni Sim"
               Height          =   255
               Index           =   60
               Left            =   8160
               TabIndex        =   124
               Top             =   1955
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Mto. Prima Uni Dif"
               Height          =   255
               Index           =   61
               Left            =   8160
               TabIndex        =   123
               Top             =   2535
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Lbl_MtoPension 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   34
               Top             =   1980
               Width           =   1215
            End
            Begin VB.Label Lbl_TasaPG 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   9120
               TabIndex        =   38
               Top             =   3390
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Lbl_TasaTir 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   37
               Top             =   3180
               Width           =   735
            End
            Begin VB.Label Lbl_TasaVta 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   36
               Top             =   2880
               Width           =   735
            End
            Begin VB.Label Lbl_TasaCE 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6615
               TabIndex        =   35
               Top             =   2580
               Width           =   735
            End
            Begin VB.Label Lbl_PrcFam 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Left            =   9240
               TabIndex        =   40
               Top             =   3105
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Factor CNU"
               Height          =   255
               Index           =   7
               Left            =   8160
               TabIndex        =   88
               Top             =   3120
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tasa Cto. Equiv."
               Height          =   255
               Index           =   8
               Left            =   4800
               TabIndex        =   87
               Top             =   2580
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tasa Venta"
               Height          =   255
               Index           =   11
               Left            =   4800
               TabIndex        =   86
               Top             =   2880
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tasa TIR"
               Height          =   255
               Index           =   13
               Left            =   4800
               TabIndex        =   85
               Top             =   3180
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Mto. Pensión "
               Height          =   255
               Index           =   14
               Left            =   4800
               TabIndex        =   84
               Top             =   1980
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Mto. Pensión Gar. "
               Height          =   255
               Index           =   15
               Left            =   4800
               TabIndex        =   83
               Top             =   2280
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tasa Int. Per. Gar."
               Height          =   255
               Index           =   17
               Left            =   7920
               TabIndex        =   82
               Top             =   3405
               Visible         =   0   'False
               Width           =   1335
            End
         End
      End
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
      TabIndex        =   119
      Top             =   0
      Width           =   8805
      Begin VB.TextBox Txt_Endoso 
         Height          =   285
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   66
         Text            =   "1"
         Top             =   360
         Width           =   585
      End
      Begin VB.TextBox Txt_NumPol 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   65
         Top             =   360
         Width           =   2625
      End
      Begin VB.CommandButton Cmd_BuscarPol 
         Height          =   495
         Left            =   6600
         Picture         =   "Frm_CalConsulta.frx":157C
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Buscar Póliza"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Póliza"
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   122
         Top             =   375
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Num. Endoso"
         Height          =   195
         Index           =   6
         Left            =   4800
         TabIndex        =   121
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
         TabIndex        =   120
         Top             =   0
         Width           =   765
      End
   End
End
Attribute VB_Name = "Frm_CalConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlElemento As String, vlNumEnd As Integer
Dim vlDir As String, vlApoderado As String
Dim vlCargo As String, vlNomAfi As String
Dim vlNumPol As String, vlRut As String
Dim vlSql As String, vlAnnosDif As String
Dim vlFechaVig As String, vlSucursal As String
Dim vlParentesco As String, vlFecNac As String
Dim vlTipoIden As String
Dim vlCodDir As String, vlTipVia As String
Dim vlNomVia As String, vlNumDmc As String
Dim vlIntDmc As String, vlTipZon As String
Dim vlNomZon As String, vlReferencia As String

'Declaración de Constantes para la Moneda de la Renta
Const clMonedaModalidad As Integer = 0
Const clMonedaBenPen    As Integer = 1
Const clMonedaBenPenGar As Integer = 2

Const clMtoCtaIndFon    As Integer = 0
Const clMtoBonoFon      As Integer = 1
Const clMtoPriUniFon    As Integer = 2
Const clMtoApoAdiFon    As Integer = 3

Dim vlRepresentante As String, vlDocum As String
Dim vlCobertura As String
Private objRep As New ClsReporte

'------------------------------------------------
'BUSCA LA GLOSA SEGUN EL CODGO DE LA TABLA TABCOD
'------------------------------------------------
Function flBuscaCodGlosa(icodtabla As String, icod As String)
On Error GoTo Err_BusDat
    vlElemento = ""
    flBuscaCodGlosa = False
    vgSql = ""
    vgSql = "SELECT gls_elemento FROM ma_tpar_tabcod WHERE "
    vgSql = vgSql & "cod_tabla= '" & icodtabla & "' AND "
    vgSql = vgSql & "cod_elemento= '" & icod & "'"
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        vlElemento = vgRs4!gls_elemento
        flBuscaCodGlosa = True
    End If
    vgRs4.Close
Exit Function
Err_BusDat:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'---------------------------------------
'BUSCA LA GLOSA DE LA CAUSA DE INVALIDEZ
'---------------------------------------
Function flBuscaGlosaCauInv(icod As String)
On Error GoTo Err_BusDat
    vlElemento = ""
    flBuscaCodGlosaCauInv = False
    vgSql = ""
    vgSql = "SELECT gls_patologia FROM ma_tpar_patologia WHERE "
    vgSql = vgSql & "cod_patologia= '" & icod & "'"
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        vlElemento = vgRs4!gls_patologia
        flBuscaCodGlosaCauInv = True
    End If
    vgRs4.Close
Exit Function
Err_BusDat:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'-----------------------------
'BUSCA LA GLOSA DE LA SUCURSAL
'-----------------------------
Function flBuscaGlosaSuc(icod As String)
On Error GoTo Err_BusDat
    vlElemento = ""
    flBuscaGlosaSuc = False
    vgSql = ""
    vgSql = "SELECT gls_sucursal FROM ma_tpar_sucursal WHERE "
    vgSql = vgSql & "cod_sucursal= '" & icod & "'"
    vgSql = vgSql & " AND cod_tipo = '" & vgTipoSucursal & "' "
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        vlElemento = vgRs4!gls_sucursal
        flBuscaGlosaSuc = True
    End If
    vgRs4.Close
Exit Function
Err_BusDat:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'---------------------------
'BUSCA LA GLOSA DE LA COMUNA
'---------------------------
Function flBuscaGlosaComuna(icod As String)
On Error GoTo Err_BusDat
    vlElemento = ""
    flBuscaGlosaComuna = False
    vgSql = ""
    vgSql = "SELECt gls_comuna FROM ma_tpar_comuna WHERE "
    vgSql = vgSql & "cod_direccion= '" & icod & "'"
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        vlElemento = vgRs4!gls_comuna
        flBuscaGlosaComuna = True
    End If
    vgRs4.Close
Exit Function
Err_BusDat:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'-------------------------------------------------------
'FUNCION BUSCA DATOS DE POLIZA PARA CARGAR EN FORMULARIO
'-------------------------------------------------------
Function flBuscaPoliza(iNumPol As String, inumend As Integer)
On Error GoTo Err_BuscaPol

    SSTab_Poliza.Tab = 0
    vlSql = ""
    vlSql = "SELECT p.num_poliza, p.num_endoso,b.gls_nomben, b.gls_patben, b.gls_matben "
    vlSql = vlSql & "FROM pd_tmae_poliza p,pd_tmae_polben b WHERE "
    vlSql = vlSql & "p.num_poliza = b.num_poliza "
    vlSql = vlSql & "AND p.num_endoso = b.num_endoso "
    vlSql = vlSql & "AND b.cod_par = '99' "
    vlSql = vlSql & "AND p.num_poliza = '" & iNumPol & "' "
    vlSql = vlSql & "AND p.num_endoso= " & inumend & ""
    Set vgRs = vgConexionBD.Execute(vlSql)
    
    If (vgRs.EOF) Then
        MsgBox "La Póliza No se encuentra en la BD", vbCritical, "Póliza No Encontrada"
        Txt_NumPol.SetFocus
        Exit Function
    End If
    
    Call flCargaCarpAfilPol(iNumPol, inumend)
    Call flCargaCarpAfilPolRep(iNumPol, inumend)  'RVF 20090914
    Call flCargaCarpCalculo(iNumPol, inumend)
    'Call flCargaCarpBono(inumpol, inumend)
    Call flCargaCarpBenef(iNumPol, inumend)

    Msf_GriAseg.Enabled = True
    'Fra_Poliza.Enabled = False
    Cmd_Poliza.Enabled = False
    SSTab_Poliza.Enabled = True
    SSTab_Poliza.Tab = 0
    
Exit Function
Err_BuscaPol:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'--------------------------------------------------
'CARGA INFORMACION EN LA CARPETA DATOS DEL AFILIADO
'--------------------------------------------------
Function flCargaCarpAfilPol(iNumPol As String, inumend As Integer)
On Error GoTo Err_CargaAfi

    vlSql = ""
    vlSql = "SELECT p.num_poliza,p.num_endoso,p.num_cot,p.num_operacion,p.fec_vigencia,"
    vlSql = vlSql & "b.cod_tipoidenben,b.num_idenben,p.num_cargas, "
    vlSql = vlSql & "b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben, "
    vlSql = vlSql & "b.cod_sexo,b.fec_nacben,b.fec_fallben, "
    vlSql = vlSql & "p.cod_tippension,b.fec_invben,b.cod_cauinv,"
    vlSql = vlSql & "p.cod_estcivil,p.cod_afp,p.cod_isapre,p.cod_direccion,"
    vlSql = vlSql & "p.gls_direccion,p.gls_fono,"
    vlSql = vlSql & "p.gls_correo,p.cod_viapago,p.cod_sucursal,"
    vlSql = vlSql & "p.cod_tipcuenta,p.cod_banco,p.num_cuenta,i.gls_tipoidencor, "
    vlSql = vlSql & "p.gls_nacionalidad,n.cod_liquidacion, "
    vlSql = vlSql & "P.gls_telben2,c.num_idencor,c.gls_nomcor || ' ' ||  c.gls_patcor || ' ' || c.gls_matcor as gls_asesor " '---RRR 08/05/2013
    vlSql = vlSql & ",p.ind_bendes "
    vlSql = vlSql & "FROM pd_tmae_poliza p,pd_tmae_polben b,ma_tpar_tipoiden i, "
    vlSql = vlSql & "pd_tmae_polprirec n, pt_tmae_corredor c "
    vlSql = vlSql & "WHERE p.num_poliza = '" & iNumPol & "' "
    vlSql = vlSql & "AND p.num_endoso = " & inumend & " "
    vlSql = vlSql & "AND p.num_poliza = b.num_poliza "
    vlSql = vlSql & "AND p.num_endoso = b.num_endoso "
    vlSql = vlSql & "AND b.cod_tipoidenben = i.cod_tipoiden AND p.num_poliza=n.num_poliza AND p.num_idencor=c.num_idencor "
    vlSql = vlSql & "AND b.cod_par = '99' "
    
    Set vgRs = vgConexionBD.Execute(vlSql)
    
    If Not (vgRs.EOF) Then
        'datos de la cabecera
        If Not IsNull(vgRs!Num_Poliza) Then Txt_NumPol = vgRs!Num_Poliza
        If Not IsNull(vgRs!Num_Endoso) Then Txt_Endoso = vgRs!Num_Endoso
        If Not IsNull(vgRs!Num_Cot) Then Lbl_NumCot = vgRs!Num_Cot
        If Not IsNull(vgRs!Num_Operacion) Then Lbl_SolOfe = vgRs!Num_Operacion
        If Not IsNull(vgRs!Fec_Vigencia) Then Lbl_FecVig = DateSerial(Mid(vgRs!Fec_Vigencia, 1, 4), Mid(vgRs!Fec_Vigencia, 5, 2), Mid(vgRs!Fec_Vigencia, 7, 2))

        'datos de la carpeta de afiliado
        If Not IsNull(vgRs!GLS_TIPOIDENCOR) Then Lbl_RutAfi = vgRs!GLS_TIPOIDENCOR
        If Not IsNull(vgRs!Num_Cargas) Then Lbl_Asegurados = vgRs!Num_Cargas
        If Not IsNull(vgRs!Num_IdenBen) Then Lbl_DgvAfi = vgRs!Num_IdenBen
        If Not IsNull(vgRs!Gls_NomBen) Then Lbl_NomAfi = vgRs!Gls_NomBen
        If Not IsNull(vgRs!Gls_NomSegBen) Then Lbl_NomAfiSeg = vgRs!Gls_NomSegBen Else Lbl_NomAfiSeg = ""
        If Not IsNull(vgRs!Gls_PatBen) Then Lbl_ApPatAfi = vgRs!Gls_PatBen
        If Not IsNull(vgRs!Gls_MatBen) Then Lbl_ApMatAfi = vgRs!Gls_MatBen
        If Not IsNull(vgRs!Fec_NacBen) Then Lbl_FecNac = DateSerial(Mid(vgRs!Fec_NacBen, 1, 4), Mid(vgRs!Fec_NacBen, 5, 2), Mid(vgRs!Fec_NacBen, 7, 2))
        If Not IsNull(vgRs!Fec_FallBen) Then Lbl_FecFall = DateSerial(Mid(vgRs!Fec_FallBen, 1, 4), Mid(vgRs!Fec_FallBen, 5, 2), Mid(vgRs!Fec_FallBen, 7, 2))
        If Not IsNull(vgRs!Fec_InvBen) Then Lbl_FecInv = DateSerial(Mid(vgRs!Fec_InvBen, 1, 4), Mid(vgRs!Fec_InvBen, 5, 2), Mid(vgRs!Fec_InvBen, 7, 2))
        If Not IsNull(vgRs!Gls_Direccion) Then Lbl_Dir = vgRs!Gls_Direccion
        If Not IsNull(vgRs!GLS_FONO) Then Lbl_Fono = vgRs!GLS_FONO
        If Not IsNull(vgRs!GLS_CORREO) Then Lbl_Correo = vgRs!GLS_CORREO
        If Not IsNull(vgRs!Num_Cuenta) Then Lbl_NumCta = vgRs!Num_Cuenta
        If Not IsNull(vgRs!gls_nacionalidad) Then Lbl_Nacionalidad = UCase(vgRs!gls_nacionalidad)
        If Not IsNull(vgRs!gls_telben2) Then lbl_teleben2 = UCase(vgRs!gls_telben2) Else lbl_teleben2 = "" '--RRR 08/05/2013
        '
        If Not IsNull(vgRs!gls_asesor) Then lbl_asesor = vgRs!Num_IdenCor & " - " & UCase(vgRs!gls_asesor) Else lbl_asesor = ""
        Lbl_NumLiq = fgBuscarNroLiquidacion(Txt_NumPol)
                        
        Call flBuscaCodGlosa(vgCodTabla_Sexo, (vgRs!Cod_Sexo))
        Lbl_SexoAfi = (vgRs!Cod_Sexo) + " - " + vlElemento
        Call flBuscaCodGlosa(vgCodTabla_TipPen, (vgRs!Cod_TipPension))
        Lbl_TipPen = (vgRs!Cod_TipPension) + " - " + vlElemento
        Call flBuscaGlosaCauInv((vgRs!Cod_CauInv))
        Lbl_CauInv = (vgRs!Cod_CauInv) + " - " + vlElemento
        Call flBuscaCodGlosa(vgCodTabla_EstCiv, (vgRs!cod_estcivil))
        Lbl_EstCivil = (vgRs!cod_estcivil) + " - " + vlElemento
        Call flBuscaCodGlosa(vgCodTabla_AFP, (vgRs!Cod_AFP))
        Lbl_Afp = (vgRs!Cod_AFP) + " - " + vlElemento
        Call flBuscaCodGlosa(vgCodTabla_InsSal, (vgRs!cod_isapre))
        Lbl_Salud = (vgRs!cod_isapre) + " - " + vlElemento
        Call flBuscaCodGlosa(vgCodTabla_ViaPago, (vgRs!Cod_ViaPago))
        Lbl_ViaPago = (vgRs!Cod_ViaPago) + " - " + vlElemento
        If (vgRs!Cod_ViaPago = "04") Then
            vgTipoSucursal = cgTipoSucursalAfp
        Else
            vgTipoSucursal = cgTipoSucursalSuc
        End If
        Call flBuscaGlosaSuc(vgRs!Cod_Sucursal)
        Lbl_Suc = (vgRs!Cod_Sucursal) + " - " + vlElemento
        Call flBuscaCodGlosa(vgCodTabla_TipCta, (vgRs!Cod_TipCuenta))
        Lbl_TipCta = (vgRs!Cod_TipCuenta) + " - " + vlElemento
        Call flBuscaCodGlosa(vgCodTabla_Bco, (vgRs!Cod_Banco))
        Lbl_Bco = (vgRs!Cod_Banco) + " - " + vlElemento
        
        vlDir = vgRs!Cod_Direccion
        Call fgBuscarNombreComunaProvinciaRegion(vlDir)
        Lbl_Provincia = vgNombreProvincia
        Lbl_Region = vgNombreRegion
'        Call flBuscaGlosaComuna(vlDir)
'        Lbl_Comuna = vlElemento
        Lbl_Comuna = vgNombreComuna
        
        If IsNull(vgRs!ind_bendes) Then
            chkExoneradoEssalud.Value = 0
        Else
            If (vgRs!ind_bendes = "N") Then
                chkExoneradoEssalud.Value = 0
            Else
                chkExoneradoEssalud.Value = 1
            End If
        End If
        
        
    End If

Exit Function
Err_CargaAfi:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function
'---------------------------------------
'CARGA INFORMACION EN LA CARPETA CALCULO
'---------------------------------------
Function flCargaCarpCalculo(iNumPol As String, inumend As Integer)
On Error GoTo Err_CarCalculo

    vlSql = ""
    vlSql = "SELECT p.fec_dev,p.fec_acepta,p.fec_pripago,p.cod_cuspp,"
    vlSql = vlSql & "p.cod_moneda,'" & vgMonedaCodOfi & "' as cod_monedafon,p.cod_tipren,p.num_mesdif,p.cod_modalidad,"
    vlSql = vlSql & "p.num_mesgar,p.prc_rentaafp,p.prc_rentatmp,p.ind_cob,"
    vlSql = vlSql & "p.mto_ctaind as mto_ctaindfon,p.cod_cobercon as prc_facpenella,p.cod_dercre,"
    vlSql = vlSql & "p.cod_dergra,p.prc_corcom,p.prc_corcomreal,p.cod_bensocial,"
    vlSql = vlSql & "p.cod_tipoidencor,p.num_idencor,p.mto_pension,"
    vlSql = vlSql & "p.mto_pensiongar,p.mto_priunisim,p.prc_tasace,"
    vlSql = vlSql & "p.prc_tasavta,p.mto_priunidif,p.prc_tasatir,"
    vlSql = vlSql & "p.mto_cnu,p.prc_tasapergar,p.mto_bono as mto_bonofon,p.mto_priuni as mto_priunifon "
    vlSql = vlSql & ",p.mto_apoadi as mto_apoadifon "
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & ",p.cod_tipreajuste,p.mto_valreajustetri,p.mto_valreajustemen "
    vlSql = vlSql & ",mtr.cod_scomp as cod_montipreaju,mtr.gls_descripcion as gls_montipreaju, mto_valmoneda "
'F--- ABV 05/02/2011 ---
    vlSql = vlSql & "FROM pd_tmae_poliza p,ma_tpar_tipoiden ti "
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & ",ma_tpar_monedatiporeaju mtr "
'F--- ABV 05/02/2011 ---
    vlSql = vlSql & "WHERE p.num_poliza = '" & iNumPol & "' "
    vlSql = vlSql & "AND p.num_endoso= " & inumend & " "
 'I--- ABV 05/02/2011 ---
    vlSql = vlSql & "AND p.cod_tipreajuste = mtr.cod_tipreajuste(+) AND "
    vlSql = vlSql & "p.cod_moneda = mtr.cod_moneda(+) "
'F--- ABV 05/02/2011 ---

    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not (vgRs.EOF) Then
        
        If Not IsNull(vgRs!Fec_Dev) Then Lbl_FecDev = DateSerial(Mid(vgRs!Fec_Dev, 1, 4), Mid(vgRs!Fec_Dev, 5, 2), Mid(vgRs!Fec_Dev, 7, 2))
        If Not IsNull(vgRs!Fec_Acepta) Then Lbl_FecIncorpora = DateSerial(Mid(vgRs!Fec_Acepta, 1, 4), Mid(vgRs!Fec_Acepta, 5, 2), Mid(vgRs!Fec_Acepta, 7, 2))
        If Not IsNull(vgRs!fec_pripago) Then Lbl_FecIniPago = DateSerial(Mid(vgRs!fec_pripago, 1, 4), Mid(vgRs!fec_pripago, 5, 2), Mid(vgRs!fec_pripago, 7, 2))
        If Not IsNull(vgRs!Cod_Cuspp) Then Lbl_CUSPP = vgRs!Cod_Cuspp
        'Tipo de Moneda de la Modalidad
        Lbl_Moneda(clMonedaBenPen) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_Moneda)
        Lbl_Moneda(clMonedaBenPenGar) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_Moneda)
        'Tipo de Moneda del Fondo
        Lbl_MonedaFon(clMtoCtaIndFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        Lbl_MonedaFon(clMtoBonoFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        Lbl_MonedaFon(clMtoPriUniFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        Lbl_MonedaFon(clMtoApoAdiFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        
'I--- ABV 05/02/2011 ---
        'Tipo de Reajuste
        vlElemento = IIf(IsNull(vgRs!cod_montipreaju), "", Trim(vgRs!cod_montipreaju))
        Lbl_ReajusteMoneda = vlElemento + " - " + IIf(IsNull(vgRs!gls_montipreaju), "", Trim(vgRs!gls_montipreaju))
        
        If Not IsNull(vgRs!Cod_TipReajuste) Then
            vlElemento = fgBuscarGlosaElemento(vgCodTabla_TipReajuste, Trim(vgRs!Cod_TipReajuste))
            Lbl_ReajusteTipo.Caption = Trim(vgRs!Cod_TipReajuste) + " - " + vlElemento
            Lbl_ReajusteValor.Caption = Format(IIf(IsNull(vgRs!Mto_ValReajusteTri), 0, vgRs!Mto_ValReajusteTri), "#0.00000000")
        Else
            Lbl_ReajusteTipo.Caption = ""
            Lbl_ReajusteValor.Caption = ""
        End If
'F--- ABV 05/02/2011 ---
        
        If Not IsNull(vgRs!Num_MesDif) Then Lbl_AnnosDif = Trim(vgRs!Num_MesDif / 12)
        If Not IsNull(vgRs!Num_MesGar) Then Lbl_MesesGar = vgRs!Num_MesGar
        If Not IsNull(vgRs!Prc_RentaAFP) Then Lbl_RentaAFP = Format(vgRs!Prc_RentaAFP, "#0.00")
        If Not IsNull(vgRs!Prc_RentaTMP) Then Lbl_PrcRtaTmp = Format(vgRs!Prc_RentaTMP, "#0.00")
        If Not IsNull(vgRs!Ind_Cob) Then
            If vgRs!Ind_Cob = "S" Then Lbl_IndCob = cgIndicadorSi Else Lbl_IndCob = cgIndicadorNo
        End If
        If Not IsNull(vgRs!Mto_CtaIndFon) Then Lbl_CtaInd = Format(vgRs!Mto_CtaIndFon, "#,#0.00")
'        If Not IsNull(vgRs!prc_facpenella) Then Lbl_FacPenElla = Format(vgRs!prc_facpenella, "#0.00")
        If Not IsNull(vgRs!Prc_FacPenElla) Then Lbl_FacPenElla = vgRs!Prc_FacPenElla
        If Not IsNull(vgRs!Cod_DerCre) Then
            If vgRs!Cod_DerCre = "S" Then Lbl_DerCre = cgIndicadorSi Else Lbl_DerCre = cgIndicadorNo
        End If
        If Not IsNull(vgRs!Cod_DerGra) Then
            If vgRs!Cod_DerGra = "S" Then Lbl_DerGra = cgIndicadorSi Else Lbl_DerGra = cgIndicadorNo
        End If
'I--- ABV 20/08/2007 ---
'        If Not IsNull(vgRs!Prc_CorCom) Then Lbl_ComInt = Format(vgRs!Prc_CorCom, "#0.00")
        If Not IsNull(vgRs!Prc_CorComReal) Then Lbl_ComInt = Format(vgRs!Prc_CorComReal, "#0.00")
'F--- ABV 20/08/2007 ---
        If Not IsNull(vgRs!Cod_BenSocial) Then
            If vgRs!Cod_BenSocial = "S" Then Lbl_BenSocial = cgIndicadorSi Else Lbl_BenSocial = cgIndicadorNo
        End If
        'Buscar la Identificación del Intermediario
        If Not IsNull(vgRs!cod_tipoidencor) Then
            Lbl_RutCorr = fgBuscarNombreTipoIden(vgRs!cod_tipoidencor)
        End If
        If (Lbl_RutCorr <> "") Then
            If Not IsNull(vgRs!Num_IdenCor) Then Lbl_DgvCorr = vgRs!Num_IdenCor
        End If
        
        lblTC = Format(vgRs!Mto_ValMoneda, "#,#0.00")
        
        
        If Not IsNull(vgRs!Mto_Pension) Then Lbl_MtoPension = Format(vgRs!Mto_Pension, "#,#0.00")
        If Not IsNull(vgRs!Mto_PensionGar) Then Lbl_MtoPenGar = Format(vgRs!Mto_PensionGar, "#,#0.00")
        If Not IsNull(vgRs!Mto_PriUniSim) Then Lbl_MtoPrimaUniSim = Format(vgRs!Mto_PriUniSim, "#,#0.000000")
        If Not IsNull(vgRs!Prc_TasaCe) Then Lbl_TasaCE = Format(vgRs!Prc_TasaCe, "#,#0.00")
        If Not IsNull(vgRs!Prc_TasaVta) Then Lbl_TasaVta = Format(vgRs!Prc_TasaVta, "#,#0.00")
        If Not IsNull(vgRs!Mto_PriUniDif) Then Lbl_MtoPrimaUniDif = Format(vgRs!Mto_PriUniDif, "#,#0.000000")
        If Not IsNull(vgRs!Prc_TasaTir) Then Lbl_TasaTIR = Format(vgRs!Prc_TasaTir, "#,#0.00")
        If Not IsNull(vgRs!Mto_CNU) Then Lbl_PrcFam = Format(vgRs!Mto_CNU, "#,#0.000000")
        If Not IsNull(vgRs!prc_tasapergar) Then Lbl_TasaPG = Format(vgRs!prc_tasapergar, "#,#0.00")
        If Not IsNull(vgRs!Mto_BonoFon) Then Lbl_BonoAct = Format(vgRs!Mto_BonoFon, "#,#0.00")
        If Not IsNull(vgRs!Mto_ApoAdiFon) Then Lbl_ApoAdi = Format(vgRs!Mto_ApoAdiFon, "#,#0.00")
        If Not IsNull(vgRs!mto_priunifon) Then Lbl_PriUnica = Format(vgRs!mto_priunifon, "#,#0.00")
        
        Call flBuscaCodGlosa(vgCodTabla_TipRen, (vgRs!Cod_TipRen))
        Lbl_TipoRenta = (vgRs!Cod_TipRen) + " - " + vlElemento
        
        Call flBuscaCodGlosa(vgCodTabla_AltPen, (vgRs!Cod_Modalidad))
        Lbl_Alter = (vgRs!Cod_Modalidad) + " - " + vlElemento
        Call flBuscaCodGlosa(vgCodTabla_TipMon, (vgRs!Cod_Moneda))
        Lbl_Moneda(0) = (vgRs!Cod_Moneda) + " - " + vlElemento
    End If
    vgRs.Close
    
    vlSql = "SELECT fec_traspaso "
    vlSql = vlSql & "FROM pd_tmae_polprirec "
    vlSql = vlSql & "WHERE num_poliza = '" & iNumPol & "' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not (vgRs.EOF) Then
        If Not IsNull(vgRs!fec_traspaso) Then Lbl_FecTraPrima = DateSerial(Mid(vgRs!fec_traspaso, 1, 4), Mid(vgRs!fec_traspaso, 5, 2), Mid(vgRs!fec_traspaso, 7, 2))
    End If
    vgRs.Close
    
Exit Function
Err_CarCalculo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'------------------------------------
'CARGA INFORMACION EN LA CARPETA BONO
'------------------------------------
'Function flCargaCarpBono(inumpol As String, inumend As Integer)
'On Error GoTo Err_CargaBono
'
'    Call flLimpiarDatosBono
'    Call flInicializaGrillaBono
'    Call flCargaGrillaBono(inumpol, inumend)
'
'Exit Function
'Err_CargaBono:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Function

'-------------------------------------------
'CARGA INFORMACION EN LA GRILLA BENEFICIARIO
'-------------------------------------------
Function flCargaCarpBenef(iNumPol As String, inumend As Integer)
Dim vlFechaNac As String
Dim vlFechaFall As String
Dim vlFechaInv As String

On Error GoTo Err_CargaBen
    Dim vlRut As String
    vlSql = ""
    vlSql = "SELECT * FROM pd_tmae_polben,ma_tpar_tipoiden WHERE "
    vlSql = vlSql & "num_poliza = '" & iNumPol & "' "
    vlSql = vlSql & "and num_endoso= " & inumend & " "
    vlSql = vlSql & "and cod_tipoidenben= cod_tipoiden "
    Set vgRs = vgConexionBD.Execute(vlSql)
    Msf_GriAseg.rows = 1
    While Not vgRs.EOF
        
        vlFechaNac = DateSerial(Mid(vgRs!Fec_NacBen, 1, 4), Mid(vgRs!Fec_NacBen, 5, 2), Mid(vgRs!Fec_NacBen, 7, 2))
        
        If Not IsNull(vgRs!Fec_InvBen) Then
            vlFechaInv = DateSerial(Mid(vgRs!Fec_InvBen, 1, 4), Mid(vgRs!Fec_InvBen, 5, 2), Mid(vgRs!Fec_InvBen, 7, 2))
        Else
            vlFechaInv = ""
        End If
        
        If Not IsNull(vgRs!Fec_FallBen) Then
            vlFechaFall = DateSerial(Mid(vgRs!Fec_FallBen, 1, 4), Mid(vgRs!Fec_FallBen, 5, 2), Mid(vgRs!Fec_FallBen, 7, 2))
        Else
            vlFechaFall = ""
        End If
                
        Msf_GriAseg.AddItem Trim(vgRs!Num_Orden) & vbTab _
                            & (Trim(vgRs!Cod_Par)) & vbTab _
                            & (Trim(vgRs!Cod_GruFam)) & vbTab _
                            & (Trim(vgRs!Cod_Sexo)) & vbTab _
                            & (Trim(vgRs!Cod_SitInv)) & vbTab _
                            & vlFechaInv & vbTab _
                            & (Trim(vgRs!Cod_CauInv)) & vbTab _
                            & (Trim(vgRs!Cod_DerPen)) & vbTab _
                            & vlFechaNac & vbTab _
                            & (Trim(vgRs!GLS_TIPOIDENCOR)) & vbTab _
                            & (Trim(vgRs!Num_IdenBen)) & vbTab _
                            & (Trim(vgRs!Gls_NomBen)) & vbTab _
                            & (Trim(vgRs!Gls_NomSegBen)) & vbTab _
                            & (Trim(vgRs!Gls_PatBen)) & vbTab _
                            & (Trim(vgRs!Gls_MatBen)) & vbTab _
                            & (Trim(vgRs!Prc_Pension)) & vbTab _
                            & (Trim(vgRs!Mto_Pension)) & vbTab _
                            & (Trim(vgRs!Mto_PensionGar)) & vbTab _
                            & (Trim(vgRs!Fec_FallBen)) & vbTab
                           
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

'Carga el representante de los beneficiarios, si lo hubiese
Function flCargaCarpAfilPolRep(iNumPol As String, inumend As Integer)
Dim vlNombresRep As String
Dim vlApepatRep As String
Dim vlApematRep As String

On Error GoTo Err_Cargarep
    'Dim vlRut As String
    
    
TP = Left(Trim(Lbl_TipPen.Caption), 2)

If TP = "08" Then
    vlSql = ""
    vlSql = "SELECT * FROM pd_tmae_polrep WHERE "
    vlSql = vlSql & "num_poliza = '" & iNumPol & "' "
    vlSql = vlSql & "and num_endoso= " & inumend & " "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!Gls_NombresRep) Then
            vlNombresRep = vgRs!Gls_NombresRep
        Else
            vlNombresRep = ""
        End If
        If Not IsNull(vgRs!Gls_ApepatRep) Then
            vlApepatRep = vgRs!Gls_ApepatRep
        Else
            vlApepatRep = ""
        End If
        If Not IsNull(vgRs!Gls_ApematRep) Then
            vlApematRep = vgRs!Gls_ApematRep
        Else
            vlApematRep = ""
        End If
        
    End If
Else
    vlSql = ""
    vlSql = "SELECT * FROM pd_tmae_polben WHERE "
    vlSql = vlSql & "num_poliza = '" & iNumPol & "' and cod_par='99'"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
          If Not IsNull(vgRs!Gls_NomBen) Then
            vlNombresRep = vgRs!Gls_NomBen
        Else
            vlNombresRep = ""
        End If
        If Not IsNull(vgRs!Gls_PatBen) Then
            vlApepatRep = vgRs!Gls_PatBen
        Else
            vlApepatRep = ""
        End If
        If Not IsNull(vgRs!Gls_MatBen) Then
            vlApematRep = vgRs!Gls_MatBen
        Else
            vlApematRep = ""
        End If
 
    End If
End If
  
  Lbl_Representante.Caption = vlNombresRep & " " & vlApepatRep & " " & vlApematRep


    vgRs.Close
 
Exit Function
Err_Cargarep:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'-----------------------------
'INICIA LA GRILLA BENEFICIARIO
'-----------------------------
Function flIniGrillaBen()
On Error GoTo Err_IniGri
    Msf_GriAseg.Clear
    Msf_GriAseg.Cols = 19
    Msf_GriAseg.rows = 1
    Msf_GriAseg.FixedCols = 0
    Msf_GriAseg.Row = 0
    
    Msf_GriAseg.Col = 0
    Msf_GriAseg.ColWidth(0) = 800
    Msf_GriAseg.Text = "Nº Orden"
    
    Msf_GriAseg.Col = 1
    Msf_GriAseg.ColWidth(1) = 900
    Msf_GriAseg.Text = "Parentesco"
    
    Msf_GriAseg.Col = 2
    Msf_GriAseg.ColWidth(2) = 800
    Msf_GriAseg.Text = "Gru. Fam."
    
    Msf_GriAseg.Col = 3
    Msf_GriAseg.ColWidth(3) = 1100
    Msf_GriAseg.Text = "Sexo"
    
    Msf_GriAseg.Col = 4
    Msf_GriAseg.ColWidth(4) = 1100
    Msf_GriAseg.Text = "Sit.Inv."
    
    Msf_GriAseg.Col = 5
    Msf_GriAseg.ColWidth(5) = 1100
    Msf_GriAseg.Text = "Fec.Inv."
        
    Msf_GriAseg.Col = 6
    Msf_GriAseg.ColWidth(6) = 1100
    Msf_GriAseg.Text = "Cau.Inv."
    
    Msf_GriAseg.Col = 7
    Msf_GriAseg.ColWidth(7) = 1100
    Msf_GriAseg.Text = "Der.Pension"

    Msf_GriAseg.Col = 8
    Msf_GriAseg.ColWidth(8) = 1100
    Msf_GriAseg.Text = "Fec.Nac."
    
    Msf_GriAseg.Col = 9
    Msf_GriAseg.ColWidth(9) = 1100
    Msf_GriAseg.Text = "Tipo Ident."
    
    Msf_GriAseg.Col = 10
    Msf_GriAseg.ColWidth(10) = 1100
    Msf_GriAseg.Text = "Nº Ident."
    
    Msf_GriAseg.Col = 11
    Msf_GriAseg.ColWidth(11) = 1500
    Msf_GriAseg.Text = " 1er. Nombre"
    
    Msf_GriAseg.Col = 12
    Msf_GriAseg.ColWidth(12) = 1500
    Msf_GriAseg.Text = " 2do. Nombre"
    
    Msf_GriAseg.Col = 13
    Msf_GriAseg.ColWidth(13) = 2000
    Msf_GriAseg.Text = "Ap. Paterno"
    
    Msf_GriAseg.Col = 14
    Msf_GriAseg.ColWidth(14) = 2000
    Msf_GriAseg.Text = "Ap. Materno"
    
    Msf_GriAseg.Col = 15
    Msf_GriAseg.ColWidth(15) = 1100
    Msf_GriAseg.Text = "Porcentaje"
    
    Msf_GriAseg.Col = 16
    Msf_GriAseg.ColWidth(16) = 1100
    Msf_GriAseg.Text = "Mto.Pension"
    
    Msf_GriAseg.Col = 17
    Msf_GriAseg.ColWidth(17) = 1200
    Msf_GriAseg.Text = "Mto. PensionGar"
    
    Msf_GriAseg.Col = 18
    Msf_GriAseg.ColWidth(18) = 1300
    Msf_GriAseg.Text = "Fec.Fallecimiento"
    
Exit Function
Err_IniGri:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'-------------------------------------
'LIMPIA LABEL DE LA CARPETA ASEGURADOS
'-------------------------------------
Function flLimpiarDatosAseg()
On Error GoTo Err_LimpiaAseg

    Lbl_RutBen = ""
    Lbl_DgvBen = ""
    Lbl_NomBen = ""
    Lbl_NomBenSeg = ""
    Lbl_ApPatBen = ""
    Lbl_ApMatBen = ""
    Lbl_FecNacBen = ""
    Lbl_FecFallBen = ""
    Lbl_PensionBen = ""
    Lbl_PenGarBen = ""
    Lbl_Par = ""
    Lbl_GruFam = ""
    Lbl_SexoBen = ""
    Lbl_SitInvBen = ""
    Lbl_FecInvBen = ""
    Lbl_CauInvBen = ""
    Lbl_DerPen = ""
    Lbl_Porcentaje = ""
        
    Exit Function
Err_LimpiaAseg:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'-----------------------------------
'LIMPIA LABEL DE LA CARPETA AFILIADO
'-----------------------------------
Function flLimpiarDatosAfi()
On Error GoTo Err_Limpiar

    Lbl_RutAfi = ""
    Lbl_Asegurados = ""
    Lbl_DgvAfi = ""
    Lbl_NomAfi = ""
    Lbl_NomAfiSeg = ""
    Lbl_ApPatAfi = ""
    Lbl_ApMatAfi = ""
    Lbl_SexoAfi = ""
    Lbl_FecNac = ""
    Lbl_FecFall = ""
    Lbl_TipPen = ""
    Lbl_FecInv = ""
    Lbl_CauInv = ""
    Lbl_EstCivil = ""
    Lbl_Afp = ""
    Lbl_Salud = ""
    Lbl_Dir = ""
    Lbl_Comuna = ""
    Lbl_Provincia = ""
    Lbl_Region = ""
    Lbl_Fono = ""
    Lbl_Correo = ""
    Lbl_ViaPago = ""
    Lbl_Suc = ""
    Lbl_TipCta = ""
    Lbl_Bco = ""
    Lbl_NumCta = ""
    Lbl_Nacionalidad = ""
    Lbl_NumLiq = ""
       
Exit Function
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'-------------------------------
'LIMPIA LABEL DE LA CARPETA BONO
'-------------------------------
'Function flLimpiarDatosBono()
'On Error GoTo Err_LimpiaBono
'
'    Lbl_TipoBono = ""
'    Lbl_ValorNomBono = ""
'    Lbl_FecEmiBono = ""
'    Lbl_FecVenBono = ""
'    Lbl_PrcIntBono = ""
'    Lbl_ValUFBono = ""
'    Lbl_ValPesosBono = ""
'    Lbl_EdadCobroBono = ""
'
'Exit Function
'Err_LimpiaBono:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Function
'-----------------------------------------
'LIMPIA LOS LABEL DE LA CARPETA DE CALCULO
'-----------------------------------------
Function flLimpiarDatosCal()
On Error GoTo Err_LimpiaCal
    
    Lbl_FecDev = ""
    Lbl_FecIncorpora = ""
    Lbl_FecIniPago = ""
    Lbl_CUSPP = ""
    Lbl_TipoRenta = ""
    Lbl_AnnosDif = ""
    Lbl_Alter = ""
    Lbl_MesesGar = ""
    Lbl_RentaAFP = ""
    Lbl_PrcRtaTmp = ""
    Lbl_IndCob = ""
    Lbl_CtaInd = ""
    Lbl_FacPenElla = ""
    Lbl_DerCre = ""
    Lbl_DerGra = ""
    Lbl_ComInt = ""
    Lbl_RutCorr = ""
    Lbl_DgvCorr = ""
    Lbl_MtoPension = ""
    Lbl_MtoPenGar = ""
    Lbl_TasaCE = ""
    Lbl_TasaVta = ""
    Lbl_TasaTIR = ""
    Lbl_TasaPG = ""
    Lbl_BonoAct = ""
    Lbl_ApoAdi = ""
    Lbl_BenSocial = ""
    Lbl_MtoPrimaUniSim = ""
    Lbl_MtoPrimaUniDif = ""
    Lbl_PrcFam = ""
    Lbl_PriUnica = ""
    Lbl_FecTraPrima = ""

'I--- ABV 05/02/2011 ---
    Lbl_ReajusteTipo.Caption = ""
    Lbl_ReajusteValor.Caption = ""
    Lbl_ReajusteMoneda.Caption = ""
'F--- ABV 05/02/2011 ---

Exit Function
Err_LimpiaCal:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'----------------------
'FUNCION IMPRIME POLIZA
'----------------------
Function flImprimirPoliza(iNumPol As String, inumend As Integer)
On Error GoTo Err_Imprimir

   'vlArchivo = strRpt & "PD_Rpt_CalConsulta.rpt"   '\Reportes
   vlArchivo = strRpt & "PD_Rpt_PolizaDef.rpt"  'Cambio para reimpresion del 20/04/2009 poliza 39
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Póliza no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Function
   End If
  
   'numero de endoso
   vlNumEnd = Trim(Txt_Endoso)

'  numero de póliza
   vlNumPol = Trim(iNumPol)
      
    'busca el nombre del Afiliado
    vlSql = ""
    vlSql = "SELECT gls_nomben,gls_patben,gls_matben,cod_tipoidenben "
    vlSql = vlSql & "FROM pd_tmae_polben WHERE "
    vlSql = vlSql & "num_poliza = '" & iNumPol & "' AND cod_par = '99' "
    vlSql = vlSql & "AND num_endoso= " & inumend & ""
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlNomAfi = Trim(vgRs!Gls_NomBen) & " " & Trim(vgRs!Gls_PatBen) & " " & Trim(vgRs!Gls_MatBen)
        vlTipoIden = Trim(vgRs!Cod_TipoIdenBen)
        vlTipoIden = fgBuscarNombreTipoIden(vlTipoIden, True)
    Else
        MsgBox "Nombre del Afiliado No encontrado o no exsiste", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        Exit Function
    End If
   
'   'busca el tipo de identificación
'   vlSql = ""
'   vlSql = "SELECT gls_tipoiden "
'   vlSql = vlSql & "FROM pd_tmae_poliza p,pd_tmae_polben b,ma_tpar_tipoiden i "
'   vlSql = vlSql & "WHERE p.num_poliza = '" & inumpol & "' "
'   vlSql = vlSql & "AND p.num_endoso = " & inumend & " "
'   vlSql = vlSql & "AND p.num_poliza = b.num_poliza "
'   vlSql = vlSql & "AND p.num_endoso = b.num_endoso "
'   vlSql = vlSql & "AND b.cod_tipoidenben = i.cod_tipoiden "
'   vlSql = vlSql & "AND b.cod_par = '99' "
'
'   Set vgRs = vgConexionBD.Execute(vlSql)
'   If Not vgRs.EOF Then
'        vlTipoIden = Trim(vgRs!gls_tipoiden)
'   Else
'       MsgBox "Nombre del Afiliado No encontrado o no exsiste", vbCritical, "Datos Incompletos"
'       Screen.MousePointer = 0
'       Exit Function
'   End If
'   vgRs.Close
'
'    'busca la identificación del beneficiario
'   vlSql = ""
'   vlSql = "SELECT gls_tipoiden "
'   vlSql = vlSql & "FROM pd_tmae_poliza p,pd_tmae_polben b,ma_tpar_tipoiden i "
'   vlSql = vlSql & "WHERE p.num_poliza = '" & inumpol & "' "
'   vlSql = vlSql & "AND p.num_endoso = " & inumend & " "
'   vlSql = vlSql & "AND p.num_poliza = b.num_poliza "
'   vlSql = vlSql & "AND p.num_endoso = b.num_endoso "
'   vlSql = vlSql & "AND b.cod_tipoidenben = i.cod_tipoiden "
'
'   Set vgRs = vgConexionBD.Execute(vlSql)
'   If Not vgRs.EOF Then
'        vlTipoIdenBen = Trim(vgRs!gls_tipoiden)
'   Else
'       MsgBox "Nombre del Afiliado No encontrado o no exsiste", vbCritical, "Datos Incompletos"
'       Screen.MousePointer = 0
'       Exit Function
'   End If
'   vgRs.Close
'
'   vlSql = ""
'   vlSql = "SELECT num_poliza FROM pd_tmae_polben WHERE "
'   vlSql = vlSql & "num_poliza= '" & inumpol & "' AND "
'   vlSql = vlSql & "num_endoso= " & inumend & ""
'   Set vgRs = vgConexionBD.Execute(vlSql)
'   Dim vlBono As Integer
'   If Not vgRs.EOF Then
'        vlBono = 1
'   Else
'        vlBono = 0
'   End If
   
   'Nombre Identificación'
   vlRut = Trim(Lbl_RutAfi)
   
   'Años Dif.'
   vlAnnosDif = Trim(Lbl_AnnosDif)
   
   'Descripción de la isapre
   vlCodIsapre = Trim(Mid(Lbl_Salud, (InStr(1, Lbl_Salud, "-") + 1), Len(Lbl_Salud)))
   
   'codigo afp
   vlCodAFP = Trim(Mid(Lbl_Afp, (InStr(1, Lbl_Afp, "-") + 1), Len(Lbl_Afp)))
      
   'Sucursal
   vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
   
   'Parentesco
   vlParentesco = Trim(Lbl_Par)
   
   'Fecha de Nacimiento
   vlFecNac = Trim(Lbl_FecNac)
   
   'Fecha de Vigencia
   vlFechaVig = Trim(Lbl_FecVig)
   
   'Nacionalidad
    vlNacionalidad = UCase(Trim(Lbl_Nacionalidad))
      
    'Número de Liquidación
    vlNumLiquidacion = Trim(Lbl_NumLiq)
    
   'Buscar la descripcion Tipo de Moneda
 
   vlSql = "SELECT cod_scomp as gls_elemento FROM ma_tpar_tabcod "
   vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_TipMon & "' "
   vlSql = vlSql & "AND cod_elemento = (SELECT cod_moneda FROM "
   vlSql = vlSql & "pd_tmae_poliza WHERE num_poliza = '" & iNumPol & "' AND "
   vlSql = vlSql & "num_endoso = '" & inumend & "')"
   Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) And Not IsNull(vgRs!gls_elemento) Then
        vlCodMoneda = vgRs!gls_elemento
   Else
        MsgBox "Tipo de Moneda no Encontrada", vbCritical, "Datos Incompleos"
        Screen.MousePointer = 0
        Exit Function
   End If
   vgRs.Close
   
    'busca tipo de renta inmediata o diferida
    vlSql = ""
    vlSql = "SELECT gls_elemento FROM ma_tpar_tabcod "
    vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_TipRen & "' "
    vlSql = vlSql & "AND cod_elemento = (SELECT cod_tipren FROM "
    vlSql = vlSql & "pd_tmae_poliza WHERE num_poliza = '" & iNumPol & "' AND "
    vlSql = vlSql & "num_endoso = '" & inumend & "')"
    Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) And Not IsNull(vgRs!gls_elemento) Then
        vlCodTipRen = vgRs!gls_elemento
   Else
        MsgBox "Tipo de Cobertura no encontrado", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        Exit Function
   End If
   vgRs.Close
   
   'busca tipo pension
    vlSql = ""
    vlSql = "SELECT gls_elemento FROM ma_tpar_tabcod "
    vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_TipPen & "' "
    vlSql = vlSql & "AND cod_elemento = (SELECT cod_tippension FROM "
    vlSql = vlSql & "pd_tmae_poliza WHERE num_poliza = '" & iNumPol & "' AND "
    vlSql = vlSql & "num_endoso = '" & inumend & "')"
    Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) And Not IsNull(vgRs!gls_elemento) Then
        vlCodTipPen = vgRs!gls_elemento
   Else
        MsgBox "Tipo de Pensión no encontrado", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        Exit Function
   End If
   vgRs.Close
 
    'busca modalidad
    vlSql = ""
    vlSql = "SELECT gls_elemento FROM ma_tpar_tabcod "
    vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_AltPen & "' "
    vlSql = vlSql & "AND cod_elemento = (SELECT cod_modalidad FROM "
    vlSql = vlSql & "pd_tmae_poliza WHERE num_poliza = '" & iNumPol & "' AND "
    vlSql = vlSql & "num_endoso = '" & inumend & "')"
    Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) And Not IsNull(vgRs!gls_elemento) Then
        vlCodAl = vgRs!gls_elemento
   Else
        MsgBox "Tipo de Modalidad no encontrado", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        Exit Function
   End If
          
  vgRs.Close
  
  'busca tipo de via de pago
    vlSql = ""
    vlSql = "SELECT gls_elemento FROM ma_tpar_tabcod "
    vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_ViaPago & "' "
    vlSql = vlSql & "AND cod_elemento = (SELECT cod_viapago FROM "
    vlSql = vlSql & "pd_tmae_poliza WHERE num_poliza = '" & iNumPol & "' AND "
    vlSql = vlSql & "num_endoso = '" & inumend & "')"
    Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) And Not IsNull(vgRs!gls_elemento) Then
        vlCodvpg = vgRs!gls_elemento
   Else
        vlCodvpg = ""
   End If
          
  vgRs.Close
  
  'busca banco
    vlSql = ""
    vlSql = "SELECT gls_elemento FROM ma_tpar_tabcod "
    vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_Bco & "' "
    vlSql = vlSql & "AND cod_elemento = (SELECT cod_banco FROM "
    vlSql = vlSql & "pd_tmae_poliza WHERE num_poliza = '" & iNumPol & "' AND "
    vlSql = vlSql & "num_endoso = '" & inumend & "')"
    Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) And Not IsNull(vgRs!gls_elemento) Then
        vlCodBco = vgRs!gls_elemento
   Else
        vlCodBco = ""
   End If
          
  vgRs.Close
  
  'tipo de cuenta
    vlSql = ""
    vlSql = "SELECT gls_elemento FROM ma_tpar_tabcod "
    vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_TipCta & "' "
    vlSql = vlSql & "AND cod_elemento = (SELECT cod_tipcuenta FROM "
    vlSql = vlSql & "pd_tmae_poliza WHERE num_poliza = '" & iNumPol & "' AND "
    vlSql = vlSql & "num_endoso = '" & inumend & "')"
    Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) And Not IsNull(vgRs!gls_elemento) Then
        vlCodTC = vgRs!gls_elemento
   Else
        vlCodTC = ""
   End If
          
   vgRs.Close
          
   Call pBuscaAntecedentes
   Call pBuscaRepresentante
   
   vgQuery = "{pd_TMAE_POLIZA.NUM_POLIZA} = '" & vlNumPol & "'"
   'vgQuery = "{pd_TMAE_POLIZA.NUM_POLIZA} = '" & vlNumPol & "' AND "
   'vgQuery = vgQuery & "{pd_TMAE_POLIZA.NUM_ENDOSO} = " & vlNumEnd & ""

   Rpt_Poliza.Reset
   Rpt_Poliza.ReportFileName = vlArchivo
   Rpt_Poliza.Connect = vgRutaDataBase
   
   'Rpt_Poliza.Formulas(0) = "Num_Pol= '" & vlNumPol & "'"
   'Rpt_Poliza.Formulas(1) = "NombreAfi = '" & vlNomAfi & "'"
   'Rpt_Poliza.Formulas(2) = "InsSalud = '" & vlCodIsapre & "'"
   'Rpt_Poliza.Formulas(3) = "Afp ='" & vlCodAFP & "'"
   'Rpt_Poliza.Formulas(4) = "TipoPension = '" & vlCodTipPen & "'"
   'Rpt_Poliza.Formulas(5) = "TipoPlan = '" & vlCodTipRen & "'"
   'Rpt_Poliza.Formulas(6) = "AnnosDif = '" & vlAnnosDif & "'"
   'Rpt_Poliza.Formulas(7) = "TipMod = '" & vlCodAl & "'"
   'Rpt_Poliza.Formulas(8) = "FechaVigencia = '" & vlFechaVig & "'"
   'Rpt_Poliza.Formulas(9) = "ViaPago = '" & vlCodvpg & "'"
   'Rpt_Poliza.Formulas(10) = "Sucursal = '" & vlNombreSucursal & "'"
   'Rpt_Poliza.Formulas(11) = "TipoCuenta = '" & vlCodTC & "'"
   'Rpt_Poliza.Formulas(12) = "Banco = '" & vlCodBco & "'"
   'Rpt_Poliza.Formulas(13) = "FecNac = '" & vlFecNac & "'"
   'Rpt_Poliza.Formulas(14) = "Moneda = '" & vlCodMoneda & "'"
   'Rpt_Poliza.Formulas(15) = "TipoIden = '" & vlTipoIden & "'"
   'Rpt_Poliza.Formulas(16) = "NivelVerTasa = '" & vgNivelIndicadorVer & "'"
   'Rpt_Poliza.Formulas(17) = "Nacionalidad = '" & vlNacionalidad & "'"
   'Rpt_Poliza.Formulas(18) = "NumLiquidacion = '" & vlNumLiquidacion & "'"
   'Rpt_Poliza.Formulas(19) = "MesGar = '" & Lbl_MesesGar & "'"
   'Rpt_Poliza.Formulas(20) = "NombreCompania = '" & UCase(vgNombreCompania) & "'"
   'Rpt_Poliza.Formulas(21) = "Concatenar = '" & vlCobertura & "'"
   
   Rpt_Poliza.Formulas(0) = "NombreAfi = '" & vlNomAfi & "'"
   Rpt_Poliza.Formulas(1) = "TipoPension = '" & vlCodTipPen & "'"
   Rpt_Poliza.Formulas(2) = "MesGar = '" & Lbl_MesesGar & "'"
   Rpt_Poliza.Formulas(3) = "NombreCompania = '" & UCase(vgNombreCompania) & "'"
   Rpt_Poliza.Formulas(4) = "Concatenar = '" & vlCobertura & "'"
   Rpt_Poliza.Formulas(5) = "Sucursal = '" & vlNombreSucursal & "'"
   'RVF 20090914
   Rpt_Poliza.Formulas(6) = "RepresentanteNom = '" & vlRepresentante & "'"
   Rpt_Poliza.Formulas(7) = "RepresentanteDoc = '" & vlDocum & "'"
   Rpt_Poliza.Formulas(8) = "CodTipPen = '" & Left(Trim(Lbl_TipPen.Caption), 2) & "'"
   '*****

   Rpt_Poliza.SelectionFormula = vgQuery
   Rpt_Poliza.Destination = crptToWindow
   Rpt_Poliza.WindowState = crptMaximized
   'Rpt_Poliza.WindowTitle = "Consulta Póliza"
   Rpt_Poliza.Action = 1
   Screen.MousePointer = 0

Exit Function
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'Function flInicializaGrillaBono()
'
'    Msf_GrillaBono.Clear
'    Msf_GrillaBono.Cols = 8
'    Msf_GrillaBono.Rows = 1
'    Msf_GrillaBono.RowHeight(0) = 250
'    Msf_GrillaBono.Row = 0
'
'    Msf_GrillaBono.Col = 0
'    Msf_GrillaBono.Text = "Tipo Bono"
'    Msf_GrillaBono.ColWidth(0) = 1000
'    Msf_Grilla.ColAlignment(0) = 1  'centrado
'
'    Msf_GrillaBono.Col = 1
'    Msf_GrillaBono.Text = "Valor Nominal"
'    Msf_GrillaBono.ColWidth(1) = 1200
'
'    Msf_GrillaBono.Col = 2
'    Msf_GrillaBono.Text = "Fec. Emisión"
'    Msf_GrillaBono.ColWidth(2) = 1200
'
'    Msf_GrillaBono.Col = 3
'    Msf_GrillaBono.Text = "Fec. Venc."
'    Msf_GrillaBono.ColWidth(3) = 1200
'
'    Msf_GrillaBono.Col = 4
'    Msf_GrillaBono.Text = "Tasa Int."
'    Msf_GrillaBono.ColWidth(4) = 1200
'
'    Msf_GrillaBono.Col = 5
'    Msf_GrillaBono.Text = "Mto. Bono UF"
'    Msf_GrillaBono.ColWidth(5) = 1200
'
'    Msf_GrillaBono.Col = 6
'    Msf_GrillaBono.Text = "Mto. Bono Pesos"
'    Msf_GrillaBono.ColWidth(6) = 1500
'
'    Msf_GrillaBono.Col = 7
'    Msf_GrillaBono.Text = "Edad Cobro"
'    Msf_GrillaBono.ColWidth(7) = 1200
'
'End Function

'Function flCargaGrillaBono(iNumPoliza As String, iNumEndoso As Integer)
'
'On Error GoTo Err_flCargaGrillaBono
'
'    vgSql = ""
'    vgSql = "SELECT  b.cod_tipobono,b.mto_valnom,b.fec_emi,b.fec_ven, "
'    vgSql = vgSql & "b.prc_tasaint,b.mto_bonoactuf,b.mto_bonoact, "
'    vgSql = vgSql & "b.num_edadcob "
'    vgSql = vgSql & "FROM pd_tmae_polbon b "
'    vgSql = vgSql & "WHERE b.num_poliza = '" & Trim(iNumPoliza) & "' AND "
'    vgSql = vgSql & "b.num_endoso = " & (iNumEndoso) & " "
'    vgSql = vgSql & "ORDER BY cod_tipobono "
'    Set vgRs = vgConexionBD.Execute(vgSql)
'    If Not vgRs.EOF Then
'       Call flInicializaGrillaBono
'
'       While Not vgRs.EOF
'
'          vltipobono = Trim(vgRs!cod_tipobono)  '& " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipBono, Trim(vgRs!cod_tipobono)))
'
'          Msf_GrillaBono.AddItem vltipobono & vbTab _
'          & (Format((vgRs!mto_valnom), "###,###,##0.00")) & vbTab _
'          & (DateSerial(Mid((vgRs!fec_emi), 1, 4), Mid((vgRs!fec_emi), 5, 2), Mid((vgRs!fec_emi), 7, 2))) & vbTab _
'          & (DateSerial(Mid((vgRs!fec_ven), 1, 4), Mid((vgRs!fec_ven), 5, 2), Mid((vgRs!fec_ven), 7, 2))) & vbTab _
'          & (Format((vgRs!prc_tasaint), "###,###,##0.00")) & vbTab _
'          & (Format((vgRs!mto_bonoactuf), "###,###,##0.00")) & vbTab _
'          & (Format((vgRs!mto_bonoact), "###,###,##0.00")) & vbTab _
'          & (Trim(vgRs!num_edadcob))
'
'          vgRs.MoveNext
'       Wend
'    End If
'    vgRs.Close
'
'Exit Function
'Err_flCargaGrillaBono:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function

Private Sub Cmd_BuscarPol_Click()
On Error GoTo Err_Buscar
    
    
'I--- ABV 04/11/2007 ---
    'Txt_Endoso = "1"
'F--- ABV 04/11/2007 ---
    
    If Trim(Txt_NumPol) <> "" And Trim(Txt_Endoso) <> "" Then
        Call flBuscaPoliza(Trim(Txt_NumPol), Trim(Txt_Endoso))
    Else
        MsgBox "Debe Ingresar Número de Póliza y Endoso", vbExclamation, "Falta Información"
        Txt_NumPol.SetFocus
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

Dim vlCodPar As String
Dim vlCodDerPen As String
Dim vlNombreSucursal, vlNombreTipoPension As String
'Dim RS As ADODB.Recordset
Dim vlFecTras As String

    vlFecTras = Lbl_FecVig.Caption
    vlFecTras = Mid(vlFecTras, 7, 4) & Mid(vlFecTras, 4, 2) & Mid(vlFecTras, 1, 2)

 'Validar el Ingreso de la Póliza
    If Txt_NumPol = "" Then
        MsgBox "Debe ingresar Póliza a Consultar.", vbCritical, "Error de Datos"
        Txt_Poliza.SetFocus
        Exit Sub
    End If
    
    
    'Valida que exista la Póliza
    If Trim(Lbl_RutAfi) = "" Then
        MsgBox "Debe Buscar Datos de la Póliza", vbCritical, "Error de Datos"
        Cmd_Buscar.SetFocus
        Exit Sub
    End If
    
    vlCodPar = "99" 'Causante
    vlCodDerPen = "10" 'Sin Derecho a Pension
    vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
    vlNombreTipoPension = fgObtenerNombre_TextoCompuesto(Lbl_TipPen)
    
On Error GoTo mierror

'    Set RS = New ADODB.Recordset
'    RS.CursorLocation = adUseClient
'   RS.Open "PD_LISTA_POLIZA.LISTAR('" & Txt_NumPol.Text & "', '" & Txt_Endoso.Text & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Call pBuscaRepresentante
    
    Dim rs      As ADODB.Recordset, cmd As ADODB.Command
    Dim conn    As ADODB.Connection
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
    Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")
    
    conn.Provider = "OraOLEDB.Oracle"
    conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
    conn.CursorLocation = adUseClient
    conn.Open
    
    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = conn
    
    objCmd.CommandText = "PD_LISTA_POLIZA.LISTAR"
    objCmd.CommandType = adCmdStoredProc
    
    Set param1 = objCmd.CreateParameter("POLIZA", adVarChar, adParamInput, 10, Txt_NumPol.Text)
    objCmd.Parameters.Append param1
    
    Set param2 = objCmd.CreateParameter("NUM_END", adInteger, adParamInput)
    param2.Value = Txt_Endoso.Text
    objCmd.Parameters.Append param2
    
    Set rs = objCmd.Execute
  
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaDef.rpt"), ".RPT", ".TTX"), 1)
    
    If objRep.CargaReporte(strRpt & "", "PD_Rpt_PolizaDef.rpt", "Póliza", rs, True, _
                            ArrFormulas("NombreAfi", Lbl_NomAfi.Caption & " " & Lbl_NomAfiSeg.Caption & " " & Lbl_ApPatAfi.Caption & " " & Lbl_ApMatAfi.Caption), _
                            ArrFormulas("TipoPension", vlNombreTipoPension), _
                            ArrFormulas("MesGar", Lbl_MesesGar.Caption), _
                            ArrFormulas("NombreCompania", UCase(vgNombreCompania)), _
                            ArrFormulas("Concatenar", vlCobertura), _
                            ArrFormulas("Sucursal", vlNombreSucursal), _
                            ArrFormulas("RepresentanteNom", vlRepresentante), _
                            ArrFormulas("RepresentanteDoc", vlDocum), _
                            ArrFormulas("CodTipPen", Left(Trim(Lbl_TipPen.Caption), 2)), _
                            ArrFormulas("TipoDocTit", Trim(Lbl_RutAfi.Caption)), _
                            ArrFormulas("NumDocTit", Trim(Lbl_DgvAfi.Caption)), _
                            ArrFormulas("fec_trasp", vlFecTras)) = False Then
            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
    
Exit Sub
mierror:
    MsgBox "No pudo cargar el reporte " & Err.Description, vbInformation
    
    
'On Error GoTo Err_Imprimir
'
'    If Trim(Txt_NumPol = "") Or Txt_Endoso = "" Then
'        MsgBox "Debe escoger una Póliza para Imprimir", vbCritical, "Datos Incompletos"
'        Exit Sub
'    End If
'
'    Txt_NumPol = Trim(Txt_NumPol)
'    Txt_Endoso = Trim(Txt_Endoso)
'
'    vlSql = ""
'    vlSql = "SELECT num_poliza FROM pd_tmae_poliza WHERE "
'    vlSql = vlSql & "num_poliza= '" & (Txt_NumPol) & "' AND "
'    vlSql = vlSql & "num_endoso= " & (Txt_Endoso) & ""
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If (vgRs.EOF) Then
'        MsgBox "La Póliza que intenta Imprimir No existe en la BD", vbCritical, "Error de Impresion"
'        Exit Sub
'    End If
'
'    Screen.MousePointer = 11
'    Call flImprimirPoliza(Txt_NumPol, Txt_Endoso)
'
'Exit Sub
'Err_Imprimir:
'Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Sub
'
'Private Sub Cmd_Limpiar_Click()
'On Error GoTo Err_Limpiar
'
'    flLimpiarDatosAfi
'    flLimpiarDatosCal
'    'flLimpiarDatosBono
'    flLimpiarDatosAseg
'    Lbl_FecVig = ""
'    Txt_NumPol = ""
'    Lbl_NumCot = ""
'    Lbl_NomAfiSeg = ""
'    Txt_Endoso = ""
'    Lbl_SolOfe = ""
'    Lbl_IdeOfe = ""
'    Lbl_SecOfe = ""
'    SSTab_Poliza.Enabled = False
'    SSTab_Poliza.Tab = 0
'    Fra_Poliza.Enabled = True
'    Cmd_Poliza.Enabled = True
'    Call flIniGrillaBen
'    Cmd_Poliza.Enabled = True
'    Txt_NumPol.SetFocus
'
'Exit Sub
'Err_Limpiar:
'Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
End Sub

Private Sub Cmd_Poliza_Click()
On Error GoTo Err_BuscarPoliza

    Frm_BuscarPolEnd.flInicio ("Frm_CalConsulta")
    
Exit Sub
Err_BuscarPoliza:
Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Salir_Click()
On Error GoTo Err_Salir

    Unload Me
    
Exit Sub
Err_Salir:
Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub cmdPlanPol_Click()
    Frm_ConsultaPolizas.Show
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form
        
    Frm_CalConsulta.Top = 0
    Frm_CalConsulta.Left = 0
    
    Fra_Poliza.Enabled = True
    Cmd_Poliza.Enabled = True
    Txt_NumPol.Enabled = True
    SSTab_Poliza.Tab = 0
    flIniGrillaBen
    Msf_GriAseg.Enabled = False
    SSTab_Poliza.Enabled = False
    SSTab_Poliza.Tab = 0
     
    If vgNivelIndicadorVer = "N" Then
        Lbl_BenSocial.Visible = False 'ABV 20/08/2007
        Lbl_TasaCE.Visible = False
        Lbl_TasaVta.Visible = False
        Lbl_TasaTIR.Visible = False
        Lbl_TasaPG.Visible = False
    Else
        Lbl_BenSocial.Visible = True 'ABV 20/08/2007
        Lbl_TasaCE.Visible = True
        Lbl_TasaVta.Visible = True
        Lbl_TasaTIR.Visible = True
        'Lbl_TasaPG.Visible = True 'ABV 04/11/2007
    End If
       
    'Call fgComboGeneral(vgCodTabla_InsSal, Cmb_Salud)
       
    Call fgCargarTablaMoneda(vgCodTabla_TipMon, egTablaMoneda(), vgNumeroTotalTablasMoneda)
    
Exit Sub
Err_Form:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_GriAseg_DblClick()
On Error GoTo Err_Grilla

    Msf_GriAseg.Col = 0
    Msf_GriAseg.Row = Msf_GriAseg.RowSel
    If (Msf_GriAseg.Text = "") Or (Msf_GriAseg.Row = 0) Then
        Exit Sub
    End If
    
    Call flLimpiarDatosAseg
    
    Lbl_NumOrden = Msf_GriAseg.Text
    
    Msf_GriAseg.Col = 1
    Call flBuscaCodGlosa(vgCodTabla_Par, Trim(Msf_GriAseg.Text))
    Lbl_Par = Trim(Msf_GriAseg.Text) + " - " + vlElemento
    
    Msf_GriAseg.Col = 2
    Call flBuscaCodGlosa(vgCodTabla_GruFam, Trim(Msf_GriAseg.Text))
    Lbl_GruFam = Trim(Msf_GriAseg.Text) + " - " + vlElemento
    
    Msf_GriAseg.Col = 3
    Call flBuscaCodGlosa(vgCodTabla_Sexo, Trim(Msf_GriAseg.Text))
    Lbl_SexoBen = Trim(Msf_GriAseg.Text) + " - " + vlElemento
    
    Msf_GriAseg.Col = 4
    Call flBuscaCodGlosa(vgCodTabla_SitInv, Trim(Msf_GriAseg.Text))
    Lbl_SitInvBen = Trim(Msf_GriAseg.Text) + " - " + vlElemento
    
    Msf_GriAseg.Col = 5
    If Msf_GriAseg.Text <> "" Then
        Lbl_FecInvBen = Msf_GriAseg.Text
    End If
    
    Msf_GriAseg.Col = 6
    Call flBuscaGlosaCauInv(Msf_GriAseg.Text)
    Lbl_CauInvBen = Trim(Msf_GriAseg.Text) + " - " + vlElemento
    
    Msf_GriAseg.Col = 7
    Call flBuscaCodGlosa(vgCodTabla_DerPen, Trim(Msf_GriAseg.Text))
    Lbl_DerPen = Trim(Msf_GriAseg.Text) + " - " + vlElemento
    
    Msf_GriAseg.Col = 8
    If Msf_GriAseg.Text <> "" Then
        Lbl_FecNacBen = Msf_GriAseg.Text
    End If
    
    Msf_GriAseg.Col = 9
    If Msf_GriAseg.Text <> "" Then
        Lbl_RutBen = Msf_GriAseg.Text
    End If
    
    Msf_GriAseg.Col = 10
    If Msf_GriAseg.Text <> "" Then
        Lbl_DgvBen = Msf_GriAseg.Text
    End If
 
    Msf_GriAseg.Col = 11
    If Msf_GriAseg.Text <> "" Then
        Lbl_NomBen = Msf_GriAseg.Text
    End If
    
    Msf_GriAseg.Col = 12
    If Msf_GriAseg.Text <> "" Then
        Lbl_NomBenSeg = Msf_GriAseg.Text
    End If
    
    Msf_GriAseg.Col = 13
    If Msf_GriAseg.Text <> "" Then
        Lbl_ApPatBen = Msf_GriAseg.Text
    End If
    
    Msf_GriAseg.Col = 14
    If Msf_GriAseg.Text <> "" Then
        Lbl_ApMatBen = Format(Msf_GriAseg.Text, "#,#0.00")
    End If
    
    Msf_GriAseg.Col = 15
    If Msf_GriAseg.Text <> "" Then
        Lbl_Porcentaje = Format(Msf_GriAseg.Text, "#,#0.00")
    End If
        
    Msf_GriAseg.Col = 16
    If Msf_GriAseg.Text <> "" Then
        Lbl_PensionBen = Format(Msf_GriAseg.Text, "#,#0.00")
    End If
        
    Msf_GriAseg.Col = 17
    If Msf_GriAseg.Text <> "" Then
        Lbl_PenGarBen = Format(Msf_GriAseg.Text, "#,#0.00")
    End If
       
    Msf_GriAseg.Col = 18
    If Msf_GriAseg.Text <> "" Then
        Lbl_FecFallBen = DateSerial(Mid(Msf_GriAseg.Text, 1, 4), Mid(Msf_GriAseg.Text, 5, 2), Mid(Msf_GriAseg.Text, 7, 2))
    End If
    
Exit Sub
Err_Grilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_GrillaBono_Click()
On Error GoTo Err_Msf_GrillaBono_Click

    Msf_GrillaBono.Col = 0
    Msf_GrillaBono.Row = Msf_GrillaBono.RowSel
    If (Msf_GrillaBono.Text = "") Or (Msf_GrillaBono.Row = 0) Then
        Exit Sub
    End If
    
'    Call flLimpiarDatosBono
    
    Msf_GrillaBono.Col = 0
    vltipobono = Trim(Msf_GrillaBono.Text) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipBono, Trim(Msf_GrillaBono.Text)))
    Lbl_TipoBono = vltipobono
    
    Msf_GrillaBono.Col = 1
    Lbl_ValorNomBono = Msf_GrillaBono.Text
    
    Msf_GrillaBono.Col = 2
    Lbl_FecEmiBono = Msf_GrillaBono.Text
    
    Msf_GrillaBono.Col = 3
    Lbl_FecVenBono = Msf_GrillaBono.Text
    
    Msf_GrillaBono.Col = 4
    Lbl_PrcIntBono = Msf_GrillaBono.Text
    
    Msf_GrillaBono.Col = 5
    Lbl_ValUFBono = Msf_GrillaBono.Text
    
    Msf_GrillaBono.Col = 6
    Lbl_ValPesosBono = Msf_GrillaBono.Text
    
    Msf_GrillaBono.Col = 7
    Lbl_EdadCobroBono = Msf_GrillaBono.Text
    
Exit Sub
Err_Msf_GrillaBono_Click:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Endoso_Change()
    If Not IsNumeric(Txt_Endoso) Then
        Txt_Endoso = ""
    End If
End Sub

Private Sub Txt_Endoso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmd_BuscarPol.SetFocus
    End If
End Sub

Private Sub Txt_NumPol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'I--- ABV 04/11/2007 ---
        'Txt_Endoso.SetFocus
        'Cmd_BuscarPol.SetFocus
'F--- ABV 04/11/2007 ---

           Dim vlRegistro As New ADODB.Recordset
        
           vgSql = "SELECT MAX(NUM_ENDOSO) AS NUM_ENDOSO FROM PD_TMAE_POLIZA WHERE NUM_POLIZA=" & Txt_NumPol & ""
        
           Set vlRegistro = vgConexionBD.Execute(vgSql)
           If Not (vlRegistro.EOF) Then
               Txt_Endoso = IIf(IsNull(vlRegistro!Num_Endoso), "1", vlRegistro!Num_Endoso)
           End If
         

'RRR 07/05/2013
        Txt_Endoso.SetFocus
'RRR
    End If
End Sub

Private Sub Txt_NumPol_LostFocus()
    If Trim(Txt_NumPol) <> "" Then
        Txt_NumPol = Format(Txt_NumPol, "000000000#")
    End If
End Sub

Private Sub pBuscaRepresentante()

On Error GoTo Err_Cargarep
Dim vlSql As String

TP = Left(Trim(Lbl_TipPen.Caption), 2)

If TP = "08" Then

    vlSql = ""
    vlSql = "SELECT * FROM pd_tmae_polrep a, ma_tpar_tipoiden b WHERE "
    vlSql = vlSql & "num_poliza = '" & Txt_NumPol & "' and a.cod_tipoidenrep = b.cod_tipoiden"
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
    vlSql = vlSql & "num_poliza = '" & Txt_NumPol & "' and cod_par='99'"
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



Private Sub pBuscaAntecedentes()
Dim vlCodPa As String
Dim vlRegistro As ADODB.Recordset
Dim vlDif As Double
Dim vlNomSeg As String
On Error GoTo Err_buscaAnt
    
    vgSql = ""
    vlCodTp = "TP"
    vlCodTr = "TR"
    vlCodAl = "AL"
    vlCodPa = "99"
    
    vgSql = ""
    vgSql = "SELECT  p.cod_modalidad,r.gls_elemento as gls_renta,"
    vgSql = vgSql & "m.gls_elemento as gls_modalidad,"
    vgSql = vgSql & "p.cod_cobercon, b.gls_cobercon, p.cod_dercre, p.cod_dergra "
    vgSql = vgSql & "FROM pd_tmae_poliza p, ma_tpar_tabcod r, "
    vgSql = vgSql & "ma_tpar_tabcod m, ma_tpar_cobercon b "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "p.num_poliza = '" & Trim(Txt_NumPol.Text) & "' AND "
    vgSql = vgSql & "r.cod_tabla = '" & Trim(vlCodTr) & "' AND "
    vgSql = vgSql & "r.cod_elemento = p.cod_tipren AND "
    vgSql = vgSql & "m.cod_tabla = '" & Trim(vlCodAl) & "' AND "
    vgSql = vgSql & "m.cod_elemento = p.cod_modalidad AND "
    vgSql = vgSql & "p.cod_cobercon = b.cod_cobercon"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro.EOF) Then
        vlCobertura = vlRegistro!Gls_Renta
        If vlRegistro!Cod_Modalidad = 1 Then
            If Not IsNull(vlRegistro!Gls_Modalidad) Then
                vlCobertura = vlCobertura & " " & vlRegistro!Gls_Modalidad
            End If
        Else
            If Not IsNull(vlRegistro!Gls_Modalidad) Then
                vlCobertura = vlCobertura & " CON P. " & vlRegistro!Gls_Modalidad
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
    Else
        vlCobertura = ""
    End If
    vlRegistro.Close
    
Exit Sub
Err_buscaAnt:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

