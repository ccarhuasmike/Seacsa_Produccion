VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Frm_CalPolizaRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recálculo de Pólizas."
   ClientHeight    =   8040
   ClientLeft      =   2325
   ClientTop       =   2295
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10665
   Begin VB.Frame Fra_Cabeza 
      Height          =   1485
      Left            =   120
      TabIndex        =   99
      Top             =   120
      Width           =   10455
      Begin VB.TextBox Txt_NumCot 
         BackColor       =   &H00E0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   885
         Width           =   1695
      End
      Begin VB.TextBox Txt_NumPol 
         BackColor       =   &H00E0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   290
         Width           =   1695
      End
      Begin VB.Label Lbl_FecCalculo 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8520
         TabIndex        =   3
         Top             =   285
         Width           =   1695
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Fecha Cálculo"
         Height          =   255
         Index           =   5
         Left            =   7200
         TabIndex        =   208
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Lbl_SecOfe 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8520
         TabIndex        =   6
         Top             =   885
         Width           =   495
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Nº Correlativo"
         Height          =   255
         Index           =   4
         Left            =   7200
         TabIndex        =   206
         Top             =   885
         Width           =   1095
      End
      Begin VB.Label Lbl_FecVig 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5160
         TabIndex        =   2
         Top             =   285
         Width           =   1695
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Nº de Cotizacion"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   103
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Nº Poliza"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   102
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Fecha Emisión"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   101
         Top             =   285
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   360
         X2              =   9960
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Nº Operación"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   100
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Lbl_SolOfe 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5160
         TabIndex        =   5
         Top             =   885
         Width           =   1695
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   10455
      Begin VB.CommandButton Cmd_Modificar 
         Height          =   675
         Left            =   3720
         Picture         =   "Frm_CalPolizaRec.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Modificar lo Calculado"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Calcular 
         Caption         =   "&Calcular"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_CalPolizaRec.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Calcular Pensión"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   6000
         Picture         =   "Frm_CalPolizaRec.frx":08E4
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Traspasar"
         Height          =   675
         Left            =   4800
         Picture         =   "Frm_CalPolizaRec.frx":0F9E
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Traspasar Datos a la Póliza"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   7080
         Picture         =   "Frm_CalPolizaRec.frx":13E0
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Cancelar Modificaciones"
         Top             =   240
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab_Poliza 
      Height          =   5175
      Left            =   120
      TabIndex        =   104
      Top             =   1680
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9128
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos del Afiliado"
      TabPicture(0)   =   "Frm_CalPolizaRec.frx":19BA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Afiliado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos de Cálculo"
      TabPicture(1)   =   "Frm_CalPolizaRec.frx":19D6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_Calculo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Datos de Beneficiarios"
      TabPicture(2)   =   "Frm_CalPolizaRec.frx":19F2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Fra_Benef"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Fra_Benef 
         Height          =   4605
         Left            =   -74880
         TabIndex        =   170
         Top             =   480
         Width           =   10215
         Begin VB.Frame Fra_DatosBenef 
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
            TabIndex        =   171
            Top             =   1560
            Width           =   9975
            Begin VB.TextBox Txt_FecInvBen 
               Height          =   285
               Left            =   6120
               MaxLength       =   25
               TabIndex        =   85
               Top             =   1400
               Width           =   1095
            End
            Begin VB.CommandButton Cmd_LimpiarBen 
               Height          =   450
               Left            =   9360
               Picture         =   "Frm_CalPolizaRec.frx":1A0E
               Style           =   1  'Graphical
               TabIndex        =   93
               ToolTipText     =   "Habilitar Ingreso Nvos Benef."
               Top             =   1680
               Width           =   495
            End
            Begin VB.TextBox Txt_FecFallBen 
               Height          =   285
               Left            =   3600
               MaxLength       =   20
               TabIndex        =   78
               Top             =   1875
               Width           =   1095
            End
            Begin VB.TextBox Txt_FecNacBen 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   77
               Top             =   1875
               Width           =   1095
            End
            Begin VB.ComboBox Cmb_SitInv 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":2080
               Left            =   6120
               List            =   "Frm_CalPolizaRec.frx":2082
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   1095
               Width           =   2895
            End
            Begin VB.ComboBox Cmb_SexoBen 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":2084
               Left            =   6120
               List            =   "Frm_CalPolizaRec.frx":2086
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   800
               Width           =   2895
            End
            Begin VB.ComboBox Cmb_GrupoFam 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":2088
               Left            =   6120
               List            =   "Frm_CalPolizaRec.frx":208A
               Style           =   2  'Dropdown List
               TabIndex        =   82
               Top             =   480
               Width           =   2895
            End
            Begin VB.ComboBox Cmb_Parentesco 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":208C
               Left            =   6120
               List            =   "Frm_CalPolizaRec.frx":208E
               Style           =   2  'Dropdown List
               TabIndex        =   81
               Top             =   170
               Width           =   2895
            End
            Begin VB.CommandButton Cmd_BuscaCauInv 
               Caption         =   "?"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9000
               TabIndex        =   87
               Top             =   1680
               Width           =   285
            End
            Begin VB.TextBox Txt_NombresBenSeg 
               Height          =   285
               Left            =   1320
               MaxLength       =   25
               TabIndex        =   74
               Top             =   1030
               Width           =   3375
            End
            Begin VB.TextBox Txt_NumIdentBen 
               Height          =   285
               Left            =   1320
               MaxLength       =   25
               TabIndex        =   72
               Top             =   480
               Width           =   2295
            End
            Begin VB.ComboBox Cmb_TipoIdentBen 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":2090
               Left            =   1320
               List            =   "Frm_CalPolizaRec.frx":2092
               Style           =   2  'Dropdown List
               TabIndex        =   71
               Top             =   180
               Width           =   3375
            End
            Begin VB.CommandButton Btn_Porcentaje 
               Height          =   450
               Left            =   9360
               Picture         =   "Frm_CalPolizaRec.frx":2094
               Style           =   1  'Graphical
               TabIndex        =   92
               ToolTipText     =   "Calcular Porcentajes"
               Top             =   1200
               Width           =   495
            End
            Begin VB.CommandButton Btn_Quita 
               Height          =   450
               Left            =   9360
               Picture         =   "Frm_CalPolizaRec.frx":2656
               Style           =   1  'Graphical
               TabIndex        =   91
               ToolTipText     =   "Eliminar Beneficiario"
               Top             =   720
               Width           =   495
            End
            Begin VB.CommandButton Btn_Agregar 
               Height          =   450
               Left            =   9360
               Picture         =   "Frm_CalPolizaRec.frx":27E0
               Style           =   1  'Graphical
               TabIndex        =   90
               ToolTipText     =   "Agregar Beneficiario"
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox Txt_ApMatBen 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   76
               Top             =   1590
               Width           =   3375
            End
            Begin VB.TextBox Txt_ApPatBen 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   75
               Top             =   1320
               Width           =   3375
            End
            Begin VB.TextBox Txt_NombresBen 
               Height          =   285
               Left            =   1320
               MaxLength       =   25
               TabIndex        =   73
               Top             =   760
               Width           =   3375
            End
            Begin VB.Label Lbl_Moneda 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   197
               Top             =   2475
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label Lbl_Moneda 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   196
               Top             =   2175
               Width           =   375
            End
            Begin VB.Label Lbl_DerPension 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6120
               TabIndex        =   88
               Top             =   1950
               Width           =   2895
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "2do. Nombre"
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   191
               Top             =   1030
               Width           =   915
            End
            Begin VB.Label Lbl_Nombre 
               AutoSize        =   -1  'True
               Caption         =   "Nº Ident."
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   190
               Top             =   480
               Width           =   630
            End
            Begin VB.Label Lbl_CauInvBen 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6120
               TabIndex        =   86
               Top             =   1680
               Width           =   2895
            End
            Begin VB.Label Lbl_PenGar 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   80
               Top             =   2475
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Lbl_PensionBen 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   79
               Top             =   2175
               Width           =   1095
            End
            Begin VB.Label Lbl_Porcentaje 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6120
               TabIndex        =   89
               Top             =   2235
               Width           =   1095
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   240
               TabIndex        =   189
               Top             =   0
               Width           =   1020
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Ident."
               Height          =   195
               Index           =   15
               Left            =   120
               TabIndex        =   188
               Top             =   220
               Width           =   765
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Pensión Gar."
               Height          =   195
               Index           =   14
               Left            =   120
               TabIndex        =   187
               Top             =   2475
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label Lbl_NumOrden 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   4320
               TabIndex        =   186
               Top             =   480
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Grupo Familiar"
               Height          =   255
               Index           =   9
               Left            =   4920
               TabIndex        =   185
               Top             =   525
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Dº a Pensión"
               Height          =   255
               Index           =   8
               Left            =   4920
               TabIndex        =   184
               Top             =   1950
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Fec. Nac."
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   183
               Top             =   1875
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Sit. Invalidez"
               Height          =   255
               Index           =   6
               Left            =   4920
               TabIndex        =   182
               Top             =   1095
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Parentesco"
               Height          =   255
               Index           =   5
               Left            =   4920
               TabIndex        =   181
               Top             =   220
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Sexo"
               Height          =   255
               Index           =   4
               Left            =   4920
               TabIndex        =   180
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Ap. Materno"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   179
               Top             =   1590
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Ap. Paterno"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   178
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "1er.  Nombre"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   177
               Top             =   765
               Width           =   915
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Porcentaje"
               Height          =   255
               Index           =   7
               Left            =   4920
               TabIndex        =   176
               Top             =   2235
               Width           =   855
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Fec. Fallec."
               Height          =   195
               Index           =   10
               Left            =   2640
               TabIndex        =   175
               Top             =   1920
               Width           =   825
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Pensión"
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   174
               Top             =   2175
               Width           =   855
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Fec. Invalidez"
               Height          =   255
               Index           =   12
               Left            =   4920
               TabIndex        =   173
               Top             =   1400
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Causal Invalidez"
               Height          =   255
               Index           =   13
               Left            =   4920
               TabIndex        =   172
               Top             =   1680
               Width           =   1215
            End
         End
         Begin MSFlexGridLib.MSFlexGrid Msf_GriAseg 
            Height          =   1335
            Left            =   120
            TabIndex        =   192
            TabStop         =   0   'False
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   2355
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
         Height          =   4455
         Left            =   120
         TabIndex        =   141
         Top             =   480
         Width           =   10215
         Begin VB.TextBox Txt_Asegurados 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4440
            TabIndex        =   9
            Top             =   500
            Width           =   495
         End
         Begin VB.TextBox Txt_NumIdent 
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   8
            Top             =   500
            Width           =   2055
         End
         Begin VB.ComboBox Cmb_TipoIdent 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPolizaRec.frx":296A
            Left            =   1560
            List            =   "Frm_CalPolizaRec.frx":296C
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   200
            Width           =   3375
         End
         Begin VB.TextBox Txt_Nacionalidad 
            Height          =   285
            Left            =   6600
            MaxLength       =   40
            TabIndex        =   31
            Top             =   1950
            Width           =   3255
         End
         Begin VB.ComboBox Cmb_Vejez 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPolizaRec.frx":296E
            Left            =   6600
            List            =   "Frm_CalPolizaRec.frx":2970
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   2240
            Width           =   3255
         End
         Begin VB.CommandButton Cmd_BuscarDir 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9870
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Efectuar Busqueda de Dirección"
            Top             =   1080
            Width           =   300
         End
         Begin VB.ComboBox Cmb_TipoPension 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPolizaRec.frx":2972
            Left            =   1560
            List            =   "Frm_CalPolizaRec.frx":2974
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2520
            Width           =   3375
         End
         Begin VB.TextBox Txt_FecFall 
            Height          =   285
            Left            =   3840
            MaxLength       =   20
            TabIndex        =   16
            Top             =   2250
            Width           =   1095
         End
         Begin VB.TextBox Txt_FecNac 
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   15
            Top             =   2250
            Width           =   1095
         End
         Begin VB.ComboBox Cmb_Afp 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPolizaRec.frx":2976
            Left            =   1560
            List            =   "Frm_CalPolizaRec.frx":2978
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   3420
            Width           =   3375
         End
         Begin VB.ComboBox Cmb_Sexo 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPolizaRec.frx":297A
            Left            =   1560
            List            =   "Frm_CalPolizaRec.frx":297C
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1930
            Width           =   3375
         End
         Begin VB.TextBox Txt_NomAfi 
            Height          =   285
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   10
            Top             =   800
            Width           =   3375
         End
         Begin VB.TextBox Txt_ApPatAfi 
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   12
            Top             =   1380
            Width           =   3375
         End
         Begin VB.TextBox Txt_ApMatAfi 
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1660
            Width           =   3375
         End
         Begin VB.TextBox Txt_Fono 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   29
            Top             =   1400
            Width           =   1815
         End
         Begin VB.TextBox Txt_Dir 
            Height          =   285
            Left            =   6600
            MaxLength       =   50
            TabIndex        =   24
            Top             =   200
            Width           =   3250
         End
         Begin VB.ComboBox Cmb_Salud 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   4040
            Width           =   3375
         End
         Begin VB.TextBox Txt_Correo 
            Height          =   285
            Left            =   6600
            MaxLength       =   40
            TabIndex        =   30
            Top             =   1680
            Width           =   3250
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
            TabIndex        =   142
            Top             =   2600
            Width           =   4455
            Begin VB.ComboBox Cmb_Bco 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   1140
               Width           =   3300
            End
            Begin VB.ComboBox Cmb_TipCta 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   840
               Width           =   3300
            End
            Begin VB.ComboBox Cmb_Suc 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   540
               Width           =   3315
            End
            Begin VB.ComboBox Cmb_ViaPago 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   240
               Width           =   3315
            End
            Begin VB.TextBox Txt_NumCta 
               Height          =   285
               Left            =   960
               MaxLength       =   15
               TabIndex        =   37
               Top             =   1450
               Width           =   3300
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Sucursal"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   147
               Top             =   560
               Width           =   810
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "N°Cuenta"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   146
               Top             =   1515
               Width           =   795
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Banco"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   145
               Top             =   1185
               Width           =   825
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Vía Pago"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   144
               Top             =   240
               Width           =   810
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Tipo Cta."
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   143
               Top             =   870
               Width           =   825
            End
         End
         Begin VB.CommandButton Cmd_CauInv 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4940
            TabIndex        =   20
            Top             =   3150
            Width           =   285
         End
         Begin VB.TextBox Txt_NomAfiSeg 
            Height          =   285
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   11
            Top             =   1090
            Width           =   3375
         End
         Begin VB.TextBox Txt_FecInv 
            Height          =   285
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   18
            Top             =   2850
            Width           =   1095
         End
         Begin VB.ComboBox Cmb_EstCivil 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPolizaRec.frx":297E
            Left            =   1560
            List            =   "Frm_CalPolizaRec.frx":2980
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   3730
            Width           =   3375
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Nacionalidad"
            Height          =   255
            Index           =   21
            Left            =   5400
            TabIndex        =   203
            Top             =   1950
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Vejez"
            Height          =   195
            Index           =   22
            Left            =   5400
            TabIndex        =   202
            Top             =   2240
            Width           =   750
         End
         Begin VB.Label Lbl_Distrito 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   27
            Top             =   1080
            Width           =   3250
         End
         Begin VB.Label Lbl_Provincia 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   26
            Top             =   780
            Width           =   3255
         End
         Begin VB.Label Lbl_Departamento 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   25
            Top             =   480
            Width           =   3250
         End
         Begin VB.Label Lbl_CauInv 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1560
            TabIndex        =   19
            Top             =   3150
            Width           =   3375
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "1er. Nombre"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   169
            Top             =   800
            Width           =   975
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Distrito"
            Height          =   255
            Index           =   12
            Left            =   5400
            TabIndex        =   168
            Top             =   1090
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   11
            Left            =   5400
            TabIndex        =   167
            Top             =   790
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   10
            Left            =   5400
            TabIndex        =   166
            Top             =   200
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   9
            Left            =   5400
            TabIndex        =   165
            Top             =   1400
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Est. Civil"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   164
            Top             =   3730
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Fecha Nac."
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   163
            Top             =   2250
            Width           =   1095
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Sexo"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   162
            Top             =   1950
            Width           =   975
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Ap. Materno"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   161
            Top             =   1650
            Width           =   975
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Ap. Paterno"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   160
            Top             =   1380
            Width           =   975
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Inst. Salud"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   159
            Top             =   4040
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Nº Benef."
            Height          =   255
            Index           =   13
            Left            =   3720
            TabIndex        =   158
            Top             =   525
            Width           =   735
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "AFP"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   157
            Top             =   3420
            Width           =   1140
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Departamento"
            Height          =   255
            Index           =   15
            Left            =   5400
            TabIndex        =   156
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Correo"
            Height          =   255
            Index           =   14
            Left            =   5400
            TabIndex        =   155
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Fecha Fallec."
            Height          =   255
            Index           =   16
            Left            =   2760
            TabIndex        =   154
            Top             =   2250
            Width           =   1095
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Tipo Pensión"
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   153
            Top             =   2550
            Width           =   1095
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Fec. Invalidez"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   152
            Top             =   2850
            Width           =   1095
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Causal Invalidez"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   151
            Top             =   3150
            Width           =   1335
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Identificación"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   150
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Nº Identificación"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   149
            Top             =   510
            Width           =   1170
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "2do. Nombre"
            Height          =   195
            Index           =   20
            Left            =   240
            TabIndex        =   148
            Top             =   1090
            Width           =   915
         End
      End
      Begin VB.Frame Fra_Calculo 
         Height          =   4660
         Left            =   -74880
         TabIndex        =   105
         Top             =   360
         Width           =   10215
         Begin VB.Frame Fra_SumaBono 
            Height          =   615
            Left            =   120
            TabIndex        =   135
            Top             =   3920
            Width           =   9975
            Begin VB.TextBox Txt_ApoAdi 
               Height          =   285
               Left            =   5640
               MaxLength       =   20
               TabIndex        =   209
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox Txt_BonoAct 
               Height          =   285
               Left            =   3120
               MaxLength       =   20
               TabIndex        =   69
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox Txt_CtaInd 
               Height          =   285
               Left            =   720
               MaxLength       =   20
               TabIndex        =   68
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "AA."
               Height          =   195
               Index           =   6
               Left            =   4920
               TabIndex        =   212
               Top             =   240
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
               Index           =   5
               Left            =   4680
               TabIndex        =   211
               Top             =   195
               Width           =   255
            End
            Begin VB.Label Lbl_MonedaFon 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   3
               Left            =   5280
               TabIndex        =   210
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Lbl_PriUnica 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   8400
               TabIndex        =   70
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Lbl_MonedaFon 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   2
               Left            =   8040
               TabIndex        =   201
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Lbl_MonedaFon 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   1
               Left            =   2760
               TabIndex        =   200
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Lbl_MonedaFon 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   0
               Left            =   315
               TabIndex        =   199
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
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
               Height          =   300
               Index           =   4
               Left            =   7200
               TabIndex        =   140
               Top             =   195
               Width           =   165
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
               Left            =   2160
               TabIndex        =   139
               Top             =   195
               Width           =   255
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "BR."
               Height          =   195
               Index           =   0
               Left            =   2400
               TabIndex        =   138
               Top             =   255
               Width           =   270
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "CI."
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   137
               Top             =   255
               Width           =   195
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "Prima U."
               Height          =   195
               Index           =   2
               Left            =   7440
               TabIndex        =   136
               Top             =   255
               Width           =   600
            End
         End
         Begin VB.Frame Fra_DatCal 
            Height          =   3800
            Left            =   120
            TabIndex        =   106
            Top             =   120
            Width           =   9975
            Begin VB.ComboBox Cmb_IndCob 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":2982
               Left            =   2040
               List            =   "Frm_CalPolizaRec.frx":2984
               Style           =   2  'Dropdown List
               TabIndex        =   49
               Top             =   3440
               Width           =   1095
            End
            Begin VB.TextBox Txt_Cuspp 
               Height          =   285
               Left            =   2040
               MaxLength       =   18
               TabIndex        =   41
               Top             =   1040
               Width           =   2655
            End
            Begin VB.ComboBox Cmb_DerGra 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":2986
               Left            =   6550
               List            =   "Frm_CalPolizaRec.frx":2988
               Style           =   2  'Dropdown List
               TabIndex        =   52
               Top             =   800
               Width           =   1095
            End
            Begin VB.ComboBox Cmb_DerCre 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":298A
               Left            =   6550
               List            =   "Frm_CalPolizaRec.frx":298C
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   500
               Width           =   1095
            End
            Begin VB.ComboBox Cmb_CobConyuge 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":298E
               Left            =   6550
               List            =   "Frm_CalPolizaRec.frx":2990
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   200
               Width           =   1095
            End
            Begin VB.ComboBox Cmb_Moneda 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":2992
               Left            =   2040
               List            =   "Frm_CalPolizaRec.frx":2994
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox Txt_FecIniPago 
               Height          =   285
               Left            =   2040
               MaxLength       =   10
               TabIndex        =   40
               Top             =   750
               Width           =   1095
            End
            Begin VB.TextBox Txt_FecIncorpora 
               Height          =   285
               Left            =   2040
               MaxLength       =   10
               TabIndex        =   39
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox Txt_PrcRentaTmp 
               Height          =   285
               Left            =   2040
               MaxLength       =   15
               TabIndex        =   48
               Top             =   3150
               Width           =   975
            End
            Begin VB.TextBox Txt_MesesGar 
               Height          =   285
               Left            =   2040
               MaxLength       =   4
               TabIndex        =   46
               Top             =   2580
               Width           =   735
            End
            Begin VB.TextBox Txt_AnnosDif 
               Height          =   285
               Left            =   2040
               MaxLength       =   3
               TabIndex        =   44
               Top             =   1960
               Width           =   735
            End
            Begin VB.ComboBox Cmb_Modalidad 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":2996
               Left            =   2040
               List            =   "Frm_CalPolizaRec.frx":2998
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   2250
               Width           =   2655
            End
            Begin VB.ComboBox Cmb_TipoRenta 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPolizaRec.frx":299A
               Left            =   2040
               List            =   "Frm_CalPolizaRec.frx":299C
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   1640
               Width           =   2655
            End
            Begin VB.TextBox Txt_FecDev 
               Height          =   285
               Left            =   2040
               MaxLength       =   10
               TabIndex        =   38
               Top             =   200
               Width           =   1095
            End
            Begin VB.TextBox Txt_PrcFam 
               Height          =   285
               Left            =   9240
               MaxLength       =   10
               TabIndex        =   67
               Top             =   3480
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox Txt_ComInt 
               Height          =   285
               Left            =   6550
               MaxLength       =   10
               TabIndex        =   53
               Top             =   1105
               Width           =   1095
            End
            Begin VB.CommandButton Cmd_BuscaCor 
               Caption         =   "?"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9560
               TabIndex        =   57
               Top             =   1410
               Width           =   285
            End
            Begin VB.Label Lbl_ComIntBen 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   9120
               TabIndex        =   54
               Top             =   720
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Com Inter C/Benef"
               Height          =   195
               Index           =   23
               Left            =   7800
               TabIndex        =   207
               Top             =   720
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Lbl_Moneda 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   3
               Left            =   7800
               TabIndex        =   205
               Top             =   3120
               Width           =   375
            End
            Begin VB.Label Lbl_Moneda 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   4
               Left            =   7800
               TabIndex        =   204
               Top             =   3420
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Ben. Social"
               Height          =   255
               Index           =   24
               Left            =   8250
               TabIndex        =   198
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Lbl_BenSocial 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Left            =   9105
               TabIndex        =   55
               Top             =   1080
               Width           =   480
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Gratificación"
               Height          =   195
               Index           =   19
               Left            =   4920
               TabIndex        =   195
               Top             =   800
               Width           =   885
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Derecho Crecer"
               Height          =   195
               Index           =   18
               Left            =   4920
               TabIndex        =   194
               Top             =   500
               Width           =   1125
            End
            Begin VB.Label Lbl_RentaAFP 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   2040
               TabIndex        =   47
               Top             =   2870
               Width           =   975
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Moneda"
               Height          =   195
               Index           =   17
               Left            =   240
               TabIndex        =   193
               Top             =   1320
               Width           =   945
            End
            Begin VB.Label Lbl_NumIdentCorr 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6555
               TabIndex        =   58
               Top             =   1680
               Width           =   3015
            End
            Begin VB.Label Lbl_TipoIdentCorr 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6555
               TabIndex        =   56
               Top             =   1395
               Width           =   3015
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Factor CNU"
               Height          =   255
               Index           =   7
               Left            =   8280
               TabIndex        =   134
               Top             =   3480
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Ident. Inter."
               Height          =   195
               Index           =   2
               Left            =   4905
               TabIndex        =   133
               Top             =   1400
               Width           =   1170
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Años Dif."
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   132
               Top             =   1960
               Width           =   735
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Fec. Devengue"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   131
               Top             =   200
               Width           =   1695
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Meses Gar."
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   130
               Top             =   2580
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Modalidad"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   129
               Top             =   2250
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tipo de Renta"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   128
               Top             =   1640
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Cobertura Cónyuge"
               Height          =   195
               Index           =   9
               Left            =   4920
               TabIndex        =   127
               Top             =   200
               Width           =   1365
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Prc. Renta Temporal"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   126
               Top             =   3150
               Width           =   1575
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Rentabilidad AFP"
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   125
               Top             =   2870
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tasa Cto. Equiv."
               Height          =   255
               Index           =   8
               Left            =   4905
               TabIndex        =   124
               Top             =   1965
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tasa Venta"
               Height          =   255
               Index           =   11
               Left            =   4905
               TabIndex        =   123
               Top             =   2265
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tasa TIR"
               Height          =   255
               Index           =   13
               Left            =   4905
               TabIndex        =   122
               Top             =   2565
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Mto. Pensión "
               Height          =   195
               Index           =   14
               Left            =   4905
               TabIndex        =   121
               Top             =   3165
               Width           =   975
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Mto. Pensión Gar. "
               Height          =   195
               Index           =   15
               Left            =   4905
               TabIndex        =   120
               Top             =   3465
               Visible         =   0   'False
               Width           =   1320
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Comisión Inter."
               Height          =   195
               Index           =   16
               Left            =   4920
               TabIndex        =   119
               Top             =   1110
               Width           =   1035
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Tasa Int. Período Gar."
               Height          =   195
               Index           =   17
               Left            =   4905
               TabIndex        =   118
               Top             =   2865
               Width           =   1590
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   20
               Left            =   3120
               TabIndex        =   117
               Top             =   2870
               Width           =   255
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   21
               Left            =   3120
               TabIndex        =   116
               Top             =   3150
               Width           =   255
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   22
               Left            =   7800
               TabIndex        =   115
               Top             =   1110
               Width           =   165
            End
            Begin VB.Label Lbl_MtoPrimaUniSim 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   8640
               TabIndex        =   65
               Top             =   2520
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Lbl_MtoPrimaUniDif 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   8640
               TabIndex        =   66
               Top             =   3120
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Lbl_TasaCtoEq 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6555
               TabIndex        =   59
               Top             =   1965
               Width           =   735
            End
            Begin VB.Label Lbl_TasaVta 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6555
               TabIndex        =   60
               Top             =   2265
               Width           =   735
            End
            Begin VB.Label Lbl_TasaTIR 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6555
               TabIndex        =   61
               Top             =   2565
               Width           =   735
            End
            Begin VB.Label Lbl_TasaPerGar 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6555
               TabIndex        =   62
               Top             =   2865
               Width           =   735
            End
            Begin VB.Label Lbl_MtoPension 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6555
               TabIndex        =   63
               Top             =   3165
               Width           =   1215
            End
            Begin VB.Label Lbl_MtoPensionGar 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6555
               TabIndex        =   64
               Top             =   3465
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Mto. Prima Uni Sim"
               Height          =   255
               Index           =   60
               Left            =   8400
               TabIndex        =   114
               Top             =   2280
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Mto. Prima Uni Dif"
               Height          =   255
               Index           =   61
               Left            =   8400
               TabIndex        =   113
               Top             =   2880
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Nº Ident. Inter."
               Height          =   255
               Index           =   7
               Left            =   4920
               TabIndex        =   112
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "CUSSP"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   111
               Top             =   1040
               Width           =   540
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Estado Cobertura"
               Height          =   255
               Index           =   37
               Left            =   240
               TabIndex        =   110
               Top             =   3480
               Width           =   1695
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Fec. Incorp. a la Póliza"
               Height          =   255
               Index           =   38
               Left            =   240
               TabIndex        =   109
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Fec. Ini. Primer Pago"
               Height          =   255
               Index           =   39
               Left            =   240
               TabIndex        =   108
               Top             =   765
               Width           =   1695
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   40
               Left            =   7800
               TabIndex        =   107
               Top             =   240
               Width           =   255
            End
         End
      End
   End
End
Attribute VB_Name = "Frm_CalPolizaRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaración de Constantes para Procesos Varios
Const clCodCauInvNoInv As String * 1 = "0"
Const clCodSitInvNoInv As String * 1 = "N"

'Declaración de Constantes para la Moneda de la Renta
Const clMonedaModalidad As Integer = 0
Const clMonedaBenPen    As Integer = 1
Const clMonedaBenPenGar As Integer = 2
Const clMonedaModPen    As Integer = 3
Const clMonedaModPenGar As Integer = 4

Const clMtoCtaIndFon    As Integer = 0
Const clMtoBonoFon      As Integer = 1
Const clMtoPriUniFon    As Integer = 2
Const clMtoApoAdiFon    As Integer = 3

Dim vlBotonEscogido As String
Dim vlSwCalIntOK As Boolean 'Sw de boton de calculo interior
Dim vlSwCalExtOK As Boolean 'Sw de boton de calculo exterior

'Variables de la Póliza
Dim vlNumPol As String, vlNumCot As String, vlNumCorrelativo As String
Dim vlCodAFP As String, vlCodIsapre As String
Dim vlCodTipPen As String, vlCodVejez As String
Dim vlCodEstCivil As String
Dim vlNumCargas As Integer
Dim vlIndCobPol As String, vlCodDerCrePol As String
Dim vlNumMesGar As Long, vlNumMesDif As Long
Dim vlMtoPensionRef As Double

'Datos del Afiliado
Dim vlTipoIden As String, vlNumIden As String

'Variables para Beneficiarios
Dim vlNumOrden As String, vlNumOrdenCot As String
Dim vlCodPar As String, vlCodGruFam As String, vlCodSexoBen As String
Dim vlCodSitInv As String, vlCodDerPen As String, vlCodDerCre As String
Dim vlCauInv As String, vlFecInv As String, vlRutBen As String
Dim vlDgvBen As String, vlNomBen As String, vlNomBenSeg As String
Dim vlPatBen As String, vlMatBen As String, vlEstPen As String
Dim vlFecFallBen As String, vlFecNacHM As String, vlFecNacBen As String
Dim vlPrcPension As Double, vlPenBen As Double, vlPenGarBen As Double
Dim vlPrcPensionLeg As Double
Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla
'Dim vlNumIdent         As Integer
Dim vlCodTipoIden As Integer 'sirve para guardar el código
Dim vlValCodSitInv As String
Dim vlPrcPensionGar As Double

Dim vlSwEncontrado      As Boolean 'Sw de boton restar beneficiario
Dim vlPos As Integer
Dim vlNumero As Long
Dim vlOpcion As String

'Variables para Cálculo
Dim vlFecVig As String, vlFecDev As String
Dim vlFecEmision As String, vlFecCalculo As String 'ABV 08/07/2007
Dim vlElemento As String
Dim vlCodDireccion As String, vlNacionalidad As String
Dim vlFecIncorporacion As String, vlFecPriPago As String
Dim vlcuspp As String, vlFecRepPrimasEst As String

Dim vlSwAfp As Boolean '22-09-2007

'I--- ABV 04/12/2009 ---
Dim vlMarcaSobDif As String
'F--- ABV 04/12/2009 ---

'Variables para el Beneficiario

Function flRecibeDireccion(iNomDepartamento As String, iNomProvincia As String, iNomDistrito As String, iCodDir As String)
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA de Dirección

    Lbl_Departamento = Trim(iNomDepartamento)
    Lbl_Provincia = Trim(iNomProvincia)
    Lbl_Distrito = Trim(iNomDistrito)
    vlCodDireccion = iCodDir
    vgCodDireccion = iCodDir
    Txt_Fono.SetFocus
    
End Function

Function flRecibeCorredor(iNomTipoIden As String, iNumIden As String, iCodTipoIden As String, iBenSocial As String, iPrcComision As Double)
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA del Corredor

    Lbl_TipoIdentCorr = iCodTipoIden & " - " & iNomTipoIden
    Lbl_NumIdentCorr = iNumIden
    If (iBenSocial <> "") Then
        Lbl_BenSocial = IIf(iBenSocial = "S", cgIndicadorSi, cgIndicadorNo)
    Else
        Lbl_BenSocial = cgIndicadorNo
    End If
    Txt_ComInt = Format(iPrcComision, "#0.00")
    Txt_CtaInd.SetFocus
    'If (Lbl_BenSocial = cgIndicadorNo) Then
    '    Lbl_ComIntBen = Txt_ComInt
    'End If

End Function

'-------------------------------------------------------------------------
'HABILITA O DESHABILITA LOS CAMPOS DE FORMA DE PAGO SEGUN LA VIA DE PAGO
'-------------------------------------------------------------------------
Function flValidaViaPago()
    vlCombo = Trim(Mid(Cmb_ViaPago, 1, (InStr(1, Cmb_ViaPago, "-") - 1)))
    If vlCombo = "01" Or vlCombo = "04" Then
        If (vlCombo = "04") Then
            vgTipoSucursal = cgTipoSucursalAfp
        Else
            vgTipoSucursal = cgTipoSucursalSuc
        End If
        fgComboSucursal Cmb_Suc, vgTipoSucursal

        Cmb_TipCta.Enabled = False
        Cmb_Bco.Enabled = False
        Txt_NumCta.Enabled = False
        Cmb_Suc.Enabled = True
        If (vlCombo = "04") Then
            vgPalabra = fgObtenerCodigo_TextoCompuesto(Cmb_Afp)
            Call fgBuscaPos(Cmb_Suc, vgPalabra)
        End If
        If Cmb_TipCta.ListCount <> 0 Then
            Cmb_TipCta.ListIndex = 0
        End If
        If Cmb_Bco.ListCount <> 0 Then
            Cmb_Bco.ListIndex = 0
        End If
    Else
        vgTipoSucursal = cgTipoSucursalSuc
        fgComboSucursal Cmb_Suc, vgTipoSucursal

        If vlCombo = "00" Or vlCombo = "05" Then
            Cmb_TipCta.Enabled = False
            Cmb_Bco.Enabled = False
            Txt_NumCta.Enabled = False
            Cmb_Suc.Enabled = False
            If Cmb_TipCta.ListCount <> 0 Then
                Cmb_TipCta.ListIndex = 0
            End If
            If Cmb_Bco.ListCount <> 0 Then
                Cmb_Bco.ListIndex = 0
            End If
        Else
            Cmb_TipCta.Enabled = True
            Cmb_Bco.Enabled = True
            Txt_NumCta.Enabled = True
            Cmb_Suc.Enabled = False
            If Cmb_Suc.ListCount <> 0 Then
                Cmb_Suc.ListIndex = 0
            End If
        End If
    End If

End Function

Function flEditarAfiliado()
Dim vlCodigo As String
Dim vlPosicion As Integer

    'Asigna los datos del formulario "Póliza" al formulario
    ' "PólizaRec" (EditarAfiliado)

    'Encabezado
    Txt_NumPol = Frm_CalPoliza.Txt_NumPol
    Lbl_FecVig = Frm_CalPoliza.Txt_FecVig
    Txt_NumCot = Frm_CalPoliza.Lbl_NumCot
    Lbl_SolOfe = Frm_CalPoliza.Lbl_SolOfe
    Lbl_SecOfe = Frm_CalPoliza.Lbl_SecOfe

    'Carpeta del Afiliado
    Cmb_TipoIdent.Text = Frm_CalPoliza.Cmb_TipoIdent.Text
'    If Not IsNull(vgRs!cod_tipoiden) Then Call fgBuscaPos(Cmb_TipoIdent, vgRs!cod_tipoiden)
'    If Not IsNull(vgRs!cod_estcivil) Then Call fgBuscaPos(Cmb_EstCivil, vgRs!cod_estcivil)
'    Call fgBuscaPos(Cmb_Salud, (vgRs!cod_isapre))


    Txt_NumIdent = Frm_CalPoliza.Txt_NumIdent
    Txt_Asegurados = Frm_CalPoliza.Txt_Asegurados
    Txt_NomAfi = Frm_CalPoliza.Txt_NomAfi
    Txt_NomAfiSeg = Frm_CalPoliza.Txt_NomAfiSeg
    Txt_ApPatAfi = Frm_CalPoliza.Txt_ApPatAfi
    Txt_ApMatAfi = Frm_CalPoliza.Txt_ApMatAfi
    Cmb_Sexo.Text = Frm_CalPoliza.Lbl_SexoAfi.Caption
    Txt_FecNac = Frm_CalPoliza.Lbl_FecNac
    Txt_FecFall = Frm_CalPoliza.Lbl_FecFall
    Cmb_TipoPension.Text = Frm_CalPoliza.Lbl_TipPen.Caption
    Txt_FecInv = Frm_CalPoliza.Txt_FecInv
    Lbl_CauInv = Frm_CalPoliza.Lbl_CauInv
    Cmb_Afp.Text = Frm_CalPoliza.Lbl_Afp.Caption
    Cmb_EstCivil.Text = Frm_CalPoliza.Cmb_EstCivil.Text
    Cmb_Salud.Text = Frm_CalPoliza.Cmb_Salud.Text
    Txt_Dir = Frm_CalPoliza.Lbl_Dir  'RVF 20090914
    Lbl_Departamento = Frm_CalPoliza.Lbl_Departamento
    Lbl_Provincia = Frm_CalPoliza.Lbl_Provincia
    Lbl_Distrito = Frm_CalPoliza.Lbl_Distrito
    Txt_Fono = Frm_CalPoliza.Txt_Fono
    Txt_Correo = Frm_CalPoliza.Txt_Correo
    Txt_Nacionalidad = Frm_CalPoliza.Txt_Nacionalidad
    Cmb_Vejez.Text = Frm_CalPoliza.Cmb_Vejez.Text
    Cmb_ViaPago.Text = Frm_CalPoliza.Cmb_ViaPago.Text
    Cmb_Suc.Text = Frm_CalPoliza.Cmb_Suc.Text
    Cmb_TipCta.Text = Frm_CalPoliza.Cmb_TipCta.Text
    Cmb_Bco.Text = Frm_CalPoliza.Cmb_Bco.Text
    Txt_NumCta = Frm_CalPoliza.Txt_NumCta

End Function

Function flEditarCalculo()

    'Carpeta de Cálculo
    Lbl_FecCalculo = Frm_CalPoliza.Lbl_FecCalculo
    Txt_FecDev = Frm_CalPoliza.Lbl_FecDev
    Txt_FecIncorpora = Frm_CalPoliza.Lbl_FecIncorpora
    Txt_FecIniPago = Frm_CalPoliza.Txt_FecIniPago
    Txt_Cuspp = Frm_CalPoliza.Lbl_CUSPP
    Cmb_Moneda = Frm_CalPoliza.Lbl_Moneda(0).Caption
    Cmb_TipoRenta.Text = Frm_CalPoliza.Lbl_TipoRenta.Caption
    Txt_AnnosDif = Frm_CalPoliza.Lbl_AnnosDif
    Cmb_Modalidad.Text = Frm_CalPoliza.Lbl_Alter.Caption
    Txt_MesesGar = Frm_CalPoliza.Lbl_MesesGar
    Lbl_RentaAFP = Frm_CalPoliza.Lbl_RentaAFP
    Txt_PrcRentaTmp = Frm_CalPoliza.Lbl_PrcRentaTmp
    If (Frm_CalPoliza.Lbl_IndCob = "Si") Then
        Cmb_IndCob = "S - " & Frm_CalPoliza.Lbl_IndCob
    Else
        Cmb_IndCob = "N - " & Frm_CalPoliza.Lbl_IndCob
    End If

    Cmb_CobConyuge.Text = Frm_CalPoliza.Lbl_FacPenElla
    If (Frm_CalPoliza.Lbl_DerCre = "Si") Then
        Cmb_DerCre = "S - " & Frm_CalPoliza.Lbl_DerCre
    Else
        Cmb_DerCre = "N - " & Frm_CalPoliza.Lbl_DerCre
    End If
    If (Frm_CalPoliza.Lbl_DerGra = "Si") Then
        Cmb_DerGra = "S - " & Frm_CalPoliza.Lbl_DerGra
    Else
        Cmb_IndCob = "N - " & Frm_CalPoliza.Lbl_DerGra
    End If
    Txt_ComInt = Frm_CalPoliza.Lbl_ComInt
    Lbl_ComIntBen = Frm_CalPoliza.Lbl_ComIntBen
    Lbl_BenSocial = Frm_CalPoliza.Lbl_BenSocial
    Lbl_TipoIdentCorr = Frm_CalPoliza.Lbl_TipoIdentCorr
    Lbl_NumIdentCorr = Frm_CalPoliza.Lbl_NumIdentCorr
    Lbl_TasaCtoEq = Frm_CalPoliza.Lbl_TasaCtoEq
    Lbl_TasaVta = Frm_CalPoliza.Lbl_TasaVta
    Lbl_TasaTIR = Frm_CalPoliza.Lbl_TasaTIR
    Lbl_TasaPerGar = Frm_CalPoliza.Lbl_TasaPerGar
    Lbl_MtoPension = Frm_CalPoliza.Lbl_MtoPension
    Lbl_MtoPensionGar = Frm_CalPoliza.Lbl_MtoPensionGar
    Lbl_MtoPrimaUniSim = Frm_CalPoliza.Lbl_MtoPrimaUniSim
    Lbl_MtoPrimaUniDif = Frm_CalPoliza.Lbl_MtoPrimaUniDif
    Txt_PrcFam = Frm_CalPoliza.Txt_PrcFam
    Txt_CtaInd = Frm_CalPoliza.Lbl_CtaInd
    Txt_BonoAct = Frm_CalPoliza.Lbl_BonoAct
    Txt_ApoAdi = Frm_CalPoliza.Lbl_ApoAdi
    Lbl_PriUnica = Frm_CalPoliza.Lbl_PriUnica

    Lbl_Moneda(clMonedaBenPen) = Frm_CalPoliza.Lbl_Moneda(clMonedaBenPen)
    Lbl_Moneda(clMonedaBenPenGar) = Frm_CalPoliza.Lbl_Moneda(clMonedaBenPenGar)

    Lbl_MonedaFon(clMtoCtaIndFon) = Frm_CalPoliza.Lbl_MonedaFon(clMtoCtaIndFon)
    Lbl_MonedaFon(clMtoBonoFon) = Frm_CalPoliza.Lbl_MonedaFon(clMtoBonoFon)
    Lbl_MonedaFon(clMtoPriUniFon) = Frm_CalPoliza.Lbl_MonedaFon(clMtoPriUniFon)
    Lbl_MonedaFon(clMtoApoAdiFon) = Frm_CalPoliza.Lbl_MonedaFon(clMtoApoAdiFon)

    Lbl_Moneda(clMonedaModPen) = Frm_CalPoliza.Lbl_Moneda(clMonedaModPen)
    Lbl_Moneda(clMonedaModPenGar) = Frm_CalPoliza.Lbl_Moneda(clMonedaModPenGar)

End Function

Function flEditarBeneficiarios()
Dim vlCasos As Long

    'Carpeta de Beneficiarios
    vlCasos = Frm_CalPoliza.Msf_GriAseg.rows
    vgI = 1

    While vgI <= vlCasos - 1

        With Frm_CalPoliza.Msf_GriAseg

            vlNumOrden = Trim(.TextMatrix(vgI, 0))
            vlCodPar = Trim(.TextMatrix(vgI, 1))
            vlCodGruFam = Trim(.TextMatrix(vgI, 2))
            vlCodSexoBen = Trim(.TextMatrix(vgI, 3))
            vlCodSitInv = Trim(.TextMatrix(vgI, 4))
            vlFecInv = Trim(.TextMatrix(vgI, 5))
            vlCauInv = Trim(.TextMatrix(vgI, 6))
            vlCodDerPen = Trim(.TextMatrix(vgI, 7))
            vlCodDerCre = Trim(.TextMatrix(vgI, 8))
            vlFecNacBen = Trim(.TextMatrix(vgI, 9))
            vlFecNacHM = Trim(.TextMatrix(vgI, 10))
            vlRutBen = " " & Trim(.TextMatrix(vgI, 11))
            vlDgvBen = Trim(.TextMatrix(vgI, 12))
            vlNomBen = Trim(.TextMatrix(vgI, 13))
            vlNomBenSeg = Trim(.TextMatrix(vgI, 14))
            vlPatBen = Trim(.TextMatrix(vgI, 15))
            vlMatBen = Trim(.TextMatrix(vgI, 16))
            vlPrcPension = Trim(.TextMatrix(vgI, 17))
            vlPenBen = Trim(.TextMatrix(vgI, 18))
            vlPenGarBen = Trim(.TextMatrix(vgI, 19))
            vlFecFallBen = Trim(.TextMatrix(vgI, 20))
            vlNumOrdenCot = Trim(.TextMatrix(vgI, 21))
            vlEstPen = Trim(.TextMatrix(vgI, 22))
            vlPrcPensionGar = Trim(.TextMatrix(vgI, 23))
            vlPrcPensionLeg = Trim(.TextMatrix(vgI, 24))

            Msf_GriAseg.AddItem (vlNumOrden) & vbTab & _
                        (vlCodPar) & vbTab & (vlCodGruFam) & vbTab & _
                        (vlCodSexoBen) & vbTab & (vlCodSitInv) & vbTab & _
                        (vlFecInv) & vbTab & (vlCauInv) & vbTab & _
                        (vlCodDerPen) & vbTab & (vlCodDerCre) & vbTab & _
                        (vlFecNacBen) & vbTab & (vlFecNacHM) & vbTab & _
                        (vlRutBen) & vbTab & (vlDgvBen) & vbTab & _
                        Trim(vlNomBen) & vbTab & Trim(vlNomBenSeg) & vbTab & _
                        Trim(vlPatBen) & vbTab & Trim(vlMatBen) & vbTab & _
                        Format(CDbl(vlPrcPension), "#,#0.000") & vbTab & _
                        Format(CDbl(vlPenBen), "#,#0.00") & vbTab & _
                        Format(CDbl(vlPenGarBen), "#,#0.00") & vbTab & _
                        Trim(vlFecFallBen) & vbTab & _
                        Trim(vlNumOrdenCot) & vbTab & vlEstPen _
                        & vbTab & vlPrcPensionGar & vbTab & vlPrcPensionLeg
        End With
        vgI = vgI + 1
    Wend

End Function

Function flLimpiarVariables()

'    vlNumBen = 0
'    vlTipoPen = ""
'    vlEstVigencia = ""
'    vlTipoRta = ""
'    vlModalidad = ""
'    vlPar = ""
'    vlCauInv = ""
'
'    vlRutAux = ""
'    vlPos = 0
''    vlTablaPoliza = ""
''    vlTablaBen = ""
''    vlTipoBuscar = ""
'    vlTipoBuscar = "N"
'    vlTablaPoliza = clTablaPolizaOri
'    vlTablaBen = clTablaBenOri
'
'    vlPalabraAux = ""
'    vlNumero = 0
'    vlOpcion = ""
'
'    vlSwValidaOK = False
    vlSwCalIntOK = False
    vlSwCalExtOK = False

'    vlSwGrabar = False
'    vlSwAprobar = False
'
'    'Beneficiario
'    vlNumOrden = 0
'    vlRutBen = ""
'    vlDgvBen = ""
'    vlNomBen = ""
'    vlPaternoBen = ""
'    vlMaternoBen = ""
'    vlCodPar = ""
'    vlCodGruFam = ""
'    vlCodSexo = ""
'    vlCodSitInv = ""
'    vlCodEstPension = ""
'    vlCodDerPen = ""
'    vlCodDerCre = ""
'    vlNumPoliza = ""
'    vlNumEndoso = 0
'    vlCodCauInv = ""
'    vlFecNacBen = ""
'    vlFecNacHM = ""
'    vlFecInvBen = ""
'    vlCodMotReqPen = ""
'    vlMtoPensionGar = 0
'    vlPrcPension = 0
'    vlFecFallBen = ""
'    vlCodCauSusBen = ""
'    vlFecSusBen = ""
'    vlFecIniPagoPen = ""
'    vlFecTerPagoPenGar = ""
'    vlFecMatrimonio = ""
'    'Poliza
'    vlCodTipPension = ""
'    vlCodEstado = ""
'    vlCodTipRen = ""
'    vlCodModalidad = ""
'    vlNumCargas = 0
'    vlFecVigencia = ""
'    vlFecTerVigencia = ""
'    vlMtoPrima = ""
'    vlMtoPension = 0
'    vlNumMesDif = 0
'    vlNumMesGar = 0
'    vlPrcTasaCe = 0
'    vlPrcTasaVta = 0
'    vlPrcTasaIntPerGar = 0
'    'Endoso
'    vlFecSolEndoso = ""
'    vlFecEndoso = ""
'    vlCodCauEndoso = ""
'    vlCodTipEndoso = ""
'    vlMtoDiferencia = 0
'
'    vlGlsUsuarioCrea = ""
'    vlFecCrea = 0
'    vlHorCrea = 0
'    vlGlsUsuarioModi = 0
'    vlFecModi = 0
'    vlHorModi = 0
'
End Function


Function flCargaDatosBeneficiariosMod(iPosicion As Integer)
On Error GoTo Err_flCargaDatosBeneficiariosMod

    'Posicionar Combos
    Msf_GriAseg.Row = iPosicion
    Msf_GriAseg.Col = 5
    Cmb_Parentesco.ListIndex = fgBuscarPosicionCodigoCombo(Trim(Msf_GriAseg.Text), Cmb_Parentesco)
    Msf_GriAseg.Col = 6
    Cmb_GrupoFam.ListIndex = fgBuscarPosicionCodigoCombo(Trim(Msf_GriAseg.Text), Cmb_GrupoFam)
    Msf_GriAseg.Col = 7
    Cmb_SexoBen.ListIndex = fgBuscarPosicionCodigoCombo(Trim(Msf_GriAseg.Text), Cmb_SexoBen)
    Msf_GriAseg.Col = 8
    Cmb_SitInv.ListIndex = fgBuscarPosicionCodigoCombo(Trim(Msf_GriAseg.Text), Cmb_SitInv)
    Msf_GriAseg.Col = 13
'    Cmb_BMCauInv.ListIndex = fgBuscarPosicionCodigoCombo(Trim(Msf_GriAseg.Text), Cmb_BMCauInv)
    Lbl_CauInvBen = Trim(Msf_GriAseg.Text) & " - " & Trim(fgBuscarGlosaCauInv(Trim(Msf_GriAseg.Text)))
    Msf_GriAseg.Col = 9
    'Cmb_BMDerPen.ListIndex = fgBuscarPosicionCodigoCombo(Trim(Msf_GriAseg.Text), Cmb_BMDerPen)
    If Msf_GriAseg.Text <> "" Then Lbl_DerPension = Trim(Msf_GriAseg.Text) & " - " & fgBuscarGlosaElemento(vgCodTabla_DerPen, Msf_GriAseg.Text)

    'Mostrar Datos
    Msf_GriAseg.Col = 0
    Lbl_NumOrden = Msf_GriAseg.Text

    Msf_GriAseg.Col = 11
    'Tipo Identificación
    If Msf_GriAseg.Text <> "" Then
        vgPalabra = Trim(Mid(Msf_GriAseg, 1, InStr(1, Msf_GriAseg, "-") - 1))
        vgI = fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_TipoIdentBen)
        If (Cmb_TipoIdentBen.ListCount > 0) Then
            Cmb_TipoIdentBen.ListIndex = vgI
        End If
    End If

    Msf_GriAseg.Col = 12
    'Número Identificación
    If Msf_GriAseg.Text <> "" Then
        Txt_NumIdentBen = Msf_GriAseg.Text
    End If

    Msf_GriAseg.Col = 2
    Txt_NombresBen = Trim(Msf_GriAseg.Text)
    Msf_GriAseg.Col = 2
    Txt_NombresBenSeg = Trim(Msf_GriAseg.Text)
    Msf_GriAseg.Col = 3
    Txt_ApPatBen = Trim(Msf_GriAseg.Text)
    Msf_GriAseg.Col = 4
    Txt_ApMatBen = Trim(Msf_GriAseg.Text)

    'I--- ABV 16/04/2005 ---
    'Las Fechas se encuentran formateadas de acuerdo a la configuración del PC
    'Por ende solo se deben traspasar las Fechas directamente
    Msf_GriAseg.Col = 16
    If (Msf_GriAseg.Text) = "" Then
       Txt_FecInvBen = ""
    Else
        Txt_FecInvBen = Trim(Msf_GriAseg.Text)
        'Txt_BMFecInv = DateSerial(Mid((Msf_GriAseg.Text), 1, 4), Mid((Msf_GriAseg.Text), 5, 2), Mid((Msf_GriAseg.Text), 7, 2))
    End If

    Msf_GriAseg.Col = 14
    Txt_FecNacBen = Trim(Msf_GriAseg.Text)
    'Txt_BMFecNac = DateSerial(Mid((Msf_GriAseg.Text), 1, 4), Mid((Msf_GriAseg.Text), 5, 2), Mid((Msf_GriAseg.Text), 7, 2))

    Msf_GriAseg.Col = 19
    If (Msf_GriAseg.Text) = "" Then
       Txt_FecFallBen = ""
    Else
        Txt_FecFallBen = Trim(Msf_GriAseg.Text)
        'Txt_BMFecFall = DateSerial(Mid((Msf_GriAseg.Text), 1, 4), Mid((Msf_GriAseg.Text), 5, 2), Mid((Msf_GriAseg.Text), 7, 2))
    End If

'    Msf_GriAseg.Col = 10
'    If (Msf_GriAseg.Text) = "" Then
'       Lbl_FecNHM = ""
'    Else
'        Lbl_FecNHM = Trim(Msf_GriAseg.Text)
'        'Lbl_BMFecNHM = DateSerial(Mid((Msf_GriAseg.Text), 1, 4), Mid((Msf_GriAseg.Text), 5, 2), Mid((Msf_GriAseg.Text), 7, 2))
'    End If

'    Msf_GriAseg.Col = 8
'    Lbl_DerAcrecer = Trim(Msf_GriAseg.Text)
    Msf_GriAseg.Col = 17
    Lbl_Porcentaje = Format(Msf_GriAseg.Text, "#0.00")
    Msf_GriAseg.Col = 18
    Lbl_PensionBen = Format(Msf_GriAseg.Text, "#,#0.00")
    Msf_GriAseg.Col = 19
    Lbl_PenGar = Format(Msf_GriAseg.Text, "#,#0.00")

'    Msf_GriAseg.Col = 11
'    vlNumPoliza = Trim(Msf_GriAseg.Text)
'    Msf_GriAseg.Col = 12
'    vlNumEndoso = Trim(Msf_GriAseg.Text)
    Msf_GriAseg.Col = 20
    vlCodEstPen = Trim(Msf_GriAseg.Text)

    'vlNumOrden = vlNumOrden

'    'I--- ABV 20/04/2005 ---
'    Call flHabilitarPenGar
'    'F--- ABV 20/04/2005 ---
'    Msf_GriAseg.Col = 22
'    vlMtoPensionGar = IIf(Trim(Msf_GriAseg.Text) = "", 0, Trim(Msf_GriAseg.Text))
'    Txt_BMMtoPensionGar = Format(vlMtoPensionGar, "#,#0.00")
'    'F--- ABV 19/04/2005 ---

Exit Function
Err_flCargaDatosBeneficiariosMod:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flValidaFecha(iFecha As String) As Boolean

    flValidaFecha = False

    If (Trim(iFecha) = "") Then
        Exit Function
    End If
    If Not IsDate(iFecha) Then
        Exit Function
    End If
    If (Year(CDate(iFecha)) < 1890) Then
        Exit Function
    End If

    flValidaFecha = True

End Function

'---------------------------------------------------------------------------
'VALIDA QUE ESTEN LLENOS LOS DATOS DEL BENEFICIRAIO
'---------------------------------------------------------------------------
Function flValidaDatosAseg() As Boolean
On Error GoTo Err_ValDatAseg

    flValidaDatosAseg = False

    If Trim(Cmb_TipoIdentBen) = "" Then
        MsgBox "Debe ingresar el Tipo de Identificación del Beneficiario.", vbExclamation, "Error de Datos"
        SSTab_Poliza.Tab = 2
        Cmb_TipoIdentBen.SetFocus
        Exit Function
    End If
    If Txt_NumIdentBen = "" Then
        MsgBox "Debe ingresar el Número de Identificación del Beneficiario.", vbExclamation, "Error de Datos"
        SSTab_Poliza.Tab = 2
        Txt_NumIdentBen.SetFocus
        Exit Function
    End If
    If Txt_NombresBen = "" Then
        MsgBox "Debe ingresar Nombre del Beneficiario.", vbExclamation, "Error de Datos"
        SSTab_Poliza.Tab = 2
        Txt_NombresBen.SetFocus
        Exit Function
    End If
'    If Txt_NombresBenSeg = "" Then
'        MsgBox "Debe ingresar el Segundo Nombre del Beneficiario.", vbExclamation, "Error de Datos"
'        SSTab_Poliza.Tab = 2
'        Txt_NombresBenSeg.SetFocus
'        Exit Function
'    End If
    If Txt_ApPatBen = "" Then
        MsgBox "Debe ingresar Apellido Paterno del Beneficiario.", vbExclamation, "Error de Datos"
        SSTab_Poliza.Tab = 2
        Txt_ApPatBen.SetFocus
        Exit Function
    End If
'    If Txt_ApMatBen = "" Then
'        MsgBox "Debe ingresar Apellido Materno del Beneficiario.", vbExclamation, "Error de Datos"
'        SSTab_Poliza.Tab = 2
'        Txt_ApMatBen.SetFocus
'        Exit Function
'    End If

    'Valida la Fecha de Nacimiento
    'If Not IsDate(Txt_FecNacBen) Then
    If (flValidaFecha(Txt_FecNacBen) = False) Then
        MsgBox "Debe ingresar un dato válido para la Fecha de Nacimiento del Beneficiario.", vbExclamation, "Error de Datos"
        SSTab_Poliza.Tab = 2
        Txt_FecNacBen.SetFocus
        Exit Function
    End If

    'Valida la Fecha de Fallecimiento
    If Trim(Txt_FecFallBen) <> "" Then
        'If Not IsDate(Txt_FecFallBen) Then
        If (flValidaFecha(Txt_FecFallBen) = False) Then
            MsgBox "Debe ingresar un dato válido para la Fecha de Fallecimiento del Beneficiario.", vbExclamation, "Error de Datos"
            SSTab_Poliza.Tab = 2
            Txt_FecFallBen.SetFocus
            Exit Function
        End If
    End If

    'Valida selección de Parentesco
    If Trim(Cmb_Parentesco) = "" Then
        MsgBox "Debe seleccionar el Parentesco del Beneficiario.", vbExclamation, "Error de Datos"
        SSTab_Poliza.Tab = 2
        Cmb_Parentesco.SetFocus
        Exit Function
    End If

    'Valida selección de Grupo Familiar
    If Trim(Cmb_GrupoFam) = "" Then
        MsgBox "Debe seleccionar el Grupo Familiar del Beneficiario.", vbExclamation, "Error de Datos"
        SSTab_Poliza.Tab = 2
        Cmb_GrupoFam.SetFocus
        Exit Function
    End If

    'Valida selección de Sexo
    If Trim(Cmb_SexoBen) = "" Then
        MsgBox "Debe seleccionar el Sexo del Beneficiario.", vbExclamation, "Error de Datos"
        SSTab_Poliza.Tab = 2
        Cmb_SexoBen.SetFocus
        Exit Function
    End If

    'Valida selección de Situación de Invalidez
    If Trim(Cmb_SitInv) = "" Then
        MsgBox "Debe seleccionar la Situación de Invalidez del Beneficiario.", vbExclamation, "Error de Datos"
        SSTab_Poliza.Tab = 2
        Cmb_SitInv.SetFocus
        Exit Function
    End If

    'Valida la Selección de la Causa de Invalidez del Beneficiario
    If Lbl_CauInvBen = "" Then
        MsgBox "Debe ingresar la Causa de Invalidez del Beneficiario.", vbExclamation, "Error de Datos"
        SSTab_Poliza.Tab = 2
        Cmd_BuscaCauInv.SetFocus
        Exit Function
    End If

    'Valida el Ingreso de la Fecha de Invalidez si se ha seleccionado una Causa
    If Trim(Mid(Lbl_CauInvBen, 1, InStr(1, Lbl_CauInvBen, "-") - 1)) <> clCodCauInvNoInv Then
        If Not IsDate(Txt_FecInvBen) Then
            MsgBox "Debe ingresar la Fecha de Invalidez del Beneficiario.", vbExclamation, "Error de Datos"
            SSTab_Poliza.Tab = 2
            Txt_FecInvBen.SetFocus
            Exit Function
        End If
    Else
        Txt_FecInvBen = ""
    End If

'**********************************************
    'Validación Lógica de los Datos, referente a las Fechas Nac., Inv., Fall.
    'Pendiente

'    'Valida Derecho a Pensión del Beneficiario
'    vlNumero = InStr(Cmb_BMPar.Text, "-")
'    vlCodPar = Trim(Mid(Cmb_BMPar.Text, 1, vlNumero - 1))
'    vlNumero = InStr(Cmb_TipoPension.Text, "-")
'    vlCodTipPension = Trim(Mid(Cmb_TipoPension.Text, 1, vlNumero - 1))
'    vgX = InStr(1, clPensionInvVejez, Trim(vlCodTipPension))
'    If (vgX = 0) Then
'        'Cuando el Causante se encuentra Muerto
'        'I--- ABV 13/06/2005 -- If (vlCodPar <> "99") Then
'        If (vlCodPar = "99") Then
'            vlNumero = InStr(Cmb_BMDerPen.Text, "-")
'            vlCodEstPension = Trim(Mid(Cmb_BMDerPen.Text, 1, vlNumero - 1))
'            If (vlCodEstPension <> "10") Then
'                'I--- ABV 13/06/2005 -- MsgBox "El Beneficiario ingresado debe estar Sin Derecho a Pensión.", vbCritical, "Error de Datos"
'                MsgBox "El Causante ingresado debe estar Sin Derecho a Pensión.", vbCritical, "Error de Datos"
'                Cmb_BMDerPen.SetFocus
'                vlSwCalIntOK = False
'                Exit Function
'            End If
'        End If
'    Else
'        If (vlCodPar = "99") Then
'        'Cuando el Causante se encuentra Vivo
'            vlNumero = InStr(Cmb_BMDerPen.Text, "-")
'            vlCodEstPension = Trim(Mid(Cmb_BMDerPen.Text, 1, vlNumero - 1))
'            'I--- ABV 13/06/2005 -- If (vlCodEstPension <> "10") Then
'            If (vlCodEstPension = "10") Then
'                'I--- ABV 13/06/2005 -- MsgBox "El Causante de la Póliza debe estar Sin Derecho a Pensión.", vbCritical, "Error de Datos"
'                MsgBox "El Causante de la Póliza debe estar Con Derecho a Pensión.", vbCritical, "Error de Datos"
'                Cmb_BMDerPen.SetFocus
'                vlSwCalIntOK = False
'                Exit Function
'            End If
'        End If
'    End If

    vlNumero = InStr(Cmb_SitInv.Text, "-")
    vlValCodSitInv = Trim(Mid(Cmb_SitInv.Text, 1, vlNumero - 1))
    If vlValCodSitInv <> Trim(clCodSitInvNoInv) Then
        If Lbl_CauInv = "" Then
            MsgBox "Debe Ingresar Causal de Invalidez para el Beneficiario o Modificar Situación.", vbCritical, "Error de Datos"
            Cmd_BuscaCauInv.SetFocus
            SSTab_Poliza.Tab = 2
            Exit Function
        End If
    End If

'**********************************************

    flValidaDatosAseg = True

Exit Function
Err_ValDatAseg:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flValidaRutGrilla(iTipoIden As String, iNumIden As String) As Boolean
Dim vlTipoIdenAfi As String
Dim vlNumIdenAfi As String

    flValidaRutGrilla = False

    vlPos = Msf_GriAseg.rows - 1
    Msf_GriAseg.Row = 1
    Msf_GriAseg.Col = 11
    For vlI = 1 To vlPos
        Msf_GriAseg.Row = vlI
        Msf_GriAseg.Col = 0
        If vlNumOrden <> Msf_GriAseg.Text Then
            'Obtener la Identificación del Beneficiario desde la Grilla
            Msf_GriAseg.Col = 11
            If Msf_GriAseg.Text <> "" Then
                vlTipoIdenAfi = Trim(Mid(Msf_GriAseg.Text, 1, InStr(1, Msf_GriAseg.Text, "-") - 1))
            Else
                vlTipoIdenAfi = ""
            End If
            Msf_GriAseg.Col = 12
            If Msf_GriAseg.Text <> "" Then
                vlNumIdenAfi = Trim(Msf_GriAseg.Text)
            Else
                vlNumIdenAfi = ""
            End If

            If (iTipoIden = vlTipoIdenAfi) And (iNumIden = vlNumIdenAfi) Then
                MsgBox "El Tipo y Número de Identificación ingresado ya existe en la lista de Beneficiarios", vbExclamation, "Proceso Cancelado"
                flValidaRutGrilla = False
                If vlNumOrden <> clNumOrden1 Then
                    If (Cmb_TipoIdentBen.ListCount <> 0) Then
                        Cmb_TipoIdentBen.ListIndex = 0
                    End If

                    Txt_NumIdentBen = ""
                    Cmb_TipoIdentBen.Enabled = True
                    Txt_NumIdentBen.Enabled = True
                    Cmb_TipoIdentBen.SetFocus
                Else
                    If (Cmb_TipoIdent.ListCount <> 0) Then
                        Cmb_TipoIdent.ListIndex = 0
                    End If
                    Txt_NumIdent = ""
                    Cmb_TipoIdent.SetFocus
                End If
                vlSw = False
                Exit Function
            End If
        End If
    Next vlI

    flValidaRutGrilla = True

End Function

'-----------------------------------------------------------------------
'PERMITE INGRESAR LOS DATOS DEL AFILIADO EN LA PRIMERA FILA DE LA GRILLA
'-----------------------------------------------------------------------
Function flDatosCompletos()
On Error GoTo Err_DatCom

    vgI = InStr(1, Cmb_TipoIdent, "-")

'    vlTipoIden = Trim(Mid(Cmb_TipoIdent, 1, vgI - 1))
    vlTipoIden = (Cmb_TipoIdent)
    vlNumIden = UCase(Trim(Txt_NumIdent))
    vlNomBen = UCase(Trim(Txt_NomAfi))
    vlNomBenSeg = UCase(Trim(Txt_NomAfiSeg))
    vlPatBen = UCase(Trim(Txt_ApPatAfi))
    vlMatBen = UCase(Trim(Txt_ApMatAfi))
    vlFecInv = Trim(Txt_FecInv)
    'If IsDate(Txt_FecInv) Then vlFecInv = Format(Txt_FecInv, "yyyymmdd")
    vlCauInv = Trim(Mid(Lbl_CauInv, 1, InStr(1, Lbl_CauInv, "-") - 1))

    If vlTipoIden = "" Then
        Exit Function
    End If

    vlNumOrden = clNumOrden1

    'Valida que el rut no exista en la grilla
    If flValidaRutGrilla(vlTipoIden, vlNumIden) = False Then
        Exit Function
    End If

    Msf_GriAseg.Row = 1
    Msf_GriAseg.Col = 5
    Msf_GriAseg.Text = vlFecInv
    Msf_GriAseg.Col = 6
    Msf_GriAseg.Text = vlCauInv
    Msf_GriAseg.Col = 11
    Msf_GriAseg.Text = (vlTipoIden)
    Msf_GriAseg.Col = 12
    Msf_GriAseg.Text = Trim(vlNumIden)
    Msf_GriAseg.Col = 13
    Msf_GriAseg.Text = vlNomBen
    Msf_GriAseg.Col = 14
    Msf_GriAseg.Text = vlNomBenSeg
    Msf_GriAseg.Col = 15
    Msf_GriAseg.Text = vlPatBen
    Msf_GriAseg.Col = 16
    Msf_GriAseg.Text = vlMatBen

    If vlNumOrden = clNumOrden1 Then
        Msf_GriAseg.Col = 6
        Msf_GriAseg.Text = Trim(Mid(Lbl_CauInv, 1, (InStr(1, Lbl_CauInv, "-") - 1)))
    End If

'    SSTab_Poliza.TabEnabled(3) = True
'    Msf_GriAseg.Enabled = True
'    Fra_DatosBenef.Enabled = True

Exit Function
Err_DatCom:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flGuardarDatosBenef_Var()

'    vlFecInv = Trim(Txt_FecInvBen)
''    If IsDate(Txt_FecInvBen) Then vlFecInv = Format(Txt_FecInvBen, "yyyymmdd")
'    vlCauInv = Trim(Mid(Lbl_CauInvBen, 1, InStr(1, Lbl_CauInvBen, "-") - 1))

    vlNumOrden = Lbl_NumOrden
    vgPalabra = fgObtenerCodigo_TextoCompuesto(Cmb_TipoIdentBen)
    vlRutBen = Trim(Cmb_TipoIdentBen)
    vlDgvBen = UCase(Trim(Txt_NumIdentBen))

    'Valida que el rut no exista en la grilla
    If flValidaRutGrilla(vgPalabra, vlDgvBen) = False Then
        Exit Function
    End If

    vlNomBen = UCase(Trim(Txt_NombresBen))
    vlNomBenSeg = UCase(Trim(Txt_NombresBenSeg))
    vlPatBen = UCase(Trim(Txt_ApPatBen))
    vlMatBen = UCase(Trim(Txt_ApMatBen))
    vlFecNacBen = CDate(Trim(Txt_FecNacBen))
    If Txt_FecFallBen <> "" Then
        'vlFecFallBen = Format(CDate(Trim(Txt_BMFecFall)), "yyyymmdd")
        vlFecFallBen = CDate(Trim(Txt_FecFallBen))
    Else
        vlFecFallBen = ""
    End If

    vlCodPar = fgObtenerCodigo_TextoCompuesto(Cmb_Parentesco)
    vlCodGruFam = fgObtenerCodigo_TextoCompuesto(Cmb_GrupoFam)
    vlCodSexoBen = fgObtenerCodigo_TextoCompuesto(Cmb_SexoBen)
    vlCodSitInv = fgObtenerCodigo_TextoCompuesto(Cmb_SitInv)
    If Trim(Txt_FecInvBen) <> "" Then
        vlFecInv = CDate(Trim(Txt_FecInvBen))
    Else
        vlFecInv = ""
    End If
    If Lbl_CauInvBen <> "" Then
        vlCauInv = Trim(Mid(Lbl_CauInvBen, 1, (InStr(1, Lbl_CauInvBen, "-") - 1)))
    Else
        vlCauInv = ""
    End If

'I--- ABV 10/08/2007 ---
'    'Corregir el Derecho a Crecer de la Cónyuge o Madre con Hijos => 11 ó 21
'    If vlCodPar = "11" Or vlCodPar = "21" Then
'        vlCodDerCre = vlCodDerCrePol
'    End If
'    'Corregir Porcentaje para la Cobertura de la Cónyuge (S/Hijos) definida en la Cotizazión
'    If (vlCodCoberCon <> "0") And (vlCodCoberCon <> "") Then
'        If vlCodPar = "10" Then vlPrcPension = vlPrcFacPenElla
'    End If
    'Corregir el Estado de Pago para la Pensión
'    vlEstPen = ""
    vlEstPen = fgCalcularEstadoPagoPension(vlFecDev, vlCodTipPen, vlCodPar, Format(vlFecNacBen, "yyyymmdd"), Format(vlFecFallBen, "yyyymmdd"), "", vlCodSitInv)
'F--- ABV 10/08/2007 ---

    vlCodDerPen = ""
    vlCodDerCre = ""
    vlFecNacHM = ""
    vlPrcPension = 0
    vlPenBen = 0
    vlPenGarBen = 0
    vlNumOrdenCot = ""
    vlPrcPensionGar = 0
    vlPrcPensionLeg = 0
End Function

Function flIngresarBenef()

    Msf_GriAseg.AddItem (vlNumOrden) & vbTab _
    & (vlCodPar) & vbTab & (vlCodGruFam) & vbTab _
    & (vlCodSexoBen) & vbTab & (vlCodSitInv) & vbTab _
    & (vlFecInv) & vbTab & (vlCauInv) & vbTab _
    & (vlCodDerPen) & vbTab & (vlCodDerCre) & vbTab _
    & (vlFecNacBen) & vbTab & (vlFecNacHM) & vbTab _
    & " " & (vlRutBen) & vbTab & (vlDgvBen) & vbTab _
    & (vlNomBen) & vbTab & (vlNomBenSeg) & vbTab _
    & (vlPatBen) & vbTab & (vlMatBen) & vbTab _
    & (vlPrcPension) & vbTab _
    & (vlPenBen) & vbTab & (vlPenGarBen) & vbTab _
    & (vlFecFallBen) & vbTab _
    & (vlNumOrden) & vbTab _
    & (vlEstPen) & vbTab & vlPrcPensionGar & vbTab & vlPrcPensionLeg

End Function

Function flModificarBenef()

    Msf_GriAseg.Row = vlNumOrden

    Msf_GriAseg.Col = 1
    Msf_GriAseg.Text = vlCodPar
    Msf_GriAseg.Col = 2
    Msf_GriAseg.Text = vlCodGruFam
    Msf_GriAseg.Col = 3
    Msf_GriAseg.Text = vlCodSexoBen
    Msf_GriAseg.Col = 4
    Msf_GriAseg.Text = vlCodSitInv
    Msf_GriAseg.Col = 5
    Msf_GriAseg.Text = vlFecInv
    Msf_GriAseg.Col = 6
    Msf_GriAseg.Text = vlCauInv
    Msf_GriAseg.Col = 7
    Msf_GriAseg.Text = vlCodDerPen
    Msf_GriAseg.Col = 8
    Msf_GriAseg.Text = vlCodDerCre
    Msf_GriAseg.Col = 9
    Msf_GriAseg.Text = vlFecNacBen
    Msf_GriAseg.Col = 10
    Msf_GriAseg.Text = vlFecNacHM
    Msf_GriAseg.Col = 11
    Msf_GriAseg.Text = " " & vlRutBen
    Msf_GriAseg.Col = 12
    Msf_GriAseg.Text = vlDgvBen
    Msf_GriAseg.Col = 13
    Msf_GriAseg.Text = vlNomBen
    Msf_GriAseg.Col = 14
    Msf_GriAseg.Text = vlNomBenSeg
    Msf_GriAseg.Col = 15
    Msf_GriAseg.Text = vlPatBen
    Msf_GriAseg.Col = 16
    Msf_GriAseg.Text = vlMatBen
    Msf_GriAseg.Col = 17
    Msf_GriAseg.Text = vlPrcPension
    Msf_GriAseg.Col = 18
    Msf_GriAseg.Text = vlPenBen
    Msf_GriAseg.Col = 19
    Msf_GriAseg.Text = vlPenGarBen
    Msf_GriAseg.Col = 20
    Msf_GriAseg.Text = vlFecFallBen
    Msf_GriAseg.Col = 21
    Msf_GriAseg.Text = vlNumOrdenCot
    Msf_GriAseg.Col = 22
    Msf_GriAseg.Text = vlEstPen
    Msf_GriAseg.Col = 23
    Msf_GriAseg.Text = vlPrcPensionGar
    Msf_GriAseg.Col = 24
    Msf_GriAseg.Text = vlPrcPensionLeg

End Function

Function flCalcularCotizacion() As Boolean
Dim vlValor As Long
On Error GoTo Err_Calcular

    flCalcularCotizacion = False

    Screen.MousePointer = 11
'0. Cargar Edades para la Tabla de Mortalidad - Hijos
    L24 = 0
    L21 = 0
    L18 = 0
    
    'Edad de 24 años
    If (fgCarga_Param("LI", "L24", vlFecCalculo) = True) Then
        L24 = vgValorParametro
    Else
        vgError = 1000
        MsgBox "No existe Edad de tope para los 24 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    L24 = L24 * 12 'Mensualizar la Edad de 24 Años
    
    'Edad de 21 años
    If (fgCarga_Param("LI", "L21", vlFecCalculo) = True) Then
        L21 = vgValorParametro
    Else
        vgError = 1000
        MsgBox "No existe Edad de tope para los 21 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    L21 = L21 * 12 'Mensualizar la Edad de 21 Años

    'Edad de 18 años
    If (fgCarga_Param("LI", "L18", vlFecCalculo) = True) Then
        L18 = vgValorParametro
    Else
        vgError = 1000
        MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    L18 = L18 * 12 'Mensualizar la Edad de 18 Años


'1. Cargar las Tablas de Mortalidad MENSUALES en la Estructura Existente
    Call fgCargarTablaMortalidad(vgTipoPeriodoMensual)

'2. Buscar el Nombre del Padre de las Tablas de Mortalidad
    'Mujeres - RV
    vgPalabra_MortalVit_F = fgComboMortalNombre(vlFecCalculo, vgTipoTablaRentista, vgTipoPeriodoMensual, "F")
    'Mujeres - IT
    vgPalabra_MortalTot_F = fgComboMortalNombre(vlFecCalculo, vgTipoTablaTotal, vgTipoPeriodoMensual, "F")
    'Mujeres - IP
    vgPalabra_MortalPar_F = fgComboMortalNombre(vlFecCalculo, vgTipoTablaParcial, vgTipoPeriodoMensual, "F")
    'Mujeres - BEN
    vgPalabra_MortalBen_F = fgComboMortalNombre(vlFecCalculo, vgTipoTablaBeneficiario, vgTipoPeriodoMensual, "F")

    'Hombres - RV
    vgPalabra_MortalVit_M = fgComboMortalNombre(vlFecCalculo, vgTipoTablaRentista, vgTipoPeriodoMensual, "M")
    'Hombres - IT
    vgPalabra_MortalTot_M = fgComboMortalNombre(vlFecCalculo, vgTipoTablaTotal, vgTipoPeriodoMensual, "M")
    'Hombres - IP
    vgPalabra_MortalPar_M = fgComboMortalNombre(vlFecCalculo, vgTipoTablaParcial, vgTipoPeriodoMensual, "M")
    'Hombres - BEN
    vgPalabra_MortalBen_M = fgComboMortalNombre(vlFecCalculo, vgTipoTablaBeneficiario, vgTipoPeriodoMensual, "M")

'3. Verificar selección de Tablas de Mortalidad
    'Mujeres - Rtas. Vitalicias
    If (vgPalabra_MortalVit_F = "") Then
        MsgBox "No se encuentra registrada en la BD Tabla de Mortalidad de Rtas. Vitalicias de Mujeres para el cálculo.", vbCritical, "Error de Datos"
        Exit Function
    End If
    'Mujeres - Invalidez Total
    If (vgPalabra_MortalTot_F = "") Then
        MsgBox "No se encuentra registrada en la BD Tabla de Mortalidad de Inv. Total de Mujeres para el cálculo.", vbCritical, "Error de Datos"
        Exit Function
    End If
    'Mujeres - Invalidez Parcial
    If (vgPalabra_MortalPar_F = "") Then
        MsgBox "No se encuentra registrada en la BD Tabla de Mortalidad de Inv. Parcial de Mujeres para el cálculo.", vbCritical, "Error de Datos"
        Exit Function
    End If
    'Mujeres - Beneficiarios
    If (vgPalabra_MortalBen_F = "") Then
        MsgBox "No se encuentra registrada en la BD Tabla de Mortalidad de Beneficiarios de Mujeres para el cálculo.", vbCritical, "Error de Datos"
        Exit Function
    End If

    'Hombres - Rtas. Vitalicias
    If (vgPalabra_MortalVit_M = "") Then
        MsgBox "No se encuentra registrada en la BD Tabla de Mortalidad de Rtas. Vitalicias de Hombres para el cálculo.", vbCritical, "Error de Datos"
        Exit Function
    End If
    'Hombres - Invalidez Total
    If (vgPalabra_MortalTot_M = "") Then
        MsgBox "No se encuentra registrada en la BD Tabla de Mortalidad de Inv. Total de Hombres para el cálculo.", vbCritical, "Error de Datos"
        Exit Function
    End If
    'Hombres - Invalidez Parcial
    If (vgPalabra_MortalPar_M = "") Then
        MsgBox "No se encuentra registrada en la BD Tabla de Mortalidad de Inv. Parcial de Hombres para el cálculo.", vbCritical, "Error de Datos"
        Exit Function
    End If
    'Hombres - Beneficiarios
    If (vgPalabra_MortalBen_M = "") Then
        MsgBox "No se encuentra registrada en la BD Tabla de Mortalidad de Beneficiarios de Hombres para el cálculo.", vbCritical, "Error de Datos"
        Exit Function
    End If

'4. Buscar el Número Correlativo de las Tablas de Mortalidad
    'Mujeres - RV
    vgMortalVit_F = fgBuscarMortalCodigo(vgPalabra_MortalVit_F)
    'Mujeres - IT
    vgMortalTot_F = fgBuscarMortalCodigo(vgPalabra_MortalTot_F)
    'Mujeres - IP
    vgMortalPar_F = fgBuscarMortalCodigo(vgPalabra_MortalPar_F)
    'Mujeres - BEN
    vgMortalBen_F = fgBuscarMortalCodigo(vgPalabra_MortalBen_F)

    'Hombres - RV
    vgMortalVit_M = fgBuscarMortalCodigo(vgPalabra_MortalVit_M)
    'Hombres - IT
    vgMortalTot_M = fgBuscarMortalCodigo(vgPalabra_MortalTot_M)
    'Hombres - IP
    vgMortalPar_M = fgBuscarMortalCodigo(vgPalabra_MortalPar_M)
    'Hombres - BEN
    vgMortalBen_M = fgBuscarMortalCodigo(vgPalabra_MortalBen_M)

'5. Determinar los Finales de Tablas de Mortalidad para cada Tipo
    'Validar Tablas de Mujeres
    vgFinTabVit_F = fgFinTab_Mortal(vgMortalVit_F)
    If (vgFinTabVit_F = -1) Then
        MsgBox "No existe la Edad Final de la Tabla de Mortalidad de Rtas. Vitalicias de Mujeres.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    vgFinTabTot_F = fgFinTab_Mortal(vgMortalTot_F)
    If (vgFinTabTot_F = -1) Then
        MsgBox "No existe la Edad Final de la Tabla de Mortalidad de Inv. Total de Mujeres.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    vgFinTabPar_F = fgFinTab_Mortal(vgMortalPar_F)
    If (vgFinTabPar_F = -1) Then
        MsgBox "No existe la Edad Final de la Tabla de Mortalidad de Inv. Parcial de Mujeres.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    vgFinTabBen_F = fgFinTab_Mortal(vgMortalBen_F)
    If (vgFinTabBen_F = -1) Then
        MsgBox "No existe la Edad Final de la Tabla de Mortalidad de Beneficiarios de Mujeres.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    'Validar Tablas de Hombres
    vgFinTabVit_M = fgFinTab_Mortal(vgMortalVit_M)
    If (vgFinTabVit_M = -1) Then
        MsgBox "No existe la Edad Final de la Tabla de Mortalidad de Rtas. Vitalicias de Hombres.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    vgFinTabTot_M = fgFinTab_Mortal(vgMortalTot_M)
    If (vgFinTabTot_M = -1) Then
        MsgBox "No existe la Edad Final de la Tabla de Mortalidad de Inv. Total de Hombres.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    vgFinTabPar_M = fgFinTab_Mortal(vgMortalPar_M)
    If (vgFinTabPar_M = -1) Then
        MsgBox "No existe la Edad Final de la Tabla de Mortalidad de Inv. Parcial de Hombres.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    vgFinTabBen_M = fgFinTab_Mortal(vgMortalBen_M)
    If (vgFinTabBen_M = -1) Then
        MsgBox "No existe la Edad Final de la Tabla de Mortalidad de Beneficiarios de Hombres.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If

    'Falta validar que el FinTab tome el mayor valor de Tablas de Mortalidad
    'Tomar el Mayor valor para el Término de la Tabla de Mortalidad
    vlValor = fgMaximo(vgFinTabVit_F, vgFinTabTot_F)
    vlValor = fgMaximo(vlValor, vgFinTabPar_F)
    vlValor = fgMaximo(vlValor, vgFinTabBen_F)
    vlValor = fgMaximo(vlValor, vgFinTabVit_M)
    vlValor = fgMaximo(vlValor, vgFinTabTot_M)
    vlValor = fgMaximo(vlValor, vgFinTabPar_M)
    vlValor = fgMaximo(vlValor, vgFinTabBen_M)
    Fintab = vlValor

    ReDim Lx(1 To 2, 1 To 3, 1 To Fintab) As Double
    ReDim Ly(1 To 2, 1 To 3, 1 To Fintab) As Double

    Call fgLimpiarVariablesGlobales

'6. Determinar el Número de Cargas
    vlNumCargas = Msf_GriAseg.rows - 1

'6. Realizar el Proceso de Cálculo
    If (fgCalcularRentaVitalicia(stPolizaMod, stBeneficiariosMod, vlNumCot, vlCodAFP, vgRentabilidadAFP, vlNumCargas) = False) Then
        MsgBox "Se ha producido un error durante el proceso de Cálculo. Vuelva a intertarlo.", vbCritical, "Estado de la Operación"
        Screen.MousePointer = 0
        Exit Function
    End If

'7. Visualizar la Nueva Pensión de Referencia
    Lbl_MtoPension = stPolizaMod.Mto_Pension
    Lbl_MtoPensionGar = stPolizaMod.Mto_PensionGar
    Lbl_MtoPrimaUniSim = stPolizaMod.Mto_PriUniSim
    Lbl_MtoPrimaUniDif = stPolizaMod.Mto_PriUniDif
    Txt_PrcFam = stPolizaMod.Mto_FacPenElla

'8. Volver a Calcular los Montos de Pensión del GF
    vlNumCargas = Msf_GriAseg.rows - 1
    vlCodDerCrePol = stPolizaMod.Cod_DerCre
    vlIndCobPol = stPolizaMod.Ind_Cob
    vlNumMesGar = stPolizaMod.Num_MesGar

    vgError = 0

    Call fgCalcularPorcentajeBenef(vlFecDev, vlNumCargas, stBeneficiariosMod, stPolizaMod.Cod_TipPension, stPolizaMod.Mto_Pension, True, vlCodDerCrePol, vlIndCobPol, False, vlNumMesGar)

    If (vgError = 0) Then
        Call fgActualizaGrillaBeneficiarios(Msf_GriAseg, stBeneficiariosMod, vlNumCargas, vlNumMesGar, vlCodDerCrePol, vlFecDev, vlCodTipPen)

'        MsgBox "El cálculo ha finalizado Correctamente.", vbInformation, "Operación Realizada"
    Else
'        MsgBox "El Cálculo no ha podido llevarse a cabo correctamente.", vbCritical, "Operación Cancelada"
    End If


'    vlSql = "update tmae_propuesta set "
'    vlSql = vlSql & "cod_calculo = 'C' "
'    'I---- ABV 10/02/2004 ---
'    vlSql = vlSql & "where "
'    vlSql = vlSql & "num_cot = '" & Txt_NumCotCal & "'"
'    'F---- ABV 10/02/2004 ---
'    vgConectarBD.Execute (vlSql)

'    vlSql = "Select num_correlcot from tmae_propuesta where "
'    vlSql = vlSql & "cod_calculo = 'C' and "
'    vlSql = vlSql & "cod_estado = 'C' "
'    'I---- ABV 10/02/2004 ---
'    vlSql = vlSql & "and num_cot = '" & Txt_NumCotCal & "'"
'    'F---- ABV 10/02/2004 ---
'    Set vgRs = vgConectarBD.Execute(vlSql)
'    While Not vgRs.EOF
'        vlNumCot = Mid(Txt_NumCotCal, 1, 13) & Format(vgRs!num_correlcot, "00") & Mid(Txt_NumCotCal, 16, 15)
'
'        vlSql = "update tmae_cotizacion set "
'        vlSql = vlSql & "cod_tipcot = 'C' "
'        vlSql = vlSql & "where "
'        vlSql = vlSql & "num_cot = '" & vlNumCot & "'"
'        vgConectarBD.Execute (vlSql)
'
'        vgRs.MoveNext
'    Wend
'    vgRs.Close

    vgCalculo = "S"
    flCalcularCotizacion = True

    Screen.MousePointer = 0

Exit Function
Err_Calcular:
    Screen.MousePointer = 0
    Unload Frm_Progress
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Btn_Agregar_Click()
On Error GoTo Err_CmdSumarClick
    
    vlSwCalIntOK = False
    vlSwCalExtOK = False
    
    'Valida el Ingreso de Datos
    If (flValidaDatosAseg = False) Then Exit Sub

    vlSwCalIntOK = True
    
    If vlSwCalIntOK = True Then

        'Fecha de Devengue
        vlFecDev = Format(Txt_FecDev, "yyyymmdd")

        If Lbl_NumOrden = "" Then
            'Si es un Beneficiario Nuevo
            vlNumero = InStr(Cmb_TipoPension.Text, "-")
            vlCodTipPension = Trim(Mid(Cmb_TipoPension.Text, 1, vlNumero - 1))
        End If

        'Ingresar Nuevo Beneficiario
        If Lbl_NumOrden.Caption = "" Then
            vgRes = MsgBox(" ¿ Está seguro que desea Ingresar este Beneficiario ?", 4 + 32 + 256, "Operación de Ingreso")
            If vgRes <> 6 Then
                Cmd_Salir.SetFocus
                Screen.MousePointer = 0
                Exit Sub
                vlSwCalIntOK = False
            End If

            vlNumero = InStr(Cmb_Parentesco.Text, "-")
            vlOpcion = Trim(Mid(Cmb_Parentesco.Text, 1, vlNumero - 1))
            If vlOpcion = cgCauCodPar Then
               MsgBox "No Puede Ingresar un Beneficiario con Parentesco Causante", vbInformation, "Información"
               vlSwCalIntOK = False
               Exit Sub
            End If

            vlNumOrden = Trim(Msf_GriAseg.rows)
            Lbl_NumOrden = Trim(vlNumOrden)
            Call flGuardarDatosBenef_Var
            Call flIngresarBenef

        Else
            'Modificar Beneficiario
            vgRes = MsgBox("¿ Está seguro que desea Modificar los Datos del Beneficiario?", 4 + 32 + 256, "Operación de Actualización")
            If vgRes <> 6 Then
                Screen.MousePointer = 0
                Exit Sub
                vlSwCalIntOK = False
            End If
            vlNumOrden = Trim(Lbl_NumOrden)
            Call flGuardarDatosBenef_Var
            Call flModificarBenef
        End If
        Txt_Asegurados.Text = (Msf_GriAseg.rows - 1)
    Else
        MsgBox "Existen problemas en el Ingreso de la Información.", vbCritical, "Proceso Cancelado"
        vlSwCalIntOK = False
    End If

    vlSwCalIntOK = False

Exit Sub
Err_CmdSumarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Btn_Porcentaje_Click()
Dim vlTipoPension As String
Dim vlMtoPensionRef As Double
Dim vlCodDerCrecerCot As String
Dim vlIndCobertura As String
Dim vlMesesGar As Long
On Error GoTo Err_Cmd_BMCalcular_Click

    Call flLimpiarDatosAseg

'    'I--- ABV 15/04/2005 ---
'    'Modifique la Fecha a utilizar para determinar los Porcentajes de
'    'Pensión - Igualmente debo preguntar a Daniela si está bien o no
'    '-----------------------------------------------------------
'    'vlFecVig = Format(CDate(Trim(Txt_PMIniVig)), "yyyymmdd")
'    'Validar el ingreso de la Fecha de Efecto del Endoso
'    If (Txt_FecVig = "") Then
'        MsgBox "Debe ingresar la Fecha de Vigencia sobre la cual se realizarán los cálculos.", vbCritical, "Operación Cancelada"
'        Txt_FecVig.SetFocus
'        Exit Sub
'    End If
'    If Not IsDate(Txt_FecVig) Then
'        MsgBox "La Fecha de Vigencia ingresada no es una fecha válida.", vbCritical, "Operación Cancelada"
'        Txt_FecVig.SetFocus
'        Exit Sub
'    End If

'    vgPalabraAux = Format(flCalculaFechaEfecto(Trim("")), "yyyymmdd")
'    If vgPalabraAux = "" Then Exit Sub
'    vgPalabra = Format(Txt_EndFecEfecto, "yyyymmdd")
'    If (vgPalabra < vgPalabraAux) Then
'        MsgBox "La Fecha de Efecto ingresada es menor a la Fecha de Cierre de la Póliza.", vbCritical, "Operación Cancelada"
'        Txt_EndFecEfecto.SetFocus
'        Exit Sub
'    End If

'Validar la Fecha de Devengue
    If (Txt_FecDev = "") Then
        MsgBox "Debe ingresar la Fecha de Devengue sobre la cual se realizarán los cálculos.", vbCritical, "Operación Cancelada"
'        Lbl_FecDev.SetFocus
        Exit Sub
    End If
    If Not IsDate(Txt_FecDev) Then
        MsgBox "La Fecha de Devengue ingresada no es una fecha válida.", vbCritical, "Operación Cancelada"
'        Lbl_FecDev.SetFocus
        Exit Sub
    End If

'Validar el Tipo de Pensión
    If (Cmb_TipoPension = "") Then
        Exit Sub
    End If
    vlTipoPension = fgObtenerCodigo_TextoCompuesto(Cmb_TipoPension)

''Validar la Pensión de Referencia
'    If Not IsNumeric(Lbl_MtoPension) Then
'        Exit Sub
'    End If
'    vlMtoPensionRef = CDbl(Lbl_MtoPension)

'Validar el Derecho a Crecer
    If (Cmb_DerCre = "") Then
        Exit Sub
    End If
    vlCodDerCrecerCot = Mid(Cmb_DerCre, 1, 1)

'Validar la Cobertura
    If (Cmb_IndCob = "") Then
        Exit Sub
    End If
    vlIndCobertura = Mid(Cmb_IndCob, 1, 1)

'Validar Meses Garantizados
    If (Txt_MesesGar = "") Then
        Exit Sub
    End If
    vlMesesGar = CLng(Txt_MesesGar)

'    vlFecVig = Format(Txt_FecVig, "yyyymmdd")
    vlFecDev = Format(Txt_FecDev, "yyyymmdd")

    'Datos de calculo
    Call fgCargaEstBenGrilla(Msf_GriAseg, stBeneficiariosMod, vlFecDev)
    vlNumCargas = (Msf_GriAseg.rows - 1)

    ''I--- ABV 15/04/2005 ---
    ''Modifique la Fecha a utilizar para determinar los Porcentajes de
    ''Pensión - Igualmente debo preguntar a Daniela si está bien o no
    ''-----------------------------------------------------------
    ''vlFecVig = Format(CDate(Trim(Txt_PMIniVig)), "yyyymmdd")
    'If Txt_EndFecEnd = "" Then
    '    vlFecVig = fgBuscaFecServ
    '    vlFecVig = Format(fgValidaFechaEfecto(Trim(vlFecVig), vlNumPoliza, vlNumOrden), "yyyymmdd")
    'Else
    '    vlFecVig = Format(Txt_EndFecEfecto, "yyyymmdd")
    'End If
    ''I--- ABV 15/04/2005 ---

    vgError = 0

    Call fgCalcularPorcentajeBenef(vlFecDev, vlNumCargas, stBeneficiariosMod, vlTipoPension, vlMtoPensionRef, False, vlCodDerCrecerCot, vlIndCobertura, True)

    If (vgError = 0) Then
'        Call flInicializaGrillaBenef(Msf_GriAseg)
        Call fgActualizaGrillaBeneficiarios(Msf_GriAseg, stBeneficiariosMod, vlNumCargas, vlMesesGar, vlCodDerCrecerCot, vlFecDev, vlTipoPension)

'        Lbl_PMNumCar = (Msf_BMGrilla.Rows - 1)

        vlSwCalIntOK = True

        MsgBox "El Cálculo ha finalizado Correctamente.", vbInformation, "Operación Realizada"
    Else
        MsgBox "El Cálculo no ha podido llevarse a cabo correctamente.", vbCritical, "Operación Cancelada"
    End If

'    Cmd_BMSumar.SetFocus

Exit Sub
Err_Cmd_BMCalcular_Click:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Btn_Quita_Click()
On Error GoTo Err_Cmd_BMRestar

    If Lbl_NumOrden.Caption <> "" Then
        Msf_GriAseg.Row = 1
        Msf_GriAseg.Col = 0
        vlPos = Msf_GriAseg.Row
        vlSwEncontrado = False
        While vlPos <= Msf_GriAseg.rows - 1
            If Trim(Lbl_NumOrden.Caption) = Trim(Msf_GriAseg.Text) Then
                'Si el Beneficiario es encontrado
                vlSwEncontrado = True
            End If
            vlPos = vlPos + 1
            If Msf_GriAseg.Row < Msf_GriAseg.rows - 1 Then
                Msf_GriAseg.Row = vlPos
            End If
        Wend
    Else
        MsgBox "Debe Seleccionar el Beneficiario que Desea Eliminar.", vbInformation, "Información"
        Exit Sub
    End If

    If vlSwEncontrado = True Then
        vgRes = MsgBox(" ¿ Está seguro que desea Eliminar este Beneficiario ?", 4 + 32 + 256, "Operación de Eliminación")
        If vgRes <> 6 Then
            'Cmd_Salir.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    Else
        'Si el beneficiario no es encontrado, NO puede ser eliminado
        MsgBox "El Beneficiario Seleccionado No puede ser Eliminado.", vbInformation, "Información"
        Exit Sub
    End If

    vlSwCalIntOK = False
    vlSwCalExtOK = False

    Msf_GriAseg.RemoveItem Msf_GriAseg.Row

    'ReAsignar Números de Orden a Registros Nuevos
    Msf_GriAseg.Row = 1
    Msf_GriAseg.Col = 0
    vlPos = Msf_GriAseg.Row
    While vlPos <= Msf_GriAseg.rows - 1
        If Trim(Msf_GriAseg.Text) <> vlPos Then
            'Si el número de línea es distinto al número de orden
            Msf_GriAseg.Text = vlPos
        End If
        vlPos = vlPos + 1
        If Msf_GriAseg.Row < Msf_GriAseg.rows - 1 Then
            Msf_GriAseg.Row = vlPos
        End If
    Wend

    Call flLimpiarDatosAseg
    Cmb_TipoIdentBen.SetFocus
    Txt_Asegurados.Text = (Msf_GriAseg.rows - 1)

'    Msf_GriAseg.Row = 1
'    Call flCargaDatosBeneficiariosMod(Msf_GriAseg.Row)

Exit Sub
Err_Cmd_BMRestar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmb_Afp_Click()
If (Cmb_Afp <> "") Then
    If (vlSwAfp = True) Then
        vlCodAFP = fgObtenerCodigo_TextoCompuesto(Cmb_Afp)
        vlFecCalculo = Format(Lbl_FecCalculo, "yyyymmdd")
        If (vlCodAFP <> "000") Then
            'Buscar la Rentabilidad de la AFP
            If (fgObtenerRentabilidadAFP(vlCodAFP, vlFecCalculo, vgRentabilidadAFP) = False) Then
                Lbl_RentaAFP = ""
                MsgBox "Inexistencia del Valor de Rentabilidad de la AFP a la Fecha de Cálculo.", vbCritical, "Proceso de Calculo Cancelado"
                Exit Sub
            Else
                Lbl_RentaAFP = Format(vgRentabilidadAFP, "#0.00")
            End If
        End If
    End If
End If
End Sub

Private Sub Cmb_TipoIdentBen_Click()
If (Cmb_TipoIdentBen <> "") Then
    vlPosicionTipoIden = Cmb_TipoIdentBen.ListIndex
    vlLargoTipoIden = Cmb_TipoIdentBen.ItemData(vlPosicionTipoIden)
    If (vlLargoTipoIden = 0) Then
        Txt_NumIdentBen.Text = "0"
        Txt_NumIdentBen.Enabled = False
    Else
        If (Cmb_TipoIdentBen.Enabled = True) Then
            Txt_NumIdentBen = ""
        End If
        Txt_NumIdentBen.MaxLength = vlLargoTipoIden
        Txt_NumIdentBen.Enabled = True
        If (Txt_NumIdentBen <> "") Then Txt_NumIdentBen.Text = Mid(Txt_NumIdentBen, 1, vlLargoTipoIden)
    End If
End If
End Sub

Private Sub Cmd_BuscaCauInv_Click()
On Error GoTo Err_Buscar

    Screen.MousePointer = 11
    Me.Enabled = False
    vgFormulario = "R"  'indica al formulario frm_buscacoti que fue llamado por el boton de Cotizaciones
    vgFormularioCarpeta = "B" 'indica carpeta de Beneficiarios
    Call Center(Frm_BuscaCauInv)
    Frm_BuscaCauInv.Show
    Screen.MousePointer = 0

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscaCor_Click()
On Error GoTo Err_Buscar

    Frm_BuscaCorredor.flInicio ("Frm_CalPolizaRec")

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarDir_Click()
On Error GoTo Err_Buscar

    Frm_BusDireccion.flInicio ("Frm_CalPolizaRec")

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Calcular_Click()
On Error GoTo Err_Calcular

'I--- ABV 17/11/2005 ---
    vgUtilizarNormativa = "N"
'F--- ABV 17/11/2005 ---

    vlSwCalExtOK = False
    
    If (vlSwCalIntOK = False) Then
        MsgBox "Debe realizar el proceso de Cálculo de Porcentajes del Grupo Familiar.", vbCritical, "Proceso Cancelado"
        SSTab_Poliza.Tab = 2
        Exit Sub
    End If

    'Verificar cálculo a realizar
    vgRes = MsgBox("¿ Está seguro que desea realizar el Proceso de Cálculo ?", 4 + 32 + 256, "Confirmación")
    If vgRes <> 6 Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    'Valida que esten todos los datos del afiliado
    If flValDatAfi = False Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    vlTipoIden = fgObtenerCodigo_TextoCompuesto(Cmb_TipoIdent.Text)
    vlNumIden = Txt_NumIdent

    'Valida que esten todos los datos del calculo
    If flValDatCal = False Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    'Valida que los beneficiarios de la grilla tengan todos los datos
    If flValidaBenGrilla = False Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    vlFecCalculo = Format(Lbl_FecCalculo, "yyyymmdd")
    vlFecEmision = Format(Lbl_FecVig, "yyyymmdd")
    vlFecVig = Mid(Format(Lbl_FecVig, "yyyymmdd"), 1, 6) & "01"
    vlFecDev = Format(Txt_FecDev, "yyyymmdd")
    vlCodTipPen = fgObtenerCodigo_TextoCompuesto(Cmb_TipoPension)
    vlCodAFP = fgObtenerCodigo_TextoCompuesto(Cmb_Afp)
    vlNumPol = Trim(Txt_NumPol)
    vlNumCot = Trim(Txt_NumCot)
    vlNumCorrelativo = Trim(Lbl_SecOfe)

    'Buscar el Factor de Ajuste del IPC
    If (fgObtenerFactorAjusteIPC(vlFecDev, vlFecCalculo) = False) Then
        MsgBox "Inexistencia del Valor de Ajuste para la Fecha de Cálculo y/o Devengue.", vbCritical, "Proceso de Calculo Cancelado"
        Exit Sub
    End If

    'Buscar la Rentabilidad de la AFP
    If (fgObtenerRentabilidadAFP(vlCodAFP, vlFecCalculo, vgRentabilidadAFP) = False) Then
        MsgBox "Inexistencia del Valor de Rentabilidad de la AFP a la Fecha de Cálculo.", vbCritical, "Proceso de Calculo Cancelado"
        Exit Sub
    End If

    Screen.MousePointer = 11

    'Guardar los Datos de la Póliza en una Estructura
    Call fgCargaEstPoliza(Frm_CalPolizaRec, stPolizaMod, vlNumPol, vlFecCalculo, vlFecVig, vlFecEmision, vlBotonEscogido, vlNumCot, vlNumCorrelativo)

    'Guardar los Datos de los Beneficiarios en una Estructura
    Call fgCargaEstBenGrilla(Msf_GriAseg, stBeneficiariosMod(), vlFecDev)

    'Calcular Cotización
    If (flCalcularCotizacion = True) Then
        MsgBox "Proceso de Cálculo finalizado Exitosamente", vbInformation, "Estado del Proceso"
        SSTab_Poliza.Tab = 0
        
        vlSwCalExtOK = True
        Fra_Afiliado.Enabled = False
        Fra_Calculo.Enabled = False
        Fra_DatosBenef.Enabled = False
    Else
        MsgBox "Proceso de Cálculo Cancelado por Errores encontrados durante su ejecución", vbCritical, "Estado del Proceso"
    End If

    Screen.MousePointer = 0

Exit Sub
Err_Calcular:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cancelar_Click()
On Error GoTo Err_Cancelar
    
    Screen.MousePointer = 11
    vlSwAfp = False
    Unload Me
    Screen.MousePointer = 0

Exit Sub
Err_Cancelar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_CauInv_Click()
On Error GoTo Err_Buscar

    Screen.MousePointer = 11
    vgFormulario = "R"  'indica al formulario frm_buscacoti que fue llamado por el boton de Cotizaciones
    Frm_BuscaCauInv.Show
    vgFormularioCarpeta = "A" 'indica carpeta de Afiliado
    Call Center(Frm_BuscaCauInv)
    Screen.MousePointer = 0

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Grabar_Click()
On Error GoTo Err_Traspasar
    
    If (vlSwCalIntOK = False) Then
        MsgBox "Debe realizar el proceso de cálculo de los Porcentajes de Beneficiarios.", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    If (vlSwCalExtOK = False) Then
        MsgBox "Debe realizar el proceso de cálculo de la Pensión de Renta Vitalicia.", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    'Llama a las funciones para traspasar datos del formulario
    Call flTraspasarAfiliado
    Call flTraspasarCalculo
    Call flTraspasarBeneficiarios
    
    vgBotonEscogido = "R"
'    vgCodDireccion = vlCodDireccion
    
'    Frm_CalPoliza.Fra_Afiliado.Enabled = False
'    Frm_CalPoliza.Fra_DatCal.Enabled = False
'    Frm_CalPoliza.Fra_DatosBenef.Enabled = False
    
    Screen.MousePointer = 0
    
    Unload Frm_CalPolizaRec

Exit Sub
Err_Traspasar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpia
    
    If (Fra_DatosBenef.Enabled = False) And (Fra_Calculo.Enabled = False) And (Fra_Afiliado.Enabled = False) Then
        MsgBox "No se puede realizar la Limpieza ya que se ha realizado el Proceso de Calculo.", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    If SSTab_Poliza.Tab = 0 Then
        If SSTab_Poliza.TabEnabled(2) = True Then
            If SSTab_Poliza.TabEnabled(2) = False Then
                If (Cmb_TipoIdent.ListCount <> 0) Then
                    Cmb_TipoIdent.ListIndex = 0
                End If
                Txt_NumIdent = ""
            End If

            Txt_NumIdent = ""
            Txt_NomAfi = ""
            Txt_NomAfiSeg = ""
            Txt_ApPatAfi = ""
            Txt_ApMatAfi = ""
            Txt_FecNac = ""
            Txt_FecFall = ""
            Txt_FecInv = ""
            Txt_Dir = ""
            Txt_Fono = ""
            Txt_Correo = ""
            Txt_Nacionalidad = ""
            Txt_NumCta = ""
            Lbl_Departamento = ""
            Lbl_Provincia = ""
            Lbl_Distrito = ""

            If (Cmb_TipoIdent.ListCount <> 0) Then
                Cmb_TipoIdent.ListIndex = 0
            End If
            If (Cmb_Sexo.ListCount <> 0) Then
                Cmb_Sexo.ListIndex = 0
            End If
            If (Cmb_TipoPension.ListCount <> 0) Then
                Cmb_TipoPension.ListIndex = 0
            End If
            If (Cmb_Afp.ListCount <> 0) Then
                Cmb_Afp.ListIndex = 0
            End If
            If (Cmb_EstCivil.ListCount <> 0) Then
                Cmb_EstCivil.ListIndex = 0
            End If
            If (Cmb_Salud.ListCount <> 0) Then
                Cmb_Salud.ListIndex = 0
            End If
            If (Cmb_Vejez.ListCount <> 0) Then
                Cmb_Vejez.ListIndex = 0
            End If
            If (Cmb_ViaPago.ListCount <> 0) Then
                Cmb_ViaPago.ListIndex = 0
            End If
            If (Cmb_Suc.ListCount <> 0) Then
                Cmb_Suc.ListIndex = 0
            End If
            If (Cmb_TipCta.ListCount <> 0) Then
                Cmb_TipCta.ListIndex = 0
            End If
            If (Cmb_Bco.ListCount <> 0) Then
                Cmb_Bco.ListIndex = 0
            End If

        Else
            Cmb_TipoIdent.Enabled = True
            Txt_NumIdent.Enabled = True
            If SSTab_Poliza.TabEnabled(2) = False Then
                If (Cmb_TipoIdent.ListCount <> 0) Then
                    Cmb_TipoIdent.ListIndex = 0
                End If
                Txt_NumIdent = ""
            End If

            Txt_NumIdent = ""
            Txt_NomAfi = ""
            Txt_NomAfiSeg = ""
            Txt_ApPatAfi = ""
            Txt_ApMatAfi = ""
            Txt_FecNac = ""
            Txt_FecFall = ""
            Txt_FecInv = ""
            Txt_Fono = ""
            Txt_Correo = ""
            Txt_Nacionalidad = ""
            Txt_NumCta = ""

            If (Cmb_Sexo.ListCount <> 0) Then
                Cmb_Sexo.ListIndex = 0
            End If
            If (Cmb_TipoPension <> 0) Then
                Cmb_TipoPension.ListIndex = 0
            End If
            If (Cmb_Afp <> 0) Then
                Cmb_Afp.ListIndex = 0
            End If
            If (Cmb_EstCivil.ListCount <> 0) Then
                Cmb_EstCivil.ListIndex = 0
            End If
            If (Cmb_Salud.ListCount <> 0) Then
                Cmb_Salud.ListIndex = 0
            End If
            If (Cmb_Vejez.ListCount <> 0) Then
                Cmb_Vejez.ListIndex = 0
            End If
            If (Cmb_ViaPago.ListCount <> 0) Then
                Cmb_ViaPago.ListIndex = 0
            End If
            If (Cmb_Suc.ListCount <> 0) Then
                Cmb_Suc.ListIndex = 0
            End If
            If (Cmb_TipCta.ListCount <> 0) Then
                Cmb_TipCta.ListIndex = 0
            End If
            If (Cmb_Bco.ListCount <> 0) Then
                Cmb_Bco.ListIndex = 0
            End If

        End If

    End If

    If SSTab_Poliza.Tab = 1 Then
        'flLimpiarDatosCal
        If (Cmb_Moneda.ListCount <> 0) Then
            Cmb_Moneda.ListIndex = 0
        End If
        If (Cmb_TipoRenta.ListCount <> 0) Then
            Cmb_TipoRenta.ListIndex = 0
        End If
        If (Cmb_Modalidad.ListCount <> 0) Then
            Cmb_Modalidad.ListIndex = 0
        End If
        If (Cmb_IndCob.ListCount <> 0) Then
            Cmb_IndCob.ListIndex = 0
        End If
        If (Cmb_CobConyuge.ListCount <> 0) Then
            Cmb_CobConyuge.ListIndex = 0
        End If
        If (Cmb_DerCre.ListCount <> 0) Then
            Cmb_DerCre.ListIndex = 0
        End If
        If (Cmb_DerGra.ListCount <> 0) Then
            Cmb_DerGra.ListIndex = 0
        End If

        Txt_FecDev = ""
        Txt_AnnosDif = ""
        Txt_MesesGar = ""
        Txt_PrcRentaTmp = ""
        Txt_ComInt = ""
        Txt_PrcFam = ""
        Txt_CtaInd = ""
        Txt_BonoAct = ""
        Txt_ApoAdi = ""
    End If

    If SSTab_Poliza.Tab = 2 Then
        If (Cmb_TipoIdentBen.ListCount <> 0) Then
            Cmb_TipoIdentBen.ListIndex = 0
            Txt_NumIdentBen = "0"
        End If
        If (Cmb_Parentesco.ListCount <> 0) Then
            Cmb_Parentesco.ListIndex = 0
        End If
        If (Cmb_GrupoFam.ListCount <> 0) Then
            Cmb_GrupoFam.ListIndex = 0
        End If
        If (Cmb_SexoBen.ListCount <> 0) Then
            Cmb_SexoBen.ListIndex = 0
        End If
        If (Cmb_SitInv.ListCount <> 0) Then
            Cmb_SitInv.ListIndex = 0
        End If

        Lbl_NumOrden = ""
        
        'Txt_NumIdentBen = ""
        Txt_NombresBen = ""
        Txt_NombresBenSeg = ""
        Txt_ApPatBen = ""
        Txt_ApMatBen = ""
        Txt_FecNacBen = ""
        Txt_FecFallBen = ""
        Lbl_PensionBen = ""
        Lbl_PenGar = ""
        
        Txt_FecInvBen = ""
        Lbl_CauInvBen = "0 - " & fgBuscarGlosaCauInv("0")
        Lbl_DerPension = ""
        Lbl_Porcentaje = ""

    End If
    
    vlSwCalIntOK = False
    vlSwCalExtOK = False

Exit Sub
Err_Limpia:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_LimpiarBen_Click()
On Error GoTo Err_LimpiaBen

    Call flLimpiarDatosAseg
    Cmb_TipoIdentBen.SetFocus

Exit Sub
Err_LimpiaBen:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Modificar_Click()
On Error GoTo Err_Modificar
    
    Screen.MousePointer = 11
    
    'Permite la modificación de los cálculos ya realizados
    vlSwCalIntOK = False
    vlSwCalExtOK = False
    
    Fra_Afiliado.Enabled = True
    Fra_Calculo.Enabled = True
    Fra_DatosBenef.Enabled = True
    
    Screen.MousePointer = 0

Exit Sub
Err_Modificar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Err_Carga

    Me.Top = 0
    Me.Left = 0
    SSTab_Poliza.Tab = 0

    Call flIniGrillaBen
'    flCargarGrilla

    'Llena los combos para Editar Afiliados
    Call fgComboTipoIdentificacion(Cmb_TipoIdent)
    Call fgComboGeneral(vgCodTabla_Sexo, Cmb_Sexo)
    Call fgComboGeneral(vgCodTabla_TipPen, Cmb_TipoPension)
    Call fgComboGeneral(vgCodTabla_AFP, Cmb_Afp)
    Call fgComboGeneral(vgCodTabla_EstCiv, Cmb_EstCivil)
    Call fgComboGeneral(vgCodTabla_InsSal, Cmb_Salud)
    Call fgComboGeneral(vgCodTabla_TipVej, Cmb_Vejez)
    Call fgComboGeneral(vgCodTabla_ViaPago, Cmb_ViaPago)
    Call fgComboGeneral(vgCodTabla_TipCta, Cmb_TipCta)
    Call fgComboGeneral(vgCodTabla_Bco, Cmb_Bco)

    'Llena los combos para Editar Cálculo
    Call fgComboMoneda(Cmb_Moneda)
    Call fgComboGeneral(vgCodTabla_TipRen, Cmb_TipoRenta)
    Call fgComboGeneral(vgCodTabla_AltPen, Cmb_Modalidad)
    Call fgComboCoberturaConyuge(Cmb_CobConyuge)
    fgComboSiNo Cmb_IndCob
    fgComboSiNo Cmb_DerCre
    fgComboSiNo Cmb_DerGra

    'Llena los combos Editar Beneficiarios
    Call fgComboTipoIdentificacion(Cmb_TipoIdentBen)
    Call fgComboGeneral(vgCodTabla_Par, Cmb_Parentesco)
    Call fgComboGeneral(vgCodTabla_GruFam, Cmb_GrupoFam)
    Call fgComboGeneral(vgCodTabla_Sexo, Cmb_SexoBen)
    Call fgComboGeneral(vgCodTabla_SitInv, Cmb_SitInv)

    vlBotonEscogido = vgBotonEscogido
    vlSwCalIntOK = False
    vlSwCalExtOK = False
    Fra_Afiliado.Enabled = True
    Fra_Calculo.Enabled = True
    Fra_DatosBenef.Enabled = True

    vlSwAfp = False
    'Cargar los Datos desde el Formulario de Pólizas
    Call flEditarAfiliado
    Call flEditarCalculo
    Call flEditarBeneficiarios

    vlSwAfp = True
    
Exit Sub
Err_Carga:
Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Function flTraspasarAfiliado()

    'Traspasa la información del formulario PolizaRec al
    'formulario PólizaRec
    Frm_CalPoliza.Txt_NumPol = Txt_NumPol
    Frm_CalPoliza.Txt_FecVig = Lbl_FecVig
    
    Frm_CalPoliza.Lbl_NumCot = Txt_NumCot
    Frm_CalPoliza.Lbl_SolOfe = Lbl_SolOfe
    Frm_CalPoliza.Txt_NumIdent = Txt_NumIdent
    Frm_CalPoliza.Txt_Asegurados = Txt_Asegurados
    Frm_CalPoliza.Txt_NomAfi = Txt_NomAfi
    Frm_CalPoliza.Txt_NomAfiSeg = Txt_NomAfiSeg
    Frm_CalPoliza.Txt_ApPatAfi = Txt_ApPatAfi
    Frm_CalPoliza.Txt_ApMatAfi = Txt_ApMatAfi
    Frm_CalPoliza.Lbl_FecNac = Txt_FecNac
    Frm_CalPoliza.Lbl_FecFall = Txt_FecFall
    Frm_CalPoliza.Txt_FecInv = Txt_FecInv
    Frm_CalPoliza.Lbl_CauInv = Lbl_CauInv
    Frm_CalPoliza.Lbl_Dir = Txt_Dir    'RVF 20090914
    Frm_CalPoliza.Lbl_Departamento = Lbl_Departamento
    Frm_CalPoliza.Lbl_Provincia = Lbl_Provincia
    Frm_CalPoliza.Lbl_Distrito = Lbl_Distrito
    Frm_CalPoliza.Txt_Fono = Txt_Fono
    Frm_CalPoliza.Txt_Correo = Txt_Correo
    Frm_CalPoliza.Txt_NumCta = Txt_NumCta
    Frm_CalPoliza.Txt_Nacionalidad = Txt_Nacionalidad
    Frm_CalPoliza.Cmb_Vejez.Text = Cmb_Vejez.Text

    Frm_CalPoliza.Lbl_SexoAfi.Caption = Cmb_Sexo.Text
    Frm_CalPoliza.Lbl_TipPen.Caption = Cmb_TipoPension.Text
    Frm_CalPoliza.Lbl_Afp.Caption = Cmb_Afp.Text

    Frm_CalPoliza.Cmb_TipoIdent.Text = Cmb_TipoIdent.Text
    Frm_CalPoliza.Cmb_EstCivil.Text = Cmb_EstCivil.Text
    Frm_CalPoliza.Cmb_Salud.Text = Cmb_Salud.Text
    Frm_CalPoliza.Cmb_ViaPago.Text = Cmb_ViaPago.Text
    Frm_CalPoliza.Cmb_Suc.Text = Cmb_Suc.Text
    Frm_CalPoliza.Cmb_TipCta.Text = Cmb_TipCta.Text
    Frm_CalPoliza.Cmb_Bco.Text = Cmb_Bco.Text

   End Function

Function flTraspasarCalculo()

    'Traspasa la información del formulario PolizaRec al
    'formulario PólizaRec
    Frm_CalPoliza.Lbl_FecDev = Txt_FecDev
    Frm_CalPoliza.Lbl_FecIncorpora = Txt_FecIncorpora
    Frm_CalPoliza.Txt_FecIniPago = Txt_FecIniPago
    Frm_CalPoliza.Lbl_CUSPP = Txt_Cuspp
    Frm_CalPoliza.Lbl_Moneda(0).Caption = Cmb_Moneda.Text
    Frm_CalPoliza.Lbl_TipoRenta.Caption = Cmb_TipoRenta.Text
    Frm_CalPoliza.Lbl_AnnosDif = Txt_AnnosDif
    Frm_CalPoliza.Lbl_Alter.Caption = Cmb_Modalidad.Text
    Frm_CalPoliza.Lbl_MesesGar = Txt_MesesGar
    Frm_CalPoliza.Lbl_RentaAFP = Lbl_RentaAFP
    Frm_CalPoliza.Lbl_PrcRentaTmp = Txt_PrcRentaTmp
    
    vgPalabra = Trim(Mid(Cmb_IndCob, InStr(1, Cmb_IndCob, "-", 0) + 1, Len(Cmb_IndCob)))
    Frm_CalPoliza.Lbl_IndCob.Caption = vgPalabra
    Frm_CalPoliza.Lbl_FacPenElla.Caption = Cmb_CobConyuge.Text
    vgPalabra = Trim(Mid(Cmb_DerCre, InStr(1, Cmb_DerCre, "-", 0) + 1, Len(Cmb_DerCre)))
    Frm_CalPoliza.Lbl_DerCre.Caption = vgPalabra
    vgPalabra = Trim(Mid(Cmb_DerGra, InStr(1, Cmb_DerGra, "-", 0) + 1, Len(Cmb_DerGra)))
    Frm_CalPoliza.Lbl_DerGra.Caption = vgPalabra
    
    'Datos del Intermediario
    Frm_CalPoliza.Lbl_ComInt = Txt_ComInt
    Frm_CalPoliza.Lbl_TipoIdentCorr = Lbl_TipoIdentCorr
    Frm_CalPoliza.Lbl_NumIdentCorr = Lbl_NumIdentCorr
    Frm_CalPoliza.Lbl_ComIntBen = Lbl_ComIntBen
    Frm_CalPoliza.Lbl_BenSocial = Lbl_BenSocial
    
    'Datos de Cálculo de Tasas
    Frm_CalPoliza.Lbl_TasaCtoEq = Lbl_TasaCtoEq
    Frm_CalPoliza.Lbl_TasaVta = Lbl_TasaVta
    Frm_CalPoliza.Lbl_TasaTIR = Lbl_TasaTIR
    Frm_CalPoliza.Lbl_TasaPerGar = Lbl_TasaPerGar
    Frm_CalPoliza.Lbl_MtoPension = Lbl_MtoPension
    Frm_CalPoliza.Lbl_MtoPensionGar = Lbl_MtoPensionGar
    Frm_CalPoliza.Lbl_MtoPrimaUniSim = Lbl_MtoPrimaUniSim
    Frm_CalPoliza.Lbl_MtoPrimaUniDif = Lbl_MtoPrimaUniDif
    Frm_CalPoliza.Txt_PrcFam = Txt_PrcFam
    Frm_CalPoliza.Lbl_CtaInd = Txt_CtaInd
    Frm_CalPoliza.Lbl_BonoAct = Txt_BonoAct
    Frm_CalPoliza.Lbl_ApoAdi = Txt_ApoAdi
    Frm_CalPoliza.Lbl_PriUnica = Lbl_PriUnica

    Frm_CalPoliza.Lbl_MonedaFon(0).Caption = Lbl_MonedaFon(0).Caption
    Frm_CalPoliza.Lbl_MonedaFon(1).Caption = Lbl_MonedaFon(1).Caption
    Frm_CalPoliza.Lbl_MonedaFon(2).Caption = Lbl_MonedaFon(2).Caption

End Function

Function flTraspasarBeneficiarios()
Dim vlCasos As Long

    'Inicializa la Grilla de Beneficiarios
    Frm_CalPoliza.flIniGrillaBen
    
    'Traspasa los datos de la Grilla de Beneficiarios a la Póliza Original
    'Carpeta de Beneficiarios
    vlCasos = Msf_GriAseg.rows
    vgI = 1

    While vgI <= vlCasos - 1

        With Msf_GriAseg

            vlNumOrden = Trim(.TextMatrix(vgI, 0))
            vlCodPar = Trim(.TextMatrix(vgI, 1))
            vlCodGruFam = Trim(.TextMatrix(vgI, 2))
            vlCodSexoBen = Trim(.TextMatrix(vgI, 3))
            vlCodSitInv = Trim(.TextMatrix(vgI, 4))
            vlFecInv = Trim(.TextMatrix(vgI, 5))
            vlCauInv = Trim(.TextMatrix(vgI, 6))
            vlCodDerPen = Trim(.TextMatrix(vgI, 7))
            vlCodDerCre = Trim(.TextMatrix(vgI, 8))
            vlFecNacBen = Trim(.TextMatrix(vgI, 9))
            vlFecNacHM = Trim(.TextMatrix(vgI, 10))
            vlRutBen = " " & Trim(.TextMatrix(vgI, 11))
            vlDgvBen = Trim(.TextMatrix(vgI, 12))
            vlNomBen = Trim(.TextMatrix(vgI, 13))
            vlNomBenSeg = Trim(.TextMatrix(vgI, 14))
            vlPatBen = Trim(.TextMatrix(vgI, 15))
            vlMatBen = Trim(.TextMatrix(vgI, 16))
            vlPrcPension = Trim(.TextMatrix(vgI, 17))
            vlPenBen = Trim(.TextMatrix(vgI, 18))
            vlPenGarBen = Trim(.TextMatrix(vgI, 19))
            vlFecFallBen = Trim(.TextMatrix(vgI, 20))
            vlNumOrdenCot = Trim(.TextMatrix(vgI, 21))
            vlEstPen = Trim(.TextMatrix(vgI, 22))
            vlPrcPensionGar = Trim(.TextMatrix(vgI, 23))
            vlPrcPensionLeg = Trim(.TextMatrix(vgI, 24))

            Frm_CalPoliza.Msf_GriAseg.AddItem (vlNumOrden) & vbTab & _
                        (vlCodPar) & vbTab & (vlCodGruFam) & vbTab & _
                        (vlCodSexoBen) & vbTab & (vlCodSitInv) & vbTab & _
                        (vlFecInv) & vbTab & (vlCauInv) & vbTab & _
                        (vlCodDerPen) & vbTab & (vlCodDerCre) & vbTab & _
                        (vlFecNacBen) & vbTab & (vlFecNacHM) & vbTab & _
                        (vlRutBen) & vbTab & (vlDgvBen) & vbTab & _
                        Trim(vlNomBen) & vbTab & Trim(vlNomBenSeg) & vbTab & _
                        Trim(vlPatBen) & vbTab & Trim(vlMatBen) & vbTab & _
                        Format(CDbl(vlPrcPension), "#,#0.000") & vbTab & _
                        Format(CDbl(vlPenBen), "#,#0.00") & vbTab & _
                        Format(CDbl(vlPenGarBen), "#,#0.00") & vbTab & _
                        Trim(vlFecFallBen) & vbTab & _
                        Trim(vlNumOrdenCot) & vbTab & vlEstPen _
                        & vbTab & vlPrcPensionGar & vbTab & vlPrcPensionLeg
        End With
        vgI = vgI + 1
    Wend

    Frm_CalPoliza.Txt_NumIdentBen = ""
    Frm_CalPoliza.Txt_NombresBen = ""
    Frm_CalPoliza.Txt_NombresBenSeg = ""
    Frm_CalPoliza.Txt_ApPatBen = ""
    Frm_CalPoliza.Txt_ApMatBen = ""
    Frm_CalPoliza.Lbl_FecNacBen = ""
    Frm_CalPoliza.Lbl_FecFallBen = ""
    Frm_CalPoliza.Lbl_PensionBen = ""
    Frm_CalPoliza.Lbl_PenGar = ""
    Frm_CalPoliza.Txt_FecInvBen = ""
    Frm_CalPoliza.Lbl_CauInvBen = ""
    Frm_CalPoliza.Lbl_DerPension = ""
    Frm_CalPoliza.Lbl_Porcentaje = ""

    'Frm_CalPoliza.Cmb_TipoIdentBen.Text = ""
    Frm_CalPoliza.Lbl_Par.Caption = ""
    Frm_CalPoliza.Lbl_Grupo.Caption = ""
    Frm_CalPoliza.Lbl_SexoBen.Caption = ""
    Frm_CalPoliza.Lbl_SitInvBen.Caption = ""

    Frm_CalPoliza.Lbl_Moneda(1) = Lbl_Moneda(1)
    Frm_CalPoliza.Lbl_Moneda(2) = Lbl_Moneda(2)
    Frm_CalPoliza.Lbl_Moneda(3) = Lbl_Moneda(3)
    Frm_CalPoliza.Lbl_Moneda(4) = Lbl_Moneda(4)
    
End Function

Function flLimpiarDatosAseg()
On Error GoTo Err_LimpiaAseg

    If (Cmb_TipoIdentBen.ListCount <> 0) Then
        Cmb_TipoIdentBen.ListIndex = 0
        Txt_NumIdentBen = "0"
    End If
    If (Cmb_Parentesco.ListCount <> 0) Then
        Cmb_Parentesco.ListIndex = 0
    End If
    If (Cmb_GrupoFam.ListCount <> 0) Then
        Cmb_GrupoFam.ListIndex = 0
    End If
    If (Cmb_SexoBen.ListCount <> 0) Then
        Cmb_SexoBen.ListIndex = 0
    End If
    If (Cmb_SitInv.ListCount <> 0) Then
        Cmb_SitInv.ListIndex = 0
    End If

    Txt_NombresBen = ""
    Txt_NombresBenSeg = ""
    Txt_ApPatBen = ""
    Txt_ApMatBen = ""
    Txt_FecNacBen = ""
    Txt_FecFallBen = ""
    Lbl_PensionBen = ""
    Lbl_PenGar = ""
    Txt_FecInvBen = ""
    Lbl_CauInvBen = "0 - " & fgBuscarGlosaCauInv("0")
    Lbl_DerPension = ""
    Lbl_Porcentaje = ""
    Lbl_NumOrden = ""

'    Lbl_Moneda(clMonedaBenPen) = ""
'    Lbl_Moneda(clMonedaBenPenGar) = ""

Exit Function
Err_LimpiaAseg:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function


Function flIniGrillaBen()
On Error GoTo Err_IniGri

    Msf_GriAseg.Clear
    Msf_GriAseg.Cols = 25
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
    Msf_GriAseg.ColWidth(3) = 600
    Msf_GriAseg.Text = "Sexo"

    Msf_GriAseg.Col = 4
    Msf_GriAseg.ColWidth(4) = 600
    Msf_GriAseg.Text = "Sit.Inv."

    Msf_GriAseg.Col = 5
    Msf_GriAseg.ColWidth(5) = 0
    Msf_GriAseg.Text = "Fec.Inv."

    Msf_GriAseg.Col = 6
    Msf_GriAseg.ColWidth(6) = 0
    Msf_GriAseg.Text = "Cau.Inv."

    Msf_GriAseg.Col = 7
    Msf_GriAseg.ColWidth(7) = 0
    Msf_GriAseg.Text = "Der.Pension"

    Msf_GriAseg.Col = 8
    Msf_GriAseg.ColWidth(8) = 0
    Msf_GriAseg.Text = "Der.Crecer"

    Msf_GriAseg.Col = 9
    Msf_GriAseg.ColWidth(9) = 0
    Msf_GriAseg.Text = "Fec.Nac."

    Msf_GriAseg.Col = 10
    Msf_GriAseg.ColWidth(10) = 0
    Msf_GriAseg.Text = "Fec.NacHM"

    Msf_GriAseg.Col = 11
    Msf_GriAseg.ColWidth(11) = 1100
    Msf_GriAseg.Text = "Tipo Ident."

    Msf_GriAseg.Col = 12
    Msf_GriAseg.ColWidth(12) = 1100
    Msf_GriAseg.Text = "Nº Ident."

    Msf_GriAseg.Col = 13
    Msf_GriAseg.ColWidth(13) = 2000
    Msf_GriAseg.Text = " 1er. Nombre"

    Msf_GriAseg.Col = 14
    Msf_GriAseg.ColWidth(14) = 2000
    Msf_GriAseg.Text = " 2do. Nombre"

    Msf_GriAseg.Col = 15
    Msf_GriAseg.ColWidth(15) = 2000
    Msf_GriAseg.Text = "Ap. Paterno"

    Msf_GriAseg.Col = 16
    Msf_GriAseg.ColWidth(16) = 2000
    Msf_GriAseg.Text = "Ap. Materno"

    Msf_GriAseg.Col = 17
    Msf_GriAseg.ColWidth(17) = 0
    Msf_GriAseg.Text = "Porcentaje"

    Msf_GriAseg.Col = 18
    Msf_GriAseg.ColWidth(18) = 0
    Msf_GriAseg.Text = "mto.pension"

    Msf_GriAseg.Col = 19
    Msf_GriAseg.ColWidth(19) = 0
    Msf_GriAseg.Text = "mto. pensiongar"

    Msf_GriAseg.Col = 20
    Msf_GriAseg.ColWidth(20) = 0
    Msf_GriAseg.Text = "Fec.fallecimiento"

    Msf_GriAseg.Col = 21
    Msf_GriAseg.ColWidth(21) = 0
    Msf_GriAseg.Text = "NumOrden Cot"

    Msf_GriAseg.Col = 22
    Msf_GriAseg.ColWidth(22) = 0
    Msf_GriAseg.Text = "EstPensión"
    Msf_GriAseg.Col = 23
    Msf_GriAseg.ColWidth(23) = 0
    Msf_GriAseg.Text = "PrcPensionGar"
    Msf_GriAseg.Col = 24
    Msf_GriAseg.ColWidth(24) = 0
    Msf_GriAseg.Text = "PrcPensionLeg"

Exit Function
Err_IniGri:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'-------------------------------------------------------------------
'SE CARGAN DATOS DEL BENEFICIARIO SELECCIONADO
'-------------------------------------------------------------------
Private Sub Msf_GriAseg_DblClick()
On Error GoTo Err_Grilla

    Msf_GriAseg.Col = 0
    Msf_GriAseg.Row = Msf_GriAseg.RowSel
    If (Msf_GriAseg.Text = "") Or (Msf_GriAseg.Row = 0) Then
        Exit Sub
    End If

    'Call flLimpiarDatosAseg
    Fra_DatosBenef.Enabled = True

    Lbl_NumOrden = Msf_GriAseg.Text

    Msf_GriAseg.Col = 1
    vlCodPar = fgObtenerCodigo_TextoCompuesto(Msf_GriAseg.Text)
    vgI = fgBuscarPosicionCodigoCombo(vlCodPar, Cmb_Parentesco)
    If (Cmb_Parentesco.ListCount > 0) Then
        Cmb_Parentesco.ListIndex = vgI
    End If

    If Msf_GriAseg.Text = cgCauCodPar Then
        MsgBox "Los Datos del Afiliado Deben ser Modificados en Carpeta Afiliado", vbExclamation, "Aviso"
        Fra_DatosBenef.Enabled = False
    Else
        If (vlSwCalExtOK = True) And (vlSwCalIntOK = True) Then
            Fra_DatosBenef.Enabled = False
        Else
            Fra_DatosBenef.Enabled = True
        End If
    End If

    Msf_GriAseg.Col = 2
    vlCodGruFam = fgObtenerCodigo_TextoCompuesto(Msf_GriAseg.Text)
    vgI = fgBuscarPosicionCodigoCombo(vlCodGruFam, Cmb_GrupoFam)
    If (Cmb_GrupoFam.ListCount > 0) Then
        Cmb_GrupoFam.ListIndex = vgI
    End If

    Msf_GriAseg.Col = 3
    vlCodSexoBen = fgObtenerCodigo_TextoCompuesto(Msf_GriAseg.Text)
    vgI = fgBuscarPosicionCodigoCombo(vlCodSexoBen, Cmb_SexoBen)
    If (Cmb_SexoBen.ListCount > 0) Then
        Cmb_SexoBen.ListIndex = vgI
    End If

    Msf_GriAseg.Col = 4
    vlCodSitInv = fgObtenerCodigo_TextoCompuesto(Msf_GriAseg.Text)
    vgI = fgBuscarPosicionCodigoCombo(vlCodSitInv, Cmb_SitInv)
    If (Cmb_SitInv.ListCount > 0) Then
        Cmb_SitInv.ListIndex = vgI
    End If

    Msf_GriAseg.Col = 5
    If Msf_GriAseg.Text <> "" Then
        Txt_FecInvBen = Msf_GriAseg.Text
    Else
        Txt_FecInvBen = ""
    End If

    Msf_GriAseg.Col = 6
    If Msf_GriAseg.Text <> "" Then
        Lbl_CauInvBen = Trim(Msf_GriAseg.Text) & " - " & fgBuscarGlosaCauInv(Msf_GriAseg.Text)
    Else
        Lbl_CauInvBen = ""
    End If

    Msf_GriAseg.Col = 7
    If Msf_GriAseg.Text <> "" Then
        Lbl_DerPension = Trim(Msf_GriAseg.Text) & " - " & fgBuscarGlosaElemento(vgCodTabla_DerPen, Msf_GriAseg.Text)
    Else
        Lbl_DerPension = ""
    End If

    Msf_GriAseg.Col = 8
    'Derecho a crecer

    Msf_GriAseg.Col = 9
    If Msf_GriAseg.Text <> "" Then
        'Lbl_FecNacBen = DateSerial(Mid(Msf_GriAseg.Text, 1, 4), Mid(Msf_GriAseg.Text, 5, 2), Mid(Msf_GriAseg.Text, 7, 2))
        Txt_FecNacBen = Msf_GriAseg.Text
    Else
        Txt_FecNacBen = ""
    End If

    Msf_GriAseg.Col = 10
    'fecha nacimiento hijo menor

    Msf_GriAseg.Col = 11
    'Tipo Identificación
    If Msf_GriAseg.Text <> "" Then
        vgPalabra = Trim(Mid(Msf_GriAseg, 1, InStr(1, Msf_GriAseg, "-") - 1))
        vgI = fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_TipoIdentBen)
        If (Cmb_TipoIdentBen.ListCount > 0) Then
            Cmb_TipoIdentBen.ListIndex = vgI
        End If
    End If

    Msf_GriAseg.Col = 12
    If Msf_GriAseg.Text <> "" Then
        Txt_NumIdentBen = Msf_GriAseg.Text
    Else
        Txt_NumIdentBen = ""
    End If

    Msf_GriAseg.Col = 13
    Txt_NombresBen = Msf_GriAseg.Text

    Msf_GriAseg.Col = 14
    Txt_NombresBenSeg = Msf_GriAseg.Text

    Msf_GriAseg.Col = 15
    Txt_ApPatBen = Msf_GriAseg.Text

    Msf_GriAseg.Col = 16
    Txt_ApMatBen = Msf_GriAseg.Text

    Msf_GriAseg.Col = 17
    Lbl_Porcentaje = Format(Msf_GriAseg.Text, "#,#0.00")

    Msf_GriAseg.Col = 18
    Lbl_PensionBen = Format(Msf_GriAseg.Text, "#,#0.00")

    Msf_GriAseg.Col = 19
    Lbl_PenGar = Format(Msf_GriAseg.Text, "#,#0.00")

    Msf_GriAseg.Col = 20
    If Msf_GriAseg.Text <> "" Then
        'Lbl_FecFallBen = DateSerial(Mid(Msf_GriAseg.Text, 1, 4), Mid(Msf_GriAseg.Text, 5, 2), Mid(Msf_GriAseg.Text, 7, 2))
        Txt_FecFallBen = Msf_GriAseg.Text
    Else
        Txt_FecFallBen = ""
    End If

    Msf_GriAseg.Col = 22
    'Estado de Pago Pensión

    If Cmb_TipoIdentBen.Enabled = True Then
        Cmb_TipoIdentBen.SetFocus
    Else
        If Txt_NombresBen.Enabled = True Then
            Txt_NombresBen.SetFocus
        End If
    End If

Exit Sub
Err_Grilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmb_TipoIdent_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_NumIdent.SetFocus
End If
End Sub

Private Sub Txt_AnnosDif_Change()
If Not IsNumeric(Txt_AnnosDif) Then
    Txt_AnnosDif.Text = ""
End If
End Sub

Private Sub Txt_ApoAdi_Change()
If Not IsNumeric(Txt_ApoAdi) Then
    Txt_ApoAdi = ""
    Lbl_PriUnica = ""
Else
    If (Len(Txt_ApoAdi) = 0) Then
        Lbl_PriUnica = ""
    End If
End If
End Sub

Private Sub Txt_ApoAdi_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_ApoAdi = Format(Txt_ApoAdi, "#,#0.00")
    SSTab_Poliza.Tab = 2
    Cmb_TipoIdentBen.SetFocus
End If
End Sub

Private Sub Txt_ApoAdi_LostFocus()
Txt_ApoAdi = Format(Txt_ApoAdi, "#,#0.00")
If (IsNumeric(Txt_BonoAct)) And (IsNumeric(Txt_CtaInd)) And (IsNumeric(Txt_ApoAdi)) Then
    Lbl_PriUnica = Format(CDbl(Txt_BonoAct) + CDbl(Txt_CtaInd) + CDbl(Txt_ApoAdi), "#,#0.00")
End If
End Sub

Private Sub Txt_BonoAct_Change()
If Not IsNumeric(Txt_BonoAct) Then
    Txt_BonoAct = ""
    Lbl_PriUnica = ""
Else
    If (Len(Txt_BonoAct) = 0) Then
        Lbl_PriUnica = ""
    End If
End If
End Sub

Private Sub Txt_BonoAct_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_BonoAct = Format(Txt_BonoAct, "#,#0.00")
    Txt_ApoAdi.SetFocus
End If
End Sub

Private Sub Txt_BonoAct_LostFocus()
Txt_BonoAct = Format(Txt_BonoAct, "#,#0.00")
If (IsNumeric(Txt_BonoAct)) And (IsNumeric(Txt_CtaInd)) And (IsNumeric(Txt_ApoAdi)) Then
    Lbl_PriUnica = Format(CDbl(Txt_BonoAct) + CDbl(Txt_CtaInd) + CDbl(Txt_ApoAdi), "#,#0.00")
End If
End Sub

Private Sub Txt_ComInt_Change()
If Not IsNumeric(txt_conint) Then
    Txt_ComInt.Text = ""
End If
End Sub

Private Sub Txt_CtaInd_Change()
If Not IsNumeric(Txt_CtaInd) Then
    Txt_CtaInd = ""
    Lbl_PriUnica = ""
Else
    If (Len(Txt_CtaInd) = 0) Then
        Lbl_PriUnica = ""
    End If
End If
End Sub

Private Sub Txt_CtaInd_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_CtaInd = Format(Txt_CtaInd, "#,#0.00")
    Txt_BonoAct.SetFocus
End If
End Sub

Private Sub Txt_CtaInd_LostFocus()
Txt_CtaInd = Format(Txt_CtaInd, "#,#0.00")
If (IsNumeric(Txt_BonoAct)) And (IsNumeric(Txt_CtaInd)) And (IsNumeric(Txt_ApoAdi)) Then
    Lbl_PriUnica = Format(CDbl(Txt_BonoAct) + CDbl(Txt_CtaInd) + CDbl(Txt_ApoAdi), "#,#0.00")
End If
End Sub

Private Sub Txt_FecNacBen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_FecFallBen.SetFocus
End If
End Sub

Private Sub Txt_FecNacBen_LostFocus()
    If Txt_FecNacBen <> "" Then
        If (flValidaFecha(Txt_FecNacBen) = False) Then
            Txt_FecNacBen = ""
            Exit Sub
        End If
        Txt_FecNacBen.Text = Format(CDate(Trim(Txt_FecNacBen)), "yyyymmdd")
        Txt_FecNacBen.Text = DateSerial(Mid((Txt_FecNacBen.Text), 1, 4), Mid((Txt_FecNacBen.Text), 5, 2), Mid((Txt_FecNacBen.Text), 7, 2))
    End If
End Sub

Private Sub Txt_FecFallBen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Cmb_Parentesco.SetFocus
End If
End Sub

Private Sub Txt_FecFallBen_LostFocus()
    If Txt_FecFallBen <> "" Then
        If (flValidaFecha(Txt_FecFallBen) = False) Then
            Txt_FecFallBen = ""
            Exit Sub
        End If
        Txt_FecFallBen.Text = Format(CDate(Trim(Txt_FecFallBen)), "yyyymmdd")
        Txt_FecFallBen.Text = DateSerial(Mid((Txt_FecFallBen.Text), 1, 4), Mid((Txt_FecFallBen.Text), 5, 2), Mid((Txt_FecFallBen.Text), 7, 2))
    End If
End Sub

Private Sub Txt_MesesGar_Change()
If Not IsNumeric(Txt_MesesGar) Then
    Txt_MesesGar.Text = ""
End If
End Sub

Private Sub Txt_NumIdent_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_NomAfi.SetFocus
End If
End Sub

Private Sub Txt_NumIdent_LostFocus()
    Txt_NumIdent = Trim(UCase(Txt_NumIdent))
End Sub

Private Sub Txt_NomAfi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_NomAfiSeg.SetFocus
End If
End Sub

Private Sub Txt_NomAfi_LostFocus()
    Txt_NomAfi = Trim(UCase(Txt_NomAfi))
End Sub

Private Sub Txt_NomAfiSeg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_ApPatAfi.SetFocus
End If
End Sub

Private Sub Txt_NomAfiSeg_LostFocus()
    Txt_NomAfiSeg = Trim(UCase(Txt_NomAfiSeg))
End Sub

Private Sub Txt_ApPatAfi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_ApMatAfi.SetFocus
    End If
End Sub

Private Sub Txt_ApPatAfi_LostFocus()
    Txt_ApPatAfi.Text = Trim(UCase(Txt_ApPatAfi.Text))
End Sub

Private Sub Txt_ApMatAfi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Cmb_Sexo.SetFocus
End If
End Sub

Private Sub Txt_ApMatAfi_LostFocus()
    Txt_ApMatAfi.Text = Trim(UCase(Txt_ApMatAfi.Text))
End Sub

Private Sub Cmb_Sexo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_FecNac.SetFocus
End If
End Sub

Private Sub Txt_FecNac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_FecFall.SetFocus
End If
End Sub

Private Sub Txt_FecNac_LostFocus()
    If Txt_FecNac <> "" Then
        If (flValidaFecha(Txt_FecNac) = False) Then
            Txt_FecNac = ""
            Exit Sub
        End If
        Txt_FecNac.Text = Format(CDate(Trim(Txt_FecNac)), "yyyymmdd")
        Txt_FecNac.Text = DateSerial(Mid((Txt_FecNac.Text), 1, 4), Mid((Txt_FecNac.Text), 5, 2), Mid((Txt_FecNac.Text), 7, 2))
    End If
End Sub

Private Sub Txt_FecFall_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Cmb_TipoPension.SetFocus
End If
End Sub

Private Sub Txt_FecFall_LostFocus()
    If Txt_FecFall <> "" Then
        If (flValidaFecha(Txt_FecFall) = False) Then
            Txt_FecFall = ""
            Exit Sub
        End If
        Txt_FecFall.Text = Format(CDate(Trim(Txt_FecFall)), "yyyymmdd")
        Txt_FecFall.Text = DateSerial(Mid((Txt_FecFall.Text), 1, 4), Mid((Txt_FecFall.Text), 5, 2), Mid((Txt_FecFall.Text), 7, 2))
    End If
End Sub

Private Sub cmb_TipoPension_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_FecInv.SetFocus
End If
End Sub

Private Sub Txt_FecInv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Txt_FecInv <> "") Then
        Cmd_CauInv.SetFocus
    Else
        Cmb_Afp.SetFocus
    End If
End If
End Sub

Private Sub Txt_FecInv_LostFocus()
    If Txt_FecInv <> "" Then
        If (flValidaFecha(Txt_FecInv) = False) Then
            Txt_FecInv = ""
            Exit Sub
        End If
        Txt_FecInv.Text = Format(CDate(Trim(Txt_FecInv)), "yyyymmdd")
        Txt_FecInv.Text = DateSerial(Mid((Txt_FecInv.Text), 1, 4), Mid((Txt_FecInv.Text), 5, 2), Mid((Txt_FecInv.Text), 7, 2))
    End If
End Sub

Private Sub Cmb_Afp_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_EstCivil.SetFocus
End If
End Sub

Private Sub Cmb_EstCivil_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_Salud.SetFocus
End If
End Sub

Private Sub Cmb_Vejez_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_ViaPago.SetFocus
End If
End Sub

Private Sub Cmb_ViaPago_Click()
    Call flValidaViaPago
End Sub

Private Sub Cmb_ViaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cmb_Suc.Enabled = True Then
            Cmb_Suc.SetFocus
        Else
            If Cmb_TipCta.Enabled = False Then
                Txt_FecDev.SetFocus
                SSTab_Poliza.Tab = 1
            Else
                 Cmb_TipCta.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Cmb_suc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SSTab_Poliza.Tab = 1
    Txt_FecDev.SetFocus
End If
End Sub

Private Sub Cmb_Salud_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_Dir.SetFocus
End If
End Sub
Private Sub Txt_Dir_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmd_BuscarDir.SetFocus
End If
End Sub

Private Sub Txt_Dir_Lostfocus()
    Txt_Dir.Text = Trim(UCase(Txt_Dir.Text))
End Sub

Private Sub Txt_Fono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_Correo.SetFocus
    End If
End Sub

Private Sub Txt_Fono_LostFocus()
    Txt_Fono = Trim(UCase(Txt_Fono))

End Sub

Private Sub Txt_correo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_Nacionalidad.SetFocus
End If
End Sub

Private Sub Txt_Nacionalidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_Nacionalidad = Trim(Txt_Nacionalidad)
    If (Txt_Nacionalidad = "") Then
        Txt_Nacionalidad = cgTipoNacionalidad
    End If
    If (Cmb_Vejez.Enabled = True) Then
        Cmb_Vejez.SetFocus
    Else
        Cmb_ViaPago.SetFocus
    End If
End If
End Sub

Private Sub Txt_Nacionalidad_LostFocus()
    Txt_Nacionalidad.Text = Trim(UCase(Txt_Nacionalidad.Text))
End Sub

Private Sub Cmb_TipCta_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_Bco.SetFocus
End If
End Sub

Private Sub Cmb_Bco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_NumCta.SetFocus
End If
End Sub
Private Sub Txt_NumCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_NumCta = Trim(Txt_NumCta)
    SSTab_Poliza.Tab = 1
    Txt_FecDev.SetFocus
End If
End Sub

Private Sub Txt_FecDev_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_FecIncorpora.SetFocus
End If
End Sub

Private Sub Txt_FecDev_LostFocus()
    If Txt_FecDev <> "" Then
        If (Trim(Txt_FecDev) = "") Then
            Txt_FecDev = ""
            Exit Sub
        End If
        If Not IsDate(Txt_FecDev) Then
            Txt_FecDev = ""
            Exit Sub
        End If
        If (CDate(Txt_FecDev) > CDate(Date)) Then
            Txt_FecDev = ""
            Exit Sub
        End If
        If (Year(CDate(Txt_FecDev)) < 1900) Then
            Txt_FecDev = ""
            Exit Sub
        End If
        Txt_FecDev.Text = Format(CDate(Trim(Txt_FecDev)), "yyyymmdd")
        Txt_FecDev.Text = DateSerial(Mid((Txt_FecDev.Text), 1, 4), Mid((Txt_FecDev.Text), 5, 2), Mid((Txt_FecDev.Text), 7, 2))
    End If
End Sub

Private Sub Txt_FecIncorpora_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_FecIniPago.SetFocus
End If
End Sub

Private Sub Txt_FecIncorpora_LostFocus()
    If Txt_FecIncorpora <> "" Then
        If (Trim(Txt_FecIncorpora) = "") Then
            Txt_FecIncorpora = ""
            Exit Sub
        End If
        If Not IsDate(Txt_FecIncorpora) Then
            Txt_FecIncorpora = ""
            Exit Sub
        End If
        If (CDate(Txt_FecIncorpora) > CDate(Date)) Then
            Txt_FecIncorpora = ""
            Exit Sub
        End If
        If (Year(CDate(Txt_FecIncorpora)) < 1900) Then
            Txt_FecIncorpora = ""
            Exit Sub
        End If
        Txt_FecIncorpora.Text = Format(CDate(Trim(Txt_FecIncorpora)), "yyyymmdd")
        Txt_FecIncorpora.Text = DateSerial(Mid((Txt_FecIncorpora.Text), 1, 4), Mid((Txt_FecIncorpora.Text), 5, 2), Mid((Txt_FecIncorpora.Text), 7, 2))
    End If
End Sub
Private Sub Txt_FecIniPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_Cuspp.SetFocus
End If
End Sub

Private Sub Txt_FecIniPago_LostFocus()
    If Txt_FecIniPago <> "" Then
        If (flValidaFecha(Txt_FecIniPago) = False) Then
            Txt_FecIniPago = ""
            Exit Sub
        End If
        Txt_FecIniPago.Text = Format(CDate(Trim(Txt_FecIniPago)), "yyyymmdd")
        Txt_FecIniPago.Text = DateSerial(Mid((Txt_FecIniPago.Text), 1, 4), Mid((Txt_FecIniPago.Text), 5, 2), Mid((Txt_FecIniPago.Text), 7, 2))
    End If
End Sub

Private Sub Txt_Cuspp_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_Moneda.SetFocus
End If
End Sub

Private Sub Txt_Cuspp_LostFocus()
    Txt_Cuspp = Trim(UCase(Txt_Cuspp))
End Sub

Private Sub Cmb_Moneda_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_TipoRenta.SetFocus
End If
End Sub

Private Sub Cmb_Moneda_Click()
If (Trim(Cmb_Moneda <> "")) Then
    Dim vlMon As String, vlMonScomp As String
    vlMon = fgObtenerCodigo_TextoCompuesto(Cmb_Moneda)
    vlMonScomp = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vlMon)
    Lbl_Moneda(1) = vlMonScomp
    Lbl_Moneda(2) = vlMonScomp
    Lbl_Moneda(3) = vlMonScomp
    Lbl_Moneda(4) = vlMonScomp
End If
End Sub

Private Sub Cmb_TipoRenta_Click()
vgPalabra = fgObtenerCodigo_TextoCompuesto(Cmb_TipoRenta)
If (vgPalabra = "1") Then
    Txt_PrcRentaTmp = Format(0, "#0.00")
    Txt_AnnosDif = 0
    Txt_PrcRentaTmp.Enabled = False
    Txt_AnnosDif.Enabled = False
Else
    Txt_PrcRentaTmp = ""
    Txt_AnnosDif = ""
    Txt_PrcRentaTmp.Enabled = True
    Txt_AnnosDif.Enabled = True
End If
End Sub

Private Sub Cmb_TipoRenta_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Txt_AnnosDif.Enabled = True) Then
        Txt_AnnosDif.SetFocus
    Else
        Cmb_Modalidad.SetFocus
    End If
End If
End Sub

Private Sub Txt_AnnosDif_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_Modalidad.SetFocus
End If
End Sub

Private Sub Cmb_Modalidad_Click()
vgPalabra = fgObtenerCodigo_TextoCompuesto(Cmb_Modalidad)
If (vgPalabra = "1") Then
    Txt_MesesGar = 0
    Txt_MesesGar.Enabled = False
Else
    Txt_MesesGar = ""
    Txt_MesesGar.Enabled = True
End If
End Sub

Private Sub Cmb_Modalidad_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Txt_MesesGar.Enabled = True) Then
        Txt_MesesGar.SetFocus
    Else
        If (Txt_PrcRentaTmp.Enabled = True) Then
            Txt_PrcRentaTmp.SetFocus
        Else
            Cmb_IndCob.SetFocus
        End If
    End If
End If
End Sub

Private Sub Txt_MesesGar_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_PrcRentaTmp.SetFocus
End If
End Sub

Private Sub Txt_PrcRentaTmp_Change()
If Not IsNumeric(Txt_PrcRentaTmp) Then
    Txt_PrcRentaTmp.Text = ""
End If
End Sub

Private Sub Txt_PrcRentaTmp_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_IndCobertura.SetFocus
End If
End Sub

Private Sub Cmb_IndCob_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_CobConyuge.SetFocus
End If
End Sub

Private Sub Cmb_CobConyuge_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_DerCre.SetFocus
End If
End Sub

Private Sub Cmb_DerCre_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_DerGra.SetFocus
End If
End Sub

Private Sub Cmb_DerGra_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_ComInt.SetFocus
End If
End Sub

Private Sub Txt_ComInt_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_ComInt = Format(Txt_ComInt, "#,#0.00")
    Cmd_BuscaCor.SetFocus
End If
End Sub

Private Sub Txt_PrcFam_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_CtaInd.SetFocus
End If
End Sub

Private Sub Cmb_TipoIdentBen_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Txt_NumIdentBen.Enabled = True) Then
        Txt_NumIdentBen.SetFocus
    Else
        Txt_NombresBen.SetFocus
    End If
End If
End Sub

Private Sub Txt_NumIdentBen_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_NombresBen.SetFocus
End If
End Sub

Private Sub Txt_NombresBen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_NombresBenSeg.SetFocus
End If
End Sub

Private Sub Txt_NombresBen_LostFocus()
    Txt_NombresBen = Trim(UCase(Txt_NombresBen))
End Sub

Private Sub Txt_NombresBenSeg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_ApPatBen.SetFocus
End If
End Sub

Private Sub Txt_NombresBenSeg_LostFocus()
    Txt_NombresBenSeg = Trim(UCase(Txt_NombresBenSeg))
End Sub

Private Sub Txt_ApPatBen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_ApMatBen.SetFocus
End If
End Sub

Private Sub Txt_ApPatBen_LostFocus()
    Txt_ApPatBen.Text = Trim(UCase(Txt_ApPatBen.Text))
End Sub

Private Sub Txt_ApMatBen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_FecNacBen.SetFocus
    End If
End Sub

Private Sub Txt_ApMatBen_LostFocus()
    Txt_ApMatBen.Text = Trim(UCase(Txt_ApMatBen.Text))
End Sub

Private Sub Cmb_Parentesco_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_GrupoFam.SetFocus
End If
End Sub

Private Sub Cmb_GrupoFam_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_SexoBen.SetFocus
End If
End Sub

Private Sub Cmb_SexoBen_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_SitInv.SetFocus
End If
End Sub

Private Sub Cmb_SitInv_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_FecInvBen.SetFocus
End If
End Sub

Private Sub Txt_FecInvBen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Txt_FecInvBen <> "" Then
        If IsDate(Txt_FecInvBen) Then
            Cmd_BuscaCauInv.SetFocus
        End If
    Else
        Btn_Agregar.SetFocus
    End If
End If
End Sub


'----------------------------------------------------------------
'Valida Datos de Carpeta Afiliado
'----------------------------------------------------------------
Function flValDatAfi()
On Error GoTo Err_Datafi

    flValDatAfi = False
    If (Trim(Cmb_TipoIdent)) = "" Then
        MsgBox "Debe seleccionar el Tipo de Identificación del Afiliado.", vbCritical, "Error de Datos"
        Cmb_TipoIdent.SetFocus
        Exit Function
    End If
    If (Trim(Txt_NumIdent.Text)) = "" Then
        MsgBox "Debe Ingresar el Número de Identificación del Afiliado.", vbCritical, "Error de Datos"
        Txt_NumIdent.SetFocus
        Exit Function
    End If
    If Trim(Txt_NomAfi) = "" Then
        MsgBox "Debe Ingresar Nombre del Afiliado", vbExclamation, "Falta Información"
        Txt_NomAfi.SetFocus
        Exit Function
    End If
    If Trim(Txt_ApPatAfi) = "" Then
        MsgBox "Debe Ingresar Apellido Paterno del Afiliado", vbExclamation, "Falta Información"
        Txt_ApPatAfi.SetFocus
        Exit Function
    End If
'    If Trim(Txt_ApMatAfi) = "" Then
'        MsgBox "Debe Ingresar Apellido Materno del Afiliado", vbExclamation, "Falta Información"
'        Txt_ApMatAfi.SetFocus
'        Exit Function
'    End If
    If (Trim(Cmb_Sexo)) = "" Then
        MsgBox "Debe seleccionar el Sexo del Afiliado.", vbCritical, "Error de Datos"
        Cmb_Sexo.SetFocus
        Exit Function
    End If
    'Validación de Fecha de Nacimiento
    If fgValidaFecha(Trim(Txt_FecNac)) = False Then
        Txt_FecNac.Text = Format(CDate(Trim(Txt_FecNac)), "yyyymmdd")
        Txt_FecNac.Text = DateSerial(Mid((Txt_FecNac.Text), 1, 4), Mid((Txt_FecNac.Text), 5, 2), Mid((Txt_FecNac.Text), 7, 2))
    Else
'        MsgBox "Debe ingresar la Fecha de Nacimiento del Afiliado.", vbCritical, "Error de Datos"
        Txt_FecNac.SetFocus
        Exit Function
    End If
    'Validación de Fecha de Fallecimiento
    If Trim(Txt_FecFall) <> "" Then
        If fgValidaFecha(Trim(Txt_FecFall)) = False Then
            Txt_FecFall.Text = Format(CDate(Trim(Txt_FecFall)), "yyyymmdd")
            Txt_FecFall.Text = DateSerial(Mid((Txt_FecFall.Text), 1, 4), Mid((Txt_FecFall.Text), 5, 2), Mid((Txt_FecFall.Text), 7, 2))
        Else
            MsgBox "Debe ingresar correctamente la Fecha de Nacimiento del Afiliado.", vbCritical, "Error de Datos"
            Txt_FecFall.SetFocus
            Exit Function
        End If
    End If
    If (Trim(Cmb_TipoPension)) = "" Then
        MsgBox "Debe seleccionar el Tipo de Pensión del Afiliado.", vbCritical, "Error de Datos"
        Cmb_TipoPension.SetFocus
        Exit Function
    End If
    If Trim(Mid(Lbl_CauInv, 1, InStr(1, Lbl_CauInv, "-") - 1)) <> clCodCauInvNoInv Then
        If Not IsDate(Txt_FecInv) Then
            MsgBox "Debe ingresar la Fecha de Invalidez del Afiliado.", vbExclamation, "Error de Datos"
            Txt_FecInv.SetFocus
            Exit Function
        End If
    Else
        Txt_FecInv = ""
    End If
    If (Trim(Lbl_CauInv)) = "" Then
        MsgBox "Debe seleccionar la Causa de Invalidez del Afiliado.", vbCritical, "Error de Datos"
        Cmd_CauInv.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_Afp)) = "" Then
        MsgBox "Debe seleccionar la AFP del Afiliado.", vbCritical, "Error de Datos"
        Cmb_Afp.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_EstCivil)) = "" Then
        MsgBox "Debe seleccionar el Estado Civil del Afiliado.", vbCritical, "Error de Datos"
        Cmb_EstCivil.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_Salud)) = "" Then
        MsgBox "Debe seleccionar la Institución de Salud del Afiliado.", vbCritical, "Error de Datos"
        Cmb_Salud.SetFocus
        Exit Function
    End If
    If Trim(Txt_Dir) = "" Then
        MsgBox "Debe Ingresar Dirección del Afiliado", vbExclamation, "Falta Información"
        Txt_Dir.SetFocus
        Exit Function
    End If
    vlCodDireccion = vgCodDireccion
    If (vlCodDireccion = "") Then
        MsgBox "Debe selecionar la Dirección (Depto-Prov-Dist) del Afiliado", vbExclamation, "Falta Información"
        Cmd_BuscarDir.SetFocus
        Exit Function
    End If
    If Trim(Txt_Nacionalidad) = "" Then
        MsgBox "Debe Ingresar la Nacionalidad del Afiliado", vbExclamation, "Falta Información"
        Txt_Nacionalidad.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_Vejez)) = "" Then
        MsgBox "Debe seleccionar el Tipo de Vejez del Afiliado.", vbCritical, "Error de Datos"
        Cmb_Vejez.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_ViaPago)) = "" Then
        MsgBox "Debe seleccionar la Vía de Pago del Afiliado.", vbCritical, "Error de Datos"
        Cmb_ViaPago.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_Suc)) = "" Then
        MsgBox "Debe seleccionar la Sucursal del Afiliado.", vbCritical, "Error de Datos"
        Cmb_Suc.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_TipCta)) = "" Then
        MsgBox "Debe seleccionar el Tipo de Cuenta del Afiliado.", vbCritical, "Error de Datos"
        Cmb_TipCta.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_Bco)) = "" Then
        MsgBox "Debe seleccionar el Banco del Afiliado.", vbCritical, "Error de Datos"
        Cmb_Bco.SetFocus
        Exit Function
    End If
    
    flValDatAfi = True
    
Exit Function
Err_Datafi:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'------------------------------------------------
'Valida Datos de Carpeta Calculo
'------------------------------------------------
Function flValDatCal()
On Error GoTo ERR_VALCAL

    flValDatCal = False
    
    'valida fecha de devengue
    If fgValidaFecha(Trim(Txt_FecDev)) = False Then
        Txt_FecDev = Format(CDate(Trim(Txt_FecDev)), "yyyymmdd")
        Txt_FecDev = DateSerial(Mid((Txt_FecDev), 1, 4), Mid((Txt_FecDev), 5, 2), Mid((Txt_FecDev), 7, 2))
    Else
        MsgBox "Debe Ingresar la Fecha Devengue", vbExclamation, "Falta Información"
        Txt_FecDev.SetFocus
        Exit Function
    End If
    'valida fecha de Incorporación o Aceptación
    If fgValidaFecha(Trim(Txt_FecIncorpora)) = False Then
        Txt_FecIncorpora = Format(CDate(Trim(Txt_FecIncorpora)), "yyyymmdd")
        Txt_FecIncorpora = DateSerial(Mid((Txt_FecIncorpora), 1, 4), Mid((Txt_FecIncorpora), 5, 2), Mid((Txt_FecIncorpora), 7, 2))
    Else
        MsgBox "Debe Ingresar la Fecha de Incorporación o Aceptación", vbExclamation, "Falta Información"
        Txt_FecIncorpora.SetFocus
        Exit Function
    End If
    'valida fecha de Primer Pago
    If flValFecVen(Trim(Txt_FecIniPago)) = False Then
        Txt_FecIniPago.Text = Format(CDate(Trim(Txt_FecIniPago)), "yyyymmdd")
        Txt_FecIniPago.Text = DateSerial(Mid((Txt_FecIniPago.Text), 1, 4), Mid((Txt_FecIniPago.Text), 5, 2), Mid((Txt_FecIniPago.Text), 7, 2))
    Else
        MsgBox "Debe Ingresar la Fecha de Primer Pago", vbExclamation, "Falta Información"
        Txt_FecIniPago.SetFocus
        Exit Function
    End If
    
    'Valida consecutividad de las 3 fechas anteriores
    '(Txt_FecDev <= Txt_FecIncorpora <= Txt_FecIniPago)
    If (CDate(Txt_FecDev) > CDate(Txt_FecIncorpora)) Then
        MsgBox "La Fecha Devengue debe ser menor Fecha Incorp. a la Póliza", vbCritical, "Error de Datos"
        Txt_FecDev.SetFocus
        Exit Function
    End If
    If (CDate(Txt_FecIncorpora) > CDate(Txt_FecIniPago)) Then
        MsgBox "La Fecha Incorp. a la Póliza debe ser menor a la Fecha Inicio Primer Pago", vbCritical, "Error de Datos"
        Txt_FecIncorpora.SetFocus
        Exit Function
    End If
    
    If (Trim(Txt_Cuspp)) = "" Then
        MsgBox "Debe Ingresar el CUSPP.", vbCritical, "Error de Datos"
        Txt_Cuspp.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_Moneda)) = "" Then
        MsgBox "Debe Ingresar el Tipo de Moneda.", vbCritical, "Error de Datos"
        Cmb_Moneda.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_TipoRenta)) = "" Then
        MsgBox "Debe Ingresar el Tipo de Renta.", vbCritical, "Error de Datos"
        Cmb_TipoRenta.SetFocus
        Exit Function
    End If
    If Not IsNumeric(Txt_AnnosDif) Then
        MsgBox "Debe Ingresar la cantidad de Años Diferidos.", vbCritical, "Error de Datos"
        Exit Function
    End If
    If (Trim(Cmb_Modalidad)) = "" Then
        MsgBox "Debe Ingresar el Tipo de Modalidad.", vbCritical, "Error de Datos"
        Cmb_Modalidad.SetFocus
        Exit Function
    End If
    If Not IsNumeric(Txt_MesesGar) Then
        MsgBox "Debe Ingresar la cantidad de Meses Garantizados.", vbCritical, "Error de Datos"
        Exit Function
    End If
    If (Trim(Lbl_RentaAFP)) = "" Then
        MsgBox "Debe Ingresar la Rentabilidad de la AFP seleccionada.", vbCritical, "Error de Datos"
        Exit Function
    End If
    If (Trim(Txt_PrcRentaTmp)) = "" Then
        MsgBox "Debe Ingresar la Renta Temporal.", vbCritical, "Error de Datos"
        Exit Function
    End If
    If (Trim(Cmb_IndCob)) = "" Then
        MsgBox "Debe Ingresar el Estado de Cobertura.", vbCritical, "Error de Datos"
        Cmb_IndCob.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_CobConyuge)) = "" Then
        MsgBox "Debe Ingresar la Cobertura de Cónyuge.", vbCritical, "Error de Datos"
        Cmb_CobConyuge.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_DerCre)) = "" Then
        MsgBox "Debe Ingresar el Derecho Crecer.", vbCritical, "Error de Datos"
        Cmb_DerCre.SetFocus
        Exit Function
    End If
    If (Trim(Cmb_DerGra)) = "" Then
        MsgBox "Debe Ingresar la Gratificación.", vbCritical, "Error de Datos"
        Cmb_DerGra.SetFocus
        Exit Function
    End If
    If (Not IsNumeric(Txt_ComInt)) Then
        MsgBox "Debe Ingresar la Comisión del Intermediario.", vbCritical, "Error de Datos"
        Txt_ComInt.SetFocus
        Exit Function
    End If
    If (Trim(Lbl_TipoIdentCorr)) = "" Then
        MsgBox "Debe Ingresar el Tipo de Identificación del Intermediario.", vbCritical, "Error de Datos"
        Cmd_BuscaCor.SetFocus
        Exit Function
    End If
    If (Trim(Lbl_NumIdentCorr)) = "" Then
        MsgBox "Debe Ingresar el Número de Identificación del Intermediario.", vbCritical, "Error de Datos"
        Cmd_BuscaCor.SetFocus
        Exit Function
    End If
    If (Not IsNumeric(Lbl_TasaCtoEq)) Then
        MsgBox "Debe Ingresar la Tasa de Costo Equivalente.", vbCritical, "Error de Datos"
        Exit Function
    End If
    If (Not IsNumeric(Txt_CtaInd)) Then
        MsgBox "Debe Ingresar la Cuenta Individual.", vbCritical, "Error de Datos"
        Txt_CtaInd.SetFocus
        Exit Function
    End If
    If (Not IsNumeric(Txt_BonoAct)) Then
        MsgBox "Debe Ingresar el Bono de Reconocimiento.", vbCritical, "Error de Datos"
        Txt_BonoAct.SetFocus
        Exit Function
    End If
    
    flValDatCal = True
    
Exit Function
ERR_VALCAL:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flValidaBenGrilla() As Boolean
On Error GoTo Err_flValidaBenGrilla
        
    flValidaBenGrilla = False
    
    vlPos = Msf_GriAseg.rows - 1
    Msf_GriAseg.Row = 1
    Msf_GriAseg.Col = 11
    For vlI = 1 To vlPos
        Msf_GriAseg.Row = vlI
        
        Msf_GriAseg.Col = 11
        'Valida Tipo de Identificación
        If Msf_GriAseg.Text = "" Then
            MsgBox "Debe Ingresar Tipo de Identificación a Todos los Beneficiarios", vbExclamation, "Proceso Cancelado"
            Msf_GriAseg_DblClick
            Cmb_TipoIdentBen.SetFocus
            Exit Function
        End If

'I--- ABV 07/08/2007 ---
'Revisar Con MiCH
        vgPalabra = fgObtenerCodigo_TextoCompuesto(Msf_GriAseg)
        If (vgPalabra = cgTipoDocSinInformacion) Then
            MsgBox "Debe Ingresar un Tipo de Identificación válido a Todos los Beneficiarios", vbExclamation, "Proceso Cancelado"
            Msf_GriAseg_DblClick
            Cmb_TipoIdentBen.SetFocus
            Exit Function
        End If
'F--- ABV 07/08/2007 ---

        Msf_GriAseg.Col = 12
        'Valida Número de Identificación
        If Msf_GriAseg.Text = "" Then
            MsgBox "Debe Ingresar Número de Identificación  a Todos los Beneficiarios", vbExclamation, "Proceso Cancelado"
            Msf_GriAseg_DblClick
            Txt_NumIdentBen.SetFocus
            Exit Function
        End If
        
        'Valida Nombres
        Msf_GriAseg.Col = 13
        If Msf_GriAseg.Text = "" Then
            MsgBox "Debe Ingresar Nombres a Todos los Beneficiarios", vbExclamation, "Proceso Cancelado"
            Msf_GriAseg_DblClick
            Exit Function
        End If
        
        'Valida Apellido Paterno
        Msf_GriAseg.Col = 15
        If Msf_GriAseg.Text = "" Then
            MsgBox "Debe Ingresar Apellido Paterno a Todos los Beneficiarios", vbExclamation, "Proceso Cancelado"
            Msf_GriAseg_DblClick
            Txt_ApPatBen.SetFocus
            Exit Function
        End If
        
'        'Valida Apellido Materno
'        Msf_GriAseg.Col = 16
'        If Msf_GriAseg.Text = "" Then
'            MsgBox "Debe Ingresar Apellido Materno a Todos los Beneficiarios", vbExclamation, "Proceso Cancelado"
'            Msf_GriAseg_DblClick
'            Txt_ApMatBen.SetFocus
'            Exit Function
'        End If

        'Valida Fecha de Invalidez
        Msf_GriAseg.Col = 5
        If Msf_GriAseg.Text <> "" Then
            If flValFecVen(Msf_GriAseg.Text) Then
                MsgBox "La Fecha de Invalidez no es una fecha Válida para el Beneficiario", vbExclamation, "Proceso Cancelado"
                Msf_GriAseg_DblClick
                Txt_FecInvBen.SetFocus
                Exit Function
            End If
        End If
        'Valida Causal de Invalidez
        Msf_GriAseg.Col = 6
        If Msf_GriAseg.Text = "" Then
            MsgBox "Debe ingresar la Causal de Invalidez a Todos los Beneficiarios", vbExclamation, "Proceso Cancelado"
            Msf_GriAseg_DblClick
            Cmd_BuscarCauInvBen.SetFocus
            Exit Function
        End If
    Next vlI
    
    flValidaBenGrilla = True

Exit Function
Err_flValidaBenGrilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'---------------------------
'VALIDA FECHA DE VENCIMIENTO
'---------------------------
Function flValFecVen(iFecha As String)
flValFecVen = False

    If (Trim(iFecha) = "") Then
        MsgBox "Falta Ingresar Fecha", vbExclamation, "Error de Datos"
        flValFecVen = True
        Exit Function
    End If
    If Not IsDate(iFecha) Then
        MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
        flValFecVen = True
        Exit Function
    End If
    If (Year(CDate(iFecha)) < 1900) Then
        MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se pruede ingresar (1900).", vbCritical, "Error de Datos"
        flValFecVen = True
        Exit Function
    End If
    
End Function

