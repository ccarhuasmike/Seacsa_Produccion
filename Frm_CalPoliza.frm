VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_CalPoliza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Pre-Pólizas."
   ClientHeight    =   8265
   ClientLeft      =   3555
   ClientTop       =   1560
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   10650
   Begin VB.Frame Fra_Cabeza 
      Height          =   1215
      Left            =   150
      TabIndex        =   11
      Top             =   0
      Width           =   9255
      Begin VB.TextBox Txt_NumPol 
         BackColor       =   &H00E0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   290
         Width           =   1695
      End
      Begin VB.TextBox Txt_FecVig 
         Height          =   285
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   12
         Top             =   290
         Width           =   1335
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Nº de Cotizacion"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   21
         Top             =   765
         Width           =   1335
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Nº Poliza"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   20
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Fecha de Emisión"
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   19
         Top             =   285
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   360
         X2              =   8880
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Nº Operación"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   18
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Lbl_SolOfe 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5280
         TabIndex        =   17
         Top             =   765
         Width           =   1695
      End
      Begin VB.Label Lbl_Cabeza 
         Caption         =   "Nº Correlativo"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   16
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Lbl_SecOfe 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8400
         TabIndex        =   15
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Lbl_NumCot 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   765
         Width           =   1695
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Email de Bienvenida"
      Top             =   7110
      Width           =   10455
      Begin VB.CommandButton cmdEnviaCorreo 
         Caption         =   "Email Bienvenida"
         Height          =   675
         Left            =   4080
         Picture         =   "Frm_CalPoliza.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   299
         Top             =   240
         Width           =   1545
      End
      Begin VB.CommandButton Cmd_CrearPol 
         Caption         =   "&Crear Pre-Póliza"
         Height          =   675
         Left            =   720
         Picture         =   "Frm_CalPoliza.frx":10CA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Crear Póliza"
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   9120
         Picture         =   "Frm_CalPoliza.frx":19DC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   7440
         Picture         =   "Frm_CalPoliza.frx":1AD6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   5760
         Picture         =   "Frm_CalPoliza.frx":2190
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Eliminar Póliza"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   3120
         Picture         =   "Frm_CalPoliza.frx":24D2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Grabar Datos de Póliza"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   6600
         Picture         =   "Frm_CalPoliza.frx":2B8C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   8280
         Picture         =   "Frm_CalPoliza.frx":3246
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Cmd_Editar 
         Caption         =   "&Recalcular"
         Height          =   675
         Left            =   2160
         Picture         =   "Frm_CalPoliza.frx":3820
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Editar Datos y Recalcular"
         Top             =   240
         Width           =   870
      End
      Begin Crystal.CrystalReport Rpt_Poliza 
         Left            =   0
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.CommandButton Cmd_Poliza 
      Caption         =   "&Pre-Póliza"
      Height          =   615
      Left            =   9600
      Picture         =   "Frm_CalPoliza.frx":393E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Póliza"
      Top             =   30
      Width           =   900
   End
   Begin VB.CommandButton Cmd_Cotizacion 
      Caption         =   "&Cotizacion"
      Height          =   585
      Left            =   9600
      Picture         =   "Frm_CalPoliza.frx":3FF8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Buscar Cotización"
      Top             =   630
      Width           =   900
   End
   Begin TabDlg.SSTab SSTab_Poliza 
      Height          =   5775
      Left            =   120
      TabIndex        =   22
      Top             =   1350
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10186
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
      TabPicture(0)   =   "Frm_CalPoliza.frx":46B2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Afiliado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_Representante"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos de Cálculo"
      TabPicture(1)   =   "Frm_CalPoliza.frx":46CE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_Calculo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Datos de Beneficiarios"
      TabPicture(2)   =   "Frm_CalPoliza.frx":46EA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "framBancoCta"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Fra_Benef"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Fra_Representante 
         Caption         =   "Representante"
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
         Height          =   2670
         Left            =   120
         TabIndex        =   35
         Top             =   2880
         Visible         =   0   'False
         Width           =   5325
         Begin VB.ComboBox cmbSexoRep 
            Height          =   315
            ItemData        =   "Frm_CalPoliza.frx":4706
            Left            =   3960
            List            =   "Frm_CalPoliza.frx":4710
            TabIndex        =   303
            Text            =   "M"
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton CmdDireccionRep 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4800
            Picture         =   "Frm_CalPoliza.frx":471A
            Style           =   1  'Graphical
            TabIndex        =   300
            ToolTipText     =   "Efectuar Busqueda de Dirección"
            Top             =   120
            Width           =   465
         End
         Begin VB.TextBox txtCorreoRep 
            Height          =   285
            Left            =   1410
            TabIndex        =   43
            Top             =   2160
            Width           =   3375
         End
         Begin VB.TextBox txtTelRep2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3780
            TabIndex        =   42
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtTelRep1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1410
            TabIndex        =   41
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox Txt_NumIdRep 
            Height          =   285
            Left            =   1410
            MaxLength       =   16
            TabIndex        =   37
            Top             =   480
            Width           =   1665
         End
         Begin VB.TextBox Txt_NomRep 
            Height          =   285
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   38
            Top             =   825
            Width           =   3350
         End
         Begin VB.TextBox Txt_ApPatRep 
            Height          =   285
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   39
            Top             =   1140
            Width           =   3350
         End
         Begin VB.TextBox Txt_ApMatRep 
            Height          =   285
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   40
            Top             =   1455
            Width           =   3350
         End
         Begin VB.CommandButton Cmd_SalirRep 
            Height          =   360
            Left            =   4785
            Picture         =   "Frm_CalPoliza.frx":4B5C
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Salir de Dirección"
            Top             =   1920
            Width           =   405
         End
         Begin VB.ComboBox Cmb_TipIdRep 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   180
            Width           =   3350
         End
         Begin VB.Label Label9 
            Caption         =   "Sexo:"
            Height          =   255
            Left            =   3120
            TabIndex        =   302
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Correo:"
            Height          =   255
            Left            =   120
            TabIndex        =   296
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Telefono 2:"
            Height          =   375
            Left            =   2760
            TabIndex        =   295
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Teléfono 1:"
            Height          =   375
            Left            =   120
            TabIndex        =   294
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Nro. Identif."
            Height          =   255
            Index           =   15
            Left            =   135
            TabIndex        =   49
            Top             =   555
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Paterno"
            Height          =   255
            Index           =   17
            Left            =   135
            TabIndex        =   48
            Top             =   1185
            Width           =   930
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo Identif."
            Height          =   255
            Index           =   18
            Left            =   135
            TabIndex        =   47
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Nombres"
            Height          =   255
            Index           =   19
            Left            =   135
            TabIndex        =   46
            Top             =   870
            Width           =   825
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Ap. Materno"
            Height          =   255
            Index           =   16
            Left            =   135
            TabIndex        =   45
            Top             =   1485
            Width           =   960
         End
      End
      Begin VB.Frame Fra_Benef 
         Height          =   5265
         Left            =   -74880
         TabIndex        =   230
         Top             =   360
         Width           =   10215
         Begin VB.TextBox Txt_FecFallBen 
            Height          =   285
            Left            =   3720
            MaxLength       =   10
            TabIndex        =   232
            Top             =   3840
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CommandButton Btn_Agregar 
            Height          =   450
            Left            =   9480
            Picture         =   "Frm_CalPoliza.frx":4C56
            Style           =   1  'Graphical
            TabIndex        =   231
            ToolTipText     =   "Agregar Cobertura"
            Top             =   1800
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid Msf_GriAseg 
            Height          =   1335
            Left            =   120
            TabIndex        =   233
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
            Height          =   3645
            Left            =   120
            TabIndex        =   234
            Top             =   1560
            Width           =   9975
            Begin VB.TextBox txtCorreoBen 
               Height          =   285
               Left            =   1320
               TabIndex        =   298
               Top             =   1920
               Width           =   3375
            End
            Begin VB.CommandButton Btn_Porcentaje 
               Height          =   450
               Left            =   9360
               Picture         =   "Frm_CalPoliza.frx":4DE0
               Style           =   1  'Graphical
               TabIndex        =   248
               ToolTipText     =   "Calcular Porcentajes"
               Top             =   1200
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CommandButton Btn_Quita 
               Height          =   450
               Left            =   9360
               Picture         =   "Frm_CalPoliza.frx":53A2
               Style           =   1  'Graphical
               TabIndex        =   249
               ToolTipText     =   "Eliminar Cobertura"
               Top             =   720
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox Txt_NombresBen 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   252
               Top             =   760
               Width           =   3375
            End
            Begin VB.TextBox Txt_ApPatBen 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   251
               Top             =   1320
               Width           =   3375
            End
            Begin VB.TextBox Txt_ApMatBen 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   250
               Top             =   1600
               Width           =   3375
            End
            Begin VB.ComboBox Cmb_TipoIdentBen 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPoliza.frx":552C
               Left            =   1320
               List            =   "Frm_CalPoliza.frx":552E
               Style           =   2  'Dropdown List
               TabIndex        =   247
               Top             =   160
               Width           =   3375
            End
            Begin VB.TextBox Txt_NumIdentBen 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   246
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox Txt_NombresBenSeg 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   245
               Top             =   1030
               Width           =   3375
            End
            Begin VB.CommandButton Cmd_BuscarCauInvBen 
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
               TabIndex        =   244
               Top             =   1478
               Width           =   285
            End
            Begin VB.TextBox Txt_FecInvBen 
               Height          =   285
               Left            =   6120
               MaxLength       =   25
               TabIndex        =   243
               Top             =   1200
               Width           =   1215
            End
            Begin VB.TextBox txtTutor 
               BackColor       =   &H80000018&
               Height          =   285
               Left            =   6120
               Locked          =   -1  'True
               TabIndex        =   242
               Top             =   2320
               Width           =   2880
            End
            Begin VB.CommandButton cmdTutor 
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
               Height          =   270
               Left            =   9015
               TabIndex        =   241
               Top             =   2320
               Width           =   285
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Datos de Cta. Bancaria"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3000
               TabIndex        =   240
               Top             =   3060
               Width           =   1695
            End
            Begin VB.ComboBox cboNacionalidadBen 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPoliza.frx":5530
               Left            =   1320
               List            =   "Frm_CalPoliza.frx":5532
               Style           =   2  'Dropdown List
               TabIndex        =   239
               Top             =   2670
               Width           =   3390
            End
            Begin VB.CheckBox chkConTratDatos_Ben 
               Caption         =   "Consentimiento de Tratamiento de Datos"
               Height          =   375
               Left            =   4920
               TabIndex        =   238
               Top             =   2880
               Width           =   4575
            End
            Begin VB.CheckBox chkConUsoDatosCom_Ben 
               Caption         =   "Consentimiento de uso de datos para fines comerciales"
               Height          =   375
               Left            =   4920
               TabIndex        =   237
               Top             =   3240
               Width           =   4575
            End
            Begin VB.TextBox Txt_Fono2_Ben 
               Height          =   285
               Left            =   8520
               MaxLength       =   15
               TabIndex        =   236
               Top             =   2640
               Width           =   1335
            End
            Begin VB.TextBox Txt_Fono1_Ben 
               Height          =   285
               Left            =   6120
               MaxLength       =   15
               TabIndex        =   235
               Top             =   2600
               Width           =   1335
            End
            Begin VB.Label Label8 
               Caption         =   "Correo"
               Height          =   255
               Left            =   120
               TabIndex        =   297
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Label Lbl_FecFallBen 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3600
               TabIndex        =   271
               Top             =   2280
               Width           =   1095
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Causal Invalidez"
               Height          =   255
               Index           =   13
               Left            =   4920
               TabIndex        =   290
               Top             =   1478
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Fec. Invalidez"
               Height          =   255
               Index           =   12
               Left            =   4920
               TabIndex        =   289
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Pensión"
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   288
               Top             =   3000
               Width           =   855
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Fec. Fallec."
               Height          =   195
               Index           =   10
               Left            =   2640
               TabIndex        =   287
               Top             =   2280
               Width           =   825
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Porcentaje"
               Height          =   255
               Index           =   7
               Left            =   4920
               TabIndex        =   286
               Top             =   2040
               Width           =   855
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "1er.  Nombre"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   285
               Top             =   765
               Width           =   915
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Ap. Paterno"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   284
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Ap. Materno"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   283
               Top             =   1600
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Sexo"
               Height          =   255
               Index           =   4
               Left            =   4920
               TabIndex        =   282
               Top             =   690
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Parentesco"
               Height          =   255
               Index           =   5
               Left            =   4920
               TabIndex        =   281
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Sit. Invalidez"
               Height          =   255
               Index           =   6
               Left            =   4920
               TabIndex        =   280
               Top             =   930
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Fec. Nac."
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   279
               Top             =   2280
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Dº a Pensión"
               Height          =   255
               Index           =   8
               Left            =   4920
               TabIndex        =   278
               Top             =   1750
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Grupo Familiar"
               Height          =   255
               Index           =   9
               Left            =   4920
               TabIndex        =   277
               Top             =   390
               Width           =   1215
            End
            Begin VB.Label Lbl_NumOrden 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   4320
               TabIndex        =   276
               Top             =   480
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Pensión Gar."
               Height          =   195
               Index           =   14
               Left            =   120
               TabIndex        =   275
               Top             =   3300
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Ident."
               Height          =   195
               Index           =   15
               Left            =   120
               TabIndex        =   274
               Top             =   240
               Width           =   765
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
               TabIndex        =   273
               Top             =   0
               Width           =   1020
            End
            Begin VB.Label Lbl_FecNacBen 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   272
               Top             =   2280
               Width           =   1095
            End
            Begin VB.Label Lbl_Porcentaje 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6120
               TabIndex        =   270
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label Lbl_PensionBen 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   269
               Top             =   3000
               Width           =   1095
            End
            Begin VB.Label Lbl_PenGar 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   268
               Top             =   3300
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Lbl_Par 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6120
               TabIndex        =   267
               Top             =   120
               Width           =   2895
            End
            Begin VB.Label Lbl_Grupo 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6120
               TabIndex        =   266
               Top             =   390
               Width           =   2895
            End
            Begin VB.Label Lbl_SexoBen 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6120
               TabIndex        =   265
               Top             =   660
               Width           =   2895
            End
            Begin VB.Label Lbl_SitInvBen 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6120
               TabIndex        =   264
               Top             =   930
               Width           =   2895
            End
            Begin VB.Label Lbl_CauInvBen 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6120
               TabIndex        =   263
               Top             =   1478
               Width           =   2895
            End
            Begin VB.Label Lbl_DerPension 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6120
               TabIndex        =   262
               Top             =   1750
               Width           =   2895
            End
            Begin VB.Label Lbl_Nombre 
               AutoSize        =   -1  'True
               Caption         =   "Nº Ident."
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   261
               Top             =   480
               Width           =   630
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "2do. Nombre"
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   260
               Top             =   1030
               Width           =   915
            End
            Begin VB.Label Lbl_Moneda 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   259
               Top             =   3060
               Width           =   375
            End
            Begin VB.Label Lbl_Moneda 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   258
               Top             =   3300
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label lbltutor 
               Caption         =   "Tutor "
               Height          =   225
               Left            =   4920
               TabIndex        =   257
               Top             =   2320
               Width           =   900
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Nacionalidad"
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   256
               Top             =   2700
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Teléfono 1"
               Height          =   255
               Left            =   4920
               TabIndex        =   255
               Top             =   4000
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "Teléfono 1"
               Height          =   255
               Left            =   4920
               TabIndex        =   254
               Top             =   2640
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "Teléfono 2"
               Height          =   255
               Left            =   7560
               TabIndex        =   253
               Top             =   2640
               Width           =   975
            End
         End
      End
      Begin VB.Frame Fra_Calculo 
         Height          =   5265
         Left            =   -74880
         TabIndex        =   141
         Top             =   360
         Width           =   10215
         Begin VB.Frame Fra_DatCal 
            Height          =   4515
            Left            =   120
            TabIndex        =   158
            Top             =   120
            Width           =   9975
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
               Left            =   9520
               TabIndex        =   161
               Top             =   1660
               Width           =   285
            End
            Begin VB.TextBox Txt_PrcFam 
               Height          =   285
               Left            =   6600
               MaxLength       =   10
               TabIndex        =   160
               Top             =   4080
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox Txt_FecIniPago 
               Height          =   285
               Left            =   3360
               MaxLength       =   10
               TabIndex        =   159
               Top             =   765
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   40
               Left            =   7380
               TabIndex        =   229
               Top             =   465
               Width           =   255
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Fec. Cálculo"
               Height          =   255
               Index           =   39
               Left            =   240
               TabIndex        =   228
               Top             =   765
               Width           =   1575
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Fec. Incorp. a la Póliza"
               Height          =   255
               Index           =   38
               Left            =   240
               TabIndex        =   227
               Top             =   465
               Width           =   1695
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Estado Cobertura"
               Height          =   255
               Index           =   37
               Left            =   4920
               TabIndex        =   226
               Top             =   165
               Width           =   1695
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "CUSPP"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   225
               Top             =   1065
               Width           =   540
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Nº Ident. Inter."
               Height          =   255
               Index           =   7
               Left            =   4920
               TabIndex        =   224
               Top             =   1965
               Width           =   1335
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Mto. Prima Uni Dif"
               Height          =   255
               Index           =   61
               Left            =   240
               TabIndex        =   223
               Top             =   4070
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Mto. Prima Uni Sim"
               Height          =   255
               Index           =   60
               Left            =   240
               TabIndex        =   222
               Top             =   3770
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Lbl_MtoPensionGar 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   221
               Top             =   3760
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Lbl_MtoPension 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   220
               Top             =   3465
               Width           =   1215
            End
            Begin VB.Label Lbl_TasaPerGar 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   219
               Top             =   3165
               Width           =   735
            End
            Begin VB.Label Lbl_TasaTIR 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   218
               Top             =   2865
               Width           =   735
            End
            Begin VB.Label Lbl_TasaVta 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   217
               Top             =   2565
               Width           =   735
            End
            Begin VB.Label Lbl_TasaCtoEq 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   216
               Top             =   2265
               Width           =   735
            End
            Begin VB.Label Lbl_MtoPrimaUniDif 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   215
               Top             =   4070
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Lbl_MtoPrimaUniSim 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   214
               Top             =   3770
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Lbl_FacPenElla 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   213
               Top             =   465
               Width           =   735
            End
            Begin VB.Label Lbl_PrcRentaTmp 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   212
               Top             =   3465
               Width           =   975
            End
            Begin VB.Label Lbl_RentaAFP 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   211
               Top             =   3165
               Width           =   975
            End
            Begin VB.Label Lbl_MesesGar 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   210
               Top             =   2865
               Width           =   735
            End
            Begin VB.Label Lbl_Alter 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   209
               Top             =   2565
               Width           =   2775
            End
            Begin VB.Label Lbl_AnnosDif 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   208
               Top             =   2265
               Width           =   735
            End
            Begin VB.Label Lbl_TipoRenta 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   207
               Top             =   1965
               Width           =   2775
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   22
               Left            =   9600
               TabIndex        =   206
               Top             =   1065
               Visible         =   0   'False
               Width           =   165
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   21
               Left            =   3000
               TabIndex        =   205
               Top             =   3465
               Width           =   255
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   20
               Left            =   3000
               TabIndex        =   204
               Top             =   3165
               Width           =   255
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Tasa Int. Período Gar."
               Height          =   195
               Index           =   17
               Left            =   4905
               TabIndex        =   203
               Top             =   3165
               Width           =   1590
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Com Inter C/Benef"
               Height          =   195
               Index           =   16
               Left            =   7350
               TabIndex        =   202
               Top             =   1065
               Visible         =   0   'False
               Width           =   1320
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Mto. Pensión Gar. "
               Height          =   195
               Index           =   15
               Left            =   4920
               TabIndex        =   201
               Top             =   3760
               Visible         =   0   'False
               Width           =   1320
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Mto. Pensión "
               Height          =   195
               Index           =   14
               Left            =   4905
               TabIndex        =   200
               Top             =   3465
               Width           =   975
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tasa TIR"
               Height          =   255
               Index           =   13
               Left            =   4905
               TabIndex        =   199
               Top             =   2865
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tasa Venta"
               Height          =   255
               Index           =   11
               Left            =   4905
               TabIndex        =   198
               Top             =   2565
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tasa Cto. Equiv."
               Height          =   255
               Index           =   8
               Left            =   4905
               TabIndex        =   197
               Top             =   2265
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Rentabilidad AFP"
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   196
               Top             =   3165
               Width           =   1455
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Prc. Renta Temporal"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   195
               Top             =   3465
               Width           =   1575
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Cobertura Cónyuge"
               Height          =   195
               Index           =   9
               Left            =   4920
               TabIndex        =   194
               Top             =   465
               Width           =   1365
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Tipo de Renta"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   193
               Top             =   1965
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Modalidad"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   192
               Top             =   2565
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Meses Gar."
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   191
               Top             =   2865
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Fec. Devengue"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   190
               Top             =   165
               Width           =   1575
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Años Dif."
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   189
               Top             =   2265
               Width           =   735
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Ident. Inter."
               Height          =   195
               Index           =   2
               Left            =   4905
               TabIndex        =   188
               Top             =   1665
               Width           =   1170
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Factor CNU"
               Height          =   195
               Index           =   7
               Left            =   4920
               TabIndex        =   187
               Top             =   4080
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Lbl_TipoIdentCorr 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   186
               Top             =   1665
               Width           =   2895
            End
            Begin VB.Label Lbl_NumIdentCorr 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   185
               Top             =   1965
               Width           =   2895
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Moneda"
               Height          =   195
               Index           =   18
               Left            =   240
               TabIndex        =   184
               Top             =   1365
               Width           =   945
            End
            Begin VB.Label Lbl_Moneda 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   0
               Left            =   3360
               TabIndex        =   183
               Top             =   360
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Derecho Crecer"
               Height          =   255
               Index           =   19
               Left            =   4920
               TabIndex        =   182
               Top             =   765
               Width           =   1455
            End
            Begin VB.Label Lbl_DerCre 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   181
               Top             =   765
               Width           =   735
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Gratificación"
               Height          =   255
               Index           =   23
               Left            =   4920
               TabIndex        =   180
               Top             =   1065
               Width           =   975
            End
            Begin VB.Label Lbl_DerGra 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   179
               Top             =   1065
               Width           =   735
            End
            Begin VB.Label Lbl_IndCob 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   178
               Top             =   165
               Width           =   735
            End
            Begin VB.Label Lbl_BenSocial 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   8760
               TabIndex        =   177
               Top             =   1365
               Width           =   720
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Ben. Social"
               Height          =   255
               Index           =   24
               Left            =   7560
               TabIndex        =   176
               Top             =   1365
               Width           =   855
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "Comisión Inter."
               Height          =   255
               Index           =   25
               Left            =   4920
               TabIndex        =   175
               Top             =   1365
               Width           =   1455
            End
            Begin VB.Label Lbl_ComIntBen 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   8760
               TabIndex        =   174
               Top             =   1065
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Lbl_Calculo 
               Caption         =   "%"
               Height          =   255
               Index           =   26
               Left            =   7380
               TabIndex        =   173
               Top             =   1365
               Width           =   165
            End
            Begin VB.Label Lbl_FecDev 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   172
               Top             =   165
               Width           =   1095
            End
            Begin VB.Label Lbl_Cuspp 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   171
               Top             =   1065
               Width           =   2775
            End
            Begin VB.Label Lbl_FecIncorpora 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   170
               Top             =   465
               Width           =   1095
            End
            Begin VB.Label Lbl_Moneda 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   3
               Left            =   7880
               TabIndex        =   169
               Top             =   3465
               Width           =   375
            End
            Begin VB.Label Lbl_Moneda 
               Caption         =   "(TM)"
               Height          =   255
               Index           =   4
               Left            =   7880
               TabIndex        =   168
               Top             =   3760
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label Lbl_ComInt 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   6600
               TabIndex        =   167
               Top             =   1365
               Width           =   735
            End
            Begin VB.Label Lbl_FecCalculo 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   166
               Top             =   765
               Width           =   1095
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Reajuste Trimestral"
               Height          =   195
               Index           =   27
               Left            =   240
               TabIndex        =   165
               Top             =   1665
               Width           =   1350
            End
            Begin VB.Label Lbl_ReajusteTipo 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3360
               TabIndex        =   164
               Top             =   1665
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Lbl_ReajusteValor 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   163
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Lbl_ReajusteMoneda 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1920
               TabIndex        =   162
               Top             =   1365
               Width           =   2775
            End
         End
         Begin VB.Frame Fra_SumaBono 
            Height          =   615
            Left            =   120
            TabIndex        =   142
            Top             =   4560
            Width           =   9975
            Begin VB.Label Lbl_PriUnica 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   8400
               TabIndex        =   157
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Lbl_BonoAct 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3120
               TabIndex        =   156
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "CI."
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   155
               Top             =   255
               Width           =   195
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "BR."
               Height          =   195
               Index           =   0
               Left            =   2400
               TabIndex        =   154
               Top             =   255
               Width           =   270
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
               TabIndex        =   153
               Top             =   195
               Width           =   255
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
               Left            =   7005
               TabIndex        =   152
               Top             =   195
               Width           =   165
            End
            Begin VB.Label Lbl_MonedaFon 
               Alignment       =   1  'Right Justify
               Caption         =   "(TM)"
               Height          =   255
               Index           =   0
               Left            =   315
               TabIndex        =   151
               Top             =   255
               Width           =   375
            End
            Begin VB.Label Lbl_MonedaFon 
               Alignment       =   1  'Right Justify
               Caption         =   "(TM)"
               Height          =   255
               Index           =   1
               Left            =   2715
               TabIndex        =   150
               Top             =   255
               Width           =   375
            End
            Begin VB.Label Lbl_MonedaFon 
               Alignment       =   1  'Right Justify
               Caption         =   "(TM)"
               Height          =   255
               Index           =   2
               Left            =   7920
               TabIndex        =   149
               Top             =   255
               Width           =   375
            End
            Begin VB.Label Lbl_CtaInd 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   720
               TabIndex        =   148
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "Prima U."
               Height          =   195
               Index           =   2
               Left            =   7320
               TabIndex        =   147
               Top             =   255
               Width           =   600
            End
            Begin VB.Label Lbl_MonedaFon 
               Alignment       =   1  'Right Justify
               Caption         =   "(TM)"
               Height          =   255
               Index           =   3
               Left            =   5115
               TabIndex        =   146
               Top             =   255
               Width           =   375
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
               Left            =   4560
               TabIndex        =   145
               Top             =   195
               Width           =   255
            End
            Begin VB.Label Lbl_SumaBono 
               AutoSize        =   -1  'True
               Caption         =   "AA."
               Height          =   195
               Index           =   6
               Left            =   4800
               TabIndex        =   144
               Top             =   255
               Width           =   255
            End
            Begin VB.Label Lbl_ApoAdi 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5520
               TabIndex        =   143
               Top             =   255
               Width           =   1455
            End
         End
      End
      Begin VB.Frame Fra_Afiliado 
         Height          =   5295
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   10215
         Begin VB.Frame Fra_Direccion 
            Caption         =   "Dirección"
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
            Height          =   2535
            Left            =   5280
            TabIndex        =   51
            Top             =   120
            Visible         =   0   'False
            Width           =   4875
            Begin VB.ComboBox Cmb_TipoVia 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   885
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   240
               Width           =   1380
            End
            Begin VB.TextBox Txt_NombreVia 
               Height          =   300
               Left            =   3060
               MaxLength       =   50
               TabIndex        =   59
               Top             =   240
               Width           =   1725
            End
            Begin VB.TextBox Txt_Numero 
               Height          =   285
               Left            =   885
               MaxLength       =   4
               TabIndex        =   58
               Top             =   645
               Width           =   495
            End
            Begin VB.TextBox Txt_Interior 
               Height          =   285
               Left            =   3060
               MaxLength       =   4
               TabIndex        =   57
               Top             =   645
               Width           =   495
            End
            Begin VB.TextBox Txt_NombreZona 
               Height          =   285
               Left            =   3060
               MaxLength       =   50
               TabIndex        =   56
               Top             =   1005
               Width           =   1725
            End
            Begin VB.ComboBox Cmb_TipoZona 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   885
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Top             =   1005
               Width           =   1380
            End
            Begin VB.TextBox Txt_Referencia 
               Height          =   285
               Left            =   885
               MaxLength       =   40
               TabIndex        =   54
               Top             =   2160
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
               Left            =   4230
               Style           =   1  'Graphical
               TabIndex        =   53
               ToolTipText     =   "Efectuar Busqueda de Dirección"
               Top             =   1800
               Width           =   300
            End
            Begin VB.CommandButton Cmd_SalirDir 
               Height          =   360
               Left            =   4185
               Picture         =   "Frm_CalPoliza.frx":5534
               Style           =   1  'Graphical
               TabIndex        =   52
               ToolTipText     =   "Salir de Dirección"
               Top             =   2145
               Width           =   405
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Nom.Vía"
               Height          =   255
               Index           =   5
               Left            =   2295
               TabIndex        =   73
               Top             =   300
               Width           =   630
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Tipo Vía"
               Height          =   255
               Index           =   11
               Left            =   45
               TabIndex        =   72
               Top             =   300
               Width           =   720
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Número"
               Height          =   255
               Index           =   9
               Left            =   45
               TabIndex        =   71
               Top             =   705
               Width           =   645
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Interior"
               Height          =   255
               Index           =   13
               Left            =   2280
               TabIndex        =   70
               Top             =   720
               Width           =   585
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Tipo Zona"
               Height          =   255
               Index           =   10
               Left            =   45
               TabIndex        =   69
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Nom.Zona"
               Height          =   255
               Index           =   12
               Left            =   2295
               TabIndex        =   68
               Top             =   1080
               Width           =   765
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Referencia"
               Height          =   255
               Index           =   14
               Left            =   45
               TabIndex        =   67
               Top             =   2205
               Width           =   1005
            End
            Begin VB.Label Lbl_Distrito 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   885
               TabIndex        =   66
               Top             =   1785
               Width           =   3255
            End
            Begin VB.Label Lbl_Provincia 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3060
               TabIndex        =   65
               Top             =   1410
               Width           =   1725
            End
            Begin VB.Label Lbl_Departamento 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   885
               TabIndex        =   64
               Top             =   1410
               Width           =   1650
            End
            Begin VB.Label Lbl_Afiliado 
               Caption         =   "Distrito"
               Height          =   255
               Index           =   12
               Left            =   45
               TabIndex        =   63
               Top             =   1860
               Width           =   615
            End
            Begin VB.Label Lbl_Afiliado 
               Caption         =   "Prov."
               Height          =   255
               Index           =   11
               Left            =   2625
               TabIndex        =   62
               Top             =   1455
               Width           =   420
            End
            Begin VB.Label Lbl_Afiliado 
               Caption         =   "Dpto."
               Height          =   255
               Index           =   15
               Left            =   45
               TabIndex        =   61
               Top             =   1455
               Width           =   405
            End
         End
         Begin VB.ComboBox cboNacionalidad 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPoliza.frx":562E
            Left            =   1440
            List            =   "Frm_CalPoliza.frx":5630
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   2250
            Width           =   3540
         End
         Begin VB.TextBox Lbl_Dir 
            Height          =   550
            Left            =   6600
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   109
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox Txt_Fono2_Afil 
            Height          =   285
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   108
            Top             =   970
            Width           =   1815
         End
         Begin VB.CheckBox chkConUsoDatosCom_Afil 
            Caption         =   "Consentimiento de uso de datos para fines comerciales"
            Height          =   375
            Left            =   5520
            TabIndex        =   107
            Top             =   2160
            Width           =   4575
         End
         Begin VB.CheckBox chkConTratDatos_Afil 
            Caption         =   "Consentimiento de Tratamiento de Datos"
            Height          =   375
            Left            =   5520
            TabIndex        =   106
            Top             =   1845
            Width           =   4575
         End
         Begin VB.TextBox Txt_FecInv 
            Height          =   285
            Left            =   1450
            MaxLength       =   25
            TabIndex        =   105
            Top             =   3210
            Width           =   1095
         End
         Begin VB.TextBox Txt_NomAfiSeg 
            Height          =   285
            Left            =   1450
            MaxLength       =   50
            TabIndex        =   104
            Top             =   1090
            Width           =   3540
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
            Left            =   5030
            TabIndex        =   103
            Top             =   3450
            Width           =   285
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
            Height          =   2565
            Left            =   5330
            TabIndex        =   89
            Top             =   2520
            Width           =   4755
            Begin VB.ComboBox cmb_MonCta 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   1050
               Style           =   2  'Dropdown List
               TabIndex        =   292
               Top             =   1520
               Width           =   3350
            End
            Begin VB.TextBox Txt_NumCta 
               Height          =   315
               Left            =   1050
               MaxLength       =   25
               TabIndex        =   95
               Top             =   1850
               Width           =   3315
            End
            Begin VB.ComboBox Cmb_ViaPago 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   1050
               Style           =   2  'Dropdown List
               TabIndex        =   94
               Top             =   240
               Width           =   3350
            End
            Begin VB.ComboBox Cmb_Suc 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               ItemData        =   "Frm_CalPoliza.frx":5632
               Left            =   1050
               List            =   "Frm_CalPoliza.frx":5634
               Style           =   2  'Dropdown List
               TabIndex        =   93
               Top             =   540
               Width           =   3350
            End
            Begin VB.ComboBox Cmb_TipCta 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   1050
               Style           =   2  'Dropdown List
               TabIndex        =   92
               Top             =   1180
               Width           =   3350
            End
            Begin VB.ComboBox Cmb_Bco 
               BackColor       =   &H00E0FFFF&
               Height          =   315
               Left            =   1050
               Style           =   2  'Dropdown List
               TabIndex        =   91
               Top             =   850
               Width           =   3350
            End
            Begin VB.TextBox txt_CCI 
               Height          =   315
               Left            =   1050
               MaxLength       =   30
               TabIndex        =   90
               Top             =   2160
               Width           =   3315
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Mon. Cta."
               Height          =   255
               Index           =   26
               Left            =   120
               TabIndex        =   293
               Top             =   1560
               Width           =   825
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Tipo Cta."
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   102
               Top             =   1230
               Width           =   825
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Vía Pago"
               Height          =   255
               Index           =   0
               Left            =   130
               TabIndex        =   101
               Top             =   240
               Width           =   810
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Banco"
               Height          =   255
               Index           =   3
               Left            =   135
               TabIndex        =   100
               Top             =   880
               Width           =   825
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "N°Cuenta"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   99
               Top             =   1900
               Width           =   795
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "Sucursal"
               Height          =   255
               Index           =   1
               Left            =   135
               TabIndex        =   98
               Top             =   555
               Width           =   930
            End
            Begin VB.Label lblTipmonCab 
               Caption         =   "Moneda Cta."
               Height          =   255
               Left            =   120
               TabIndex        =   97
               Top             =   555
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Lbl_Nombre 
               Caption         =   "N°CCI"
               Height          =   255
               Index           =   23
               Left            =   120
               TabIndex        =   96
               Top             =   2220
               Width           =   795
            End
         End
         Begin VB.TextBox Txt_Correo 
            Height          =   285
            Left            =   6600
            MaxLength       =   40
            TabIndex        =   88
            Top             =   1260
            Width           =   3255
         End
         Begin VB.ComboBox Cmb_Salud 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   3555
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   4095
            Width           =   1455
         End
         Begin VB.TextBox Txt_Fono 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   86
            Top             =   670
            Width           =   1815
         End
         Begin VB.TextBox Txt_NomAfi 
            Height          =   285
            Left            =   1450
            MaxLength       =   50
            TabIndex        =   85
            Top             =   800
            Width           =   3540
         End
         Begin VB.TextBox Txt_ApPatAfi 
            Height          =   285
            Left            =   1450
            MaxLength       =   50
            TabIndex        =   84
            Top             =   1380
            Width           =   3540
         End
         Begin VB.TextBox Txt_ApMatAfi 
            Height          =   285
            Left            =   1450
            MaxLength       =   50
            TabIndex        =   83
            Top             =   1650
            Width           =   3540
         End
         Begin VB.ComboBox Cmb_EstCivil 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPoliza.frx":5636
            Left            =   1450
            List            =   "Frm_CalPoliza.frx":5638
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   4095
            Width           =   1305
         End
         Begin VB.ComboBox Cmb_TipoIdent 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPoliza.frx":563A
            Left            =   1450
            List            =   "Frm_CalPoliza.frx":563C
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   180
            Width           =   3540
         End
         Begin VB.TextBox Txt_NumIdent 
            Height          =   285
            Left            =   1450
            MaxLength       =   20
            TabIndex        =   80
            Top             =   500
            Width           =   2055
         End
         Begin VB.TextBox Txt_Asegurados 
            BackColor       =   &H00E0FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4410
            TabIndex        =   79
            Top             =   500
            Width           =   585
         End
         Begin VB.TextBox Txt_Nacionalidad 
            Height          =   285
            Left            =   6600
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   78
            Top             =   1590
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.ComboBox Cmb_Vejez 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPoliza.frx":563E
            Left            =   6600
            List            =   "Frm_CalPoliza.frx":5640
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   1540
            Width           =   3255
         End
         Begin VB.CommandButton Cmd_Direccion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   9690
            Picture         =   "Frm_CalPoliza.frx":5642
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Efectuar Busqueda de Dirección"
            Top             =   195
            Width           =   465
         End
         Begin VB.CommandButton Cmd_Representante 
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
            Left            =   5030
            TabIndex        =   75
            Top             =   4365
            Width           =   285
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Fecha Fallec."
            Height          =   240
            Index           =   16
            Left            =   2640
            TabIndex        =   132
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Lbl_FecFall 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3750
            TabIndex        =   117
            Top             =   2610
            Width           =   1245
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "2do. Nombre"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   140
            Top             =   1090
            Width           =   915
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Nº Identificación"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   139
            Top             =   510
            Width           =   1170
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Identificación"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   138
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label Lbl_TipPen 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1455
            TabIndex        =   137
            Top             =   2910
            Width           =   3540
         End
         Begin VB.Label Lbl_SexoAfi 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1450
            TabIndex        =   136
            Top             =   1950
            Width           =   3540
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Causal Invalidez"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   135
            Top             =   3510
            Width           =   1335
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Fec. Invalidez"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   134
            Top             =   3210
            Width           =   1095
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Tipo Pensión"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   133
            Top             =   2910
            Width           =   1095
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Correo"
            Height          =   255
            Index           =   14
            Left            =   5505
            TabIndex        =   131
            Top             =   1305
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "AFP"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   130
            Top             =   3780
            Width           =   1140
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Nº Benef."
            Height          =   255
            Index           =   13
            Left            =   3640
            TabIndex        =   129
            Top             =   525
            Width           =   735
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Inst.Salud"
            Height          =   255
            Index           =   6
            Left            =   2805
            TabIndex        =   128
            Top             =   4125
            Width           =   750
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Ap. Paterno"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   127
            Top             =   1380
            Width           =   975
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Ap. Materno"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   126
            Top             =   1650
            Width           =   975
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Sexo"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   125
            Top             =   1950
            Width           =   975
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Fecha Nac."
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   124
            Top             =   2610
            Width           =   1095
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Est. Civil"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   123
            Top             =   4095
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Teléfono 1"
            Height          =   255
            Index           =   9
            Left            =   5505
            TabIndex        =   122
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   10
            Left            =   5505
            TabIndex        =   121
            Top             =   195
            Width           =   765
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "1er. Nombre"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   120
            Top             =   800
            Width           =   975
         End
         Begin VB.Label Lbl_CauInv 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1455
            TabIndex        =   119
            Top             =   3510
            Width           =   3540
         End
         Begin VB.Label Lbl_Afp 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1455
            TabIndex        =   118
            Top             =   3780
            Width           =   3540
         End
         Begin VB.Label Lbl_FecNac 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1455
            TabIndex        =   116
            Top             =   2610
            Width           =   1095
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Nacionalidad"
            Height          =   255
            Index           =   21
            Left            =   5505
            TabIndex        =   115
            Top             =   1590
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Lbl_Afiliado 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Vejez"
            Height          =   195
            Index           =   22
            Left            =   5505
            TabIndex        =   114
            Top             =   1610
            Width           =   750
         End
         Begin VB.Label Lbl_Representante 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1455
            TabIndex        =   113
            Top             =   4455
            Width           =   3540
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Representante"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   112
            Top             =   4470
            Width           =   1335
         End
         Begin VB.Label Lbl_Afiliado 
            Caption         =   "Nacionalidad"
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   111
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label LblTelefono2 
            Caption         =   "Teléfono 2"
            Height          =   255
            Left            =   5520
            TabIndex        =   110
            Top             =   970
            Width           =   855
         End
      End
      Begin VB.Frame framBancoCta 
         Caption         =   "Datos de Cuenta"
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
         Height          =   2535
         Left            =   -70080
         TabIndex        =   23
         Top             =   2040
         Visible         =   0   'False
         Width           =   4395
         Begin VB.ComboBox cmbBancoCtaBen 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   960
            Width           =   3350
         End
         Begin VB.ComboBox cmbMonctaBen 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            ItemData        =   "Frm_CalPoliza.frx":5A84
            Left            =   960
            List            =   "Frm_CalPoliza.frx":5A86
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Width           =   3350
         End
         Begin VB.ComboBox cmbTipoCtaBen 
            BackColor       =   &H00E0FFFF&
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   240
            Width           =   3350
         End
         Begin VB.TextBox txtNumctaBen 
            Height          =   285
            Left            =   960
            MaxLength       =   25
            TabIndex        =   26
            Top             =   1320
            Width           =   3350
         End
         Begin VB.CommandButton cmd_bancoCta 
            Caption         =   "OK"
            Height          =   375
            Left            =   3840
            TabIndex        =   25
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox txt_CCIBen 
            Height          =   285
            Left            =   960
            MaxLength       =   30
            TabIndex        =   24
            Top             =   1660
            Width           =   3350
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Mon. Cta."
            Height          =   255
            Index           =   20
            Left            =   135
            TabIndex        =   34
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N°Cuenta"
            Height          =   255
            Index           =   21
            Left            =   135
            TabIndex        =   33
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Banco"
            Height          =   255
            Index           =   22
            Left            =   135
            TabIndex        =   32
            Top             =   960
            Width           =   825
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "Tipo Cta."
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N°CCI"
            Height          =   255
            Index           =   25
            Left            =   135
            TabIndex        =   30
            Top             =   1680
            Width           =   555
         End
      End
   End
   Begin VB.Label Lbl_Nombre 
      Caption         =   "Nro. Identif."
      Height          =   255
      Index           =   27
      Left            =   0
      TabIndex        =   301
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Lbl_Afiliado 
      Caption         =   "Nacionalidad"
      Height          =   255
      Index           =   24
      Left            =   360
      TabIndex        =   291
      Top             =   4230
      Width           =   1545
   End
End
Attribute VB_Name = "Frm_CalPoliza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlExiste As Boolean, vlSw As Boolean
'VARIABLES DE POLIZA
Dim vlNumPol As String, vlNumCot As String, vlCodAFP As String
Dim vlCodIsapre As String, vlCodTipPen As String, vlCodVejez As String
Dim vlCodEstCivil As String

'Const vlPrcFacPenella = "0"
Const vlMtoUf = 99999.99
Dim vlMontoUf As Double
Dim vlEstCivil As String, vlTipoIden As String, vlDgv As String, vlNumIden As String
Dim vlDir As String, vlCodDir As String, vlFono As String
Dim vlCorreo As String, vlCodViaPago As String, vlCodTipCta As String
Dim vlCodBco As String, vlNumCta As String, vlCodSuc As String, vlCodMonCta As String
'GCP-FRACTAL 08042019
Dim vlNum_Cuenta_CCI As String
Dim vlFecVig As String, vlFecDev As String, vlAnnoJub As Integer
Dim vlFecEmision As String 'ABV 08/07/2007
Dim vlFecCalculo As String 'ABV 10/08/2007
Dim vlNumBen As Integer, vlTipoIdencor As String, vlNumIdenCor As String
Dim vlComCor As String, vlValUf As String, vlCodTipRen As String
Dim vlMesDif As Long, vlCodMod As String, vlMesGar As Long
Dim vlRtaAfp As String, vlRtaTmp As String, vlFacPenElla As String
Dim vlCuoMor As String, vlCtaInd As String, vlMtoBono As String, vlMtoApoAdi As String
Dim vlPriUni As String, vlTasaTCE As String, vlTasaVta As String
Dim vlTasaTir As String, vlTasaPG As String, vlCNU As String
Dim vlPension As String, vlPenGar As String, vlSucUsu As String
Dim vlPriUniSim As String, vlPriUniDif As String, vlCtaIndAfp As String
Dim vlRtaTmpAFP As String, vlResMat As String, vlValPPTmp As String
Dim vlSexoAfi As String, vlAnnoNac As Integer, vlAño As String
Dim vlNum_idensup As String, vlNum_idenjef As String


'-- Begin: ADD by : ricardo.huerta 10-01-2018
Dim vl_nacionalidad As String
Dim vl_nacionalidadben As String
Dim vl_nacionalidadben_descripcion As String
'-- End   : ADD by : ricardo.huerta 10-01-2018


Dim vlCodMonedaFon As String, vlMtoValMonedaFon As String
Dim vlPriUniFon As String, vlCtaIndFon As String, vlMtoBonoFon As String
Dim vlApoAdiFon As String, vlTasaRetPro As String
Dim vlIndBenSocial As String, vlIndCobertura As String
Dim vlCodCoberCon As String, vlDerGra As String
Dim vlPriUniMod As String, vlCtaIndMod As String, vlMtoBonoMod As String, vlMtoApoAdiMod As String
Dim vlFecIniPerDif As String, vlFecIniPerGar As String
Dim vlFecFinPerDif As String, vlFecFinPerGar As String
Dim vlFecIniPagoPen As String
Dim vlReCalculo As String
Dim vlMtoSumPension As String, vlMtoPenAnual As String
Dim vlMtoRMPension As String, vlMtoRMGtoSep As String
Dim vlMtoRMGtoSepRV As String
Dim vlSucCorredor As String
Dim vlDerCrePol As String
Dim vlMtoAjusteIPC As String

'VARIABLES DE BENEFICIARIO
Dim vlNumOrden As Integer, vlCodPar As String, vlGruFam As String
Dim vlSexoBen As String, vlSitInv As String, vlFecInv As String
Dim vlDerPen As String, vlDerCre As String, vlFecNacBen As String
Dim vlFecFallBen As String, vlNomBen As String, vlNomBenSeg As String, vlPatBen As String
Dim vlMatBen As String, vlEdadHM As Integer, vlPrcPension As String
Dim vlPrcPensionLeg As String, vlPrcPensionRep As String
Dim vlPrcPensionGar As String
Dim vlRutBen As Integer, vlDgvBen As String, vlPenBen As String
Dim vlPenGarBen As String, vlCauInv As String, vlJ As Integer
Dim vlRutGrilla As String
Dim vlEstPen As String
'RRR
Dim vlcod_tipcta, vlcod_monbco, vlCod_Banco, vlnum_ctabco As String
Dim vlBTipoEnvio As Boolean

'VARIABLES DE BENEFICIARIO
Dim vlBonoActPesos As String, vlChk As String, vlFecNac As String
Dim vlNumEdadCobro As Integer, vlFecVen As String

'Variables Nuevas CMV 20050524
Dim vlCodTipoBono As String

Dim vlCodSecOfe As Integer
Dim vlNumPoliza As String
Dim vlNumCorrelativo As Integer
Dim vlCodSolOfe As String
Dim vlCodIdeOfe As String
Dim vlNumArchivo As Integer
Dim vlGlsDir As String
Dim vlFecSolicitud As String
Dim vlCodMoneda As String
Dim vlMtoValMoneda As Double
Dim vlCodTipReajuste As String, vlMtoValReajusteTri As Double, vlMtoValReajusteMen As Double  'I--- ABV 05/02/2011 ---
Dim vlPrcCorCom As Double, vlPrcCorComReal As Double
Dim vlCodCorAfeIva As String
Dim vlNumAnnoJub As Integer
Dim vlNumCargas As Integer
Dim vlNumMesDif As Integer
Dim vlPrcRentaAfpOri As String
Dim vlMtoFacPenElla As Double
Dim vlMtoCuoMor As Double
Dim vlMtoCorCom As Double
Dim vlMtoPerCon As Double
Dim vlPrcPerCon As Double
Dim vlCodEld As String
Dim vlMtoEld As Double
Dim vlCodTipEld As String
Dim vlMtoEldOfeUf As Double
Dim vlCodTipCot As String
Dim vlCodEstCot As String
Dim vlCodUsuario As String
Dim vlPrcFacPenElla As Double
Dim vlFecNacHM As String
Dim vlDgvGrilla As String

Dim vlTablaDetCotizacion As String

'Variables para creacion de registro de bono
'Dim vlCodTipoBono As String
Dim vlMtoValNom As Double
Dim vlFecEmi As String
'Dim vlFecVen As String
Dim vlPrcTasaInt As Double
Dim vlMtoBonoAct As Double
Dim vlMtoBonoActUF As Double
Dim vlMtoCompra As Double
Dim vlCodAfeLey As String
Dim vlNumEdadCob As Integer
Dim vlNumOrdenCot As Integer
Dim vlBotonEscogido As String

'Estado de Cotización Aceptada (No tiene Poliza)
Const clCodEstCotA As String * 1 = "A"
'Estado de Cotización Poliza (Tiene Poliza)
Const clCodEstCotP As String * 1 = "P"
Const clNumOrdenCau As Integer = 99
Const clNumOrden1 As Integer = 1

'Const clCodTipCotOfe As String * 1 = "O"
'Const clCodTipCotExt As String * 1 = "E"
'Const clCodTipCotRmt As String * 1 = "R"

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

Dim vlElemento As String
Dim vlCodDireccion As String, vlNacionalidad As String
Dim vlFecIncorporacion As String, vlFecPriPago As String
Dim vlcuspp As String, vlFecRepPrimasEst As String

Dim vlLargoTipoIden    As Integer 'sirve para llenar la grilla
Dim vlPosicionTipoIden As Integer 'sirve para llenar la grilla
Dim vlNumIdent         As Integer
Dim vlCodTipoIden As Integer 'sirve para guardar el código
Dim vlSwViaPago As Boolean
Dim vlRepresentante As String, vlDocum As String
Dim objRep As New ClsReporte 'RRR1305/2013
'I--- ABV 04/12/2009 ---
Dim vlMarcaSobDif As String
'F--- ABV 04/12/2009 ---
'RRR 18/9/13
Dim vlfecDevSol As String

'INICIO GCP-FRACTAL 11042019
Dim vl_Fono As String
Dim vl_Fono2_Afil As String
Dim VlConTratDatos_Afil As String
Dim VlConUsoDatosCom_Afil As String
Dim vl_Fono1_Ben As String
Dim vl_Fono2_Ben As String
Dim vl_ConTratDatos_Ben As String
Dim vl_ConUsoDatosCom_Ben As String
Dim vl_Correo_Ben As String
'FIN GCP-FRACTAL 11042019



'Integracion GobiernoDeDatos_(Se crean las variables a usar)
Dim CadenaJSON As String
Dim codComuna As String
Dim codRegion As String
Dim codProvincia As String
Public pgCodDireccion As String
Public pTipoTelefono As String
Public pNumTelefono As String
Public pCodigoTelefono As String
Public pTipoTelefono2 As String
Public pNumTelefono2 As String
Public pCodigoTelefono2 As String
Public pTipoVia As String
Public pDireccion As String
Public pNumero As String
Public pTipoPref As String
Public pInterior As String
Public pManzana As String
Public pLote As String
Public pEtapa As String
Public pTipoConj As String
Public pConjHabit As String
Public pTipoBlock As String
Public pNumBlock As String
Public pReferencia As String
Dim vDireccionConcat As String

'Datos a enviar
Dim vTipoDoc  As String
Dim vNumDoc As String
Dim vNomben As String
Dim vApePben As String
Dim vApeMben As String
Dim vSexoBen As String
Dim vIncapaciben As String
Dim vBirtdhayben As String
Dim vFecIncapaBen As String
Dim vFechaMatriBen As String
Dim vEstadoCivil As String
Dim vNacionaBen As String
Dim vParentesco As String
Dim vNumPoliza As String
Dim vlNumCoti As String
Dim vlNumOrd As String

Dim vl_Correo_afil As String

Private Type DatosPoliza
    Num_Cot As String
    Num_Poliza As String
    Cod_AFP As String
    Cod_Cuspp As String
    GLS_CORREO As String
    Cod_Moneda As String
    MTO_PRIUNI As Double
    MONEDA_RRVV As String
    COD_USUARIO As String
    GLS_TIPOIDEN As String
    TIPO_DOC As String
    Num_IdenBen As String
    nombres As String
    Gls_CorreoBen As String
    GLS_FONO As String
    MODALIDAD_RENTA As String
    TIPO_PRESTACION As String
    Summary As String
    Nombre_Afp As String
    Tipo_Renta As String
    NombreRep As String
    ApeRep As String
    Num_idenrep As String
    cod_tipoidenRep As String
    celularRep As String
    gls_correorep As String
    TIPO_PENSION As String
    VAL_TIPO_PENSION As String
   
End Type

Private Type Firmantes
    tipo As String
    NUM_IDEN  As String
    GLS_NOMBRES As String
    GLS_APEPAT As String
    GLS_APEMAT As String
    GLS_CORREO As String
    ID_FIRMANTE As String
    TIPO_FIRMA As String
    FIRMA_TIPO As String
    PARENTESCO As String
    celular As String
    genero As String
    TIPO_DOCUMENTO As String
    MENORDEEDAD As String
    Direccion As String
    departamento As String
    provincia As String
    distrito As String
End Type


Private Type RepresentaGC
        P_DATOS As String
        P_NTIPCONT As String
        P_SIDDOC As String
        P_NIDDOC_TYPE As String
        P_SNOMBRES As String
        P_SAPEPAT As String
        P_SAPEMAT As String
        P_SPHONE As String
End Type

Private Type AgentesCom
      NombreAsesor As String
      MailAsesor As String
      NombreSupervisor As String
      MailSupervisor As String
End Type


Dim dPol As DatosPoliza
Dim oFirmante As Firmantes
Dim LST_FIRMANTES() As Firmantes
Dim v_id_transac As Integer
Dim d_AgenteCom As AgentesCom




  

'Fin Integracion GobiernoDeDatos_

Function flIniGrillaBen()
On Error GoTo Err_IniGri

    Msf_GriAseg.Clear
    Msf_GriAseg.Cols = 37
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
    Msf_GriAseg.ColWidth(13) = 1500
    Msf_GriAseg.Text = " 1er. Nombre"
    
    Msf_GriAseg.Col = 14
    Msf_GriAseg.ColWidth(14) = 1500
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
    
'I--- ABV 10/08/2007 ---
    Msf_GriAseg.Col = 22
    Msf_GriAseg.ColWidth(22) = 0
    Msf_GriAseg.Text = "EstPensión"
    
    Msf_GriAseg.Col = 23
    'Msf_GriAseg.ColWidth(23) = 0
    Msf_GriAseg.Text = "PrcPensionGar"
    
    Msf_GriAseg.Col = 24
    'Msf_GriAseg.ColWidth(24) = 0
    Msf_GriAseg.Text = "PrcPensionLeg"
'F--- ABV 10/08/2007 ---
'´RRR 26122013
    Msf_GriAseg.Col = 25
    Msf_GriAseg.ColWidth(25) = 0
    Msf_GriAseg.Text = "TipoCta"
    
    Msf_GriAseg.Col = 26
    Msf_GriAseg.ColWidth(26) = 0
    Msf_GriAseg.Text = "Mon.Cuenta"
    
    Msf_GriAseg.Col = 27
    Msf_GriAseg.ColWidth(27) = 0
    Msf_GriAseg.Text = "Banco"
    
    Msf_GriAseg.Col = 28
    Msf_GriAseg.ColWidth(28) = 0
    Msf_GriAseg.Text = "Numero"
    
    '-- Begin : Modify by : ricardo.huerta
    Msf_GriAseg.Col = 29
    Msf_GriAseg.ColWidth(29) = 0
    Msf_GriAseg.Text = "cod_nacionalidad"
    
    Msf_GriAseg.Col = 30
    Msf_GriAseg.ColWidth(30) = 1000
    Msf_GriAseg.Text = "Nacionalidad"
    '-- End    :  Modify by : ricardo.huerta
    
'´RRR 26122013

'Inicio GCP 22032019
    Msf_GriAseg.Col = 31
    Msf_GriAseg.ColWidth(31) = 0
    Msf_GriAseg.Text = "numCCI"
  
    Msf_GriAseg.Col = 32
    Msf_GriAseg.ColWidth(32) = 0
    Msf_GriAseg.Text = "telefono1"
  
    Msf_GriAseg.Col = 33
    Msf_GriAseg.ColWidth(33) = 0
    Msf_GriAseg.Text = "telefono2"
    
    Msf_GriAseg.Col = 34
    Msf_GriAseg.ColWidth(34) = 0
    Msf_GriAseg.Text = "chkConTratDatos_Afil"
    
    Msf_GriAseg.Col = 35
    Msf_GriAseg.ColWidth(35) = 0
    Msf_GriAseg.Text = "chkConUsoDatosCom_Afil"
'END GCP 22032019
 
 'Inicio GCP 16092021
 'Se añade campo para Correo
    Msf_GriAseg.Col = 36
    Msf_GriAseg.ColWidth(36) = 1000
    Msf_GriAseg.Text = "Email"

 'Fin GCP 16092021


Exit Function
Err_IniGri:
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
''    Msf_Grilla.ColAlignment(0) = 1  'centrado
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

Function flLimpiarDatosAfi()
On Error GoTo Err_Limpiar

'I--- ABV 07/08/2007
'    If SSTab_Poliza.TabEnabled(2) = False Then
'        'Lbl_Cuspp = ""
'        If (Cmb_TipoIdent.ListCount <> 0) Then
'            Cmb_TipoIdent.ListIndex = 0
'        End If
'        Txt_NumIdent = ""
'    End If
    If (Cmb_TipoIdent.ListCount <> 0) Then
        Cmb_TipoIdent.ListIndex = 0
    End If
'I--- ABV 07/08/2007
    
    Txt_Asegurados = ""
    Txt_NomAfi = ""
    Txt_NomAfiSeg = ""
    Txt_ApPatAfi = ""
    Txt_ApMatAfi = ""
    Txt_FecInv = ""
    'Txt_Dir = ""
    Txt_NombreVia = ""
    Txt_Numero = ""
    Txt_Interior = ""
    Txt_NombreZona = ""
    Txt_Referencia = ""
    
    Txt_Fono = ""
    'Integracion GobiernoDeDatos()_
    Txt_Fono2_Afil = ""
    'Integracion GobiernoDeDatos()__
    Txt_Correo = ""
    Txt_Nacionalidad = ""
    Txt_NumCta = ""
    txt_CCI = ""
    
    Cmb_ViaPago.Clear
     
    Call fgComboGeneral(vgCodTabla_ViaPago, Cmb_ViaPago)
        
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
    If (Cmb_TipoVia.ListCount <> 0) Then
        Cmb_TipoVia.ListIndex = 0
    End If
    If (Cmb_TipoZona.ListCount <> 0) Then
        Cmb_TipoZona.ListIndex = 0
    End If
    
    Lbl_SexoAfi = ""
    Lbl_FecNac = ""
    Lbl_FecFall = ""
    Lbl_TipPen = ""
    Cmd_Representante.Enabled = False
    Lbl_CauInv = ""
    Lbl_Afp = ""
    Lbl_Departamento = ""
    Lbl_Provincia = ""
    Lbl_Distrito = ""
    Lbl_Dir = ""
    
Exit Function
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLimpiarDatosCal()
On Error GoTo Err_LimpiaCal
    
    Lbl_FecDev = ""
    Lbl_FecIncorpora = ""
    Txt_FecIniPago = ""
    Lbl_CUSPP = ""
    Lbl_FecCalculo = ""
    
    Lbl_TipoRenta = ""
    Lbl_AnnosDif = ""
    Lbl_Alter = ""
    Lbl_MesesGar = ""
    Lbl_RentaAFP = ""
    Lbl_PrcRentaImp = ""
    Lbl_IndCob = ""
    Lbl_CobConyuge = ""
    Lbl_DerCre = ""
    Lbl_DerGra = ""
    
    Lbl_ComInt = ""
    Lbl_BenSocial = ""
    Lbl_ComIntBen = ""
    Lbl_TipoIdentCorr = ""
    Lbl_NumIdentCorr = ""
    
    Lbl_MtoPrimaUniSim = ""
    Lbl_MtoPrimaUniDif = ""
    Lbl_TasaCtoEq = ""
    Lbl_TasaVta = ""
    Lbl_TasaTIR = ""
    Lbl_TasaPerGar = ""
    Lbl_MtoPension = ""
    Lbl_MtoPensionGar = ""
    
    Lbl_CtaInd = ""
    Lbl_BonoAct = ""
    Lbl_ApoAdi = ""
    Lbl_PriUnica = ""

    Lbl_Moneda(clMonedaModalidad) = ""
    Lbl_Moneda(clMonedaModPen) = ""
    Lbl_Moneda(clMonedaModPenGar) = ""
    Lbl_MonedaFon(clMtoCtaIndFon) = ""
    Lbl_MonedaFon(clMtoBonoFon) = ""
    Lbl_MonedaFon(clMtoPriUniFon) = ""
    
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

Function flLimpiarDatosBono()
On Error GoTo Err_LimpiaBono

    Lbl_TipoBono = ""
    Lbl_ValorNomBono = ""
    Lbl_FecEmiBono = ""
    Lbl_FecVenBono = ""
    Lbl_PrcIntBono = ""
    Lbl_ValUFBono = ""
    Lbl_ValPesosBono = ""
    Lbl_EdadCobroBono = ""

Exit Function
Err_LimpiaBono:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flLimpiarDatosAseg()
On Error GoTo Err_LimpiaAseg

    If (Cmb_TipoIdentBen.ListCount <> 0) Then
        Cmb_TipoIdentBen.ListIndex = 0
        Txt_NumIdentBen = "0"
    End If
    
'    Txt_NumIdentBen = ""
    Txt_NombresBen = ""
    Txt_NombresBenSeg = ""
    Txt_ApPatBen = ""
    Txt_ApMatBen = ""
    Me.txtCorreoBen.Text = ""
    Txt_FecInvBen = ""
    Lbl_NumOrden = ""
    txtTutor.Text = ""
'    Fra_DatosBenef.Enabled = True
'    Txt_RutBen.Enabled = True
'    Txt_DgvBen.Enabled = True
    
    Lbl_FecNacBen = ""
    Lbl_FecFallBen = ""
    Txt_FecFallBen.Text = ""    'DC 20091125
    Lbl_Porcentaje = ""
    Lbl_PensionBen = ""
    Lbl_PenGar = ""
    Lbl_Par = ""
    Lbl_Grupo = ""
    Lbl_SexoBen = ""
    Lbl_SitInvBen = ""
    Lbl_CauInvBen = ""
    Lbl_DerPension = ""
    
    Txt_Fono1_Ben = ""
    Txt_Fono2_Ben = ""
    chkConTratDatos_Ben = "0"
    chkConUsoDatosCom_Ben = "0"
    
    
    
    Lbl_Moneda(clMonedaBenPen) = ""
    Lbl_Moneda(clMonedaBenPenGar) = ""
    
Exit Function
Err_LimpiaAseg:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrillaBono(iNumPoliza As String)
On Error GoTo Err_flCargaGrillaBono
    
    vgSql = ""
    vgSql = "SELECT  b.cod_tipobono,b.mto_valnom,b.fec_emi,b.fec_ven, "
    vgSql = vgSql & "b.prc_tasaint,b.mto_bonoactuf,b.mto_bonoact, "
    vgSql = vgSql & "b.num_edadcob "
    vgSql = vgSql & "FROM pd_tmae_oripolbon b "
    vgSql = vgSql & "WHERE b.num_poliza = '" & Trim(iNumPoliza) & "' "
    vgSql = vgSql & "ORDER BY cod_tipobono "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       '''Call flInicializaGrillaBono
       
       While Not vgRs.EOF
       
          vltipobono = Trim(vgRs!cod_tipobono)  '& " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipBono, Trim(vgRs!cod_tipobono)))
                 
          Msf_GrillaBono.AddItem vltipobono & vbTab _
          & (Format((vgRs!mto_valnom), "###,###,##0.00")) & vbTab _
          & (DateSerial(Mid((vgRs!fec_emi), 1, 4), Mid((vgRs!fec_emi), 5, 2), Mid((vgRs!fec_emi), 7, 2))) & vbTab _
          & (DateSerial(Mid((vgRs!fec_ven), 1, 4), Mid((vgRs!fec_ven), 5, 2), Mid((vgRs!fec_ven), 7, 2))) & vbTab _
          & (Format((vgRs!prc_tasaint), "###,###,##0.00")) & vbTab _
          & (Format((vgRs!mto_bonoactuf), "###,###,##0.00")) & vbTab _
          & (Format((vgRs!mto_bonoact), "###,###,##0.00")) & vbTab _
          & (Trim(vgRs!num_edadcob))
          
          vgRs.MoveNext
       Wend
    End If
    vgRs.Close

Exit Function
Err_flCargaGrillaBono:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrillaBonoCot(iNumCot As String)

On Error GoTo Err_flCargaGrillaBonoCot
    
    vgSql = ""
    vgSql = "SELECT  b.cod_tipobono,b.mto_valnom,b.fec_emi,b.fec_ven, "
    vgSql = vgSql & "b.prc_tasaint,b.mto_bonoactuf,b.mto_bonoact, "
    vgSql = vgSql & "b.num_edadcob "
    vgSql = vgSql & "FROM pt_tmae_cotbono b "
    vgSql = vgSql & "WHERE b.num_cot = '" & Trim(iNumCot) & "' "
    vgSql = vgSql & "ORDER BY cod_tipobono "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       'Call flInicializaGrillaBono
       
       While Not vgRs.EOF
       
          vltipobono = Trim(vgRs!cod_tipobono)  '& " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipBono, Trim(vgRs!cod_tipobono)))
                 
          Msf_GrillaBono.AddItem vltipobono & vbTab _
          & (Format((vgRs!mto_valnom), "###,###,##0.00")) & vbTab _
          & (DateSerial(Mid((vgRs!fec_emi), 1, 4), Mid((vgRs!fec_emi), 5, 2), Mid((vgRs!fec_emi), 7, 2))) & vbTab _
          & (DateSerial(Mid((vgRs!fec_ven), 1, 4), Mid((vgRs!fec_ven), 5, 2), Mid((vgRs!fec_ven), 7, 2))) & vbTab _
          & (Format((vgRs!prc_tasaint), "###,###,##0.00")) & vbTab _
          & (Format((vgRs!mto_bonoactuf), "###,###,##0.00")) & vbTab _
          & (Format((vgRs!mto_bonoact), "###,###,##0.00")) & vbTab _
          & (Trim(vgRs!num_edadcob))
          
          vgRs.MoveNext
       Wend
    End If
    vgRs.Close

Exit Function
Err_flCargaGrillaBonoCot:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'------------------------------- POLIZA ----------------------------------
'-------------------------------------------------------------------------
'FUNCION QUE BUSCA LA POLIZA Y LLENA LOS DATOS
'-------------------------------------------------------------------------
Function flBuscaPoliza(iNumPol As String)
On Error GoTo Err_BuscaPol
Dim pasoErr As String
pasoErr = 0


    SSTab_Poliza.Tab = 0
    
    vlSql = "SELECT p.num_poliza, b.gls_nomben, b.gls_patben, b.gls_matben "
    vlSql = vlSql & "FROM pd_tmae_oripoliza p,pd_tmae_oripolben b WHERE "
    vlSql = vlSql & "p.num_poliza = b.num_poliza "
    vlSql = vlSql & "AND b.cod_par = '99' "
    vlSql = vlSql & "AND p.num_poliza = '" & iNumPol & "'"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If IsNull(vgRs.EOF) Then
        MsgBox "La Póliza no fue encontrada", vbCritical, "¡Error...!"
        Exit Function
    End If
    vgRs.Close
    
    pasoErr = 1

    Cmd_CrearPol.Enabled = False
'    Cmd_Nuevo.Enabled = False
    Cmd_Grabar.Enabled = True
    cmdEnviaCorreo.Enabled = True
    Cmd_Eliminar.Enabled = True
    
    If (vgNivelIndicadorBoton = "S") Then
        Cmd_Editar.Enabled = True
        Cmd_Eliminar.Enabled = True
    Else
        Cmd_Editar.Enabled = False
        Cmd_Eliminar.Enabled = False
    End If

pasoErr = 2
    'llenar los objetos de texto y los combos con la póliza escogida
'    Msf_GriAfiliado.Col = 0
'    Msf_GriAfiliado.Row = 1
'    vlNumero = Msf_GriAfiliado.Text
    vlNumero = iNumPol

    'Integracion GobiernoDeDatos()_
     Call LimpiarVariables
    'Fin Integracion GobiernoDeDatos_
    
    pasoErr = 3
    flCargaCarpAfilPol (vlNumero)
    pasoErr = 4
    flCargaCarpCalculo (vlNumero)
    pasoErr = 5
    'flCargaCarpBono (vlNumero)
    flCargaCarpBenef (vlNumero)
    pasoErr = 6
    pCargaCarpRep (vlNumero)

    Fra_Cabeza.Enabled = True
    Fra_Afiliado.Enabled = True
    Fra_PagPension.Enabled = True
    Fra_DatCal.Enabled = True
    Fra_SumaBono.Enabled = True
'    Fra_DatosBenef.Enabled = True
    Msf_GriAseg.Enabled = True
    
    
    pasoErr = 7
    'Call flValidaViaPago
    'Lbl_Cuspp.Enabled = True
    'Txt_Digito.Enabled = True
    SSTab_Poliza.Enabled = True
    SSTab_Poliza.Tab = 0
    Txt_FecVig.Enabled = True
'    Msf_GriAfiliado.Enabled = False
    
    Cmd_Poliza.Enabled = False
    Cmd_Cotizacion.Enabled = False
    
'    Txt_FecVig.SetFocus
    
Exit Function
Err_BuscaPol:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave #" & pasoErr & "[ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'------------------------------- POLIZA ----------------------------------
'-------------------------------------------------------------------
'CARGA LOS CAMPOS DE LA CARPETA AFILIADO
'-------------------------------------------------------------------
Function flCargaCarpAfilPol(iNumPol As String)
On Error GoTo Err_CargaAfi
    
    Call flLimpiarDatosAfi
    
    vlSql = "SELECT p.num_poliza,p.num_cot,p.fec_vigencia,"
    vlSql = vlSql & "p.cod_tipoidenafi as cod_tipoiden,p.num_idenafi as num_iden,p.num_cargas, "
    vlSql = vlSql & "p.num_operacion,p.num_correlativo as cod_secofe, "
    vlSql = vlSql & "b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben, "
    vlSql = vlSql & "b.cod_sexo,b.fec_nacben,b.fec_fallben, "
    vlSql = vlSql & "p.cod_tippension,b.fec_invben,b.cod_cauinv,"
    vlSql = vlSql & "p.cod_estcivil,p.cod_afp,p.cod_isapre,"
    vlSql = vlSql & "p.gls_direccion,p.cod_direccion,p.gls_fono,"
    vlSql = vlSql & "p.gls_correo,p.cod_viapago,p.cod_sucursal,"
    vlSql = vlSql & "p.cod_tipcuenta,p.cod_banco,p.num_cuenta,p.num_cuenta_cci, "
    vlSql = vlSql & "p.gls_nacionalidad,p.cod_vejez,p.Fec_Emision, "
    vlSql = vlSql & "p.cod_tipvia,p.gls_nomvia,p.gls_numdmc,p.gls_intdmc, "
    vlSql = vlSql & "p.cod_tipzon,p.gls_nomzon,p.gls_referencia, p.cod_nacionalidad, "
    'GCP-FRACTAL 11042019
    vlSql = vlSql & "p.GLS_FONO, b.GLS_FONO2, nvl(b.CONS_TRAINFO,'0') as CONS_TRAINFO, nvl(b.CONS_DATCOMER,'0') as CONS_DATCOMER, p.cod_moncta "
    vlSql = vlSql & "FROM pd_tmae_oripoliza p, pd_tmae_oripolben b "
    vlSql = vlSql & "WHERE p.num_poliza = '" & iNumPol & "' "
    vlSql = vlSql & "AND p.num_poliza = b.num_poliza "
    vlSql = vlSql & "AND b.cod_par = '99' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    
    If Not (vgRs.EOF) Then
        'datos de la cabecera
        If Not IsNull(vgRs!Num_Cot) Then Lbl_NumCot = vgRs!Num_Cot
        If Not IsNull(vgRs!Num_Poliza) Then Txt_NumPol = vgRs!Num_Poliza
'        If Not IsNull(vgRs!Fec_Vigencia) Then Txt_FecVig = DateSerial(Mid(vgRs!Fec_Vigencia, 1, 4), Mid(vgRs!Fec_Vigencia, 5, 2), Mid(vgRs!Fec_Vigencia, 7, 2))
        If Not IsNull(vgRs!Fec_Emision) Then Txt_FecVig = DateSerial(Mid(vgRs!Fec_Emision, 1, 4), Mid(vgRs!Fec_Emision, 5, 2), Mid(vgRs!Fec_Emision, 7, 2))
        If Not IsNull(vgRs!Num_Operacion) Then Lbl_SolOfe = vgRs!Num_Operacion
        If Not IsNull(vgRs!cod_secofe) Then Lbl_SecOfe = vgRs!cod_secofe

        'datos de la carpeta de afiliado
        If Not IsNull(vgRs!NUM_IDEN) Then Txt_NumIdent = vgRs!NUM_IDEN
        If Not IsNull(vgRs!Num_Cargas) Then Txt_Asegurados = vgRs!Num_Cargas
        If Not IsNull(vgRs!Gls_NomBen) Then Txt_NomAfi = vgRs!Gls_NomBen
        If Not IsNull(vgRs!Gls_NomSegBen) Then Txt_NomAfiSeg = vgRs!Gls_NomSegBen
        If Not IsNull(vgRs!Gls_PatBen) Then Txt_ApPatAfi = vgRs!Gls_PatBen
        If Not IsNull(vgRs!Gls_MatBen) Then Txt_ApMatAfi = vgRs!Gls_MatBen
        If Not IsNull(vgRs!Fec_NacBen) Then Lbl_FecNac = DateSerial(Mid(vgRs!Fec_NacBen, 1, 4), Mid(vgRs!Fec_NacBen, 5, 2), Mid(vgRs!Fec_NacBen, 7, 2))
        If Not IsNull(vgRs!Fec_FallBen) Then Lbl_FecFall = DateSerial(Mid(vgRs!Fec_FallBen, 1, 4), Mid(vgRs!Fec_FallBen, 5, 2), Mid(vgRs!Fec_FallBen, 7, 2))
        If Not IsNull(vgRs!Fec_InvBen) Then Lbl_FecInv = DateSerial(Mid(vgRs!Fec_InvBen, 1, 4), Mid(vgRs!Fec_InvBen, 5, 2), Mid(vgRs!Fec_InvBen, 7, 2))
        
        If Not IsNull(vgRs!GLS_FONO) Then Txt_Fono = vgRs!GLS_FONO
        If Not IsNull(vgRs!GLS_CORREO) Then Txt_Correo = vgRs!GLS_CORREO
        
        If Not IsNull(vgRs!GLS_FONO) Then Txt_Fono = vgRs!GLS_FONO
        If Not IsNull(vgRs!gls_fono2) Then Txt_Fono2_Afil = vgRs!gls_fono2
        If Not IsNull(vgRs!CONS_TRAINFO) Then chkConTratDatos_Afil = vgRs!CONS_TRAINFO
        If Not IsNull(vgRs!CONS_DATCOMER) Then chkConUsoDatosCom_Afil = vgRs!CONS_DATCOMER
        
        Dim arTiPension(4) As String
        Dim SwBbloqueaChk As Boolean
        
        SwBbloqueaChk = False
        
        arTiPension(0) = "08"
        arTiPension(1) = "09"
        arTiPension(2) = "10"
        arTiPension(3) = "11"
        arTiPension(4) = "12"
        
        For i = 0 To 4
            If arTiPension(i) = Trim(vgRs!Cod_TipPension) Then
                    SwBbloqueaChk = True
                    Exit For
            End If
        Next
        
        If SwBbloqueaChk Then
           chkConTratDatos_Afil.Enabled = False
            chkConUsoDatosCom_Afil.Enabled = False
       
        Else
            chkConTratDatos_Afil.Enabled = True
            chkConUsoDatosCom_Afil.Enabled = True
            
       End If
        
       'FIN GCP-FRACTAL 11042019
   

        If Not IsNull(vgRs!Cod_Sexo) Then Lbl_SexoAfi = Trim(vgRs!Cod_Sexo) & " - " & fgBuscarGlosaElemento(vgCodTabla_Sexo, vgRs!Cod_Sexo)
        If Not IsNull(vgRs!Cod_TipPension) Then Lbl_TipPen = Trim(vgRs!Cod_TipPension) & " - " & fgBuscarGlosaElemento(vgCodTabla_TipPen, vgRs!Cod_TipPension)
        'RVF 20090914
        If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
            Cmd_Representante.Enabled = True
            'DC 20091125
            Txt_FecFallBen.Visible = True
            Lbl_FecFallBen.Visible = False
        Else
            Cmd_Representante.Enabled = False
            'DC 20091125
            Txt_FecFallBen.Visible = False
            Lbl_FecFallBen.Visible = True
        End If
        '*****
        
       
             
        If Not IsNull(vgRs!Cod_CauInv) Then Lbl_CauInv = Trim(vgRs!Cod_CauInv) & " - " & fgBuscarGlosaCauInv(vgRs!Cod_CauInv)
        If Not IsNull(vgRs!Cod_AFP) Then Lbl_Afp = Trim(vgRs!Cod_AFP) & " - " & fgBuscarGlosaElemento(vgCodTabla_AFP, vgRs!Cod_AFP)

        If Not IsNull(vgRs!cod_tipoiden) Then Call fgBuscaPos(Cmb_TipoIdent, vgRs!cod_tipoiden)
        If Not IsNull(vgRs!cod_estcivil) Then Call fgBuscaPos(Cmb_EstCivil, vgRs!cod_estcivil)
        If Not IsNull(vgRs!cod_vejez) Then Call fgBuscaPos(Cmb_Vejez, vgRs!cod_vejez)
        
        Call fgBuscaPos(Cmb_Salud, (vgRs!cod_isapre))
        vlSwViaPago = False
        
        
       'INICIO GCP-FRACTAL [REQUERIMIENTO SD-5041] 21032019
        Dim TP As String
        TP = Left(Trim(Lbl_TipPen.Caption), 2)
        If (TP = "04" Or TP = "05") Then
          'Se Remueve el via pago TRANSFERENCIA AFP
          Call RemoveItemCombo(Cmb_ViaPago, "04")

       Else
        'Se Remueve el via pago DEPOSITO EN CUENTA
          Call RemoveItemCombo(Cmb_ViaPago, "02")
       End If
       'FIN GCP-FRACTAL [REQUERIMIENTO SD-5041] 21032019
        vlSwViaPago = True
        Call fgBuscaPos(Cmb_ViaPago, (vgRs!Cod_ViaPago))
        vlSwViaPago = False
        vlSw = False
        Call flValidaViaPago
        fgComboSucursal Cmb_Suc, "S"
        
        Call fgBuscaPos(Cmb_Suc, (vgRs!Cod_Sucursal))
        Call fgBuscaPos(Cmb_TipCta, (vgRs!Cod_TipCuenta))
        Call fgBuscaPos(Cmb_Bco, (vgRs!Cod_Banco))
        Call fgBuscaPos(cmb_MonCta, (vgRs!COD_MONCTA))
        If Not IsNull(vgRs!Num_Cuenta) Then Txt_NumCta = vgRs!Num_Cuenta
        If Not IsNull(vgRs!num_cuenta_cci) Then txt_CCI = vgRs!num_cuenta_cci
        
        vlSwViaPago = False
        If Not IsNull(vgRs!Cod_Direccion) Then
            vlCodDireccion = vgRs!Cod_Direccion
            vgCodDireccion = vgRs!Cod_Direccion
            'Integracion GobiernoDeDatos_
            pgCodDireccion = vgRs!Cod_Direccion
            'Fin Integracion GobiernoDeDatos_
            Call fgBuscarNombreComunaProvinciaRegion(vlCodDireccion)
            Lbl_Departamento = vgNombreRegion
            Lbl_Provincia = vgNombreProvincia
            Lbl_Distrito = vgNombreComuna
        End If
    
        'RVF 20090914
        If Not IsNull(vgRs!cod_tipvia) Then
            Call fgBuscaPos(Cmb_TipoVia, (vgRs!cod_tipvia))
        Else
            Cmb_TipoVia.ListIndex = -1
        End If
        If Not IsNull(vgRs!gls_nomvia) Then Txt_NombreVia.Text = vgRs!gls_nomvia
        If Not IsNull(vgRs!gls_numdmc) Then Txt_Numero.Text = vgRs!gls_numdmc
        If Not IsNull(vgRs!gls_intdmc) Then Txt_Interior.Text = vgRs!gls_intdmc
        If Not IsNull(vgRs!cod_tipzon) Then
            Call fgBuscaPos(Cmb_TipoZona, (vgRs!cod_tipzon))
        Else
            Cmb_TipoZona.ListIndex = -1
        End If
        If Not IsNull(vgRs!gls_nomzon) Then Txt_NombreZona.Text = vgRs!gls_nomzon
        If Not IsNull(vgRs!gls_referencia) Then Txt_Referencia.Text = vgRs!gls_referencia
                 
         '-- Begin : Modify by : ricardo.huerta YEAH
            If Not IsNull(vgRs!cod_nacionalidad) Then
                Call fgBuscaPos(cboNacionalidad, vgRs!cod_nacionalidad)
                vlNacionalidad = fg_obtener_descripcion_nacionalidad(vgRs!cod_nacionalidad)
                Txt_Nacionalidad = vlNacionalidad
            End If
            
         '-- End    :  Modify by : ricardo.huerta

        Call pConcatenaDireccion
        '*****
        
        'Integracion GobiernoDeDatos(metodo que llena las variables de direccion y telefono)
        Call ObtenerDIrecTelef(iNumPol)
        'Fin Integracion GobiernoDeDatos
    End If

Exit Function
Err_CargaAfi:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
 'INICIO GCP-FRACTAL 10052019
 Sub RemoveItemCombo(ByRef gCombo As ComboBox, ByVal CodItem As String)
        Dim i As Integer
        Dim vPosicion As Integer
        Dim vCodigo As String
        'vlSwViaPago = True
        For i = gCombo.ListCount - 1 To 0 Step -1
            gCombo.ListIndex = i
            
            vPosicion = InStr(1, gCombo.Text, "-")
            vCodigo = Trim(Mid(gCombo.Text, 1, vPosicion - 1))
           
           If vCodigo = CodItem Then
              gCombo.RemoveItem i
           End If
           
        Next
        
 End Sub
 Sub RemoveItemCombo2(ByRef gCombo As ComboBox, ByVal CodItem As String)
        Dim i As Integer
        Dim vPosicion As Integer
        Dim vCodigo As String

        For i = gCombo.ListCount - 1 To 0 Step -1
            gCombo.ListIndex = i
            
            vPosicion = InStr(1, gCombo.Text, "-")
            vCodigo = Trim(Mid(gCombo.Text, 1, vPosicion - 1))
           
           If vCodigo <> CodItem Then
              gCombo.RemoveItem i
           End If
           
        Next

 End Sub
 'FIN GCP-FRACTAL 10052019
  Sub DejarBcoPrinicipales(ByRef gCombo As ComboBox)
        'Dejar en el combo solo estos Bancos
        '02  BANCO DE CREDITO DEL PERU
        '11' BANCO CONTINENTAL
        '41' SCOTIABANK PERU
        '03' INTERBANK
        
        Dim Bancos(5) As String
        Dim SwBorra As Boolean
        
    
        
        
        Bancos(0) = "02"
        Bancos(1) = "03"
        Bancos(2) = "11"
        Bancos(3) = "41"
        Bancos(4) = "00"
        
        
        For i = gCombo.ListCount - 1 To 0 Step -1
            gCombo.ListIndex = i
             SwBorra = True
            
            vPosicion = InStr(1, gCombo.Text, "-")
            vCodigo = Trim(Mid(gCombo.Text, 1, vPosicion - 1))
            
            For j = 0 To UBound(Bancos)
               
                If vCodigo = Bancos(j) Then
                        SwBorra = False
                End If
            Next
            
            If SwBorra Then
              gCombo.RemoveItem i
            End If
            
                  
        Next
  
 End Sub
 
 
 

'------------------------------- POLIZA ----------------------------------
'--------------------------------------------------------------------
'CARGA LOS CAMPOS DE LA CARPETA DATOS DE CALCULO
'--------------------------------------------------------------------
Function flCargaCarpCalculo(iNumPol As String)
On Error GoTo Err_CarCalculo
    
    Call flLimpiarDatosCal

    vlSql = "SELECT cod_tipren,num_mesdif,cod_modalidad,cod_dercre,cod_dergra,"
    vlSql = vlSql & "num_mesgar,prc_rentaafp,prc_rentatmp,cod_cuspp,ind_cob,"
    vlSql = vlSql & "cod_cobercon,mto_facpenella,prc_facpenella,cod_tippension,"
    vlSql = vlSql & "prc_corcom,prc_corcomreal,cod_bensocial,cod_tipoidencor,num_idencor,"
    vlSql = vlSql & "fec_dev,p.cod_moneda,fec_acepta,mto_cnu,prc_tasace as prc_tasatce,"
    vlSql = vlSql & "prc_tasavta,prc_tasatir,prc_tasapergar,"
    vlSql = vlSql & "mto_pension,mto_pensiongar,mto_priuni,'" & vgMonedaCodOfi & "' as cod_monedafon,"
    vlSql = vlSql & "mto_ctaind,"
    vlSql = vlSql & "mto_bono,mto_priunisim,mto_priunidif,fec_pripago "
    vlSql = vlSql & ",fec_calculo,mto_apoadi "
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & ",p.cod_tipreajuste,p.mto_valreajustetri,p.mto_valreajustemen "
    vlSql = vlSql & ",mtr.cod_scomp as cod_montipreaju,mtr.gls_descripcion as gls_montipreaju, cod_nacionalidad"
'F--- ABV 05/02/2011 ---
    vlSql = vlSql & " FROM pd_tmae_oripoliza p "
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & ",ma_tpar_monedatiporeaju mtr "
'F--- ABV 05/02/2011 ---
    vlSql = vlSql & "WHERE num_poliza = '" & iNumPol & "'"
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & "AND p.cod_tipreajuste = mtr.cod_tipreajuste(+) AND "
    vlSql = vlSql & "p.cod_moneda = mtr.cod_moneda(+) "
'F--- ABV 05/02/2011 ---
    Set vgRs = vgConexionBD.Execute(vlSql)

    If Not (vgRs.EOF) Then
        
        Lbl_FecDev = DateSerial(Mid(vgRs!Fec_Dev, 1, 4), Mid(vgRs!Fec_Dev, 5, 2), Mid(vgRs!Fec_Dev, 7, 2))
        Lbl_FecIncorpora = DateSerial(Mid(vgRs!Fec_Acepta, 1, 4), Mid(vgRs!Fec_Acepta, 5, 2), Mid(vgRs!Fec_Acepta, 7, 2))
        Txt_FecIniPago = DateSerial(Mid(vgRs!fec_pripago, 1, 4), Mid(vgRs!fec_pripago, 5, 2), Mid(vgRs!fec_pripago, 7, 2))
        Lbl_CUSPP = vgRs!Cod_Cuspp
        Lbl_FecCalculo = DateSerial(Mid(vgRs!Fec_Calculo, 1, 4), Mid(vgRs!Fec_Calculo, 5, 2), Mid(vgRs!Fec_Calculo, 7, 2))
                
        If Not IsNull(vgRs!Num_MesDif) Then Lbl_AnnosDif = Trim(vgRs!Num_MesDif / 12)
        If Not IsNull(vgRs!Num_MesGar) Then Lbl_MesesGar = vgRs!Num_MesGar
        If Not IsNull(vgRs!Prc_RentaAFP) Then Lbl_RentaAFP = Format(vgRs!Prc_RentaAFP, "#0.00")
        If Not IsNull(vgRs!Prc_RentaTMP) Then Lbl_PrcRentaTmp = Format(vgRs!Prc_RentaTMP, "#0.00")
        If Not IsNull(vgRs!Cod_CoberCon) Then Lbl_FacPenElla = vgRs!Cod_CoberCon
        If Not IsNull(vgRs!Prc_CorCom) Then Lbl_ComIntBen = Format(vgRs!Prc_CorCom, "#0.00")
        If Not IsNull(vgRs!Prc_CorComReal) Then Lbl_ComInt = Format(vgRs!Prc_CorComReal, "#0.00")
                
        'Buscar la Identificación del Intermediario
        If Not IsNull(vgRs!cod_tipoidencor) Then
            Lbl_TipoIdentCorr = vgRs!cod_tipoidencor & " - " & fgBuscarNombreTipoIden(vgRs!cod_tipoidencor)
        End If
        If Not IsNull(vgRs!Num_IdenCor) Then
            If Not IsNull(vgRs!Num_IdenCor) Then Lbl_NumIdentCorr = vgRs!Num_IdenCor
        End If

        If Not IsNull(vgRs!Mto_CNU) Then Txt_PrcFam = Format(vgRs!Mto_CNU, "#,#0.000000")
        If Not IsNull(vgRs!Mto_PriUniSim) Then Lbl_MtoPrimaUniSim = Format(vgRs!Mto_PriUniSim, "#,#0.00")
        If Not IsNull(vgRs!Mto_PriUniDif) Then Lbl_MtoPrimaUniDif = Format(vgRs!Mto_PriUniDif, "#,#0.00")
        If Not IsNull(vgRs!prc_tasatce) Then Lbl_TasaCtoEq = Format(vgRs!prc_tasatce, "#,#0.00")
        If Not IsNull(vgRs!Prc_TasaVta) Then Lbl_TasaVta = Format(vgRs!Prc_TasaVta, "#0.00")
        If Not IsNull(vgRs!Prc_TasaTir) Then Lbl_TasaTIR = Format(vgRs!Prc_TasaTir, "#0.00")
        If Not IsNull(vgRs!prc_tasapergar) Then Lbl_TasaPerGar = Format(vgRs!prc_tasapergar, "#0.00")
        If Not IsNull(vgRs!Mto_Pension) Then Lbl_MtoPension = Format(vgRs!Mto_Pension, "#,#0.00")
        If Not IsNull(vgRs!Mto_PensionGar) Then Lbl_MtoPensionGar = Format(vgRs!Mto_PensionGar, "#,#0.00")
        If Not IsNull(vgRs!Mto_CtaInd) Then Lbl_CtaInd = Format(vgRs!Mto_CtaInd, "#,#0.00")
        If Not IsNull(vgRs!Mto_Bono) Then Lbl_BonoAct = Format(vgRs!Mto_Bono, "#,#0.00")
        If Not IsNull(vgRs!Mto_ApoAdi) Then Lbl_ApoAdi = Format(vgRs!Mto_ApoAdi, "#,#0.00")
        If Not IsNull(vgRs!MTO_PRIUNI) Then Lbl_PriUnica = Format(vgRs!MTO_PRIUNI, "#,#0.00")
        
        If Not IsNull(vgRs!Cod_TipRen) Then Lbl_TipoRenta = Trim(vgRs!Cod_TipRen) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipRen, Trim(vgRs!Cod_TipRen)))
        If Not IsNull(vgRs!Cod_Modalidad) Then Lbl_Alter = Trim(vgRs!Cod_Modalidad) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_AltPen, Trim(vgRs!Cod_Modalidad)))
        
'Nuevos Campos
        If Not IsNull(vgRs!Ind_Cob) Then
            If vgRs!Ind_Cob = "S" Then Lbl_IndCob = cgIndicadorSi Else Lbl_IndCob = cgIndicadorNo
        End If
        If Not IsNull(vgRs!Cod_DerCre) Then
            If vgRs!Cod_DerCre = "S" Then Lbl_DerCre = cgIndicadorSi Else Lbl_DerCre = cgIndicadorNo
        End If
        If Not IsNull(vgRs!Cod_DerGra) Then
            If vgRs!Cod_DerGra = "S" Then Lbl_DerGra = cgIndicadorSi Else Lbl_DerGra = cgIndicadorNo
        End If
        If Not IsNull(vgRs!Cod_BenSocial) Then
            If vgRs!Cod_BenSocial = "S" Then Lbl_BenSocial = cgIndicadorSi Else Lbl_BenSocial = cgIndicadorNo
        End If
        
        vlElemento = fgBuscarGlosaElemento(vgCodTabla_TipMon, vgRs!Cod_Moneda)
        Lbl_Moneda(0) = (vgRs!Cod_Moneda) + " - " + vlElemento

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
        
        'Tipo de Moneda de la Modalidad
        Lbl_Moneda(clMonedaBenPen) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_Moneda)
        Lbl_Moneda(clMonedaBenPenGar) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_Moneda)
        
        Lbl_Moneda(clMonedaModPen) = Lbl_Moneda(clMonedaBenPen)
        Lbl_Moneda(clMonedaModPenGar) = Lbl_Moneda(clMonedaBenPen)

        'Tipo de Moneda del Fondo
        Lbl_MonedaFon(clMtoCtaIndFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        Lbl_MonedaFon(clMtoBonoFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        Lbl_MonedaFon(clMtoPriUniFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        Lbl_MonedaFon(clMtoApoAdiFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        
'I--- ABV 10/08/2007 ---
        'Guardar en Variables los Datos para la corrección de Datos en Grilla de Beneficiarios
        vlFecDev = vgRs!Fec_Dev
        vlCodTipPen = Trim(vgRs!Cod_TipPension)
        vlCodCoberCon = Trim(vgRs!Cod_CoberCon)
        vlMtoFacPenElla = vgRs!Mto_FacPenElla
        vlPrcFacPenElla = vgRs!Prc_FacPenElla
        vlDerCrePol = Trim(vgRs!Cod_DerCre)
'F--- ABV 10/08/2007 ---


        '-- Begin : Modify by : ricardo.huerta
            If Not IsNull(vgRs!cod_nacionalidad) Then Call fgBuscaPos(cboNacionalidad, vgRs!cod_nacionalidad)
        '-- End    :  Modify by : ricardo.huerta
        
    End If

Exit Function
Err_CarCalculo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'------------------------------- POLIZA ----------------------------------
'-------------------------------------------------------------------------
'FUNCION CARGA DATOS DEL BONO
'-------------------------------------------------------------------------
Function flCargaCarpBono(iNumPol As String)
On Error GoTo Err_CargaBono

    Call flLimpiarDatosBono
    '''Call flInicializaGrillaBono - MC 29/05/2007
    Call flCargaGrillaBono(iNumPol)
    
Exit Function
Err_CargaBono:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Sub pCargaCarpRep(iNumPol As String)
Dim vlTipoIden As String
On Error GoTo Err_Cargarep

Call LimpiarDireccionRepresentante
    
    vlSql = "SELECT "
    vlSql = vlSql & "num_poliza,cod_tipoidenrep,num_idenrep,gls_nombresrep,gls_apepatrep,"
    vlSql = vlSql & "gls_apematrep,cod_usuariocrea,fec_crea,hor_crea, "
    vlSql = vlSql & "GLS_TELREP1,GLS_TELREP2, GLS_CORREOREP, cod_area_telrep1, cod_area_telrep2, cod_sexo "
    vlSql = vlSql & "FROM pd_tmae_oripolrep WHERE "
    vlSql = vlSql & "num_poliza = '" & iNumPol & "' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!cod_tipoidenRep) Then Call fgBuscaPos(Cmb_TipIdRep, vgRs!cod_tipoidenRep)
        If Not IsNull(vgRs!Num_idenrep) Then Txt_NumIdRep.Text = (vgRs!Num_idenrep) Else Txt_NumIdRep.Text = ""
        If Not IsNull(vgRs!Gls_NombresRep) Then Txt_NomRep.Text = (vgRs!Gls_NombresRep) Else Txt_NomRep.Text = ""
        If Not IsNull(vgRs!Gls_ApepatRep) Then Txt_ApPatRep.Text = (vgRs!Gls_ApepatRep) Else Txt_ApPatRep.Text = ""
        If Not IsNull(vgRs!Gls_ApematRep) Then Txt_ApMatRep.Text = (vgRs!Gls_ApematRep) Else Txt_ApMatRep.Text = ""
        
        If Not IsNull(vgRs!GLS_TELREP1) Then
            Me.txtTelRep1.Text = (vgRs!GLS_TELREP1)
            DirRep.vNumTelefono = (vgRs!GLS_TELREP1)
            DirRep.vCodigoTelefono = IIf(IsNull(vgRs!cod_area_telrep1), "", vgRs!cod_area_telrep1)
        
        Else
            Me.txtTelRep1.Text = ""
        End If
        
        If Not IsNull(vgRs!GLS_TELREP2) Then
            txtTelRep2.Text = (vgRs!GLS_TELREP2)
            DirRep.vNumTelefono2 = (vgRs!GLS_TELREP2)
            DirRep.vCodigoTelefono2 = IIf(IsNull(vgRs!cod_area_telrep2), "", vgRs!cod_area_telrep2)
            
        Else
            txtTelRep2.Text = ""
        End If
        
        Me.cmbSexoRep.Text = IIf(IsNull(vgRs!Cod_Sexo), "", vgRs!Cod_Sexo)
        
        If Not IsNull(vgRs!gls_correorep) Then Me.txtCorreoRep.Text = (vgRs!gls_correorep) Else txtCorreoRep.Text = ""
        
        Call pConcatenaRepresentante
    End If
    vgRs.Close
    
    
Call ObtieneDireccionRepresentante(iNumPol)

Exit Sub
Err_Cargarep:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub
'------------------------------- POLIZA ----------------------------------
'------------------------------------------------------------------------
'CARGA LA GRILLA CON LOS DATOS DE LOS BENEFICIARIOS
'------------------------------------------------------------------------
Function flCargaCarpBenef(iNumPol As String)
Dim vlTipoIden As String
On Error GoTo Err_CargaBen
    
    Msf_GriAseg.rows = 1
    
    vlSql = "SELECT "
    vlSql = vlSql & "num_orden,cod_grufam,cod_par,cod_sexo,cod_sitinv,"
    vlSql = vlSql & "cod_cauinv,cod_derpen,cod_dercre,"
    vlSql = vlSql & "cod_tipoidenben as cod_tipoiden,num_idenben as num_iden,"
    vlSql = vlSql & "gls_nomben,gls_nomsegben,gls_patben,gls_matben,"
    vlSql = vlSql & "fec_nacben,fec_invben,fec_fallben,Fec_NacHM,"
    vlSql = vlSql & "prc_pension,prc_pensionleg,mto_pension,mto_pensiongar "
    vlSql = vlSql & ",cod_estpension,prc_pensiongar,cod_tipcta, cod_monbco, cod_banco, num_ctabco, cod_nacionalidad, num_cuenta_cci"
    'GCP-FRACTAL 11042019
    vlSql = vlSql & ",GLS_FONO, GLS_FONO2, nvl(CONS_TRAINFO,'0') as CONS_TRAINFO, nvl(CONS_DATCOMER,'0') as CONS_DATCOMER, gls_correoben "
    vlSql = vlSql & " FROM pd_tmae_oripolben WHERE "
    vlSql = vlSql & "num_poliza = '" & iNumPol & "' "
    vlSql = vlSql & "ORDER BY num_orden ASC "
    Set vgRs = vgConexionBD.Execute(vlSql)
    While Not vgRs.EOF
    
        vlCodPar = Trim(vgRs!Cod_Par)  '& " - " & fgBuscarGlosaElemento(vgCodTabla_Par, vgRs!cod_par)
        vlGruFam = Trim(vgRs!Cod_GruFam)  '& " - " & fgBuscarGlosaElemento(vgCodTabla_GruFam, vgRs!cod_grufam)
        vlSexoBen = Trim(vgRs!Cod_Sexo)  '& " - " & fgBuscarGlosaElemento(vgCodTabla_Sexo, vgRs!cod_sexo)
        vlSitInv = Trim(vgRs!Cod_SitInv)  '& " - " & fgBuscarGlosaElemento(vgCodTabla_SitInv, vgRs!cod_sitinv)
        vlDerPen = Trim(vgRs!Cod_DerPen)  '& " - " & fgBuscarGlosaElemento(vgCodTabla_DerPen, vgRs!cod_derpen)
    
        If Not IsNull(vgRs!Cod_DerCre) Then vlDerCre = Trim(vgRs!Cod_DerCre) Else vlDerCre = ""
        If Not IsNull(vgRs!Cod_CauInv) Then vlCauInv = Trim(vgRs!Cod_CauInv)  '& " - " & fgBuscarGlosaCauInv(vgRs!cod_cauinv)
        If Not IsNull(vgRs!Fec_InvBen) Then
            vlFecInv = DateSerial(Mid(Trim(vgRs!Fec_InvBen), 1, 4), Mid(Trim(vgRs!Fec_InvBen), 5, 2), Mid(Trim(vgRs!Fec_InvBen), 7, 2))
        Else
            vlFecInv = ""
        End If
        
        If Not IsNull(vgRs!cod_tipoiden) Then
            vlRutGrilla = " " & vgRs!cod_tipoiden & " - " & fgBuscarNombreTipoIden(vgRs!cod_tipoiden)
        Else
            vlRutGrilla = ""
        End If
        
        If Not IsNull(vgRs!NUM_IDEN) Then
            vlDgvGrilla = Trim(vgRs!NUM_IDEN)
        Else
            vlDgvGrilla = ""
        End If
        
        If Not IsNull(vgRs!Gls_NomBen) Then vlNomBen = (vgRs!Gls_NomBen) Else vlNomBen = ""
        If Not IsNull(vgRs!Gls_NomSegBen) Then vlNomBenSeg = (vgRs!Gls_NomSegBen) Else vlNomBenSeg = ""
        If Not IsNull(vgRs!Gls_PatBen) Then vlPatBen = (vgRs!Gls_PatBen) Else vlPatBen = ""
        If Not IsNull(vgRs!Gls_MatBen) Then vlMatBen = (vgRs!Gls_MatBen) Else vlMatBen = ""
        
        'RRR
        If Not IsNull(vgRs!cod_tipcta) Then vlcod_tipcta = (vgRs!cod_tipcta) Else vlcod_tipcta = ""
        If Not IsNull(vgRs!cod_monbco) Then vlcod_monbco = (vgRs!cod_monbco) Else vlcod_monbco = ""
        If Not IsNull(vgRs!Cod_Banco) Then vlCod_Banco = (vgRs!Cod_Banco) Else vlCod_Banco = ""
        If Not IsNull(vgRs!num_ctabco) Then vlnum_ctabco = (vgRs!num_ctabco) Else vlnum_ctabco = ""
        
        If Not IsNull(vgRs!Fec_FallBen) Then
            vlFecFallBen = DateSerial(Mid(Trim(vgRs!Fec_FallBen), 1, 4), Mid(Trim(vgRs!Fec_FallBen), 5, 2), Mid(Trim(vgRs!Fec_FallBen), 7, 2))
        Else
            vlFecFallBen = ""
        End If
        If Not IsNull(vgRs!Fec_NacHM) Then
            vlFecNacHM = DateSerial(Mid(Trim(vgRs!Fec_NacHM), 1, 4), Mid(Trim(vgRs!Fec_NacHM), 5, 2), Mid(Trim(vgRs!Fec_NacHM), 7, 2))
        Else
            vlFecNacHM = ""
        End If
        
'I--- ABV 10/08/2007 ---
        'No es necesario recalcular los datos del D° Crecer, Est. Pensión y Montos de Pensión
        vlEstPen = Trim(vgRs!Cod_EstPension)
        vlPrcPension = vgRs!Prc_Pension
        vlPrcPensionLeg = vgRs!Prc_PensionLeg
        vlPenBen = vgRs!Mto_Pension
        vlPenGarBen = vgRs!Mto_PensionGar
        vlFecNacBen = vgRs!Fec_NacBen
        vlPrcPensionGar = vgRs!Prc_PensionGar
'F--- ABV 10/08/2007 ---

        '-- Begin : Modify by : ricardo.huerta
            If Not IsNull(vgRs!cod_nacionalidad) Then 'yeah
                vl_nacionalidadben = (vgRs!cod_nacionalidad)
                vl_nacionalidadben_descripcion = fg_obtener_descripcion_nacionalidad(vl_nacionalidadben)
            End If
            
            'INICIO GCP-FRACTAL 11042019
            vlnum_CCI = IIf(IsNull(vgRs!num_cuenta_cci), "", Trim(vgRs!num_cuenta_cci))
            vl_Fono1_Ben = IIf(IsNull(vgRs!GLS_FONO), "", Trim(vgRs!GLS_FONO))
            vl_Fono2_Ben = IIf(IsNull(vgRs!gls_fono2), "", Trim(vgRs!gls_fono2))
            vl_ConTratDatos_Ben = IIf(IsNull(vgRs!CONS_TRAINFO), "", Trim(vgRs!CONS_TRAINFO))
            vl_ConUsoDatosCom_Ben = IIf(IsNull(vgRs!CONS_DATCOMER), "", Trim(vgRs!CONS_DATCOMER))
            vl_Correo_Ben = IIf(IsNull(vgRs!Gls_CorreoBen), "", Trim(vgRs!Gls_CorreoBen))
            'FIN GCP-FRACTAL 11042019
            
        '-- End   : Modify by : ricardo.huerta
        
        Msf_GriAseg.AddItem Trim(vgRs!Num_Orden) & vbTab & _
                    Trim(vlCodPar) & vbTab & _
                    Trim(vlGruFam) & vbTab & _
                    Trim(vlSexoBen) & vbTab & _
                    Trim(vlSitInv) & vbTab & _
                    Trim(vlFecInv) & vbTab & _
                    Trim(vlCauInv) & vbTab & _
                    Trim(vlDerPen) & vbTab & _
                    Trim(vlDerCre) & vbTab & _
                    DateSerial(Mid(Trim(vlFecNacBen), 1, 4), Mid(Trim(vlFecNacBen), 5, 2), Mid(Trim(vlFecNacBen), 7, 2)) & vbTab & _
                    Trim(vlFecNacHM) & vbTab & _
                    (vlRutGrilla) & vbTab & _
                    (vlDgvGrilla) & vbTab & _
                    Trim(vlNomBen) & vbTab & Trim(vlNomBenSeg) & vbTab & _
                    Trim(vlPatBen) & vbTab & Trim(vlMatBen) & vbTab & _
                    Format(CDbl(vlPrcPension), "#,#0.00") & vbTab & _
                    Format(CDbl(vlPenBen), "#,#0.00") & vbTab & _
                    Format(CDbl(vlPenGarBen), "#,#0.00") & vbTab & _
                    Trim(vlFecFallBen) & vbTab & _
                    Trim(vgRs!Num_Orden) & vbTab & vlEstPen _
                    & vbTab & vlPrcPensionGar & vbTab & vlPrcPensionLeg _
                    & vbTab & vlcod_tipcta & vbTab & vlcod_monbco & vbTab & vlCod_Banco & vbTab & vlnum_ctabco & vbTab & vl_nacionalidadben & vbTab & vl_nacionalidadben_descripcion & vbTab & vlnum_CCI _
                    & vbTab & vl_Fono1_Ben & vbTab & vl_Fono2_Ben & vbTab & vl_ConTratDatos_Ben & vbTab & vl_ConUsoDatosCom_Ben & vbTab & vl_Correo_Ben
        vgRs.MoveNext
    Wend
    
    'No es necesario, por que los beneficiarios ya tienen el numero correcto
    'Call flModificarNumOrden
    
    vgRs.Close

Exit Function
Err_CargaBen:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'----------------------------- COTIZACION --------------------------------
'-------------------------------------------------------------------------
'FUNCION QUE BUSCA LA POLIZA Y LLENA LOS DATOS
'-------------------------------------------------------------------------
Function flBuscaCotizacion(iNumCot As String, iCodTipCot As String)
Dim vlNumero As String
On Error GoTo Err_flBuscaCotizacion

'    If Trim(iCodTipCot) = clCodTipCotOfe Then
'        vlTablaDetCotizacion = "pt_tmae_detcotizacion"
'    End If
'    If Trim(iCodTipCot) = clCodTipCotExt Then
'        vlTablaDetCotizacion = "pt_tmae_detcotexterna"
'    End If
'    If Trim(iCodTipCot) = clCodTipCotRmt Then
'        vlTablaDetCotizacion = "pt_tmae_detcotremate"
'    End If
    vlTablaDetCotizacion = "pt_tmae_detcotizacion"

    SSTab_Poliza.Tab = 0
    vlSql = ""
    vlSql = "SELECT c.num_cot, b.gls_nomben, b.gls_patben, b.gls_matben "
    vlSql = vlSql & "FROM pt_tmae_cotizacion c,pt_tmae_cotben b," & vlTablaDetCotizacion & " d "
    vlSql = vlSql & "WHERE "
    vlSql = vlSql & "c.num_cot = b.num_cot "
    vlSql = vlSql & "AND c.num_cot = d.num_cot "
    vlSql = vlSql & "AND b.cod_par = '99' "
    vlSql = vlSql & "AND d.cod_estcot = '" & clCodEstCotA & "' "
    vlSql = vlSql & "AND c.num_cot = '" & iNumCot & "'"
    Set vgRs = vgConexionBD.Execute(vlSql)

'    If IsNull(vgRs.EOF) Then
    If (vgRs.EOF) Then
        vgRs.Close
        MsgBox "La Cotización No fue Encontrada", vbCritical, "¡Error...!"
        Exit Function
    End If
    vgRs.Close

    Cmd_CrearPol.Enabled = True
'    Cmd_Nuevo.Enabled = False
    Cmd_Grabar.Enabled = False
    cmdEnviaCorreo.Enabled = False
    Cmd_Eliminar.Enabled = False
    If (vgNivelIndicadorBoton = "S") Then
        Cmd_Editar.Enabled = True
        Cmd_Eliminar.Enabled = True
    Else
        Cmd_Editar.Enabled = False
        Cmd_Eliminar.Enabled = False
    End If

'    'llenar los objetos de texto y los combos con la cotizacion escogida
'    Msf_GriAfiliado.Col = 0
'    Msf_GriAfiliado.Row = 1
'    vlNumero = Msf_GriAfiliado.Text
    vlNumero = iNumCot

    'Integracion GobiernoDeDatos
     Call LimpiarVariables
    'Fin Integracion GobiernoDeDatos_
    
    Call flCargaCotCarpAfilPol(vlNumero)
    Call flCargaCotCarpCalculo(vlNumero)
'    Call flCargaCotCarpBono(vlNumero)


    'RRR 15/09/2016
    Dim sumPrcPen As Double
    
    sumPrcPen = 0
    
    vlSql = "select sum(prc_pension) sumprc from pt_tmae_cotben where num_cot = '" & iNumCot & "'"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        sumPrcPen = CDbl(vgRs!sumPrc)
    End If
    'RRR
    Call flCargaCotCarpBenef(vlNumero, sumPrcPen)
    
'I--- ABV 08/07/2007 ---
    Txt_FecVig = DateSerial(Year(Now), Month(Now), Day(Now))
'F--- ABV 08/07/2007 ---
    
    Fra_Cabeza.Enabled = True
    Fra_Afiliado.Enabled = True
    Fra_PagPension.Enabled = True
    Fra_DatCal.Enabled = True
    Fra_SumaBono.Enabled = True
'    Fra_DatosBenef.Enabled = True
    Msf_GriAseg.Enabled = True

    'RogerPase regresar para el pase de pago doble
    Call flValidaViaPago
    
    'Lbl_Cuspp.Enabled = True
    'Txt_Digito.Enabled = True
    SSTab_Poliza.Enabled = True
    SSTab_Poliza.Tab = 0
'    Msf_GriAfiliado.Enabled = False
    
    Cmd_Poliza.Enabled = False
    Cmd_Cotizacion.Enabled = False
    Txt_FecVig.Enabled = True
'    Txt_FecVig.SetFocus
    
Exit Function
Err_flBuscaCotizacion:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'----------------------------- COTIZACION --------------------------------
'-------------------------------------------------------------------
'CARGA LOS CAMPOS DE LA CARPETA AFILIADO
'-------------------------------------------------------------------
Function flCargaCotCarpAfilPol(iNumCot As String)
On Error GoTo Err_flCargaCotCarpAfilPol
    
    Call flLimpiarDatosAfi
    
    vlSql = ""
    vlSql = "SELECT c.num_cot,"
    vlSql = vlSql & "c.cod_tipoiden,c.num_iden,c.num_cargas, "
    vlSql = vlSql & "b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben, "
    vlSql = vlSql & "b.cod_sexo,b.fec_nacben,b.fec_falben, "
    vlSql = vlSql & "c.cod_tippension,b.fec_sitinv,b.cod_cauinv,"
    vlSql = vlSql & "c.cod_estcivil,c.cod_afp,c.cod_isapre,"
    vlSql = vlSql & "c.gls_direccion,c.cod_direccion,c.gls_fono,"
    vlSql = vlSql & "c.gls_correo,'" & cgTipoNacionalidad & "' as gls_nacionalidad, "
    vlSql = vlSql & "c.cod_viapago,c.cod_sucursal,"
    vlSql = vlSql & "c.cod_tipcuenta,c.cod_banco,c.num_cuenta, "
    vlSql = vlSql & "d.num_operacion,d.num_correlativo as cod_secofe "
    vlSql = vlSql & ",c.cod_vejez "
    vlSql = vlSql & "FROM pt_tmae_cotizacion c, pt_tmae_cotben b, " & vlTablaDetCotizacion & " d "
    vlSql = vlSql & "WHERE c.num_cot = '" & iNumCot & "' "
    vlSql = vlSql & "AND c.num_cot = d.num_cot "
    vlSql = vlSql & "AND d.cod_estcot = '" & clCodEstCotA & "' "
    vlSql = vlSql & "AND c.num_cot = b.num_cot "
    vlSql = vlSql & "AND b.cod_par = '99' "
    Set vgRs = vgConexionBD.Execute(vlSql)

    If Not (vgRs.EOF) Then
        'datos de la cabecera
        If Not IsNull(vgRs!Num_Cot) Then Lbl_NumCot = vgRs!Num_Cot
        If Not IsNull(vgRs!Num_Operacion) Then Lbl_SolOfe = vgRs!Num_Operacion
        If Not IsNull(vgRs!cod_secofe) Then Lbl_SecOfe = vgRs!cod_secofe

        'datos de la carpeta de afiliado
        If Not IsNull(vgRs!NUM_IDEN) Then Txt_NumIdent = vgRs!NUM_IDEN
        If Not IsNull(vgRs!Num_Cargas) Then Txt_Asegurados = vgRs!Num_Cargas
        If Not IsNull(vgRs!Gls_NomBen) Then Txt_NomAfi = vgRs!Gls_NomBen
        If Not IsNull(vgRs!Gls_NomSegBen) Then Txt_NomAfiSeg = vgRs!Gls_NomSegBen
        If Not IsNull(vgRs!Gls_PatBen) Then Txt_ApPatAfi = vgRs!Gls_PatBen
        If Not IsNull(vgRs!Gls_MatBen) Then Txt_ApMatAfi = vgRs!Gls_MatBen
        If Not IsNull(vgRs!Fec_NacBen) Then Lbl_FecNac = DateSerial(Mid(vgRs!Fec_NacBen, 1, 4), Mid(vgRs!Fec_NacBen, 5, 2), Mid(vgRs!Fec_NacBen, 7, 2))
        If Not IsNull(vgRs!fec_falben) Then Lbl_FecFall = DateSerial(Mid(vgRs!fec_falben, 1, 4), Mid(vgRs!fec_falben, 5, 2), Mid(vgRs!fec_falben, 7, 2))
        If Not IsNull(vgRs!fec_sitinv) Then Txt_FecInv = DateSerial(Mid(vgRs!fec_sitinv, 1, 4), Mid(vgRs!fec_sitinv, 5, 2), Mid(vgRs!fec_sitinv, 7, 2))
        If Not IsNull(vgRs!Gls_Direccion) Then Txt_Dir = vgRs!Gls_Direccion
        If Not IsNull(vgRs!GLS_FONO) Then Txt_Fono = vgRs!GLS_FONO
        If Not IsNull(vgRs!GLS_CORREO) Then Txt_Correo = vgRs!GLS_CORREO
        If Not IsNull(vgRs!gls_nacionalidad) Then Txt_Nacionalidad = Trim(vgRs!gls_nacionalidad)
        If Not IsNull(vgRs!Num_Cuenta) Then Txt_NumCta = vgRs!Num_Cuenta
        
        If Not IsNull(vgRs!Cod_Sexo) Then Lbl_SexoAfi = Trim(vgRs!Cod_Sexo) & " - " & fgBuscarGlosaElemento(vgCodTabla_Sexo, vgRs!Cod_Sexo)
        If Not IsNull(vgRs!Cod_TipPension) Then Lbl_TipPen = Trim(vgRs!Cod_TipPension) & " - " & fgBuscarGlosaElemento(vgCodTabla_TipPen, vgRs!Cod_TipPension)
        'RVF 20090914
        If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
            Cmd_Representante.Enabled = True
            'DC 20091125
            Txt_FecFallBen.Visible = True
            Lbl_FecFallBen.Visible = False
        Else
            Cmd_Representante.Enabled = False
            'DC 20091125
            Txt_FecFallBen.Visible = False
            Lbl_FecFallBen.Visible = True
        End If
        '*****
        If Not IsNull(vgRs!Cod_CauInv) Then Lbl_CauInv = Trim(vgRs!Cod_CauInv) & " - " & fgBuscarGlosaCauInv(vgRs!Cod_CauInv)
        If Not IsNull(vgRs!Cod_AFP) Then Lbl_Afp = Trim(vgRs!Cod_AFP) & " - " & fgBuscarGlosaElemento(vgCodTabla_AFP, vgRs!Cod_AFP)
        
        If Not IsNull(vgRs!cod_tipoiden) Then Call fgBuscaPos(Cmb_TipoIdent, vgRs!cod_tipoiden)
        If Not IsNull(vgRs!cod_estcivil) Then Call fgBuscaPos(Cmb_EstCivil, vgRs!cod_estcivil)
        If Not IsNull(vgRs!cod_vejez) Then Call fgBuscaPos(Cmb_Vejez, vgRs!cod_vejez)
        
'I--- ABV 20/08/2007
'        Call fgBuscaPos(Cmb_Salud, (vgRs!cod_isapre))
'        Call fgBuscaPos(Cmb_ViaPago, (vgRs!Cod_ViaPago))
'        Call fgBuscaPos(Cmb_Suc, (vgRs!Cod_Sucursal))
'        Call fgBuscaPos(Cmb_TipCta, (vgRs!Cod_TipCuenta))
'        Call fgBuscaPos(Cmb_Bco, (vgRs!Cod_Banco))
Dim vlSaludCod  As String, vlSaludMod As String, vlSaludMto As Double
Dim vlFechaActual As String
Dim vlViaPagoCod As String, vlViaPagoSuc As String, vlViaPagoTC  As String
Dim vlViaPagoBco As String, vlViaPagoNumCta As String

        vlFechaActual = Format(Now, "yyyymmdd")
        
        Sql = "SELECT cod_inssalud,cod_modsalud,cod_viapago,"
        Sql = Sql & "cod_tipcuenta, cod_banco, num_cuenta,cod_sucursal "
        Sql = Sql & "FROM ma_tcod_general"
        Set vgRs4 = vgConexionBD.Execute(Sql)
        If Not vgRs4.EOF Then
            If (vgRs!cod_isapre = "00") Then
                vlSaludCod = vgRs4!Cod_InsSalud
            Else
                vlSaludCod = vgRs!cod_isapre
            End If
            vlSaludMod = vgRs4!Cod_ModSalud
            
            'se saca el mto de salud
            Sql = "SELECT mto_elemento FROM ma_tpar_tabcodvig "
            Sql = Sql & "WHERE cod_tabla= '" & vgCodTabla_PrcSal & "' AND "
            Sql = Sql & "cod_elemento = 'PSM' AND "
            Sql = Sql & "fec_inivig <= '" & (vlFechaActual) & "' AND "
            Sql = Sql & "fec_tervig >= '" & (vlFechaActual) & "'"
            Set vlRegistro = vgConexionBD.Execute(Sql)
            If Not vlRegistro.EOF Then
                vlSaludMto = Format(vlRegistro!mto_elemento, "#0.00")
            Else
                vlSaludMto = 0
            End If
            vlRegistro.Close
            
            'via pago Estándar
            If vgRs!Cod_ViaPago = "00" Then
                vlViaPagoCod = vgRs4!Cod_ViaPago
                vlViaPagoSuc = vgRs4!Cod_Sucursal
                vlViaPagoTC = vgRs4!Cod_TipCuenta
                vlViaPagoBco = vgRs4!Cod_Banco
                If Not IsNull(vgRs4!Num_Cuenta) Then
                    vlViaPagoNumCta = vgRs4!Num_Cuenta
                End If
            Else
                vlViaPagoCod = vgRs!Cod_ViaPago
                vlViaPagoSuc = vgRs!Cod_Sucursal
                vlViaPagoTC = vgRs!Cod_TipCuenta
                vlViaPagoBco = vgRs!Cod_Banco
                If Not IsNull(vgRs!Num_Cuenta) Then
                    vlViaPagoNumCta = vgRs!Num_Cuenta
                End If
            End If
        Else
            vlSaludCod = vgRs!cod_isapre
            vlViaPagoCod = vgRs!Cod_ViaPago
            vlViaPagoSuc = vgRs!Cod_Sucursal
            vlViaPagoTC = vgRs!Cod_TipCuenta
            vlViaPagoBco = vgRs!Cod_Banco
        End If
        vgRs4.Close
        
        If (vlViaPagoCod = "05") Then
        Else
        End If
        
        Call fgBuscaPos(Cmb_Salud, vlSaludCod)
        vlSwViaPago = True
        Call fgBuscaPos(Cmb_ViaPago, vlViaPagoCod)
        vlSwViaPago = False
        Call fgBuscaPos(Cmb_Suc, vlViaPagoSuc)
        Call fgBuscaPos(Cmb_TipCta, vlViaPagoTC)
        Call fgBuscaPos(Cmb_Bco, vlViaPagoBco)
'F--- ABV 21/08/2007

        If Not IsNull(vgRs!Cod_Direccion) And (vgRs!Cod_Direccion <> 0) Then
            vlCodDireccion = vgRs!Cod_Direccion
            Call fgBuscarNombreComunaProvinciaRegion(vlCodDireccion)
            Lbl_Departamento = vgNombreRegion
            Lbl_Provincia = vgNombreProvincia
            Lbl_Distrito = vgNombreComuna
        End If
        
        'Rogerpase descomentar con el pase a produccion pago directo
        
        Cmb_ViaPago.Clear
        Call fgComboGeneral(vgCodTabla_ViaPago, Cmb_ViaPago)
        
       'INICIO GCP-FRACTAL [REQUERIMIENTO SD-5041] 21032019
        Dim TP As String
        TP = Left(Trim(Lbl_TipPen.Caption), 2)
        If (TP = "04" Or TP = "05") Then
          'Se Remueve el via pago TRANSFERENCIA AFP
          Call RemoveItemCombo(Cmb_ViaPago, "04")

       Else
        'Se Remueve el via pago DEPOSITO EN CUENTA
          Call RemoveItemCombo(Cmb_ViaPago, "02")
       End If
       'FIN GCP-FRACTAL [REQUERIMIENTO SD-5041] 21032019

            
    End If

Exit Function
Err_flCargaCotCarpAfilPol:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'----------------------------- COTIZACION --------------------------------
'--------------------------------------------------------------------
'CARGA LOS CAMPOS DE LA CARPETA DATOS DE CALCULO
'--------------------------------------------------------------------
Function flCargaCotCarpCalculo(iNumCot As String)
On Error GoTo Err_flCargaCotCarpCalculo
Dim vlDias As Long

    Call flLimpiarDatosCal
    
    vlMarcaSobDif = "N"

    vlSql = ""
    vlSql = "SELECT d.cod_tipren,d.num_mesdif,d.cod_modalidad,d.cod_dercre,d.cod_dergra, "
    vlSql = vlSql & "d.num_mesgar,d.prc_rentaafp,d.prc_rentatmp,c.cod_cuspp,c.ind_cob,"
    vlSql = vlSql & "d.cod_cobercon,d.mto_facpenella,d.prc_facpenella,c.cod_tippension,"
    vlSql = vlSql & "d.prc_corcom,d.prc_corcomreal,c.cod_bensocial,c.cod_tipoidencor,c.num_idencor, "
    vlSql = vlSql & "c.fec_dev,d.cod_moneda,d.fec_acepta,d.mto_cnu,d.prc_tasatce,"
    vlSql = vlSql & "d.prc_tasavta,d.prc_tasatir,d.prc_tasapergar, "
    vlSql = vlSql & "d.mto_pension,d.mto_pensiongar,c.mto_priuni,'" & vgMonedaCodOfi & "' as cod_monedafon,"
'I--- ABV 26/06/2006 ---
'    vlSql = vlSql & "d.mto_ctaindafp, "
    vlSql = vlSql & "c.mto_ctaind, "
'F--- ABV 26/06/2006 ---
    vlSql = vlSql & "c.mto_bono,d.mto_priunisim,d.mto_priunidif "
    vlSql = vlSql & ",d.fec_calculo,c.mto_apoadi "
'I--- ABV 04/12/2009 ---
    vlSql = vlSql & ",d.ind_calsobdif "
'F--- ABV 04/12/2009 ---
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & ",d.cod_tipreajuste,d.mto_valreajustetri,d.mto_valreajustemen "
    vlSql = vlSql & ",mtr.cod_scomp as cod_montipreaju,mtr.gls_descripcion as gls_montipreaju "
'F--- ABV 05/02/2011 ---
    vlSql = vlSql & "FROM pt_tmae_cotizacion c, " & vlTablaDetCotizacion & " d "
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & ",ma_tpar_monedatiporeaju mtr "
'F--- ABV 05/02/2011 ---
    vlSql = vlSql & "WHERE c.num_cot = '" & iNumCot & "' AND "
    vlSql = vlSql & "c.num_cot = d.num_cot AND "
    vlSql = vlSql & "d.cod_estcot = '" & clCodEstCotA & "' "
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & "AND d.cod_tipreajuste = mtr.cod_tipreajuste(+) AND "
    vlSql = vlSql & "d.cod_moneda = mtr.cod_moneda(+) "
'F--- ABV 05/02/2011 ---
    Set vgRs = vgConexionBD.Execute(vlSql)

    If Not (vgRs.EOF) Then

        Lbl_FecDev = DateSerial(Mid(vgRs!Fec_Dev, 1, 4), Mid(vgRs!Fec_Dev, 5, 2), Mid(vgRs!Fec_Dev, 7, 2))
        If Not IsNull(vgRs!Fec_Acepta) Then
         Lbl_FecIncorpora = DateSerial(Mid(vgRs!Fec_Acepta, 1, 4), Mid(vgRs!Fec_Acepta, 5, 2), Mid(vgRs!Fec_Acepta, 7, 2))
           vlFecRepPrimasEst = Format(DateSerial(Mid(vgRs!Fec_Acepta, 1, 4), Mid(vgRs!Fec_Acepta, 5, 2), CLng(Mid(vgRs!Fec_Acepta, 7, 2)) + 3), "yyyymmdd")
        Txt_FecIniPago = fgCalcularFechaPrimerPagoEst(vgRs!Fec_Acepta, vgRs!Fec_Dev, vlFecRepPrimasEst, vgRs!Cod_TipPension, vgRs!Cod_TipRen, vgRs!Num_MesDif, vlDias)
        End If
        
       
        vlDias = fgObtieneDiasPrimerPagoEst(vgRs!Cod_TipPension)
      
        Lbl_CUSPP = vgRs!Cod_Cuspp
        Lbl_FecCalculo = DateSerial(Mid(vgRs!Fec_Calculo, 1, 4), Mid(vgRs!Fec_Calculo, 5, 2), Mid(vgRs!Fec_Calculo, 7, 2))
                
        If Not IsNull(vgRs!Num_MesDif) And vgRs!Num_MesDif > 0 Then Lbl_AnnosDif = Trim(vgRs!Num_MesDif / 12)
        If Not IsNull(vgRs!Num_MesGar) Then Lbl_MesesGar = vgRs!Num_MesGar
        If Not IsNull(vgRs!Prc_RentaAFP) Then Lbl_RentaAFP = Format(vgRs!Prc_RentaAFP, "#0.00")
        If Not IsNull(vgRs!Prc_RentaTMP) Then Lbl_PrcRentaTmp = Format(vgRs!Prc_RentaTMP, "#0.00")
        If Not IsNull(vgRs!Cod_CoberCon) Then Lbl_FacPenElla = vgRs!Cod_CoberCon
        If Not IsNull(vgRs!Prc_CorCom) Then Lbl_ComIntBen = Format(vgRs!Prc_CorCom, "#0.00")
        If Not IsNull(vgRs!Prc_CorComReal) Then Lbl_ComInt = Format(vgRs!Prc_CorComReal, "#0.00")
        
        'Buscar la Identificación del Intermediario
        If Not IsNull(vgRs!cod_tipoidencor) Then
            Lbl_TipoIdentCorr = vgRs!cod_tipoidencor & " - " & fgBuscarNombreTipoIden(vgRs!cod_tipoidencor)
        End If
        If Not IsNull(vgRs!Num_IdenCor) Then
            If Not IsNull(vgRs!Num_IdenCor) Then Lbl_NumIdentCorr = vgRs!Num_IdenCor
        End If
        
        If Not IsNull(vgRs!Mto_CNU) Then Txt_PrcFam = Format(vgRs!Mto_CNU, "#,#0.000000")
        If Not IsNull(vgRs!Mto_PriUniSim) Then Lbl_MtoPrimaUniSim = Format(vgRs!Mto_PriUniSim, "#,#0.00")
        If Not IsNull(vgRs!Mto_PriUniDif) Then Lbl_MtoPrimaUniDif = Format(vgRs!Mto_PriUniDif, "#,#0.00")
        If Not IsNull(vgRs!prc_tasatce) Then Lbl_TasaCtoEq = Format(vgRs!prc_tasatce, "#0.00")
        If Not IsNull(vgRs!Prc_TasaVta) Then Lbl_TasaVta = Format(vgRs!Prc_TasaVta, "#0.00")
        If Not IsNull(vgRs!Prc_TasaTir) Then Lbl_TasaTIR = Format(vgRs!Prc_TasaTir, "#0.00")
        If Not IsNull(vgRs!prc_tasapergar) Then Lbl_TasaPerGar = Format(vgRs!prc_tasapergar, "#0.00")
        If Not IsNull(vgRs!Mto_Pension) Then Lbl_MtoPension = Format(vgRs!Mto_Pension, "#,#0.00")
        If Not IsNull(vgRs!Mto_PensionGar) Then Lbl_MtoPensionGar = Format(vgRs!Mto_PensionGar, "#,#0.00")
        If Not IsNull(vgRs!Mto_CtaInd) Then Lbl_CtaInd = Format(vgRs!Mto_CtaInd, "#,#0.00")
        If Not IsNull(vgRs!Mto_Bono) Then Lbl_BonoAct = Format(vgRs!Mto_Bono, "#,#0.00")
        If Not IsNull(vgRs!Mto_ApoAdi) Then Lbl_ApoAdi = Format(vgRs!Mto_ApoAdi, "#,#0.00")
        If Not IsNull(vgRs!MTO_PRIUNI) Then Lbl_PriUnica = Format(vgRs!MTO_PRIUNI, "#,#0.00")

        If Not IsNull(vgRs!Cod_TipRen) Then Lbl_TipoRenta = Trim(vgRs!Cod_TipRen) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_TipRen, Trim(vgRs!Cod_TipRen)))
        If Not IsNull(vgRs!Cod_Modalidad) Then Lbl_Alter = Trim(vgRs!Cod_Modalidad) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_AltPen, Trim(vgRs!Cod_Modalidad)))
        
'Nuevos Campos
        If Not IsNull(vgRs!Ind_Cob) Then
            If vgRs!Ind_Cob = "S" Then Lbl_IndCob = cgIndicadorSi Else Lbl_IndCob = cgIndicadorNo
        End If
        If Not IsNull(vgRs!Cod_DerCre) Then
            If vgRs!Cod_DerCre = "S" Then Lbl_DerCre = cgIndicadorSi Else Lbl_DerCre = cgIndicadorNo
        End If
        If Not IsNull(vgRs!Cod_DerGra) Then
            If vgRs!Cod_DerGra = "S" Then Lbl_DerGra = cgIndicadorSi Else Lbl_DerGra = cgIndicadorNo
        End If
        If Not IsNull(vgRs!Cod_BenSocial) Then
            If vgRs!Cod_BenSocial = "S" Then Lbl_BenSocial = cgIndicadorSi Else Lbl_BenSocial = cgIndicadorNo
        End If
        vlElemento = fgBuscarGlosaElemento(vgCodTabla_TipMon, vgRs!Cod_Moneda)
        
        Lbl_Moneda(0) = (vgRs!Cod_Moneda) + " - " + vlElemento

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
        
        'Tipo de Moneda de la Modalidad
        Lbl_Moneda(clMonedaBenPen) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_Moneda)
        Lbl_Moneda(clMonedaBenPenGar) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_Moneda)
        
        Lbl_Moneda(clMonedaModPen) = Lbl_Moneda(clMonedaBenPen)
        Lbl_Moneda(clMonedaModPenGar) = Lbl_Moneda(clMonedaBenPen)
        
        'Tipo de Moneda del Fondo
        Lbl_MonedaFon(clMtoCtaIndFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        Lbl_MonedaFon(clMtoBonoFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        Lbl_MonedaFon(clMtoPriUniFon) = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_MonedaFon)
        
'I--- ABV 10/08/2007 ---
        'Guardar en Variables los Datos para la corrección de Datos en Grilla de Beneficiarios
        vlFecDev = vgRs!Fec_Dev
        vlCodTipPen = Trim(vgRs!Cod_TipPension)
        vlIndCobertura = Trim(vgRs!Ind_Cob)
        vlCodCoberCon = Trim(vgRs!Cod_CoberCon)
        vlMtoFacPenElla = vgRs!Mto_FacPenElla
        vlPrcFacPenElla = vgRs!Prc_FacPenElla
        vlDerCrePol = Trim(vgRs!Cod_DerCre)
        vlPension = vgRs!Mto_Pension
        vlMesGar = vgRs!Num_MesGar
'F--- ABV 10/08/2007 ---
        
'I--- ABV 04/12/2009 ---
        If Not IsNull(vgRs!Ind_CalSobDif) Then
            vlMarcaSobDif = Trim(vgRs!Ind_CalSobDif)
        End If
'F--- ABV 04/12/2009 ---

    End If
    vgRs.Close

Exit Function
Err_flCargaCotCarpCalculo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'----------------------------- COTIZACION --------------------------------
'-------------------------------------------------------------------------
'FUNCION CARGA DATOS DEL BONO
'-------------------------------------------------------------------------
Function flCargaCotCarpBono(iNumCot As String)
On Error GoTo Err_flCargaCotCarpBono

    Call flLimpiarDatosBono
    'Call flInicializaGrillaBono
    Call flCargaGrillaBonoCot(iNumCot)

Exit Function
Err_flCargaCotCarpBono:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'----------------------------- COTIZACION --------------------------------
'------------------------------------------------------------------------
'CARGA LA GRILLA CON LOS DATOS DE LOS BENEFICIARIOS
'------------------------------------------------------------------------
Function flCargaCotCarpBenef(iNumCot As String, sumPrc As Double)
On Error GoTo Err_flCargaCotCarpBenef
    Dim vlTipoIden As String
    Dim valIsSob As Double

    valIsSob = flase

    Msf_GriAseg.rows = 1
    
    vlSql = "SELECT "
    vlSql = vlSql & "num_orden,cod_grufam,cod_par,cod_sexo,cod_sitinv, "
    vlSql = vlSql & "cod_cauinv,cod_derpen,cod_dercre,cod_tipoiden,num_iden, "
    vlSql = vlSql & "gls_nomben,gls_nomsegben,gls_patben,gls_matben, "
    vlSql = vlSql & "fec_nacben,fec_sitinv,fec_falben,Fec_NacHM,"
    vlSql = vlSql & "prc_pension,prc_pensionleg,mto_pension,mto_pensiongar"
    ', cod_tipcta, cod_monbco, cod_banco, num_ctabco "
'I--- ABV 04/12/2009 ---
    vlSql = vlSql & ",prc_pensionsobdif "
'F--- ABV 04/12/2009 ---
    vlSql = vlSql & "FROM pt_tmae_cotben WHERE "
    vlSql = vlSql & "num_cot = '" & iNumCot & "' "
    vlSql = vlSql & "ORDER BY num_orden ASC "
    Set vgRs = vgConexionBD.Execute(vlSql)
    While Not vgRs.EOF
        vlCodPar = Trim(vgRs!Cod_Par)  '& " - " & fgBuscarGlosaElemento(vgCodTabla_Par, vgRs!cod_par)
        vlGruFam = Trim(vgRs!Cod_GruFam)  '& " - " & fgBuscarGlosaElemento(vgCodTabla_GruFam, vgRs!cod_grufam)
        vlSexoBen = Trim(vgRs!Cod_Sexo)  '& " - " & fgBuscarGlosaElemento(vgCodTabla_Sexo, vgRs!cod_sexo)
        vlSitInv = Trim(vgRs!Cod_SitInv)  '& " - " & fgBuscarGlosaElemento(vgCodTabla_SitInv, vgRs!cod_sitinv)
        vlDerPen = Trim(vgRs!Cod_DerPen)  '& " - " & fgBuscarGlosaElemento(vgCodTabla_DerPen, vgRs!cod_derpen)
        
        If Not IsNull(vgRs!Cod_DerCre) Then vlDerCre = Trim(vgRs!Cod_DerCre) Else vlDerCre = ""
        If Not IsNull(vgRs!Cod_CauInv) Then vlCauInv = Trim(vgRs!Cod_CauInv)  '& " - " & fgBuscarGlosaCauInv(vgRs!cod_cauinv)
        If Not IsNull(vgRs!fec_sitinv) Then
            vlFecInv = DateSerial(Mid(Trim(vgRs!fec_sitinv), 1, 4), Mid(Trim(vgRs!fec_sitinv), 5, 2), Mid(Trim(vgRs!fec_sitinv), 7, 2))
        Else
            vlFecInv = ""
        End If
        
        If Not IsNull(vgRs!cod_tipoiden) Then
            vlRutGrilla = " " & vgRs!cod_tipoiden & " - " & fgBuscarNombreTipoIden(vgRs!cod_tipoiden)
        Else
            vlRutGrilla = ""
        End If
        
        If Not IsNull(vgRs!NUM_IDEN) Then
            vlDgvGrilla = Trim(vgRs!NUM_IDEN)
        Else
            vlDgvGrilla = ""
        End If
        
        If Not IsNull(vgRs!Gls_NomBen) Then vlNomBen = (vgRs!Gls_NomBen) Else vlNomBen = ""
        If Not IsNull(vgRs!Gls_NomSegBen) Then vlNomBenSeg = (vgRs!Gls_NomSegBen) Else vlNomBenSeg = ""
        If Not IsNull(vgRs!Gls_PatBen) Then vlPatBen = (vgRs!Gls_PatBen) Else vlPatBen = ""
        If Not IsNull(vgRs!Gls_MatBen) Then vlMatBen = (vgRs!Gls_MatBen) Else vlMatBen = ""
        
'        'RRR
'        If Not IsNull(vgRs!cod_tipcta) Then vlcod_tipcta = (vgRs!cod_tipcta) Else vlcod_tipcta = ""
'        If Not IsNull(vgRs!cod_monbco) Then vlcod_monbco = (vgRs!cod_monbco) Else vlcod_monbco = ""
'        If Not IsNull(vgRs!Cod_Banco) Then vlCod_Banco = (vgRs!Cod_Banco) Else vlCod_Banco = ""
'        If Not IsNull(vgRs!num_ctabco) Then vlnum_ctabco = (vgRs!num_ctabco) Else vlnum_ctabco = ""
        
        If Not IsNull(vgRs!fec_falben) Then
            valIsSob = True
            vlFecFallBen = DateSerial(Mid(Trim(vgRs!fec_falben), 1, 4), Mid(Trim(vgRs!fec_falben), 5, 2), Mid(Trim(vgRs!fec_falben), 7, 2))
        Else
            vlFecFallBen = ""
        End If
        If Not IsNull(vgRs!Fec_NacHM) Then
            vlFecNacHM = DateSerial(Mid(Trim(vgRs!Fec_NacHM), 1, 4), Mid(Trim(vgRs!Fec_NacHM), 5, 2), Mid(Trim(vgRs!Fec_NacHM), 7, 2))
        Else
            vlFecNacHM = ""
        End If

'I--- ABV 04/12/2009 ---
'        vlPrcPension = vgRs!Prc_Pension
'        vlPrcPensionLeg = vgRs!Prc_PensionLeg
        If vlMarcaSobDif = "S" Then
            vlPrcPension = vgRs!prc_pensionsobdif
            vlPrcPensionLeg = vgRs!prc_pensionsobdif
        Else
            If (vlCodCoberCon <> "0") Then
                'Si tiene definido el Código de Cobertura a la Cónyuge, asignarle a ella el % como % de Pago
                If vlCodPar = "10" Then
                    vlPrcPension = vlPrcFacPenElla
                Else           '31/03/2010 DCM
                    vlPrcPension = vgRs!Prc_Pension
                End If
            Else
                vlPrcPension = vgRs!Prc_Pension
            End If
            vlPrcPensionLeg = vgRs!Prc_PensionLeg
        End If
'F--- ABV 04/12/2009 ---
        
        vlFecNacBen = vgRs!Fec_NacBen
        vlPenBen = vgRs!Mto_Pension
        vlPenGarBen = vgRs!Mto_PensionGar
        
'I--- ABV 10/08/2007 ---
        If valIsSob = True Then
            If (vlMesGar > 0) Then
                vlPrcPensionGar = CDbl(vlPrcPension / sumPrc) * 100
            Else
                vlPrcPensionGar = "0"
            End If
        Else
            If (vlMesGar > 0) Then
                If vlCodPar = 99 Then
                    vlPrcPensionGar = vlPrcPension
                Else
                    vlPrcPensionGar = 0
                End If
            Else
                vlPrcPensionGar = "0"
            End If
        End If
'I--- ABV 04/12/2009 ---
        'Cuando se trate de los casos de Sobrevivencia con Periodo Diferido (Marca), el Tipo de Parentesco
        'debe cambiar a Cónyuge o Madre Sin Hijos
        If vlMarcaSobDif = "S" Then
            If vlCodPar = "11" Then
                vlCodPar = "10"
            End If
            If vlCodPar = "21" Then
                vlCodPar = "20"
            End If
        End If
'F--- ABV 04/12/2009 ---
        
        'Corregir el Derecho a Crecer de la Cónyuge o Madre con Hijos => 11 ó 21
        If vlCodPar = "11" Or vlCodPar = "21" Then
            vlDerCre = vlDerCrePol
        End If
        ''Corregir Porcentaje para la Cobertura de la Cónyuge (S/Hijos) definida en la Cotizazión
        'If (vlCodCoberCon <> "0") And (vlCodCoberCon <> "") Then
        '    If vlCodPar = "10" Then vlPrcPension = vlPrcFacPenElla
        'End If
        
        'Corregir el Estado de Pago para la Pensión
'I--- ABV 04/12/2009 ---
'        vlEstPen = fgCalcularEstadoPagoPension(vlFecDev, vlCodTipPen, vlCodPar, vlFecNacBen, vlFecFallBen, "", vlSitInv)
        If vlMarcaSobDif = "S" And vlCodPar = "30" Then
            'Si se encuentra como Sobrevivencia Dif. y es un Hijo, este se deja como Sin Derecho a Pago de Pensión
            vlEstPen = 10 'Sin Derecho a Pago de Pensiones
        Else
            vlEstPen = fgCalcularEstadoPagoPension(vlFecDev, vlCodTipPen, vlCodPar, vlFecNacBen, vlFecFallBen, "", vlSitInv)
        End If
'F--- ABV 04/12/2009 ---

'F--- ABV 10/08/2007 ---
        
        
        If Trim(vgRs!Num_Orden) = clNumOrdenCau Then
        
            Msf_GriAseg.AddItem Trim(vgRs!Num_Orden) & vbTab & _
                            Trim(vlCodPar) & vbTab & _
                            Trim(vlGruFam) & vbTab & _
                            Trim(vlSexoBen) & vbTab & _
                            Trim(vlSitInv) & vbTab & _
                            Trim(vlFecInv) & vbTab & _
                            Trim(vlCauInv) & vbTab & _
                            Trim(vlDerPen) & vbTab & _
                            Trim(vlDerCre) & vbTab & _
                            DateSerial(Mid(Trim(vlFecNacBen), 1, 4), Mid(Trim(vlFecNacBen), 5, 2), Mid(Trim(vlFecNacBen), 7, 2)) & vbTab & _
                            Trim(vlFecNacHM) & vbTab & _
                            (vlRutGrilla) & vbTab & _
                            (vlDgvGrilla) & vbTab & _
                            Trim(vlNomBen) & vbTab & Trim(vlNomBenSeg) & vbTab & _
                            Trim(vlPatBen) & vbTab & Trim(vlMatBen) & vbTab & _
                            Format(CDbl(vlPrcPension), "#,#0.000") & vbTab & _
                            Format(CDbl(vlPenBen), "#,#0.00") & vbTab & _
                            Format(CDbl(vlPenGarBen), "#,#0.00") & vbTab & _
                            Trim(vlFecFallBen) & vbTab & _
                            Trim(vgRs!Num_Orden) & vbTab & vlEstPen _
                            & vbTab & vlPrcPensionGar & vbTab & vlPrcPensionLeg, 1
                            '& vbTab & vlcod_tipcta & vbTab & vlcod_monbco & vbTab & vlCod_Banco & vbTab & vlnum_ctabco, 1
        Else
            Msf_GriAseg.AddItem Trim(vgRs!Num_Orden) & vbTab & _
                            Trim(vlCodPar) & vbTab & _
                            Trim(vlGruFam) & vbTab & _
                            Trim(vlSexoBen) & vbTab & _
                            Trim(vlSitInv) & vbTab & _
                            Trim(vlFecInv) & vbTab & _
                            Trim(vlCauInv) & vbTab & _
                            Trim(vlDerPen) & vbTab & _
                            Trim(vlDerCre) & vbTab & _
                            DateSerial(Mid(Trim(vlFecNacBen), 1, 4), Mid(Trim(vlFecNacBen), 5, 2), Mid(Trim(vlFecNacBen), 7, 2)) & vbTab & _
                            Trim(vlFecNacHM) & vbTab & _
                            (vlRutGrilla) & vbTab & _
                            (vlDgvGrilla) & vbTab & _
                            Trim(vlNomBen) & vbTab & Trim(vlNomBenSeg) & vbTab & _
                            Trim(vlPatBen) & vbTab & Trim(vlMatBen) & vbTab & _
                            Format(CDbl(vlPrcPension), "#,#0.000") & vbTab & _
                            Format(CDbl(vlPenBen), "#,#0.00") & vbTab & _
                            Format(CDbl(vlPenGarBen), "#,#0.00") & vbTab & _
                            Trim(vlFecFallBen) & vbTab & _
                            Trim(vgRs!Num_Orden) & vbTab & vlEstPen _
                            & vbTab & vlPrcPensionGar & vbTab & vlPrcPensionLeg
                            '& vbTab & vlcod_tipcta & vbTab & vlcod_monbco & vbTab & vlCod_Banco & vbTab & vlnum_ctabco
        End If
        vgRs.MoveNext
    Wend
    
    Call flModificarNumOrden
    
    vgRs.Close

''I--- ABV 10/08/2007 ---
''Recalcular los Montos de Pensión
'    Call fgCargaEstBenGrilla(Msf_GriAseg, stBeneficiariosMod, vlFecDev)
'    vlNumCargas = (Msf_GriAseg.Rows - 1)
'
'    Call fgCalcularPorcentajeBenef(vlFecDev, vlNumCargas, stBeneficiariosMod, vlCodTipPen, CDbl(vlPension), True, vlDerCrePol, vlIndCobertura, False, vlMesGar)
'
'    Call fgActualizaGrillaBeneficiarios(Msf_GriAseg, stBeneficiariosMod, vlNumCargas)
''F--- ABV 10/08/2007 ---

Exit Function
Err_flCargaCotCarpBenef:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flModificarNumOrden()

    For vlJ = 1 To (Msf_GriAseg.rows - 1)
        Msf_GriAseg.Row = vlJ
        Msf_GriAseg.Col = 0
        Msf_GriAseg.Text = Msf_GriAseg.Row
    Next vlJ

End Function
'-------------------------------------------------------------------------
'HABILITA O DESHABILITA LOS CAMPOS DE FORMA DE PAGO SEGUN LA VIA DE PAGO
'-------------------------------------------------------------------------
Function flValidaViaPago()

'If Cmb_ViaPago.Enabled = True Then
'
'
'End If
    
If vlSwViaPago = False Then

    vlNumero = InStr(Cmb_ViaPago.Text, "-")
    vlOpcion = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
    
'Opción 01 = Via de Pago CAJA
    If vlSw = False Then
        fgComboSucursal Cmb_Suc, "A"
        If vlOpcion = "00" Or vlOpcion = "05" Or vlOpcion = "08" Then
            Cmb_Suc.Enabled = False
            Cmb_TipCta.Enabled = False
            Cmb_Bco.Enabled = False
            cmb_MonCta.Enabled = False
            Txt_NumCta.Enabled = False
            txt_CCI.Enabled = False
            Cmb_Suc.ListIndex = 0
            Cmb_TipCta.ListIndex = 0
            Cmb_Bco.ListIndex = 0
            Txt_NumCta.Text = ""
            txt_CCI.Text = ""
        Else
            If vlOpcion = "01" Or vlOpcion = "04" Then
                If (vlOpcion = "04") Then
                    vgTipoSucursal = cgTipoSucursalAfp
                Else
                    vgTipoSucursal = cgTipoSucursalSuc
                End If
                fgComboSucursal Cmb_Suc, vgTipoSucursal
                
                vgPalabra = fgObtenerCodigo_TextoCompuesto(Lbl_Afp)
                Call fgBuscaPos(Cmb_Suc, vgPalabra)
                Cmb_TipCta.Enabled = False
                Cmb_Bco.Enabled = False
                Txt_NumCta.Enabled = False
                Cmb_Suc.Enabled = True
                Cmb_TipCta.ListIndex = 0
                Cmb_Bco.ListIndex = 0
                
                
'                If (vlCombo = "04") Then
'                    vgPalabra = fgObtenerCodigo_TextoCompuesto(Lbl_Afp)
'                    Call fgBuscaPos(Cmb_Suc, vgPalabra)
'                End If
                'If Cmb_TipCta.ListCount <> 0 Then
                    
                'End If
                'If Cmb_Bco.ListCount <> 0 Then
                    
                'End If
                
'                Cmb_Sucursal.Enabled = True
'                cmbTipoCtaBen.Enabled = False
'                cmbBancoCtaBen.Enabled = False
'                cmbMonctaBen.Enabled = False
'                txtNumctaBen.Enabled = False
'                cmbTipoCtaBen.ListIndex = 0
'                cmbBancoCtaBen.ListIndex = 0
'                If (vlOpcion = "04") Then
'                  '  vgPalabra = fgObtenerCodigo_TextoCompuesto(vlAfp)
'                  '  Call fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_Sucursal)
'                End If
               ' txtNumctaBen.Text = ""
            Else
                If (vlOpcion = "02") Then
                   Call fgComboGeneral(vgCodTabla_Bco, Cmb_Bco)
                   Cmb_Suc.Enabled = False
                   Cmb_TipCta.Enabled = True
                   Cmb_Bco.Enabled = True
                   cmb_MonCta.Enabled = True
                   Txt_NumCta.Enabled = True
                   txt_CCI.Enabled = True
                   Cmb_Suc.ListIndex = 0
                   
'                   Txt_NumCuenta.Text = ""

                ElseIf (vlOpcion = "07") Then
                   Call fgComboGeneral(vgCodTabla_Bco, Cmb_Bco)
                   Cmb_Suc.Enabled = False
                   Cmb_TipCta.Enabled = True
                   Cmb_Bco.Enabled = True
                   cmb_MonCta.Enabled = True
                   Txt_NumCta.Enabled = True
                   txt_CCI.Enabled = True
                   Cmb_Suc.ListIndex = 0
'                   Txt_NumCuenta.Text = ""
                ElseIf (vlOpcion = "06") Then
                    Cmb_Suc.Enabled = False
                    Cmb_TipCta.Enabled = False
                    Call DejarBcoPrinicipales(Cmb_Bco)
                    'Cmb_Bco.ListIndex = fgBuscarPosicionCodigoCombo(Lbl_MonPension(2).Caption, cmbMonctaBen)
                    Cmb_Bco.Enabled = True
                    cmb_MonCta.Enabled = False
                    cmbMonctaBen.ListIndex = fgBuscarPosicionCodigoCombo("NS", cmb_MonCta)
                    Txt_NumCta.Enabled = False
                    Cmb_Suc.ListIndex = 0
                    Cmb_TipCta.ListIndex = 0
                    'Cmb_Bco.ListIndex = 0
                    Txt_NumCta.Text = ""
                    txt_CCI.Text = ""
'                Else
'                    Cmb_Sucursal.Enabled = False
'                    cmbTipoCtaBen.Enabled = False
'                    cmbBancoCtaBen.Enabled = False
'                    cmbMonctaBen.Enabled = False
'                    txtNumctaBen.Enabled = False
'                    Cmb_Sucursal.ListIndex = 0
'                    cmbTipoCtaBen.ListIndex = 0
'                    cmbBancoCtaBen.ListIndex = 0
'                    txtNumctaBen.Text = ""
'                    txt_CCI.Text = ""
                End If
            End If
        End If
'    Else
'        fgComboSucursal Cmb_Suc, "A"
'        If vlOpcion = "00" Or vlOpcion = "05" Or vlOpcion = "08" Then
'            Cmb_Suc.Enabled = False
'            Cmb_TipCta.Enabled = False
'            Cmb_Bco.Enabled = False
'            cmb_MonCta.Enabled = False
'            Txt_NumCta.Enabled = False
'            txt_CCI.Enabled = False
'            Cmb_Suc.ListIndex = 0
'            Cmb_TipCta.ListIndex = 0
'            Cmb_Bco.ListIndex = 0
'            Txt_NumCta.Text = ""
'            txt_CCI.Text = ""
'        Else
'            If vlOpcion = "01" Or vlOpcion = "04" Then
'                If (vlOpcion = "04") Then
'                    vgTipoSucursal = cgTipoSucursalAfp
'                Else
'                    vgTipoSucursal = cgTipoSucursalSuc
'                End If
'                fgComboSucursal Cmb_Suc, vgTipoSucursal
'
'                vgPalabra = fgObtenerCodigo_TextoCompuesto(Lbl_Afp)
'                Call fgBuscaPos(Cmb_Suc, vgPalabra)
'                Cmb_TipCta.Enabled = False
'                Cmb_Bco.Enabled = False
'                Txt_NumCta.Enabled = False
'                Cmb_Suc.Enabled = True
'                Cmb_TipCta.ListIndex = 0
'                Cmb_Bco.ListIndex = 0
'
'
''                If (vlCombo = "04") Then
''                    vgPalabra = fgObtenerCodigo_TextoCompuesto(Lbl_Afp)
''                    Call fgBuscaPos(Cmb_Suc, vgPalabra)
''                End If
'                'If Cmb_TipCta.ListCount <> 0 Then
'
'                'End If
'                'If Cmb_Bco.ListCount <> 0 Then
'
'                'End If
'
''                Cmb_Sucursal.Enabled = True
''                cmbTipoCtaBen.Enabled = False
''                cmbBancoCtaBen.Enabled = False
''                cmbMonctaBen.Enabled = False
''                txtNumctaBen.Enabled = False
''                cmbTipoCtaBen.ListIndex = 0
''                cmbBancoCtaBen.ListIndex = 0
''                If (vlOpcion = "04") Then
''                  '  vgPalabra = fgObtenerCodigo_TextoCompuesto(vlAfp)
''                  '  Call fgBuscarPosicionCodigoCombo(vgPalabra, Cmb_Sucursal)
''                End If
'               ' txtNumctaBen.Text = ""
'            Else
'                If (vlOpcion = "02") Then
'                   Call fgComboGeneral(vgCodTabla_Bco, Cmb_Bco)
'                   Cmb_Suc.Enabled = False
'                   Cmb_TipCta.Enabled = True
'                   Cmb_Bco.Enabled = True
'                   cmb_MonCta.Enabled = True
'                   Txt_NumCta.Enabled = True
'                   txt_CCI.Enabled = True
'                   Cmb_Suc.ListIndex = 0
'
''                   Txt_NumCuenta.Text = ""
'
'                ElseIf (vlOpcion = "07") Then
'                   Call fgComboGeneral(vgCodTabla_Bco, Cmb_Bco)
'                   Cmb_Suc.Enabled = False
'                   Cmb_TipCta.Enabled = True
'                   Cmb_Bco.Enabled = True
'                   cmb_MonCta.Enabled = True
'                   Txt_NumCta.Enabled = True
'                   txt_CCI.Enabled = True
'                   Cmb_Suc.ListIndex = 0
''                   Txt_NumCuenta.Text = ""
'                ElseIf (vlOpcion = "06") Then
'                    Cmb_Suc.Enabled = False
'                    Cmb_TipCta.Enabled = False
'                    Call DejarBcoPrinicipales(Cmb_Bco)
'                    'Cmb_Bco.ListIndex = fgBuscarPosicionCodigoCombo(Lbl_MonPension(2).Caption, cmbMonctaBen)
'                    Cmb_Bco.Enabled = True
'                    cmb_MonCta.Enabled = False
'                    cmbMonctaBen.ListIndex = fgBuscarPosicionCodigoCombo("NS", cmb_MonCta)
'                    Txt_NumCta.Enabled = False
'                    Cmb_Suc.ListIndex = 0
'                    Cmb_TipCta.ListIndex = 0
'                    'Cmb_Bco.ListIndex = 0
'                    Txt_NumCta.Text = ""
'                    txt_CCI.Text = ""
''                Else
''                    Cmb_Sucursal.Enabled = False
''                    cmbTipoCtaBen.Enabled = False
''                    cmbBancoCtaBen.Enabled = False
''                    cmbMonctaBen.Enabled = False
''                    txtNumctaBen.Enabled = False
''                    Cmb_Sucursal.ListIndex = 0
''                    cmbTipoCtaBen.ListIndex = 0
''                    cmbBancoCtaBen.ListIndex = 0
''                    txtNumctaBen.Text = ""
''                    txt_CCI.Text = ""
'                End If
'            End If
'        End If
    
    End If
End If
    
    
    
Exit Function
Err_CmbViaPagoClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select



'txt_CCI.Text = ""
'Txt_NumCta.Text = ""
'vlCombo = Trim(Mid(Cmb_ViaPago, 1, (InStr(1, Cmb_ViaPago, "-") - 1)))
'    If vlCombo = "01" Or vlCombo = "04" Then
'        If (vlCombo = "04") Then
'            vgTipoSucursal = cgTipoSucursalAfp
'        Else
'            vgTipoSucursal = cgTipoSucursalSuc
'        End If
'        fgComboSucursal Cmb_Suc, vgTipoSucursal
'
'        Cmb_TipCta.Enabled = False
'        Cmb_Bco.Enabled = False
'        Txt_NumCta.Enabled = False
'        Cmb_Suc.Enabled = True
'        If (vlCombo = "04") Then
'            vgPalabra = fgObtenerCodigo_TextoCompuesto(Lbl_Afp)
'            Call fgBuscaPos(Cmb_Suc, vgPalabra)
'        End If
'        If Cmb_TipCta.ListCount <> 0 Then
'            Cmb_TipCta.ListIndex = 0
'        End If
'        If Cmb_Bco.ListCount <> 0 Then
'            Cmb_Bco.ListIndex = 0
'        End If
'    Else
'        vgTipoSucursal = cgTipoSucursalSuc
'        fgComboSucursal Cmb_Suc, vgTipoSucursal
'
'        If vlCombo = "00" Or vlCombo = "05" Then
'            Cmb_TipCta.Enabled = False
'            Cmb_Bco.Enabled = False
'            Txt_NumCta.Enabled = False
'            Cmb_Suc.Enabled = False
'            If Cmb_TipCta.ListCount <> 0 Then
'                Cmb_TipCta.ListIndex = 0
'            End If
'            If Cmb_Bco.ListCount <> 0 Then
'                Cmb_Bco.ListIndex = 0
'            End If
'        Else
'            Cmb_TipCta.Enabled = True
'            Cmb_Bco.Enabled = True
'            Txt_NumCta.Enabled = True
'            Cmb_Suc.Enabled = False
'            If Cmb_Suc.ListCount <> 0 Then
'                Cmb_Suc.ListIndex = 0
'            End If
'        End If
'
'        If vlCombo = "02" Then
'            vlBTipoEnvio = True
'            Call fgComboGeneral("MP", Cmb_Suc)
'            lblTipmonCab.Visible = True
'            Cmb_Suc.Enabled = True
'        Else
'            fgComboSucursal Cmb_Suc, vgTipoSucursal
'            lblTipmonCab.Visible = False
'            Cmb_Suc.Enabled = False
'            Txt_NumCta = ""
'        End If
'    End If
'
'    'INICIO GCP - FRACTAL 13052019
'    If vlCombo = "06" Then
'             'Pago ventanilla debe quedar solo bancos princiales
'            Call DejarBcoPrinicipales(Cmb_Bco)
'            'Deshabilitar Sucursal, tipcta, nrcci
'            'Habilitar Banco, Nro Cuenta.
'            Cmb_Suc.Enabled = False
'            Cmb_TipCta.Enabled = False
'
'            txt_CCI.Enabled = False
'            Cmb_Bco.Enabled = True
'            Txt_NumCta.Enabled = False
'
'     ElseIf vlCombo = "05" Then
'
'            Cmb_Suc.Enabled = False
'            Cmb_TipCta.Enabled = False
'
'            txt_CCI.Enabled = False
'            Cmb_Bco.Enabled = True
'            Txt_NumCta.Enabled = False
'
'      Else
'        Cmb_Bco.Clear
'        Call fgComboGeneral(vgCodTabla_Bco, Cmb_Bco)
'    End If
'
'    'FIN GCP - FRACTAL 13052019
'
'
'
'
    
    

'  vlCombo = Trim(Mid(Cmb_ViaPago, 1, (InStr(1, Cmb_ViaPago, "-") - 1)))
'
'   Dim TP As String
'   TP = Left(Trim(Lbl_TipPen.Caption), 2)
'   If (TP <> "04" And TP <> "05") Then
'    If vlCombo = "02" Then
'        MsgBox "La forma de pago '02 - DEPOSITO EN BANCO' es exclusiva para pólizas de tipo de pensión JUBILACION legal o anticipada.", vbExclamation, "Error de Datos"
'        Cmb_ViaPago.ListIndex = 0
'        vlCombo = "00"
'     End If
'   End If
'
'    If vlCombo = "01" Or vlCombo = "04" Then
'        If (vlCombo = "04") Then
'            vgTipoSucursal = cgTipoSucursalAfp
'        Else
'            vgTipoSucursal = cgTipoSucursalSuc
'        End If
'        fgComboSucursal Cmb_Suc, vgTipoSucursal
'
'        Cmb_TipCta.Enabled = False
'        Cmb_Bco.Enabled = False
'        Txt_NumCta.Enabled = False
'        txt_CCI.Enabled = False
'        Cmb_Suc.Enabled = True
'        If (vlCombo = "04") Then
'            vgPalabra = fgObtenerCodigo_TextoCompuesto(Lbl_Afp)
'            Call fgBuscaPos(Cmb_Suc, vgPalabra)
'        End If
'        If Cmb_TipCta.ListCount <> 0 Then
'            Cmb_TipCta.ListIndex = 0
'        End If
'        If Cmb_Bco.ListCount <> 0 Then
'            Cmb_Bco.ListIndex = 0
'        End If
'    Else
'
'        vgTipoSucursal = cgTipoSucursalSuc
'        fgComboSucursal Cmb_Suc, vgTipoSucursal
'
'        If vlCombo = "00" Or vlCombo = "05" Then
'            Cmb_TipCta.Enabled = False
'            Cmb_Bco.Enabled = False
'            Txt_NumCta.Enabled = False
'            txt_CCI.Enabled = False
'            Cmb_Suc.Enabled = False
'            txt_CCI.Enabled = False
'            If Cmb_TipCta.ListCount <> 0 Then
'                Cmb_TipCta.ListIndex = 0
'            End If
'            If Cmb_Bco.ListCount <> 0 Then
'                Cmb_Bco.ListIndex = 0
'            End If
'        Else
'            Cmb_TipCta.Enabled = True
'            Cmb_Bco.Enabled = True
'            Txt_NumCta.Enabled = True
'            txt_CCI.Enabled = True
'
'            Cmb_Suc.Enabled = False
'            If Cmb_Suc.ListCount <> 0 Then
'                Cmb_Suc.ListIndex = 0
'            End If
'        End If
'
'        If vlCombo = "02" Then
'            vlBTipoEnvio = True
'            Call fgComboGeneral("MP", Cmb_Suc)
'            lblTipmonCab.Visible = True
'            Cmb_Suc.Enabled = True
'        Else
'            fgComboSucursal Cmb_Suc, vgTipoSucursal
'            lblTipmonCab.Visible = False
'            Cmb_Suc.Enabled = False
'            Txt_NumCta = ""
'            txt_CCI = ""
'        End If
'    End If
    
End Function

'---------------------------------------------------------------------------
'VALIDA QUE ESTEN LLENOS LOS DATOS DEL BENEFICIRAIO
'---------------------------------------------------------------------------
Function flValidaDatosAseg()
On Error GoTo Err_ValDatAseg

    flValidaDatosAseg = False
    
    If Trim(Cmb_TipoIdentBen) = "" Then
        MsgBox "Debe ingresar el Tipo de Identificación del Beneficiario.", vbExclamation, "Error de Datos"
        flValidaDatosAseg = True
        SSTab_Poliza.Tab = 2
        Cmb_TipoIdentBen.SetFocus
        Exit Function
    End If
    If Txt_NumIdentBen = "" Then
        MsgBox "Debe ingresar el Número de Identificación del Beneficiario.", vbExclamation, "Error de Datos"
        flValidaDatosAseg = True
        SSTab_Poliza.Tab = 2
        Txt_NumIdentBen.SetFocus
        Exit Function
    End If
    If Txt_NombresBen = "" Then
        MsgBox "Debe ingresar Nombre del Beneficiario.", vbExclamation, "Error de Datos"
        flValidaDatosAseg = True
        SSTab_Poliza.Tab = 2
        Txt_NombresBen.SetFocus
        Exit Function
    End If
'    If Txt_NombresBenSeg = "" Then
'        MsgBox "Debe ingresar el Segundo Nombre del Beneficiario.", vbExclamation, "Error de Datos"
'        flValidaDatosAseg = True
'        SSTab_Poliza.Tab = 2
'        Txt_NombresBenSeg.SetFocus
'        Exit Function
'    End If
    If Txt_ApPatBen = "" Then
        MsgBox "Debe ingresar Apellido Paterno del Beneficiario.", vbExclamation, "Error de Datos"
        flValidaDatosAseg = True
        SSTab_Poliza.Tab = 2
        Txt_ApPatBen.SetFocus
        Exit Function
    End If
'    If Txt_ApMatBen = "" Then
'        MsgBox "Debe ingresar Apellido Materno del Beneficiario.", vbExclamation, "Error de Datos"
'        flValidaDatosAseg = True
'        SSTab_Poliza.Tab = 2
'        Txt_ApMatBen.SetFocus
'        Exit Function
'    End If


    
    If Lbl_CauInvBen = "" Then
        MsgBox "Debe ingresar la Causa de Invalidez del Beneficiario.", vbExclamation, "Error de Datos"
        flValidaDatosAseg = True
        SSTab_Poliza.Tab = 2
        Cmd_BuscarCauInvBen.SetFocus
        Exit Function
    End If
    
    If Trim(Mid(Lbl_CauInvBen, 1, InStr(1, Lbl_CauInvBen, "-") - 1)) <> "0" Then
        If Not IsDate(Txt_FecInvBen) Then
            MsgBox "Debe ingresar la Fecha de Invalidez del Beneficiario.", vbExclamation, "Error de Datos"
            flValidaDatosAseg = True
            SSTab_Poliza.Tab = 2
            Txt_FecInvBen.SetFocus
            Exit Function
        End If
    Else
        Txt_FecInvBen = ""
    End If
    
    If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
        If Left(Trim(Lbl_Par.Caption), 2) <> "99" Then
            If Len(Trim(Txt_FecFallBen.Text)) > 0 Then
                If MsgBox("Seguro de Ingresar Fecha de Fallecimiento de Beneficio NO TITULAR.", vbQuestion + vbYesNo, "Confirmar") = vbNo Then
                    SSTab_Poliza.Tab = 2
                    Txt_FecFallBen.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
If Not Comprobar_Mail(Trim(Me.txtCorreoBen.Text)) Then
    
    MsgBox "El correo ingresado no es válido.", vbExclamation, "Error de Datos"
      flValidaDatosAseg = True
        SSTab_Poliza.Tab = 2
        Me.txtCorreoBen.SetFocus
        Exit Function
End If

Exit Function
Err_ValDatAseg:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Sub pGrabaRepresentante()
Dim vlTipoIdenRep As String

On Error GoTo ManejoError

    vlNumPol = Trim(Txt_NumPol)
        
    vlSql = ""
    vlSql = "DELETE FROM pd_tmae_oripolrep WHERE num_poliza= '" & vlNumPol & "'"
    vgConectarBD.Execute vlSql
    
    vlSql = ""
    vlSql = "INSERT INTO pd_tmae_oripolrep ("
    vlSql = vlSql & "num_poliza,cod_tipoidenrep,num_idenrep,gls_nombresrep,gls_apepatrep,"
    vlSql = vlSql & "gls_apematrep,cod_usuariocrea,fec_crea,hor_crea,"
    vlSql = vlSql & "GLS_TELREP1,GLS_TELREP2,GLS_CORREOREP, cod_area_telrep1, cod_area_telrep2, cod_sexo"
    vlSql = vlSql & ") values ("
    vlSql = vlSql & "'" & Trim(vlNumPol) & "', "
    vlTipoIdenRep = fgObtenerCodigo_TextoCompuesto(Cmb_TipIdRep)
    vlSql = vlSql & "'" & Trim(vlTipoIdenRep) & "', "
    vlSql = vlSql & "'" & Trim(Txt_NumIdRep.Text) & "', "
    vlSql = vlSql & "'" & Trim(Txt_NomRep.Text) & "', "
    vlSql = vlSql & "'" & Trim(Txt_ApPatRep.Text) & "', "
    vlSql = vlSql & "'" & Trim(Txt_ApMatRep.Text) & "', "
    vlSql = vlSql & "'" & vgUsuario & "',"
    vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
    vlSql = vlSql & "'" & Format(Time, "hhmmss") & "', "
    vlSql = vlSql & "'" & Trim(Me.txtTelRep1.Text) & "', "
    vlSql = vlSql & "'" & Trim(Me.txtTelRep2.Text) & "', "
    vlSql = vlSql & "'" & Trim(Me.txtCorreoRep.Text) & "',"
    vlSql = vlSql & "'" & Trim(DirRep.vCodigoTelefono) & "',"
    vlSql = vlSql & "'" & Trim(DirRep.vCodigoTelefono2) & "', "
    vlSql = vlSql & "'" & Trim(cmbSexoRep.Text) & "' "
    vlSql = vlSql & " ) "
    vgConectarBD.Execute (vlSql)
    
    If DirRep.vCodDireccion <> "" Then
 
        Call GrabarDireccionRepresentante(Trim(vlNumPol), CInt(fgObtenerCodigo_TextoCompuesto(Cmb_TipIdRep)), Trim(Txt_NumIdRep.Text), _
        Format(Date, "yyyymmdd"), DirRep.vTipoVia, DirRep.vDireccion, DirRep.vNumero, DirRep.vTipoBlock, DirRep.vNumBlock, DirRep.vTipoPref, DirRep.vInterior, _
        DirRep.vTipoConj, DirRep.vConjHabit, DirRep.vEtapa, DirRep.vManzana, DirRep.vLote, DirRep.vReferencia, 1, DirRep.vcodeDepar, DirRep.vcodeProv, _
        DirRep.vCodeDistr, DirRep.vgls_desdirebusq, vgUsuario, Format(Date, "yyyymmdd"), Format(Time, "hhmmss"), vgUsuario, DirRep.vCodDireccion)
    
    Else
        MsgBox "Debe ingresar la dirección del representante.", vbCritical
        Exit Sub
    
    End If
    
    Exit Sub
    
ManejoError:
      MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."

End Sub

Function flGrabaBeneficiario()

    vlNumPol = Trim(Txt_NumPol)
        
    vlSql = ""
    vlSql = "DELETE FROM pd_tmae_oripolben WHERE num_poliza= '" & vlNumPol & "'"
    vgConectarBD.Execute vlSql
    
   'Integracion GobiernoDeDatos_
   'KEVIN CORDOVA -- SE BORRA LA INFORMACION PARA LUEGO REGISTRARLO NUEVAMENTE
   vlSql = ""
   vlSql = "DELETE FROM PP_TMAE_BEN_DIRECCION WHERE num_poliza= '" & vlNumPol & "'"
   vgConectarBD.Execute vlSql
   
   vlSql = ""
   vlSql = "DELETE FROM PP_TMAE_BEN_TELEFONO WHERE num_poliza= '" & vlNumPol & "'"
   vgConectarBD.Execute vlSql
   'Fin Integracion GobiernoDeDatos_
    
    For vlJ = 1 To (Msf_GriAseg.rows - 1)
        Msf_GriAseg.Row = vlJ
        
'        vlNumOrden = 0
        vlRutGrilla = ""
        vlDgvGrilla = ""
        vlNomBen = ""
        vlNomBenSeg = ""
        vlPatBen = ""
        vlMatBen = ""
        
        vlCodPar = ""
        vlGruFam = ""
        vlSexoBen = ""
        vlSitInv = ""
        vlFecInv = ""
        vlCauInv = ""
        vlDerPen = ""
        vlDerCre = ""
        vlFecNacBen = ""
        vlFecFallBen = ""
        vlFecNacHM = ""
        vlPrcPension = ""
        vlPrcPensionLeg = ""
        vlPrcPensionRep = ""
        vlPenBen = ""
        vlPenGarBen = ""
        
        'INICIO GCP-FRACTAL 11042019
        vlnum_CCI = ""
        vl_Fono1_Ben = ""
        vl_Fono2_Ben = ""
        vl_ConTratDatos_Ben = ""
        vl_ConUsoDatosCom_Ben = ""
        vl_CorreoBen = ""
        'FIN GCP-FRACTAL 11042019
              

        'Obtener datos desde la grilla(pueden ser modificados)
        If vlBotonEscogido = "C" Then
            Msf_GriAseg.Col = 21
            vlNumOrden = Msf_GriAseg.Text
        Else
            Msf_GriAseg.Col = 0
            vlNumOrden = Msf_GriAseg.Text
        End If

'        If vlNumOrden = clNumOrden1 Then
            Msf_GriAseg.Col = 5
            vlFecInv = Format(Msf_GriAseg.Text, "yyyymmdd")
            Msf_GriAseg.Col = 6
            vlCauInv = Msf_GriAseg.Text
'        End If

        Msf_GriAseg.Col = 11
        vlRutGrilla = fgObtenerCodigo_TextoCompuesto(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 12
        vlDgvGrilla = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 13
        vlNomBen = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 14
        vlNomBenSeg = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 15
        vlPatBen = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 16
        vlMatBen = Trim(Msf_GriAseg.Text)
        
'I--- ABV 10/08/2007 ---
'Se deben Actualizar los siguientes campos calculados y registrados en la grilla
'Derecho a Crecer
        Msf_GriAseg.Col = 8
        If Trim(Msf_GriAseg.Text) <> "" Then
            vlDerCre = Trim(Msf_GriAseg.Text)
        Else
            vlDerCre = "N"
        End If
'I -Corrección => KVR me los entregará calculados
''Porcentaje Pensión a Pagar
'        Msf_GriAseg.Col = 17
'        vlPrcPension = Trim(Msf_GriAseg.Text)
''Monto de la Pensión Normal
'        Msf_GriAseg.Col = 18
'        vlPenBen = Trim(Msf_GriAseg.Text)
''Monto de la Pensión Garantizada
'        Msf_GriAseg.Col = 19
'        vlPenGarBen = Trim(Msf_GriAseg.Text)
'F -Corrección => KVR me los entregará calculados

'Estado de Pago de la Pensión
        Msf_GriAseg.Col = 22
        vlEstPen = Trim(Msf_GriAseg.Text)
'Porcentaje Pensión Garantizada
        Msf_GriAseg.Col = 23
        vlPrcPensionGar = Trim(Msf_GriAseg.Text)
'F--- ABV 10/08/2007 ---
        
        'RRR 26/12/2013
        Msf_GriAseg.Col = 25
        vlcod_tipcta = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 26
        vlcod_monbco = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 27
        vlCod_Banco = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 28
        vlnum_ctabco = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 29
        vl_nacionalidadben = Trim(Msf_GriAseg.Text)
        
        
        'INICIO GCP-FRACTAL 11042019
        Msf_GriAseg.Col = 31
        vlnum_CCI = Trim(Msf_GriAseg.Text)
   
        
        Msf_GriAseg.Col = 32
        vl_Fono1_Ben = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 33
        vl_Fono2_Ben = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 34
        vl_ConTratDatos_Ben = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 35
        vl_ConUsoDatosCom_Ben = Trim(Msf_GriAseg.Text)
        'FIN GCP-FRACTAL 11042019
        
        Msf_GriAseg.Col = 36
        vl_CorreoBen = Trim(Msf_GriAseg.Text)
        
        
 
If (vgBotonEscogido <> "R") Then
        'Obtener datos desde tabla de cotizacion (no pueden ser modificados)
        vlSql = ""
        vlSql = "SELECT num_orden,cod_par,cod_grufam,cod_sexo,"
        vlSql = vlSql & "cod_sitinv," 'fec_invben,cod_cauinv,"
        vlSql = vlSql & "cod_derpen,cod_dercre,fec_nacben,"
        vlSql = vlSql & "fec_fallben,fec_nachm,prc_pension,"
        vlSql = vlSql & "prc_pensionleg,prc_pensionrep,"
        vlSql = vlSql & "mto_pension,mto_pensiongar "
        vlSql = vlSql & "FROM pd_tmae_oripolben "
        vlSql = vlSql & "WHERE num_poliza= '" & vlNumPol & "' AND "
        vlSql = vlSql & "num_orden = '" & vlNumOrden & "'"
        Set vgRs = vgConexionBD.Execute(vlSql)
        If Not vgRs.EOF Then
'            vlNumPol
'            vlNumOrden
'            vlRutGrilla
'            vldgvgrilla
'            vlNomBen
'            vlPatBen
'            vlMatBen
            vlNumOrden = (vgRs!Num_Orden)
            vlCodPar = Trim(vgRs!Cod_Par)
            vlGruFam = Trim(vgRs!Cod_GruFam)
            vlSexoBen = Trim(vgRs!Cod_Sexo)
            vlSitInv = Trim(vgRs!Cod_SitInv)
'            If Not IsNull(vgRs!fec_invben) Then
'                vlFecInv = Trim(vgRs!fec_invben)
'            Else
'                vlFecInv = ""
'            End If
'            If vlNumOrden <> clNumOrden1 Then
'                If Not IsNull(vgRs!cod_cauinv) Then
'                    vlCauInv = Trim(vgRs!cod_cauinv)
'                Else
'                    vlCauInv = ""
'                End If
'            End If
            If Not IsNull(vgRs!Cod_DerPen) Then
                vlDerPen = Trim(vgRs!Cod_DerPen)
            Else
                vlDerPen = ""
            End If
            vlFecNacBen = Trim(vgRs!Fec_NacBen)
            If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
                Msf_GriAseg.Col = 20
                If Trim(Msf_GriAseg.Text) <> "" Then
                    vlFecFallBen = Mid(Msf_GriAseg.Text, 7, 4) & Mid(Msf_GriAseg.Text, 4, 2) & Mid(Msf_GriAseg.Text, 1, 2)
                Else
                    vlFecFallBen = ""
                End If
            Else
                If Not IsNull(vgRs!Fec_FallBen) Then
                    vlFecFallBen = Trim(vgRs!Fec_FallBen)
                Else
                    vlFecFallBen = ""
                End If
            End If
            If Not IsNull(vgRs!Fec_NacHM) Then
                vlFecNacHM = Trim(vgRs!Fec_NacHM)
            Else
                vlFecNacHM = ""
            End If
            vlPrcPensionLeg = Trim(vgRs!Prc_PensionLeg)
            vlPrcPensionRep = Trim(vgRs!Prc_PensionRep)

'I--- ABV 10/08/2007 ---
'Se van a obtener desde la Grilla
'Corrección => KVR me los entregará calculados
            vlPrcPension = Trim(vgRs!Prc_Pension)
            vlPenBen = Trim(vgRs!Mto_Pension)
            vlPenGarBen = Trim(vgRs!Mto_PensionGar)
'            If Not IsNull(vgRs!Cod_DerCre) Then
'                vlDerCre = Trim(vgRs!Cod_DerCre)
'            Else
'                vlDerCre = ""
'            End If
'F--- ABV 10/08/2007 ---
        End If
Else
    'Obtener los datos desde la Estructura de Beneficiarios
        Msf_GriAseg.Col = 1
        vlCodPar = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 2
        vlGruFam = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 3
        vlSexoBen = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 4
        vlSitInv = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 7
        If Trim(Msf_GriAseg.Text) <> "" Then
            vlDerPen = Trim(Msf_GriAseg.Text)
        Else
            vlDerPen = "10"
        End If
        Msf_GriAseg.Col = 9
        vlFecNacBen = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
        Msf_GriAseg.Col = 20
        If Trim(Msf_GriAseg.Text) <> "" Then
            vlFecFallBen = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
        Else
            vlFecFallBen = ""
        End If
        Msf_GriAseg.Col = 10
        If Trim(Msf_GriAseg.Text) <> "" Then
            vlFecNacHM = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
        Else
            vlFecNacHM = ""
        End If
        
        Msf_GriAseg.Col = 17
        vlPrcPension = Trim(Msf_GriAseg.Text)
        vlPrcPensionRep = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 24
        vlPrcPensionLeg = Trim(Msf_GriAseg.Text)

        Msf_GriAseg.Col = 18
        vlPenBen = Trim(Msf_GriAseg.Text)

        Msf_GriAseg.Col = 19
        vlPenGarBen = Trim(Msf_GriAseg.Text)
        
        'RRR 26/12/2013
        Msf_GriAseg.Col = 25
        vlcod_tipcta = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 26
        vlcod_monbco = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 27
        vlCod_Banco = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 28
        vlnum_ctabco = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 29
        vl_nacionalidadben = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 31
        vlnum_CCI = Trim(Msf_GriAseg.Text)
        
End If

vlNumCtaCCI = Trim(txt_CCI)
vl_Fono = Trim(Txt_Fono)
vl_Fono2_Afil = Trim(Txt_Fono2_Afil)
VlConTratDatos_Afil = Trim(chkConTratDatos_Afil.Value)
VlConUsoDatosCom_Afil = Trim(chkConUsoDatosCom_Afil.Value)

''Borrar Después
'        Msf_GriAseg.Col = 1
'        vlCodPar = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 2
'        vlGruFam = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 3
'        vlSexoBen = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 4
'        vlSitInv = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 7
'        If Trim(Msf_GriAseg.Text) <> "" Then
'            vlDerPen = Trim(Msf_GriAseg.Text)
'        Else
'            vlDerPen = "10"
'        End If
'        Msf_GriAseg.Col = 8
'        If Trim(Msf_GriAseg.Text) <> "" Then
'            vlDerCre = Trim(Msf_GriAseg.Text)
'        Else
'            vlDerCre = "N"
'        End If
'        Msf_GriAseg.Col = 9
'        vlFecNacBen = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
'        Msf_GriAseg.Col = 20
'        If Trim(Msf_GriAseg.Text) <> "" Then
'            vlFecFallBen = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
'        Else
'            vlFecFallBen = ""
'        End If
'        Msf_GriAseg.Col = 10
'        If Trim(Msf_GriAseg.Text) <> "" Then
'            vlFecNacHM = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
'        Else
'            vlFecNacHM = ""
'        End If
'        Msf_GriAseg.Col = 17
'        vlPrcPension = Trim(Msf_GriAseg.Text)
'        vlPrcPensionLeg = Trim(Msf_GriAseg.Text)
'        vlPrcPensionRep = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 18
'        vlPenBen = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 19
'        vlPenGarBen = Trim(Msf_GriAseg.Text)
''Borrar Después


'Integracion GobiernoDeDatos_
    
   'INSERTAR EN LA TABLA PP_TMAE_BEN_TELEFONO
            vlSql = "INSERT INTO PP_TMAE_BEN_TELEFONO "
            vlSql = vlSql + "(NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN,cod_tipoidenben,NUM_IDENBEN,FEC_INGRESO,"
            vlSql = vlSql + "COD_TIPO_FONOBEN,COD_AREA_FONOBEN,GLS_FONOBEN,cod_tipo_telben2,cod_area_Telben2,gls_telben2,COD_USUARIOCREA,"
            vlSql = vlSql + "FEC_CREA,HOR_CREA)"
            vlSql = vlSql + "VALUES("
            vlSql = vlSql & "'" & Trim(Txt_NumPol) & "', "
            vlSql = vlSql & "1,"
            Msf_GriAseg.Col = 0
            vlSql = vlSql & " " & str(Msf_GriAseg.Text) & ", "
            vlSql = vlSql & " " & str(vlRutGrilla) & ", "
            vlSql = vlSql & "'" & Trim(vlDgvGrilla) & "', "
            vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
            
            Dim NumOrden  As String
            Msf_GriAseg.Col = 0
            NumOrden = Msf_GriAseg.Text
            
     If NumOrden = 1 Then
        
            If Trim(pTipoTelefono) <> "" Then
            vlSql = vlSql & Trim(pTipoTelefono) & ", "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pCodigoTelefono) <> "" Then
            vlSql = vlSql & Trim(pCodigoTelefono) & ", "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pNumTelefono) <> "" Then
            vlSql = vlSql & "'" & Trim(pNumTelefono) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pTipoTelefono2) <> "" Then
            vlSql = vlSql & Trim(pTipoTelefono2) & ", "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pCodigoTelefono2) <> "" Then
            vlSql = vlSql & Trim(pCodigoTelefono2) & ", "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pNumTelefono2) <> "" Then
            vlSql = vlSql & "'" & Trim(pNumTelefono2) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            
        Else
                vlSql = vlSql & "'4',"
                vlSql = vlSql & "'1',"
                vlSql = vlSql & "'" & Trim(vl_Fono1_Ben) & "', "
                vlSql = vlSql & "'2',"
                vlSql = vlSql & "NULL,"
                vlSql = vlSql & "'" & Trim(vl_Fono2_Ben) & "', "
          End If
        
            
            vlSql = vlSql & "'" & vgUsuario & "',"
            vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
            vlSql = vlSql & "'" & Format(Time, "hhmmss") & "') "
            
            
        vgConectarBD.Execute (vlSql)
            
   'INSERTAR EN LA TABLA MAE_BEN_DIRECCION
   
            vlSql = "INSERT INTO PP_TMAE_BEN_DIRECCION "
            vlSql = vlSql + "(NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN,cod_tipoidenben,NUM_IDENBEN,FEC_INGRESO"
            vlSql = vlSql & ",COD_DIRE_VIA,"
            vlSql = vlSql + "GLS_DIRECCION,NUM_DIRECCION,COD_BLOCKCHALET,GLS_BLOCKCHALET,COD_INTERIOR,"
            vlSql = vlSql + "NUM_INTERIOR,COD_CJHT,GLS_NOM_CJHT,GLS_ETAPA,GLS_MANZANA,GLS_LOTE,"
            vlSql = vlSql + "GLS_REFERENCIA,COD_PAIS,COD_DEPARTAMENTO,COD_PROVINCIA,COD_DISTRITO,"
            vlSql = vlSql + "GLS_DESDIREBUSQ,COD_USUARIOCREA,FEC_CREA,HOR_CREA)"
            vlSql = vlSql + "VALUES("
            
            vlSql = vlSql & "'" & Trim(Txt_NumPol) & "', "
            vlSql = vlSql & "1,"
            Msf_GriAseg.Col = 0
            vlSql = vlSql & " " & str(Msf_GriAseg.Text) & ", "
            vlSql = vlSql & " " & str(vlRutGrilla) & ", "
            vlSql = vlSql & "'" & Trim(vlDgvGrilla) & "', "
            vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
            
            If Trim(pTipoVia) <> "" Then
            vlSql = vlSql & "'" & Trim(pTipoVia) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pDireccion) <> "" Then
            vlSql = vlSql & "'" & Trim(pDireccion) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pNumero) <> "" Then
            vlSql = vlSql & "'" & Trim(pNumero) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pTipoBlock) <> "" Then
            vlSql = vlSql & "'" & Trim(pTipoBlock) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pNumBlock) <> "" Then
            vlSql = vlSql & "'" & Trim(pNumBlock) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pTipoPref) <> "" Then
            vlSql = vlSql & "'" & Trim(pTipoPref) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pInterior) <> "" Then
            vlSql = vlSql & "'" & Trim(pInterior) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pTipoConj) <> "" Then
            vlSql = vlSql & "'" & Trim(pTipoConj) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pConjHabit) <> "" Then
            vlSql = vlSql & "'" & Trim(pConjHabit) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pEtapa) <> "" Then
            vlSql = vlSql & "'" & Trim(pEtapa) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pManzana) <> "" Then
            vlSql = vlSql & "'" & Trim(pManzana) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pLote) <> "" Then
            vlSql = vlSql & "'" & Trim(pLote) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pReferencia) <> "" Then
            vlSql = vlSql & "'" & Trim(pReferencia) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(fgObtenerCodigo_TextoCompuesto(cboNacionalidad)) <> "" Then
            vlSql = vlSql & "'" & Trim(fgObtenerCodigo_TextoCompuesto(cboNacionalidad)) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(vgCodigoRegion) <> "" Then
            vlSql = vlSql & "'" & Trim(vgCodigoRegion) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            If Trim(vgCodigoProvincia) <> "" Then
            vlSql = vlSql & "'" & Trim(vgCodigoProvincia) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            If Trim(vgCodigoComuna) <> "" Then
            vlSql = vlSql & "'" & Trim(vgCodigoComuna) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(Lbl_Dir) <> "" Then
            vlSql = vlSql & "'" & Trim(vDireccionConcat) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            vlSql = vlSql & "'" & vgUsuario & "',"
            vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
            vlSql = vlSql & "'" & Format(Time, "hhmmss") & "') "
            
            vgConectarBD.Execute (vlSql)
     'Fin Integracion GobiernoDeDatos_
        
       'SI ES UN NUEVO REGISTRO
        vlSql = ""
        vlSql = "INSERT INTO pd_tmae_oripolben ("
        vlSql = vlSql & "num_poliza,num_orden,cod_par,cod_grufam,"
        vlSql = vlSql & "cod_sexo,cod_sitinv,fec_invben,cod_cauinv,"
        vlSql = vlSql & "cod_derpen,cod_dercre,cod_tipoidenben,num_idenben,"
        vlSql = vlSql & "gls_nomben,gls_nomsegben,gls_patben,gls_matben,fec_nacben,"
        vlSql = vlSql & "fec_fallben,fec_nachm,prc_pension,prc_pensionleg,"
        vlSql = vlSql & "prc_pensionrep,mto_pension,"
        vlSql = vlSql & "mto_pensiongar,cod_usuariocrea,fec_crea,hor_crea"
        vlSql = vlSql & ",cod_estpension,prc_pensiongar "
        'RRR 26/12/2013
        If vlcod_tipcta <> "" Then vlSql = vlSql & ", cod_tipcta "
        If vlcod_monbco <> "" Then vlSql = vlSql & ", cod_monbco "
        If vlCod_Banco <> "" Then vlSql = vlSql & ", Cod_Banco "
        If vlnum_ctabco <> "" Then vlSql = vlSql & ", num_ctabco "
        
        'INICIO GCP-FRACTAL 11042019
        vlSql = vlSql & ", cod_nacionalidad, num_cuenta_cci "
        vlSql = vlSql & ", GLS_FONO, GLS_FONO2, CONS_TRAINFO, CONS_DATCOMER,GLS_CORREOBEN "
        'FIN GCP-FRACTAL 11042019
        
        'RRR
        vlSql = vlSql & ") values ("
        vlSql = vlSql & "'" & Trim(vlNumPol) & "', "
        vlSql = vlSql & " " & str(vlNumOrden) & ", "
        vlSql = vlSql & "'" & Trim(vlCodPar) & "', "
        vlSql = vlSql & "'" & Trim(vlGruFam) & "', "
        vlSql = vlSql & "'" & Trim(vlSexoBen) & "', "
        vlSql = vlSql & "'" & Trim(vlSitInv) & "', "
        If vlFecInv <> "" Then
            vlSql = vlSql & "'" & Trim(vlFecInv) & "', "
        Else
            vlSql = vlSql & "NULL,"
        End If
        If Trim(vlCauInv) <> "" Then
            vlSql = vlSql & "'" & Trim(vlCauInv) & "', "
        Else
            vlSql = vlSql & "'0',"
        End If
        vlSql = vlSql & "'" & Trim(vlDerPen) & "', "
        vlSql = vlSql & "'" & Trim(vlDerCre) & "', "
        vlSql = vlSql & " " & str(vlRutGrilla) & ", "
        vlSql = vlSql & "'" & Trim(vlDgvGrilla) & "', "
        vlSql = vlSql & "'" & Trim(vlNomBen) & "', "
        If Trim(vlNomBenSeg) <> "" Then
            vlSql = vlSql & "'" & Trim(vlNomBenSeg) & "', "
        Else
            vlSql = vlSql & "NULL,"
        End If
        vlSql = vlSql & "'" & Trim(vlPatBen) & "', "
        If Trim(vlMatBen) <> "" Then
            vlSql = vlSql & "'" & Trim(vlMatBen) & "', "
        Else
            vlSql = vlSql & "NULL,"
        End If
        vlSql = vlSql & "'" & Trim(vlFecNacBen) & "', "
        If Trim(vlFecFallBen) <> "" Then
            vlSql = vlSql & "'" & Trim(vlFecFallBen) & "', "
        Else
            vlSql = vlSql & " NULL,"
        End If
        If Trim(vlFecNacHM) <> "" Then
            vlSql = vlSql & "'" & Trim(vlFecNacHM) & "', "
        Else
            vlSql = vlSql & " NULL,"
        End If
        vlSql = vlSql & " " & str(vlPrcPension) & ", "
        vlSql = vlSql & " " & str(vlPrcPensionLeg) & ", "
        vlSql = vlSql & " " & str(vlPrcPensionRep) & ", "
        vlSql = vlSql & " " & str(vlPenBen) & ", "
        vlSql = vlSql & " " & str(vlPenGarBen) & ","
        vlSql = vlSql & "'" & vgUsuario & "',"
        vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
        vlSql = vlSql & "'" & Format(Time, "hhmmss") & "' "
'I--- ABV 07/08/2007 ---
        vlSql = vlSql & ",'" & Trim(vlEstPen) & "' "
        vlSql = vlSql & "," & str(vlPrcPensionGar) & " "
'F--- ABV 07/08/2007 ---
        'RRR 26/12/2013
        If vlcod_tipcta <> "" Then vlSql = vlSql & ",'" & Trim(vlcod_tipcta) & "' "
        If vlcod_monbco <> "" Then vlSql = vlSql & ",'" & Trim(vlcod_monbco) & "' "
        If vlCod_Banco <> "" Then vlSql = vlSql & ",'" & Trim(vlCod_Banco) & "' "
        If vlnum_ctabco <> "" Then vlSql = vlSql & ",'" & Trim(vlnum_ctabco) & "' "
        vlSql = vlSql & ",'" & vl_nacionalidadben & "' "
        If vlCodPar = "99" Then
            vlSql = vlSql & ",'" & vlNumCtaCCI & "' "
            vlSql = vlSql & ",'" & vl_Fono & "' "
            vlSql = vlSql & ",'" & vl_Fono2_Afil & "' "
            vlSql = vlSql & ",'" & VlConTratDatos_Afil & "' "
            vlSql = vlSql & ",'" & VlConUsoDatosCom_Afil & "' "
        Else
          vlSql = vlSql & ",'" & vlnum_CCI & "' "
          vlSql = vlSql & ",'" & vl_Fono1_Ben & "' "
          vlSql = vlSql & ",'" & vl_Fono2_Ben & "' "
          vlSql = vlSql & ",'" & vl_ConTratDatos_Ben & "' "
          vlSql = vlSql & ",'" & vl_ConUsoDatosCom_Ben & "' "
       End If
          
          vlSql = vlSql & ",'" & vl_CorreoBen & "' "
     
     
        'RRR
        vlSql = vlSql & ") "
        vgConectarBD.Execute (vlSql)
    
Next vlJ
    
End Function

Function flGrabaBeneficiarioCot()


    vlNumPol = Trim(Txt_NumPol)
        
    vlSql = ""
    vlSql = "DELETE FROM pd_tmae_oripolben WHERE num_poliza= '" & vlNumPol & "'"
    vgConectarBD.Execute vlSql
    
    'INICIO GCP-FRACTAL 11042019
    vl_Fono = Trim(Txt_Fono)
    vl_Fono2_Afil = Trim(Txt_Fono2_Afil)
    VlConTratDatos_Afil = Trim(chkConTratDatos_Afil)
    VlConUsoDatosCom_Afil = Trim(chkConUsoDatosCom_Afil)
    'FIN GCP-FRACTAL 11042019
    
    vl_Correo_afil = Trim(Txt_Correo)
       
        
     
    For vlJ = 1 To (Msf_GriAseg.rows - 1)
        Msf_GriAseg.Row = vlJ
        
'        vlNumOrden = 0
        vlRutGrilla = ""
        vlDgvGrilla = ""
        vlNomBen = ""
        vlNomBenSeg = ""
        vlPatBen = ""
        vlMatBen = ""
        
        vlCodPar = ""
        vlGruFam = ""
        vlSexoBen = ""
        vlSitInv = ""
        vlFecInv = ""
        vlCauInv = ""
        vlDerPen = ""
        vlDerCre = ""
        vlFecNacBen = ""
        vlFecFallBen = ""
        vlFecNacHM = ""
        vlPrcPension = ""
        vlPrcPensionLeg = ""
        vlPrcPensionRep = ""
        vlPenBen = ""
        vlPenGarBen = ""
        vlEstPen = ""
        vlcod_tipcta = ""
        vlcod_monbco = ""
        vlCod_Banco = ""
        vlnum_ctabco = ""
        vlnum_CCI = ""
        'INICIO GCP-FRACTAL 11042019
        vl_Fono1_Ben = ""
        vl_Fono2_Ben = ""
        vl_ConTratDatos_Ben = ""
        vl_ConUsoDatosCom_Ben = ""
        
        'FIN GCP-FRACTAL 11042019
        
      

        'Obtener datos desde la grilla(pueden ser modificados)
        If vlBotonEscogido = "C" Then
            Msf_GriAseg.Col = 21
            vlNumOrden = Msf_GriAseg.Text
        Else
            Msf_GriAseg.Col = 0
            vlNumOrden = Msf_GriAseg.Text
        End If
        
'        If vlNumOrden = clNumOrden1 Then
            Msf_GriAseg.Col = 5
            vlFecInv = Msf_GriAseg.Text
            Msf_GriAseg.Col = 6
            vlCauInv = Msf_GriAseg.Text
'        End If
        
        Msf_GriAseg.Col = 11
        vlRutGrilla = fgObtenerCodigo_TextoCompuesto(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 12
        vlDgvGrilla = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 13
        vlNomBen = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 14
        vlNomBenSeg = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 15
        vlPatBen = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 16
        vlMatBen = Trim(Msf_GriAseg.Text)
        
'I--- ABV 10/08/2007 ---
'Se deben Actualizar los siguientes campos calculados y registrados en la grilla
'Derecho a Crecer
        Msf_GriAseg.Col = 8
        If Trim(Msf_GriAseg.Text) <> "" Then
            vlDerCre = Trim(Msf_GriAseg.Text)
        Else
            vlDerCre = "N"
        End If
'I -Corrección => KVR me los entregará calculados
''Porcentaje Pensión a Pagar
'        Msf_GriAseg.Col = 17
'        vlPrcPension = Trim(Msf_GriAseg.Text)
''Monto de la Pensión Normal
'        Msf_GriAseg.Col = 18
'        vlPenBen = Trim(Msf_GriAseg.Text)
''Monto de la Pensión Garantizada
'        Msf_GriAseg.Col = 19
'        vlPenGarBen = Trim(Msf_GriAseg.Text)
'F -Corrección => KVR me los entregará calculados

'Estado de Pago de la Pensión
        Msf_GriAseg.Col = 22
        vlEstPen = Trim(Msf_GriAseg.Text)
'Porcentaje Pensión Garantizada
        Msf_GriAseg.Col = 23
        vlPrcPensionGar = Trim(Msf_GriAseg.Text)
'F--- ABV 10/08/2007 ---
        
        'RRR 26/12/2013
        Msf_GriAseg.Col = 25
        vlcod_tipcta = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 26
        vlcod_monbco = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 27
        vlCod_Banco = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 28
        vlnum_ctabco = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 29
        vl_nacionalidadben = Trim(Msf_GriAseg.Text)
        
        
        'INICIO GCP-FRACTAL 11042019
        
        Msf_GriAseg.Col = 31
        vlnum_CCI = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 32
        vl_Fono1_Ben = Trim(Msf_GriAseg.Text)
        
        
        Msf_GriAseg.Col = 33
        vl_Fono2_Ben = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 34
        vl_ConTratDatos_Ben = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 35
        vl_ConUsoDatosCom_Ben = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 36
       vl_Correo_Ben = Trim(Msf_GriAseg.Text)
        
        'FIN GCP-FRACTAL 11042019
        
        
If (vgBotonEscogido <> "R") Then
        'Obtener datos desde tabla de cotizacion (no pueden ser modificados)
        vlSql = ""
        vlSql = "SELECT cod_par,cod_grufam,cod_sexo,"
        vlSql = vlSql & "cod_sitinv," 'fec_sitinv,cod_cauinv,"
        vlSql = vlSql & "cod_derpen,cod_dercre,fec_nacben,"
        vlSql = vlSql & "fec_falben,fec_nachm,prc_pension,prc_pensionleg,prc_pensionrep,"
        vlSql = vlSql & "mto_pension,mto_pensiongar "
'I--- ABV 04/12/2009 ---
        vlSql = vlSql & ",prc_pensionsobdif "
'F--- ABV 04/12/2009 ---
        vlSql = vlSql & "FROM pt_tmae_cotben "
        vlSql = vlSql & "WHERE num_cot = '" & vlNumCot & "' AND "
        vlSql = vlSql & "num_orden = '" & vlNumOrden & "'"
        Set vgRs = vgConexionBD.Execute(vlSql)
        If Not vgRs.EOF Then
'            vlNumPol
'            vlNumOrden
'            vlRutGrilla
'            vldgvgrilla
'            vlNomBen
'            vlPatBen
'            vlMatBen
            vlCodPar = Trim(vgRs!Cod_Par)
            vlGruFam = Trim(vgRs!Cod_GruFam)
            vlSexoBen = Trim(vgRs!Cod_Sexo)
            vlSitInv = Trim(vgRs!Cod_SitInv)
'            If Not IsNull(vgRs!fec_sitinv) Then
'                vlFecInv = Trim(vgRs!fec_sitinv)
'            Else
'                vlFecInv = ""
'            End If
'            If vlNumOrden <> clNumOrden1 Then
'                If Not IsNull(vgRs!cod_cauinv) Then
'                    vlCauInv = Trim(vgRs!cod_cauinv)
'                Else
'                    vlCauInv = ""
'                End If
'            End If
            If Not IsNull(vgRs!Cod_DerPen) Then
                vlDerPen = Trim(vgRs!Cod_DerPen)
            Else
                vlDerPen = "10"
            End If
            vlFecNacBen = Trim(vgRs!Fec_NacBen)
            'DC 20091125
            If Not IsNull(vgRs!fec_falben) Then
                If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
                    Msf_GriAseg.Col = 20
                    If Trim(Msf_GriAseg.Text) <> "" Then
                        vlFecFallBen = Mid(Msf_GriAseg.Text, 7, 4) & Mid(Msf_GriAseg.Text, 4, 2) & Mid(Msf_GriAseg.Text, 1, 2)
                    Else
                        vlFecFallBen = ""
                    End If
                Else
                    vlFecFallBen = Trim(vgRs!fec_falben)
                End If
            Else
                vlFecFallBen = ""
            End If
            If Not IsNull(vgRs!Fec_NacHM) Then
                vlFecNacHM = Trim(vgRs!Fec_NacHM)
            Else
                vlFecNacHM = ""
            End If
            
'I--- ABV 04/12/2009 ---
            If (vlMarcaSobDif = "S") Then
                vlPrcPensionLeg = Trim(vgRs!prc_pensionsobdif)
                vlPrcPensionRep = Trim(vgRs!prc_pensionsobdif)
                vlPrcPension = Trim(vgRs!prc_pensionsobdif)
            Else
                vlPrcPensionLeg = Trim(vgRs!Prc_PensionLeg)
                vlPrcPensionRep = Trim(vgRs!Prc_PensionRep)
                vlPrcPension = Trim(vgRs!Prc_Pension)
            End If
'F--- ABV 04/12/2009 ---

'I--- ABV 10/08/2007 ---
'Se van a obtener desde la Grilla
'Corrección => KVR me los entregará calculados
            
'I--- ABV 04/12/2009 ---
'            vlPrcPension = Trim(vgRs!Prc_Pension) 'Lo asigne arriba
'F--- ABV 04/12/2009 ---

            vlPenBen = Trim(vgRs!Mto_Pension)
            vlPenGarBen = Trim(vgRs!Mto_PensionGar)
'            If Not IsNull(vgRs!Cod_DerCre) Then
'                vlDerCre = Trim(vgRs!Cod_DerCre)
'            Else
'                vlDerCre = "N"
'            End If
'F--- ABV 10/08/2007 ---
        End If
Else
    'Obtener los datos desde la Estructura de Beneficiarios
        Msf_GriAseg.Col = 1
        vlCodPar = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 2
        vlGruFam = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 3
        vlSexoBen = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 4
        vlSitInv = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 7
        If Trim(Msf_GriAseg.Text) <> "" Then
            vlDerPen = Trim(Msf_GriAseg.Text)
        Else
            vlDerPen = "10"
        End If
        Msf_GriAseg.Col = 9
        vlFecNacBen = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
        Msf_GriAseg.Col = 20
        If Trim(Msf_GriAseg.Text) <> "" Then
            vlFecFallBen = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
        Else
            vlFecFallBen = ""
        End If
        Msf_GriAseg.Col = 10
        If Trim(Msf_GriAseg.Text) <> "" Then
            vlFecNacHM = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
        Else
            vlFecNacHM = ""
        End If
        
        Msf_GriAseg.Col = 17
        vlPrcPension = Trim(Msf_GriAseg.Text)
        vlPrcPensionRep = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 24
        vlPrcPensionLeg = Trim(Msf_GriAseg.Text)

        Msf_GriAseg.Col = 18
        vlPenBen = Trim(Msf_GriAseg.Text)

        Msf_GriAseg.Col = 19
        vlPenGarBen = Trim(Msf_GriAseg.Text)
        '
        'RRR 26/12/2013
        Msf_GriAseg.Col = 25
        vlcod_tipcta = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 26
        vlcod_monbco = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 27
        vlCod_Banco = Trim(Msf_GriAseg.Text)
        Msf_GriAseg.Col = 28
        vlnum_ctabco = Trim(Msf_GriAseg.Text)
        
        Msf_GriAseg.Col = 29
        vlnum_ctabco = Trim(Msf_GriAseg.Text)
End If
        
''Borrar Después
'        Msf_GriAseg.Col = 1
'        vlCodPar = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 2
'        vlGruFam = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 3
'        vlSexoBen = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 4
'        vlSitInv = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 7
'        If Trim(Msf_GriAseg.Text) <> "" Then
'            vlDerPen = Trim(Msf_GriAseg.Text)
'        Else
'            vlDerPen = "10"
'        End If
'        Msf_GriAseg.Col = 8
'        If Trim(Msf_GriAseg.Text) <> "" Then
'            vlDerCre = Trim(Msf_GriAseg.Text)
'        Else
'            vlDerCre = "N"
'        End If
'        Msf_GriAseg.Col = 9
'        vlFecNacBen = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
'        Msf_GriAseg.Col = 20
'        If Trim(Msf_GriAseg.Text) <> "" Then
'            vlFecFallBen = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
'        Else
'            vlFecFallBen = ""
'        End If
'        Msf_GriAseg.Col = 10
'        If Trim(Msf_GriAseg.Text) <> "" Then
'            vlFecNacHM = Format(CDate(Msf_GriAseg.Text), "yyyymmdd")
'        Else
'            vlFecNacHM = ""
'        End If
'        Msf_GriAseg.Col = 17
'        vlPrcPension = Trim(Msf_GriAseg.Text)
'        vlPrcPensionLeg = Trim(Msf_GriAseg.Text)
'        vlPrcPensionRep = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 18
'        vlPenBen = Trim(Msf_GriAseg.Text)
'        Msf_GriAseg.Col = 19
'        vlPenGarBen = Trim(Msf_GriAseg.Text)
''Borrar Después

        'Integracion GobiernoDeDatos_
        'INSERTAR EN LA TABLA PP_TMAE_BEN_TELEFONO
            vlSql = "INSERT INTO PP_TMAE_BEN_TELEFONO "
            vlSql = vlSql + "(NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN,cod_tipoidenben,NUM_IDENBEN,FEC_INGRESO,"
            vlSql = vlSql + "cod_tipo_fonoben,cod_area_fonoben,gls_fonoben,cod_tipo_telben2,cod_area_telben2,gls_telben2,COD_USUARIOCREA,"
            vlSql = vlSql + "FEC_CREA,HOR_CREA)"
            vlSql = vlSql + "VALUES("
            vlSql = vlSql & "'" & Trim(Txt_NumPol) & "', "
            vlSql = vlSql & "1,"
            Msf_GriAseg.Col = 0
            vlSql = vlSql & " " & str(Msf_GriAseg.Text) & ", "
            vlSql = vlSql & " " & str(vlRutGrilla) & ", "
            vlSql = vlSql & "'" & Trim(vlDgvGrilla) & "', "
            vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
            
            Dim NumOrden As String
            Msf_GriAseg.Col = 0
            NumOrden = Msf_GriAseg.Text
                        
           If NumOrden = 1 Then
        
            If Trim(pTipoTelefono) <> "" Then
            vlSql = vlSql & Trim(pTipoTelefono) & ", "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pCodigoTelefono) <> "" Then
            vlSql = vlSql & Trim(pCodigoTelefono) & ", "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pNumTelefono) <> "" Then
            vlSql = vlSql & "'" & Trim(pNumTelefono) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pTipoTelefono2) <> "" Then
            vlSql = vlSql & Trim(pTipoTelefono2) & ", "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pCodigoTelefono2) <> "" Then
            vlSql = vlSql & Trim(pCodigoTelefono2) & ", "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pNumTelefono2) <> "" Then
            vlSql = vlSql & "'" & Trim(pNumTelefono2) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
        Else
                vlSql = vlSql & "'4',"
                vlSql = vlSql & "'1',"
                vlSql = vlSql & "'" & Trim(vl_Fono1_Ben) & "', "
                vlSql = vlSql & "'2',"
                vlSql = vlSql & "NULL,"
                vlSql = vlSql & "'" & Trim(vl_Fono2_Ben) & "', "
            End If
            
            vlSql = vlSql & "'" & vgUsuario & "',"
            vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
            vlSql = vlSql & "'" & Format(Time, "hhmmss") & "') "
            
            
        vgConectarBD.Execute (vlSql)
        
            
    
    
   'INSERTAR EN LA TABLA MAE_BEN_DIRECCION
   
            vlSql = "INSERT INTO PP_TMAE_BEN_DIRECCION "
            vlSql = vlSql + "(NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN,cod_tipoidenben,NUM_IDENBEN,FEC_INGRESO"
            vlSql = vlSql & ",COD_DIRE_VIA,"
            vlSql = vlSql + "GLS_DIRECCION,NUM_DIRECCION,COD_BLOCKCHALET,GLS_BLOCKCHALET,COD_INTERIOR,"
            vlSql = vlSql + "NUM_INTERIOR,COD_CJHT,GLS_NOM_CJHT,GLS_ETAPA,GLS_MANZANA,GLS_LOTE,"
            vlSql = vlSql + "GLS_REFERENCIA,COD_PAIS,COD_DEPARTAMENTO,COD_PROVINCIA,COD_DISTRITO,"
            vlSql = vlSql + "GLS_DESDIREBUSQ,COD_USUARIOCREA,FEC_CREA,HOR_CREA)"
            vlSql = vlSql + "VALUES("
            
            vlSql = vlSql & "'" & Trim(Txt_NumPol) & "', "
            vlSql = vlSql & "1,"
            Msf_GriAseg.Col = 0
            vlSql = vlSql & " " & str(Msf_GriAseg.Text) & ", "
            vlSql = vlSql & " " & str(vlRutGrilla) & ", "
            vlSql = vlSql & "'" & Trim(vlDgvGrilla) & "', "
            vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
            
            If Trim(pTipoVia) <> "" Then
            vlSql = vlSql & "'" & Trim(pTipoVia) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pDireccion) <> "" Then
            vlSql = vlSql & "'" & Trim(pDireccion) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pNumero) <> "" Then
            vlSql = vlSql & "'" & Trim(pNumero) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pTipoBlock) <> "" Then
            vlSql = vlSql & "'" & Trim(pTipoBlock) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pNumBlock) <> "" Then
            vlSql = vlSql & "'" & Trim(pNumBlock) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pTipoPref) <> "" Then
            vlSql = vlSql & "'" & Trim(pTipoPref) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pInterior) <> "" Then
            vlSql = vlSql & "'" & Trim(pInterior) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pTipoConj) <> "" Then
            vlSql = vlSql & "'" & Trim(pTipoConj) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pConjHabit) <> "" Then
            vlSql = vlSql & "'" & Trim(pConjHabit) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pEtapa) <> "" Then
            vlSql = vlSql & "'" & Trim(pEtapa) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pManzana) <> "" Then
            vlSql = vlSql & "'" & Trim(pManzana) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pLote) <> "" Then
            vlSql = vlSql & "'" & Trim(pLote) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(pReferencia) <> "" Then
            vlSql = vlSql & "'" & Trim(pReferencia) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(fgObtenerCodigo_TextoCompuesto(cboNacionalidad)) <> "" Then
            vlSql = vlSql & "'" & Trim(fgObtenerCodigo_TextoCompuesto(cboNacionalidad)) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(vgCodigoRegion) <> "" Then
            vlSql = vlSql & "'" & Trim(vgCodigoRegion) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            If Trim(vgCodigoProvincia) <> "" Then
            vlSql = vlSql & "'" & Trim(vgCodigoProvincia) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            If Trim(vgCodigoComuna) <> "" Then
            vlSql = vlSql & "'" & Trim(vgCodigoComuna) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            If Trim(Lbl_Dir) <> "" Then
            vlSql = vlSql & "'" & Trim(vDireccionConcat) & "', "
            Else
                vlSql = vlSql & "NULL,"
            End If
            
            vlSql = vlSql & "'" & vgUsuario & "',"
            vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
            vlSql = vlSql & "'" & Format(Time, "hhmmss") & "') "
            vgConectarBD.Execute (vlSql)
    'Fin Integracion GobiernoDeDatos_
    
        
       'SI ES UN NUEVO REGISTRO
        vlSql = ""
        vlSql = "INSERT INTO pd_tmae_oripolben ("
        vlSql = vlSql & "num_poliza,num_orden,cod_par,cod_grufam,"
        vlSql = vlSql & "cod_sexo,cod_sitinv,fec_invben,cod_cauinv,"
        vlSql = vlSql & "cod_derpen,cod_dercre,cod_tipoidenben,num_idenben,"
        vlSql = vlSql & "gls_nomben,gls_nomsegben,gls_patben,gls_matben,fec_nacben,"
        vlSql = vlSql & "fec_fallben,fec_nachm,prc_pension,prc_pensionleg,"
        vlSql = vlSql & "prc_pensionrep,mto_pension,"
        vlSql = vlSql & "mto_pensiongar,cod_usuariocrea,fec_crea,hor_crea"
'I--- ABV 07/08/2007 ---
        vlSql = vlSql & ",cod_estpension,prc_pensiongar"
'F--- ABV 07/08/2007 ---
'RRR 26/12/2013
        If vlcod_tipcta <> "" Then vlSql = vlSql & ", cod_tipcta "
        If vlcod_monbco <> "" Then vlSql = vlSql & ", cod_monbco "
        If vlCod_Banco <> "" Then vlSql = vlSql & ", Cod_Banco "
        If vlnum_ctabco <> "" Then vlSql = vlSql & ", num_ctabco "
        
        vlSql = vlSql & ", cod_nacionalidad, num_cuenta_cci"
        'GCP-FRACTAL 11042019
        vlSql = vlSql & ", GLS_FONO, GLS_FONO2, CONS_TRAINFO, CONS_DATCOMER,gls_correoben  "
        
'RRR
        vlSql = vlSql & ") values ("
        vlSql = vlSql & "'" & Trim(vlNumPol) & "', "
        Msf_GriAseg.Col = 0
        vlSql = vlSql & " " & str(Msf_GriAseg.Text) & ", "
        vlSql = vlSql & "'" & Trim(vlCodPar) & "', "
        vlSql = vlSql & "'" & Trim(vlGruFam) & "', "
        vlSql = vlSql & "'" & Trim(vlSexoBen) & "', "
        vlSql = vlSql & "'" & Trim(vlSitInv) & "', "
        If vlFecInv <> "" Then
            vlSql = vlSql & "'" & Trim(vlFecInv) & "', "
        Else
            vlSql = vlSql & "NULL,"
        End If
        If Trim(vlCauInv) <> "" Then
            vlSql = vlSql & "'" & Trim(vlCauInv) & "', "
        Else
            vlSql = vlSql & "'0',"
        End If
        vlSql = vlSql & "'" & Trim(vlDerPen) & "', "
        vlSql = vlSql & "'" & Trim(vlDerCre) & "', "
        vlSql = vlSql & " " & str(vlRutGrilla) & ", "
        vlSql = vlSql & "'" & Trim(vlDgvGrilla) & "', "
        vlSql = vlSql & "'" & Trim(vlNomBen) & "', "
        If Trim(vlNomBenSeg) <> "" Then
            vlSql = vlSql & "'" & Trim(vlNomBenSeg) & "', "
        Else
            vlSql = vlSql & "NULL,"
        End If
        vlSql = vlSql & "'" & Trim(vlPatBen) & "', "
        If Trim(vlMatBen) <> "" Then
            vlSql = vlSql & "'" & Trim(vlMatBen) & "', "
        Else
            vlSql = vlSql & "NULL,"
        End If
        vlSql = vlSql & "'" & Trim(vlFecNacBen) & "', "
        If Trim(vlFecFallBen) <> "" Then
            vlSql = vlSql & "'" & Trim(vlFecFallBen) & "', "
        Else
            vlSql = vlSql & " NULL,"
        End If
        If Trim(vlFecNacHM) <> "" Then
            vlSql = vlSql & "'" & Trim(vlFecNacHM) & "', "
        Else
            vlSql = vlSql & " NULL,"
        End If
        vlSql = vlSql & " " & str(vlPrcPension) & ", "
        vlSql = vlSql & " " & str(vlPrcPensionLeg) & ", "
        vlSql = vlSql & " " & str(vlPrcPensionRep) & ", "
        vlSql = vlSql & " " & str(vlPenBen) & ", "
        vlSql = vlSql & " " & str(vlPenGarBen) & ","
        vlSql = vlSql & "'" & vgUsuario & "',"
        vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
        vlSql = vlSql & "'" & Format(Time, "hhmmss") & "' "
'I--- ABV 07/08/2007 ---
        vlSql = vlSql & ",'" & Trim(vlEstPen) & "' "
        vlSql = vlSql & "," & str(vlPrcPensionGar) & " "
'F--- ABV 07/08/2007 ---
        'RRR 26/12/2013
        If vlcod_tipcta <> "" Then vlSql = vlSql & ",'" & Trim(vlcod_tipcta) & "' "
        If vlcod_monbco <> "" Then vlSql = vlSql & ",'" & Trim(vlcod_monbco) & "' "
        If vlCod_Banco <> "" Then vlSql = vlSql & ",'" & Trim(vlCod_Banco) & "' "
        If vlnum_ctabco <> "" Then vlSql = vlSql & ",'" & Trim(vlnum_ctabco) & "' "
        
        vlSql = vlSql & ",'" & vl_nacionalidadben & "' "
        
        If Trim(vlCodPar) = "99" Then
             vlSql = vlSql & ",'" & vlNum_Cuenta_CCI & "' "
             
            vlSql = vlSql & ",'" & vl_Fono & "' "
            vlSql = vlSql & ",'" & vl_Fono2_Afil & "' "
            vlSql = vlSql & ",'" & VlConTratDatos_Afil & "' "
            vlSql = vlSql & ",'" & VlConUsoDatosCom_Afil & "' "
            vlSql = vlSql & ",'" & vl_Correo_afil & "' "
   
        Else
     
            vlSql = vlSql & ",'" & vlnum_CCI & "' "
       
            vlSql = vlSql & ",'" & vl_Fono1_Ben & "' "
            vlSql = vlSql & ",'" & vl_Fono2_Ben & "' "
            vlSql = vlSql & ",'" & vl_ConTratDatos_Ben & "' "
            vlSql = vlSql & ",'" & vl_ConUsoDatosCom_Ben & "' "
            vlSql = vlSql & ",'" & vl_Correo_Ben & "' "
            
            
            
        End If
        
       
      
        'RRR
        vlSql = vlSql & ") "
        vgConectarBD.Execute (vlSql)

Next vlJ

End Function

''------------------------------------------
''GRABA DATOS DEL BONO EN LA TABLA ORIBONPOL
''------------------------------------------
'Function flGrabaBono()
'
'    vlNumPol = Trim(Txt_NumPol)
'
'    vlSql = ""
'    vlSql = "SELECT cod_tipobono,mto_valnom,fec_emi,fec_ven,"
'    vlSql = vlSql & "prc_tasaint,mto_bonoact,mto_bonoactuf,"
'    vlSql = vlSql & "mto_compra,cod_afeley,num_edadcob "
'    vlSql = vlSql & "FROM pd_tmae_oripolbon WHERE "
'    vlSql = vlSql & "num_poliza = '" & Trim(vlNumPol) & "' "
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If Not vgRs.EOF Then
''        vlNumPol
'        vlCodTipoBono = (vgRs!cod_tipobono)
'        vlMtoValNom = (vgRs!mto_valnom)
'        vlFecEmi = (vgRs!fec_emi)
'        vlFecVen = (vgRs!fec_ven)
'        vlPrcTasaInt = (vgRs!prc_tasaint)
'        vlMtoBonoAct = (vgRs!mto_bonoact)
'        vlMtoBonoActUF = (vgRs!mto_bonoactuf)
'        vlMtoCompra = (vgRs!mto_compra)
'        vlCodAfeLey = (vgRs!cod_afeley)
'        vlNumEdadCob = (vgRs!num_edadcob)
'    End If
'
'    vlSql = ""
'    vlSql = "SELECT num_poliza FROM pd_tmae_oripolbon WHERE "
'    vlSql = vlSql & "num_poliza = '" & vlNumPol & "' AND "
'    vlSql = vlSql & "cod_tipobono = '" & Trim(vlCodTipoBono) & "' "
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If vgRs.EOF Then
'        'ES NUEVO
'        vlSql = ""
'        vlSql = "INSERT INTO pd_tmae_oripolbon ("
'        vlSql = vlSql & "num_poliza,cod_tipobono,mto_valnom,fec_emi,"
'        vlSql = vlSql & "fec_ven,prc_tasaint,mto_bonoact,mto_bonoactuf,"
'        vlSql = vlSql & "mto_compra,cod_afeley,num_edadcob,"
'        vlSql = vlSql & "cod_usuariocrea,fec_crea,hor_crea "
'        vlSql = vlSql & ") values("
'        vlSql = vlSql & "'" & Trim(vlNumPol) & "', "
'        vlSql = vlSql & "'" & Trim(vlCodTipoBono) & "',"
'        vlSql = vlSql & "" & Str(vlMtoValNom) & ","
'        vlSql = vlSql & "'" & Trim(vlFecEmi) & "',"
'        vlSql = vlSql & "'" & Trim(vlFecVen) & "',"
'        vlSql = vlSql & " " & Str(vlPrcTasaInt) & ","
'        vlSql = vlSql & " " & Str(vlMtoBonoAct) & ","
'        vlSql = vlSql & " " & Str(vlMtoBonoActUF) & ","
'        vlSql = vlSql & " " & Str(vlMtoCompra) & ","
'        vlSql = vlSql & "'" & Trim(vlCodAfeLey) & "',"
'        vlSql = vlSql & " " & Str(vlNumEdadCob) & ","
'        vlSql = vlSql & "'" & vgUsuario & "',"
'        vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
'        vlSql = vlSql & "'" & Format(Time, "hhmmss") & "'"
'        vlSql = vlSql & ") "
'
'        vgConectarBD.Execute (vlSql)
'    End If
'
'
'End Function


'------------------------------------------
'GRABA DATOS DEL BONO EN LA TABLA ORIBONPOL
'------------------------------------------
Function flGrabaBono()
Dim vlExiste As Boolean
    
    vlNumPol = Trim(Txt_NumPol)
    vlNumCot = Trim(Lbl_NumCot)
    vlExiste = False
    
    vlSql = ""
    vlSql = "SELECT cod_tipobono,mto_valnom,fec_emi,fec_ven,"
    vlSql = vlSql & "prc_tasaint,mto_bonoact,mto_bonoactuf,"
    vlSql = vlSql & "mto_compra,cod_afeley,num_edadcob "
    vlSql = vlSql & "FROM pt_tmae_cotbono WHERE "
    vlSql = vlSql & "num_cot = '" & Trim(vlNumCot) & "' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
'        vlNumPol
        vlCodTipoBono = (vgRs!cod_tipobono)
        vlMtoValNom = (vgRs!mto_valnom)
        vlFecEmi = (vgRs!fec_emi)
        vlFecVen = (vgRs!fec_ven)
        vlPrcTasaInt = (vgRs!prc_tasaint)
        vlMtoBonoAct = (vgRs!mto_bonoact)
        vlMtoBonoActUF = (vgRs!mto_bonoactuf)
        vlMtoCompra = (vgRs!mto_compra)
        vlCodAfeLey = (vgRs!cod_afeley)
        vlNumEdadCob = (vgRs!num_edadcob)
        vlExiste = True
    End If
    vgRs.Close
    
    If (vlExiste = True) Then
        vlSql = ""
        vlSql = "SELECT num_poliza FROM pd_tmae_oripolbon WHERE "
        vlSql = vlSql & "num_poliza = '" & vlNumPol & "' AND "
        vlSql = vlSql & "cod_tipobono = '" & Trim(vlCodTipoBono) & "' "
        Set vgRs = vgConexionBD.Execute(vlSql)
        If vgRs.EOF Then
            'ES NUEVO
            vlSql = ""
            vlSql = "INSERT INTO pd_tmae_oripolbon ("
            vlSql = vlSql & "num_poliza,cod_tipobono,mto_valnom,fec_emi,"
            vlSql = vlSql & "fec_ven,prc_tasaint,mto_bonoact,mto_bonoactuf,"
            vlSql = vlSql & "mto_compra,cod_afeley,num_edadcob,"
            vlSql = vlSql & "cod_usuariocrea,fec_crea,hor_crea "
            vlSql = vlSql & ") values("
            vlSql = vlSql & "'" & Trim(vlNumPol) & "', "
            vlSql = vlSql & "'" & Trim(vlCodTipoBono) & "',"
            vlSql = vlSql & "" & str(vlMtoValNom) & ","
            vlSql = vlSql & "'" & Trim(vlFecEmi) & "',"
            vlSql = vlSql & "'" & Trim(vlFecVen) & "',"
            vlSql = vlSql & " " & str(vlPrcTasaInt) & ","
            vlSql = vlSql & " " & str(vlMtoBonoAct) & ","
            vlSql = vlSql & " " & str(vlMtoBonoActUF) & ","
            vlSql = vlSql & " " & str(vlMtoCompra) & ","
            vlSql = vlSql & "'" & Trim(vlCodAfeLey) & "',"
            vlSql = vlSql & " " & str(vlNumEdadCob) & ","
            vlSql = vlSql & "'" & vgUsuario & "',"
            vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
            vlSql = vlSql & "'" & Format(Time, "hhmmss") & "'"
            vlSql = vlSql & ") "
    
            vgConectarBD.Execute (vlSql)
        End If
    End If
        
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
        MsgBox "Debe Ingresar el Número de IDentificación del Afiliado.", vbCritical, "Error de Datos"
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
    If (Trim(Lbl_CauInv)) = "" Then
        MsgBox "Debe seleccionar la Causa de Invalidez del Afiliado.", vbCritical, "Error de Datos"
        Cmd_CauInv.SetFocus
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
    If Trim(Lbl_Dir) = "" Then
        MsgBox "Debe Ingresar Dirección del Afiliado", vbExclamation, "Falta Información"
        Cmd_Direccion.SetFocus
        Exit Function
    End If
    If (vlCodDireccion = "") Then
        MsgBox "Debe selecionar la Dirección (Depto-Prov-Dist) del Afiliado", vbExclamation, "Falta Información"
        Cmd_BuscarDir.SetFocus
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
    If Left(Trim(Lbl_TipPen.Caption), 2) = "08" And Trim(Lbl_Representante.Caption) = "" Then
        MsgBox "Debe indicar los datos del Representante.", vbCritical, "Error de Datos"
        Cmd_Representante.SetFocus
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
    If (Trim(Lbl_FecDev) = "") Then
        MsgBox "Debe Ingresar la Fecha Devengue", vbExclamation, "Error de Datos"
        Exit Function
    End If
    If Not IsDate(Lbl_FecDev) Then
        MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
        Exit Function
    End If
    If (Year(CDate(Lbl_FecDev)) < 1900) Then
        MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
        Exit Function
    End If
'    If fgValidaFecha(Trim(Lbl_FecDev)) = False Then
'        Lbl_FecDev = Format(CDate(Trim(Lbl_FecDev)), "yyyymmdd")
'        Lbl_FecDev = DateSerial(Mid((Lbl_FecDev), 1, 4), Mid((Lbl_FecDev), 5, 2), Mid((Lbl_FecDev), 7, 2))
'    Else
'        MsgBox "Debe Ingresar la Fecha Devengue", vbExclamation, "Falta Información"
''        Lbl_FecDev.SetFocus
'        Exit Function
'    End If
    'valida fecha de Incorporación o Aceptación
    If fgValidaFecha(Trim(Lbl_FecIncorpora)) = False Then
        Lbl_FecIncorpora = Format(CDate(Trim(Lbl_FecIncorpora)), "yyyymmdd")
        Lbl_FecIncorpora = DateSerial(Mid((Lbl_FecIncorpora), 1, 4), Mid((Lbl_FecIncorpora), 5, 2), Mid((Lbl_FecIncorpora), 7, 2))
    Else
        MsgBox "Debe Ingresar la Fecha de Incorporación o Aceptación", vbExclamation, "Falta Información"
'        Lbl_FecIncorpora.SetFocus
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
    
    If (Trim(Lbl_CUSPP)) = "" Then
        MsgBox "Debe Ingresar Rut del Afiliado.", vbCritical, "Error de Datos"
'        Lbl_CUSPP.SetFocus
        Exit Function
    End If
    If (Trim(Lbl_TipoIdentCorr)) = "" Then
        MsgBox "Debe Ingresar el Tipo de Identificación del Intermediario.", vbCritical, "Error de Datos"
        Cmd_BuscaCor.SetFocus
        Exit Function
    End If
    If (Not IsNumeric(Lbl_ComInt)) Then
        MsgBox "Debe Ingresar el Porcentaje de Comisión del Intermediario.", vbCritical, "Error de Datos"
        Cmd_BuscaCor.SetFocus
        Exit Function
    End If
    
'I--- ABV 05/02/2011 ---
    'Validar la existencia del Tipo de Reajuste
    If (Trim(Lbl_ReajusteTipo)) = "" Then
        MsgBox "No se encuentra indicado el Tipo de Reajuste de la Póliza. Favor revisar la información de la Cotización y/o Póliza.", vbCritical, "Error de Datos"
'        Lbl_CUSPP.SetFocus
        Exit Function
    End If
'F--- ABV 05/02/2011 ---
    
    flValDatCal = True
    
Exit Function
ERR_VALCAL:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'-------------------------------
'MODIFICA LOS DATOS DE LA POLIZA
'-------------------------------
Function flModificaPoliza(iNumPol As String)
On Error GoTo Err_ModPol
    
    vlSw = True
    
    If Not fgConexionBaseDatos(vgConectarBD) Then
        MsgBox "Fallo la Conexion con la Base de Datos", vbCritical, "Error de Conexión"
        vlSw = False
        Exit Function
    End If
    
    'comenzar la transaccion
    vgConectarBD.BeginTrans
    
    vlFecVig = Mid(Format(Trim(Txt_FecVig), "yyyymmdd"), 1, 6) & "01"
    vlFecEmision = Format(Trim(Txt_FecVig), "yyyymmdd")
    vlNumCtaCCI = Trim(txt_CCI)
    
    vlNumPol = Trim(Txt_NumPol)
    Sql = ""
    Sql = "SELECT num_poliza,prc_corcom,prc_corcomreal,mto_corcom,mto_priuni "
    Sql = Sql & ",fec_inipencia,fec_calculo "
    Sql = Sql & "FROM pd_tmae_oripoliza "
    Sql = Sql & "WHERE num_poliza = '" & Trim(vlNumPol) & "'"
    Set vgRs = vgConexionBD.Execute(Sql)
    If ((vgRs.EOF) = True) Then
        vgConectarBD.RollbackTrans
        vgConectarBD.Close
        MsgBox "La póliza que está Modificando No se Encuentra en la Base de Datos", vbCritical, "Error"
        vlSw = False
        Exit Function
    Else
        
        vlTipoIden = fgObtenerCodigo_TextoCompuesto(Cmb_TipoIdent.Text)
        vlNumIden = Trim(Txt_NumIdent)
        
        'vlCodDir = Trim(Mid(Cmb_Departamento.Text, 1, (InStr(1, Cmb_Departamento, "-") - 1)))
        vlDir = Trim(Left(Lbl_Dir.Text, 50))   'RVF 20090914
        vlCodDir = vlCodDireccion
        vlTipVia = Left(Cmb_TipoVia.Text, 2)
        vlNomVia = Trim(Txt_NombreVia.Text)
        vlNumDmc = Trim(Txt_Numero.Text)
        vlIntDmc = Trim(Txt_Interior.Text)
        vlTipZon = Left(Cmb_TipoZona.Text, 2)
        vlNomZon = Trim(Txt_NombreZona.Text)
        vlReferencia = Trim(Txt_Referencia.Text)
        '*****
        'codigo isapre
        vlCodIsapre = fgObtenerCodigo_TextoCompuesto(Cmb_Salud)
        vlCodVejez = fgObtenerCodigo_TextoCompuesto(Cmb_Vejez.Text)
        vlEstCivil = fgObtenerCodigo_TextoCompuesto(Cmb_EstCivil.Text)
        vlFono = Trim(Txt_Fono)
        vlCorreo = Trim(Txt_Correo)
        vlCodViaPago = Trim(Mid(Cmb_ViaPago.Text, 1, (InStr(1, Cmb_ViaPago, "-") - 1)))
        vlCodTipCta = Trim(Mid(Cmb_TipCta.Text, 1, (InStr(1, Cmb_TipCta, "-") - 1)))
        vlCodMonCta = Trim(Mid(cmb_MonCta.Text, 1, (InStr(1, cmb_MonCta, "-") - 1)))
        vlCodBco = Trim(Mid(Cmb_Bco.Text, 1, (InStr(1, Cmb_Bco, "-") - 1)))
        vlNumCta = Trim(Txt_NumCta)
        vlNumCtaCCI = Trim(txt_CCI)
        
        
        vlCodSuc = Trim(Mid(Cmb_Suc.Text, 1, (InStr(1, Cmb_Suc, "-") - 1)))
        
        'Identificación del Corredor
        If Lbl_TipoIdentCorr <> "" Then
            vlTipoIdencor = fgObtenerCodigo_TextoCompuesto(Lbl_TipoIdentCorr)
        Else
            vlTipoIdencor = ""
        End If
        vlNumIdenCor = Trim(Lbl_NumIdentCorr)
        'Busca la sucursal del corredor
        vlSucCorredor = fgBuscarSucCorredor(vlTipoIdencor, vlNumIdenCor)
        
        vlFecPriPago = Format(Trim(Txt_FecIniPago), "yyyymmdd")
        
        If (Txt_Nacionalidad <> "") Then
            vlNacionalidad = Trim(Txt_Nacionalidad)
        Else
            vlNacionalidad = cgTipoNacionalidad
        End If
        
        vlFecDev = Format(Trim(Lbl_FecDev), "yyyymmdd")
        vlFecIncorporacion = Format(Trim(Lbl_FecIncorpora), "yyyymmdd")
'        vlcuspp = Trim(Lbl_CUSPP)

'        vlPriUni = (vgRs!mto_priuni)
'        vlPrcCorCom = vgRs!Prc_CorCom
'        vlPrcCorComReal = vgRs!Prc_CorComReal
'        vlMtoCorCom = (vgRs!Mto_CorCom)

'        'Verificar Porcentaje de Comision intermediario
'        If vlPrcCorCom = Lbl_ComInt Then
'            vlPrcCorCom = (vgRs!prc_corcom)
'            vlMtoCorCom = (vgRs!mto_corcom)
'            vlPrcCorComReal = vgRs!prc_corcomreal
'        Else
'            vlPrcCorCom = Lbl_ComInt
'            vlMtoCorCom = Format((vlPrcCorCom / 100) * vlPriUni, "##0.00")
'            vlPrcCorComReal = Format((vlPrcCorCom * 100) / (100 + vgPrcBenSocial), "#0.00")
'        End If
'        Lbl_ComIntBen = Format(vlPrcCorComReal, "#0.00")

        vlSql = "UPDATE pd_tmae_oripoliza SET "
        vlSql = vlSql & "cod_tipoidenafi = '" & (vlTipoIden) & "',"
        vlSql = vlSql & "num_idenafi = '" & (vlNumIden) & "',"
        vlSql = vlSql & "cod_direccion = " & vlCodDir & ","
        vlSql = vlSql & "gls_direccion = '" & vlDir & "',"
        
        '-- Begin : Modify by : ricardo.huerta
        vlSql = vlSql & "cod_nacionalidad = '" & vl_nacionalidad & "',"
        '-- End   : Modify by : ricardo.huerta

        
        If vlFono <> "" Then
            vlSql = vlSql & "gls_fono= '" & vlFono & "',"
        Else
            vlSql = vlSql & "gls_fono= Null,"
        End If
        If vlCorreo <> "" Then
            vlSql = vlSql & "gls_correo= '" & vlCorreo & "',"
        Else
            vlSql = vlSql & "gls_correo=Null,"
        End If
        vlSql = vlSql & "cod_viapago = '" & vlCodViaPago & "',"
        vlSql = vlSql & "cod_tipcuenta= '" & vlCodTipCta & "',"
        vlSql = vlSql & "cod_banco= '" & vlCodBco & "',"
        vlSql = vlSql & "cod_MonCta= '" & vlCodMonCta & "',"
        
        
        If Txt_NumCta.Enabled = True Then
            vlSql = vlSql & "num_cuenta= '" & vlNumCta & "',"
        Else
            vlSql = vlSql & "num_cuenta= Null,"
        End If
        
        If txt_CCI.Enabled = True Then
            vlSql = vlSql & "num_cuenta_CCI= '" & vlNumCtaCCI & "',"
        Else
            vlSql = vlSql & "num_cuenta_CCI= Null,"
        End If
        
        
        vlSql = vlSql & "cod_sucursal= '" & vlCodSuc & "',"
        'vlSql = vlSql & "fec_dev = '" & vlFecDev & "',"
        vlSql = vlSql & "cod_tipoidencor = '" & vlTipoIdencor & "',"
        vlSql = vlSql & "num_idencor = '" & vlNumIdenCor & "',"
        vlSql = vlSql & "cod_succorredor = '" & vlSucCorredor & "',"
'        vlSql = vlSql & "prc_corcom = " & Str(vlPrcCorCom) & ","
'        vlSql = vlSql & "mto_corcom = " & Str(vlMtoCorCom) & ","
'        vlSql = vlSql & "prc_corcomreal = " & Str(vlPrcCorComReal) & ","
        vlSql = vlSql & "cod_usuariomodi= '" & vgUsuario & "',"
        vlSql = vlSql & "fec_modi= '" & Format(Date, "yyyymmdd") & "', "
        vlSql = vlSql & "hor_modi= '" & Format(Time, "hhmmss") & "'"
        vlSql = vlSql & ",gls_nacionalidad = '" & vlNacionalidad & "',"
        vlSql = vlSql & "cod_vejez = '" & vlCodVejez & "',"
        vlSql = vlSql & "cod_estcivil = '" & vlEstCivil & "',"
        vlSql = vlSql & "cod_isapre = '" & vlCodIsapre & "',"
        'vlSql = vlSql & "fec_acepta = '" & vlFecIncorporacion & "',"
        vlSql = vlSql & "fec_pripago = '" & vlFecPriPago & "',"
        'vlSql = vlSql & ",cod_cuspp = '" & vlcuspp & "' "
        vlSql = vlSql & "fec_vigencia = '" & Trim(vlFecVig) & "',"
        vlSql = vlSql & "fec_emision = '" & Trim(vlFecEmision) & "',"
        'RVF 20090914
        vlSql = vlSql & "cod_tipvia = '" & vlTipVia & "',"
        vlSql = vlSql & "gls_nomvia = '" & vlNomVia & "',"
        vlSql = vlSql & "gls_numdmc = '" & vlNumDmc & "',"
        vlSql = vlSql & "gls_intdmc = '" & vlIntDmc & "',"
        vlSql = vlSql & "cod_tipzon = '" & vlTipZon & "',"
        vlSql = vlSql & "gls_nomzon = '" & vlNomZon & "',"
        vlSql = vlSql & "gls_referencia = '" & vlReferencia & "' "
        '*****
        If (vgBotonEscogido = "R") Then
        
            'Obtener los datos desde la Estructura de la Póliza Recalculada
            With stPolizaMod
            
                vlNumArchivo = .Num_Archivo
                vlCodAFP = .Cod_AFP
                vlCodTipPen = .Cod_TipPension
                
        '        Trim (Mid(Cmb_Departamento.Text, 1, (InStr(1, Cmb_Departamento, "-") - 1)))
                vlFecSolicitud = .Fec_Solicitud
                'vlFecVig = Format(Trim(Txt_FecVig), "yyyymmdd")
                vlFecDev = .Fec_Dev
                vlFecIncorporacion = .Fec_Acepta
                vlFecCalculo = .Fec_Calculo
                vlFecEmision = .Fec_Emision
                vlcuspp = .Cod_Cuspp
                
                vlCodMonedaFon = .Cod_MonedaFon
                vlMtoValMonedaFon = .Mto_MonedaFon
                vlPriUniFon = .Mto_PrimaFon
                vlCtaIndFon = .Mto_CtaIndFon
                vlMtoBonoFon = .Mto_BonoFon
                vlApoAdiFon = .Mto_ApoAdiFon
                vlApoAdi = .Mto_ApoAdi
                vlPriUni = .Mto_Prima
                vlCtaInd = .Mto_CtaInd
                vlMtoBono = .Mto_Bono
                vlTasaRetPro = .Prc_TasaRPRT
        
                'Verificar Porcentaje de Comision intermediario
                vlPrcCorCom = .Prc_CorCom
                vlMtoCorCom = .Mto_CorCom
                vlPrcCorComReal = .Prc_CorComReal
        '        If CDbl(vlPrcCorCom) = CDbl(Lbl_ComInt) Then
        '            vlPrcCorCom = (vgRs!prc_corcom)
        '            vlMtoCorCom = (vgRs!mto_corcom)
        '            vlPrcCorComReal = vgRs!prc_corcomreal
        '        Else
        '            vlPrcCorCom = Lbl_ComInt
        '            vlMtoCorCom = Format((vlPrcCorCom / 100) * vlPriUni, "##0.00")
        '            vlPrcCorComReal = Format((vlPrcCorCom * 100) / (100 + vgPrcBenSocial), "#0.00")
        '        End If
        '        Lbl_ComIntBen = Format(vlPrcCorComReal, "#0.00")
                
                vlNumAnnoJub = .Num_AnnoJub
                vlNumCargas = .Num_Cargas
                vlIndCobertura = .Ind_Cob
                vlIndBenSocial = .Cod_BenSocial
                
                vlCodMoneda = .Cod_Moneda
                If (vgMonedaCodOfi = .Cod_Moneda) Then
                    vlMtoValMoneda = cgMonedaValorNS
                Else
                    vlMtoValMoneda = .Mto_ValMoneda
                End If
                
'I--- ABV 05/02/2011 ---
                vlCodTipReajuste = .Cod_TipReajuste
                vlMtoValReajusteTri = .Mto_ValReajusteTri
                vlMtoValReajusteMen = .Mto_ValReajusteMen
'F--- ABV 05/02/2011 ---
                
                vlPriUniMod = .Mto_PriUniMod
                vlCtaIndMod = .Mto_CtaIndMod
                vlMtoBonoMod = .Mto_BonoMod
                vlMtoApoAdiMod = .Mto_ApoAdiMod
                
                vlCodTipRen = (.Cod_TipRen)
                vlNumMesDif = (.Num_MesDif)
                vlCodMod = (.Cod_Modalidad)
                vlMesGar = (.Num_MesGar)
                vlCodCoberCon = (.Cod_CoberCon)
                vlMtoFacPenElla = (.Mto_FacPenElla)
                vlPrcFacPenElla = (.Prc_FacPenElla)
                vlDerCre = (.Cod_DerCre)
                vlDerGra = (.Cod_DerGra)
                
                vlRtaAfp = (.Prc_RentaAFP)
                vlPrcRentaAfpOri = (.Prc_RentaAFP)
                vlRtaTmp = (.Prc_RentaTMP)
                vlMtoCuoMor = (.Mto_CuoMor)
                
                vlTasaTCE = (.Prc_TasaCe)
                vlTasaVta = (.Prc_TasaVta)
                vlTasaTir = (.Prc_TasaTir)
                vlTasaPG = (.Prc_TasaIntPerGar)
                vlCNU = (.Mto_CNU)
                vlPriUniSim = (.Mto_PriUniSim)
                vlPriUniDif = (.Mto_PriUniDif)
                vlPension = (.Mto_Pension)
                vlPenGar = (.Mto_PensionGar)
                vlCtaIndAfp = (.Mto_CtaIndAFP)
                vlRtaTmpAFP = (.Mto_RentaTMPAFP)
                vlResMat = (.Mto_ResMat)
                vlValPPTmp = (.Mto_ValPrePenTmp)
                vlMtoPerCon = (.Mto_PerCon)
                vlPrcPerCon = (.Prc_PerCon)
                
                vlMtoSumPension = (.Mto_SumPension)
                vlMtoPenAnual = (.Mto_PenAnual)
                vlMtoRMPension = (.Mto_RMPension)
                vlMtoRMGtoSep = (.Mto_RMGtoSep)
                vlMtoRMGtoSepRV = (.Mto_RMGtoSepRV)
                vlMtoAjusteIPC = .Mto_AjusteIPC
                
                vlCodTipCot = (.Cod_TipoCot)
                vlCodEstCot = (.Cod_EstCot)
                
                vlFecFinPerDif = fgCalcularFechaFinPerDiferido(vlFecDev, CLng(vlNumMesDif))
                vlFecFinPerGar = fgCalcularFechaFinPerGarantizado(vlFecDev, CLng(vlNumMesDif), CLng(vlMesGar))
                vlFecIniPagoPen = fgCalcularFechaIniPagoPensiones(vlFecDev, CLng(vlNumMesDif))
                
                
        
            End With
        
            vlReCalculo = "S"
        
            vlSql = vlSql & ", num_archivo = " & vlNumArchivo & ", "
            vlSql = vlSql & "cod_afp = '" & Trim(vlCodAFP) & "',"
            vlSql = vlSql & "cod_tippension = '" & Trim(vlCodTipPen) & "',"
            
            vlSql = vlSql & "cod_cuspp = '" & (vlcuspp) & "',"
            vlSql = vlSql & "fec_solicitud = '" & Trim(vlFecSolicitud) & "',"
            vlSql = vlSql & "fec_dev = '" & Trim(vlFecDev) & "',"
            vlSql = vlSql & "fec_acepta = '" & Trim(vlFecIncorporacion) & "',"
            
            If Trim(vlCodMonedaFon) <> "" Then
                vlSql = vlSql & "cod_monedafon = '" & Trim(vlCodMonedaFon) & "',"
            Else
                vlSql = vlSql & "cod_monedafon = NULL" & ","
            End If
            
            vlSql = vlSql & "mto_monedafon = " & str(vlMtoValMonedaFon) & ","
            vlSql = vlSql & "mto_priunifon = " & str(vlPriUniFon) & ","
            vlSql = vlSql & "mto_ctaindfon = " & str(vlCtaIndFon) & ","
            vlSql = vlSql & "mto_bonofon = " & str(vlMtoBonoFon) & ","
            vlSql = vlSql & "mto_apoadifon = " & str(vlMtoApoAdiFon) & ","
            vlSql = vlSql & "mto_apoadi = " & str(vlApoAdi) & ","
            vlSql = vlSql & "mto_priuni = " & str(vlPriUni) & ","
            vlSql = vlSql & "mto_ctaind = " & str(vlCtaInd) & ","
            vlSql = vlSql & "mto_bono = " & str(vlMtoBono) & ","
            vlSql = vlSql & "prc_tasarprt = " & str(vlTasaRetPro) & ","
            vlSql = vlSql & "prc_corcom = " & str(vlPrcCorCom) & ","
            vlSql = vlSql & "prc_corcomreal = " & str(vlPrcCorComReal) & ","
            vlSql = vlSql & "mto_corcom = " & str(vlMtoCorCom) & ","
            vlSql = vlSql & "num_annojub = " & str(vlNumAnnoJub) & " ,"
            vlSql = vlSql & "num_cargas = " & str(vlNumCargas) & " ,"
            vlSql = vlSql & "ind_cob = '" & (vlIndCobertura) & "',"
            vlSql = vlSql & "cod_bensocial = '" & (vlIndBenSocial) & "',"
            If Trim(vlCodMoneda) <> "" Then
                vlSql = vlSql & "cod_moneda = '" & Trim(vlCodMoneda) & "',"
            Else
                vlSql = vlSql & "cod_moneda = NULL" & ","
            End If
            vlSql = vlSql & "mto_valmoneda = " & str(vlMtoValMoneda) & ","
            vlSql = vlSql & "mto_priunimod = " & str(vlPriUniMod) & ","
            vlSql = vlSql & "mto_ctaindmod = " & str(vlCtaIndMod) & ","
            vlSql = vlSql & "mto_bonomod = " & str(vlMtoBonoMod) & ","
            vlSql = vlSql & "mto_apoadimod = " & str(vlMtoApoAdiMod) & ","
            vlSql = vlSql & "cod_tipren = '" & Trim(vlCodTipRen) & "',"
            vlSql = vlSql & "num_mesdif = " & str(vlNumMesDif) & " ,"
            vlSql = vlSql & "cod_modalidad = '" & Trim(vlCodMod) & "',"
            vlSql = vlSql & "num_mesgar = " & str(vlMesGar) & " ,"
            vlSql = vlSql & "cod_cobercon = '" & Trim(vlCodCoberCon) & "',"
            vlSql = vlSql & "mto_facpenella = " & str(vlMtoFacPenElla) & " ,"
            vlSql = vlSql & "prc_facpenella = " & str(vlPrcFacPenElla) & " ,"
            vlSql = vlSql & "cod_dercre = '" & Trim(vlDerCre) & "',"
            vlSql = vlSql & "cod_dergra = '" & Trim(vlDerGra) & "',"
            vlSql = vlSql & "prc_rentaafp = " & str(vlRtaAfp) & " ,"
            vlSql = vlSql & "prc_rentaafpori = " & str(vlPrcRentaAfpOri) & " ,"
            vlSql = vlSql & "prc_rentatmp = " & str(vlRtaTmp) & " ,"
            vlSql = vlSql & "mto_cuomor = " & str(vlMtoCuoMor) & " ,"
            vlSql = vlSql & "prc_tasace = " & str(vlTasaTCE) & " ,"
            vlSql = vlSql & "prc_tasavta = " & str(vlTasaVta) & " ,"
            vlSql = vlSql & "prc_tasatir = " & str(vlTasaTir) & " ,"
            vlSql = vlSql & "prc_tasapergar = " & str(vlTasaPG) & " ,"
            vlSql = vlSql & "mto_cnu = " & str(vlCNU) & " ,"
            vlSql = vlSql & "mto_priunisim = " & str(vlPriUniSim) & " ,"
            vlSql = vlSql & "mto_priunidif = " & str(vlPriUniDif) & " ,"
            vlSql = vlSql & "mto_pension = " & str(vlPension) & " ,"
            vlSql = vlSql & "mto_pensiongar = " & str(vlPenGar) & " ,"
            vlSql = vlSql & "mto_ctaindafp = " & str(vlCtaIndAfp) & " ,"
            vlSql = vlSql & "mto_rentatmpafp = " & str(vlRtaTmpAFP) & " ,"
            vlSql = vlSql & "mto_resmat = " & str(vlResMat) & " ,"
            vlSql = vlSql & "mto_valprepentmp = " & str(vlValPPTmp) & " ,"
            vlSql = vlSql & "mto_percon = " & str(vlMtoPerCon) & " ,"
            vlSql = vlSql & "prc_percon = " & str(vlPrcPerCon) & " ,"
            
            vlSql = vlSql & "mto_sumpension = " & str(vlMtoSumPension) & " ,"
            vlSql = vlSql & "mto_penanual = " & str(vlMtoPenAnual) & " ,"
            vlSql = vlSql & "mto_rmpension = " & str(vlMtoRMPension) & " ,"
            vlSql = vlSql & "mto_rmgtosep = " & str(vlMtoRMGtoSep) & " ,"
            vlSql = vlSql & "mto_rmgtoseprv = " & str(vlMtoRMGtoSepRV) & " ,"
            
'            vlSql = vlSql & "cod_tipcot = '" & Trim(vlCodTipCot) & "',"
'            vlSql = vlSql & "cod_estcot = '" & Trim(vlCodEstCot) & "',"
'            If Trim(vlCodUsuario) <> "" Then
'                vlSql = vlSql & "'" & Trim(vlCodUsuario) & "',"
'            Else
'                vlSql = vlSql & "NULL" & ","
'            End If
'            If Trim(vlSucUsu) <> "" Then
'                vlSql = vlSql & "'" & Trim(vlSucUsu) & "',"
'            Else
'                vlSql = vlSql & "NULL" & ","
'            End If
                
            If Trim(vlFecFinPerDif) <> "" Then
                vlSql = vlSql & "fec_finperdif = '" & Trim(vlFecFinPerDif) & "',"
            Else
                vlSql = vlSql & "fec_finperdif = NULL" & ","
            End If
                
            If Trim(vlFecFinPerGar) <> "" Then
                vlSql = vlSql & "fec_finpergar = '" & Trim(vlFecFinPerGar) & "',"
            Else
                vlSql = vlSql & "fec_finpergar = NULL" & ","
            End If
                
            vlSql = vlSql & "ind_recalculo = '" & Trim(vlReCalculo) & "',"
            vlSql = vlSql & "fec_inipencia = '" & Trim(vlFecIniPagoPen) & "', "
            vlSql = vlSql & "fec_calculo = '" & Trim(vlFecCalculo) & "' "
            vlSql = vlSql & ",mto_ajusteIPC = " & str(vlMtoAjusteIPC) & " "
            
'I--- ABV 05/02/2011 ---
            vlSql = vlSql & ",cod_tipreajuste = '" & vlCodTipReajuste & "',"
            vlSql = vlSql & "mto_valreajustetri = " & str(vlMtoValReajusteTri) & ","
            vlSql = vlSql & "mto_valreajustemen = " & str(vlMtoValReajusteMen) & " "
'F--- ABV 05/02/2011 ---
        End If
        
        vlSql = vlSql & " WHERE num_poliza = '" & vlNumPol & "'"
        vgConectarBD.Execute vlSql
        
        'actualiza los beneficiarios
        If vlBotonEscogido = "C" Then
            Call flGrabaBeneficiarioCot
        Else
            Call flGrabaBeneficiario
        End If
        'RVF 20090914
      
        
        vlCodTipPen = Trim(Mid(Lbl_TipPen, 1, (InStr(1, Lbl_TipPen, "-") - 1)))
        
        'El Bono no se actualiza, ya que ninguno de sus datos se modifica
        'desde pantalla
'        'actualiza bono
'        If vlCodTipPen = "05" Then
'            Call flGrabaBono
'        Else
'            vlSql = ""
'            vlSql = "DELETE FROM pd_tmae_oripolbon WHERE num_poliza = '" & vlNumPol & "'"
'            vgConectarBD.Execute vlSql
'        End If



   If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
            Call pGrabaRepresentante
   End If

      
   
        
        If vlSw = True Then
            'ejecutar la transaccion
            vgConectarBD.CommitTrans
            'vgConectarBD.RollbackTrans
        
            'cerrar la transaccion
            vgConectarBD.Close
            
     
            MsgBox "Los Datos de la Póliza han sido Modificados Satisfactoriamente", vbInformation, "Proceso Completado"
'            Cmd_Cancelar_Click
            Cmd_CrearPol.Enabled = False
            Cmd_Grabar.Enabled = True
            cmdEnviaCorreo.Enabled = True
            Cmd_Eliminar.Enabled = True
            
            If (vgNivelIndicadorBoton = "S") Then
                Cmd_Editar.Enabled = True
            Else
                Cmd_Editar.Enabled = False
            End If
            
            'Call flCargaCarpBenef(Txt_NumPol.Text)
        Else
            'Deshacer la transaccion
            vgConectarBD.RollbackTrans
            'cerrar la transaccion
            vgConectarBD.Close
        
        End If
        Screen.MousePointer = 0
    
    End If

Exit Function
Err_ModPol:
    Screen.MousePointer = 0
    vgConectarBD.RollbackTrans
    vgConectarBD.Close
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'-------------------------------------
'ELIMINA LA POLIZA Y TODO SU CONTENIDO
'-------------------------------------
Function flEliminarPoliza(iNumPol As String)
Dim vlUltimoNumPol As String
On Error GoTo err_eli
    
   If Not fgConexionBaseDatos(vgConectarBD) Then
        MsgBox "Fallo la Conexion con la Base de Datos", vbCritical, "Error de Conexión"
        Exit Function
    End If
    
    'comenzar la transaccion
    vgConectarBD.BeginTrans
    
'I--- ABV 07/08/2007 ---
'Validar que al Eliminar se trata de la última Póliza
    vlUltimoNumPol = ""
    vlSql = "SELECT max(num_poliza) as numero FROM pd_tmae_gennumpol "
    Set vgRs = vgConectarBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlUltimoNumPol = Trim(vgRs!Numero)
    End If
    vgRs.Close
    
    vgI = Len(vlUltimoNumPol)
    vlUltimoNumPol = String(10 - vgI, "0") & vlUltimoNumPol
    
    If (vlUltimoNumPol <> iNumPol) Then
        vgConectarBD.RollbackTrans
        vgConectarBD.Close
        MsgBox "El Número de Póliza no corresponde con la última Generación" & Chr(13) & "de la Numeración de Pólizas.", vbCritical, "Operación Cancelada"
        Screen.MousePointer = 0
        Exit Function
    End If
'F--- ABV 07/08/2007 ---
    
'    'Eliminar Registros de Bono
'    vlSql = ""
'    vlSql = "SELECT num_poliza FROM pd_tmae_oripolbon WHERE "
'    vlSql = vlSql & "num_poliza= '" & iNumPol & "'"
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If Not vgRs.EOF Then
'        vlSql = ""
'        vlSql = "DELETE FROM pd_tmae_oripolbon WHERE "
'        vlSql = vlSql & "num_poliza = '" & iNumPol & "'"
'        vgConectarBD.Execute (vlSql)
'    End If
    
'I--- ABV 15/10/2007 ---
'Eliminar Tabla de Recepción de Primas
        vlSql = "DELETE FROM pd_tmae_polprirec WHERE "
        vlSql = vlSql & "num_poliza = '" & iNumPol & "' "
        vgConectarBD.Execute (vlSql)
'F--- ABV 15/10/2007 ---
    
    'Eliminar Registros de Beneficiarios
'    vlSql = ""
'    vlSql = "SELECT num_poliza FROM pd_tmae_oripolben WHERE "
'    vlSql = vlSql & "num_poliza= '" & iNumPol & "'"
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If Not vgRs.EOF Then
        vlSql = "DELETE FROM pd_tmae_oripolben WHERE "
        vlSql = vlSql & "num_poliza = '" & iNumPol & "' "
        vgConectarBD.Execute (vlSql)
'    End If
    
    'Eliminar Registro de Póliza
    vlSql = ""
    vlSql = "DELETE FROM pd_tmae_oripoliza WHERE "
    vlSql = vlSql & "num_poliza = '" & iNumPol & "' "
    vgConectarBD.Execute (vlSql)

'I--- ABV 07/08/2007 ---
'    'Modificar Estado de la Cotización
'    vlSql = ""
'    vlSql = "SELECT num_cot FROM pt_tmae_detcotizacion WHERE "
'    vlSql = vlSql & "num_cot = '" & Trim(vlNumCot) & "' AND "
'    vlSql = vlSql & "cod_estcot = '" & clCodEstCotP & "' "
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If Not vgRs.EOF Then
'        vlSql = ""
'        vlSql = " UPDATE pt_tmae_detcotizacion SET "
'        vlSql = vlSql & "cod_estcot = '" & clCodEstCotA & "' "
'        vlSql = vlSql & "WHERE num_cot = '" & vlNumCot & "' ANd "
'        vlSql = vlSql & "cod_estcot = '" & clCodEstCotA & "' "
'        vgConectarBD.Execute (vlSql)
'    End If
    
    'Modificar el Estado de la Cotización que se genero como Póliza a Aceptada
'    vlSql = "SELECT num_cot FROM pt_tmae_detcotizacion WHERE "
'    vlSql = vlSql & "num_cot = '" & Trim(vlNumCot) & "' AND "
'    vlSql = vlSql & "cod_estcot = '" & clCodEstCotA & "' "
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If Not vgRs.EOF Then
        vlSql = " UPDATE pt_tmae_detcotizacion SET "
        vlSql = vlSql & "cod_estcot = '" & clCodEstCotA & "' "
        vlSql = vlSql & "WHERE "
        vlSql = vlSql & "num_cot = '" & vlNumCot & "' AND "
        vlSql = vlSql & "num_correlativo = " & vlNumCorrelativo & " "
        vgConectarBD.Execute (vlSql)
'    End If
'    vgRs.Close
    
    'Actualizar el Ultimo Número de la Generación de Números de Póliza
    vlSql = "UPDATE pd_tmae_gennumpol SET "
    vlSql = vlSql & "num_poliza = " & CDbl(vlUltimoNumPol) - 1
    vgConectarBD.Execute (vlSql)
'F--- ABV 07/08/2007 ---
    
    'marco---23/03/2010
    vlSql = ""
    vlSql = "UPDATE PD_TMAE_ORITUTOR SET NUM_POLIZA=' ' WHERE NUM_COTIZACION='" & vlNumCot & "'"
    vgConectarBD.Execute (vlSql)
    
    'ejecutar la transaccion
    vgConectarBD.CommitTrans
    
    'cerrar la transaccion
    vgConectarBD.Close
    
    Cmd_Cancelar_Click
    SSTab_Poliza.Tab = 0
       
    MsgBox "Los Datos de la Póliza seleccionada han sido completamente Eliminados", vbInformation, "Proceso de Eliminación"

Exit Function
err_eli:
    Screen.MousePointer = 0
    vgConectarBD.RollbackTrans
    vgConectarBD.Close
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'-----------------------------------------------
'FUNCION QUE GRABA LOS DATOS DE UNA NUEVA POLIZA
'-----------------------------------------------
Function flCreaPoliza()
 Dim vlFono2 As String
 

    vlSw = True
    If (Trim(Txt_NumPol) = "") Then
        MsgBox "No se ha podido generar el número de Póliza", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        vlSw = False
        Exit Function
    End If
    
    vlNumCot = Trim(Lbl_NumCot)
    vlNumPol = Trim(Txt_NumPol)
    vlNumCorrelativo = Trim(Lbl_SecOfe)
    vlCodSolOfe = Trim(Lbl_SolOfe) 'N° Operación
    
'I--- ABV 08/07/2007 ---
'    vlFecVig = Trim(Txt_FecVig)
'    vlFecVig = Format(vlFecVig, "yyyymmdd")
    vlFecEmision = Trim(Txt_FecVig)
    vlFecEmision = Format(vlFecEmision, "yyyymmdd")
    vlFecVig = Trim(Txt_FecVig)
    vlFecVig = Mid(Format(vlFecVig, "yyyymmdd"), 1, 6) & "01"
'I--- ABV 08/07/2007 ---

    'VALIDA SI YA EXISTE LA POLIZA EN LA BD
    vlSql = ""
    vlSql = "SELECT num_poliza FROM pd_tmae_oripoliza WHERE "
    vlSql = vlSql & "num_poliza = '" & (Txt_NumPol) & "'"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        MsgBox "El número de Póliza ya existe en la BD", vbCritical, "Error"
        Screen.MousePointer = 0
        vlSw = False
        Exit Function
    End If
    vgRs.Close

'I--- ABV 04/12/2009 ---
    vlMarcaSobDif = "N"
'F--- ABV 04/12/2009 ---
    
If (vgBotonEscogido <> "R") Then
    'Seleccionar Datos desde tablas de cotizacion, para generar poliza
    vlSql = ""
    vlSql = "SELECT c.cod_afp,c.cod_tippension," ',cod_isapre
    vlSql = vlSql & "c.num_annojub,c.num_cargas, " 'c.cod_vejez,c.cod_estcivil,
    vlSql = vlSql & "" 'c.cod_tipoidencor,c.num_idencor,
    vlSql = vlSql & "d.prc_corcom,d.prc_corcomreal,d.mto_corcom," 'c.cod_viapago,
    vlSql = vlSql & "d.num_correlativo,d.num_archivo,d.num_operacion,"
    vlSql = vlSql & "c.fec_suscripcion as fec_solicitud,"
    vlSql = vlSql & "c.cod_monedafon,c.mto_monedafon,mto_priunifon,"
    vlSql = vlSql & "c.mto_ctaindfon,c.mto_bonofon,c.mto_apoadi,c.prc_tasarprt,"
    vlSql = vlSql & "c.mto_ctaind,c.mto_priuni,c.mto_bono,"
    vlSql = vlSql & "c.ind_cob,c.cod_bensocial,d.cod_cobercon,"
    vlSql = vlSql & "d.cod_moneda,d.mto_valmoneda,"
    vlSql = vlSql & "d.mto_ctaindmod,d.mto_priunimod,d.mto_bonomod,"
    vlSql = vlSql & "d.cod_tipren,d.num_mesdif, d.cod_modalidad,d.num_mesgar,d.prc_rentaafp,d.prc_rentaafpori, "
    vlSql = vlSql & "d.mto_facpenella,d.prc_facpenella,d.cod_dercre,d.cod_dergra,"
    vlSql = vlSql & "d.prc_rentaafp,d.prc_rentaafpori,d.prc_rentatmp,"
    vlSql = vlSql & "d.mto_cuomor, "
    vlSql = vlSql & "d.prc_tasatce,d.prc_tasavta,d.prc_tasatir,d.prc_tasapergar,"
    vlSql = vlSql & "d.mto_cnu,d.mto_priunisim,d.mto_priunidif,d.mto_pension,"
    vlSql = vlSql & "d.mto_pensiongar,d.mto_ctaindafp, "
    vlSql = vlSql & "d.mto_rentatmpafp,d.mto_resmat,d.mto_valprepentmp, "
    vlSql = vlSql & "d.mto_sumpension,d.mto_penanual,d.mto_rmpension,d.mto_rmgtosep, "
    vlSql = vlSql & "d.mto_rmgtoseprv, "
    vlSql = vlSql & "d.mto_percon,d.prc_percon,d.cod_estcot,c.cod_tipcot "
    vlSql = vlSql & ",d.fec_calculo,d.mto_ajusteIPC,c.fec_dev,d.Fec_Acepta "
'I--- ABV 04/12/2009 ---
    vlSql = vlSql & ",d.ind_calsobdif "
'F--- ABV 04/12/2009 ---
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & ",d.cod_tipreajuste,d.mto_valreajustetri,d.mto_valreajustemen, fec_devsol "
    vlSql = vlSql & ",c.cod_tipoidencor,c.num_idencor, "
    'Inicio Cuando el supervisor y el jefe vienen vacios
    vlSql = vlSql & " nvl(c.num_idensup,cr.num_idenjefe) as num_idensup, "
    vlSql = vlSql & " nvl(c.num_idenjef, cr.num_idenjefes) As num_idenjef "
    'Fin  el supervisor y el jefe vienen vacios
'F--- ABV 05/02/2011 ---
'I--- ABV 22/06/2006 ---
'    vlSql = vlSql & "FROM pt_tmae_cotizacion c, pt_tmae_detcotizacion d "
    vlSql = vlSql & "FROM pt_tmae_cotizacion c, " & vlTablaDetCotizacion & " d,  pt_tmae_corredor cr "
'F--- ABV 22/06/2006 ---
    vlSql = vlSql & "WHERE c.num_cot = '" & Trim(Lbl_NumCot) & "' AND "
    vlSql = vlSql & "c.num_cot = d.num_cot AND "
    vlSql = vlSql & "c.num_idencor = cr.num_idencor AND "
    vlSql = vlSql & "d.cod_estcot = '" & clCodEstCotA & "' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
       
        vlNumArchivo = (vgRs!Num_Archivo)
        vlCodAFP = (vgRs!Cod_AFP)
        vlCodTipPen = (vgRs!Cod_TipPension)
        
        vlTipoIdencor = vgRs!cod_tipoidencor
        VlNum_idencor = vgRs!Num_IdenCor
        vlNum_idensup = vgRs!num_idensup
        vlNum_idenjef = vgRs!num_idenjef
   
        
        
'        Trim (Mid(Cmb_Departamento.Text, 1, (InStr(1, Cmb_Departamento, "-") - 1)))
        vlFecSolicitud = (vgRs!Fec_Solicitud)
        'vlFecVig = Format(Trim(Txt_FecVig), "yyyymmdd")
        vlFecDev = vgRs!Fec_Dev
        vlFecIncorporacion = vgRs!Fec_Acepta
        vlFecCalculo = vgRs!Fec_Calculo
        vlcuspp = Trim(Lbl_CUSPP)
        
        vlCodMonedaFon = vgRs!Cod_MonedaFon
        vlMtoValMonedaFon = vgRs!Mto_MonedaFon
        vlPriUniFon = vgRs!mto_priunifon
        vlCtaIndFon = vgRs!Mto_CtaIndFon
        vlMtoBonoFon = vgRs!Mto_BonoFon
        vlApoAdiFon = Format(vgRs!Mto_ApoAdi / vgRs!Mto_MonedaFon, "#0.00")
        vlPriUni = (vgRs!MTO_PRIUNI)
        vlCtaInd = (vgRs!Mto_CtaInd)
        vlMtoBono = (vgRs!Mto_Bono)
        vlMtoApoAdi = vgRs!Mto_ApoAdi
        vlTasaRetPro = vgRs!Prc_TasaRPRT

        'Verificar Porcentaje de Comision intermediario
        vlPrcCorCom = vgRs!Prc_CorCom
        vlMtoCorCom = vgRs!Mto_CorCom
        vlPrcCorComReal = vgRs!Prc_CorComReal
'        If CDbl(vlPrcCorCom) = CDbl(Lbl_ComInt) Then
'            vlPrcCorCom = (vgRs!prc_corcom)
'            vlMtoCorCom = (vgRs!mto_corcom)
'            vlPrcCorComReal = vgRs!prc_corcomreal
'        Else
'            vlPrcCorCom = Lbl_ComInt
'            vlMtoCorCom = Format((vlPrcCorCom / 100) * vlPriUni, "##0.00")
'            vlPrcCorComReal = Format((vlPrcCorCom * 100) / (100 + vgPrcBenSocial), "#0.00")
'        End If
'        Lbl_ComIntBen = Format(vlPrcCorComReal, "#0.00")
        
        vlNumAnnoJub = (vgRs!Num_AnnoJub)
        vlNumCargas = (vgRs!Num_Cargas)
        vlIndCobertura = (vgRs!Ind_Cob)
        vlIndBenSocial = (vgRs!Cod_BenSocial)
        
        vlCodMoneda = (vgRs!Cod_Moneda)
        If (vgMonedaCodOfi = vgRs!Cod_Moneda) Then
            vlMtoValMoneda = cgMonedaValorNS
        Else
            vlMtoValMoneda = (vgRs!Mto_ValMoneda)
        End If
        
'I--- ABV 05/02/2011 ---
        vlCodTipReajuste = Trim(vgRs!Cod_TipReajuste)
        vlMtoValReajusteTri = vgRs!Mto_ValReajusteTri
        vlMtoValReajusteMen = vgRs!Mto_ValReajusteMen
'F--- ABV 05/02/2011 ---
        
        vlPriUniMod = vgRs!Mto_PriUniMod
        vlCtaIndMod = vgRs!Mto_CtaIndMod
        vlMtoBonoMod = vgRs!Mto_BonoMod
        If (vgMonedaCodOfi = vgRs!Cod_Moneda) Then
            vlMtoApoAdiMod = vgRs!Mto_ApoAdi
        Else
            vlMtoApoAdiMod = Format(vgRs!Mto_ApoAdi / vlMtoValMoneda, "#0.00")
        End If
        
        vlCodTipRen = (vgRs!Cod_TipRen)
        vlNumMesDif = (vgRs!Num_MesDif)
        vlCodMod = (vgRs!Cod_Modalidad)
        vlMesGar = (vgRs!Num_MesGar)
        vlCodCoberCon = (vgRs!Cod_CoberCon)
        vlMtoFacPenElla = (vgRs!Mto_FacPenElla)
        vlPrcFacPenElla = (vgRs!Prc_FacPenElla)
        vlDerCre = (vgRs!Cod_DerCre)
        vlDerGra = (vgRs!Cod_DerGra)
        
        vlRtaAfp = (vgRs!Prc_RentaAFP)
        vlPrcRentaAfpOri = (vgRs!prc_rentaafpori)
        vlRtaTmp = (vgRs!Prc_RentaTMP)
        vlMtoCuoMor = (vgRs!Mto_CuoMor)
        
        vlTasaTCE = (vgRs!prc_tasatce)
        vlTasaVta = (vgRs!Prc_TasaVta)
        vlTasaTir = (vgRs!Prc_TasaTir)
        vlTasaPG = (vgRs!prc_tasapergar)
        vlCNU = (vgRs!Mto_CNU)
        vlPriUniSim = (vgRs!Mto_PriUniSim)
        vlPriUniDif = (vgRs!Mto_PriUniDif)
        vlPension = (vgRs!Mto_Pension)
        If vlCodTipPen = "08" Then
            vlPenGar = (vgRs!Mto_SumPension)
        Else
            vlPenGar = (vgRs!Mto_PensionGar)
        End If
        vlCtaIndAfp = (vgRs!Mto_CtaIndAFP)
        vlRtaTmpAFP = (vgRs!Mto_RentaTMPAFP)
        vlResMat = (vgRs!Mto_ResMat)
        vlValPPTmp = (vgRs!Mto_ValPrePenTmp)
        vlMtoPerCon = (vgRs!Mto_PerCon)
        vlPrcPerCon = (vgRs!Prc_PerCon)
        
        vlMtoSumPension = (vgRs!Mto_SumPension)
        vlMtoPenAnual = (vgRs!Mto_PenAnual)
        vlMtoRMPension = (vgRs!Mto_RMPension)
        vlMtoRMGtoSep = (vgRs!Mto_RMGtoSep)
        vlMtoRMGtoSepRV = (vgRs!Mto_RMGtoSepRV)
        vlMtoAjusteIPC = (vgRs!Mto_AjusteIPC)
        
        vlCodTipCot = (vgRs!cod_tipcot)
        vlCodEstCot = (vgRs!Cod_EstCot)
        If vlCodTipRen = "6" Then
            vlFecFinPerDif = fgCalcularFechaFinPerDiferido(vlFecDev, CLng(vlNumMesDif))
            vlFecFinPerGar = fgCalcularFechaFinPerGarantizado(vlFecDev, 0, CLng(vlMesGar))
            vlFecIniPagoPen = vlFecDev
        Else
            vlFecFinPerDif = fgCalcularFechaFinPerDiferido(vlFecDev, CLng(vlNumMesDif))
            vlFecFinPerGar = fgCalcularFechaFinPerGarantizado(vlFecDev, CLng(vlNumMesDif), CLng(vlMesGar))
            vlFecIniPagoPen = fgCalcularFechaIniPagoPensiones(vlFecDev, CLng(vlNumMesDif))
        End If
        
        
        vlfecDevSol = IIf(IsNull((vgRs!fec_devsol)), "", (vgRs!fec_devsol)) '--18/9/13
'I--- ABV 04/12/2009 ---
        If Not IsNull(vgRs!Ind_CalSobDif) Then
            vlMarcaSobDif = Trim(vgRs!Ind_CalSobDif)
        End If
'F--- ABV 04/12/2009 ---
        
        vlReCalculo = "N"
    End If
Else
    'Obtener los datos desde la Estructura de la Póliza Recalculada
    With stPolizaMod
    
        vlNumArchivo = .Num_Archivo
        vlCodAFP = .Cod_AFP
        vlCodTipPen = .Cod_TipPension
        
'        Trim (Mid(Cmb_Departamento.Text, 1, (InStr(1, Cmb_Departamento, "-") - 1)))
        vlFecSolicitud = .Fec_Solicitud
        'vlFecVig = Format(Trim(Txt_FecVig), "yyyymmdd")
        vlFecDev = .Fec_Dev
        vlFecIncorporacion = .Fec_Acepta
        vlFecCalculo = .Fec_Calculo
        vlcuspp = .Cod_Cuspp
        
        vlCodMonedaFon = .Cod_MonedaFon
        vlMtoValMonedaFon = .Mto_MonedaFon
        vlPriUniFon = .Mto_PrimaFon
        vlCtaIndFon = .Mto_CtaIndFon
        vlMtoBonoFon = .Mto_BonoFon
        'vlApoAdiFon = .Mto_ApoAdiFon
        If (vgMonedaCodOfi = .Cod_MonedaFon) Then
            vlApoAdiFon = .Mto_ApoAdi
        Else
            vlApoAdiFon = Format(.Mto_ApoAdi / .Mto_MonedaFon, "#0.00")
        End If
        
        vlPriUni = .Mto_Prima
        vlCtaInd = .Mto_CtaInd
        vlMtoBono = .Mto_Bono
        vlMtoApoAdi = .Mto_ApoAdi
        vlTasaRetPro = .Prc_TasaRPRT

        'Verificar Porcentaje de Comision intermediario
        vlPrcCorCom = .Prc_CorCom
        vlMtoCorCom = .Mto_CorCom
        vlPrcCorComReal = .Prc_CorComReal
'        If CDbl(vlPrcCorCom) = CDbl(Lbl_ComInt) Then
'            vlPrcCorCom = (vgRs!prc_corcom)
'            vlMtoCorCom = (vgRs!mto_corcom)
'            vlPrcCorComReal = vgRs!prc_corcomreal
'        Else
'            vlPrcCorCom = Lbl_ComInt
'            vlMtoCorCom = Format((vlPrcCorCom / 100) * vlPriUni, "##0.00")
'            vlPrcCorComReal = Format((vlPrcCorCom * 100) / (100 + vgPrcBenSocial), "#0.00")
'        End If
'        Lbl_ComIntBen = Format(vlPrcCorComReal, "#0.00")
        
        vlNumAnnoJub = .Num_AnnoJub
        vlNumCargas = .Num_Cargas
        vlIndCobertura = .Ind_Cob
        vlIndBenSocial = .Cod_BenSocial
        
        vlCodMoneda = .Cod_Moneda
        If (vgMonedaCodOfi = .Cod_Moneda) Then
            vlMtoValMoneda = cgMonedaValorNS
        Else
            vlMtoValMoneda = .Mto_ValMoneda
        End If
        
'I--- ABV 05/02/2011 ---
        vlCodTipReajuste = .Cod_TipReajuste
        vlMtoValReajusteTri = .Mto_ValReajusteTri
        vlMtoValReajusteMen = .Mto_ValReajusteMen
'F--- ABV 05/02/2011 ---
        
        vlPriUniMod = .Mto_PriUniMod
        vlCtaIndMod = .Mto_CtaIndMod
        vlMtoBonoMod = .Mto_BonoMod
'        vlMtoApoAdiMod = .Mto_ApoAdiMod
        If (vgMonedaCodOfi = .Cod_Moneda) Then
            vlMtoApoAdiMod = .Mto_ApoAdiMod
        Else
            vlMtoApoAdiMod = Format(.Mto_ApoAdiMod / vlMtoValMoneda, "#0.00")
        End If
        
        vlCodTipRen = (.Cod_TipRen)
        vlNumMesDif = (.Num_MesDif)
        vlCodMod = (.Cod_Modalidad)
        vlMesGar = (.Num_MesGar)
        vlCodCoberCon = (.Cod_CoberCon)
        vlMtoFacPenElla = (.Mto_FacPenElla)
        vlPrcFacPenElla = (.Prc_FacPenElla)
        vlDerCre = (.Cod_DerCre)
        vlDerGra = (.Cod_DerGra)
        
        vlRtaAfp = (.Prc_RentaAFP)
        vlPrcRentaAfpOri = (.Prc_RentaAFP)
        vlRtaTmp = (.Prc_RentaTMP)
        vlMtoCuoMor = (.Mto_CuoMor)
        
        vlTasaTCE = (.Prc_TasaCe)
        vlTasaVta = (.Prc_TasaVta)
        vlTasaTir = (.Prc_TasaTir)
        vlTasaPG = (.Prc_TasaIntPerGar)
        vlCNU = (.Mto_CNU)
        vlPriUniSim = (.Mto_PriUniSim)
        vlPriUniDif = (.Mto_PriUniDif)
        vlPension = (.Mto_Pension)
        vlPenGar = (.Mto_PensionGar)
        vlCtaIndAfp = (.Mto_CtaIndAFP)
        vlRtaTmpAFP = (.Mto_RentaTMPAFP)
        vlResMat = (.Mto_ResMat)
        vlValPPTmp = (.Mto_ValPrePenTmp)
        vlMtoPerCon = (.Mto_PerCon)
        vlPrcPerCon = (.Prc_PerCon)
        
        vlMtoSumPension = (.Mto_SumPension)
        vlMtoPenAnual = (.Mto_PenAnual)
        vlMtoRMPension = (.Mto_RMPension)
        vlMtoRMGtoSep = (.Mto_RMGtoSep)
        vlMtoRMGtoSepRV = (.Mto_RMGtoSepRV)
        vlMtoAjusteIPC = .Mto_AjusteIPC
        
        vlCodTipCot = (.Cod_TipoCot)
        vlCodEstCot = (.Cod_EstCot)
        
        If vlCodTipRen = "6" Then
            vlFecFinPerDif = fgCalcularFechaFinPerDiferido(vlFecDev, CLng(vlNumMesDif))
            vlFecFinPerGar = fgCalcularFechaFinPerGarantizado(vlFecDev, 0, CLng(vlMesGar))
            vlFecIniPagoPen = vlFecDev
        Else
            vlFecFinPerDif = fgCalcularFechaFinPerDiferido(vlFecDev, CLng(vlNumMesDif))
            vlFecFinPerGar = fgCalcularFechaFinPerGarantizado(vlFecDev, CLng(vlNumMesDif), CLng(vlMesGar))
            vlFecIniPagoPen = fgCalcularFechaIniPagoPensiones(vlFecDev, CLng(vlNumMesDif))
        End If
        
'I--- ABV 04/12/2009 ---
        vlMarcaSobDif = Trim(.Ind_CalSobDif)
'F--- ABV 04/12/2009 ---

    End With

    vlReCalculo = "S"
End If

    'Datos que se pueden registrar desde la Pantalla y no afectan al cálculo de la Pensión
    vlTipoIden = fgObtenerCodigo_TextoCompuesto(Cmb_TipoIdent) 'Trim(Format(Lbl_Cuspp, "#0"))
    vlNumIden = Trim(Txt_NumIdent)
    vlGlsDir = Trim(Left(Lbl_Dir, 50))
    vlCodDir = vlCodDireccion
    vlCodIsapre = fgObtenerCodigo_TextoCompuesto(Cmb_Salud)  '(vgRs!cod_isapre)
    vlCodVejez = fgObtenerCodigo_TextoCompuesto(Cmb_Vejez) '(vgRs!cod_vejez)
    vlEstCivil = fgObtenerCodigo_TextoCompuesto(Cmb_EstCivil) '(vgRs!cod_estcivil)
    vlFono = Trim(Txt_Fono)
    vlFono2 = Trim(Me.Txt_Fono2_Afil)
    vlCorreo = Trim(Txt_Correo)
    vlCodViaPago = Trim(Mid(Cmb_ViaPago.Text, 1, (InStr(1, Cmb_ViaPago, "-") - 1)))
    vlCodTipCta = Trim(Mid(Cmb_TipCta.Text, 1, (InStr(1, Cmb_TipCta, "-") - 1)))
    vlCodBco = Trim(Mid(Cmb_Bco.Text, 1, (InStr(1, Cmb_Bco, "-") - 1)))
    vlCodMonCta = Trim(Mid(cmb_MonCta.Text, 1, (InStr(1, cmb_MonCta, "-") - 1)))
    vlNumCta = Trim(Txt_NumCta)
    vlCodSuc = Trim(Mid(Cmb_Suc.Text, 1, (InStr(1, Cmb_Suc, "-") - 1)))
    vlNum_Cuenta_CCI = Trim(txt_CCI)
    
    
    'Identificación del Corredor
    If Lbl_TipoIdentCorr <> "" Then
        vlTipoIdencor = fgObtenerCodigo_TextoCompuesto(Lbl_TipoIdentCorr)
    Else
        vlTipoIdencor = ""
    End If
    vlNumIdenCor = Trim(Lbl_NumIdentCorr)
        
    vlCodUsuario = (vgUsuario)
    vlSucUsu = vgUsuarioSuc
    vlSucCorredor = fgBuscarSucCorredor(vlTipoIdencor, vlNumIdenCor)
    
    vlFecPriPago = Format(Trim(Txt_FecIniPago), "yyyymmdd")
    
    If (Txt_Nacionalidad <> "") Then
        vlNacionalidad = Trim(Txt_Nacionalidad)
    Else
        vlNacionalidad = cgTipoNacionalidad
    End If
    
    'RVF 20090914
    vlTipVia = Left(Cmb_TipoVia.Text, 2)
    vlNomVia = Trim(Txt_NombreVia.Text)
    vlNumDmc = Trim(Txt_Numero.Text)
    vlIntDmc = Trim(Txt_Interior.Text)
    vlTipZon = Left(Cmb_TipoZona.Text, 2)
    vlNomZon = Trim(Txt_NombreZona.Text)
    vlReferencia = Trim(Txt_Referencia.Text)
    '*****
    
    'SE GRABA LA POLIZA
    vlSql = ""
    vlSql = "INSERT INTO pd_TMAE_ORIPOLIZA ("
    vlSql = vlSql & "num_poliza,num_cot,num_correlativo,num_operacion,"
    vlSql = vlSql & "num_archivo,cod_afp,cod_isapre,cod_tippension,"
    vlSql = vlSql & "cod_vejez,cod_estcivil,cod_cuspp,cod_tipoidenafi,num_idenafi,gls_direccion,"
    vlSql = vlSql & "cod_direccion,gls_fono,gls_correo,cod_viapago,cod_tipcuenta,"
    vlSql = vlSql & "cod_banco,num_cuenta,cod_sucursal,fec_solicitud,"
    vlSql = vlSql & "fec_vigencia,fec_dev,fec_acepta,cod_monedafon,mto_monedafon,"
    vlSql = vlSql & "mto_priunifon,mto_ctaindfon,mto_bonofon,mto_apoadi,"
    vlSql = vlSql & "mto_priuni,mto_ctaind,mto_bono,prc_tasarprt,"
    vlSql = vlSql & "cod_tipoidencor,num_idencor,"
    vlSql = vlSql & "prc_corcom,prc_corcomreal,mto_corcom,num_annojub,"
    vlSql = vlSql & "num_cargas,ind_cob,cod_bensocial,cod_moneda,mto_valmoneda,"
    vlSql = vlSql & "mto_priunimod,mto_ctaindmod,mto_bonomod,"
    vlSql = vlSql & "cod_tipren,num_mesdif,cod_modalidad,num_mesgar,cod_cobercon,"
    vlSql = vlSql & "mto_facpenella,prc_facpenella,cod_dercre,cod_dergra,"
    vlSql = vlSql & "prc_rentaafp,prc_rentaafpori,prc_rentatmp,"
    vlSql = vlSql & "mto_cuomor,"
    vlSql = vlSql & "prc_tasace,prc_tasavta,prc_tasatir,prc_tasapergar,mto_cnu,"
    vlSql = vlSql & "mto_priunisim,mto_priunidif,mto_pension,mto_pensiongar,"
    vlSql = vlSql & "mto_ctaindafp,mto_rentatmpafp,mto_resmat,mto_valprepentmp,"
    vlSql = vlSql & "mto_percon,prc_percon,mto_sumpension,mto_penanual,"
    vlSql = vlSql & "mto_rmpension,mto_rmgtosep,"
    vlSql = vlSql & "mto_rmgtoseprv,"
    vlSql = vlSql & "cod_tipcot,cod_estcot,cod_usuario,cod_sucursalusu,"
    vlSql = vlSql & "cod_usuariocrea, fec_crea, hor_crea,"
    vlSql = vlSql & "cod_succorredor,fec_finperdif,fec_finpergar,"
    vlSql = vlSql & "gls_nacionalidad,ind_recalculo,fec_pripago "
    vlSql = vlSql & ",fec_emision,fec_inipencia "
    vlSql = vlSql & ",fec_calculo,mto_ajusteIPC "
    vlSql = vlSql & ",mto_apoadifon,mto_apoadimod "
    vlSql = vlSql & ",cod_tipvia,gls_nomvia,gls_numdmc,gls_intdmc"
    vlSql = vlSql & ",cod_tipzon,gls_nomzon,gls_referencia"
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & ",cod_tipreajuste,mto_valreajustetri,mto_valreajustemen, fec_devsol, cod_nacionalidad "
'F--- ABV 05/02/2011 ---
    'GCP-FRACTAL 08042019
    vlSql = vlSql & ", NUM_CUENTA_CCI, COD_MONCTA"
    vlSql = vlSql & ", num_idensup, num_idenjef, GLS_FONO2"
    vlSql = vlSql & ") VALUES ("
    vlSql = vlSql & "'" & Trim(vlNumPol) & "', "
    vlSql = vlSql & "'" & Trim(vlNumCot) & "', "
    vlSql = vlSql & "" & vlNumCorrelativo & ", "
    vlSql = vlSql & "'" & Trim(vlCodSolOfe) & "', "
    vlSql = vlSql & "" & vlNumArchivo & ", "
    vlSql = vlSql & "'" & Trim(vlCodAFP) & "',"
    If Trim(vlCodIsapre) <> "" Then
        vlSql = vlSql & "'" & Trim(vlCodIsapre) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    vlSql = vlSql & "'" & Trim(vlCodTipPen) & "',"
    If Trim(vlCodVejez) <> "" Then
        vlSql = vlSql & "'" & Trim(vlCodVejez) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    vlSql = vlSql & "'" & Trim(vlEstCivil) & "',"
    vlSql = vlSql & "'" & (vlcuspp) & "',"
    vlSql = vlSql & " " & (vlTipoIden) & ","
    If Trim(vlNumIden) <> "" Then
        vlSql = vlSql & "'" & Trim(vlNumIden) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    If Trim(vlGlsDir) <> "" Then
        vlSql = vlSql & "'" & Trim(vlGlsDir) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    vlSql = vlSql & " " & vlCodDir & ","
    If Trim(vlFono) <> "" Then
        vlSql = vlSql & "'" & Trim(vlFono) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    If Trim(vlCorreo) <> "" Then
        vlSql = vlSql & "'" & Trim(vlCorreo) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    If Trim(vlCodViaPago) <> "" Then
        vlSql = vlSql & "'" & Trim(vlCodViaPago) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    If Trim(vlCodTipCta) <> "" Then
        vlSql = vlSql & "'" & Trim(vlCodTipCta) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    If Trim(vlCodBco) <> "" Then
        vlSql = vlSql & "'" & Trim(vlCodBco) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    If Trim(vlNumCta) <> "" Then
        vlSql = vlSql & "'" & Trim(vlNumCta) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    If Trim(vlCodSuc) <> "" Then
        vlSql = vlSql & "'" & Trim(vlCodSuc) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    vlSql = vlSql & "'" & Trim(vlFecSolicitud) & "',"
    vlSql = vlSql & "'" & Trim(vlFecVig) & "',"
    vlSql = vlSql & "'" & Trim(vlFecDev) & "',"
    vlSql = vlSql & "'" & Trim(vlFecIncorporacion) & "',"
    
    If Trim(vlCodMonedaFon) <> "" Then
        vlSql = vlSql & "'" & Trim(vlCodMonedaFon) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    vlSql = vlSql & " " & str(vlMtoValMonedaFon) & ","
    vlSql = vlSql & " " & str(vlPriUniFon) & ","
    vlSql = vlSql & " " & str(vlCtaIndFon) & ","
    vlSql = vlSql & " " & str(vlMtoBonoFon) & ","
    vlSql = vlSql & " " & str(vlMtoApoAdi) & ","
    vlSql = vlSql & " " & str(vlPriUni) & ","
    vlSql = vlSql & " " & str(vlCtaInd) & ","
    vlSql = vlSql & " " & str(vlMtoBono) & ","
    vlSql = vlSql & " " & str(vlTasaRetPro) & ","
    vlSql = vlSql & "'" & (vlTipoIdencor) & "',"
    If Trim(vlNumIdenCor) <> "" Then
        vlSql = vlSql & "'" & Trim(vlNumIdenCor) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    vlSql = vlSql & " " & str(vlPrcCorCom) & ","
    vlSql = vlSql & " " & str(vlPrcCorComReal) & ","
    vlSql = vlSql & " " & str(vlMtoCorCom) & ","
    vlSql = vlSql & " " & str(vlNumAnnoJub) & " ,"
    vlSql = vlSql & " " & str(vlNumCargas) & " ,"
    vlSql = vlSql & "'" & (vlIndCobertura) & "',"
    vlSql = vlSql & "'" & (vlIndBenSocial) & "',"
    If Trim(vlCodMoneda) <> "" Then
        vlSql = vlSql & "'" & Trim(vlCodMoneda) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    vlSql = vlSql & " " & str(vlMtoValMoneda) & ","
    vlSql = vlSql & " " & str(vlPriUniMod) & ","
    vlSql = vlSql & " " & str(vlCtaIndMod) & ","
    vlSql = vlSql & " " & str(vlMtoBonoMod) & ","
    vlSql = vlSql & "'" & Trim(vlCodTipRen) & "',"
    vlSql = vlSql & " " & str(vlNumMesDif) & " ,"
    vlSql = vlSql & "'" & Trim(vlCodMod) & "',"
    vlSql = vlSql & " " & str(vlMesGar) & " ,"
    vlSql = vlSql & "'" & Trim(vlCodCoberCon) & "',"
    vlSql = vlSql & " " & str(vlMtoFacPenElla) & " ,"
    vlSql = vlSql & " " & str(vlPrcFacPenElla) & " ,"
    vlSql = vlSql & "'" & Trim(vlDerCre) & "',"
    vlSql = vlSql & "'" & Trim(vlDerGra) & "',"
    vlSql = vlSql & " " & str(vlRtaAfp) & " ,"
    vlSql = vlSql & " " & str(vlPrcRentaAfpOri) & " ,"
    vlSql = vlSql & " " & str(vlRtaTmp) & " ,"
    vlSql = vlSql & " " & str(vlMtoCuoMor) & " ,"
    vlSql = vlSql & " " & str(vlTasaTCE) & " ,"
    vlSql = vlSql & " " & str(vlTasaVta) & " ,"
    vlSql = vlSql & " " & str(vlTasaTir) & " ,"
    vlSql = vlSql & " " & str(vlTasaPG) & " ,"
    vlSql = vlSql & " " & str(vlCNU) & " ,"
    vlSql = vlSql & " " & str(vlPriUniSim) & " ,"
    vlSql = vlSql & " " & str(vlPriUniDif) & " ,"
    vlSql = vlSql & " " & str(vlPension) & " ,"
    vlSql = vlSql & " " & str(vlPenGar) & " ,"
    vlSql = vlSql & " " & str(vlCtaIndAfp) & " ,"
    vlSql = vlSql & " " & str(vlRtaTmpAFP) & " ,"
    vlSql = vlSql & " " & str(vlResMat) & " ,"
    vlSql = vlSql & " " & str(vlValPPTmp) & " ,"
    vlSql = vlSql & " " & str(vlMtoPerCon) & " ,"
    vlSql = vlSql & " " & str(vlPrcPerCon) & " ,"
    
    vlSql = vlSql & " " & str(vlMtoSumPension) & " ,"
    vlSql = vlSql & " " & str(vlMtoPenAnual) & " ,"
    vlSql = vlSql & " " & str(vlMtoRMPension) & " ,"
    vlSql = vlSql & " " & str(vlMtoRMGtoSep) & " ,"
    vlSql = vlSql & " " & str(vlMtoRMGtoSepRV) & " ,"
    
    vlSql = vlSql & "'" & Trim(vlCodTipCot) & "',"
    vlSql = vlSql & "'" & Trim(vlCodEstCot) & "',"
        
    If Trim(vlCodUsuario) <> "" Then
        vlSql = vlSql & "'" & Trim(vlCodUsuario) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    If Trim(vlSucUsu) <> "" Then
        vlSql = vlSql & "'" & Trim(vlSucUsu) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
        
    vlSql = vlSql & "'" & vgUsuario & "',"
    vlSql = vlSql & "'" & Format(Date, "yyyymmdd") & "', "
    vlSql = vlSql & "'" & Format(Time, "hhmmss") & "',"
        
    If Trim(vlSucCorredor) <> "" Then
        vlSql = vlSql & "'" & Trim(vlSucCorredor) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
        
    If Trim(vlFecFinPerDif) <> "" Then
        vlSql = vlSql & "'" & Trim(vlFecFinPerDif) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
        
    If Trim(vlFecFinPerGar) <> "" Then
        vlSql = vlSql & "'" & Trim(vlFecFinPerGar) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
        
    vlSql = vlSql & "'" & Trim(vlNacionalidad) & "',"
    vlSql = vlSql & "'" & Trim(vlReCalculo) & "',"
    vlSql = vlSql & "'" & Trim(vlFecPriPago) & "' "
    vlSql = vlSql & ",'" & Trim(vlFecEmision) & "' "
    vlSql = vlSql & ",'" & Trim(vlFecIniPagoPen) & "' "
    vlSql = vlSql & ",'" & Trim(vlFecCalculo) & "' "
    vlSql = vlSql & "," & str(vlMtoAjusteIPC) & " "
    
    vlSql = vlSql & "," & str(vlMtoBonoMod) & " "
    vlSql = vlSql & "," & str(vlMtoApoAdiMod) & ", "
    
    'RVF 20090914
    vlSql = vlSql & "'" & vlTipVia & "',"
    vlSql = vlSql & "'" & vlNomVia & "',"
    vlSql = vlSql & "'" & vlNumDmc & "',"
    vlSql = vlSql & "'" & vlIntDmc & "',"
    vlSql = vlSql & "'" & vlTipZon & "',"
    vlSql = vlSql & "'" & vlNomZon & "',"
    vlSql = vlSql & "'" & vlReferencia & "' "
    '*****
    
'I--- ABV 05/02/2011 ---
    vlSql = vlSql & ",'" & vlCodTipReajuste & "',"
    vlSql = vlSql & " " & str(vlMtoValReajusteTri) & ","
    vlSql = vlSql & " " & str(vlMtoValReajusteMen) & ", "
'F--- ABV 05/02/2011 ---
    'RRR 18/9/13
    
    vlSql = vlSql & " '" & Trim(vlfecDevSol) & "', "
    '-- Begin : Modify by : ricardo.huerta 11-12-2018
    vlSql = vlSql & " '" & vl_nacionalidad & "', "
    '-- End    :  Modify by : ricardo.huerta 11-12-2018
    'GCP-FRACTAL 08042019
     vlSql = vlSql & " '" & vlNum_Cuenta_CCI & "', "
     vlSql = vlSql & " '" & vlCodMonCta & "', "
     
    If Trim(vlNum_idensup) <> "" Then
        vlSql = vlSql & "'" & Trim(vlNum_idensup) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
    
    If Trim(vlNum_idenjef) <> "" Then
        vlSql = vlSql & "'" & Trim(vlNum_idenjef) & "',"
    Else
        vlSql = vlSql & "NULL" & ","
    End If
     
    If Trim(vlFono2) <> "" Then
        vlSql = vlSql & "'" & Trim(vlFono2) & "'"
    Else
        vlSql = vlSql & "NULL" & ""
    End If
    
     vlSql = vlSql & ") "

    vgConectarBD.Execute (vlSql)

    'SE GRABAN LOS BENEFICIARIOS
'    Call flGrabaBeneficiario
    If vlBotonEscogido = "C" Then
        Call flGrabaBeneficiarioCot
    Else
        Call flGrabaBeneficiario
    End If
    If Trim(Lbl_Representante.Caption) <> "" Then
        Call pGrabaRepresentante
    End If
    
'    'SE GRABA EL BONO
''I--- ABV 23/06/2006 ---
''Determinar realmente si por Tipo de Cobertura se debe grabar el Bono
''Por ahora dejare que lo guarde para todos los que lo tengan
''    If vlCodTipPen = "05" Then
'        Call flGrabaBono
''    End If
''F--- ABV 23/06/2006 ---
    
    'Modificar Estado de la Cotización
    vlSql = ""
'I--- ABV 23/06/2006 ---
'    vlSql = "SELECT num_cot FROM pt_tmae_detcotizacion WHERE "
    vlSql = "SELECT num_cot FROM " & vlTablaDetCotizacion & " WHERE "
'F--- ABV 23/06/2006 ---
    vlSql = vlSql & "num_cot = '" & Trim(vlNumCot) & "' AND "
    vlSql = vlSql & "cod_estcot = '" & clCodEstCotA & "' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlSql = ""
'I--- ABV 23/06/2006 ---
'        vlSql = " UPDATE pt_tmae_detcotizacion SET "
        vlSql = " UPDATE " & vlTablaDetCotizacion & " SET "
'F--- ABV 23/06/2006 ---
        vlSql = vlSql & "cod_estcot = '" & clCodEstCotP & "' "
        vlSql = vlSql & "WHERE num_cot = '" & vlNumCot & "' AND "
        vlSql = vlSql & "cod_estcot = '" & clCodEstCotA & "' "
        vgConectarBD.Execute (vlSql)
    End If
    vgRs.Close

End Function

Function flImprimirPoliza(iNumPol)
Dim vlPriUni As Double
Dim vlCobertura As String
Dim rs As ADODB.Recordset

   vlArchivo = strRpt & "PD_Rpt_Poliza.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Póliza no se encuentra en el directorio de la aplicación", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Function
   End If

'   busca el Apoderado
'   vlSql = "SELECT gls_nomapo, gls_carapo FROM pd_TMAE_APODERADO "
'   Set vgRs = vgConexionBD.Execute(vlSql)
'   If Not vgRs.EOF And Not IsNull(vgRs!gls_nomApo) Then
'     vlApoderado = vgRs!gls_nomApo
'     vlCargo = vgRs!gls_carApo
'   Else
'        MsgBox "Apoderado Inexistente", vbCritical, "Datos Incompletos"
'        Exit Function
'   End If
'   vgRs.Close

'  numero de póliza
   vlNumPol = Trim(iNumPol)
   
   'busca el nombre del Afiliado
   vlSql = ""
   vlSql = "SELECT gls_nomben, gls_patben, gls_matben "
   vlSql = vlSql & "FROM pd_tmae_oripolben WHERE "
   vlSql = vlSql & "num_poliza = '" & iNumPol & "' AND cod_par = '99' "
   Set vgRs = vgConexionBD.Execute(vlSql)
   If Not vgRs.EOF Then
       vlNomAfi = Trim(vgRs!Gls_NomBen) & " " & Trim(vgRs!Gls_PatBen) & " " & Trim(vgRs!Gls_MatBen)
   Else
       MsgBox "Nombre del Afiliado No encontrado o no exsiste", vbCritical, "Datos Incompletos"
       Screen.MousePointer = 0
       Exit Function
   End If
 
   'codigo isapre
   vlCodIsapre = Trim(Mid(Cmb_Salud.Text, (InStr(1, Cmb_Salud, "-") + 1), Len(Cmb_Salud)))
   
   'vlCodScomp = fgObtenerCodMonedaScomp(egTablaMoneda(), vgMonedaCodOf, vgRs!cod_Moneda)
   
   'codigo afp
   vlCodAFP = Trim(Mid(Lbl_Afp, (InStr(1, Lbl_Afp, "-") + 1), Len(Lbl_Afp)))
      
   'Buscar la descripcion Tipo de Moneda
   vlSql = ""
   vlSql = "SELECT gls_elemento FROM ma_tpar_tabcod "
   vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_TipMon & "' "
   vlSql = vlSql & "AND cod_elemento = (SELECT cod_moneda FROM "
   vlSql = vlSql & "pd_tmae_oripoliza WHERE num_poliza = '" & iNumPol & "')"
   Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) And Not IsNull(vgRs!gls_elemento) Then
        vlCodMoneda = vgRs!gls_elemento
   Else
        MsgBox "Tipo de Moneda no Encontrada", vbCritical, "Datos Incompleos"
        Screen.MousePointer = 0
        Exit Function
   End If
   vgRs.Close

   'busca tipo pension
    vlSql = ""
    vlSql = "SELECT gls_elemento FROM ma_tpar_tabcod "
    vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_TipPen & "' "
    vlSql = vlSql & "AND cod_elemento = (SELECT cod_tippension FROM "
    vlSql = vlSql & "pd_tmae_oripoliza WHERE num_poliza = '" & iNumPol & "')"
    Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) And Not IsNull(vgRs!gls_elemento) Then
        vlCodTipPen = vgRs!gls_elemento
   Else
        MsgBox "Tipo de Pensión no encontrado", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        Exit Function
   End If
   vgRs.Close
   
   'busca la glosa segun codigo de  renta
    vlSql = ""
    vlSql = "SELECT gls_elemento FROM ma_tpar_tabcod "
    vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_TipRen & "' "
    vlSql = vlSql & "AND cod_elemento = (SELECT cod_tipren FROM "
    vlSql = vlSql & "pd_tmae_oripoliza WHERE num_poliza = '" & iNumPol & "')"
    Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) And Not IsNull(vgRs!gls_elemento) Then
        vlCodTipRen = vgRs!gls_elemento
   Else
        MsgBox "Tipo de Renta no encontrado", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        Exit Function
   End If
   vgRs.Close
   
   'busca el numero de liquidación
   vlSql = ""
   vlSql = "SELECT cod_liquidacion FROM pd_tmae_oripolben a, "
   vlSql = vlSql & "pd_tmae_polprirec b where a.num_poliza=b.num_poliza"
   Set vgRs = vgConexionBD.Execute(vlSql)
   If Not (vgRs.EOF) Then
         vlNumLiquidacion = vgRs!cod_liquidacion
   End If
   vgRs.Close
   
   vlNumTipIden = Txt_NumIdent
   vlNacionalidad = Txt_Nacionalidad
   vlCobConyuge = Lbl_FacPenElla
   vlIndDeCre = Lbl_DerCre
   vlIndGra = Lbl_DerGra
   vlMesGar = Lbl_MesesGar
   vlAñoGar = Lbl_AnnosDif
   vlFecNac = Lbl_FecNac
   vlParentesco = Lbl_Par

   vlCobertura = vlCodTipRen
   vlModalidad = fgObtenerCodigo_TextoCompuesto(Lbl_Alter)
   vlGlosaModalidad = Mid(Lbl_Alter, InStr(1, Lbl_Alter, "-") + 1, Len(Lbl_Alter))
   vlGlosaCobConyuge = fgBuscarGlosaCobConyuge(vlCobConyuge)
    
   vlNombreCompania = UCase(vgNombreCompania)
    
   'Sucursal         RVF 20090914
   vlNombreSucursal = fgObtenerNombreSuc_Usuario(vgUsuarioSuc)
   vlRepresentante = IIf(Trim(Lbl_Representante.Caption) <> "", Lbl_Representante.Caption, "")
   vlDocum = IIf(Trim(Cmb_TipIdRep.Text) <> "", Mid(Trim(Cmb_TipIdRep.Text), 5), "") & " " & IIf(Trim(Txt_NumIdRep.Text) <> "", Trim(Txt_NumIdRep.Text), "")
   '*****
   
   If vlModalidad = 1 Then
       If vlGlosaModalidad <> "" Then
           vlCobertura = vlCobertura & " " & vlGlosaModalidad
       End If
   Else
       If vlGlosaModalidad <> "" Then
           vlCobertura = vlCobertura & " CON P. " & vlGlosaModalidad
       End If
   End If
   If vlCobConyuge <> 0 Then
       If vlGlosaCobConyuge <> "" Then
           vlCobertura = vlCobertura & " CON " & vlGlosaCobConyuge
       End If
   End If
   If Mid(vlIndDeCre, 1, 1) = "S" Then
       vlCobertura = vlCobertura & " CON D.CRECER"
   End If
   
   If Mid(vlIndGra, 1, 1) = "S" Then
       vlCobertura = vlCobertura & " Y CON GRATIFICACIÓN"
   End If
   
   'Obtiene el mes en palabra, y la envía al reporte.
   vlfechahoy = Date
   vlMes = MonthName(Month(Date), False)
   vlDia = Day(Date)
   vlAño = Year(Date)
   vlfechapalabra = vlDia & " de " & vlMes & " de " & vlAño
  
   'vgQuery = "{pd_TMAE_ORIPOLIZA.NUM_POLIZA} = '" & vlNumPol & "'"
   
   'Set RS = New ADODB.Recordset
   'RS.CursorLocation = adUseClient
   
   
   vgQuery = "select a.NUM_POLIZA, a.GLS_DIRECCION, a.GLS_FONO, a.FEC_DEV, MTO_APOADI, MTO_PRIUNI, MTO_CTAIND, NUM_MESDIF, "
   vgQuery = vgQuery & " COD_MODALIDAD, NUM_MESGAR, a.COD_DERCRE, COD_DERGRA, a.MTO_PENSION, FEC_FINPERDIF, FEC_FINPERGAR, GLS_NACIONALIDAD,"
   vgQuery = vgQuery & " FEC_INIPENCIA, b.NUM_ORDEN, b.COD_SEXO, b.COD_SITINV, NUM_IDENBEN, GLS_NOMBEN, GLS_NOMSEGBEN, GLS_PATBEN, GLS_MATBEN,"
   vgQuery = vgQuery & " FEC_NACBEN, b.MTO_PENSION pensionben, c.cod_cobercon, gls_cobercon, gls_comuna, c.cod_scomp, gls_provincia, gls_region, h.gls_elemento,"
   vgQuery = vgQuery & " i.gls_tipoiden ben, j.gls_tipoiden afi, num_idenafi, a.cod_moneda, e.cod_scomp monajus, k.gls_elemento parentesco, mto_bonofon, mto_apoadi,"
   vgQuery = vgQuery & " b.GLS_FONO2, a.GLS_CORREO" 'JVB 202104267 Nuevos campos a mostrar en el reporte
   vgQuery = vgQuery & " from pd_tmae_oripoliza a"
   vgQuery = vgQuery & " join pd_tmae_oripolben b on a.num_poliza=b.num_poliza"
   vgQuery = vgQuery & " join ma_tpar_cobercon c on a.cod_cobercon=c.cod_cobercon"
   vgQuery = vgQuery & " join ma_tpar_comuna d on a.cod_direccion=d.cod_direccion"
   vgQuery = vgQuery & " join ma_tpar_monedatiporeaju e on a.cod_moneda=e.cod_moneda and a.cod_tipreajuste=e.cod_tipreajuste"
   vgQuery = vgQuery & " join ma_tpar_provincia f on d.cod_provincia=f.cod_provincia"
   vgQuery = vgQuery & " join ma_tpar_region g on f.cod_region=g.cod_region"
   vgQuery = vgQuery & " join ma_tpar_tabcod h on a.cod_modalidad=h.cod_elemento and cod_tabla='AL'"
   vgQuery = vgQuery & " join ma_tpar_tipoiden i on b.cod_tipoidenben=i.cod_tipoiden"
   vgQuery = vgQuery & " join ma_tpar_tipoiden j on a.cod_tipoidenafi=j.cod_tipoiden"
   vgQuery = vgQuery & " join ma_tpar_tabcod k on b.cod_par=k.cod_elemento and k.cod_tabla='PA'"
   vgQuery = vgQuery & " where a.num_poliza='" & vlNumPol & "'"
   'RS.Open vgQuery
   
   Set vgRs = vgConexionBD.Execute(vgQuery)
   'RS.Open "PP_LISTA_BIENVENIDA.LISTAR('" & Txt_Poliza.Text & "','" & clCodTipPensionSob & "','" & Trim(vlCodDerPen) & "','" & Trim(vlCodPar) & "', '" & Trim(Txt_Endoso) & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
   Dim LNGa As Long
   LNGa = CreateFieldDefFile(vgRs, Replace(UCase(strRpt & "Estructura\PD_Rpt_Poliza.rpt"), ".RPT", ".TTX"), 1)

   If objRep.CargaReporte(strRpt & "", "PD_Rpt_Poliza.rpt", "PrePoliza", vgRs, True, _
                            ArrFormulas("NombreAfi", vlNomAfi), _
                            ArrFormulas("TipoPension", vlCodTipPen), _
                            ArrFormulas("MesGar", vlMesGar), _
                            ArrFormulas("NombreCompania", vlNombreCompania), _
                            ArrFormulas("Concatenar", vlCobertura), _
                            ArrFormulas("Sucursal", "Surquillo"), _
                            ArrFormulas("RepresentanteNom", vlRepresentante), _
                            ArrFormulas("RepresentanteDoc", vlDocum), _
                            ArrFormulas("CodTipPen", Mid(Lbl_TipPen.Caption, 1, 2)), _
                            ArrFormulas("InsSalud", vlCodIsapre), _
                            ArrFormulas("Afp", vlCodAFP)) = False Then
            
      'If objRep.CargaReporte(strRpt & "", "PD_Rpt_Poliza.rpt", "PrePoliza", vgRs, True) = False Then
            
       MsgBox "No se pudo abrir el reporte", vbInformation
       'Exit Sub
    End If
   
   
'   Rpt_Poliza.Reset
'   Rpt_Poliza.WindowState = crptMaximized
'   Rpt_Poliza.ReportFileName = vlArchivo
'   Rpt_Poliza.Connect = vgRutaDataBase
'   Rpt_Poliza.SelectionFormula = ""
'   Rpt_Poliza.SelectionFormula = vgQuery
'
'   Rpt_Poliza.Formulas(0) = ""
'   Rpt_Poliza.Formulas(1) = ""
'   Rpt_Poliza.Formulas(2) = ""
'   Rpt_Poliza.Formulas(3) = ""
'   Rpt_Poliza.Formulas(4) = ""
'   Rpt_Poliza.Formulas(5) = ""
'   Rpt_Poliza.Formulas(6) = ""
'   Rpt_Poliza.Formulas(7) = ""
'   Rpt_Poliza.Formulas(8) = ""
'   Rpt_Poliza.Formulas(9) = ""
'   Rpt_Poliza.Formulas(10) = ""
'
'   'Rpt_Poliza.Formulas(0) = "InsSalud = '" & vlCodIsapre & "'"
'   'Rpt_Poliza.Formulas(1) = "Afp = '" & vlCodAFP & "'"
'   'Rpt_Poliza.Formulas(2) = "GlsMoneda = '" & vlCodMoneda & "'"
'   'Rpt_Poliza.Formulas(5) = "Num_pol= '" & vlNumPol & "'"
'   'Rpt_Poliza.Formulas(7) = "NombreSistema= '" & vgNombreSistema & "'"
'   'Rpt_Poliza.Formulas(8) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
'   'Rpt_Poliza.Formulas(9) = "NumTipIden='" & vlNumTipIden & "'"
'   '''Rpt_Poliza.Formulas(10) = "FecNac='" & vlFecNac & "'"
'   'Rpt_Poliza.Formulas(11) = "Nacionalidad='" & vlNacionalidad & "'"
'   'Rpt_Poliza.Formulas(12) = "CobConyuge='" & vlCobConyuge & "'"
'   'Rpt_Poliza.Formulas(13) = "IndCrecer='" & vlIndDeCre & "'"
'   'Rpt_Poliza.Formulas(14) = "IndGratificacion='" & vlIndGra & "'"
'   'Rpt_Poliza.Formulas(16) = "AnnosDif='" & vlAñoGar & "'"
'   'Rpt_Poliza.Formulas(17) = "GlosaTipoRenta='" & vlCodTipRen & "'"
'   'Rpt_Poliza.Formulas(19) = "Mes = '" & vlfechapalabra & "'"
'   'Rpt_Poliza.Formulas(20) = "NumLiquidacion='" & vlNumLiquidacion & "'"
'
'   'RVF 20090914
'   Rpt_Poliza.Formulas(0) = "NombreAfi = '" & vlNomAfi & "'"
'   Rpt_Poliza.Formulas(1) = "TipoPension = '" & vlCodTipPen & "'"
'   Rpt_Poliza.Formulas(2) = "MesGar='" & vlMesGar & " '"
'   Rpt_Poliza.Formulas(3) = "NombreCompania='" & vlNombreCompania & "'"
'   Rpt_Poliza.Formulas(4) = "Concatenar = '" & vlCobertura & "'"
'   Rpt_Poliza.Formulas(5) = "Sucursal = '" & vlNombreSucursal & "'"
'   Rpt_Poliza.Formulas(6) = "RepresentanteNom = '" & vlRepresentante & "'"
'   Rpt_Poliza.Formulas(7) = "RepresentanteDoc = '" & vlDocum & "'"
'   Rpt_Poliza.Formulas(8) = "CodTipPen = '" & Left(Trim(Lbl_TipPen.Caption), 2) & "'"
'   '*****
'
'   'DC 20091125
'   Rpt_Poliza.Formulas(9) = "InsSalud = '" & vlCodIsapre & "'"
'   Rpt_Poliza.Formulas(10) = "Afp = '" & vlCodAFP & "'"
'
'   Rpt_Poliza.SubreportToChange = ""
'   Rpt_Poliza.Destination = crptToWindow
'   Rpt_Poliza.WindowTitle = "Pre-Póliza"
'   Rpt_Poliza.Action = 1
'   Screen.MousePointer = 0
   
End Function

Function flEliminaCotizacion(iNumCot As String)
On Error GoTo Err_EliCot

    vlSql = "DELETE FROM tmae_boncot WHERE "
    vlSql = vlSql & "mid(num_cot,1,13) = '" & Mid(iNumCot, 1, 13) & "' and "
    vlSql = vlSql & "mid(num_cot,16,15) = '" & Mid(iNumCot, 16, 15) & "'"
    vgConectarBD.Execute (vlSql)
    
    vlSql = "DELETE FROM tmae_bencot where "
    vlSql = vlSql & "mid(num_cot,1,13) = '" & Mid(iNumCot, 1, 13) & "' and "
    vlSql = vlSql & "mid(num_cot,16,15) = '" & Mid(iNumCot, 16, 15) & "'"
    vgConectarBD.Execute (vlSql)
    
    vlSql = "DELETE FROM tmae_cotizacion where "
    vlSql = vlSql & "mid(num_cot,1,13) = '" & Mid(iNumCot, 1, 13) & "' and "
    vlSql = vlSql & "mid(num_cot,16,15) = '" & Mid(iNumCot, 16, 15) & "'"
    vgConectarBD.Execute (vlSql)
    
    vlSql = "DELETE FROM tmae_benpro where "
    vlSql = vlSql & "mid(num_cot,1,13) = '" & Mid(iNumCot, 1, 13) & "' and "
    vlSql = vlSql & "mid(num_cot,16,15) = '" & Mid(iNumCot, 16, 15) & "'"
    vgConectarBD.Execute (vlSql)
    
    vlSql = "DELETE FROM tmae_propuesta where "
    vlSql = vlSql & "mid(num_cot,1,13) = '" & Mid(iNumCot, 1, 13) & "' and "
    vlSql = vlSql & "mid(num_cot,16,15) = '" & Mid(iNumCot, 16, 15) & "'"
    vgConectarBD.Execute (vlSql)
Exit Function
Err_EliCot:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flHabilitaModif(iopcion As String)
On Error GoTo Err_HabMod
    If Trim(iopcion) = "H" Then
        Txt_NomAfi.Enabled = True
        Txt_ApPatAfi.Enabled = True
        Txt_ApMatAfi.Enabled = True
        Txt_Fono.Enabled = True
        Txt_Dir.Enabled = True
        Txt_Distrito.Enabled = True
        Txt_Ciudad.Enabled = True
        Cmb_SexoAfi.Enabled = True
        Cmb_Salud.Enabled = True
        
        Txt_NomAseg.Enabled = True
        txt_AppatAseg.Enabled = True
        txt_ApMatAseg.Enabled = True
        
        Cmd_Refrescar.Enabled = True
    Else
        If Trim(iopcion) = "D" Then
            Txt_NomAfi.Enabled = False
            Txt_ApPatAfi.Enabled = False
            Txt_ApMatAfi.Enabled = False
            Txt_Fono.Enabled = False
            Txt_Dir.Enabled = False
            Txt_Distrito.Enabled = False
            Txt_Ciudad.Enabled = False
            Cmb_SexoAfi.Enabled = False
            Cmb_Salud.Enabled = False
            
            Txt_NomAseg.Enabled = False
            txt_AppatAseg.Enabled = False
            txt_ApMatAseg.Enabled = False
            
            Cmd_Refrescar.Enabled = False
        End If
    End If
    
Exit Function
Err_HabMod:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'--------------------------------------------------------------------------
'MODIFICA UN BENEFICIARIO EN LA GRILLA
'--------------------------------------------------------------------------
Function flModificaGrilla()
Dim vlNombre As String, vlNombreSeg As String
Dim vlApBen  As String, vlAmBen     As String
Dim vlFecInv As String, vlCauInv    As String
Dim vlCorreoBen As String


Dim vlTipcta, vlMoncta, vlBancocta, vlCtanum As String
On Error GoTo Err_ModGri

    vlNumOrden = Lbl_NumOrden
    vlTipoIden = (Cmb_TipoIdentBen)
    vlNumIden = Trim(Txt_NumIdentBen)
    vlNombre = Trim(Txt_NombresBen)
    vlNombreSeg = Trim(Txt_NombresBenSeg)
    vlApBen = Trim(Txt_ApPatBen)
    vlAmBen = Trim(Txt_ApMatBen)
    vlCorreoBen = Trim(Me.txtCorreoBen.Text)
    vlFecInv = Trim(Txt_FecInvBen)
'    If IsDate(Txt_FecInvBen) Then vlFecInv = Format(Txt_FecInvBen, "yyyymmdd")
    vlCauInv = Trim(Mid(Lbl_CauInvBen, 1, InStr(1, Lbl_CauInvBen, "-") - 1))
    
    vlTipcta = Trim(Mid(cmbTipoCtaBen.Text, 1, InStr(1, cmbTipoCtaBen.Text, "-") - 1))
    vlMoncta = Trim(Mid(cmbMonctaBen.Text, 1, InStr(1, cmbMonctaBen.Text, "-") - 1))
    vlBancocta = Trim(Mid(cmbBancoCtaBen.Text, 1, InStr(1, cmbBancoCtaBen.Text, "-") - 1))
    vlCtanum = Trim(txtNumctaBen)
    'vlNumCCI = Trim(txt_CCIBen)
    
    
    'INICIO GCP-FRACTAL 11042019
    vlNumCCI = Trim(txt_CCIBen)
    vl_Fono1_Ben = Trim(Txt_Fono1_Ben)
    vl_Fono2_Ben = Trim(Txt_Fono2_Ben)
    vl_ConTratDatos_Ben = Trim(chkConTratDatos_Ben)
    vl_ConUsoDatosCom_Ben = Trim(chkConUsoDatosCom_Ben)
    'FIN GCP-FRACTAL 11042019
     
    'Valida que el rut no exista en la grilla
    If flValidaRutGrilla(vlTipoIden, vlNumIden) = False Then
        Exit Function
    End If
    
    Msf_GriAseg.Row = vlNumOrden
    
    Msf_GriAseg.Col = 5
    Msf_GriAseg.Text = vlFecInv
    
    Msf_GriAseg.Col = 6
    Msf_GriAseg.Text = vlCauInv
    
    Msf_GriAseg.Col = 11
    Msf_GriAseg.Text = vlTipoIden
            
    Msf_GriAseg.Col = 12
    Msf_GriAseg.Text = vlNumIden
            
    Msf_GriAseg.Col = 13
    Msf_GriAseg.Text = vlNombre
    
    Msf_GriAseg.Col = 14
    Msf_GriAseg.Text = vlNombreSeg
    
    Msf_GriAseg.Col = 15
    Msf_GriAseg.Text = vlApBen
    
    Msf_GriAseg.Col = 16
    Msf_GriAseg.Text = vlAmBen
              
    If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
        Msf_GriAseg.Col = 20
        Msf_GriAseg.Text = Trim(Txt_FecFallBen.Text)
    End If
              
    Msf_GriAseg.Col = 25
    Msf_GriAseg.Text = vlTipcta
    
    Msf_GriAseg.Col = 26
    Msf_GriAseg.Text = vlMoncta
    
    Msf_GriAseg.Col = 27
    Msf_GriAseg.Text = vlBancocta
    
    Msf_GriAseg.Col = 28
    Msf_GriAseg.Text = vlCtanum
    
    'Inicio GCP-Fractal 22032019
    Msf_GriAseg.Col = 31
    Msf_GriAseg.Text = vlNumCCI
    
    Msf_GriAseg.Col = 32
    Msf_GriAseg.Text = vl_Fono1_Ben
  
    Msf_GriAseg.Col = 33
    Msf_GriAseg.Text = vl_Fono2_Ben
    
    Msf_GriAseg.Col = 34
    Msf_GriAseg.Text = vl_ConTratDatos_Ben
    
    Msf_GriAseg.Col = 35
    Msf_GriAseg.Text = vl_ConUsoDatosCom_Ben
    
    Msf_GriAseg.Col = 36
    Msf_GriAseg.Text = vlCorreoBen
   
    'Fin GCP-Fractal 22032019
    
    Dim vl_IndexNacionalidad As Integer
        
    If vlNumOrden = 1 Then
    
        Msf_GriAseg.Col = 29
        Msf_GriAseg.Text = vl_nacionalidad
        vl_IndexNacionalidad = Trim(Mid(cboNacionalidad.Text, 1, InStr(1, cboNacionalidad.Text, "-") - 1))
        If vl_IndexNacionalidad > 0 Then
            Msf_GriAseg.Col = 30
            Msf_GriAseg.Text = fg_obtener_descripcion_nacionalidad(vl_nacionalidad)
        End If
        
    End If
    
    If vlNumOrden > 1 Then
  
        Msf_GriAseg.Col = 29
        Msf_GriAseg.Text = vl_nacionalidadben
        vl_IndexNacionalidad = Trim(Mid(cboNacionalidadBen.Text, 1, InStr(1, cboNacionalidadBen.Text, "-") - 1))
        If vl_IndexNacionalidad > 0 Then
            Msf_GriAseg.Col = 30
            Msf_GriAseg.Text = fg_obtener_descripcion_nacionalidad(vl_nacionalidadben)
        End If
        
    End If
              
              
Exit Function
Err_ModGri:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
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

Private Sub Btn_Porcentaje_Click()
Dim vlTipoPension As String
Dim vlMtoPensionRef As Double
Dim vlDerCrecerCot As String
Dim vlIndCobertura As String
Dim vlMesesGar As Long
On Error GoTo Err_Cmd_BMCalcular_Click
    
    'I--- ABV 15/04/2005 ---
    'Modifique la Fecha a utilizar para determinar los Porcentajes de
    'Pensión - Igualmente debo preguntar a Daniela si está bien o no
    '-----------------------------------------------------------
    'vlFecVig = Format(CDate(Trim(Txt_PMIniVig)), "yyyymmdd")
    'Validar el ingreso de la Fecha de Efecto del Endoso
    If (Txt_FecVig = "") Then
        MsgBox "Debe ingresar la Fecha de Vigencia sobre la cual se realizarán los cálculos.", vbCritical, "Operación Cancelada"
        Txt_FecVig.SetFocus
        Exit Sub
    End If
    If Not IsDate(Txt_FecVig) Then
        MsgBox "La Fecha de Vigencia ingresada no es una fecha válida.", vbCritical, "Operación Cancelada"
        Txt_FecVig.SetFocus
        Exit Sub
    End If
    
'    vgPalabraAux = Format(flCalculaFechaEfecto(Trim("")), "yyyymmdd")
'    If vgPalabraAux = "" Then Exit Sub
'    vgPalabra = Format(Txt_EndFecEfecto, "yyyymmdd")
'    If (vgPalabra < vgPalabraAux) Then
'        MsgBox "La Fecha de Efecto ingresada es menor a la Fecha de Cierre de la Póliza.", vbCritical, "Operación Cancelada"
'        Txt_EndFecEfecto.SetFocus
'        Exit Sub
'    End If
    
'Validar la Fecha de Devengue
    If (Lbl_FecDev = "") Then
        MsgBox "Debe ingresar la Fecha de Devengue sobre la cual se realizarán los cálculos.", vbCritical, "Operación Cancelada"
'        Lbl_FecDev.SetFocus
        Exit Sub
    End If
    If Not IsDate(Lbl_FecDev) Then
        MsgBox "La Fecha de Devengue ingresada no es una fecha válida.", vbCritical, "Operación Cancelada"
'        Lbl_FecDev.SetFocus
        Exit Sub
    End If
    
'Validar el Tipo de Pensión
    If (Lbl_TipPen = "") Then
        Exit Sub
    End If
    vlTipoPension = fgObtenerCodigo_TextoCompuesto(Lbl_TipPen)
    
'Validar la Pensión de Referencia
    If Not IsNumeric(Lbl_MtoPension) Then
        Exit Sub
    End If
    vlMtoPensionRef = CDbl(Lbl_MtoPension)
    
'Validar el Derecho a Crecer
    If (Lbl_DerCre = "") Then
        Exit Sub
    End If
    vlDerCrecerCot = Mid(Lbl_DerCre, 1, 1)

'Validar la Cobertura
    If (Lbl_IndCob = "") Then
        Exit Sub
    End If
    vlIndCobertura = Mid(Lbl_IndCob, 1, 1)

'Validar Meses Garantizados
    If (Lbl_MesesGar = "") Then
        Exit Sub
    End If
    vlMesesGar = CLng(Lbl_MesesGar)

    vlFecVig = Format(Txt_FecVig, "yyyymmdd")
    vlFecDev = Format(Lbl_FecDev, "yyyymmdd")
    
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
    
    Call fgCalcularPorcentajeBenef(vlFecDev, vlNumCargas, stBeneficiariosMod, vlTipoPension, vlMtoPensionRef, True, vlDerCrecerCot, vlIndCobertura, True, vlMesesGar)
    
    If (vgError = 0) Then
'        Call flInicializaGrillaBenef(Msf_GriAseg)
        Call fgActualizaGrillaBeneficiarios(Msf_GriAseg, stBeneficiariosMod, vlNumCargas, vlMesesGar, vlDerCrecerCot, vlFecDev, vlTipoPension)
        
'        Lbl_PMNumCar = (Msf_BMGrilla.Rows - 1)
        
'        vlSwCalIntOK = True
        
        MsgBox "El cálculo ha finalizado Correctamente.", vbInformation, "Operación Realizada"
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





Private Sub cboNacionalidad_Click()
vl_nacionalidad = cboNacionalidad.ItemData(cboNacionalidad.ListIndex)
End Sub

Private Sub cboNacionalidadBen_Click()
vl_nacionalidadben = cboNacionalidadBen.ItemData(cboNacionalidadBen.ListIndex)
End Sub

Private Sub chkConTratDatos_Afil_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    chkConUsoDatosCom_Afil.SetFocus
End If
End Sub




Private Sub chkConTratDatos_Ben_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    chkConUsoDatosCom_Ben.SetFocus
End If

End Sub

Private Sub chkConUsoDatosCom_Afil_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_ViaPago.SetFocus
End If

End Sub

Private Sub chkConUsoDatosCom_Ben_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmd_SalirDir.SetFocus
End If

End Sub

'Private Sub cboNacionalidad_Click()
'Me.Caption = cboNacionalidad.ItemData(cboNacionalidad.ListIndex)
'End Sub

Private Sub Cmb_EstCivil_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Cmb_Salud.SetFocus
End If
End Sub

Private Sub Cmb_TipIdRep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_NumIdRep.SetFocus
    End If
End Sub

Private Sub Cmb_TipIdRep_LostFocus()
    Call pConcatenaRepresentante

End Sub

Private Sub Cmb_TipoIdent_Click()
If (Cmb_TipoIdent <> "") Then
    vlPosicionTipoIden = Cmb_TipoIdent.ListIndex
    vlLargoTipoIden = Cmb_TipoIdent.ItemData(vlPosicionTipoIden)
    If (vlLargoTipoIden = 0) Then
        Txt_NumIdent.Text = "0"
        Txt_NumIdent.Enabled = False
    Else
        If (Cmb_TipoIdent.Enabled = True) Then
'            Txt_NumIdent = ""
        End If
        Txt_NumIdent.MaxLength = vlLargoTipoIden
        Txt_NumIdent.Enabled = True
        If (Txt_NumIdent <> "") Then Txt_NumIdent.Text = Mid(Txt_NumIdent, 1, vlLargoTipoIden)
    End If
End If
End Sub

Private Sub Cmb_TipoIdent_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_NumIdent.SetFocus
End If
End Sub

Private Sub Cmb_TipoIdent_LostFocus()
    Call flDatosCompletos
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

Private Sub Cmb_TipoIdentBen_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_NumIdentBen.SetFocus
End If
End Sub

Private Sub pConcatenaDireccion()

'Implementacion GestorCliente
    Lbl_Dir.Text = Trim(Trim(Right(Cmb_TipoVia.Text, 6)) & " " & Trim(Txt_NombreVia.Text) & " " & Trim(Txt_Numero.Text) _
        & " " & Trim(Txt_Interior.Text) & " " & Trim(Right(Cmb_TipoZona.Text, 6)) & " " & Trim(Txt_NombreZona.Text) _
        & " " & Trim(Lbl_Distrito.Caption) & " " & Trim(Lbl_Provincia.Caption) & " " & Trim(Lbl_Departamento.Caption) _
        & " " & Trim(Txt_Referencia.Text))
'Fin Implementacion GestorCliente
End Sub

Private Sub pConcatenaRepresentante()
    Lbl_Representante.Caption = Trim(Txt_NomRep.Text) & " " & Trim(Txt_ApPatRep.Text) & " " & Trim(Txt_ApMatRep.Text)

End Sub

Private Sub Cmb_TipoVia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_NombreVia.SetFocus
    End If

End Sub

Private Sub Cmb_TipoVia_LostFocus()
    Call pConcatenaDireccion
End Sub

Private Sub Cmb_TipoZona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_NombreZona.SetFocus
    End If

End Sub

Private Sub Cmb_TipoZona_LostFocus()
    Call pConcatenaDireccion
End Sub

Private Sub Cmb_Vejez_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    chkConTratDatos_Afil.SetFocus
End If
End Sub

Private Sub cmd_bancoCta_Click()


 'INICIO GCP-FRACTAL [REQUERIMIENTO SD-5041] 21032019
    'Validamos si se ha ingresado nmro de cuenta y cci en el caso que sea deposito a cuenta
If Not Left(Trim(cmbTipoCtaBen.Text), 2) = "00" Then
     If Left(Trim(cmbBancoCtaBen.Text), 2) = "00" Then
        MsgBox "No se encuentra definido el banco. ", vbCritical, "Error de Datos"
            Screen.MousePointer = 0
            cmbBancoCtaBen.SetFocus
            Exit Sub
     End If
 
    If txtNumctaBen = "" Then
       MsgBox "No se encuentra definido el Nro Cuenta y/o Nro CCI ", vbCritical, "Error de Datos"
       Screen.MousePointer = 0
       cmd_bancoCta.SetFocus
       Exit Sub
    End If
    
    
    
End If
  'FIN GCP-FRACTAL [REQUERIMIENTO SD-5041] 21032019




    framBancoCta.Visible = False
    Btn_Agregar.SetFocus
End Sub

Private Sub Cmd_BuscaCor_Click()
On Error GoTo Err_Buscar

    Frm_BuscaCorredor.flInicio ("Frm_CalPoliza")
    
Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarCauInvBen_Click()
On Error GoTo Err_Buscar

    Screen.MousePointer = 11
    Frm_CalPoliza.Enabled = False
    vgFormulario = "P"  'indica al formulario frm_buscacoti que fue llamado por el boton de Cotizaciones
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

Private Sub Cmd_BuscarDir_Click()
On Error GoTo Err_Buscar

    Frm_BusDireccion.flInicio ("Frm_CalPoliza")
    
Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_CauInv_Click()
On Error GoTo Err_Buscar

    Screen.MousePointer = 11
    Frm_CalPoliza.Enabled = False
    vgFormulario = "P"  'indica al formulario frm_buscacoti que fue llamado por el boton de Cotizaciones
    vgFormularioCarpeta = "A" 'indica carpeta de Afiliado
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

Private Sub Cmd_Direccion_Click()
    'RVF 20090914
    'Integracion GobiernoDeDatos(Se cambio a donde dirigia el boton de direccion por otro formulario)_
    'Fra_Direccion.Visible = True_
    Screen.MousePointer = 11
    If pNumTelefono = "" Then
    pNumTelefono = Txt_Fono.Text
    End If
    If pNumTelefono2 = "" Then
    pNumTelefono2 = Txt_Fono2_Afil.Text
    End If
    Call Frm_EditCamposPoliza.flIniciarValores(vlCodDireccion, pTipoTelefono, pNumTelefono, pCodigoTelefono, pTipoTelefono2, pNumTelefono2, pCodigoTelefono2, pTipoVia, pDireccion, pNumero, pTipoPref, pInterior, pManzana, pLote, pEtapa, pTipoConj, pConjHabit, pTipoBlock, pNumBlock, pReferencia, Lbl_Dir)
    Frm_CalPoliza.Enabled = False
    Frm_EditCamposPoliza.Show
    Screen.MousePointer = 0
    'Fin Integracion GobiernoDeDatos_
    
End Sub

Private Sub Cmd_Editar_Click()
On Error GoTo Err_Editar

    'Determinar el Tipo de Tabla desde la cual se obtiene la información
    vgBotonEscogido = vlBotonEscogido
    
    Screen.MousePointer = 11
    Frm_CalPolizaRec.Show
    Screen.MousePointer = 0

'    'Llama a las funciones para editar pólizas
'    Call flEditarAfiliado
'    Call flEditarCalculo
'    Call flEditarBeneficiarios
    
Exit Sub
Err_Editar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Poliza_Click()

On Error GoTo Err_PolizaClic
    
    flLimpiarDatosAfi
    flLimpiarDatosCal
    '''flLimpiarDatosBono
    flLimpiarDatosAseg

    Call flIniGrillaBen
    '''Call flInicializaGrillaBono
    
    Screen.MousePointer = 11
    Frm_CalPoliza.Enabled = False
    Frm_BuscaPol.Show
    Screen.MousePointer = 0
    
    'variable que se utiliza en la carga de datos desde la msf_griafiliado a las carpetas
    vlBotonEscogido = "P"   'determina que se hizo clic en el boton Póliza
    vgBotonEscogido = "P"
    
    Msf_GriAseg.rows = 1
    
    SSTab_Poliza.Enabled = False
    SSTab_Poliza.Tab = 0
    Fra_Cabeza.Enabled = False
    Cmd_CrearPol.Enabled = False
    Cmd_Grabar.Enabled = False
    cmdEnviaCorreo.Enabled = False
    Cmd_Editar.Enabled = False
    
Exit Sub
Err_PolizaClic:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cotizacion_Click()
On Error GoTo Err_Coti

    flLimpiarDatosAfi
    flLimpiarDatosAseg
    flLimpiarDatosCal

    Screen.MousePointer = 11
    Frm_CalPoliza.Enabled = False
    vgFormulario = "P"  'indica al formulario frm_buscacoti que fue llamado por el boton de Cotizaciones
    Frm_BuscaCoti.Show
    Screen.MousePointer = 0
    
    'determina que se hizo clic en el boton cotización
    'variable que se utiliza en la carga de datos desde la msf_GriAfiliado a las carpetas
    vlBotonEscogido = "C"
    vgBotonEscogido = "C"
    
    Msf_GriAseg.rows = 1
    
    SSTab_Poliza.Enabled = False
    SSTab_Poliza.Tab = 0
    Cmd_CrearPol.Enabled = False
    Cmd_Grabar.Enabled = False
    cmdEnviaCorreo.Enabled = False
    Cmd_Editar.Enabled = False

Exit Sub
Err_Coti:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_CrearPol_Click()
Dim vlExiste As Boolean
Dim rs As ADODB.Recordset
On Error GoTo Err_CreaPol

    'Integracion GobiernoDeDatos(Se invoca al servicio de Validacion de gestor de clientes)_
    Dim RpstServicio As Boolean
    For vlJ = 1 To (Msf_GriAseg.rows - 1)
        Msf_GriAseg.Row = vlJ
        Dim Mensaje As String
        RpstServicio = EnviarGestorCliente("Validar", Mensaje)
        If (RpstServicio = False) Then
           MsgBox Mensaje & " Beneficiario #" & " " & vlJ
           Exit Sub
        End If
    Next vlJ
    'Fin Integracion GobiernoDeDatos_
    

    'VALIDA SI HAY ALGUN HIJO MENOR O HIJO INVALIDO NO TIENE TUTOR
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "PP_MENORES_INVALIDOS_SIN_TUTOR.LISTAR('" & Lbl_NumCot.Caption & "')", vgConexionBD, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        If MsgBox("Hay hijos Menores de edad o hijos invalidos que no tienen asignado un tutor. Desea continuar con la creación de la Pre-Póliza", vbQuestion + vbYesNo, "Pre-Poliza") = vbYes Then
            
        Else
            Exit Sub
        End If
    End If


    Screen.MousePointer = 11
    
    'Validar si se trata de un Recálculo de Datos
    If (vgBotonEscogido = "R") Then
        vlCodDireccion = vgCodDireccion
    End If
    
    'validacion de datos
    If (Trim(Txt_FecVig) = "") Then
        MsgBox "Debe ingresar la Fecha de Vigencia de Póliza", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        Txt_FecVig.SetFocus
        Exit Sub
    End If
    
    'Validación de Fecha de Vigencia
    If fgValidaFecha(Trim(Txt_FecVig)) = False Then
        Txt_FecVig.Text = Format(CDate(Trim(Txt_FecVig)), "yyyymmdd")
        Txt_FecVig.Text = DateSerial(Mid((Txt_FecVig.Text), 1, 4), Mid((Txt_FecVig.Text), 5, 2), Mid((Txt_FecVig.Text), 7, 2))
    Else
        Txt_FecVig.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'Valida que esten todos los datos del afiliado
    If flValDatAfi = False Then
        Screen.MousePointer = 0
        SSTab_Poliza.Tab = 0
        Exit Sub
    End If
    
    If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
        If MsgBox("Seguro de haber actualizado fecha de fallecimiento", vbYesNo, "Confirmar") = vbNo Then
            Screen.MousePointer = 0
            SSTab_Poliza.Tab = 2
            Exit Sub
        End If
    End If
    
    If Not ValidaBancoPrincipal Then
        MsgBox "Debe indicar el CCI para los bancos no principales", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        SSTab_Poliza.Tab = 0
        
        Exit Sub
    
    End If
    
     If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
        If DirRep.vCodDireccion = "" Then
            MsgBox "Debe ingresar la direccion del representante", vbCritical, "Datos Incompletos"
           Screen.MousePointer = 0
           SSTab_Poliza.Tab = 0
           
           Exit Sub
        
        End If
        
        If cmbSexoRep.Text = "" Then
              MsgBox "Debe ingresar el sexo del representante", vbCritical, "Datos Incompletos"
           Screen.MousePointer = 0
           SSTab_Poliza.Tab = 0
           
           Exit Sub
        
        End If
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
    
    vlCodTipPen = Trim(Mid(Lbl_TipPen, 1, (InStr(1, Lbl_TipPen, "-") - 1)))

    If Not fgConexionBaseDatos(vgConectarBD) Then
        MsgBox "Fallo la Conexion con la Base de Datos", vbCritical, "Error de Conexión"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
'I--- ABV 22/06/2006 ---
    vlNumCot = Trim(Lbl_NumCot)
    
    'Determinar el Tipo de Cotización (SO, OE, PR)
    vlSql = "SELECT c.num_cot,c.cod_tipcot "
    vlSql = vlSql & " FROM pt_tmae_cotizacion c WHERE num_cot = '" & vlNumCot & "' "
    Set vgRs1 = vgConexionBD.Execute(vlSql)
    If Not vgRs1.EOF Then
        vlCodTipCot = (vgRs1!cod_tipcot)
    End If
    vgRs1.Close
    
'I--- ABV 22/06/2006 ---
'    If vlCodTipCot = clCodTipCotOfe Then
'        vlTablaDetCotizacion = "pt_tmae_detcotizacion"
'    End If
'    If vlCodTipCot = clCodTipCotExt Then
'        vlTablaDetCotizacion = "pt_tmae_detcotexterna"
'    End If
'    If vlCodTipCot = clCodTipCotRmt Then
'        vlTablaDetCotizacion = "pt_tmae_detcotremate"
'    End If
    vlTablaDetCotizacion = "pt_tmae_detcotizacion"
'F--- ABV 22/06/2006 ---

    'comenzar la transaccion
    vgConectarBD.BeginTrans
    
    vlExiste = True
    'genera nuevo numero de póliza
    Dim vlNewNumPol As Long
    vlSql = ""
    vgSql = "SELECT num_poliza FROM pd_tmae_gennumpol WHERE num_poliza = "
    vgSql = vgSql & " (SELECT MAX(num_poliza) FROM pd_tmae_gennumpol)"
    Set vgRs = vgConectarBD.Execute(vgSql)
    If Not vgRs.EOF Then
        vlNewNumPol = CLng((vgRs!Num_Poliza)) + 1
    Else
        vlExiste = False
        vlNewNumPol = 1
    End If
    vgRs.Close
    
    If (vlExiste = True) Then
        vlSql = "UPDATE pd_tmae_gennumpol SET "
        vlSql = vlSql & "num_poliza = '" & str(vlNewNumPol) & "'"
        vgConectarBD.Execute (vlSql)
    Else
        vlSql = "INSERT INTO pd_tmae_gennumpol "
        vlSql = vlSql & "(num_poliza) VALUES ("
        vlSql = vlSql & "'" & str(vlNewNumPol) & "'"
        vlSql = vlSql & ")"
        vgConectarBD.Execute (vlSql)
    End If
    
    Txt_NumPol = Format(vlNewNumPol, "0000000000")
    
    'GCRUZ -- ACTUALIZANDO num_idensup Y num_idenjef en pt_tmae_cotizacion
    'cuando los tiene null
    
     vlSql = ""
     vlSql = vlSql & "Update pt_tmae_cotizacion " & Chr(13)
     vlSql = vlSql & "set num_idensup= (select num_idenjefe " & Chr(13)
     vlSql = vlSql & "                  from pt_tmae_corredor cor " & Chr(13)
     vlSql = vlSql & "                  join pt_tmae_cotizacion cot on cor.num_idencor=cot.num_idencor " & Chr(13)
     vlSql = vlSql & "                   where cot.num_cot='" & vlNumCot & "'), " & Chr(13)
     vlSql = vlSql & "num_idenjef=(select num_idenjefes " & Chr(13)
     vlSql = vlSql & "              from pt_tmae_corredor cor " & Chr(13)
     vlSql = vlSql & "              join pt_tmae_cotizacion cot on cor.num_idencor=cot.num_idencor " & Chr(13)
     vlSql = vlSql & "               where cot.num_cot='" & vlNumCot & "') " & Chr(13)
     vlSql = vlSql & "where num_cot='" & vlNumCot & "' " & Chr(13)
     'Ini GCP 14/01/2021  Se comenta para que toda cotizacion sea actualizada
     ' vlSql = vlSql & " and num_idensup is null "
     'Fin Gcp 14/01/2022
     
     vgConectarBD.Execute (vlSql)
  
    'Crear la Pre-Póliza
    Call flCreaPoliza
 
    'marco---23/03/2010
    vlSql = ""
    vlSql = "UPDATE PD_TMAE_ORITUTOR SET NUM_POLIZA='" & Txt_NumPol & "' WHERE NUM_COTIZACION='" & vlNumCot & "'"
    vgConectarBD.Execute (vlSql)
    
    
    'RRR---26/08/2014
    
    'Dim vlExist As Integer
    'vlSql = "select count(*) as existe from pc_tmae_anticipo where cod_cuspp='" & Trim(Lbl_Cuspp) & "' and num_poliza ='XXXXXXXXXX' and num_cot='XXXXXXXXXX' and cod_estemision=3"
    'Set vgRs = vgConectarBD.Execute(vlSql)
    'If Not vgRs.EOF Then
    '    vlExist = vgRs!existe
    'End If
    'vgRs.Close
    'If vlExist > 0 Then
    '    vlSql = ""
    '    vlSql = "UPDATE pc_tmae_anticipo SET NUM_POLIZA='" & Txt_NumPol & "', NUM_COT='" & vlNumCot & "' WHERE cod_cuspp='" & Lbl_Cuspp & "' and cod_estemision=3"
    '    vgConectarBD.Execute (vlSql)
    'End If
   
    
    
    'Integracion GobiernoDeDatos(Se envia y agrega la data al gestor de clientes)_
   Dim Error As Integer
    For vlJ = 1 To (Msf_GriAseg.rows - 1)
    Msf_GriAseg.Row = vlJ
    RpstServicio = EnviarGestorCliente("Agregar", Mensaje)
    If (RpstServicio = False) Then
       Error = 1
       MsgBox Mensaje & " Beneficiario #" & " " & vlJ
    End If
    Next vlJ
    
    If vlSw = True And Error <> 1 Then
    vlBotonEscogido = "P"
    'Fin Integracion GobiernoDeDatos_
       'ejecutar la transaccion
        'vgConectarBD.CommitTrans
        vgConectarBD.RollbackTrans
    
       'cerrar la transaccion
        vgConectarBD.Close
    
        MsgBox "La Póliza ha sido correctamente Creada", vbInformation, "Proceso Completado"
'        Call Cmd_Cancelar_Click
        Cmd_CrearPol.Enabled = False
        Cmd_Grabar.Enabled = True
        cmdEnviaCorreo.Enabled = True
        
        If (vgNivelIndicadorBoton = "S") Then
            Cmd_Editar.Enabled = True
        Else
            Cmd_Editar.Enabled = False
        End If
        Cmd_Eliminar.Enabled = True
    Else
        'Deshacer la Transacción
        vgConectarBD.RollbackTrans
        
        'cerrar la transaccion
        vgConectarBD.Close

        Txt_NumPol = ""
        
    End If
    Screen.MousePointer = 0
    Fra_Representante.Visible = False
'    Lbl_Representante.Caption = ""
'    Txt_NumIdRep.Text = ""
'    Txt_NomRep.Text = ""
'    Txt_ApPatRep.Text = ""
'    Txt_ApMatRep.Text = ""
'    Me.txtTelRep1.Text = ""
'    Me.txtTelRep2.Text = ""
'    Me.txtCorreoRep.Text = ""
'
'    Call LimpiarDireccionRepresentante
    
    
    
    
Exit Sub
Err_CreaPol:
    'deshacer la transaccion
    vgConectarBD.RollbackTrans
    
    'cerrar la transaccion
    vgConectarBD.Close
    
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

'------------------------------
'Inserta los registros de beneficiarios a la grilla de beneficiarios
'------------------------------
Private Sub Btn_Agregar_Click()
On Error GoTo Err_Agregar
Screen.MousePointer = 11
If flValidaDatosAseg = False Then

'    Btn_Porcentaje.Visible = True

vl_nacionalidadben = cboNacionalidadBen.ItemData(cboNacionalidadBen.ListIndex)
If cboNacionalidadBen.ListIndex = 0 Then
    vl_nacionalidadben = ""
End If

    vlSw = True
    If Lbl_NumOrden <> "" Then
        vgRes = MsgBox("¿ Está seguro que desea Modificar los Datos del Beneficiario?", 4 + 32 + 256, "Operación de Actualización")
        If vgRes <> 6 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        Call flModificaGrilla
    End If
    If vlSw = False Then
        Screen.MousePointer = 0
        vlSw = True
        Exit Sub
    End If

    Call flLimpiarDatosAseg
    Fra_DatosBenef.Enabled = False
'    Txt_RutBen.Enabled = True
'    Txt_DgvBen.Enabled = True
'    Txt_RutBen.SetFocus
    
End If
Screen.MousePointer = 0
Exit Sub
Err_Agregar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmb_Bco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_NumCta.SetFocus
End If
End Sub

Private Sub Cmb_Salud_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_Dir.SetFocus
End If
End Sub

Private Sub Cmb_suc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SSTab_Poliza.Tab = 1
    Txt_FecIniPago.SetFocus
End If
End Sub

Private Sub Cmb_TipCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Cmb_Bco.SetFocus
End If
End Sub

Private Sub Cmb_ViaPago_Click()
If (vlSwViaPago = False) Then
    'vlSw = True
    Call flValidaViaPago
End If
'Call flValidaViaPago
End Sub

Private Sub Cmb_ViaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cmb_Suc.Enabled = True Then
            Cmb_Suc.SetFocus
        Else
            If Cmb_TipCta.Enabled = False Then
                Txt_FecIniPago.SetFocus
                SSTab_Poliza.Tab = 1
            Else
                 Cmb_TipCta.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Cmd_Cancelar_Click()
On Error GoTo Err_Cancelar
    
    vlCodDireccion = ""
    vgCodDireccion = ""
    vgBotonEscogido = ""

    flLimpiarDatosAfi
    flLimpiarDatosCal
    flLimpiarDatosBono
    flLimpiarDatosAseg
    Lbl_CUSPP = ""
    'Lbl_Cuspp.Enabled = True
'    Txt_NumIdent = ""  'ABV 07/08/2007
    'Txt_Digito.Enabled = True
    Txt_FecVig = ""
    Txt_NumPol = ""
    Lbl_NumCot = ""
    Lbl_SolOfe = ""
    Lbl_IdeOfe = ""
    Lbl_SecOfe = ""
    Fra_Cabeza.Enabled = False
    Fra_Afiliado.Enabled = False
    Fra_DatCal.Enabled = False
    Fra_SumaBono.Enabled = False
'''    Fra_DatBon.Enabled = False -  MC 29/05/2007
    Fra_DatosBenef.Enabled = False
    SSTab_Poliza.Enabled = False
    SSTab_Poliza.Tab = 0
    Call flIniGrillaBen
    '''Call flInicializaGrillaBono - MC 29/05/2007
'    Msf_GriAfiliado.Enabled = False
    Cmd_Poliza.Enabled = True
    Cmd_Cotizacion.Enabled = True
    Cmd_CrearPol.Enabled = False
    Cmd_Grabar.Enabled = False
    cmdEnviaCorreo.Enabled = False
    Cmd_Editar.Enabled = False
    Cmd_Eliminar.Enabled = False
    vlSw = False
    Fra_Representante.Visible = False
    Lbl_Representante.Caption = ""
    Txt_NumIdRep.Text = ""
    Txt_NomRep.Text = ""
    Txt_ApPatRep.Text = ""
    Txt_ApMatRep.Text = ""
    Me.txtTelRep1.Text = ""
    Me.txtTelRep2.Text = ""
    Me.txtCorreoRep.Text = ""
    
    Call LimpiarDireccionRepresentante
    
    
Exit Sub
Err_Cancelar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Eliminar_Click()
On Error GoTo err_eli
Dim vlPoliza As String
Dim vlSwEliminar As Boolean

    'valida que exista el número de póliza
    If (Trim(Txt_NumPol) = "") Then
        MsgBox "Debe Seleccionar una Póliza para Eliminar", vbExclamation, "Datos Incompletos"
        Exit Sub
    End If
    
    vlResp = MsgBox(" ¿ Está seguro que desea Eliminar TODOS LOS DATOS de la Póliza Seleccionada ?", 4 + 32 + 256, "Proceso de Eliminación de Datos")
    If vlResp <> 6 Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    vlNumPol = Trim(Txt_NumPol)
    
    vlSwEliminar = False
    
    vlSql = "SELECT num_poliza,num_cot,num_correlativo "
    vlSql = vlSql & "FROM pd_tmae_oripoliza WHERE "
    vlSql = vlSql & "num_poliza= '" & vlNumPol & "'"
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
        vlNumCot = (vgRs!Num_Cot)
        vlNumCorrelativo = vgRs!Num_Correlativo
        vlSwEliminar = True
    End If
    vgRs.Close

    If (vlSwEliminar = True) Then
        Call flEliminarPoliza(vlNumPol)
    Else
        MsgBox "La Póliza Ingresada No se Encuentra Registrada en la Base de Datos", vbExclamation, "Datos Incompletos"
    End If
    
    Screen.MousePointer = 0

Exit Sub
err_eli:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Grabar_Click()
On Error GoTo Err_Grabar


    'Integracion GobiernoDeDatos(Se valida la informacion antes de enviarla al gestor)_
    If Not vlBotonEscogido = "C" Then
        ObtenerDatosIniciales (Trim(Txt_NumPol))
    End If
    
     Dim RpstServicio As Boolean
    For vlJ = 1 To (Msf_GriAseg.rows - 1)
    Msf_GriAseg.Row = vlJ
    Dim Mensaje As String

    
    RpstServicio = EnviarGestorCliente("Validar", Mensaje)
    If (RpstServicio = False) Then
       MsgBox Mensaje & " Beneficiario #" & " " & vlJ
       Exit Sub
    End If
    Next vlJ

    'Fin Integracion GobiernoDeDatos_
   
    'validacion de datos
    If (Trim(Txt_FecVig) = "") Then
        MsgBox "Debe ingresar la Fecha de Vigencia de Póliza", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        Txt_FecVig.SetFocus
        Exit Sub
    End If
    
    'Validación de Fecha de Vigencia
    If fgValidaFecha(Trim(Txt_FecVig)) = False Then
        Txt_FecVig.Text = Format(CDate(Trim(Txt_FecVig)), "yyyymmdd")
        Txt_FecVig.Text = DateSerial(Mid((Txt_FecVig.Text), 1, 4), Mid((Txt_FecVig.Text), 5, 2), Mid((Txt_FecVig.Text), 7, 2))
    Else
        Txt_FecVig.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'Valida que esten todos los datos del afiliado
    If flValDatAfi = False Then
        Screen.MousePointer = 0
        SSTab_Poliza.Tab = 0
        Exit Sub
    End If
    
    If cboNacionalidad.ListIndex = 0 Then
        vl_nacionalidad = ""
    End If
    
    vlTipoIden = fgObtenerCodigo_TextoCompuesto(Cmb_TipoIdent.Text)
    vlNumIden = Txt_NumIdent
    vlFecVig = Format(Txt_FecVig, "yyyymmdd")
    
    'Buscar el Procentaje del Beneficio Social
    If (fgBuscarPorcBenSocial(vlFecVig, vgPrcBenSocial) = False) Then
        MsgBox "No se encuentra definido el porcentaje para el Beneficio Social", vbCritical, "Error de Datos"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'Valida que esten todos los datos del calculo
    If flValDatCal = False Then
        Screen.MousePointer = 0
        SSTab_Poliza.Tab = 1
        Exit Sub
    End If
    
    'Valida que los beneficiarios de la grilla tengan todos los datos
    If flValidaBenGrilla = False Then
        Screen.MousePointer = 0
        SSTab_Poliza.Tab = 2
        Exit Sub
    End If
    
    
    'INICIO GCP-FRACTAL [REQUERIMIENTO SD-5041] 21032019
    'Validamos si se ha ingresado nmro de cuenta y cci en el caso que sea deposito a cuenta
    If Left(Trim(Cmb_ViaPago.Text), 2) = "02" Then
      If Left(Trim(Cmb_TipCta.Text), 2) = "00" Then
        MsgBox "No se encuentra definido el tipo de cuenta. ", vbCritical, "Error de Datos"
            Screen.MousePointer = 0
            Cmb_TipCta.SetFocus
            Exit Sub
     End If
     
     If Left(Trim(Cmb_Bco.Text), 2) = "00" Then
        MsgBox "No se encuentra definido el banco. ", vbCritical, "Error de Datos"
            Screen.MousePointer = 0
            Cmb_Bco.SetFocus
            Exit Sub
     End If
   
       Dim TP As String
        TP = Left(Trim(Lbl_TipPen.Caption), 2)
        If (TP = "04" Or TP = "05") Then
         If Txt_NumCta = "" Or txt_CCI = "" Then
            MsgBox "No se encuentra definido el Nro Cuenta y/o Nro CCI ", vbCritical, "Error de Datos"
            Screen.MousePointer = 0
            Exit Sub
         End If
      End If
   End If
   
     If Not ValidaBancoPrincipal Then
        MsgBox "Debe indicar el CCI para los bancos no principales", vbCritical, "Datos Incompletos"
        Screen.MousePointer = 0
        SSTab_Poliza.Tab = 0
        
        Exit Sub
    
    End If
    
    If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
        If DirRep.vCodDireccion = "" Then
            MsgBox "Debe ingresar la direccion del representante", vbCritical, "Datos Incompletos"
           Screen.MousePointer = 0
           SSTab_Poliza.Tab = 0
           
           Exit Sub
        
        End If
        
        If cmbSexoRep.Text = "" Then
              MsgBox "Debe ingresar el sexo del representante", vbCritical, "Datos Incompletos"
           Screen.MousePointer = 0
           SSTab_Poliza.Tab = 0
           
           Exit Sub
        
        End If
 
     End If
 
    'FIN GCP-FRACTAL [REQUERIMIENTO SD-5041] 21032019
      
    vlResp = MsgBox(" ¿ Está seguro que desea Modificar los datos de la Póliza ?", 4 + 32 + 256, "Proceso de Modificación de Datos")
    If vlResp <> 6 Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
   Call flModificaPoliza(Trim(Txt_NumPol))
  
    'Cambios GestorCliente_
     Dim Error As Integer
        For vlJ = 1 To (Msf_GriAseg.rows - 1)
        Msf_GriAseg.Row = vlJ
        RpstServicio = EnviarGestorCliente("Agregar", Mensaje)
        If (RpstServicio = False) Then
           Error = 1
           Screen.MousePointer = 0
           MsgBox Mensaje & " Beneficiario #" & " " & vlJ
        End If
        Next vlJ

    
    If vlSw = True And Error <> 1 Then
    'Fin Cambios GestorCliente_
        'deshabilito el objeto de las carpetas
        SSTab_Poliza.Tab = 0
        SSTab_Poliza.Enabled = True
        Cmd_Grabar.Enabled = True
        cmdEnviaCorreo.Enabled = True
        
        If (vgNivelIndicadorBoton = "S") Then
            Cmd_Editar.Enabled = True
        Else
            Cmd_Editar.Enabled = False
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

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir

    If Trim(Txt_NumPol = "") Then
        MsgBox "Debe escoger una Póliza para Imprimir", vbCritical, "Datos Incompletos"
        Exit Sub
    End If
   
    Screen.MousePointer = 11
      
    Call flImprimirPoliza(Trim(Txt_NumPol))

Exit Sub
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpia
    
    If SSTab_Poliza.Tab = 0 Then
        If SSTab_Poliza.TabEnabled(2) = True Then
'            flLimpiarDatosAfi
            If SSTab_Poliza.TabEnabled(2) = False Then
                If (Cmb_TipoIdent.ListCount <> 0) Then
                    Cmb_TipoIdent.ListIndex = 0
                End If
                'Txt_NumIdent = ""
            End If
            
            Txt_NomAfi = ""
            Txt_NomAfiSeg = ""
            Txt_ApPatAfi = ""
            Txt_ApMatAfi = ""
            Txt_FecInv = ""
            Txt_Dir = ""
            Txt_Fono = ""
            Txt_Correo = ""
            Txt_Nacionalidad = ""
            Txt_NumCta = ""
                
            If (Cmb_EstCivil.ListCount <> 0) Then
                Cmb_EstCivil.ListIndex = 0
            End If
            If (Cmb_Salud.ListCount <> 0) Then
                Cmb_Salud.ListIndex = 0
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
            'flLimpiarDatosAfi
            If SSTab_Poliza.TabEnabled(2) = False Then
                If (Cmb_TipoIdent.ListCount <> 0) Then
                    Cmb_TipoIdent.ListIndex = 0
                End If
                'Txt_NumIdent = ""
            End If
            
            Txt_NomAfi = ""
            Txt_NomAfiSeg = ""
            Txt_ApPatAfi = ""
            Txt_ApMatAfi = ""
            Txt_FecInv = ""
            Txt_Dir = ""
            Txt_Fono = ""
            Txt_Correo = ""
            Txt_Nacionalidad = ""
            Txt_NumCta = ""

            If (Cmb_EstCivil.ListCount <> 0) Then
                Cmb_EstCivil.ListIndex = 0
            End If
            If (Cmb_Salud.ListCount <> 0) Then
                Cmb_Salud.ListIndex = 0
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
'        Lbl_FecDev = ""
'        Lbl_FecIncorpora = ""
'        Txt_FecIniPago = ""
'        Lbl_Cuspp = ""
'        Txt_PrcFam = ""
    End If
    
'    If SSTab_Poliza.Tab = 2 Then
''        flLimpiarDatosBono
'    End If
    
    If SSTab_Poliza.Tab = 2 Then
'        flLimpiarDatosAseg
        If (Cmb_TipoIdentBen.ListCount <> 0) Then
            Cmb_TipoIdentBen.ListIndex = 0
        End If
        
'        Txt_NumIdentBen = ""
        Txt_NombresBen = ""
        Txt_NombresBenSeg = ""
        Txt_ApPatBen = ""
        Txt_ApMatBen = ""
        Txt_FecInvBen = ""
    End If

Call LimpiarDireccionRepresentante

Exit Sub
Err_Limpia:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Representante_Click()
    'RVF 20090914
    Fra_Representante.Visible = True

End Sub

Private Sub Cmd_Salir_Click()
On Error GoTo Err_Descargar

    Screen.MousePointer = 11
    Unload Me
    Screen.MousePointer = 0

Exit Sub
Err_Descargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_SalirDir_Click()
    'RVF 20090914
    Call pConcatenaDireccion
    Fra_Direccion.Visible = False
    
End Sub

Private Sub Cmd_SalirRep_Click()
    'RVF 20090914
    Call pConcatenaRepresentante
    Fra_Representante.Visible = False
    
End Sub
'Private Sub LlamadasAPI()
'    Call ObtenerToken
'    Call CreaTicket
'End Sub

Private Sub CmdDireccionRep_Click()
    Frm_DireccionRep.Show
End Sub

Private Sub Command1_Click()
    framBancoCta.Visible = True
End Sub

Private Sub cmdTutor_Click()

Msf_GriAseg.Col = 0
Frm_AntTutores.EstadoInvalidez = Mid(Trim(Lbl_SitInvBen.Caption), 1, 1)
Frm_AntTutores.TIPODOC = Cmb_TipoIdentBen
Frm_AntTutores.NumDoc = Txt_NumIdentBen
Frm_AntTutores.NombreCompleto = Txt_NombresBen.Text & " " & Txt_NombresBenSeg.Text & " " & Txt_ApPatBen.Text & " " & Txt_ApMatBen.Text
Frm_AntTutores.Cotizacion = Lbl_NumCot.Caption
Frm_AntTutores.BotonSel = vlBotonEscogido
Frm_AntTutores.Orden = Msf_GriAseg.Text
Frm_AntTutores.fechaNac = Lbl_FecNacBen.Caption
Frm_AntTutores.Show

End Sub






Private Sub Form_Load()
On Error GoTo Err_Ejec
    
    SSTab_Poliza.Tab = 0
    Frm_CalPoliza.Top = 0
    Frm_CalPoliza.Left = 0
    
    Call fgComboGeneral(vgCodTabla_InsSal, Cmb_Salud)
'    Call fgComboGeneral(vgCodTabla_ViaPago, Cmb_ViaPago)
    Call fgComboGeneral(vgCodTabla_TipCta, Cmb_TipCta)
    Call fgComboGeneral(vgCodTabla_Bco, Cmb_Bco)
    Call fgComboGeneral(vgCodTabla_TipVej, Cmb_Vejez)
    Call fgComboGeneral(vgCodTabla_EstCiv, Cmb_EstCivil)
    
    Call fgComboNacionalidad(cboNacionalidad)
    Call fgComboNacionalidad(cboNacionalidadBen)
    
    Call fgComboTipoIdentificacion(Cmb_TipoIdent)
    Call fgComboTipoIdentificacion(Cmb_TipoIdentBen)
    Call fgComboTipoIdentificacion(Cmb_TipIdRep)   'RVF 20090914
    
    Call fgComboGeneralDirec("DVI", Cmb_TipoVia)   'RVF 20090914
    Call fgComboGeneralDirec("DZO", Cmb_TipoZona)  'RVF 20090914
    
    'RRR 26122013
    Call fgComboGeneral(vgCodTabla_TipCta, cmbTipoCtaBen)
    Call fgComboGeneral(vgCodTabla_Bco, cmbBancoCtaBen)
    
    Call fgComboGeneral("MP", cmb_MonCta)
    Call fgComboGeneral("MP", cmbMonctaBen)
    
    'Llenar los casilleros de Información
    flIniGrillaBen
    '''Call flInicializaGrillaBono

    Fra_Cabeza.Enabled = False
    Fra_Afiliado.Enabled = False
    Fra_PagPension.Enabled = False
    Fra_DatCal.Enabled = False
    Fra_SumaBono.Enabled = False
    '''Fra_DatBon.Enabled = False
    Fra_DatosBenef.Enabled = False
    Fra_Direccion.Visible = False
    Msf_GriAseg.Enabled = False
    SSTab_Poliza.Tab = 0
    SSTab_Poliza.Enabled = False
'    Msf_GriAfiliado.Enabled = False
    Cmd_Grabar.Enabled = False
    cmdEnviaCorreo.Enabled = False
    Cmd_Editar.Enabled = False
    Cmd_Eliminar.Enabled = False
    Cmd_CrearPol.Enabled = False

'    If vgNivelIndicadorBoton = "S" Then
'        Cmd_Editar.Enabled = True
'    Else
'        Cmd_Editar.Enabled = False
'    End If
       
    Call fgCargarTablaMoneda(vgCodTabla_TipMon, egTablaMoneda(), vgNumeroTotalTablasMoneda)

    lbltutor.Visible = False
    cmdTutor.Visible = False
    txtTutor.Visible = False
    
    
    'Intregracion GobiernoDeDatos(Se bloquean los campos)
    Lbl_Dir.Locked = True
    Txt_Fono.Locked = True
    Txt_Fono2_Afil.Locked = True
    'Fin Intregracion GobiernoDeDatos(Se bloquean los campos)
    
    Call LimpiarDireccionRepresentante
    
    
Exit Sub
Err_Ejec:
Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Frm_CalPoliza = Nothing
    
End Sub



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
    
    Call flLimpiarDatosAseg
     Fra_DatosBenef.Enabled = True
'    Txt_RutBen.Enabled = True
 '   Txt_DgvBen.Enabled = True
    
    Lbl_NumOrden = Msf_GriAseg.Text
    
    Msf_GriAseg.Col = 1
    If Msf_GriAseg.Text <> "" Then Lbl_Par = Trim(Msf_GriAseg.Text) & " - " & fgBuscarGlosaElemento(vgCodTabla_Par, Msf_GriAseg.Text)
    'Lbl_Par = Trim(Msf_GriAseg.Text)
    
    If Msf_GriAseg.Text = "99" Then
        MsgBox "Los Datos del Afiliado Deben ser Modificados en Carpeta Afiliado", vbExclamation, "Aviso"
        Fra_DatosBenef.Enabled = False
    Else
        Fra_DatosBenef.Enabled = True
    End If
    
    Msf_GriAseg.Col = 2
    If Msf_GriAseg.Text <> "" Then Lbl_Grupo = Trim(Msf_GriAseg.Text) & " - " & fgBuscarGlosaElemento(vgCodTabla_GruFam, Msf_GriAseg.Text)
    'Lbl_Grupo = Trim(Msf_GriAseg.Text)
    
    Msf_GriAseg.Col = 3
    If Msf_GriAseg.Text <> "" Then Lbl_SexoBen = Trim(Msf_GriAseg.Text) & " - " & fgBuscarGlosaElemento(vgCodTabla_Sexo, Msf_GriAseg.Text)
    'Lbl_SexoBen = Trim(Msf_GriAseg.Text)
    
    Msf_GriAseg.Col = 4
    If Msf_GriAseg.Text <> "" Then Lbl_SitInvBen = Trim(Msf_GriAseg.Text) & " - " & fgBuscarGlosaElemento(vgCodTabla_SitInv, Msf_GriAseg.Text)
    'Lbl_SitInvBen = Trim(Msf_GriAseg.Text)
    
    Msf_GriAseg.Col = 5
    If Msf_GriAseg.Text <> "" Then
        'Txt_FecInvBen = DateSerial(Mid(Msf_GriAseg.Text, 1, 4), Mid(Msf_GriAseg.Text, 5, 2), Mid(Msf_GriAseg.Text, 7, 2))
        Txt_FecInvBen = Msf_GriAseg.Text
    End If
    
    Msf_GriAseg.Col = 6
    If Msf_GriAseg.Text <> "" Then Lbl_CauInvBen = Trim(Msf_GriAseg.Text) & " - " & fgBuscarGlosaCauInv(Msf_GriAseg.Text)
    'Lbl_CauInvBen = Trim(Msf_GriAseg.Text)
    
    Msf_GriAseg.Col = 7
    If Msf_GriAseg.Text <> "" Then Lbl_DerPension = Trim(Msf_GriAseg.Text) & " - " & fgBuscarGlosaElemento(vgCodTabla_DerPen, Msf_GriAseg.Text)
    'Lbl_DerPension = Trim(Msf_GriAseg.Text)
    
    Msf_GriAseg.Col = 8
    'derecho a crecer
    
    Msf_GriAseg.Col = 9
    If Msf_GriAseg.Text <> "" Then
        'Lbl_FecNacBen = DateSerial(Mid(Msf_GriAseg.Text, 1, 4), Mid(Msf_GriAseg.Text, 5, 2), Mid(Msf_GriAseg.Text, 7, 2))
        Lbl_FecNacBen = Msf_GriAseg.Text
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
    'Número Identificación
    If Msf_GriAseg.Text <> "" Then
        Txt_NumIdentBen = Msf_GriAseg.Text
    End If
    
    Msf_GriAseg.Col = 13
    Txt_NombresBen = Msf_GriAseg.Text
    
    Msf_GriAseg.Col = 14
    'Segundo Nombre
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
        Lbl_FecFallBen = Msf_GriAseg.Text
        Txt_FecFallBen.Text = Msf_GriAseg.Text
    End If
    
    '-- Begin : Modify by : ricardo.huerta
    Msf_GriAseg.Col = 29
    If Msf_GriAseg.Text <> "" Then
        Call fgBuscaPos(cboNacionalidadBen, Msf_GriAseg.Text)
    End If
    '-- End    :  Modify by : ricardo.huerta

   
    If Cmb_TipoIdentBen.Enabled = True Then
        Cmb_TipoIdentBen.SetFocus
    Else
        If Txt_NombresBen.Enabled = True Then
            Txt_NombresBen.SetFocus
        End If
    End If
    
'    Txt_RutBen.Enabled = False
'    Txt_DgvBen.Enabled = False

 'busca el ultimo tutor registrado
    Msf_GriAseg.Col = 0
    vgSql = ""
    vgSql = "SELECT *  FROM Pd_TMAE_oriTUTOR "
    If vlBotonEscogido = "C" Then
        vgSql = vgSql & "WHERE num_cotizacion = '" & Trim(Lbl_NumCot.Caption) & "' AND "
    Else
        vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_NumPol.Text) & "' AND "
    End If
    vgSql = vgSql & "num_orden = " & Trim(Msf_GriAseg.Text) & " "
    Set vgRs = vgConexionBD.Execute(vgSql)
    
    If Not vgRs.EOF Then
        txtTutor.Text = vgRs!gls_nomtut & " " & vgRs!gls_nomsegtut & " " & vgRs!GLS_PATTUT & " " & vgRs!GLS_MATTUT
    Else
        txtTutor.Text = ""
    End If
    
    'RRR 26/12/2013
 Msf_GriAseg.Col = 25
'Tipo Cuenta
    If Msf_GriAseg.Text <> "" Then
        vgPalabra = Msf_GriAseg.Text
        vgI = fgBuscarPosicionCodigoCombo(vgPalabra, cmbTipoCtaBen)
        If (cmbTipoCtaBen.ListCount > 0) Then
            cmbTipoCtaBen.ListIndex = vgI
        End If
    Else
        cmbTipoCtaBen.ListIndex = 0
    End If

 Msf_GriAseg.Col = 26
'Tipo Moneda Cuenta
    If Msf_GriAseg.Text <> "" Then
        vgPalabra = Msf_GriAseg.Text
        vgI = fgBuscarPosicionCodigoCombo(vgPalabra, cmbMonctaBen)
        If (cmbMonctaBen.ListCount > 0) Then
            cmbMonctaBen.ListIndex = vgI
        End If
    Else
        cmbMonctaBen.ListIndex = 0
    End If

 Msf_GriAseg.Col = 27
'Tipo Banco
    If Msf_GriAseg.Text <> "" Then
        vgPalabra = Msf_GriAseg.Text
        vgI = fgBuscarPosicionCodigoCombo(vgPalabra, cmbBancoCtaBen)
        If (cmbBancoCtaBen.ListCount > 0) Then
            cmbBancoCtaBen.ListIndex = vgI
        End If
    Else
        cmbBancoCtaBen.ListIndex = 0
    End If
 'Numero Cta
 Msf_GriAseg.Col = 28
 txtNumctaBen = Msf_GriAseg.Text
 
'INICIO GCP-FRACTAL 12042019
Msf_GriAseg.Col = 31
txt_CCIBen = Msf_GriAseg.Text

Msf_GriAseg.Col = 32
Txt_Fono1_Ben = Msf_GriAseg.Text

Msf_GriAseg.Col = 33
Txt_Fono2_Ben = Msf_GriAseg.Text

Msf_GriAseg.Col = 34
'chkConTratDatos_Ben.Value = IIf(Msf_GriAseg.Text = "", 0, 1)
chkConTratDatos_Ben.Value = IIf(Msf_GriAseg.Text = "", 0, Msf_GriAseg.Text) 'JVB 20210413 Correcion pq ponia 1 cuando venia 0

Msf_GriAseg.Col = 35
'chkConUsoDatosCom_Ben = IIf(Msf_GriAseg.Text = "", 0, 1)
chkConUsoDatosCom_Ben = IIf(Msf_GriAseg.Text = "", 0, Msf_GriAseg.Text) 'JVB 20210413 Correcion pq ponia 1 cuando venia 0
'FIN GCP-FRACTAL 12042019
 
Msf_GriAseg.Col = 36
Me.txtCorreoBen = Msf_GriAseg.Text

    

'MARCO -----18/03/2010
Msf_GriAseg.Col = 1
 If Mid(Lbl_SitInvBen.Caption, 1, 1) <> "N" And Msf_GriAseg.Text = "30" Then
    lbltutor.Visible = True
    cmdTutor.Visible = True
    txtTutor.Visible = True
    Exit Sub
 Else
    lbltutor.Visible = False
    cmdTutor.Visible = False
    txtTutor.Visible = False
 End If
 
 If ValidaMenorEdad(CDate(Lbl_FecNacBen.Caption)) = True And Msf_GriAseg.Text = "30" Then
    lbltutor.Visible = True
    cmdTutor.Visible = True
    txtTutor.Visible = True
 Else
    lbltutor.Visible = False
    cmdTutor.Visible = False
    txtTutor.Visible = False
 End If
 
 '***** Movido mas arriba *******
' If txtTutor.Visible = True Then
'    'busca el ultimo tutor registrado
'    Msf_GriAseg.Col = 0
'    vgSql = ""
'    vgSql = "SELECT *  FROM Pd_TMAE_oriTUTOR "
'    If vlBotonEscogido = "C" Then
'        vgSql = vgSql & "WHERE num_cotizacion = '" & Trim(Lbl_NumCot.Caption) & "' AND "
'    Else
'        vgSql = vgSql & "WHERE num_poliza = '" & Trim(Txt_NumPol.Text) & "' AND "
'    End If
'    vgSql = vgSql & "num_orden = " & Trim(Msf_GriAseg.Text) & " "
'    Set vgRs = vgConexionBD.Execute(vgSql)
'
'    If Not vgRs.EOF Then
'        txtTutor.Text = vgRs!gls_nomtut & " " & vgRs!gls_nomsegtut & " " & vgRs!GLS_PATTUT & " " & vgRs!GLS_MATTUT
'    Else
'        txtTutor.Text = ""
'    End If
' End If
 '***** Movido mas arriba *******
 
 
 'FIN MARCO




Exit Sub
Err_Grilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Function BuscaTutor(TIPODOC As String, DOC As String) As String
Dim cadena As String
Dim rs As ADODB.Recordset

On Error GoTo mierror
Set rs = New ADODB.Recordset
cadena = ""


Exit Function
mierror:
    MsgBox "No pudo cargar Tutor", vbInformation
    

End Function





Private Sub SSTab_Poliza_Click(PreviousTab As Integer)
    If vlBTipoEnvio = True Then
        If Msf_GriAseg.Row <> 0 Then
            Msf_GriAseg.Row = 1
            Msf_GriAseg.Col = 25
            Msf_GriAseg.Text = Trim(Mid(Cmb_TipCta.Text, 1, InStr(1, Cmb_TipCta.Text, "-") - 1))
            
            Msf_GriAseg.Col = 26
            Msf_GriAseg.Text = "NS"
            
            Msf_GriAseg.Col = 27
            Msf_GriAseg.Text = Trim(Mid(Cmb_Bco.Text, 1, InStr(1, Cmb_Bco.Text, "-") - 1))
            
            Msf_GriAseg.Col = 28
            Msf_GriAseg.Text = Txt_NumCta.Text
        End If
    End If
End Sub


Private Sub Txt_ApMatAfi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_FecInv.SetFocus
End If
End Sub

Private Sub Txt_ApMatAfi_LostFocus()

    Txt_ApMatAfi = Trim(UCase(Txt_ApMatAfi))
'    If (Msf_GriAseg.Rows - 1) = 0 Then
        Call flDatosCompletos
'    Else
'        If Trim(Lbl_Cuspp) = "" Or Trim(Txt_Digito) = "" Or _
'            Trim(Txt_NomAfi) = "" Or Trim(Txt_ApPatAfi) = "" Or Trim(Txt_ApMatAfi) = "" Then
'            SSTab_Poliza.TabEnabled(3) = True
'            SSTab_Poliza.Tab = 0
'        Else
'            SSTab_Poliza.TabEnabled(3) = True
'        End If
'        For vlJ = 1 To (Msf_GriAseg.Rows - 1)
'            Msf_GriAseg.Col = 1
'            Msf_GriAseg.Row = vlJ
'            If Msf_GriAseg.Text = "99" Then
'                Msf_GriAseg.Col = 14
'                Msf_GriAseg.Row = vlJ
'                Msf_GriAseg.Text = Trim(Txt_ApMatAfi)
'                vlJ = Msf_GriAseg.Rows
'            End If
'        Next
'    End If
    
End Sub

Private Sub Txt_ApMatBen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_FecInvBen.SetFocus
    End If
End Sub

Private Sub Txt_ApMatBen_LostFocus()
    Txt_ApMatBen = Trim(UCase(Txt_ApMatBen))
End Sub

Private Sub Txt_ApMatRep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmd_SalirRep.SetFocus
    End If

End Sub

Private Sub Txt_ApMatRep_LostFocus()
    Txt_ApMatRep.Text = Trim(UCase(Txt_ApMatRep.Text))
    Call pConcatenaRepresentante
End Sub

Private Sub Txt_ApPatAfi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_ApMatAfi.SetFocus
    End If
End Sub

Private Sub Txt_ApPatAfi_LostFocus()
    Txt_ApPatAfi = Trim(UCase(Txt_ApPatAfi))
    Call flDatosCompletos
End Sub

Private Sub Txt_ApPatBen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_ApMatBen.SetFocus
End If
End Sub

Private Sub Txt_ApPatBen_LostFocus()
    Txt_ApPatBen = Trim(UCase(Txt_ApPatBen))
End Sub

Private Sub Txt_ApPatRep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_ApMatRep.SetFocus
    End If

End Sub

Private Sub Txt_ApPatRep_LostFocus()
    Txt_ApPatRep.Text = Trim(UCase(Txt_ApPatRep.Text))
    Call pConcatenaRepresentante
End Sub





Private Sub txt_CCIBen_Change()
Dim Keychar As String
If KeyAscii > 31 Then
    Keychar = Chr(KeyAscii)
    If Not IsNumeric(Keychar) Then
        KeyAscii = 0
    End If
Else
    If KeyAscii = 13 Then
      cmd_bancoCta.SetFocus
    End If
End If
End Sub

Private Sub Txt_correo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_Correo = Trim(Txt_Correo)
        Cmb_Vejez.SetFocus
    End If
End Sub

Private Sub Txt_Dir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmd_BuscarDir.SetFocus
    End If
End Sub

Private Sub Txt_Dir_Lostfocus()
    Txt_Dir = Trim(UCase(Txt_Dir))
End Sub

Private Sub Txt_Correo_LostFocus()
 Dim bVerifica As Boolean
 
 bVerifica = Comprobar_Mail(Me.Txt_Correo.Text)
 If Not bVerifica Then
    MsgBox "El correo ingresado es inválido."
    Txt_Correo.SetFocus

 End If
    
      
End Sub

Private Sub Txt_FecFallBen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If (Trim(Txt_FecFallBen) = "") Then
         MsgBox "Debe Ingresar una Fecha para el Valor Fallecimiento", vbCritical, "Error de Datos"
         Txt_FecFallBen.SetFocus
         Exit Sub
      End If
      If Not IsDate(Txt_FecFallBen.Text) Then
         MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
         Txt_FecFallBen.SetFocus
         Exit Sub
      End If
      If (CDate(Txt_FecFallBen) > CDate(Date)) Then
         MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
         Txt_FecFallBen.SetFocus
         Exit Sub
      End If
      If (Year(Txt_FecFallBen) < 1900) Then
         MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
         Txt_FecFallBen.SetFocus
         Exit Sub
      End If
      Txt_FecFallBen.Text = Format(CDate(Trim(Txt_FecFallBen)), "yyyymmdd")
      Txt_FecFallBen.Text = DateSerial(Mid((Txt_FecFallBen.Text), 1, 4), Mid((Txt_FecFallBen.Text), 5, 2), Mid((Txt_FecFallBen.Text), 7, 2))
    End If
End Sub

Private Sub Txt_FecFallBen_LostFocus()
    If (Trim(Txt_FecFallBen) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_FecFallBen.Text) Then
       Exit Sub
    End If
    If (CDate(Txt_FecFallBen) > CDate(Date)) Then
       Exit Sub
    End If
    If (Year(Txt_FecFallBen) < 1900) Then
       Exit Sub
    End If
    Txt_FecFallBen.Text = Format(CDate(Trim(Txt_FecFallBen)), "yyyymmdd")
    Txt_FecFallBen.Text = DateSerial(Mid((Txt_FecFallBen.Text), 1, 4), Mid((Txt_FecFallBen.Text), 5, 2), Mid((Txt_FecFallBen.Text), 7, 2))
End Sub

Private Sub Txt_FecIniPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmd_BuscaCor.SetFocus
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

Private Sub Txt_FecInv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Txt_FecInv <> "" Then
        If IsDate(Txt_FecInv) Then
            Cmd_CauInv.SetFocus
        End If
    Else
        Cmb_EstCivil.SetFocus
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

Private Sub Txt_FecInvBen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Txt_FecInvBen <> "" Then
        If IsDate(Txt_FecInvBen) Then
            Cmd_BuscarCauInvBen.SetFocus
        End If
    Else
        Btn_Agregar.SetFocus
    End If
End If
End Sub

Private Sub Txt_FecInvBen_LostFocus()
    If Txt_FecInvBen <> "" Then
        If (flValidaFecha(Txt_FecInvBen) = False) Then
            Txt_FecInvBen = ""
            Exit Sub
        End If
        Txt_FecInvBen.Text = Format(CDate(Trim(Txt_FecInvBen)), "yyyymmdd")
        Txt_FecInvBen.Text = DateSerial(Mid((Txt_FecInvBen.Text), 1, 4), Mid((Txt_FecInvBen.Text), 5, 2), Mid((Txt_FecInvBen.Text), 7, 2))
    End If
End Sub

Private Sub Txt_FecVig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Cmb_TipoIdent.Enabled = True Then
        Cmb_TipoIdent.SetFocus
    Else
        Txt_NumIdent.SetFocus
    End If
End If
End Sub

Private Sub Txt_FecVig_LostFocus()
On Error GoTo Err_FecVig

    If Txt_FecVig <> "" Then
        If (Trim(Txt_FecVig) = "") Then
            Txt_FecVig = ""
            Exit Sub
        End If
        If Not IsDate(Txt_FecVig) Then
            Txt_FecVig = ""
            Exit Sub
        End If
        If (CDate(Txt_FecVig) > CDate(Date)) Then
            Txt_FecVig = ""
            Exit Sub
        End If
        If (Year(CDate(Txt_FecVig)) < 1900) Then
            Txt_FecVig = ""
            Exit Sub
        End If
        Txt_FecVig.Text = Format(CDate(Trim(Txt_FecVig)), "yyyymmdd")
        Txt_FecVig.Text = DateSerial(Mid((Txt_FecVig.Text), 1, 4), Mid((Txt_FecVig.Text), 5, 2), Mid((Txt_FecVig.Text), 7, 2))
    End If
Exit Sub
Err_FecVig:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub

Private Sub Txt_Fono_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Txt_Fono2_Afil.SetFocus
    End If
    
End Sub

Private Sub Txt_Fono_LostFocus()

    Txt_Fono = Trim(UCase(Txt_Fono))
    
End Sub





Private Sub Txt_Fono1_Ben_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_Fono1_Ben.SetFocus
    End If
End Sub

Private Sub Txt_Fono2_Afil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_Correo.SetFocus
    End If

End Sub



Private Sub Txt_Fono2_Ben_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkConTratDatos_Ben.SetFocus
    End If
End Sub

Private Sub Txt_Interior_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmb_TipoZona.SetFocus
    End If

End Sub

Private Sub Txt_Interior_LostFocus()
    Txt_Interior.Text = Trim(UCase(Txt_Interior.Text))
    Call pConcatenaDireccion
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
    Txt_Nacionalidad = UCase(Trim(Txt_Nacionalidad))
End Sub

Private Sub Txt_NomAfi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_NomAfiSeg.SetFocus
End If
End Sub

Private Sub Txt_NomAfi_LostFocus()
    Txt_NomAfi = Trim(UCase(Txt_NomAfi))
    Call flDatosCompletos
End Sub

Private Sub Txt_NomAfiSeg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_ApPatAfi.SetFocus
End If
End Sub

Private Sub Txt_NomAfiSeg_LostFocus()
    Txt_NomAfiSeg = Trim(UCase(Txt_NomAfiSeg))
    Call flDatosCompletos
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

Private Sub Txt_NombreVia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_Numero.SetFocus
    End If

End Sub

Private Sub Txt_NombreVia_LostFocus()
    Txt_NombreVia.Text = Trim(UCase(Txt_NombreVia.Text))
    Call pConcatenaDireccion
End Sub

Private Sub Txt_NombreZona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmd_BuscarDir.SetFocus
    End If

End Sub

Private Sub Txt_NombreZona_LostFocus()
    Txt_NombreZona.Text = Trim(UCase(Txt_NombreZona.Text))
    Call pConcatenaDireccion
End Sub

Private Sub Txt_NomRep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_ApPatRep.SetFocus
    End If
End Sub

Private Sub Txt_NomRep_LostFocus()
    Call pConcatenaRepresentante
    Txt_NomRep.Text = Trim(UCase(Txt_NomRep.Text))
End Sub

Private Sub Txt_NumCta_KeyPress(KeyAscii As Integer)

Dim Keychar As String
If KeyAscii > 31 Then
    Keychar = Chr(KeyAscii)
    If Not IsNumeric(Keychar) Then
        KeyAscii = 0
    End If
Else
    If KeyAscii = 13 Then
        Txt_NumCta = Trim(Txt_NumCta)
        txt_CCI.SetFocus
    End If
End If



End Sub
Private Sub txt_CCI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt_CCI = Trim(txt_CCI)
    SSTab_Poliza.Tab = 1
    Txt_FecIniPago.SetFocus
Else
  If KeyAscii > 31 Then
    Keychar = Chr(KeyAscii)
    If Not IsNumeric(Keychar) Then
        KeyAscii = 0
    End If
  End If
  
End If
End Sub

Private Sub Txt_Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_Interior.SetFocus
 
    End If

End Sub

Private Sub Txt_Numero_LostFocus()
    Txt_Numero.Text = Trim(UCase(Txt_Numero.Text))
    Call pConcatenaDireccion
End Sub

Private Sub Txt_NumIdent_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_NumIdent = UCase(Trim(Txt_NumIdent))
    Txt_NomAfi.SetFocus
End If
End Sub

Private Sub Txt_NumIdent_LostFocus()
    Txt_NumIdent = Trim(UCase(Txt_NumIdent))
    Call flDatosCompletos
End Sub

Private Sub Txt_NumIdentBen_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Txt_NumIdentBen = UCase(Trim(Txt_NumIdentBen))
    Txt_NombresBen.SetFocus
End If
End Sub

Private Sub Txt_NumIdentBen_LostFocus()
Txt_NumIdentBen = UCase(Trim(Txt_NumIdentBen))
End Sub

Function flRecibeDireccion(iNomDepartamento As String, iNomProvincia As String, iNomDistrito As String, iCodDir As String)
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA de Dirección
    
    Lbl_Departamento = Trim(iNomDepartamento)
    Lbl_Provincia = Trim(iNomProvincia)
    Lbl_Distrito = Trim(iNomDistrito)
    vlCodDireccion = iCodDir
    
 
    
    'Integracion GobiernoDeDatos_
    vgCodDireccion = iCodDir
    'fin Integracion GobiernoDeDatos_

End Function
Friend Sub RecibeDireccionRepr(DRep As DireccionRep)
   txtTelRep1.Text = DirRep.vNumTelefono
   txtTelRep2.Text = DirRep.vNumTelefono2
   
  
End Sub






Function flRecibeCorredor(iNomTipoIden As String, iNumIden As String, iCodTipoIden As String, iBenSocial As String, iPrcComision As Double)
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA del Corredor

    Lbl_TipoIdentCorr = iCodTipoIden & " - " & iNomTipoIden
    Lbl_NumIdentCorr = iNumIden
    If (iBenSocial <> "") Then
        Lbl_BenSocial = IIf(iBenSocial = "S", cgIndicadorSi, cgIndicadorNo)
    Else
        Lbl_BenSocial = cgIndicadorNo
    End If
    If (Lbl_BenSocial = cgIndicadorNo) Then
        Lbl_ComIntBen = Lbl_ComInt
    End If

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

Private Sub Txt_NumIdRep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_NomRep.SetFocus
    End If
End Sub

Private Sub Txt_NumIdRep_LostFocus()
    Txt_NumIdRep.Text = Trim(UCase(Txt_NumIdRep.Text))
    Call pConcatenaRepresentante
End Sub

Private Sub Txt_Referencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmd_SalirDir.SetFocus
    End If

End Sub

Private Sub Txt_Referencia_LostFocus()
    Txt_Referencia.Text = Trim(UCase(Txt_Referencia.Text))
    Call pConcatenaDireccion
End Sub
Private Sub txtCorreoRep_LostFocus()
    Me.txtCorreoRep.Text = Trim(UCase(Me.txtCorreoRep.Text))
    
bVerifica = Comprobar_Mail(Me.txtCorreoRep.Text)
 If Not bVerifica Then
    MsgBox "El correo ingresado es inválido."
    txtCorreoRep.SetFocus

 End If
End Sub

Private Sub txtNumctaBen_KeyPress(KeyAscii As Integer)
Dim Keychar As String
If KeyAscii > 31 Then
    Keychar = Chr(KeyAscii)
    If Not IsNumeric(Keychar) Then
        KeyAscii = 0
    End If
Else
    If KeyAscii = 13 Then
         txt_CCIBen.SetFocus
    End If
End If
End Sub


'Integracion GobiernoDeDatos(Metodos y funciones agregadas para el flujo de envio al gestor de clientes)_

Public Function EnviarGestorCliente(Operacion As String, Mensaje As String) As Boolean
On Error GoTo Err_Descargar

Call fgBuscarCodigos(pgCodDireccion)

Dim Texto  As String
Dim cadena As String
Dim sInputJson As String
Dim httpURL As Object
Dim response As String

Set httpURL = CreateObject("WinHttp.WinHttpRequest.5.1")

If Operacion = "Validar" Then
        'Pruebas
        'cadena = "http://10.10.1.56/WSGestorClienteQA/Api/Cliente/ValidarCliente"

        'Produccion
        cadena = "http://10.10.1.58/WSGestorCliente/Api/Cliente/ValidarCliente"
        'cadena = "https://soatservicios.protectasecurity.pe/WSGestorCliente/Api/Cliente/ValidarCliente"
Else
        'Pruebas
        'cadena = "http://10.10.1.56/WSGestorClienteQA/Api/Cliente/GestionarCliente"

        'Produccion
        cadena = "http://10.10.1.58/WSGestorCliente/Api/Cliente/GestionarCliente"
        'cadena = "https://soatservicios.protectasecurity.pe/WSGestorCliente/Api/Cliente/GestionarCliente"
End If

 sInputJson = CreateJSON

 httpURL.Open "POST", cadena, False
 httpURL.SetRequestHeader "Content-Type", "application/json"
 httpURL.Send sInputJson
 response = httpURL.ResponseText
  
  
Dim p As Object
Dim CodReturn As String
Dim RspMensaje As String
Set p = json.parse(response)
RspMensaje = p.Item("P_SMESSAGE")
CodReturn = p.Item("P_NCODE")
If CodReturn = "1" Then
Mensaje = RspMensaje
EnviarGestorCliente = False
Else
EnviarGestorCliente = True
End If
Exit Function
Err_Descargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'Se arma el json a enviar_
Private Function CreateJSON() As String
Dim Texto As String
'JSON DATOS CLIENTE _

Dim prueba As String
Dim TelefonoEnviar As String
Dim TelefonoEnviar2 As String
Dim CodigTelEnviar As String
Dim lstRep As RepresentaGC


vNumPoliza = Txt_NumPol.Text
 vlNumCoti = Lbl_NumCot
 
Call Limpiar_RepresentaGC
Call RepresentanteToGC(vNumPoliza, lstRep)



Msf_GriAseg.Col = 21
vlNumOrd = Msf_GriAseg.Text
 
 
vlSql = "SELECT TO_CHAR(SYSDATE,'dd-mm-YYYY') as fecha from dual"
        Set vgRs = vgConexionBD.Execute(vlSql)
        If Not vgRs.EOF Then
        vfechasistema = IIf(IsNull(vgRs!fecha), "", vgRs!fecha)
        End If


vlSql = "SELECT cod_par,cod_grufam,cod_sexo,"
        vlSql = vlSql & "cod_sitinv,TO_CHAR(to_date(fec_sitinv),'dd-MM-YYYY') as fec_sitinv ," 'fec_sitinv,cod_cauinv,"
        vlSql = vlSql & "cod_derpen,cod_dercre,fec_nacben,"
        vlSql = vlSql & "fec_falben,fec_nachm,prc_pension,prc_pensionleg,prc_pensionrep,"
        vlSql = vlSql & "mto_pension,mto_pensiongar "
'I--- ABV 04/12/2009 ---
        vlSql = vlSql & ",prc_pensionsobdif "
'F--- ABV 04/12/2009 ---
        vlSql = vlSql & "FROM pt_tmae_cotben "
        vlSql = vlSql & "WHERE num_cot = '" & vlNumCoti & "' AND "
        vlSql = vlSql & "num_orden = '" & vlNumOrd & "'"
        Set vgRs = vgConexionBD.Execute(vlSql)
        If Not vgRs.EOF Then
        vParentesco = IIf(IsNull(vgRs!Cod_Par), "", vgRs!Cod_Par)
        vSexoBen = IIf(IsNull(vgRs!Cod_Sexo), "", vgRs!Cod_Sexo)
        vIncapaciben = IIf(IsNull(vgRs!Cod_SitInv), "", vgRs!Cod_SitInv)
        vFecIncapaBen = IIf(IsNull(vgRs!fec_sitinv), "", vgRs!fec_sitinv)
        End If
        
        If vParentesco = "99" Then
            vTipoDoc = CStr(fgObtenerCodigo_TextoCompuesto(Cmb_TipoIdent))
            vNumDoc = CStr(Txt_NumIdent.Text)
            vNomben = Trim(CStr(Txt_NomAfi.Text) + " " + CStr(Txt_NomAfiSeg.Text))
            vApePben = CStr(Txt_ApPatAfi.Text)
            vApeMben = CStr(Txt_ApMatAfi.Text)
            vFecIncapaBen = ChangeFechaFormat(CStr(Txt_FecInv.Text))
            Msf_GriAseg.Col = 4
            vIncapaciben = Msf_GriAseg.Text
            vEstadoCivil = CStr(fgObtenerCodigo_TextoCompuesto(Cmb_EstCivil))
            vNacionaBen = CStr(fgObtenerCodigo_TextoCompuesto(cboNacionalidad))
            TelefonoEnviar = pNumTelefono
            TelefonoEnviar2 = pNumTelefono2
            CodigTelEnviar = pCodigoTelefono
                    
            vl_ConTratDatos_Ben = IIf(Trim(chkConTratDatos_Afil.Value) = "1", "1", 2)
            vl_ConUsoDatosCom_Ben = IIf(Trim(chkConUsoDatosCom_Afil.Value) = "1", "1", 2)
        Else
        
                Msf_GriAseg.Col = 11
                vTipoDoc = fgObtenerCodigo_TextoCompuesto(Msf_GriAseg.Text)
                
                Msf_GriAseg.Col = 12
                vNumDoc = Trim(Msf_GriAseg.Text)
                
                Msf_GriAseg.Col = 13
                vNomben = Trim(Msf_GriAseg.Text)
                
                Msf_GriAseg.Col = 14
                vNomben = Trim(vNomben + " " + Msf_GriAseg.Text)
                        
                Msf_GriAseg.Col = 15
                vApePben = Trim(Msf_GriAseg.Text)
        
                Msf_GriAseg.Col = 16
                vApeMben = Trim(Msf_GriAseg.Text)
                                        
                 Msf_GriAseg.Col = 29
                vNacionaBen = Trim(Msf_GriAseg.Text)
                
                
                pTipoTelefono = "4"
                pTipoTelefono2 = "2"
                CodigTelEnviar = pCodigoTelefono
                'CodigTelEnviar = "1"_
                Msf_GriAseg.Col = 32
                TelefonoEnviar = Trim(Msf_GriAseg.Text)
                Msf_GriAseg.Col = 33
                TelefonoEnviar2 = Trim(Msf_GriAseg.Text)
                
                Msf_GriAseg.Col = 34
                vl_ConTratDatos_Ben = IIf(Trim(Msf_GriAseg.Text) = "1", "1", "2")
        
                Msf_GriAseg.Col = 35
                vl_ConUsoDatosCom_Ben = IIf(Trim(Msf_GriAseg.Text) = "1", "1", "2")
        End If
        
        Msf_GriAseg.Col = 9
        vBirtdhayben = ChangeFechaFormat(Format(CDate(Msf_GriAseg.Text), "dd-mm-yyyy"))
           
        
        vFechaMatriBen = ""
        
        
        
        '------------------------------------------------------------------

Texto = "{" + _
        "p_CodAplicacion : 'SEACSA', " + _
        "p_TipOper : 'INS', " + _
        "P_NUSERCODE : '" + vgUsuario + "', " + _
        "P_NIDDOC_TYPE : '" + vTipoDoc + "', " + _
        "P_SIDDOC : '" + vNumDoc + "', " + _
        "P_SFIRSTNAME : '" + vNomben + "', " + _
        "P_SLASTNAME : '" + vApePben + "', " + _
        "P_SLASTNAME2 : '" + vApeMben + "', " + _
        "P_SLEGALNAME : '" + vNomben + "', " + _
        "P_SSEXCLIEN : '" + vSexoBen + "', " + _
        "P_NINCAPACITY : '" + vIncapaciben + "', " + _
        "P_DBIRTHDAT : '" + vBirtdhayben + "', "
'------------------------------------------------------------------


Texto = Texto + _
        "p_DINCAPACITY : '" + CambiarVacioxNull(vFecIncapaBen) + "', " + _
        "P_NSPECIALITY : '99', " + _
        "P_NCIVILSTA : '" + vEstadoCivil + "', " + _
        "P_NTITLE : '99', " + _
        "P_NAFP : '" + vParentesco + "', " + _
        "P_NNATIONALITY : '" + vNacionaBen + "', " + _
        "P_SBAJAMAIL_IND  : '" + vl_ConUsoDatosCom_Ben + "', " + _
        "P_SPROTEG_DATOS_IND  : '" + vl_ConTratDatos_Ben + "', " + _
        "P_SISCLIENT_IND : '1', " + _
        "P_SISRENIEC_IND : '2' , "
        ' "P_SPOLIZA_ELECT_IND : '1', " +
        ' "P_SPROTEG_DATOS_IND : '1', " + _
        ' "P_COD_CUSPP : '123456789012', " + _

'JSON DATOS DIRECCION _

Texto = Texto + "EListAddresClient : [ "
   
    If Not vlBotonEscogido = "C" Then
        If stPolizaBenDirec(vlJ).pDireccionConcat <> "" Then
                    If stPolizaBenDirec(vlJ).pDireccionConcat <> vDireccionConcat Then
                    
                    ' "P_NCOUNTRY : '604', " + _

                      Texto = Texto + _
                        "{" + _
                        "P_TIPOPER : 'DEL', " + _
                        "P_SRECTYPE : '2', " + _
                        "P_NCOUNTRY : '1', " + _
                        "P_NPROVINCE : '" + stPolizaBenDirec(vlJ).codRegion + "', " + _
                        "P_NLOCAL : '" + stPolizaBenDirec(vlJ).codProvincia + "', " + _
                        "P_NMUNICIPALITY : '" + stPolizaBenDirec(vlJ).codComuna + "', " + _
                        "P_STI_DIRE : '" + stPolizaBenDirec(vlJ).pTipoVia + "', " + _
                        "P_SNOM_DIRECCION : '" + stPolizaBenDirec(vlJ).pDireccion + "', " + _
                        "P_SNUM_DIRECCION : '" + stPolizaBenDirec(vlJ).pNumero + "', " + _
                        "P_STI_BLOCKCHALET : '" + stPolizaBenDirec(vlJ).pTipoBlock + "', " + _
                        "P_SBLOCKCHALET : '" + stPolizaBenDirec(vlJ).pNumBlock + "', " + _
                        "P_STI_INTERIOR : '" + stPolizaBenDirec(vlJ).pTipoPref + "', " + _
                        "P_SNUM_INTERIOR : '" + stPolizaBenDirec(vlJ).pInterior + "', " + _
                        "P_STI_CJHT : '" + stPolizaBenDirec(vlJ).pTipoConj + "', " + _
                        "P_SNOM_CJHT : '" + stPolizaBenDirec(vlJ).pConjHabit + "', " + _
                        "P_SETAPA  : '" + stPolizaBenDirec(vlJ).pEtapa + "', " + _
                        "P_SMANZANA : '" + stPolizaBenDirec(vlJ).pManzana + "', " + _
                        "P_SLOTE : '" + stPolizaBenDirec(vlJ).pLote + "', " + _
                        "P_SREFERENCIA : '" + stPolizaBenDirec(vlJ).pReferencia + "'" + _
                        "},"
                    End If
        End If
 End If
   
    If Trim(Lbl_Dir) <> "" Then
    
    Texto = Texto + _
            "{" + _
            "P_SRECTYPE : '2', " + _
            "P_NCOUNTRY : '1', " + _
            "P_NPROVINCE : '" + codRegion + "', " + _
            "P_NLOCAL : '" + codProvincia + "', " + _
            "P_NMUNICIPALITY : '" + codComuna + "', " + _
            "P_STI_DIRE : '" + CambiarVacioxNull(pTipoVia) + "', " + _
            "P_SNOM_DIRECCION : '" + CambiarVacioxNull(pDireccion) + "', " + _
            "P_SNUM_DIRECCION : '" + CambiarVacioxNull(pNumero) + "', " + _
            "P_STI_BLOCKCHALET : '" + CambiarVacioxNull(pTipoBlock) + "', " + _
            "P_SBLOCKCHALET : '" + CambiarVacioxNull(pNumBlock) + "', " + _
            "P_STI_INTERIOR : '" + CambiarVacioxNull(pTipoPref) + "', " + _
            "P_SNUM_INTERIOR : '" + CambiarVacioxNull(pInterior) + "', " + _
            "P_STI_CJHT : '" + CambiarVacioxNull(pTipoConj) + "', " + _
            "P_SNOM_CJHT : '" + CambiarVacioxNull(pConjHabit) + "', " + _
            "P_SETAPA  : '" + CambiarVacioxNull(pEtapa) + "', " + _
            "P_SMANZANA : '" + CambiarVacioxNull(pManzana) + "', " + _
            "P_SLOTE : '" + CambiarVacioxNull(pLote) + "', " + _
            "P_SREFERENCIA : '" + CambiarVacioxNull(pReferencia) + "'" + _
            "} "
    End If


Texto = Mid(Texto, 1, Len(Texto) - 1)


'JSON DATOS TELEFONO_

Texto = Texto + "],EListPhoneClient : [ "

    If Not vlBotonEscogido = "C" Then
        If stPolizaBenDirec(vlJ).cod_tip_fonoben <> Trim(pTipoTelefono) Or stPolizaBenDirec(vlJ).cod_area_fonoben <> pCodigoTelefono Or stPolizaBenDirec(vlJ).Gls_FonoBen <> TelefonoEnviar Then
            Texto = Texto + _
            "{" + _
            "P_TIPOPER : 'DEL', " + _
            "P_NAREA_CODE : '" + stPolizaBenDirec(vlJ).cod_area_fonoben + "', " + _
            "P_SPHONE : '" + stPolizaBenDirec(vlJ).Gls_FonoBen + "', " + _
            "P_NPHONE_TYPE : '" + stPolizaBenDirec(vlJ).cod_tip_fonoben + "'" + _
            "},"
            
        End If
        
            If stPolizaBenDirec(vlJ).cod_tipo_telben2 <> Trim(pTipoTelefono2) Or stPolizaBenDirec(vlJ).cod_area_telben2 <> pCodigoTelefono2 Or stPolizaBenDirec(vlJ).gls_telben2 <> TelefonoEnviar2 Then
            Texto = Texto + _
            "{" + _
            "P_TIPOPER : 'DEL', " + _
            "P_NAREA_CODE : '" + stPolizaBenDirec(vlJ).cod_area_telben2 + "', " + _
            "P_SPHONE : '" + stPolizaBenDirec(vlJ).gls_telben2 + "', " + _
            "P_NPHONE_TYPE : '" + stPolizaBenDirec(vlJ).cod_tipo_telben2 + "'" + _
            "},"
        
        End If
        
    End If

    If Trim(TelefonoEnviar) <> "" Then
        Texto = Texto + _
        "{" + _
        "P_NAREA_CODE : '" + CambiarVacioxNull(CodigTelEnviar) + "', " + _
        "P_SPHONE : '" + CambiarVacioxNull(TelefonoEnviar) + "', " + _
        "P_NPHONE_TYPE : '" + CambiarVacioxNull(pTipoTelefono) + "'" + _
        "},"
    End If
    
    If Trim(TelefonoEnviar2) <> "" Then
            Texto = Texto + _
            "{" + _
            "P_NAREA_CODE : '" + CambiarVacioxNull(pCodigoTelefono2) + "', " + _
            "P_SPHONE : '" + CambiarVacioxNull(TelefonoEnviar2) + "', " + _
            "P_NPHONE_TYPE : '" + CambiarVacioxNull(pTipoTelefono2) + "'" + _
            "} "
    End If

Texto = Mid(Texto, 1, Len(Texto) - 1)

'JSON DATOS EMAIL_

Texto = Texto + "],EListEmailClient : [ "

   If Not vlBotonEscogido = "C" Then
        If stPolizaBenDirec(vlJ).pGlsCorreo <> Trim(Txt_Correo.Text) Then
        Texto = Texto + _
           "{" + _
           "P_TIPOPER : 'DEL', " + _
           "P_SRECTYPE : '5', " + _
           "P_SE_MAIL : '" + UCase(stPolizaBenDirec(0).pGlsCorreo) + "'  " + _
           "},"
        End If
    End If


    If Trim(Txt_Correo.Text) <> "" Then
    
    Texto = Texto + _
    "{" + _
    "P_NROW : '1', " + _
    "P_SRECTYPE : '4', " + _
    "P_SE_MAIL : '" + CambiarVacioxNull(Txt_Correo.Text) + "'  " + _
    "} "
    
    End If

Texto = Mid(Texto, 1, Len(Texto) - 1)


'JSON DATOS CONTACTO_


Texto = Texto + "],EListContactClient : [ "

If vParentesco = "99" Then

        For vlJP = 1 To (Msf_GriAseg.rows - 1)
         Msf_GriAseg.Row = vlJP
         Dim CodPar As String
         Dim nombre As String
         Dim ApePat As String
         Dim ApeMat As String
         Dim TipDoc As String
         Dim NumDoc As String
         Dim telefono As String
         Msf_GriAseg.Col = 1
         CodPar = Msf_GriAseg.Text
         
         Msf_GriAseg.Col = 11
         TipDoc = fgObtenerCodigo_TextoCompuesto(Msf_GriAseg.Text)
             
         Msf_GriAseg.Col = 12
         NumDoc = Trim(Msf_GriAseg.Text)
         
         Msf_GriAseg.Col = 13
         nombre = Trim(Msf_GriAseg.Text)
         
         Msf_GriAseg.Col = 14
         nombre = nombre & " " & Trim(Msf_GriAseg.Text)
         
         Msf_GriAseg.Col = 15
         ApePat = Trim(Msf_GriAseg.Text)
         
         Msf_GriAseg.Col = 16
         ApeMat = Trim(Msf_GriAseg.Text)
         
        Msf_GriAseg.Col = 32
        telefono = Trim(Msf_GriAseg.Text)
        If telefono = "" Then
        Msf_GriAseg.Col = 33
        telefono = Trim(Msf_GriAseg.Text)
        End If
        
        
         If CodPar <> "99" Then
         
         Texto = Texto + _
             "{" + _
             "P_TIPOPER : 'DEL', " + _
             "P_NTIPCONT : '" & CodPar & "', " + _
             "P_NIDDOC_TYPE :'" & TipDoc & "', " + _
             "P_SIDDOC :'" & NumDoc & "' " + _
             "},"

         
         Texto = Texto + _
             "{" + _
             "P_NTIPCONT : '" & CodPar & "', " + _
             "P_NIDDOC_TYPE :'" & TipDoc & "', " + _
             "P_SIDDOC :'" & NumDoc & "', " + _
             "P_SNOMBRES :'" & nombre & "', " + _
             "P_SAPEPAT :'" & ApePat & "', " + _
             "P_SAPEMAT :'" & ApeMat & "', " + _
             "P_SPHONE :'" & IIf(telefono = "", pNumTelefono, telefono) & "'" + _
             "}"
             
             If vlJP <> Msf_GriAseg.rows - 1 Then
             Texto = Texto + ","
             End If
             
        End If
        Next vlJP

End If

'INICIO***********Añadimos al representante si es que no es beneficiario*********************

If lstRep.P_DATOS = "S" Then

  Texto = Texto + _
             "{" + _
             "P_TIPOPER : 'DEL', " + _
             "P_NTIPCONT : '10 ', " + _
             "P_NIDDOC_TYPE :'" & lstRep.P_NIDDOC_TYPE & "', " + _
             "P_SIDDOC :'" & lstRep.P_SIDDOC & "' " + _
             "},"

         Texto = Texto + _
             "{" + _
             "P_NTIPCONT : '10 ', " + _
             "P_NIDDOC_TYPE :'" & lstRep.P_NIDDOC_TYPE & "', " + _
             "P_SIDDOC :'" & lstRep.P_SIDDOC & "', " + _
             "P_SNOMBRES :'" & lstRep.P_SNOMBRES & "', " + _
             "P_SAPEPAT :'" & lstRep.P_SAPEPAT & "', " + _
             "P_SAPEMAT :'" & lstRep.P_SAPEMAT & "', " + _
             "P_SPHONE :'" & lstRep.P_SPHONE & "'" + _
             "}"
End If
'FIN***********Añadimos al representante*********************

 Texto = Texto + _
         "]," + _
         "EListCIIUClient : Null " + _
         "}"


Dim valor As Object
Set valor = json.parse(Texto)
CreateJSON = json.toString(valor)



End Function

Function fgBuscarCodigos(vlCodDir As String) As String
Dim vlRegistroDir As ADODB.Recordset
On Error GoTo Err_Buscar
     
    vgSql = "SELECT c.cod_comuna,c.cod_provincia,c.cod_region  "
    vgSql = vgSql & " FROM MA_TPAR_COMUNA c "
    vgSql = vgSql & " Where c.Cod_Direccion = '" & vlCodDir & "'"
    
    Set vlRegistroDir = vgConexionBD.Execute(vgSql)
    If Not vlRegistroDir.EOF Then
       codComuna = (vlRegistroDir!cod_comuna)
       codProvincia = (vlRegistroDir!COD_PROVINCIA)
       codRegion = (vlRegistroDir!cod_region)
    End If
    vlRegistroDir.Close
Exit Function
Err_Buscar:
   Screen.MousePointer = 0
   Select Case Err
       Case Else
       MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
   End Select
End Function

Private Function CambiarVacioxNull(ByVal valor As String) As String
If Trim(valor) = "" Then
 valor = ""
End If
CambiarVacioxNull = valor
End Function
Private Function ChangeFechaFormat(ByVal fecha As String) As String
Dim dia As String
Dim mes As String
Dim periodo As String
If fecha <> "" Then
dia = Mid(fecha, 1, 2)
mes = Mid(fecha, 4, 2)
periodo = Mid(fecha, 7, 4)
ChangeFechaFormat = dia & "-" & mes & "-" & periodo
Else
ChangeFechaFormat = fecha
End If
End Function
Function flRecibeParametrosEdit(vgCodDireccion As String, vTipoTelefono As String, vNumTelefono As String, vCodigoTelefono As String, vTipoTelefono2 As String, vNumTelefono2 As String, vCodigoTelefono2 As String, vTipoVia As String, vDireccion As String, vNumero As String, vTipoPref As String, vInterior As String, vManzana As String, vLote As String, vEtapa As String, vTipoConj As String, vConjHabit As String, vTipoBlock As String, vNumBlock As String, vReferencia As String)
pgCodDireccion = UCase(NullxVacio(vgCodDireccion))
vlCodDireccion = UCase(NullxVacio(vgCodDireccion))
pTipoTelefono = UCase(NullxVacio(vTipoTelefono))
pNumTelefono = UCase(NullxVacio(vNumTelefono))
pCodigoTelefono = UCase(NullxVacio(vCodigoTelefono))
pTipoTelefono2 = UCase(NullxVacio(vTipoTelefono2))
pNumTelefono2 = UCase(NullxVacio(vNumTelefono2))
pCodigoTelefono2 = UCase(NullxVacio(vCodigoTelefono2))
pTipoVia = UCase(NullxVacio(vTipoVia))
pDireccion = UCase(NullxVacio(vDireccion))
pNumero = UCase(NullxVacio(vNumero))
pTipoPref = UCase(NullxVacio(vTipoPref))
pInterior = UCase(NullxVacio(vInterior))
pManzana = UCase(NullxVacio(vManzana))
pLote = UCase(NullxVacio(vLote))
pEtapa = UCase(NullxVacio(vEtapa))
pTipoConj = UCase(NullxVacio(vTipoConj))
pConjHabit = UCase(NullxVacio(vConjHabit))
pTipoBlock = UCase(NullxVacio(vTipoBlock))
pNumBlock = UCase(NullxVacio(vNumBlock))
pReferencia = UCase(NullxVacio(vReferencia))
vlCodDireccion = UCase(NullxVacio(vgCodDireccion))
vgCodDireccion = UCase(NullxVacio(vgCodDireccion))
Txt_Fono = UCase(NullxVacio(vNumTelefono))
Txt_Fono2_Afil = UCase(NullxVacio(vNumTelefono2))
Frm_CalPoliza.Enabled = True
Call ConcatenarDireccion
End Function
Private Sub ConcatenarDireccion()
Dim vlRegistroDir As ADODB.Recordset
On Error GoTo Err_Buscar
 Dim tipo_via As String
 Dim tipo_bloque As String
 Dim tipo_interior As String
 Dim tipo_cjht As String
Call fgBuscarNombreComunaProvinciaRegion(pgCodDireccion)

     vgSql = "SELECT"
     vgSql = vgSql + "(SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_VIA T "
     vgSql = vgSql + " WHERE T.COD_DIRE_VIA = '" + pTipoVia + "') as TIPO_VIA,"
     vgSql = vgSql + " (SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_BLOQUE T"
     vgSql = vgSql + " WHERE T.COD_BLOCKCHALET =  '" + pTipoBlock + "') AS TIPO_BLOQUE,"
     vgSql = vgSql + " (SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_INTERIOR T"
     vgSql = vgSql + " WHERE T.COD_INTERIOR = '" + pTipoPref + "') AS TIPO_INTERIOR,"
     vgSql = vgSql + "(SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_CJHT T"
     vgSql = vgSql + " WHERE T.COD_CJHT = '" + pTipoConj + "' ) AS TIPO_CJHT"
     vgSql = vgSql + " FROM DUAL"
     Set vlRegistroDir = vgConexionBD.Execute(vgSql)
     If Not vlRegistroDir.EOF Then
        tipo_via = IIf(IsNull(vlRegistroDir!tipo_via), "", vlRegistroDir!tipo_via)
        tipo_bloque = IIf(IsNull(vlRegistroDir!tipo_bloque), "", vlRegistroDir!tipo_bloque)
        tipo_interior = IIf(IsNull(vlRegistroDir!tipo_interior), "", vlRegistroDir!tipo_interior)
        tipo_cjht = IIf(IsNull(vlRegistroDir!tipo_cjht), "", vlRegistroDir!tipo_cjht)
     End If
     vlRegistroDir.Close
     Dim Strmanzana As String
     If Trim(pManzana) <> "" Then
       Strmanzana = "Manzana " & pManzana & " "
     Else
       Strmanzana = ""
     End If
     
     Dim StrLote As String
     If Trim(pLote) <> "" Then
       StrLote = "Lote " & pLote & " "
     Else
       StrLote = ""
     End If
     Dim StrEtapa As String
     If Trim(pEtapa) <> "" Then
       StrEtapa = "Etapa " & pEtapa & " "
     Else
       StrEtapa = ""
     End If
     Dim strBloque As String
     If Trim(tipo_bloque) <> "" And Trim(pNumBlock) <> "" Then
     strBloque = tipo_bloque & " " & Trim(pNumBlock) & " "
     Else
     strBloque = ""
     End If
     Dim StrInterior As String
     If Trim(tipo_interior) <> "" And Trim(pInterior) <> "" Then
     StrInterior = tipo_interior & " " & Trim(pInterior) & " "
     Else
     StrInterior = ""
     End If
     Dim StrCjht As String
     If pTipoConj = "99" Then
      StrCjht = pConjHabit & " "
     Else
        If Trim(tipo_cjht) <> "" And Trim(pConjHabit) <> "" Then
        StrCjht = tipo_cjht & " " & pConjHabit & " "
        Else
        StrCjht = ""
        End If
     End If
     Dim StrDireccion As String
     If pTipoVia = "99" Or pTipoVia = "88" Then
     StrDireccion = "" & pDireccion
    Else
        If Trim(tipo_via) <> "" And Trim(pDireccion) <> "" Then
        StrDireccion = tipo_via & " " & pDireccion
        Else
        StrDireccion = ""
        End If
     
     End If
     
    vDireccionConcat = Trim(UCase(StrDireccion & " " & IIf(Trim(pNumero) = "", "", Trim(pNumero) & " ") _
        & strBloque & StrInterior _
        & StrCjht _
        & StrEtapa & Strmanzana & StrLote _
        & " " & IIf(Trim(pReferencia) = "", "", Trim(pReferencia) & " ") _
        & IIf(Trim(vgNombreRegion) = "", "", Trim(vgNombreRegion) & " ") _
        & IIf(Trim(vgNombreProvincia) = "", "", Trim(vgNombreProvincia) & " ") _
        & IIf(Trim(vgNombreComuna) = "", "", Trim(vgNombreComuna) & " ")))
        
       Lbl_Dir = vDireccionConcat

        
 Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Function NullxVacio(valor As String) As String
If valor = "Null" Then
valor = ""
End If
NullxVacio = valor
End Function
Private Sub LimpiarVariables()
codComuna = ""
codRegion = ""
codProvincia = ""
pgCodDireccion = ""
pTipoTelefono = ""
pNumTelefono = ""
pCodigoTelefono = ""
pgCodDireccion2 = ""
pTipoTelefono2 = ""
pNumTelefono2 = ""
pTipoVia = ""
pDireccion = ""
pNumero = ""
pTipoPref = ""
pInterior = ""
pManzana = ""
pLote = ""
pEtapa = ""
pTipoConj = ""
pConjHabit = ""
pTipoBlock = ""
pNumBlock = ""
pReferencia = ""
End Sub
Private Sub ObtenerDIrecTelef(ByVal iNumPol As String)

    vlSql = "SELECT "
    vlSql = vlSql & " P.NUM_POLIZA,P.NUM_ENDOSO,P.COD_DIRE_VIA,P.GLS_DIRECCION,P.NUM_DIRECCION,"
    vlSql = vlSql & " p.COD_BLOCKCHALET,p.GLS_BLOCKCHALET,p.COD_INTERIOR,p.NUM_INTERIOR,P.COD_CJHT,"
    vlSql = vlSql & " P.GLS_NOM_CJHT,p.GLS_ETAPA,p.GLS_MANZANA, p.GLS_LOTE,P.GLS_REFERENCIA,P.COD_PAIS,P.COD_DISTRITO,P.GLS_DESDIREBUSQ"
    vlSql = vlSql & " FROM PP_TMAE_BEN_DIRECCION p "
    vlSql = vlSql & " WHERE P.NUM_POLIZA = '" & iNumPol & "' "
    vlSql = vlSql & " AND P.NUM_ENDOSO = (SELECT MAX(DIRE.NUM_ENDOSO) FROM PP_TMAE_BEN_DIRECCION DIRE"
    vlSql = vlSql & " WHERE DIRE.NUM_POLIZA = '" & iNumPol & "')"
    vlSql = vlSql & " AND P.NUM_ORDEN = 1"
    
    Set vgRs = vgConexionBD.Execute(vlSql)
    
    If Not (vgRs.EOF) Then
    pTipoVia = IIf(IsNull(vgRs!cod_dire_via), "", vgRs!cod_dire_via)
    pDireccion = IIf(IsNull(vgRs!Gls_Direccion), "", vgRs!Gls_Direccion)
    pNumero = IIf(IsNull(vgRs!num_direccion), "", vgRs!num_direccion)
    pTipoPref = IIf(IsNull(vgRs!cod_interior), "", vgRs!cod_interior)
    pInterior = IIf(IsNull(vgRs!num_interior), "", vgRs!num_interior)
    pManzana = IIf(IsNull(vgRs!gls_manzana), "", vgRs!gls_manzana)
    pLote = IIf(IsNull(vgRs!gls_lote), "", vgRs!gls_lote)
    pEtapa = IIf(IsNull(vgRs!gls_etapa), "", vgRs!gls_etapa)
    pTipoConj = IIf(IsNull(vgRs!cod_cjht), "", vgRs!cod_cjht)
    pConjHabit = IIf(IsNull(vgRs!gls_nom_cjht), "", vgRs!gls_nom_cjht)
    pTipoBlock = IIf(IsNull(vgRs!cod_blockchalet), "", vgRs!cod_blockchalet)
    pNumBlock = IIf(IsNull(vgRs!gls_blockchalet), "", vgRs!gls_blockchalet)
    pReferencia = IIf(IsNull(vgRs!gls_referencia), "", vgRs!gls_referencia)
    Lbl_Dir = Trim(IIf(IsNull(vgRs!gls_desdirebusq), "", vgRs!gls_desdirebusq))
    End If
    
    vlSql = "SELECT "
    vlSql = vlSql & " P.NUM_POLIZA,P.NUM_ENDOSO,P.COD_TIPO_FONOBEN,P.COD_AREA_FONOBEN,P.GLS_FONOBEN,P.COD_TIPO_TELBEN2,P.COD_AREA_TELBEN2,P.GLS_TELBEN2"
    vlSql = vlSql & " FROM PP_TMAE_BEN_TELEFONO p "
    vlSql = vlSql & " WHERE P.NUM_POLIZA = '" & iNumPol & "' "
    vlSql = vlSql & " AND P.NUM_ENDOSO = (SELECT MAX(TELE.NUM_ENDOSO) FROM PP_TMAE_BEN_TELEFONO TELE"
    vlSql = vlSql & " WHERE TELE.NUM_POLIZA = '" & iNumPol & "')"
    vlSql = vlSql & " AND P.NUM_ORDEN = 1"
    
    Set vgRs = vgConexionBD.Execute(vlSql)
    
    If Not (vgRs.EOF) Then
    pTipoTelefono = IIf(IsNull(vgRs!COD_TIPO_FONOBEN), "", vgRs!COD_TIPO_FONOBEN)
    pNumTelefono = IIf(IsNull(vgRs!Gls_FonoBen), "", vgRs!Gls_FonoBen)
    pCodigoTelefono = IIf(IsNull(vgRs!cod_area_fonoben), "", vgRs!cod_area_fonoben)
    pTipoTelefono2 = IIf(IsNull(vgRs!cod_tipo_telben2), "", vgRs!cod_tipo_telben2)
    pNumTelefono2 = IIf(IsNull(vgRs!gls_telben2), "", vgRs!gls_telben2)
    pCodigoTelefono2 = IIf(IsNull(vgRs!cod_area_telben2), "", vgRs!cod_area_telben2)
    End If
End Sub
Private Sub ObtenerDatosIniciales(pNumPoliza As String)

ReDim stPolizaBenDirec(Msf_GriAseg.rows - 1) As TyBeneficiariosEst


For vlJ = 1 To (Msf_GriAseg.rows - 1)
Msf_GriAseg.Row = vlJ

Dim Orden  As Integer
Msf_GriAseg.Col = 21
Orden = Msf_GriAseg.Text

vlSql = "SELECT D.num_poliza,D.NUM_ENDOSO,D.NUM_ORDEN,"
    vlSql = vlSql & "D.COD_DIRE_VIA as COD_DIRE_VIA,D.GLS_DIRECCION as GLS_DIRECCION ,D.NUM_DIRECCION AS NUM_DIRECCION, "
    vlSql = vlSql & "D.COD_BLOCKCHALET,D.GLS_BLOCKCHALET as GLS_BLOCKCHALET , "
    vlSql = vlSql & "D.COD_INTERIOR AS COD_INTERIOR,D.NUM_INTERIOR as NUM_INTERIOR , "
    vlSql = vlSql & "D.COD_CJHT AS COD_CJHT ,D.GLS_NOM_CJHT as GLS_NOM_CJHT , "
    vlSql = vlSql & "D.GLS_ETAPA AS GLS_ETAPA ,D.GLS_MANZANA,D.GLS_LOTE,D.GLS_REFERENCIA,D.COD_PAIS,D.COD_DEPARTAMENTO,D.COD_PROVINCIA,D.COD_DISTRITO,D.GLS_DESDIREBUSQ,D.COD_PAIS,D.COD_DEPARTAMENTO,D.COD_PROVINCIA,D.COD_DISTRITO "
    vlSql = vlSql & "FROM PP_TMAE_BEN_DIRECCION D "
    vlSql = vlSql & "WHERE D.NUM_POLIZA = '" & pNumPoliza & "' "
    vlSql = vlSql & "AND D.NUM_ORDEN = " & Orden
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not (vgRs.EOF) Then
    stPolizaBenDirec(vlJ).pTipoVia = IIf(IsNull(vgRs!cod_dire_via), "", vgRs!cod_dire_via)
    stPolizaBenDirec(vlJ).pDireccion = IIf(IsNull(vgRs!Gls_Direccion), "", vgRs!Gls_Direccion)
    stPolizaBenDirec(vlJ).pNumero = IIf(IsNull(vgRs!num_direccion), "", vgRs!num_direccion)
    stPolizaBenDirec(vlJ).pTipoBlock = IIf(IsNull(vgRs!cod_blockchalet), "", vgRs!cod_blockchalet)
    stPolizaBenDirec(vlJ).pNumBlock = IIf(IsNull(vgRs!gls_blockchalet), "", vgRs!gls_blockchalet)
    stPolizaBenDirec(vlJ).pTipoPref = IIf(IsNull(vgRs!cod_interior), "", vgRs!cod_interior)
    stPolizaBenDirec(vlJ).pInterior = IIf(IsNull(vgRs!num_interior), "", vgRs!num_interior)
    stPolizaBenDirec(vlJ).pTipoConj = IIf(IsNull(vgRs!cod_cjht), "", vgRs!cod_cjht)
    stPolizaBenDirec(vlJ).pConjHabit = IIf(IsNull(vgRs!gls_nom_cjht), "", vgRs!gls_nom_cjht)
    stPolizaBenDirec(vlJ).pEtapa = IIf(IsNull(vgRs!gls_etapa), "", vgRs!gls_etapa)
    stPolizaBenDirec(vlJ).pManzana = IIf(IsNull(vgRs!gls_manzana), "", vgRs!gls_manzana)
    stPolizaBenDirec(vlJ).pLote = IIf(IsNull(vgRs!gls_lote), "", vgRs!gls_lote)
    stPolizaBenDirec(vlJ).pReferencia = IIf(IsNull(vgRs!gls_referencia), "", vgRs!gls_referencia)
    stPolizaBenDirec(vlJ).pDireccionConcat = IIf(IsNull(vgRs!gls_desdirebusq), "", vgRs!gls_desdirebusq)
    stPolizaBenDirec(vlJ).codComuna = IIf(IsNull(vgRs!cod_distrito), "", vgRs!cod_distrito)
    stPolizaBenDirec(vlJ).codProvincia = IIf(IsNull(vgRs!COD_PROVINCIA), "", vgRs!COD_PROVINCIA)
    stPolizaBenDirec(vlJ).codRegion = IIf(IsNull(vgRs!cod_departamento), "", vgRs!cod_departamento)
    stPolizaBenDirec(vlJ).cod_pais = IIf(IsNull(vgRs!cod_pais), "", vgRs!cod_pais)
    
    End If
    
vlSql = "SELECT D.num_poliza,D.NUM_ENDOSO,D.NUM_ORDEN,"
    vlSql = vlSql & "D.COD_TIPO_FONOBEN as COD_TIPO_FONOBEN,D.COD_AREA_FONOBEN as COD_AREA_FONOBEN,D.GLS_FONOBEN AS GLS_FONOBEN,D.COD_TIPO_TELBEN2,D.COD_AREA_TELBEN2,D.GLS_TELBEN2  "
    vlSql = vlSql & "FROM PP_tMAE_BEN_TELEFONO D "
    vlSql = vlSql & "WHERE D.NUM_POLIZA = '" & pNumPoliza & "' "
    vlSql = vlSql & "AND D.NUM_ORDEN = " & Orden
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not (vgRs.EOF) Then
    stPolizaBenDirec(vlJ).cod_tip_fonoben = IIf(IsNull(vgRs!COD_TIPO_FONOBEN), "", vgRs!COD_TIPO_FONOBEN)
    stPolizaBenDirec(vlJ).cod_area_fonoben = IIf(IsNull(vgRs!cod_area_fonoben), "", vgRs!cod_area_fonoben)
    stPolizaBenDirec(vlJ).Gls_FonoBen = IIf(IsNull(vgRs!Gls_FonoBen), "", vgRs!Gls_FonoBen)
    stPolizaBenDirec(vlJ).cod_tipo_telben2 = IIf(IsNull(vgRs!cod_tipo_telben2), "", vgRs!cod_tipo_telben2)
    stPolizaBenDirec(vlJ).cod_area_telben2 = IIf(IsNull(vgRs!cod_area_telben2), "", vgRs!cod_area_telben2)
    stPolizaBenDirec(vlJ).gls_telben2 = IIf(IsNull(vgRs!gls_telben2), "", vgRs!gls_telben2)
    End If

vlSql = "SELECT p.num_poliza,p.num_cot,p.fec_vigencia,"
    vlSql = vlSql & "p.cod_tipoidenafi as cod_tipoiden , "
    vlSql = vlSql & "p.gls_correo "
    vlSql = vlSql & "FROM pd_tmae_oripoliza p, pd_tmae_oripolben b "
    vlSql = vlSql & "WHERE p.num_poliza = '" & pNumPoliza & "' "
    vlSql = vlSql & "AND p.num_poliza = b.num_poliza "
    vlSql = vlSql & "AND b.cod_par = '99' "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not (vgRs.EOF) Then
    stPolizaBenDirec(vlJ).pGlsCorreo = IIf(IsNull(vgRs!GLS_CORREO), "", vgRs!GLS_CORREO)
    End If
    
    Next vlJ
End Sub
Private Sub txtTelRep1_LostFocus()
 Me.txtTelRep1.Text = Trim(UCase(Me.txtTelRep1.Text))
 
End Sub
Private Sub txtTelRep2_LostFocus()
Me.txtTelRep2.Text = Trim(UCase(Me.txtTelRep2.Text))
End Sub

Private Function Comprobar_Mail(Direccion As String) As Boolean

On Error GoTo ErrFunction
    
    Dim oReg As RegExp
   ' Crea un Nuevo objeto RegExp
    Set oReg = New RegExp
    
   If Len(Trim(Direccion)) > 0 Then
        ' Expresión regular
        oReg.Pattern = "^[\w-\.]+@\w+\.\w+$"
        ' Comprueba y Retorna TRue o false
        Comprobar_Mail = oReg.Test(Direccion)
    
    Else
       Comprobar_Mail = True
   End If
 
    Set oReg = Nothing
    Exit Function
    
    'Error
ErrFunction:
    
    MsgBox Err.Description, vbCritical
    
    If Not oReg Is Nothing Then
    Set oReg = Nothing
    End If

End Function

Public Function ValidaBancoPrincipal() As Boolean
        
    vlNumero = InStr(Cmb_ViaPago.Text, "-")
    vlOpcion = Trim(Mid(Cmb_ViaPago.Text, 1, vlNumero - 1))
    
    'If vlOpcion = "07" Or vlOpcion = "02" Then
    If txt_CCI.Enabled Then
    
            
        
        If Left(Trim(Cmb_Bco.Text), 2) <> "02" And _
                Left(Trim(Cmb_Bco.Text), 2) <> "03" And _
                Left(Trim(Cmb_Bco.Text), 2) <> "11" And _
                Left(Trim(Cmb_Bco.Text), 2) <> "41" Then
                
                   'BCP = 02
                   'INTERBANK = 03
                   'BBVA = 11
                   'SCOTIABANK = 41
                    
                    If Len(Trim(txt_CCI)) = 0 Then
                          ValidaBancoPrincipal = False
                          
                          Else
                          
                          ValidaBancoPrincipal = True
                          
                       End If
          Else
            ValidaBancoPrincipal = True
            
          End If
          
    Else
     ValidaBancoPrincipal = True
    
    End If
   
 
End Function
Private Sub cmdEnviaCorreo_Click()

    Dim Correos() As String
    Dim MensajeError As String
    Dim SeguirSinCorreo As Byte
    
    'SeguirSinCorreo
    '0: Sigue Flujo completo
    '1: No Continnua por validacion de Correos
    '2: Solo se graba ticket pero no debe continuar flujo
      
    Dim NumError As Integer
    Dim ListaPDFS As String
    Dim ResValidaEnvio As String
    
    
    On Error GoTo Msgerror
    
    NumError = 0
    MensajeError = ""
    SeguirSinCorreo = 0
    
    ResValidaEnvio = ValidaEnvio(Lbl_NumCot.Caption)
     
   If ResValidaEnvio <> "N" Then
    
        MsgBox "Esta póliza ya fue enviada electrónicamente." + Chr(13) + ResValidaEnvio, vbExclamation, "Envio Eléctronico."
        Exit Sub

   End If
   
    Screen.MousePointer = 11

    Call cargaDatos_poliza(Txt_NumPol.Text)
    Call EnvioEmails_poliza(Txt_NumPol.Text, Correos, MensajeError)
    
    If Len(MensajeError) > 0 Then
       If MsgBox(MensajeError + " ¿Desea seguir el proceso sin envío de email?", vbQuestion + vbYesNo, "Verificacion de datos") = vbYes Then
                SeguirSinCorreo = 2
        Else
                SeguirSinCorreo = 1
               
       End If
    End If
    
    If SeguirSinCorreo = 1 Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    
        Dim glob As New ChilkatGlobal
        Dim success As Long
        Dim vTokenDavidCloud As String
        Dim VidDocumento As String
        Dim vToken As String
        Dim vTicket As String
        
        success = glob.UnlockBundle("GVNCRZ.CB1032023_x4BpcXzLDR4D")
        If (success <> 1) Then
            Debug.Print glob.LastErrorText
            Exit Sub
        End If
   
        vToken = ObtenerToken
    
        If vToken = "" Then
             MsgBox "No se Pudo obtener el token", vbCritical, "Notificación Electronica"
             Screen.MousePointer = 0
             Exit Sub
        End If
        
        vTicket = CreaTicket(vToken, SeguirSinCorreo)
        
        If vTicket = "error" Then
            MsgBox "Ocurrió un error en la creacion de ticket.", vbCritical, "Notificación Electrónica"
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        If SeguirSinCorreo = 2 Then
            MsgBox "Se creó el ticket: " & vTicket & Chr(13) & "No continuó con el flujo de envío electrónico.", vbInformation, "Notificación Electrónica"
            Screen.MousePointer = 0
            Exit Sub
        End If
   
        'Call ConexionDavidCloud
        
        vTokenDavidCloud = TokenDavicloud
        
        Call Get_Firmantes(Txt_NumPol.Text) 'LLENA LOS FIRMANTES EN LST_FIRMANTES
        
          
        For x = 0 To UBound(LST_FIRMANTES)
        
            Call RegistrarFirmantes(vTokenDavidCloud, LST_FIRMANTES(x))
        
        Next
             
        VidDocumento = RegistraDocumento(vTokenDavidCloud, Txt_NumPol.Text, vTicket)
        Call IniciarProcesoFirma(VidDocumento, vTokenDavidCloud, LST_FIRMANTES)
        
        MsgBox "Se envió la notificación eléctronica. Póliza: " & Txt_NumPol.Text & " Ticket: " & vTicket, vbInformation, "Notificación Electrónica"
        Screen.MousePointer = 0
        
        Exit Sub
        
Msgerror:
        Screen.MousePointer = 0
         MsgBox "No se pudo enviar la notificación eléctronica. Póliza: " & Txt_NumPol.Text & " Error: " & Err.Description, vbCritical, "Notificación Electrónica"
        
   
End Sub

Private Function ObtenerToken() As String
          
        Dim rest As New ChilkatRest
        Dim success As Long
        Dim vToken As String
        
        vToken = ""
        
        Dim bTls As Long
        bTls = 1
        Dim port As Long
        port = 443
        Dim bAutoReconnect As Long
        
        Result = True
        bAutoReconnect = 1
        success = rest.Connect("nntp-user-pool-stg.auth.us-east-1.amazoncognito.com", port, bTls, bAutoReconnect)
        If (success <> 1) Then
            Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
            Debug.Print rest.LastErrorText
            Exit Function
        End If
        
        success = rest.AddQueryParam("grant_type", "client_credentials")
        success = rest.AddQueryParam("scope", "https://api-stg.soporte.protectasecuritycloud.pe/jira_basic")
        
        success = rest.AddHeader("Cookie", "XSRF-TOKEN=e2e403f1-6fd6-4e84-8027-0424a4c168d9")
        success = rest.AddHeader("Authorization", "Basic NWpxNjgyN2lkM2ZxaGdraW9yMmNyZTc0cnY6cXRmazhnanU0bHFlbmIwaTNhbm5vbnQxMWdmbjFwcG51aGV2dG5uY29uc2gzbnQwMzB0")
        
        Dim strResponseBody As String
        Dim Rpta() As String
  
        strResponseBody = rest.FullRequestFormUrlEncoded("POST", "/oauth2/token")
        
        Dim p As Object
        Set p = json.parse(strResponseBody)
        vToken = p.Item("access_token")
        
        If (rest.LastMethodSuccess <> 1) Then
            Debug.Print rest.LastErrorText
            ObtenerToken = vToken
            Exit Function
        End If
        
        Dim respStatusCode As Long
        respStatusCode = rest.ResponseStatusCode
        Debug.Print "response status code = " & respStatusCode
        
        If (respStatusCode >= 400) Then
           SP_LOG_API_TOKEN_JIRA vgUsuario, dPol.Num_Cot, "nntp-user-pool-stg.auth.us-east-1.amazoncognito.com/oauth2/token", "1", p.Item("error"), v_id_transac
'            Debug.Print "Response Status Code = " & respStatusCode
'            Debug.Print "Response Header:"
'            Debug.Print rest.ResponseHeader
'            Debug.Print "Response Body:"
'            MsgBox strResponseBody
             Result = False

            ObtenerToken = vToken
            Exit Function
        Else
            ObtenerToken = vToken
            SP_LOG_API_TOKEN_JIRA vgUsuario, dPol.Num_Cot, "nntp-user-pool-stg.auth.us-east-1.amazoncognito.com/oauth2/token", "0", "Obtencion Token", v_id_transac
        End If
   
End Function
Public Function CreaTicket(ByVal vToken As String, ByVal Flujo As Byte) As String
            
            Dim rest As New ChilkatRest
            Dim success As Long
            Dim sTicket As String
            Dim vTipoFirma As String
            Dim oDatosAgentes As AgentesCom
            
            oDatosAgentes = getAgentesComerciales(Txt_NumPol.Text)
            

            sTicket = ""
            
            'URL: https://api-stg.soporte.protectasecuritycloud.pe/jira/v1/issues
            Dim bTls As Long
            bTls = 1
            Dim port As Long
            port = 443
            Dim bAutoReconnect As Long
            bAutoReconnect = 1
            success = rest.Connect("api-stg.soporte.protectasecuritycloud.pe", port, bTls, bAutoReconnect)
            If (success <> 1) Then
                MsgBox "ConnectFailReason: " & rest.ConnectFailReason
                MsgBox rest.LastErrorText
                Exit Function
            End If
            
            
            
            Dim jsonChilka As New ChilkatJsonObject
            success = jsonChilka.UpdateString("system", "cliente360")
            success = jsonChilka.UpdateString("fields[0].id", "project")
            success = jsonChilka.UpdateString("fields[0].value.key", "ERV")
            success = jsonChilka.UpdateString("fields[1].id", "issuetype")
            success = jsonChilka.UpdateString("fields[1].value.id", "11400")
            success = jsonChilka.UpdateString("fields[2].id", "summary")
            success = jsonChilka.UpdateString("fields[2].value", dPol.Summary)
            success = jsonChilka.UpdateString("fields[3].id", "description")
            success = jsonChilka.UpdateString("fields[3].value", dPol.Summary)
            
            'Tipo de Renta
            success = jsonChilka.UpdateString("fields[4].id", "customfield_13221")
            success = jsonChilka.UpdateString("fields[4].value.id", "15415")
            
            'Nombre de la AFP
            success = jsonChilka.UpdateString("fields[5].id", "customfield_13242")
            success = jsonChilka.UpdateString("fields[5].value", dPol.Nombre_Afp)
            
            'CUSPP
            success = jsonChilka.UpdateString("fields[6].id", "customfield_13244")
            success = jsonChilka.UpdateString("fields[6].value", dPol.Cod_Cuspp)
            
            'NOMBRE DEL CLIENTE
            success = jsonChilka.UpdateString("fields[7].id", "customfield_12215")
            success = jsonChilka.UpdateString("fields[7].value", dPol.nombres)
            
            'TIPO DE DOCUMENTO DEL CLIENTE
            success = jsonChilka.UpdateString("fields[8].id", "customfield_12213")
            success = jsonChilka.UpdateString("fields[8].value.id", dPol.TIPO_DOC)
            
            'NUMERO DE DOCUMENTO DEL CLIENTE
            success = jsonChilka.UpdateString("fields[9].id", "customfield_12214")
            success = jsonChilka.UpdateString("fields[9].value", dPol.Num_IdenBen)
            
            'EMAIL DEL CLIENTE
            success = jsonChilka.UpdateString("fields[10].id", "customfield_12310")
            success = jsonChilka.UpdateString("fields[10].value", dPol.GLS_CORREO)
            
            'TELEFONO DEL CLIENTE
            success = jsonChilka.UpdateString("fields[11].id", "customfield_13238")
            success = jsonChilka.UpdateString("fields[11].value", dPol.GLS_FONO)
            
            'MONEDA RRVV
            success = jsonChilka.UpdateString("fields[12].id", "customfield_13249")
            success = jsonChilka.UpdateString("fields[12].value.id", dPol.MONEDA_RRVV)
            
            'PRIMA
            success = jsonChilka.UpdateString("fields[13].id", "customfield_11725")
            success = jsonChilka.UpdateString("fields[13].value", dPol.MTO_PRIUNI)
            
                 
            'MODALIDAD RENTA
            success = jsonChilka.UpdateString("fields[14].id", "customfield_13239")
            success = jsonChilka.UpdateString("fields[14].value.id", dPol.MODALIDAD_RENTA)
            
            'TIPO DE PRESTACION
            success = jsonChilka.UpdateString("fields[15].id", "customfield_13243")
            success = jsonChilka.UpdateString("fields[15].value.id", dPol.TIPO_PRESTACION)
            
            'TIPO PENSION
            success = jsonChilka.UpdateString("fields[16].id", "customfield_13252")
            success = jsonChilka.UpdateString("fields[16].value.id", dPol.VAL_TIPO_PENSION)
            
            'Indicador Tipo Firma
            If Flujo = 2 Then
               vTipoFirma = 15424
            Else
               vTipoFirma = 15423
            End If
  
             success = jsonChilka.UpdateString("fields[17].id", "customfield_13225")
             success = jsonChilka.UpdateString("fields[17].value.id", vTipoFirma)
            
            
             success = jsonChilka.UpdateString("fields[18].id", "customfield_13432")
             success = jsonChilka.UpdateString("fields[18].value", oDatosAgentes.NombreAsesor)
             
             
             success = jsonChilka.UpdateString("fields[19].id", "customfield_13433")
             success = jsonChilka.UpdateString("fields[19].value", oDatosAgentes.MailAsesor)
             
             
             success = jsonChilka.UpdateString("fields[20].id", "customfield_13434")
             success = jsonChilka.UpdateString("fields[20].value", oDatosAgentes.NombreSupervisor)
             
             
             success = jsonChilka.UpdateString("fields[21].id", "customfield_13435")
             success = jsonChilka.UpdateString("fields[21].value", oDatosAgentes.MailSupervisor)
             
             success = jsonChilka.UpdateString("fields[22].id", "customfield_11784")
             success = jsonChilka.UpdateString("fields[22].value", Txt_NumPol.Text)
           
            '1. Nombre del Asesor: customfield_13432
            '2. Email del Asesor: customfield_13433
            '3. Nombre del Supervisor: customfield_13434
            '4. Email del Supervisor: customfield_13435

            If dPol.NombreRep <> "" Then
            
            
                'Indicador Representante
                success = jsonChilka.UpdateString("fields[23].id", "customfield_13303")
                success = jsonChilka.UpdateString("fields[23].value.id", "15639")
    
                'Nombre
                success = jsonChilka.UpdateString("fields[24].id", "customfield_13307")
                success = jsonChilka.UpdateString("fields[24].value", dPol.NombreRep & " " & dPol.ApeRep)
    
                'Email
                success = jsonChilka.UpdateString("fields[25].id", "customfield_13308")
                success = jsonChilka.UpdateString("fields[25].value", dPol.gls_correorep) ' "rtarazona@protectasecurity.pe")
    
                'Telefono
                success = jsonChilka.UpdateString("fields[26].id", "customfield_13340")
                success = jsonChilka.UpdateString("fields[26].value", dPol.celularRep)
                
            Else
                       
                'Indicador Representante
                success = jsonChilka.UpdateString("fields[23].id", "customfield_13303")
                success = jsonChilka.UpdateString("fields[23].value.id", "15638")
       
            
            End If
            
        
           
            success = rest.AddHeader("Authorization", "Bearer " & vToken)
            success = rest.AddHeader("Content-Type", "application/json")
            
            Dim sbRequestBody As New ChilkatStringBuilder
            success = jsonChilka.EmitSb(sbRequestBody)
            Dim sbResponseBody As New ChilkatStringBuilder
            success = rest.FullRequestSb("POST", "/jira/v1/issues", sbRequestBody, sbResponseBody)
            If (success <> 1) Then
                MsgBox rest.LastErrorText
                Exit Function
            End If
            
            Dim respStatusCode As Long
            respStatusCode = rest.ResponseStatusCode
            Debug.Print "response status code = " & respStatusCode
            
             
            Dim p As Object
            Dim CodReturn As String
            Dim RspMensaje As String
            Set p = json.parse(sbResponseBody.GetAsString)
            sTicket = p.Item("id")
               
            CreaTicket = sTicket
               
               
           
            If (respStatusCode >= 400) Then
                SP_LOG_API_TIKET_JIRA vgUsuario, v_id_transac, p.Item("id"), "api-stg.soporte.protectasecuritycloud.pe/jira/v1/issues", "1", sbResponseBody.GetAsString
'                Debug.Print "Response Status Code = " & respStatusCode
'                Debug.Print "Response Header:"
'                Debug.Print rest.ResponseHeader
'                Debug.Print "Response Body:"
                MsgBox sbResponseBody.GetAsString()
                CreaTicket = "error"
                Exit Function
            Else
            
                SP_LOG_API_TIKET_JIRA vgUsuario, v_id_transac, p.Item("id"), "api-stg.soporte.protectasecuritycloud.pe/jira/v1/issues", "0", "Ticket Creado correctamente"
              
            End If

End Function
Public Sub ConexionDavidCloud()
Dim rest As New ChilkatRest
Dim success As Long

'  URL: https://test36.davicloud.com/API/sign/v1/api_rest.php/servicio
Dim bTls As Long
bTls = 1
Dim port As Long
port = 443
Dim bAutoReconnect As Long
bAutoReconnect = 1
'PRUEBAS
'success = rest.Connect("test36.davicloud.com", port, bTls, bAutoReconnect)
'PROD
success = rest.Connect("protecta.davicloud.com", port, bTls, bAutoReconnect)
If (success <> 1) Then
    Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
    Debug.Print rest.LastErrorText
    Exit Sub
End If


Dim strResponseBody As String
strResponseBody = rest.FullRequestFormUrlEncoded("GET", "/API/sign/v1/api_rest.php/servicio")

If (rest.LastMethodSuccess <> 1) Then
    Debug.Print rest.LastErrorText
    Exit Sub
End If

Dim respStatusCode As Long
respStatusCode = rest.ResponseStatusCode
Debug.Print "response status code = " & respStatusCode
If (respStatusCode >= 400) Then
    Debug.Print "Response Status Code = " & respStatusCode
    Debug.Print "Response Header:"
    Debug.Print rest.ResponseHeader
    Debug.Print "Response Body:"
    MsgBox strResponseBody
    Exit Sub
End If

MsgBox strResponseBody
End Sub
Private Sub cargaDatos_poliza(ByVal pnum_poliza As String)

                Dim Mensaje As String
           
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
                
                objCmd.CommandText = "PKG_API_FIRMARRVV.sp_datos_poliza"
                objCmd.CommandType = adCmdStoredProc
                
'                pnum_poliza  varchar2,
'                pnum_orden   number,
'                p_outNumError    out number,
'                p_outMsgError    out varchar2
'                pcursor      out sys_refcursor,
                
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 10, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("pnum_orden", adDouble, adParamInput)
                param2.Value = 2
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
    
                
                If Len(Trim(Mensaje)) = 0 Then
                    
                    dPol.Num_Cot = IIf(IsNull(rs!Num_Cot), "", rs!Num_Cot)
                    dPol.Num_Poliza = IIf(IsNull(rs!Num_Poliza), "", rs!Num_Poliza)
                    dPol.Cod_AFP = IIf(IsNull(rs!Cod_AFP), "", rs!Cod_AFP)
                    dPol.Cod_Cuspp = IIf(IsNull(rs!Cod_Cuspp), "", rs!Cod_Cuspp)
                    dPol.GLS_CORREO = IIf(IsNull(rs!GLS_CORREO), "", rs!GLS_CORREO)
                    dPol.Cod_Moneda = IIf(IsNull(rs!Cod_Moneda), "", rs!Cod_Moneda)
                    dPol.MTO_PRIUNI = IIf(IsNull(rs!MTO_PRIUNI), "", rs!MTO_PRIUNI)
                    dPol.MONEDA_RRVV = IIf(IsNull(rs!MONEDA_RRVV), "", rs!MONEDA_RRVV)
                    dPol.COD_USUARIO = IIf(IsNull(rs!COD_USUARIO), "", rs!COD_USUARIO)
                    dPol.GLS_TIPOIDEN = IIf(IsNull(rs!GLS_TIPOIDEN), "", rs!GLS_TIPOIDEN)
                    dPol.TIPO_DOC = IIf(IsNull(rs!TIPO_DOC), "", rs!TIPO_DOC)
                    dPol.Num_IdenBen = IIf(IsNull(rs!Num_IdenBen), "", rs!Num_IdenBen)
                    dPol.nombres = IIf(IsNull(rs!nombres), "", rs!nombres)
                    dPol.Gls_CorreoBen = IIf(IsNull(rs!Gls_CorreoBen), "", rs!Gls_CorreoBen)
                    dPol.GLS_FONO = IIf(IsNull(rs!GLS_FONO), "", rs!GLS_FONO)
                    dPol.MODALIDAD_RENTA = IIf(IsNull(rs!MODALIDAD_RENTA), "", rs!MODALIDAD_RENTA)
                    dPol.TIPO_PRESTACION = IIf(IsNull(rs!TIPO_PRESTACION), "", rs!TIPO_PRESTACION)
                    dPol.Summary = IIf(IsNull(rs!Summary), "", rs!Summary)
                    dPol.Nombre_Afp = IIf(IsNull(rs!Nombre_Afp), "", rs!Nombre_Afp)
                    dPol.cod_tipoidenRep = IIf(IsNull(rs!cod_tipoidenRep), "", rs!cod_tipoidenRep)
                    dPol.Num_idenrep = IIf(IsNull(rs!Num_idenrep), "", rs!Num_idenrep)
                    dPol.NombreRep = IIf(IsNull(rs!Gls_NombresRep), "", rs!Gls_NombresRep)
                    dPol.ApeRep = IIf(IsNull(rs!Gls_ApepatRep), "", rs!Gls_ApepatRep) & " " & IIf(IsNull(rs!Gls_ApematRep), "", rs!Gls_ApematRep)
                    dPol.celularRep = IIf(IsNull(rs!GLS_TELREP2), "", rs!GLS_TELREP2)
                    dPol.VAL_TIPO_PENSION = IIf(IsNull(rs!VAL_TIPO_PENSION), "", rs!VAL_TIPO_PENSION)
                    dPol.TIPO_PENSION = IIf(IsNull(rs!TIPO_PENSION), "", rs!TIPO_PENSION)
                    dPol.Tipo_Renta = IIf(IsNull(rs!Tipo_Renta), "", rs!Tipo_Renta)
                    dPol.gls_correorep = IIf(IsNull(rs!gls_correorep), "", rs!gls_correorep)
         
                            
                Else
                
                    MsgBox Mensaje, vbCritical, "Error"
                
                End If
                
  conn.Close
  Set objCmd = Nothing
  Set rs = Nothing
  Set conn = Nothing
  
  
            
End Sub
Private Sub EnvioEmails_poliza(ByVal pnum_poliza As String, _
                                   ByRef pemails() As String, _
                                   ByRef MensajeError As String)

              
                Dim lstCorreos As String
                
      
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
                
                objCmd.CommandText = "PKG_API_FIRMARRVV.sp_validacion_correos"
                objCmd.CommandType = adCmdStoredProc
                
'                pnum_poliza  varchar2,
'                pemails      out varchar2,
'                p_outNumError    out number,
'                p_outMsgError    out varchar2
                
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 10, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("pemails", adVarChar, adParamOutput, 1000)
                objCmd.Parameters.Append param2
            
                 Set param3 = objCmd.CreateParameter("p_outNumError", adDouble, adParamOutput)
                objCmd.Parameters.Append param3
                
                Set param4 = objCmd.CreateParameter("p_outMsgError", adVarChar, adParamOutput, 200)
                objCmd.Parameters.Append param4
                
                Set rs = objCmd.Execute
                
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                End If
                
                If Not IsNull(objCmd.Parameters.Item("pemails").Value) Then
                    lstCorreos = objCmd.Parameters.Item("pemails").Value
                Else
                    lstCorreos = ""
                End If
                
                pemails = Split(lstCorreos, ";")
                
                          

                
  conn.Close
  Set objCmd = Nothing
  Set rs = Nothing
  Set conn = Nothing
  
  
            
End Sub

Private Sub SP_LOG_API_TOKEN_JIRA(ByVal p_usuario As String, _
                                   ByVal p_numcot As String, _
                                   ByVal p_urlapi As String, _
                                   ByVal p_error As String, _
                                   ByVal p_mensaje As String, _
                                   ByRef p_id_transac As Integer)
                                   
                
                                   
                Dim conn    As ADODB.Connection
                Set conn = New ADODB.Connection
                Set rs = New ADODB.Recordset
                Set objCmd = New ADODB.Command
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
                
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMA_DOCUMENTOS.SP_LOG_API_TOKEN_JIRA"
                objCmd.CommandType = adCmdStoredProc
                
                Set param1 = objCmd.CreateParameter("p_usuario", adVarChar, adParamInput, 10, p_usuario)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("p_numcot", adVarChar, adParamInput, 12, p_numcot)
                objCmd.Parameters.Append param2
                
                Set param3 = objCmd.CreateParameter("p_codsistema", adVarChar, adParamInput, 2, "RV")
                objCmd.Parameters.Append param3
                
                Set param4 = objCmd.CreateParameter("p_urlapi", adVarChar, adParamInput, 255, p_urlapi)
                objCmd.Parameters.Append param4
                
                Set param5 = objCmd.CreateParameter("p_error", adVarChar, adParamInput, 1, p_error)
                objCmd.Parameters.Append param5
                
                Set param6 = objCmd.CreateParameter("p_mensaje", adVarChar, adParamInput, 300, p_mensaje)
                objCmd.Parameters.Append param6
                
                Set param7 = objCmd.CreateParameter("p_id_transac", adDouble, adParamOutput)
                objCmd.Parameters.Append param7
                
                Set rs = objCmd.Execute
                
                If Not IsNull(objCmd.Parameters.Item("p_id_transac").Value) Then
                    p_id_transac = Trim(objCmd.Parameters.Item("p_id_transac").Value)
                Else
                    p_id_transac = 0
                End If
                                  
                conn.Close
                Set objCmd = Nothing
                Set rs = Nothing
                Set conn = Nothing
  
End Sub

Private Sub SP_LOG_API_TIKET_JIRA(ByVal p_usuario As String, _
                                 ByVal p_id_transac As Integer, _
                                 ByVal p_ticket As String, _
                                 ByVal p_urlapi As String, _
                                 ByVal p_error As String, _
                                 ByVal p_mensaje As String)
                                 
        Dim conn    As ADODB.Connection
        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Set objCmd = New ADODB.Command
        
        conn.Provider = "OraOLEDB.Oracle"
        conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
        conn.CursorLocation = adUseClient
        conn.Open
        
        Set objCmd = CreateObject("ADODB.Command")
        Set objCmd.ActiveConnection = conn
        
        objCmd.CommandText = "PKG_API_FIRMA_DOCUMENTOS.SP_LOG_API_TIKET_JIRA"
        objCmd.CommandType = adCmdStoredProc
'
'        p_usuario varchar2,
'                                  p_id_transac number,
'                                  p_ticket varchar2,
'                                  p_urlapi varchar2,
'                                  p_error char,
'                                  p_mensaje varchar2
'                                 ) AS
        
        Set param1 = objCmd.CreateParameter("p_usuario", adVarChar, adParamInput, 10, p_usuario)
        objCmd.Parameters.Append param1
        
        Set param2 = objCmd.CreateParameter("p_id_transac", adInteger, adParamInput)
        param2.Value = p_id_transac
        objCmd.Parameters.Append param2
        
        Set param3 = objCmd.CreateParameter("p_ticket", adVarChar, adParamInput, 10, p_ticket)
        objCmd.Parameters.Append param3
              
        Set param4 = objCmd.CreateParameter("p_urlapi", adVarChar, adParamInput, 255, p_urlapi)
        objCmd.Parameters.Append param4
        
        Set param5 = objCmd.CreateParameter("p_error", adChar, adParamInput, 1, p_error)
        objCmd.Parameters.Append param5
        
        Set param6 = objCmd.CreateParameter("p_mensaje", adVarChar, adParamInput, 300, p_mensaje)
        objCmd.Parameters.Append param6
        
        Set rs = objCmd.Execute
        
        conn.Close
        Set objCmd = Nothing
        Set rs = Nothing
        Set conn = Nothing

                                 
End Sub

Private Sub SP_LOG_API_DOC(ByVal p_usuario As String, _
                           ByVal p_id_transac As Integer, _
                           ByVal p_urlapi As String, _
                           ByVal p_error As String, _
                           ByVal p_mensaje As String)

        Dim conn    As ADODB.Connection
        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Set objCmd = New ADODB.Command
        
        conn.Provider = "OraOLEDB.Oracle"
        conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
        conn.CursorLocation = adUseClient
        conn.Open
        
        Set objCmd = CreateObject("ADODB.Command")
        Set objCmd.ActiveConnection = conn
        
        objCmd.CommandText = "PKG_API_FIRMA_DOCUMENTOS.SP_LOG_API_DOC"
        objCmd.CommandType = adCmdStoredProc
   
        Set param1 = objCmd.CreateParameter("p_usuario", adVarChar, adParamInput, 10, p_usuario)
        objCmd.Parameters.Append param1
        
        Set param2 = objCmd.CreateParameter("p_id_transac", adDouble, adParamInput)
        param2.Value = p_id_transac
        objCmd.Parameters.Append param2
        
        Set param4 = objCmd.CreateParameter("p_urlapi", adVarChar, adParamInput, 255, p_urlapi)
        objCmd.Parameters.Append param4
        
        Set param5 = objCmd.CreateParameter("p_error", adVarChar, adParamInput, 1, p_error)
        objCmd.Parameters.Append param5
        
        Set param6 = objCmd.CreateParameter("p_mensaje", adVarChar, adParamInput, 300, p_mensaje)
        objCmd.Parameters.Append param6
        
        objCmd.Execute
        
        
        
        
        conn.Close
        Set objCmd = Nothing
        Set rs = Nothing
        Set conn = Nothing
 

End Sub
Private Sub Get_Firmantes(ByVal pnum_poliza As String)
                                   
                Dim conn    As ADODB.Connection
                Set conn = New ADODB.Connection
                Set rs = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
                Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")
                
                On Error GoTo ManejoError
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
                
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMARRVV.sp_listaFirmantesPrepoliza"
                objCmd.CommandType = adCmdStoredProc
                
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 10, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("p_outNumError", adDouble, adParamOutput)
                objCmd.Parameters.Append param2
                
                Set param3 = objCmd.CreateParameter("p_outMsgError", adVarChar, adParamOutput, 200)
                objCmd.Parameters.Append param3
                
                       
                Set rs = objCmd.Execute
        
                
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                    Dim i As Integer
                    i = 0
                               
                  While Not rs.EOF()
                    i = i + 1
                    ReDim Preserve LST_FIRMANTES(i - 1)
                    
                     '  , TIPO, NUM_IDENBEN, GLS_NOMBEN, GLS_PATBEN, GLS_MATBEN, GLS_CORREOBEN, CELULAR, COD_SEXO, COD_TIPOIDENBEN, MENOREDAD
                    
                    LST_FIRMANTES(i - 1).tipo = IIf(IsNull(rs!tipo), "", rs!tipo)
                    LST_FIRMANTES(i - 1).NUM_IDEN = IIf(IsNull(rs!Num_IdenBen), "", rs!Num_IdenBen)
                    LST_FIRMANTES(i - 1).GLS_NOMBRES = IIf(IsNull(rs!GLS_NOMBRES), "", rs!GLS_NOMBRES)
                    LST_FIRMANTES(i - 1).GLS_APEPAT = IIf(IsNull(rs!Gls_PatBen), "", rs!Gls_PatBen)
                    LST_FIRMANTES(i - 1).GLS_APEMAT = IIf(IsNull(rs!Gls_MatBen), "", rs!Gls_MatBen)
                    LST_FIRMANTES(i - 1).GLS_CORREO = IIf(IsNull(rs!Gls_CorreoBen), "", rs!Gls_CorreoBen)
                    LST_FIRMANTES(i - 1).celular = IIf(IsNull(rs!celular), "", rs!celular)
                    LST_FIRMANTES(i - 1).Direccion = " "
                    LST_FIRMANTES(i - 1).genero = IIf(IsNull(rs!genero), "", rs!genero)
                    LST_FIRMANTES(i - 1).TIPO_DOCUMENTO = IIf(IsNull(rs!Cod_TipoIdenBen), "", rs!Cod_TipoIdenBen)
                    LST_FIRMANTES(i - 1).MENORDEEDAD = IIf(IsNull(rs!MENOREDAD), "", rs!MENOREDAD)
                    LST_FIRMANTES(i - 1).Direccion = IIf(IsNull(rs!Direccion), "", rs!Direccion)
                    LST_FIRMANTES(i - 1).departamento = IIf(IsNull(rs!departamento), "", rs!departamento)
                    LST_FIRMANTES(i - 1).provincia = IIf(IsNull(rs!provincia), "", rs!provincia)
                    LST_FIRMANTES(i - 1).distrito = IIf(IsNull(rs!distrito), "", rs!distrito)
                    LST_FIRMANTES(i - 1).PARENTESCO = IIf(IsNull(rs!PARENTESCO), "", rs!PARENTESCO)
              
          
                    Select Case LST_FIRMANTES(i - 1).tipo
                        Case "CONT"
                                  LST_FIRMANTES(i - 1).FIRMA_TIPO = "A"
                                 
                                  
                                  If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
                                    LST_FIRMANTES(i - 1).TIPO_FIRMA = "NF"
                                   Else
                                    LST_FIRMANTES(i - 1).TIPO_FIRMA = "FE"
                                  End If
            
                                  
                        Case "REP"
                                  LST_FIRMANTES(i - 1).FIRMA_TIPO = "R"
                                  LST_FIRMANTES(i - 1).PARENTESCO = "OTROS"
                                  
                                  If Left(Trim(Lbl_TipPen.Caption), 2) = "08" Then
                                    LST_FIRMANTES(i - 1).TIPO_FIRMA = "FE"
                                   Else
                                    LST_FIRMANTES(i - 1).TIPO_FIRMA = "NF"
                                  End If
                                  
                        Case "BEN"
                                  LST_FIRMANTES(i - 1).FIRMA_TIPO = "B"
                                  LST_FIRMANTES(i - 1).TIPO_FIRMA = "NF"
                                  
'********************No se envian estos datos**************************
'                        Case "TUT"
'                                  LST_FIRMANTES(i - 1).FIRMA_TIPO = ""
'                                  LST_FIRMANTES(i - 1).TIPO_FIRMA = "NF"
'                        Case "ASES"
'                                  LST_FIRMANTES(i - 1).FIRMA_TIPO = "A"
'
'                                  LST_FIRMANTES(i - 1).TIPO_FIRMA = "NF"
'                        Case "SUPERVISOR"
'                                  LST_FIRMANTES(i - 1).FIRMA_TIPO = "A"
                                
                    
                    End Select
                    
                
                    
                    rs.MoveNext
                    
    
                  Wend
                  
                  
                  
               
                End If
                
                conn.Close
                Set objCmd = Nothing
                Set rs = Nothing
                Set conn = Nothing
                Exit Sub
                
                
ManejoError:
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                End If
                MsgBox Err.Description + MensajeError, vbCritical
                        
End Sub

Private Function TokenDavicloud() As String


Dim rest As New ChilkatRest
Dim success As Long

' URL: https://test36.davicloud.com/API/sign/v1/api_rest.php/access_token
Dim bTls As Long
bTls = 1
Dim port As Long
port = 443
Dim bAutoReconnect As Long
bAutoReconnect = 1
'PRUEBAS
'success = rest.Connect("test36.davicloud.com", port, bTls, bAutoReconnect)
'PROD
success = rest.Connect("protecta.davicloud.com", port, bTls, bAutoReconnect)
If (success <> 1) Then
    Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
    Debug.Print rest.LastErrorText
    Exit Function
End If

success = rest.AddHeader("Id-Organizacion", "PROTECTA")
'success = rest.AddHeader("Authorization", "Basic QVBJUFJPVEVDVEE6UHJvdGVjdGEuMjAyMSM=")
success = rest.AddHeader("Authorization", "Basic QVBJUFJPVEVDVEE6UHJvdGVjdGFQcmQuMjAyMSM=")
'success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "bca9b2ead50d08bbc075e7ba661c6d3f")
success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "c12b12d9d8b815adc2733e72aee4540b")

Dim sbResponseBody As New ChilkatStringBuilder
success = rest.FullRequestNoBodySb("POST", "/API/sign/v1/api_rest.php/access_token", sbResponseBody)
If (success <> 1) Then
    Debug.Print rest.LastErrorText
    Exit Function
End If

Dim respStatusCode As Long
respStatusCode = rest.ResponseStatusCode
Debug.Print "response status code = " & respStatusCode
If (respStatusCode >= 400) Then
       SP_LOG_API_DOC vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/access_token", "1", sbResponseBody.GetAsString
'    Debug.Print "Response Status Code = " & respStatusCode
'    Debug.Print "Response Header:"
'    Debug.Print rest.ResponseHeader
'    Debug.Print "Response Body:"
'    MsgBox sbResponseBody.GetAsString()
    Exit Function
Else
        SP_LOG_API_DOC vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/access_token", "0", "Tocken davidCloud obtenido correctamente"
End If

Dim Rpta() As String
Dim vTokenDavidCloud As String

Dim p As Object
Set p = json.parse(sbResponseBody.GetAsString())
vTokenDavidCloud = p.Item("access_token")
        

TokenDavicloud = vTokenDavidCloud


End Function

Private Sub RegistrarFirmantes(ByVal pToken As String, ByRef vFirmante As Firmantes)
    
        Dim rest As New ChilkatRest
        Dim success As Long
        
        ' URL: https://test36.davicloud.com/API/sign/v1/api_rest.php/registra_firmante
        Dim bTls As Long
        bTls = 1
        Dim port As Long
        port = 443
        Dim bAutoReconnect As Long
        bAutoReconnect = 1
        'PRUEBAS
        'success = rest.Connect("test36.davicloud.com", port, bTls, bAutoReconnect)
        'PROD
        success = rest.Connect("protecta.davicloud.com", port, bTls, bAutoReconnect)
        If (success <> 1) Then
            Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
            Debug.Print rest.LastErrorText
            Exit Sub
        End If
      
        
        Dim JSONchk As New ChilkatJsonObject
        success = JSONchk.UpdateString("nacionalidad", "PE")
        success = JSONchk.UpdateString("tipodocumento", "1")
        success = JSONchk.UpdateString("numerodocumento", vFirmante.NUM_IDEN)
        success = JSONchk.UpdateString("nombre", vFirmante.GLS_NOMBRES)
        success = JSONchk.UpdateString("apellidos", vFirmante.GLS_APEPAT & " " & vFirmante.GLS_APEMAT)
        success = JSONchk.UpdateString("correo", vFirmante.GLS_CORREO) ' "giovanni.cruz@efit11ec.pe")
        success = JSONchk.UpdateString("celular", vFirmante.celular)
        success = JSONchk.UpdateString("direccion", vFirmante.Direccion)
        success = JSONchk.UpdateString("departamento", vFirmante.departamento)
        success = JSONchk.UpdateString("provincia", vFirmante.provincia)
        success = JSONchk.UpdateString("distrito", vFirmante.distrito)
        success = JSONchk.UpdateString("genero", vFirmante.genero)
        
        success = rest.AddHeader("Id-Organizacion", "PROTECTA")
        success = rest.AddHeader("Content-Type", "application/json")
        success = rest.AddHeader("Authorization", "Bearer " & pToken)
        'success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "bca9b2ead50d08bbc075e7ba661c6d3f")
        success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "c12b12d9d8b815adc2733e72aee4540b")
        
        Dim sbRequestBody As New ChilkatStringBuilder
        success = JSONchk.EmitSb(sbRequestBody)
        
        Dim sbResponseBody As New ChilkatStringBuilder
        
        success = rest.FullRequestSb("POST", "/API/sign/v1/api_rest.php/registra_firmante", sbRequestBody, sbResponseBody)
        
        If (success <> 1) Then
            Debug.Print rest.LastErrorText
            Exit Sub
        End If
        
        Dim p As Object
        Dim vID_FIRMANTE As String
             
        Set p = json.parse(sbResponseBody.GetAsString)
        vFirmante.ID_FIRMANTE = p.Item("idfirmante")
   
   
        Dim respStatusCode As Long
        respStatusCode = rest.ResponseStatusCode
        Debug.Print "response status code = " & respStatusCode
        
        If (respStatusCode >= 400) Then
              SP_LOG_API_DOC vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/registra_firmante", "1", sbResponseBody.GetAsString
            Debug.Print "Response Status Code = " & respStatusCode
            Debug.Print "Response Header:"
            Debug.Print rest.ResponseHeader
            Debug.Print "Response Body:"
            Debug.Print sbResponseBody.GetAsString()
            Exit Sub
        Else
            SP_LOG_API_DOC vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/registra_firmante", "0", p.Item("message") + " - " + vFirmante.tipo + " - " + vFirmante.NUM_IDEN + " - " + vFirmante.GLS_NOMBRES + " " + vFirmante.GLS_APEPAT + " - " + vFirmante.GLS_CORREO
        End If
        
  

End Sub
Private Function RegistraDocumento(ByVal vToken As String, ByVal vPoliza As String, ByVal vTicketDavidCloud As String) As String

        Dim rest As New ChilkatRest
        Dim success As Long
        
        ' URL: https://test36.davicloud.com/API/sign/v1/api_rest.php/registra_documento_vitalicia
        Dim bTls As Long
        bTls = 1
        Dim port As Long
        port = 443
        Dim bAutoReconnect As Long
        bAutoReconnect = 1
        'PRUEBAS
        'success = rest.Connect("test36.davicloud.com", port, bTls, bAutoReconnect)
        'PROD
        success = rest.Connect("protecta.davicloud.com", port, bTls, bAutoReconnect)
        If (success <> 1) Then
            Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
            Debug.Print rest.LastErrorText
            Exit Function
        End If
        
        Dim JSONchk As New ChilkatJsonObject
        success = JSONchk.UpdateString("documentotipo", "18")
        success = JSONchk.UpdateString("documentocreadorfirmacorreo", "juan.davila@bigdavi.com")
        success = JSONchk.UpdateString("documentodescripcion", dPol.Summary)
        success = JSONchk.UpdateString("documentonombre", "Nombre archivo.pdf")
        success = JSONchk.UpdateString("documentoclavepdf", "")
        success = JSONchk.UpdateString("documentocodigoqr", "0")
        success = JSONchk.UpdateString("documentonumticket", vTicketDavidCloud)
        success = JSONchk.UpdateString("documentonumcotizacion", vPoliza)
        success = JSONchk.UpdateString("documentoformatoact", dPol.Tipo_Renta) 'invalidez, sobrevivencia
        success = JSONchk.UpdateString("documentoproducto", "RV")
        success = JSONchk.UpdateString("documentocontent", "")
        
        success = rest.AddHeader("Id-Organizacion", "PROTECTA")
        success = rest.AddHeader("Content-Type", "application/json")
        success = rest.AddHeader("Authorization", "Bearer " & vToken)
        'success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "bca9b2ead50d08bbc075e7ba661c6d3f")
        success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "c12b12d9d8b815adc2733e72aee4540b")
        
        Dim sbRequestBody As New ChilkatStringBuilder
        success = JSONchk.EmitSb(sbRequestBody)
        
        Dim sbResponseBody As New ChilkatStringBuilder
        success = rest.FullRequestSb("POST", "/API/sign/v1/api_rest.php/registra_documento_vitalicia", sbRequestBody, sbResponseBody)
        
        Dim VidDocumento As String
        Dim p As Object

        Set p = json.parse(sbResponseBody.GetAsString)
        VidDocumento = p.Item("iddocumento")
       
        RegistraDocumento = VidDocumento
        
        If (success <> 1) Then
            Debug.Print rest.LastErrorText
            Exit Function
        End If
        
        Dim respStatusCode As Long
        respStatusCode = rest.ResponseStatusCode
         
         If (respStatusCode >= 400) Then
          SP_LOG_API_DOC vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/registra_documento_vitalicia", "1", sbResponseBody.GetAsString
'            Debug.Print "Response Status Code = " & respStatusCode
'            Debug.Print "Response Header:"
'            Debug.Print rest.ResponseHeader
'            Debug.Print "Response Body:"
'            Debug.Print sbResponseBody.GetAsString()
            Exit Function
        Else
            SP_LOG_API_DOC vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/registra_documento_vitalicia", "0", p.Item("message")
        End If
        
 
        
        End Function
Private Sub IniciarProcesoFirma(ByVal VidDocumento As String, ByVal vToken As String, ByRef lstFirmantes() As Firmantes)
        
        Dim rest As New ChilkatRest
        Dim success As Long
        
        ' URL: https://test36.davicloud.com/API/sign/v1/api_rest.php/proceso_firma_vitalicia
        Dim bTls As Long
        bTls = 1
        Dim port As Long
        port = 443
        Dim bAutoReconnect As Long
        bAutoReconnect = 1
        'PRUEBAS
        'success = rest.Connect("test36.davicloud.com", port, bTls, bAutoReconnect)
        'PROD
        success = rest.Connect("protecta.davicloud.com", port, bTls, bAutoReconnect)
        If (success <> 1) Then
            Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
            Debug.Print rest.LastErrorText
            Exit Sub
        End If
        
        
        Dim JSONchk As New ChilkatJsonObject
        success = JSONchk.UpdateString("iddocumento", VidDocumento)
        Dim NumFirmante As String
        
        For x = 0 To UBound(lstFirmantes)

            NumFirmante = "firmantes.firmante" & x + 1
            
            success = JSONchk.UpdateString(NumFirmante & ".idfirmante", lstFirmantes(x).ID_FIRMANTE)
            success = JSONchk.UpdateString(NumFirmante & ".tipofirma", lstFirmantes(x).TIPO_FIRMA)
            success = JSONchk.UpdateString(NumFirmante & ".firmatipo", lstFirmantes(x).FIRMA_TIPO)
            success = JSONchk.UpdateString(NumFirmante & ".parentesco", lstFirmantes(x).PARENTESCO)
            success = JSONchk.UpdateString(NumFirmante & ".menordeedad", lstFirmantes(x).MENORDEEDAD)

        Next
        
        success = rest.AddHeader("Id-Organizacion", "PROTECTA")
        success = rest.AddHeader("Content-Type", "application/json")
        success = rest.AddHeader("Authorization", "Bearer " & vToken)
        'success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "bca9b2ead50d08bbc075e7ba661c6d3f")
        success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "c12b12d9d8b815adc2733e72aee4540b")
        
        Dim sbRequestBody As New ChilkatStringBuilder
        success = JSONchk.EmitSb(sbRequestBody)
        
        Dim sbResponseBody As New ChilkatStringBuilder
        success = rest.FullRequestSb("POST", "/API/sign/v1/api_rest.php/proceso_firma_vitalicia", sbRequestBody, sbResponseBody)
        
        If (success <> 1) Then
            Debug.Print rest.LastErrorText
            Exit Sub
        End If
        
        Dim respStatusCode As Long
        respStatusCode = rest.ResponseStatusCode
        
        Dim p As Object
        Set p = json.parse(sbResponseBody.GetAsString)
       
        
        'Debug.Print "response status code = " & respStatusCode
        If (respStatusCode >= 400) Then
        SP_LOG_API_DOC vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/proceso_firma_vitalicia", "1", sbResponseBody.GetAsString
'            Debug.Print "Response Status Code = " & respStatusCode
'            Debug.Print "Response Header:"
'            Debug.Print rest.ResponseHeader
'            Debug.Print "Response Body:"
'            Debug.Print sbResponseBody.GetAsString()
            Exit Sub
        Else
            SP_LOG_API_DOC vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/proceso_firma_vitalicia", "0", p.Item("message")
        End If
End Sub

Private Sub GrabarDireccionRepresentante(ByVal pnum_poliza As String, ByVal pcod_tipoidenrep As Integer, _
                                        ByVal pnum_idenrep As String, ByVal pfec_ingreso As String, _
                                        ByVal pcod_dire_via As String, ByVal pgls_direccion As String, _
                                        ByVal pnum_direccion As String, ByVal pcod_blockchalet As String, _
                                        ByVal pgls_blockchalet As String, ByVal pcod_interior As String, _
                                        ByVal pnum_interior As String, ByVal pcod_cjht As String, _
                                        ByVal pgls_nom_cjht As String, ByVal pgls_etapa As String, _
                                        ByVal pgls_manzana As String, ByVal pgls_lote As String, _
                                        ByVal pgls_referencia As String, ByVal pcod_pais As Integer, _
                                        ByVal pcod_departamento As Integer, ByVal pcod_provincia As Integer, _
                                        ByVal pcod_distrito As Integer, ByVal pgls_desdirebusq As String, _
                                        ByVal pcod_usuariocrea As String, ByVal pfec_crea As String, _
                                        ByVal phor_crea As String, ByVal pcod_usuariocrea_vt As String, ByVal pcod_direccion As Integer)
                                        
                                        
                Dim objCmd As ADODB.Command
                Dim rs As ADODB.Recordset
                Dim conn As ADODB.Connection
                    
                Set rs = New ADODB.Recordset
                Set conn = New ADODB.Connection
                
                Dim Texto As String
                
                
                Set conn = New ADODB.Connection
                Set rs = New ADODB.Recordset
                Set objCmd = New ADODB.Command
                
                On Error GoTo ManejoError
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
          
                               
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMARRVV.sp_insDireccionRepresentante"
                objCmd.CommandType = adCmdStoredProc
                
              
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 15, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("pcod_tipoidenrep", adNumeric, adParamInput, 3, pcod_tipoidenrep)
                objCmd.Parameters.Append param2
                
                Set param3 = objCmd.CreateParameter("pnum_idenrep", adVarChar, adParamInput, 16, pnum_idenrep)
                objCmd.Parameters.Append param3
                
                Set param4 = objCmd.CreateParameter("pfec_ingreso", adVarChar, adParamInput, 8, pfec_ingreso)
                objCmd.Parameters.Append param4
                
                Set param5 = objCmd.CreateParameter("pcod_dire_via", adVarChar, adParamInput, 2, pcod_dire_via)
                objCmd.Parameters.Append param5
                    
                Set param6 = objCmd.CreateParameter("pgls_direccion", adVarChar, adParamInput, 60, pgls_direccion)
                objCmd.Parameters.Append param6
                
                Set param7 = objCmd.CreateParameter("pnum_direccion", adVarChar, adParamInput, 4, pnum_direccion)
                objCmd.Parameters.Append param7
                
                Set param8 = objCmd.CreateParameter("pcod_blockchalet", adChar, adParamInput, 2, pcod_blockchalet)
                objCmd.Parameters.Append param8
                
                Set param9 = objCmd.CreateParameter("pgls_blockchalet", adVarChar, adParamInput, 3, pgls_blockchalet)
                objCmd.Parameters.Append param9
                
                Set param10 = objCmd.CreateParameter("pcod_interior", adChar, adParamInput, 2, pcod_interior)
                objCmd.Parameters.Append param10
                
                Set param11 = objCmd.CreateParameter("pnum_interior", adVarChar, adParamInput, 8, pnum_interior)
                objCmd.Parameters.Append param11
                
                Set param12 = objCmd.CreateParameter("pcod_cjht", adChar, adParamInput, 2, pcod_cjht)
                objCmd.Parameters.Append param12
                
                Set param13 = objCmd.CreateParameter("pgls_nom_cjht", adVarChar, adParamInput, 30, pgls_nom_cjht)
                objCmd.Parameters.Append param13
                
                
                Set param14 = objCmd.CreateParameter("pgls_etapa", adVarChar, adParamInput, 4, pgls_etapa)
                objCmd.Parameters.Append param14
                
                
                Set param15 = objCmd.CreateParameter("pgls_manzana", adVarChar, adParamInput, 4, pgls_manzana)
                objCmd.Parameters.Append param15
                

                Set param16 = objCmd.CreateParameter("pgls_lote", adVarChar, adParamInput, 4, pgls_lote)
                objCmd.Parameters.Append param16
                
                Set param17 = objCmd.CreateParameter("pgls_referencia", adVarChar, adParamInput, 30, pgls_referencia)
                objCmd.Parameters.Append param17

                Set param18 = objCmd.CreateParameter("pcod_pais", adNumeric, adParamInput, 5, pcod_pais)
                objCmd.Parameters.Append param18
       
                Set param22 = objCmd.CreateParameter("pgls_desdirebusq", adVarChar, adParamInput, 200, pgls_desdirebusq)
                objCmd.Parameters.Append param22

                Set param23 = objCmd.CreateParameter("pcod_usuariocrea", adVarChar, adParamInput, 10, pcod_usuariocrea)
                objCmd.Parameters.Append param23
                
                
                Set param24 = objCmd.CreateParameter("pcod_direccion", adNumeric, adParamInput, 4, pcod_direccion)
                objCmd.Parameters.Append param24
                
                Set param25 = objCmd.CreateParameter("pfec_crea", adVarChar, adParamInput, 8, pfec_crea)
                objCmd.Parameters.Append param25

                Set param26 = objCmd.CreateParameter("phor_crea", adVarChar, adParamInput, 6, phor_crea)
                objCmd.Parameters.Append param26
                
                Set param27 = objCmd.CreateParameter("p_outNumError", adNumeric, adParamOutput, 5, p_outNumError)
                objCmd.Parameters.Append param27
                
                Set param28 = objCmd.CreateParameter("p_outMsgError", adNumeric, adParamOutput, 200, p_outMsgError)
                objCmd.Parameters.Append param28
        
                objCmd.Execute
                
          
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                End If
                
                   
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        
        
        
        Exit Sub
        
ManejoError:
                conn.Close
                Set rs = Nothing
                Set conn = Nothing
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                End If
                MsgBox Err.Description + MensajeError, vbCritical
    

End Sub


Private Sub ObtieneDireccionRepresentante(ByVal pnum_poliza As String)
                
                
                Dim objCmd As ADODB.Command
                Dim rs As ADODB.Recordset
                Dim conn As ADODB.Connection
                    
                Set rs = New ADODB.Recordset
                Set conn = New ADODB.Connection
                
                On Error GoTo ManejoError
           
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
                
                objCmd.CommandText = "PKG_API_FIRMARRVV.sp_GetDireccionRepresentante"
                objCmd.CommandType = adCmdStoredProc
                
                    
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 15, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("p_outNumError", adNumeric, adParamOutput, 5, p_outNumError)
                objCmd.Parameters.Append param2
                
                Set param3 = objCmd.CreateParameter("p_outMsgError", adNumeric, adParamOutput, 200, p_outMsgError)
                objCmd.Parameters.Append param3
                
            
                Set rs = objCmd.Execute
                
                While Not rs.EOF()
                    
                    With DirRep
                    
                       .vTipoVia = IIf(IsNull(rs!cod_dire_via), "", rs!cod_dire_via)
                        .vDireccion = IIf(IsNull(rs!Gls_Direccion), "", rs!Gls_Direccion)
                        .vNumero = IIf(IsNull(rs!num_direccion), "", rs!num_direccion)
                        .vTipoPref = IIf(IsNull(rs!cod_interior), "", rs!cod_interior)
                        .vInterior = IIf(IsNull(rs!num_interior), "", rs!num_interior)
                        .vManzana = IIf(IsNull(rs!gls_manzana), "", rs!gls_manzana)
                        .vLote = IIf(IsNull(rs!gls_lote), "", rs!gls_lote)
                        .vEtapa = IIf(IsNull(rs!gls_etapa), "", rs!gls_etapa)
                        .vTipoConj = IIf(IsNull(rs!cod_cjht), "", rs!cod_cjht)
                        .vConjHabit = IIf(IsNull(rs!gls_nom_cjht), "", rs!gls_nom_cjht)
                        .vTipoBlock = IIf(IsNull(rs!cod_blockchalet), "", rs!cod_blockchalet)
                        .vNumBlock = IIf(IsNull(rs!gls_blockchalet), "", rs!gls_blockchalet)
                        .vReferencia = IIf(IsNull(rs!gls_referencia), "", rs!gls_referencia)
                        .vCodDireccion = IIf(IsNull(rs!Cod_Direccion), "", rs!Cod_Direccion)
                  
                    End With

                        
                     rs.MoveNext
                
                Wend
                
                
                   
         conn.Close
        Set rs = Nothing
        Set conn = Nothing

        Exit Sub
    
ManejoError:
                conn.Close
                Set rs = Nothing
                Set conn = Nothing
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                End If
                MsgBox Err.Description + MensajeError, vbCritical






End Sub

Private Sub LimpiarDireccionRepresentante()
    
    With DirRep
    
     .vTipoTelefono = ""
    .vNumTelefono = ""
    .vCodigoTelefono = ""
    .vTipoTelefono2 = ""
    .vNumTelefono2 = ""
    .vCodigoTelefono2 = ""
    .vTipoVia = ""
    .vDireccion = ""
    .vNumero = ""
    .vTipoPref = ""
    .vInterior = ""
    .vManzana = ""
    .vLote = ""
    .vEtapa = ""
    .vTipoConj = ""
    .vConjHabit = ""
    .vTipoBlock = ""
    .vNumBlock = ""
    .vReferencia = ""
    .vcodeDepar = 0
    .vcodeProv = 0
    .vCodeDistr = 0
    .vCodLoad = 0
    .vNomDepartamento = ""
    .vNomProvincia = ""
    .vNomDistrito = ""
    .vCodDireccion = ""
    .vgls_desdirebusq = ""
    
    
    End With
    

End Sub

Private Function DatosRepresentanteGestorCliente(ByVal pnum_poliza As String, ByVal Operacion As String) As Boolean

            Dim rest As New ChilkatRest
            Dim success As Long
            
            Dim glob As New ChilkatGlobal
           success = glob.UnlockBundle("GVNCRZ.CB1032023_x4BpcXzLDR4D")
            If (success <> 1) Then
                Debug.Print glob.LastErrorText
                DatosRepresentanteGestorCliente = False
                Exit Function
            End If
            
           ' URL: https://soatservicios.protectasecurity.pe/WSGestorCliente/Api/Cliente/ValidarCliente
            Dim bTls As Long
            bTls = 1
            Dim port As Long
            port = 443
            Dim bAutoReconnect As Long
            bAutoReconnect = 1
            'success = rest.Connect("soatservicios.protectasecurity.pe", port, bTls, bAutoReconnect)
            'success = rest.Connect("10.10.1.51", port, bTls, bAutoReconnect)
            success = rest.Connect("10.10.1.58", port, bTls, bAutoReconnect)
            If (success <> 1) Then
                Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
                Debug.Print rest.LastErrorText
                 DatosRepresentanteGestorCliente = False
                Exit Function
            End If
            
            
            Dim JSONchk As New ChilkatJsonObject

            Dim objCmd As ADODB.Command
            On Error GoTo ManejoError
                 
            Set objCmd = New ADODB.Command
            
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
             
             Dim conn As ADODB.Connection
             Set conn = New ADODB.Connection
             
            conn.Provider = "OraOLEDB.Oracle"
            conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
            conn.CursorLocation = adUseClient
            conn.Open
                
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMARRVV.sp_DatosRepresentanteGC"
                objCmd.CommandType = adCmdStoredProc
                
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 15, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("p_outNumError", adNumeric, adParamOutput, 5)
                objCmd.Parameters.Append param2
                
                Set param3 = objCmd.CreateParameter("p_outMsgError", adNumeric, adParamOutput, 200)
                objCmd.Parameters.Append param3
        
            
                Set rs = objCmd.Execute
                
                
                If Not rs.EOF() Then
                
                
                    success = JSONchk.UpdateString("p_CodAplicacion", "SEACSA")
                    success = JSONchk.UpdateString("p_TipOper", "INS")
                    success = JSONchk.UpdateString("P_NUSERCODE", vgUsuario)
                    success = JSONchk.UpdateString("P_NIDDOC_TYPE", rs!P_NIDDOC_TYPE)
                    success = JSONchk.UpdateString("P_SIDDOC", rs!P_SIDDOC)
                    success = JSONchk.UpdateString("P_SFIRSTNAME", rs!P_SFIRSTNAME)
                    success = JSONchk.UpdateString("P_SLASTNAME", rs!P_SLASTNAME)
                    success = JSONchk.UpdateString("P_SLASTNAME2", rs!P_SLASTNAME2)
                    success = JSONchk.UpdateString("P_SLEGALNAME", rs!P_SLEGALNAME)
                    success = JSONchk.UpdateString("P_SSEXCLIEN", rs!P_SSEXCLIEN)
                    success = JSONchk.UpdateString("P_NINCAPACITY", rs!P_NINCAPACITY)
                    success = JSONchk.UpdateString("P_DBIRTHDAT", rs!P_DBIRTHDAT)
                    success = JSONchk.UpdateString("p_DINCAPACITY", rs!p_DINCAPACITY)
                    success = JSONchk.UpdateString("P_NSPECIALITY", rs!P_NSPECIALITY)
                    success = JSONchk.UpdateString("P_NCIVILSTA", rs!P_NCIVILSTA)
                    success = JSONchk.UpdateString("P_NTITLE", rs!P_NTITLE)
                    success = JSONchk.UpdateString("P_NAFP", rs!P_NAFP)
                    success = JSONchk.UpdateString("P_NNATIONALITY", rs!P_NNATIONALITY)
                    success = JSONchk.UpdateString("P_SBAJAMAIL_IND", rs!P_SBAJAMAIL_IND)
                    success = JSONchk.UpdateString("P_SPROTEG_DATOS_IND", rs!P_SPROTEG_DATOS_IND)
                    success = JSONchk.UpdateString("P_SISCLIENT_IND", rs!P_SISCLIENT_IND)
                    success = JSONchk.UpdateString("P_SISRENIEC_IND", rs!P_SISRENIEC_IND)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_SRECTYPE", "2")
                    success = JSONchk.UpdateString("EListAddresClient[0].P_NCOUNTRY", "1")
                    success = JSONchk.UpdateString("EListAddresClient[0].P_NPROVINCE", rs!P_NPROVINCE)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_NLOCAL", rs!P_NLOCAL)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_NMUNICIPALITY", rs!P_NMUNICIPALITY)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_STI_DIRE", rs!P_STI_DIRE)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_SNOM_DIRECCION", rs!P_SNOM_DIRECCION)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_SNUM_DIRECCION", rs!P_SNUM_DIRECCION)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_STI_BLOCKCHALET", rs!P_STI_BLOCKCHALET)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_SBLOCKCHALET", rs!P_SBLOCKCHALET)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_STI_INTERIOR", rs!P_STI_INTERIOR)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_SNUM_INTERIOR", rs!P_SNUM_INTERIOR)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_STI_CJHT", rs!P_STI_CJHT)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_SNOM_CJHT", rs!P_SNOM_CJHT)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_SETAPA", rs!P_SETAPA)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_SMANZANA", rs!P_SMANZANA)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_SLOTE", rs!P_SLOTE)
                    success = JSONchk.UpdateString("EListAddresClient[0].P_SREFERENCIA", rs!P_SREFERENCIA)
                    success = JSONchk.UpdateString("EListPhoneClient[0].P_TIPOPER", "DEL")
                    success = JSONchk.UpdateString("EListPhoneClient[0].P_NAREA_CODE", rs!P_NAREA_CODE)
                    success = JSONchk.UpdateString("EListPhoneClient[0].P_SPHONE", rs!P_SPHONE)
                    success = JSONchk.UpdateString("EListPhoneClient[0].P_NPHONE_TYPE", rs!P_NPHONE_TYPE)
                    success = JSONchk.UpdateString("EListPhoneClient[1].P_NAREA_CODE", "1")
                    success = JSONchk.UpdateString("EListPhoneClient[1].P_SPHONE", rs!P_SPHONE2)
                    success = JSONchk.UpdateString("EListPhoneClient[1].P_NPHONE_TYPE", "2")
                    success = JSONchk.UpdateString("EListEmailClient[0].P_NROW", "1")
                    success = JSONchk.UpdateString("EListEmailClient[0].P_SRECTYPE", "4")
                    success = JSONchk.UpdateString("EListEmailClient[0].P_SE_MAIL", rs!P_SE_MAIL)

          
            End If
            
            success = rest.AddHeader("Content-Type", "application/json")
            
            Dim sbRequestBody As New ChilkatStringBuilder
            success = JSONchk.EmitSb(sbRequestBody)
            Dim sbResponseBody As New ChilkatStringBuilder
            'success = rest.FullRequestSb("POST", "/WSGestorCliente/Api/Cliente/ValidarCliente", sbRequestBody, sbResponseBody)
            success = rest.FullRequestSb("POST", "/WSGestorCliente/Api/Cliente/GestionarCliente", sbRequestBody, sbResponseBody)

            If (success <> 1) Then
                Debug.Print rest.LastErrorText
                 DatosRepresentanteGestorCliente = False
                Exit Function
                
            End If
            
            Dim respStatusCode As Long
            respStatusCode = rest.ResponseStatusCode
            Debug.Print "response status code = " & respStatusCode
            If (respStatusCode >= 400) Then
                Debug.Print "Response Status Code = " & respStatusCode
                Debug.Print "Response Header:"
                Debug.Print rest.ResponseHeader
                Debug.Print "Response Body:"
                Debug.Print sbResponseBody.GetAsString()
                 DatosRepresentanteGestorCliente = False
                Exit Function
            End If
            
        Dim VidDocumento As String
        Dim p As Object

        Set p = json.parse(sbResponseBody.GetAsString)
        If p.Item("P_NCODE") <> "0" Then
        
            MsgBox "Error en Representante: " & p.Item("P_SMESSAGE"), vbCritical
             DatosRepresentanteGestorCliente = False
            Exit Function
        End If
        
        
          
        conn.Close
        Set rs = Nothing
        Set conn = Nothing

        DatosRepresentanteGestorCliente = True
        Exit Function
    
ManejoError:
                conn.Close
                Set rs = Nothing
                Set conn = Nothing
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                End If
                DatosRepresentanteGestorCliente = False
                MsgBox Err.Description + MensajeError, vbCritical


End Function
Private Sub RepresentanteToGC(ByVal vNum_poliza As String, ByRef vRep As RepresentaGC)

        Dim objCmd As ADODB.Command
        On Error GoTo ManejoError
             
        Set objCmd = New ADODB.Command
        
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
         
        Dim conn As ADODB.Connection
        Set conn = New ADODB.Connection
        
        conn.Provider = "OraOLEDB.Oracle"
        conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
        conn.CursorLocation = adUseClient
        conn.Open

        Set objCmd = CreateObject("ADODB.Command")
        Set objCmd.ActiveConnection = conn
        
        objCmd.CommandText = "PKG_API_FIRMARRVV.sp_DatosRepresentanteGC"
        objCmd.CommandType = adCmdStoredProc
        
        Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 15, vNum_poliza)
        objCmd.Parameters.Append param1
        
        Set param2 = objCmd.CreateParameter("p_outNumError", adNumeric, adParamOutput, 5)
        objCmd.Parameters.Append param2
                
        Set param3 = objCmd.CreateParameter("p_outMsgError", adNumeric, adParamOutput, 200)
        objCmd.Parameters.Append param3
        
        
        Set rs = objCmd.Execute
        If objCmd.Parameters.Item("p_outNumError").Value <> 0 Then GoTo ManejoError
        
          vRep.P_DATOS = "N"
          
     While Not rs.EOF()
            vRep.P_NIDDOC_TYPE = rs!P_NIDDOC_TYPE
            vRep.P_NTIPCONT = rs!P_NTIPCONT
            vRep.P_SAPEPAT = rs!P_SAPEPAT
            vRep.P_SAPEMAT = rs!P_SAPEMAT
            vRep.P_SIDDOC = rs!P_SIDDOC
            vRep.P_SNOMBRES = rs!P_SNOMBRES
            vRep.P_SPHONE = rs!P_SPHONE
            vRep.P_DATOS = "S"
            
            
            rs.MoveNext
    
    Wend
 
                 
      conn.Close
      Set rs = Nothing
      Set conn = Nothing
      Exit Sub
        
ManejoError:
                conn.Close
                Set rs = Nothing
                Set conn = Nothing
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                End If
                
                MsgBox Err.Description + MensajeError, vbCritical

End Sub
Private Function ValidaEnvio(ByVal pnum_cot As String) As String


                Dim conn    As ADODB.Connection
                Set conn = New ADODB.Connection
                Set rs = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
                Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")
                Dim MensajeError As String
                Dim vResultado As String
                
                
                On Error GoTo ManejoError
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
                
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMA_DOCUMENTOS.sp_validarEnvio"
                objCmd.CommandType = adCmdStoredProc
                
                Set param1 = objCmd.CreateParameter("pnum_cot", adVarChar, adParamInput, 12, pnum_cot)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("p_EnviadoAntes", adVarChar, adParamOutput, 50)
                objCmd.Parameters.Append param2
                
                Set param3 = objCmd.CreateParameter("p_outNumError", adNumeric, adParamOutput, 5)
                objCmd.Parameters.Append param3
                
                Set param4 = objCmd.CreateParameter("p_outMsgError", adNumeric, adParamOutput, 200)
                objCmd.Parameters.Append param4
                
                objCmd.Execute
          
                
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                    GoTo ManejoError
               End If
               
               
               vResultado = objCmd.Parameters.Item("p_EnviadoAntes").Value
               
               
               ValidaEnvio = vResultado
               Exit Function
   
ManejoError:
                conn.Close
                Set rs = Nothing
                Set conn = Nothing
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                End If
                
                MsgBox Err.Description + MensajeError, vbCritical

End Function

Private Function getAgentesComerciales(ByVal pnum_poliza As String) As AgentesCom


                Dim objCmd As ADODB.Command
                Dim rs As ADODB.Recordset
                Dim conn As ADODB.Connection
                Dim oAgentesCom As AgentesCom
                
                On Error GoTo ManejoError
              
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
                
                objCmd.CommandText = "PKG_API_FIRMARRVV.sp_DatosAgentesComerciales"
                objCmd.CommandType = adCmdStoredProc
                

                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 15, pnum_poliza)
                objCmd.Parameters.Append param1
       
                Set param2 = objCmd.CreateParameter("p_outNumError", adNumeric, adParamOutput, 5)
                objCmd.Parameters.Append param2
                
                Set param3 = objCmd.CreateParameter("p_outMsgError", adNumeric, adParamOutput, 200)
                objCmd.Parameters.Append param3
                
        
                Set rs = objCmd.Execute
                
                While Not rs.EOF
                                    
                    oAgentesCom.NombreAsesor = IIf(IsNull(rs!NombreAsesor), "", rs!NombreAsesor)
                    oAgentesCom.MailAsesor = IIf(IsNull(rs!correoasesor), "", rs!correoasesor)
                    
                    oAgentesCom.NombreSupervisor = IIf(IsNull(rs!nombresuper), "", rs!nombresuper)
                    oAgentesCom.MailSupervisor = IIf(IsNull(rs!correosuper), "", rs!correosuper)
                        
                    rs.MoveNext
               Wend
        
                   
         conn.Close
        Set rs = Nothing
        Set conn = Nothing
        
        getAgentesComerciales = oAgentesCom
        Exit Function
        
        
ManejoError:
                conn.Close
                Set rs = Nothing
                Set conn = Nothing
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                End If
                
                MsgBox Err.Description + MensajeError, vbCritical

End Function
    
Private Sub Limpiar_RepresentaGC()

        P_DATOS = ""
        P_NTIPCONT = ""
        P_SIDDOC = ""
        P_NIDDOC_TYPE = ""
        P_SNOMBRES = ""
        P_SAPEPAT = ""
        P_SAPEMAT = ""
        P_SPHONE = ""
End Sub


