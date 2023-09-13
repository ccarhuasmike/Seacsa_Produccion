VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_RecalculoPol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recalculo de Pólizas."
   ClientHeight    =   7260
   ClientLeft      =   2085
   ClientTop       =   1665
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   11490
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   60
      TabIndex        =   49
      Top             =   5850
      Width           =   11325
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   10500
         Picture         =   "Frm_RecalculoPol.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab_Poliza 
      Height          =   5775
      Left            =   90
      TabIndex        =   13
      Top             =   60
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10186
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "C"
      TabPicture(0)   =   "Frm_RecalculoPol.frx":00FA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Msf_GriAseg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ProgressBar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "R"
      TabPicture(1)   =   "Frm_RecalculoPol.frx":0116
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   345
         Left            =   7470
         TabIndex        =   51
         Top             =   690
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
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
         Height          =   915
         Left            =   420
         TabIndex        =   44
         Top             =   210
         Width           =   8295
         Begin MSComCtl2.DTPicker DTPDesde 
            Height          =   345
            Left            =   1560
            TabIndex        =   8
            Top             =   480
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            _Version        =   393216
            Format          =   74579969
            CurrentDate     =   43566
         End
         Begin VB.CommandButton Cmd_Excel 
            Caption         =   "&Excel"
            Enabled         =   0   'False
            Height          =   675
            Left            =   6270
            Picture         =   "Frm_RecalculoPol.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Exportar  a Excel"
            Top             =   150
            Width           =   675
         End
         Begin VB.CommandButton Cmd_BuscarCotizaciones 
            Caption         =   "&Buscar"
            Height          =   675
            Left            =   5460
            Picture         =   "Frm_RecalculoPol.frx":05E9
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Buscar Todas las Póliza"
            Top             =   150
            Width           =   720
         End
         Begin VB.TextBox Txt_Poliza 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPHasta 
            Height          =   345
            Left            =   3330
            TabIndex        =   9
            Top             =   450
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            _Version        =   393216
            Format          =   74579969
            CurrentDate     =   43566
         End
         Begin VB.Label Lbl_Buscador 
            AutoSize        =   -1  'True
            Caption         =   "Póliza"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Lbl_Buscador 
            Caption         =   "Hasta"
            Height          =   255
            Index           =   2
            Left            =   3330
            TabIndex        =   46
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Lbl_Buscador 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   0
            Left            =   1560
            TabIndex        =   45
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Recepción de Fecha de Cargos Varios."
         Height          =   5415
         Left            =   -74580
         TabIndex        =   14
         Top             =   120
         Width           =   10725
         Begin VB.TextBox Txt_Endoso 
            Height          =   285
            Left            =   5340
            MaxLength       =   10
            TabIndex        =   1
            Text            =   "1"
            Top             =   240
            Width           =   585
         End
         Begin VB.CommandButton Cmd_BuscarPol 
            Height          =   495
            Left            =   6030
            Picture         =   "Frm_RecalculoPol.frx":06EB
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Póliza"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox Txt_NumPol 
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   0
            Top             =   240
            Width           =   2625
         End
         Begin VB.Frame Frame4 
            Caption         =   "Registro de Interación Externos"
            Height          =   2145
            Left            =   180
            TabIndex        =   17
            Top             =   3030
            Width           =   10455
            Begin MSComCtl2.DTPicker DTPRecalculo 
               Height          =   345
               Left            =   6060
               TabIndex        =   3
               Top             =   360
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   74579969
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin VB.CommandButton Cmd_Grabar 
               Caption         =   "&Grabar"
               Enabled         =   0   'False
               Height          =   675
               Left            =   9660
               Picture         =   "Frm_RecalculoPol.frx":07ED
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Grabar Datos de Póliza"
               Top             =   1380
               Width           =   720
            End
            Begin MSComCtl2.DTPicker DTPPoliza 
               Height          =   345
               Left            =   6060
               TabIndex        =   4
               Top             =   720
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   74579969
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker DTPFichas 
               Height          =   345
               Left            =   6060
               TabIndex        =   5
               Top             =   1080
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   74579969
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Recepción de cargos de fichas de datos del cliente"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   43
               Top             =   1200
               Width           =   4365
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Recepción de cargos de la póliza (Entrega de la póliza al cliente)"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   42
               Top             =   810
               Width           =   5295
            End
            Begin VB.Label Lbl_Calculo 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Recepción de cargos AFP (Recalculos)"
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   41
               Top             =   420
               Width           =   3525
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Datos Representate de la Póliza"
            Height          =   2265
            Left            =   5760
            TabIndex        =   16
            Top             =   660
            Width           =   4905
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Número Ident."
               Height          =   195
               Index           =   17
               Left            =   90
               TabIndex        =   40
               Top             =   585
               Width           =   1005
            End
            Begin VB.Label Lbl_DgvRep 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1290
               TabIndex        =   39
               Top             =   585
               Width           =   2775
            End
            Begin VB.Label Lbl_RutRep 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1290
               TabIndex        =   38
               Top             =   270
               Width           =   2775
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Nombres"
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   37
               Top             =   900
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Ap. Paterno"
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   36
               Top             =   1215
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               Caption         =   "Ap. Materno"
               Height          =   255
               Index           =   2
               Left            =   90
               TabIndex        =   35
               Top             =   1500
               Width           =   1215
            End
            Begin VB.Label Lbl_Benef 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Ident."
               Height          =   195
               Index           =   15
               Left            =   90
               TabIndex        =   34
               Top             =   300
               Width           =   765
            End
            Begin VB.Label Lbl_NomRep 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1290
               TabIndex        =   33
               Top             =   900
               Width           =   3375
            End
            Begin VB.Label Lbl_ApPatRep 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1290
               TabIndex        =   32
               Top             =   1215
               Width           =   3375
            End
            Begin VB.Label Lbl_ApMatRep 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1290
               TabIndex        =   31
               Top             =   1500
               Width           =   3375
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Datos Cliente Contratante"
            Height          =   2265
            Left            =   180
            TabIndex        =   15
            Top             =   660
            Width           =   5475
            Begin VB.Label Lbl_Afiliado 
               AutoSize        =   -1  'True
               Caption         =   "Ap. Paterno"
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   30
               Top             =   1530
               Width           =   840
            End
            Begin VB.Label Lbl_Afiliado 
               AutoSize        =   -1  'True
               Caption         =   "Ap. Materno"
               Height          =   195
               Index           =   3
               Left            =   90
               TabIndex        =   29
               Top             =   1815
               Width           =   870
            End
            Begin VB.Label Lbl_Afiliado 
               AutoSize        =   -1  'True
               Caption         =   "2do. Nombre"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   28
               Top             =   1230
               Width           =   915
            End
            Begin VB.Label Lbl_Afiliado 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Identificación"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   27
               Top             =   345
               Width           =   1305
            End
            Begin VB.Label Lbl_RutAfi 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1485
               TabIndex        =   26
               Top             =   300
               Width           =   2295
            End
            Begin VB.Label Lbl_DgvAfi 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1485
               TabIndex        =   25
               Top             =   600
               Width           =   2295
            End
            Begin VB.Label Lbl_NomAfi 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1485
               TabIndex        =   24
               Top             =   900
               Width           =   3660
            End
            Begin VB.Label Lbl_ApMatAfi 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1485
               TabIndex        =   23
               Top             =   1770
               Width           =   3660
            End
            Begin VB.Label Lbl_ApPatAfi 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1485
               TabIndex        =   22
               Top             =   1485
               Width           =   3660
            End
            Begin VB.Label Lbl_Afiliado 
               AutoSize        =   -1  'True
               Caption         =   "Nº Identificación"
               Height          =   195
               Index           =   20
               Left            =   90
               TabIndex        =   21
               Top             =   645
               Width           =   1170
            End
            Begin VB.Label Lbl_NomAfiSeg 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1485
               TabIndex        =   20
               Top             =   1185
               Width           =   3660
            End
            Begin VB.Label Lbl_Afiliado 
               AutoSize        =   -1  'True
               Caption         =   "1er. Nombre"
               Height          =   195
               Index           =   21
               Left            =   90
               TabIndex        =   19
               Top             =   930
               Width           =   870
            End
         End
         Begin VB.Label Lbl_Nombre 
            AutoSize        =   -1  'True
            Caption         =   "Num. Endoso"
            Height          =   195
            Index           =   6
            Left            =   4290
            TabIndex        =   48
            Top             =   285
            Width           =   960
         End
         Begin VB.Label Lbl_Nombre 
            Caption         =   "N° Póliza"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   18
            Top             =   255
            Width           =   855
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Msf_GriAseg 
         Height          =   4515
         Left            =   390
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1140
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   7964
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
   End
End
Attribute VB_Name = "Frm_RecalculoPol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_BuscarCotizaciones_Click()
  Call flCargaCarpBenef(Txt_Poliza)

End Sub

Private Sub Cmd_BuscarPol_Click()
On Error GoTo Err_Buscar
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



Private Sub Cmd_Grabar_Click()
On Error GoTo Err_Salir
    vlSql = "UPDATE PD_TMAE_POLIZA SET FEC_RECCARAFP=TO_CHAR(TO_DATE('" & DTPRecalculo & "','dd/mm/yyyy'), 'YYYYMMDD'), "
    vlSql = vlSql & "FEC_RECCARPOL=TO_CHAR(TO_DATE('" & DTPPoliza & "','dd/mm/yyyy'), 'YYYYMMDD'), "
    vlSql = vlSql & "FEC_RECCARCLI=TO_CHAR(TO_DATE('" & DTPFichas & "','dd/mm/yyyy'), 'YYYYMMDD') "
    vlSql = vlSql & "WHERE NUM_POLIZA='" & Txt_NumPol & "' AND NUM_ENDOSO=" & Txt_Endoso & ""
    vgConexionBD.Execute (vlSql)
    MsgBox "grabado.", vbInformation, "Mensaje..."
Exit Sub
Err_Salir:
Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Excel_Click()
Dim xlapp As Excel.Application
Dim sArchivo As String
Dim ix As Long
Dim Sql As String
Dim vlRegistro As ADODB.Recordset

    Screen.MousePointer = 11
    
    Set xlapp = CreateObject("excel.application")

    sArchivo = App.Path & "\Plantilla_polizas.xls"

    xlapp.Visible = True 'para ver vista previa
    xlapp.WindowState = 2 ' minimiza excel
    xlapp.Workbooks.Open (sArchivo)

    ix = 2
    
    vlSql = ""
    vlSql = "SELECT A.* FROM (SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.FEC_TRASPASO FEC_EMISION, FEC_RECCARAFP, FEC_RECCARPOL, FEC_RECCARCLI "
    vlSql = vlSql & "FROM PD_TMAE_POLIZA A JOIN PD_TMAE_POLPRIREC B ON A.NUM_POLIZA=B.NUM_POLIZA "
    vlSql = vlSql & "WHERE A.NUM_POLIZA='" & Txt_Poliza & "' AND B.FEC_TRASPASO >= TO_CHAR(TO_DATE('" & DTPDesde & "','dd/mm/yyyy'), 'YYYYMMDD') AND B.FEC_TRASPASO <= TO_CHAR(TO_DATE('" & DTPHasta & "','dd/mm/yyyy'), 'YYYYMMDD') order by num_endoso desc )A WHERE ROWNUM=1"
    
    Set vlRegistro = vgConexionBD.Execute(vlSql)
   
    Dim registros As Variant
    
    registros = vlRegistro.GetRows()
    
    'MsgBox "Cantidad de registros: " & UBound(registros, 2) + 1
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = UBound(registros, 2) + 1
    
    vlRegistro.MoveFirst
    
    If Not vlRegistro.EOF Then
        While Not vlRegistro.EOF
            xlapp.Range("A" & ix) = IIf(IsNull(vlRegistro(0).Value) = True, "", vlRegistro(0).Value)
            xlapp.Range("B" & ix) = IIf(IsNull(vlRegistro(1).Value) = True, "", vlRegistro(1).Value)
                       
            xlapp.Range("C" & ix) = IIf(IsNull(vlRegistro(2).Value) = True, "", Mid(vlRegistro(2).Value, 1, 4) + "/" + Mid(vlRegistro(2).Value, 5, 2) + "/" + Mid(vlRegistro(2).Value, 7, 2))
            xlapp.Range("D" & ix) = IIf(IsNull(vlRegistro(3).Value) = True, "", Mid(vlRegistro(3).Value, 1, 4) + "/" + Mid(vlRegistro(3).Value, 5, 2) + "/" + Mid(vlRegistro(3).Value, 7, 2))
            xlapp.Range("E" & ix) = IIf(IsNull(vlRegistro(4).Value) = True, "", Mid(vlRegistro(4).Value, 1, 4) + "/" + Mid(vlRegistro(4).Value, 5, 2) + "/" + Mid(vlRegistro(4).Value, 7, 2))
            xlapp.Range("F" & ix) = IIf(IsNull(vlRegistro(5).Value) = True, "", Mid(vlRegistro(5).Value, 1, 4) + "/" + Mid(vlRegistro(5).Value, 5, 2) + "/" + Mid(vlRegistro(5).Value, 7, 2))
'            xlapp.Range("D" & ix) = IIf(IsNull(vlRegistro(3).Value) = True, "", vlRegistro(3).Value)
'            xlapp.Range("E" & ix) = IIf(IsNull(vlRegistro(4).Value) = True, "", vlRegistro(4).Value)
'            xlapp.Range("F" & ix) = IIf(IsNull(vlRegistro(5).Value) = True, "", vlRegistro(5).Value)
            ix = ix + 1
            ProgressBar1.Value = ix - 2
            vlRegistro.MoveNext
        Wend
    End If
    Screen.MousePointer = 0
    xlapp.WindowState = xlMaximized
    'xlapp.Workbooks.Close (sArchivo)
    ProgressBar1.Value = 0
    MsgBox ("Datos Exportados!")
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

Private Sub Form_Load()
On Error GoTo Err_Form
        
    Frm_RecalculoPol.Top = 0
    Frm_RecalculoPol.Left = 0
        
    Txt_NumPol.Enabled = True
    SSTab_Poliza.Tab = 1
   
    'SSTab_Poliza.Enabled = False
    
    Msf_GriAseg.Clear
    Msf_GriAseg.Cols = 6
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
    
    Msf_GriAseg.Col = 3
    Msf_GriAseg.ColWidth(3) = 1500
    Msf_GriAseg.Text = "FECHA REC. AFP"
    
    Msf_GriAseg.Col = 4
    Msf_GriAseg.ColWidth(4) = 1500
    Msf_GriAseg.Text = "FECHA REC. POLIZA"
    
    Msf_GriAseg.Col = 5
    Msf_GriAseg.ColWidth(5) = 1500
    Msf_GriAseg.Text = "FECHA REC. FICHA DE ACTUALIZACION DATOS"
    
    DTPRecalculo.Value = DTPRecalculo.MinDate
    DTPRecalculo.Value = Null

    DTPPoliza.Value = DTPPoliza.MinDate
    DTPPoliza.Value = Null

    DTPFichas.Value = DTPFichas.MinDate
    DTPFichas.Value = Null


     
 
    
Exit Sub
Err_Form:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Form_LostFocus()
Txt_NumPol.SetFocus
End Sub

Private Sub SSTab_Poliza_Click(PreviousTab As Integer)
On Error GoTo Err_

    If PreviousTab = 0 Then
        Txt_NumPol.SetFocus
        SendKeys "{home}+{end}"
    ElseIf PreviousTab = 1 Then
        Txt_Poliza.SetFocus
        SendKeys "{home}+{end}"
        
        DTPDesde.Value = Date - 1
        DTPHasta.Value = Date
    End If
Err_:
'Screen.MousePointer = 0
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
        If Trim(Txt_NumPol) <> "" And Trim(Txt_Endoso) <> "" Then
           Dim vlRegistro As New ADODB.Recordset
           vgSql = "SELECT MAX(NUM_ENDOSO) AS NUM_ENDOSO FROM PD_TMAE_POLIZA WHERE NUM_POLIZA=" & Txt_NumPol & ""
        
           Set vlRegistro = vgConexionBD.Execute(vgSql)
           If Not (vlRegistro.EOF) Then
               Txt_Endoso = IIf(IsNull(vlRegistro!Num_Endoso), "1", vlRegistro!Num_Endoso)
           End If
           Txt_Endoso.SetFocus
           
        Else
            MsgBox "Debe Ingresar Número de Póliza y Endoso", vbExclamation, "Falta Información"
            Txt_NumPol.SetFocus
        End If

          
    End If
End Sub

'-------------------------------------------------------
'FUNCION BUSCA DATOS DE POLIZA PARA CARGAR EN FORMULARIO
'-------------------------------------------------------
Function flBuscaPoliza(iNumPol As String, inumend As Integer)
On Error GoTo Err_BuscaPol

    SSTab_Poliza.Tab = 1
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
        Cmd_Grabar.Enabled = False
        
        Lbl_RutAfi.Caption = ""
        Lbl_DgvAfi.Caption = ""
        Lbl_NomAfi.Caption = ""
        Lbl_NomAfiSeg.Caption = ""
        Lbl_ApPatAfi.Caption = ""
        Lbl_ApMatAfi.Caption = ""
        
        DTPRecalculo.Value = DTPRecalculo.MinDate
        DTPRecalculo.Value = Null
        
        DTPPoliza.Value = DTPPoliza.MinDate
        DTPPoliza.Value = Null
        
        DTPFichas.Value = DTPFichas.MinDate
        DTPFichas.Value = Null
        
        Lbl_RutRep.Caption = ""
        Lbl_DgvRep.Caption = ""
        Lbl_NomRep.Caption = ""
        Lbl_ApPatRep.Caption = ""
        Lbl_ApMatRep.Caption = ""
                
        
        Exit Function
    End If
    Cmd_Grabar.Enabled = True
    Call flCargaCarpAfilPol(iNumPol, inumend)
    Call flCargaCarpAfilPolRep(iNumPol, inumend)
    

    'Msf_GriAseg.Enabled = True
    'Fra_Poliza.Enabled = False
    'Cmd_Poliza.Enabled = False
    SSTab_Poliza.Enabled = True
    SSTab_Poliza.Tab = 1
    
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
    vlSql = "SELECT p.num_poliza,p.num_endoso,p.num_cot,p.num_operacion,p.fec_vigencia,p.FEC_RECCARAFP,p.FEC_RECCARPOL,p.FEC_RECCARCLI, "
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
        
        'datos de la carpeta de afiliado
        If Not IsNull(vgRs!GLS_TIPOIDENCOR) Then Lbl_RutAfi = vgRs!GLS_TIPOIDENCOR
        If Not IsNull(vgRs!Num_IdenBen) Then Lbl_DgvAfi = vgRs!Num_IdenBen
        If Not IsNull(vgRs!Gls_NomBen) Then Lbl_NomAfi = vgRs!Gls_NomBen
        If Not IsNull(vgRs!Gls_NomSegBen) Then Lbl_NomAfiSeg = vgRs!Gls_NomSegBen Else Lbl_NomAfiSeg = ""
        If Not IsNull(vgRs!Gls_PatBen) Then Lbl_ApPatAfi = vgRs!Gls_PatBen
        If Not IsNull(vgRs!Gls_MatBen) Then Lbl_ApMatAfi = vgRs!Gls_MatBen
        
        
        If Not IsNull(vgRs!FEC_RECCARAFP) Then
            Dim fechaAFP As String
            fechaAFP = Mid(vgRs!FEC_RECCARAFP, 1, 4) + "/" + Mid(vgRs!FEC_RECCARAFP, 5, 2) + "/" + Mid(vgRs!FEC_RECCARAFP, 7, 2)
            DTPRecalculo.Value = CDate(fechaAFP)
        Else
           
            DTPRecalculo.Value = DTPRecalculo.MinDate
            DTPRecalculo.Value = Null
        End If
        
        If Not IsNull(vgRs!FEC_RECCARPOL) Then
            Dim fechaPOL As String
            fechaPOL = Mid(vgRs!FEC_RECCARPOL, 1, 4) + "/" + Mid(vgRs!FEC_RECCARPOL, 5, 2) + "/" + Mid(vgRs!FEC_RECCARPOL, 7, 2)
            DTPPoliza.Value = CDate(fechaPOL)
        Else
            DTPPoliza.Value = DTPPoliza.MinDate
            DTPPoliza.Value = Null
        End If
        
        If Not IsNull(vgRs!FEC_RECCARCLI) Then
            Dim fechaCLI As String
            fechaCLI = Mid(vgRs!FEC_RECCARCLI, 1, 4) + "/" + Mid(vgRs!FEC_RECCARCLI, 5, 2) + "/" + Mid(vgRs!FEC_RECCARCLI, 7, 2)
            DTPFichas.Value = CDate(fechaCLI)
        Else
            DTPFichas.Value = DTPFichas.MinDate
            DTPFichas.Value = Null
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

'Carga el representante de los beneficiarios, si lo hubiese
Function flCargaCarpAfilPolRep(iNumPol As String, inumend As Integer)
Dim vlNombresRep As String
Dim vlApepatRep As String
Dim vlApematRep As String

On Error GoTo Err_Cargarep
    Dim vlRut As String
    vlSql = ""
    vlSql = "SELECT * FROM pd_tmae_polrep b , ma_tpar_tipoiden i WHERE b.cod_tipoidenrep = i.cod_tipoiden "
    vlSql = vlSql & "and num_poliza = '" & iNumPol & "' "
    vlSql = vlSql & "and num_endoso= " & inumend & " "
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not vgRs.EOF Then
    
        If Not IsNull(vgRs!GLS_TIPOIDENCOR) Then
            Lbl_RutRep = vgRs!GLS_TIPOIDENCOR
        Else
            Lbl_RutRep = ""
        End If
    
        If Not IsNull(vgRs!NUM_IDENREP) Then
            Lbl_DgvRep = vgRs!NUM_IDENREP
        Else
            Lbl_DgvRep = ""
        End If
        
        'Lbl_NomRepSeg
        
        If Not IsNull(vgRs!Gls_NombresRep) Then
            Lbl_NomRep = vgRs!Gls_NombresRep
        Else
            Lbl_NomRep = ""
        End If
        
         If Not IsNull(vgRs!Gls_ApepatRep) Then
            Lbl_ApPatRep = vgRs!Gls_ApepatRep
        Else
            Lbl_ApPatRep = ""
        End If
        If Not IsNull(vgRs!Gls_ApematRep) Then
            Lbl_ApMatRep = vgRs!Gls_ApematRep
        Else
            Lbl_ApMatRep = ""
        End If
    
       
       
        'Lbl_Representante.Caption = vlNombresRep & " " & vlApepatRep & " " & vlApematRep
    End If
    vgRs.Close
 
Exit Function
Err_Cargarep:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Private Sub Txt_NumPol_LostFocus()
    If Trim(Txt_NumPol) <> "" Then
        Txt_NumPol = Format(Txt_NumPol, "000000000#")
    End If
End Sub

'-------------------------------------------
'CARGA INFORMACION EN LA GRILLA
'-------------------------------------------
Function flCargaCarpBenef(iNumPol As String)
Dim vlFechaEmis As String
Dim vlFechaRECCARAFP As String
Dim vlFechaRECCARPOL As String
Dim vlFechaRECCARCLI As String


On Error GoTo Err_CargaBen
    Dim vlRut As String
    vlSql = ""
    vlSql = "SELECT A.* FROM (SELECT A.NUM_POLIZA, A.NUM_ENDOSO, B.FEC_TRASPASO FEC_EMISION, FEC_RECCARAFP, FEC_RECCARPOL, FEC_RECCARCLI "
    vlSql = vlSql & "FROM PD_TMAE_POLIZA A JOIN PD_TMAE_POLPRIREC B ON A.NUM_POLIZA=B.NUM_POLIZA "
    vlSql = vlSql & "WHERE A.NUM_POLIZA='" & iNumPol & "' AND B.FEC_TRASPASO >= TO_CHAR(TO_DATE('" & DTPDesde & "','dd/mm/yyyy'), 'YYYYMMDD') AND B.FEC_TRASPASO <= TO_CHAR(TO_DATE('" & DTPHasta & "','dd/mm/yyyy'), 'YYYYMMDD') order by num_endoso desc) A WHERE ROWNUM=1"
    Set vgRs = vgConexionBD.Execute(vlSql)
    Msf_GriAseg.Rows = 1
    
    If Not vgRs.EOF Then
        Cmd_Excel.Enabled = True
    Else
        Cmd_Excel.Enabled = False
        MsgBox "No existe Registros con los criterios dados...", vbInformation, "Mensaje"
        
    End If
    
    While Not vgRs.EOF
        
        If Not IsNull(vgRs!Fec_Emision) Then
            vlFechaEmis = DateSerial(Mid(vgRs!Fec_Emision, 1, 4), Mid(vgRs!Fec_Emision, 5, 2), Mid(vgRs!Fec_Emision, 7, 2))
        Else
            vlFechaEmis = ""
        End If

        If Not IsNull(vgRs!FEC_RECCARAFP) Then
            vlFechaRECCARAFP = DateSerial(Mid(vgRs!FEC_RECCARAFP, 1, 4), Mid(vgRs!FEC_RECCARAFP, 5, 2), Mid(vgRs!FEC_RECCARAFP, 7, 2))
        Else
            vlFechaRECCARAFP = ""
        End If

        If Not IsNull(vgRs!FEC_RECCARPOL) Then
            vlFechaRECCARPOL = DateSerial(Mid(vgRs!FEC_RECCARPOL, 1, 4), Mid(vgRs!FEC_RECCARPOL, 5, 2), Mid(vgRs!FEC_RECCARPOL, 7, 2))
        Else
            vlFechaRECCARPOL = ""
        End If
        
        If Not IsNull(vgRs!FEC_RECCARCLI) Then
            vlFechaRECCARCLI = DateSerial(Mid(vgRs!FEC_RECCARCLI, 1, 4), Mid(vgRs!FEC_RECCARCLI, 5, 2), Mid(vgRs!FEC_RECCARCLI, 7, 2))
        Else
            vlFechaRECCARCLI = ""
        End If

                
        Msf_GriAseg.AddItem Trim(vgRs!Num_Poliza) & vbTab _
                            & vgRs!Num_Endoso & vbTab _
                            & vlFechaEmis & vbTab _
                            & vlFechaRECCARAFP & vbTab _
                            & vlFechaRECCARPOL & vbTab _
                            & vlFechaRECCARCLI & vbTab
                           
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



Private Sub Txt_Poliza_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Txt_Poliza = Format(Txt_Poliza, "000000000#")
    Cmd_BuscarCotizaciones_Click
      
  End If
End Sub

Private Sub Txt_Poliza_LostFocus()
   If Trim(Txt_Poliza) <> "" Then
        Txt_Poliza = Format(Txt_Poliza, "000000000#")
    End If
End Sub
