VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_ContableArch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Archivo Contable "
   ClientHeight    =   6315
   ClientLeft      =   2160
   ClientTop       =   1740
   ClientWidth     =   9495
   Icon            =   "Frm_ContableArch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9495
   Begin VB.Frame Frame3 
      Caption         =   "Opciones de Carga"
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
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Width           =   9255
      Begin VB.OptionButton Opt_Recepcionadas 
         Caption         =   "Recepcionadas"
         Height          =   195
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   1545
      End
      Begin VB.OptionButton Opt_Todas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   210
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones de Impresión"
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
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   9255
      Begin VB.OptionButton Opt_DetPen 
         Caption         =   "Detalle por Pensiones"
         Height          =   255
         Left            =   6360
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Opt_DetMov 
         Caption         =   "Detalle por Movimiento"
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Opt_Resumen 
         Caption         =   "Resumen"
         Height          =   195
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proceso de Carga"
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
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3375
      Begin VB.Label Lbl_FecCierre 
         Alignment       =   2  'Center
         BackColor       =   &H00E8FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Fra_Datos 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   9255
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Salir"
         Height          =   755
         Left            =   5880
         Picture         =   "Frm_ContableArch.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Generar"
         Height          =   755
         Left            =   2280
         Picture         =   "Frm_ContableArch.frx":053C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Generación de Archivo Contable"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   755
         Left            =   3480
         Picture         =   "Frm_ContableArch.frx":0D5E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Resumen de Carga"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmd_limpiar 
         Caption         =   "&Limpiar"
         Height          =   755
         Left            =   4680
         Picture         =   "Frm_ContableArch.frx":1418
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Fra_Datos 
      Height          =   735
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   9255
      Begin VB.CommandButton CmdContable 
         Height          =   375
         Left            =   8280
         Picture         =   "Frm_ContableArch.frx":1AD2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblArchivo 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Destino de Datos       :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Fra_Datos 
      Caption         =   " Fecha de ..."
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
      Height          =   1575
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   2880
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Lbl_NumArchivo 
         Alignment       =   2  'Center
         BackColor       =   &H00E8FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   15
         Top             =   1180
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Hasta       :"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Desde      :"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Periodos 
      Height          =   2535
      Left            =   3600
      TabIndex        =   12
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      BackColor       =   14745599
   End
   Begin MSComDlg.CommonDialog ComDialogo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_ContableArch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vlFecDesde          As String, vlFecHasta As String
Dim vlFecCrea           As String, vlHorCrea As String
Dim vlMoneda            As String
Dim vlSql               As String
Dim vlArchivoCont       As String
Dim vlArchPriUni        As Boolean
Dim vlArchPriPagos      As Boolean

'-------- Constantes para generar el archivo de Prima Unica -------------------
Const clTipRegPU        As String = "1"     '1 -Tipo de Registro
Const clSucPU           As String = "50"    '3 -SUCURSAL
Const clRamoContPU      As String = "76"    '10-RAMO CONTABLE
Const clFrecPagPU       As String = "0"     '13-FRECUENCIA DE PAGO
Const clRegTratFacPU    As String = "1"     '15-REGISTRO EN TRATAMIENTO (Factura)
Const clRegTratCredPU   As String = "2"     '15-REGISTRO EN TRATAMIENTO (Nota de Crédito)
Const clTipInterPU      As String = "A"     '22-TIPO DE INTERMEDIARIO
Const clReaPU           As String = "2"     '23-REASEGURADOR
Const clTipMovPU        As String = "32"    '29-TIPO MOVIMIENTO

'-------- Constantes para generar el archivo de Primeros Pagos ----------------
Const clTipRegPP        As String = "5"     '1-Tipo de Registro
Const clRamoContPP      As String = "76"    '10-RAMO CONTABLE
Const clFrecPagPP       As String = "0"     '13-FRECUENCIA DE PAGO
Const clReaPP           As String = "2"     '23-REASEGURADOR
Const clTipMovSinLiqPP  As String = "35"    '61-TIPO MOVIMIENTO SINIESTRO (Liquido)
Const clTipMovSinSalPP  As String = "36"    '61-TIPO MOVIMIENTO SINIESTRO (Salud)
Const clTipMovSinResPP  As String = "1"    '61-TIPO MOVIMIENTO SINIESTRO (Resumen)
Const clTipPerNatPP     As String = "N"     'Tipo de persona (jurídico / natural)
Const clTipPerJurPP     As String = "J"     'Tipo de persona (jurídico / natural)
'------------------------------------------------------------------------------

Dim vlVar1  As String       'Tipo de Registro
Dim vlVar2  As String       'POLIZA
Dim vlVar3  As String       '--SUCURSAL
Dim vlVar4  As String       'VIGENCIA "DESDE" DE LA POLIZA
Dim vlVar5  As String       '--VIGENCIA "HASTA" DE LA POLIZA
Dim vlVar6  As String       '--VIGENCIA 'DESDE' ORIGINAL
Dim vlVar7  As String       'FECHA CONTABLE(MES / ANO)
Dim vlVar8  As String       'MONEDA DEL MOVIMIENTO
Dim vlVar9  As String       'COBERTURA
Dim vlVar10 As String       'RAMO CONTABLE
Dim vlVar11 As String       'CONTRATANTE RUT
Dim vlVar12 As String       'CONTRATANTE NOMBRE
Dim vlVar13 As String       'FRECUENCIA DE PAGO (INMEDIATA)
Dim vlVar14 As String       '--INTERMERDIARIO GTO DE COBRANZA
Dim vlVar15 As String       '--REGISTRO EN TRATAMIENTO
Dim vlVar16 As String       '--FECHA EFECTO REGISTRO (MOV)
Dim vlVar17 As String       'NOMBRE
Dim vlVar18 As String       'RUT
Dim vlVar19_Pen As String       'SUCURSAL
Dim vlVar19_Sal As String       'SUCURSAL
Dim vlVar20 As String       '--NOMBRE INTERMEDIARIO
Dim vlVar21 As String       '--RUT INNTERMEDIARIO
Dim vlVar22 As String       'TIPO DE INTERMEDIARIO
Dim vlVar23 As String       '--REASEGURADOR
Dim vlVar24 As String       'NACIONALIDAD REASEGURADOR
Dim vlVar25 As String       'CONTRATO DE REASEGURO
Dim vlVar26 As String       'TIPO DE REASEGURO
Dim vlVar27 As String       'NUMERO DE SINIESTRO
Dim vlVar28 As String       'ESTADO DEL SINIESTRO
Dim vlVar29 As String       'TIPO MOVIMIENTO
Dim vlVar30 As String       'NUMERO DE MOVIMIENTO
Dim vlVar31 As String       'MONTO EXENTO PRIMA
Dim vlVar32 As String       '--MONTO AFECTO PRIMA
Dim vlVar33 As String       '--MONTO IGV PRIMA
Dim vlVar34 As String       '--MONTO BRUTO PRIMA
Dim vlVar35 As String       '--MONTO NETO PRIMA DEVENGADA
Dim vlVar36 As String       '--CAPITALES ASEGURADOS
Dim vlVar37 As String       '--ORIGEN DEL RECIBO
Dim vlVar38 As String       '--TIPO DE MOVIMIENTO
Dim vlVar39 As String       '--MONTO PRIMA CEDIDA ANTES DSCTO
Dim vlVar40 As String       '--MONTO DESC. POR PRIMA CEDIDA
Dim vlVar41 As String       '--MONTO IMPUESTO 2%
Dim vlVar42 As String       '--MONTO EXCESO DE PERDIDA
Dim vlVar43 As String       '--CAPITALES CEDIDOS
Dim vlVar44 As String       '--TIPO DE RESERVA
Dim vlVar45 As String       '--MONTO RESERVA MATEMATICA
Dim vlVar47 As String       '-- % DE COMSION SOBRE LA PRIMA
Dim vlVar48 As String       '--TIPO DE COMISION
Dim vlVar49 As String       '--MONTO COMISION NETA
Dim vlVar50 As String       '--MONTO IGV COMISION
Dim vlVar51 As String       '--MONTO BRUTO COMISION
Dim vlVar52 As String       '--PERIODO DE GRACIA
Dim vlVar53 As String       '--MONTO NETO COMISION
Dim vlVar54 As String       '--ESQUEMA DE PAGO
Dim vlVar55 As String       '--fecha desde
Dim vlVar56 As String       '--fecha Hasta
Dim vlVar57 As String       '--RAMO
Dim vlVar58 As String       '--PRODUCTO
Dim vlVar59 As String       '--POLIZA
Dim vlVar60 As String       '--RUT DEL CLIENTE
Dim vlVar61 As String       '--TIPO DE MOVIMIENTO SINIESTRO
Dim vlVar62 As String       '--MONTO
Dim vlVar63 As String       '--MONTO CEDIDO EN EL MES
Dim vlVar64 As String       '-- % DE COMISION DE GASTOS DE COB
Dim vlVar65 As String       '--MTO. GASTOS DE COB. PRIMA REC.
Dim vlVar66 As String       '--MTO. GASTOS DE COB. PRIMA DEV.
Dim vlVar67 As String       '--Tipo de persona (jurídico / natural)
'-------------------------------------------------------------------------------
'Numero de archivo creado
Dim vlNumArchivo        As Integer

'Variables de Prima Unica
Dim vlNumCasosPriUni    As Long
Dim vlMtoPrimas         As Double

'Variables de Primeros Pagos
Dim vlNumCasosPPPension As Long
Dim vlNumCasosPPSalud   As Long
Dim vlMtoPPPension      As Double
Dim vlMtoPPSalud        As Double
Dim rs As ADODB.Recordset


Function flLmpGrillaPriUnica()

    Msf_Periodos.Clear
    Msf_Periodos.rows = 1
    Msf_Periodos.Cols = 8
    Msf_Periodos.RowHeight(0) = 250
    Msf_Periodos.Row = 0
    Msf_Periodos.ColWidth(0) = 0
    Msf_Periodos.Col = 1
    Msf_Periodos.Text = "Desde"
    Msf_Periodos.ColWidth(1) = 1100
    Msf_Periodos.Col = 2
    Msf_Periodos.Text = "Hasta"
    Msf_Periodos.ColWidth(2) = 1100
    Msf_Periodos.Col = 3
    Msf_Periodos.Text = "Nº Casos"
    Msf_Periodos.ColWidth(3) = 1200
    Msf_Periodos.Col = 4
    Msf_Periodos.Text = "Usuario"
    Msf_Periodos.ColWidth(4) = 1500
    Msf_Periodos.Col = 5
    Msf_Periodos.Text = "Fecha"
    Msf_Periodos.ColWidth(5) = 1200
    Msf_Periodos.Col = 6
    Msf_Periodos.Text = "Hora"
    Msf_Periodos.ColWidth(6) = 1200
    Msf_Periodos.Col = 7
    Msf_Periodos.Text = "Nº Archivo"
    Msf_Periodos.ColWidth(7) = 1000
    
End Function

Function flLmpGrillaPriPago()

    Msf_Periodos.Clear
    Msf_Periodos.rows = 1
    Msf_Periodos.Cols = 9
    Msf_Periodos.RowHeight(0) = 250
    Msf_Periodos.Row = 0
    Msf_Periodos.ColWidth(0) = 0
    Msf_Periodos.Col = 1
    Msf_Periodos.Text = "Desde"
    Msf_Periodos.ColWidth(1) = 1100
    Msf_Periodos.Col = 2
    Msf_Periodos.Text = "Hasta"
    Msf_Periodos.ColWidth(2) = 1100
    Msf_Periodos.Col = 3
    Msf_Periodos.Text = "Nº Casos Pensión"
    Msf_Periodos.ColWidth(3) = 1500
    Msf_Periodos.Col = 4
    Msf_Periodos.Text = "Nº Casos Salud"
    Msf_Periodos.ColWidth(4) = 1200
    Msf_Periodos.Col = 5
    Msf_Periodos.Text = "Usuario"
    Msf_Periodos.ColWidth(5) = 1500
    Msf_Periodos.Col = 6
    Msf_Periodos.Text = "Fecha"
    Msf_Periodos.ColWidth(6) = 1200
    Msf_Periodos.Col = 7
    Msf_Periodos.Text = "Hora"
    Msf_Periodos.ColWidth(7) = 1200
    Msf_Periodos.Col = 8
    Msf_Periodos.Text = "Nº Archivo"
    Msf_Periodos.ColWidth(8) = 1000
    
End Function

Function flActGrillaPriUnica()

vgQuery = ""
vgQuery = "SELECT num_archivo,fec_desde,fec_hasta,sum(num_casos) as num_casos,"
vgQuery = vgQuery & "cod_usuariocrea,fec_crea,hor_crea "
vgQuery = vgQuery & "FROM PD_TMAE_CONTABLEPRIUNI "
vgQuery = vgQuery & "GROUP BY num_archivo,fec_desde,fec_hasta,"
vgQuery = vgQuery & "cod_usuariocrea,fec_crea,hor_crea "
vgQuery = vgQuery & "ORDER BY fec_desde desc,fec_hasta desc,fec_crea desc,num_archivo desc "
Set vgRs = vgConexionBD.Execute(vgQuery)
If Not (vgRs.EOF) Then
    vgI = 1
    While Not (vgRs.EOF)
        Msf_Periodos.AddItem (vgI)
        Msf_Periodos.Row = vgI
        
        Msf_Periodos.Col = 1
        Msf_Periodos.Text = DateSerial(Mid(vgRs!Fec_desde, 1, 4), Mid(vgRs!Fec_desde, 5, 2), Mid(vgRs!Fec_desde, 7, 2))
        Msf_Periodos.Col = 2
        Msf_Periodos.Text = DateSerial(Mid(vgRs!fec_hasta, 1, 4), Mid(vgRs!fec_hasta, 5, 2), Mid(vgRs!fec_hasta, 7, 2))
        Msf_Periodos.Col = 3
        Msf_Periodos.Text = Format(vgRs!num_casos, "#,#")
        Msf_Periodos.Col = 4
        Msf_Periodos.Text = Trim(vgRs!Cod_UsuarioCrea)
        Msf_Periodos.Col = 5
        Msf_Periodos.Text = DateSerial(Mid(vgRs!Fec_Crea, 1, 4), Mid(vgRs!Fec_Crea, 5, 2), Mid(vgRs!Fec_Crea, 7, 2))
        Msf_Periodos.Col = 6
        Msf_Periodos.Text = Trim(Mid(vgRs!Hor_Crea, 1, 2) & ":" & Mid(vgRs!Hor_Crea, 3, 2) & ":" & Mid(vgRs!Hor_Crea, 5, 2))
        Msf_Periodos.Col = 7
        Msf_Periodos.Text = Trim(vgRs!Num_Archivo)
        
        vgI = vgI + 1
        vgRs.MoveNext
    Wend
End If
vgRs.Close

End Function

Function flActGrillaPriPago()

vgQuery = ""
vgQuery = "SELECT distinct num_archivo,fec_desde,fec_hasta,"
vgQuery = vgQuery & "(select sum(num_casos) from pd_tmae_contablepripago where "
vgQuery = vgQuery & "num_archivo=p.num_archivo and cod_tipmov='" & clTipMovSinLiqPP & "')as num_casospension,"
vgQuery = vgQuery & "(select sum(num_casos) from pd_tmae_contablepripago where "
vgQuery = vgQuery & "num_archivo=p.num_archivo and cod_tipmov='" & clTipMovSinSalPP & "')as num_casossalud "
vgQuery = vgQuery & ",cod_usuariocrea,fec_crea,hor_crea "
vgQuery = vgQuery & "FROM PD_TMAE_CONTABLEPRIPAGO P "
vgQuery = vgQuery & "ORDER BY fec_desde desc,fec_hasta desc,fec_crea desc,hor_crea desc "
Set vgRs = vgConexionBD.Execute(vgQuery)
If Not (vgRs.EOF) Then
    vgI = 1
    While Not (vgRs.EOF)
        Msf_Periodos.AddItem (vgI)
        Msf_Periodos.Row = vgI
        
        Msf_Periodos.Col = 1
        Msf_Periodos.Text = DateSerial(Mid(vgRs!Fec_desde, 1, 4), Mid(vgRs!Fec_desde, 5, 2), Mid(vgRs!Fec_desde, 7, 2))
        Msf_Periodos.Col = 2
        Msf_Periodos.Text = DateSerial(Mid(vgRs!fec_hasta, 1, 4), Mid(vgRs!fec_hasta, 5, 2), Mid(vgRs!fec_hasta, 7, 2))
        Msf_Periodos.Col = 3
        Msf_Periodos.Text = Format(vgRs!num_casospension, "#,#0")
        Msf_Periodos.Col = 4
        Msf_Periodos.Text = Format(vgRs!num_casossalud, "#,#0")
        Msf_Periodos.Col = 5
        Msf_Periodos.Text = Trim(vgRs!Cod_UsuarioCrea)
        Msf_Periodos.Col = 6
        Msf_Periodos.Text = DateSerial(Mid(vgRs!Fec_Crea, 1, 4), Mid(vgRs!Fec_Crea, 5, 2), Mid(vgRs!Fec_Crea, 7, 2))
        Msf_Periodos.Col = 7
        Msf_Periodos.Text = Trim(Mid(vgRs!Hor_Crea, 1, 2) & ":" & Mid(vgRs!Hor_Crea, 3, 2) & ":" & Mid(vgRs!Hor_Crea, 5, 2))
        Msf_Periodos.Col = 8
        Msf_Periodos.Text = Trim(vgRs!Num_Archivo)
        
        vgI = vgI + 1
        vgRs.MoveNext
    Wend
End If
vgRs.Close

End Function

Sub flImpresion()
Dim vlArchivo As String

Err.Clear
On Error GoTo Errores1
   
    Screen.MousePointer = 11
    
    If (Trim(Lbl_NumArchivo) = "") Then
        MsgBox "Debe seleccionar un Periodo a Imprimir.", vbInformation, "Falta Información"
        Screen.MousePointer = 0
        Exit Sub
    End If
        
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
   
    Dim objRep As New ClsReporte
    Dim LNGa As Long
    If (vgNomForm = "PriUni") Then
    
        'Sql = "SELECT * FROM PD_TMAE_CONTABLEPRIUNI where num_archivo=" & Trim(Lbl_NumArchivo)
                
        'Sql = "SELECT C.NUM_ARCHIVO,C.NUM_CASOS,C.COD_MONEDA,C.COD_TIPMOV,C.COD_TIPREG,C.MTO_PRIMAS,C.FEC_DESDE,C.FEC_HASTA,C.HOR_CREA,C.COD_USUARIOCREA"
        'Sql = Sql & ",sum(Decode(p.ind_estsun,3,1,0)) as NUM_CORRECTOS,Sum(Decode(p.ind_estsun,null,1,2,1,0)) as NUM_INCORRECTOS "
        'Sql = Sql & "FROM PD_TMAE_CONTABLEPRIUNI C "
        'Sql = Sql & "inner join PD_TMAE_CONTABLEDETPRIUNI D "
        'Sql = Sql & "on C.NUM_ARCHIVO=D.NUM_ARCHIVO AND C.COD_TIPREG=D.COD_TIPREG AND C.COD_TIPMOV=D.COD_TIPMOV AND C.COD_MONEDA=D.COD_MONEDA "
        'Sql = Sql & "LEFT JOIN PD_TMAE_POLIZA P ON D.NUM_POLIZA=P.NUM_POLIZA and P.NUM_ENDOSO=1 "
        'Sql = Sql & "Where C.num_archivo=" & Trim(Lbl_NumArchivo) & " "
        'Sql = Sql & "GROUP BY C.NUM_ARCHIVO, C.NUM_CASOS, C.COD_MONEDA, C.COD_TIPMOV, C.COD_TIPREG, C.MTO_PRIMAS, C.FEC_DESDE, C.FEC_HASTA, C.HOR_CREA, C.COD_USUARIOCREA"
        
    
        rs.Open "PD_LISTA_ARCH_CONT_PRI_UNICA.LISTAR_RESUMEN('" & Lbl_NumArchivo.Caption & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
        LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PD_Rpt_ContableResPriUnica.rpt"), ".RPT", ".TTX"), 1)
        If objRep.CargaReporte(strRpt & "", "PD_Rpt_ContableResPriUnica.rpt", "Informe Detalle Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
            
            MsgBox "No se pudo abrir el reporte", vbInformation
            Exit Sub
        End If
    Else
        Sql = "SELECT * FROM PD_TMAE_CONTABLEPRIPAGO where num_archivo=" & Trim(Lbl_NumArchivo)
    
        '"PD_LISTA_ARCH_CONT_PAGOS.LISTAR('" & Lbl_NumArchivo.Caption & "')"
    
        rs.Open Sql, vgConexionBD, adOpenForwardOnly, adLockReadOnly
        LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PD_Rpt_ContableResPriPagos.rpt"), ".RPT", ".TTX"), 1)
        If objRep.CargaReporte(strRpt & "", "PD_Rpt_ContableResPriPagos.rpt", "Informe Detalle Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
            
            MsgBox "No se pudo abrir el reporte", vbInformation
            Exit Sub
        End If
    End If
        
  '  If (vgNomForm = "PriUni") Then
  '      vlArchivo = strRpt & "PD_Rpt_ContableResPriUnica.rpt"   '\Reportes
  '      vgQuery = "{PD_TMAE_CONTABLEPRIUNI.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
  '  Else
  '      vlArchivo = strRpt & "PD_Rpt_ContableResPriPagos.rpt"   '\Reportes
  '      vgQuery = "{PD_TMAE_CONTABLEPRIPAGO.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
  '  End If
  '
  '  If Not fgExiste(vlArchivo) Then     ', vbNormal
  '      MsgBox "Archivo de Reporte de Resumen de Archivo Contable no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
  '      Screen.MousePointer = 0
  '      Exit Sub
  '  End If
  '
  '  Rpt_General.Reset
  '  Rpt_General.WindowState = crptMaximized
  '  Rpt_General.ReportFileName = vlArchivo
  '  Rpt_General.Connect = vgRutaDataBase
  '  Rpt_General.Destination = crptToWindow
  '  Rpt_General.SelectionFormula = ""
  '  Rpt_General.SelectionFormula = vgQuery
  '
  '  Rpt_General.Formulas(0) = ""
  '  Rpt_General.Formulas(1) = ""
  '  Rpt_General.Formulas(2) = ""
  '  Rpt_General.Formulas(3) = ""
  '
  '  Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
  '  Rpt_General.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
  '  Rpt_General.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
  '
  '  Rpt_General.WindowTitle = "Informe Resumen Archivo Contable"
  '  Rpt_General.Action = 1
    
    Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

'FIN DE LAS FUNCIONES Y PROCEDIMIENTOS CREADOS
'-----------------------------------------------------

Private Sub Cmd_Cargar_Click()
On Error GoTo Err_Cargar
    
    Lbl_FecCierre = Format(Now, "dd/mm/yyyy Hh:Nn:Ss AMPM")
    
    'Validación de Datos
    'Periodo Desde
    If Txt_Desde = "" Then
       MsgBox "Debe Ingresar Fecha Desde.", vbInformation, "Falta Información"
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_Desde) Then
        MsgBox "La Fecha Desde ingresada no es válida.", vbInformation, "Error de Datos"
        Screen.MousePointer = 0
        Txt_Desde.Text = ""
        Exit Sub
    End If
    
    'Periodo Hasta
    If Txt_Hasta = "" Then
       MsgBox "Debe Ingresar Fecha Hasta.", vbInformation, "Falta Información"
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_Hasta) Then
        MsgBox "La Fecha Hasta ingresada no es válida.", vbInformation, "Error de Datos"
        Screen.MousePointer = 0
        Txt_Hasta.Text = ""
        Exit Sub
    End If
    
    If CDate(Txt_Desde) > CDate(Txt_Hasta) Then
       MsgBox "La Fecha Desde debe ser menor o igual a la Fecha Hasta.", vbInformation, "Error de Datos"
       Txt_Desde.SetFocus
       Exit Sub
    End If

    If LblArchivo = "" Then
        MsgBox "Debe seleccionar Archivo a generar.", vbInformation, "Falta Información"
        Screen.MousePointer = 0
        CmdContable.SetFocus
        Exit Sub
    End If
    
    vlFecDesde = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFecHasta = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    
    vgSw = True
    
    vgRes = MsgBox(" ¿ Está Seguro que desea 'Generar' el Archivo Contable para el Período definido ? ", 4 + 32 + 256, "Archivo Contable")
    If vgRes = 6 Then
        vgSw = False
    Else
        Exit Sub
    End If
    
    Screen.MousePointer = 11
                    
    If (vgNomForm = "PriUni") Then
        'Genera el archivo de Prima Unica
        vlArchPriUni = flExportarPriUnica(vlFecDesde, vlFecHasta)
        If (vlArchPriUni = False) Then
            MsgBox "Se Ha Producido un error durante el proceso de Generación del Archivo Contable", vbCritical, "Proceso Cancelado"
            Exit Sub
        Else
            'Actualizar la Grilla de Datos
            Call flLmpGrillaPriUnica
            Call flActGrillaPriUnica
            Screen.MousePointer = 0
            MsgBox "El Proceso ha finalizado Exitosamente.", vbInformation, "Proceso Generación."
        End If
    Else
        'Genera el archivo de Primeros Pagos
        vlArchPriPagos = flExportarPriPagos(vlFecDesde, vlFecHasta)
        If (vlArchPriPagos = False) Then
            MsgBox "Se Ha Producido un error durante el proceso de Generación del Archivo Contable", vbCritical, "Proceso Cancelado"
            Exit Sub
        Else
            'Actualizar la Grilla de Datos
            Call flLmpGrillaPriPago
            Call flActGrillaPriPago
            Screen.MousePointer = 0
            MsgBox "El Proceso ha finalizado Exitosamente.", vbInformation, "Proceso Generación."
        End If
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

Private Sub flImpDetalle()
Dim vlArchivo As String
'Dim objRep As New ClsReporte
Dim vlNomRep As String

Err.Clear
On Error GoTo Errores1
   
    Screen.MousePointer = 11
    
    If (Trim(Lbl_NumArchivo) = "") Then
        MsgBox "Debe seleccionar un Periodo a Imprimir.", vbInformation, "Falta Información"
        Screen.MousePointer = 0
        Exit Sub
    End If
        
'    If (vgNomForm = "PriUni") Then
'        vlArchivo = strRpt & ""   '\Reportes
'        vlNomRep = "PD_Rpt_ContableDetPriUnica.rpt"
'        'vgQuery = "{PD_TMAE_CONTABLEDETPRIUNI.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
'    Else
'        vlArchivo = strRpt & ""  '\Reportes
'        vlNomRep = "PD_Rpt_ContableDetPriPagos.rpt"
'        'vgQuery = "{PD_TMAE_CONTABLEDETPRIPAGO.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo)
'    End If
'
'    If Not fgExiste(vlArchivo) Then     ', vbNormal
'        MsgBox "Archivo de Reporte de Detalle de Archivo Contable no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'
'    Rpt_General.Reset
'    Rpt_General.WindowState = crptMaximized
'    Rpt_General.ReportFileName = vlArchivo
'    Rpt_General.Connect = vgRutaDataBase
'    Rpt_General.Destination = crptToWindow
'    Rpt_General.SelectionFormula = ""
'    Rpt_General.SelectionFormula = vgQuery
'
'    Rpt_General.Formulas(0) = ""
'    Rpt_General.Formulas(1) = ""
'    Rpt_General.Formulas(2) = ""
'    Rpt_General.Formulas(3) = ""
'
'    Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
'    Rpt_General.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
'    Rpt_General.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
'
'    Rpt_General.WindowTitle = "Informe Detalle Archivo Contable"
'    Rpt_General.Action = 1

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Dim objRep As New ClsReporte
    
    Dim LNGa As Long
    If (vgNomForm = "PriUni") Then
        rs.Open "PD_LISTA_ARCH_CONT_PRI_UNICA.LISTAR('" & Lbl_NumArchivo.Caption & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
        LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PD_Rpt_ContableDetPriUnica.rpt"), ".RPT", ".TTX"), 1)
        If objRep.CargaReporte(strRpt & "", "PD_Rpt_ContableDetPriUnica.rpt", "Informe Detalle Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
            
            MsgBox "No se pudo abrir el reporte", vbInformation
            Exit Sub
        End If
    Else
        rs.Open "PD_LISTA_ARCH_CONT_PAGOS.LISTAR('" & Lbl_NumArchivo.Caption & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
        LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PD_Rpt_ContableDetPriPagos.rpt"), ".RPT", ".TTX"), 1)
        If objRep.CargaReporte(strRpt & "", "PD_Rpt_ContableDetPriPagos.rpt", "Informe Detalle Archivo Contable", rs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
            
            MsgBox "No se pudo abrir el reporte", vbInformation
            Exit Sub
        End If
    End If
    
    
    
    Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir

    'Imprime el Reporte de Resumen
    If (Opt_Resumen) Then
        flImpresion
    End If
    'Imprime el Reporte de Detalle por Movimiento
    If (Opt_DetMov) Then
        flImpDetalle
    End If
    'Imprime el Reporte de Detalle por Pensionado
    If (Opt_DetPen) Then
        flImpDetallePen
    End If
    
Exit Sub
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_Limpiar
    
    Lbl_FecCierre = Format(Now, "dd/mm/yyyy Hh:Nn:Ss AMPM")
    Txt_Desde = ""
    Txt_Hasta = ""
    LblArchivo.Caption = ""
    Lbl_NumArchivo.Caption = ""
    
    If (vgNomForm = "PriUni") Then
        Call flLmpGrillaPriUnica
        Call flActGrillaPriUnica
    Else
        Call flLmpGrillaPriPago
        Call flActGrillaPriPago
    End If
    Opt_Resumen.Value = True
    
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

Private Sub CmdContable_Click()
Dim ilargo As Long
Dim iTexto As Long
Dim fecdesde As String, fechasta As String
On Error GoTo Err_Carga
    
    If Not IsDate(Txt_Desde) Then
        MsgBox "Debe Ingresar la Fecha Desde.", vbInformation, "Falta Información"
        Txt_Desde.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Txt_Hasta) Then
        MsgBox "Debe Ingresar la Fecha Hasta.", vbInformation, "Falta Información"
        Txt_Hasta.SetFocus
        Exit Sub
    End If
    
    If CDate(Txt_Desde) > CDate(Txt_Hasta) Then
       MsgBox "La Fecha Desde debe ser menor o igual a la Fecha Hasta.", vbInformation, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    
    fecdesde = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    fechasta = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    
    'Selección del Archivo del cual se generará el archivo contable
    vlArchivoCont = ""
    ComDialogo.CancelError = True
    ComDialogo.FileName = "SEA_" & vgNomForm & "_" & fecdesde & "_" & fechasta & ".txt" '17/06/2008
    ComDialogo.DialogTitle = "Archivo Contable"
    ComDialogo.Filter = "*.txt"
    ComDialogo.FilterIndex = 1
    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    ComDialogo.ShowSave
    vlArchivoCont = ComDialogo.FileName
    LblArchivo.Caption = vlArchivoCont
    If (Len(vlArchivoCont) > 65) Then
        While Len(LblArchivo) > 65
            ilargo = InStr(1, LblArchivo, "\")
            LblArchivo = Mid(LblArchivo, ilargo + 1, Len(LblArchivo))
        Wend
        LblArchivo.Caption = "\\" & LblArchivo
    End If
    If vlArchivoCont = "" Then
        Exit Sub
    End If
    
Exit Sub
Err_Carga:
    If (Err.Number = 32755) Then
        Exit Sub
    End If
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar
  
    Me.Top = 0
    Me.Left = 0
    Lbl_FecCierre = Format(Now, "dd/mm/yyyy Hh:Nn:Ss AMPM")
    If (vgNomForm = "PriUni") Then
        Fra_Datos(0).Caption = " Fecha de Traspaso Prima"
        Frame3.Visible = True
        Fra_Datos(1).Top = 5160
        Frm_ContableArch.Height = 6885
        Call flOcultarOption
        Call flLmpGrillaPriUnica
        Call flActGrillaPriUnica
        Me.Caption = "Generación de Archivo Contable de Prima Unica."
    Else
        Fra_Datos(0).Caption = " Fecha de Pago"
        Frame3.Visible = False
        Fra_Datos(1).Top = 4440
        Frm_ContableArch.Height = 6180
        Call flMostrarOption
        Call flLmpGrillaPriPago
        Call flActGrillaPriPago
        Me.Caption = "Generación de Archivo Contable de Primeros Pagos."
    End If
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_Periodos_Click()
On Error GoTo Err_Periodo

    If (Msf_Periodos.Text = "") Or (Msf_Periodos.Row = 0) Then
        Exit Sub
    End If
    
    vgI = Msf_Periodos.Row
    
    Txt_Desde = Msf_Periodos.TextMatrix(vgI, 1)
    Txt_Hasta = Msf_Periodos.TextMatrix(vgI, 2)
    
    If (vgNomForm = "PriUni") Then
        Lbl_NumArchivo = Msf_Periodos.TextMatrix(vgI, 7)
    Else
        Lbl_NumArchivo = Msf_Periodos.TextMatrix(vgI, 8)
    End If
        
Exit Sub
Err_Periodo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
 
 '----- RESCATA EL NUMERO DE ENTRADA CORRESPONDIENTE AL ARCHIVO DE PRIMA UNICA -----
Private Function flNumArchivoPriUnica() As Integer
    vgQuery = ""
    vgQuery = "SELECT NUM_ARCHIVO FROM PD_TMAE_CONTABLEPRIUNI "
    vgQuery = vgQuery & " ORDER BY NUM_ARCHIVO DESC"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        flNumArchivoPriUnica = CInt(vgRs!Num_Archivo) + 1
    Else
        flNumArchivoPriUnica = 1
    End If
End Function

 '----- RESCATA EL NUMERO DE ENTRADA CORRESPONDIENTE AL ARCHIVO DE PRIMEROS PAGOS -----
Private Function flNumArchivoPriPago() As Integer
    vgQuery = ""
    vgQuery = "SELECT NUM_ARCHIVO FROM PD_TMAE_CONTABLEPRIPAGO "
    vgQuery = vgQuery & " ORDER BY NUM_ARCHIVO DESC"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        flNumArchivoPriPago = CInt(vgRs!Num_Archivo) + 1
    Else
        flNumArchivoPriPago = 1
    End If
End Function

Private Function flExportarPriUnica(iFecDesde As String, iFecHasta As String) As Boolean
Dim vlLinea As String
Dim vlArchivo As String, vlOpen As Boolean
Dim vlVarRutCont As String, vlVarRut As String
Dim vlMtoPri As Double, vlFecTrasp As String
Dim vlafp As String, vlCodPen As String
Dim vlNumOrd As Integer
On Error GoTo Errores

flExportarPriUnica = False

    vlArchivo = vlArchivoCont 'LblArchivo
    
    Screen.MousePointer = 11
    
    Open vlArchivo For Output As #1
    Me.Refresh
    vlOpen = True
        
    Call flInicializaVar
    
    'Obtiene el nº de archivo a crear
    vlNumArchivo = flNumArchivoPriUnica
    vlFecCrea = Format(Date, "yyyymmdd")
    vlHorCrea = Format(Time, "hhmmss")
    
''    vlSql = "SELECT COD_ELEMENTO FROM MA_TPAR_TABCOD "
''    vlSql = vlSql & "WHERE COD_TABLA='" & vgCodTabla_TipMon & "' "
''    vlSql = vlSql & "AND (COD_SISTEMA IS NULL OR COD_SISTEMA<>'PP') "
''    vlSql = vlSql & "ORDER BY COD_ELEMENTO "
''    Set vgRs1 = vgConexionBD.Execute(vlSql)
''    While Not (vgRs1.EOF)
''        vlMoneda = Trim(vgRs1!cod_elemento)
        
        Call flEstadisticaPriUnica
        
        vlNumCasosPriUni = 0
        vlMtoPrimas = 0
        
        Sql = "SELECT r.num_poliza,p.fec_vigencia,r.fec_traspaso,r.cod_monedapriinf,"
        Sql = Sql & "p.cod_tippension,p.cod_tipoidenafi,p.num_idenafi,b.gls_nomben,"
        Sql = Sql & "b.gls_nomsegben,b.gls_patben,b.gls_matben,p.cod_afp,p.gls_direccion,"
        Sql = Sql & "r.cod_liquidacion,r.mto_prirec,cod_renvit,b.num_orden "
        Sql = Sql & "FROM pd_tmae_polprirec r,pd_tmae_poliza p,pd_tmae_polben b "
        Sql = Sql & "WHERE r.fec_traspaso between '" & iFecDesde & "' and '" & iFecHasta & "' "
        ''Sql = Sql & "AND p.cod_moneda='" & vlMoneda & "' "
        Sql = Sql & "AND p.num_poliza=r.num_poliza "
        Sql = Sql & "and p.num_poliza=b.num_poliza "
        Sql = Sql & "and p.num_endoso=b.num_endoso "
        Sql = Sql & "and b.cod_par='99' "
        Sql = Sql & "and p.num_endoso=1 "
        
        If (Opt_Recepcionadas) Then 'Filtra solo los recepcionados
        Sql = Sql & "and p.IND_ESTSUN=3 "
        End If
        
        If (Opt_Todas) Then 'Filtra los recepcionados, rechazados y no enviados
        Sql = Sql & "and (p.IND_ESTSUN=3 or p.IND_ESTSUN=2 or p.IND_ESTSUN is null) "
        End If
                        
        Sql = Sql & "ORDER BY r.num_poliza "
        Set vgRs = vgConexionBD.Execute(Sql)
        While Not (vgRs.EOF)
            
            vlNumOrd = CInt(vgRs!Num_Orden)
            vlFecTrasp = Trim(vgRs!fec_traspaso)
            vlVar1 = Format(Trim(clTipRegPU), "00000")
            vlVar2 = Format(Trim(vgRs!Num_Poliza), "0000000000")
            vlVar3 = Format(Trim(clSucPU), "00000")
            If Trim(vgRs!fec_traspaso) <> "" Then
                vlVar4 = DateSerial(Mid(vgRs!fec_traspaso, 1, 4), Mid(vgRs!fec_traspaso, 5, 2), Mid(vgRs!fec_traspaso, 7, 2))
            Else
                vlVar4 = Space(10)
            End If
            If (Len(Trim(vgRs!cod_monedapriinf)) <= 5) Then
                vlVar8 = Trim(vgRs!cod_monedapriinf) & Space(5 - Len(Trim(vgRs!cod_monedapriinf)))
            Else
                vlVar8 = Mid(Trim(vgRs!cod_monedapriinf), 1, 5)
            End If
            vlMoneda = Trim(vlVar8)
            vlCodPen = Trim(vgRs!Cod_TipPension)
            vlVar9 = Format(Trim(vgRs!Cod_TipPension), "00000")
            'vlVar10 = Format(Trim(clRamoContPU), "00000")
            Select Case vgRs!Cod_TipPension
             Case "04"
                vlVar10 = Format(Trim("76"), "00000")
             Case "05"
                vlVar10 = Format(Trim("76"), "00000")
             Case "06"
                vlVar10 = Format(Trim("94"), "00000")
             Case "07"
                vlVar10 = Format(Trim("94"), "00000")
             Case "08"
                vlVar10 = Format(Trim("95"), "00000")
            End Select
            
            If (Len(Trim(vgRs!num_idenafi)) <= 12) Then
                vlVar11 = Format(Trim(vgRs!cod_tipoidenafi), "00") & (Trim(vgRs!num_idenafi) & Space(12 - Len(Trim(vgRs!num_idenafi))))
            Else
                vlVar11 = Format(Trim(vgRs!cod_tipoidenafi), "00") & Mid(Trim(vgRs!num_idenafi), 1, 12)
            End If
            vlVar12 = fgFormarNombreCompleto(IIf(IsNull(vgRs!Gls_NomBen), "", Trim(vgRs!Gls_NomBen)), IIf(IsNull(vgRs!Gls_NomSegBen), "", Trim(vgRs!Gls_NomSegBen)), IIf(IsNull(vgRs!Gls_PatBen), "", Trim(vgRs!Gls_PatBen)), IIf(IsNull(vgRs!Gls_MatBen), "", Trim(vgRs!Gls_MatBen)))
            If (Len(Trim(vlVar12)) <= 60) Then
                vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
            Else
                vlVar12 = Mid(vlVar12, 1, 60)
            End If
            vlVar13 = Format(Trim(clFrecPagPU), "00000")
            vlVar15 = Format(Trim(clRegTratFacPU), "00000")
            vlVar17 = fgFormarNombreCompleto(IIf(IsNull(vgRs!Gls_NomBen), "", Trim(vgRs!Gls_NomBen)), IIf(IsNull(vgRs!Gls_NomSegBen), "", Trim(vgRs!Gls_NomSegBen)), IIf(IsNull(vgRs!Gls_PatBen), "", Trim(vgRs!Gls_PatBen)), IIf(IsNull(vgRs!Gls_MatBen), "", Trim(vgRs!Gls_MatBen)))
            If (Len(Trim(vlVar17)) <= 60) Then
                vlVar17 = vlVar17 & Space(60 - Len(Trim(vlVar17)))
            Else
                vlVar17 = Mid(vlVar17, 1, 60)
            End If
            vlafp = Trim(vgRs!Cod_AFP)
            vlVar19_Pen = Format(Trim(vgRs!Cod_AFP), "00000")
            vlVar23 = Format(Trim(clReaPU), "0000000000")
            vlVar24 = Format(Trim(vgRs!cod_renvit), "B000") & "-" & Format(Trim(vgRs!cod_liquidacion), "00000000") 'AB-10/11/2007
            If (Len(Trim(vlVar24)) <= 40) Then
                vlVar24 = Trim(vlVar24) & Space(40 - Len(Trim(vlVar24)))
            Else
                vlVar24 = Mid(Trim(vlVar24), 1, 40)
            End If
            vlVar25 = vlVar9
            vlVar29 = Format(Trim(clTipMovPU), "00000")
            vlVar31 = IIf(IsNull(vgRs!MTO_PRIREC), 0, Format(vgRs!MTO_PRIREC, "#0.00"))
            vlVar31 = flFormatNum18_2(vlVar31)
            
            vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                      (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                      (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                      (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19_Pen) & (vlVar20) & _
                      (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                      (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                      (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                      (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                      (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                      (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                      (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                      (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                      (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                      (vlVar66) & (vlVar67)
    
            vlLinea = Replace(vlLinea, ",", ".")
            Print #1, vlLinea
            
            'Guarda el detalle de la póliza informada
            vlMtoPri = IIf(IsNull(vgRs!MTO_PRIREC), 0, Format(vgRs!MTO_PRIREC, "#0.00"))
            Call flGrabaDetPriUnica(vlVar2, vlNumOrd, vlFecTrasp, vlCodPen, vlafp, vlVar24, vlVar12, vlMtoPri)
            
            vlNumCasosPriUni = vlNumCasosPriUni + 1
            vlMtoPrimas = vlMtoPrimas + IIf(IsNull(vgRs!MTO_PRIREC), 0, (vgRs!MTO_PRIREC))
            
            vgRs.MoveNext
        Wend
        vgRs.Close
        
        'Actualiza la cantidad de casos y mtos informados por tipo moneda
        Call flActEstadisticaPriUnica
        
''        vgRs1.MoveNext
''    Wend
''    vgRs1.Close
    
Close #1
vlOpen = False

flExportarPriUnica = True

Exit Function
Errores:
Screen.MousePointer = vbDefault
If Err.Number <> 0 Then
    If vlOpen Then
        Close #1
''        Unload Frm_BarraProg
    End If
    MsgBox "Se ha producido el siguiente error : " & Err.Description, vbCritical, "Error"
End If
End Function

Private Function flExportarPriPagos(iFecDesde As String, iFecHasta As String) As Boolean
Dim vlLinea As String
Dim vlArchivo As String, vlOpen As Boolean
Dim vlRutContrat As String, vlNomContrat As String
Dim vlFecPago As String, vlNumPol As String
Dim vlMtoPen As Double, vlMtoSal As Double
Dim vlafp As String, vlCodPen As String
Dim vlNumOrd As Integer
Dim vlMtoPenTot As Double
On Error GoTo Errores

flExportarPriPagos = False

    vlArchivo = vlArchivoCont 'LblArchivo
    
    Screen.MousePointer = 11
    
    Open vlArchivo For Output As #1
    
    vlOpen = True
    
    Call flInicializaVar
    
    'Obtiene el nº de archivo a crear
    vlNumArchivo = flNumArchivoPriPago
    vlFecCrea = Format(Date, "yyyymmdd")
    vlHorCrea = Format(Time, "hhmmss")
    
    vlSql = "SELECT COD_ELEMENTO FROM MA_TPAR_TABCOD "
    vlSql = vlSql & "WHERE COD_TABLA='" & vgCodTabla_TipMon & "' "
    vlSql = vlSql & "AND (COD_SISTEMA IS NULL OR COD_SISTEMA<>'PP') "
    vlSql = vlSql & "ORDER BY COD_ELEMENTO "
    Set vgRs1 = vgConexionBD.Execute(vlSql)
    While Not (vgRs1.EOF)
        vlMoneda = Trim(vgRs1!cod_elemento)
        
        Call flEstadisticaPriPago(clTipMovSinLiqPP)
        Call flEstadisticaPriPago(clTipMovSinSalPP)
        
        vlNumCasosPPPension = 0
        vlNumCasosPPSalud = 0
        vlMtoPPPension = 0
        vlMtoPPSalud = 0
        
'        Sql = "SELECT l.num_poliza,l.cod_viapago,p.fec_vigencia,l.fec_pago,p.cod_moneda,"
'        Sql = Sql & "l.cod_tippension,p.cod_tipoidenafi,p.num_idenafi,b.gls_nomben,"
'        Sql = Sql & "b.gls_nomsegben,b.gls_patben,b.gls_matben,b.cod_tipoidenben,"
'        Sql = Sql & "b.num_idenben,p.cod_afp,l.gls_direccion,sum(l.mto_liqpagar) "
'        Sql = Sql & "as mto_liqpagar,sum(l.mto_plansalud) as mto_plansalud,"
'        Sql = Sql & "l.cod_inssalud,p.num_endoso,b.num_orden"
'        Sql = Sql & ",sum(l.mto_pension) as mto_pensiontot " '01/12/2007
'        Sql = Sql & "FROM pd_tmae_liqpagopen l,pd_tmae_poliza p, pd_tmae_polben b "
'        Sql = Sql & "WHERE fec_pago between '" & iFecDesde & "' and '" & iFecHasta & "' "
'        Sql = Sql & "AND p.num_poliza=l.num_poliza and b.num_poliza=l.num_poliza "
'        Sql = Sql & "and b.num_orden=l.num_orden "
'        Sql = Sql & "and p.cod_moneda ='" & vlMoneda & "' "
'        Sql = Sql & "GROUP BY l.num_poliza,l.cod_viapago,p.fec_vigencia,l.fec_pago,"
'        Sql = Sql & "p.cod_moneda,l.cod_tippension,p.cod_tipoidenafi,p.num_idenafi,"
'        Sql = Sql & "b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,"
'        Sql = Sql & "b.cod_tipoidenben,b.num_idenben,p.cod_afp,l.gls_direccion,"
'        Sql = Sql & "l.cod_inssalud,p.num_endoso,b.num_orden "
'        Sql = Sql & "ORDER BY l.num_poliza "
        vlSql = ""
        '???? CAMBIO SELEC 29/03/2010
'        vlSql = "SELECT l.num_poliza,l.cod_viapago,p.fec_vigencia,l.fec_pago,p.cod_moneda,l.cod_tippension,p.cod_tipoidenafi,p.num_idenafi,"
'        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.gls_nomtut else b.gls_nomben end)AS gls_nomben,"
'        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.gls_nomsegtut else b.gls_nomsegben end)AS gls_nomsegben,"
'        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.gls_pattut else b.gls_patben end)AS gls_patben,"
'        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.gls_mattut else b.gls_matben end)AS gls_matben,"
'        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.cod_tipoidentut else b.cod_tipoidenben end)AS cod_tipoidenben,"
'        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.num_identut else b.num_idenben end)AS num_idenben,"
'        vlSql = vlSql & " p.cod_afp,l.gls_direccion,sum(l.mto_liqpagar) as mto_liqpagar,sum(l.mto_plansalud) as mto_plansalud,l.cod_inssalud,"
'        vlSql = vlSql & " p.num_endoso,b.num_orden,sum(l.mto_pension) as mto_pensiontot"
'        vlSql = vlSql & " FROM pd_tmae_liqpagopen l"
'        vlSql = vlSql & ", pd_tmae_poliza p, pd_tmae_polben b, pd_tmae_poltutor t"
'        vlSql = vlSql & " where p.num_poliza =l.num_poliza "
'        vlSql = vlSql & " AND b.num_poliza = l.num_poliza AND b.num_orden = l.num_orden "
'        vlSql = vlSql & " AND p.num_endoso=b.num_endoso"
'        vlSql = vlSql & " and b.num_poliza =t.num_poliza  (+)"
'        vlSql = vlSql & " AND b.num_endoso  =t.num_endoso (+)"
'        vlSql = vlSql & " AND b.num_orden  =t.num_orden (+)"
'        vlSql = vlSql & " AND fec_pago between '" & iFecDesde & "' and '" & iFecHasta & "'  and p.cod_moneda ='" & vlMoneda & "'"
'        vlSql = vlSql & " AND p.num_endoso=1"
'        vlSql = vlSql & " GROUP BY l.num_poliza,l.cod_viapago,p.fec_vigencia,l.fec_pago,p.cod_moneda,l.cod_tippension,p.cod_tipoidenafi,"
'        vlSql = vlSql & " p.num_idenafi,b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,b.cod_tipoidenben,b.num_idenben,p.cod_afp,"
'        vlSql = vlSql & " l.gls_direccion,l.cod_inssalud,p.num_endoso,b.num_orden,"
'        vlSql = vlSql & " t.gls_nomtut,t.gls_nomsegtut,t.gls_pattut,t.gls_mattut,t.cod_tipoidentut,t.num_identut ORDER BY l.num_poliza"

            
            
        vlSql = "SELECT l.num_poliza,p.cod_viapago,p.fec_vigencia,l.fec_pago,p.cod_moneda,l.cod_tippension,p.cod_tipoidenafi,p.num_idenafi,"
        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.gls_nomtut else b.gls_nomben end)AS gls_nomben,"
        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.gls_nomsegtut else b.gls_nomsegben end)AS gls_nomsegben,"
        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.gls_pattut else b.gls_patben end)AS gls_patben,"
        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.gls_mattut else b.gls_matben end)AS gls_matben,"
        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.cod_tipoidentut else b.cod_tipoidenben end)AS cod_tipoidenben,"
        vlSql = vlSql & " (case when nvl(t.NUM_IDENTUT,' ')<>' ' then t.num_identut else b.num_idenben end)AS num_idenben, p.cod_afp,l.gls_direccion,sum(l.mto_liqpagar) as mto_liqpagar,sum(l.mto_plansalud) as mto_plansalud,"
        vlSql = vlSql & " l.cod_inssalud, p.num_endoso,b.num_orden,sum(l.mto_pension) as mto_pensiontot, p.cod_banco, p.num_cuenta, p.num_cuenta_cci, m.gls_elemento gls_banco"
        vlSql = vlSql & " FROM pd_tmae_liqpagopen l"
        vlSql = vlSql & " join pd_tmae_polben b on b.num_poliza = l.num_poliza and b.num_orden = l.num_orden"
        vlSql = vlSql & " join pd_tmae_poliza p on p.num_poliza =b.num_poliza and p.num_endoso=b.num_endoso"
        vlSql = vlSql & " left join pd_tmae_poltutor t on b.num_poliza =t.num_poliza and b.num_endoso=t.num_endoso and b.num_orden=t.num_orden"
        vlSql = vlSql & " left join ma_tpar_tabcod m on p.cod_banco=m.cod_elemento and m.cod_tabla='BCO'"
        vlSql = vlSql & " where fec_pago between '" & iFecDesde & "' and '" & iFecHasta & "' and p.cod_moneda ='" & vlMoneda & "'"
        vlSql = vlSql & " and p.num_endoso=(select max(num_endoso) from pd_tmae_poliza where num_poliza=l.num_poliza)"
        vlSql = vlSql & " GROUP BY l.num_poliza,p.cod_viapago,p.fec_vigencia,l.fec_pago,p.cod_moneda,l.cod_tippension,p.cod_tipoidenafi, p.num_idenafi,b.gls_nomben,b.gls_nomsegben,b.gls_patben,b.gls_matben,b.cod_tipoidenben,b.num_idenben,p.cod_afp,"
        vlSql = vlSql & " l.Gls_Direccion , l.Cod_InsSalud, p.Num_Endoso, b.Num_Orden, t.gls_nomtut, t.gls_nomsegtut, t.gls_pattut, t.gls_mattut, t.cod_tipoidentut, t.num_identut, p.Cod_Banco, p.Num_Cuenta, p.num_cuenta_cci, m.gls_elemento"
        vlSql = vlSql & " ORDER BY l.num_poliza"

        Set vgRs = vgConexionBD.Execute(vlSql)
        While Not (vgRs.EOF)
            vlMtoSal = 0
            
            vlNumPol = Trim(vgRs!Num_Poliza)
            vlNumOrd = CInt(vgRs!Num_Orden)
            vlFecPago = Trim(vgRs!Fec_Pago)
            'Obtiene la información del causante
            vgSql = ""
            vgSql = "SELECT cod_tipoidenben,num_idenben,gls_nomben,"
            vgSql = vgSql & "gls_nomsegben,gls_patben,gls_matben "
            vgSql = vgSql & "FROM pd_tmae_polben b "
            vgSql = vgSql & "WHERE cod_par='99' "
            vgSql = vgSql & "and num_poliza='" & Trim(vgRs!Num_Poliza) & "' "
            vgSql = vgSql & "and num_endoso=(select max(num_endoso) from pd_tmae_poliza where num_poliza=b.num_poliza) "
            Set vgRs2 = vgConexionBD.Execute(vgSql)
            If Not (vgRs2.EOF) Then
                If (Len(Trim(vgRs!Num_IdenBen)) <= 12) Then
                    vlRutContrat = Format(Trim(vgRs!Cod_TipoIdenBen), "00") & (Trim(vgRs!Num_IdenBen) & Space(12 - Len(Trim(vgRs!Num_IdenBen))))
                Else
                    vlRutContrat = Format(Trim(vgRs!Cod_TipoIdenBen), "00") & Mid(Trim(vgRs!Num_IdenBen), 1, 12)
                End If
                vlNomContrat = fgFormarNombreCompleto(IIf(IsNull(vgRs!Gls_NomBen), "", Trim(vgRs!Gls_NomBen)), IIf(IsNull(vgRs!Gls_NomSegBen), "", Trim(vgRs!Gls_NomSegBen)), IIf(IsNull(vgRs!Gls_PatBen), "", Trim(vgRs!Gls_PatBen)), IIf(IsNull(vgRs!Gls_MatBen), "", Trim(vgRs!Gls_MatBen)))
                If (Len(Trim(vlNomContrat)) <= 60) Then
                    vlNomContrat = vlNomContrat & Space(60 - Len(Trim(vlNomContrat)))
                Else
                    vlNomContrat = Mid(vlNomContrat, 1, 60)
                End If
            End If
            
            '****************** Movimiento de Pensión ****************************
            vlVar1 = Format(Trim(clTipRegPP), "00000")
            vlVar2 = Format(Trim(vgRs!Num_Poliza), "0000000000")
            vlVar3 = Format(Trim(vgRs!Cod_ViaPago), "00000")
            If Trim(vgRs!Fec_Pago) <> "" Then
                vlVar4 = DateSerial(Mid(vgRs!Fec_Pago, 1, 4), Mid(vgRs!Fec_Pago, 5, 2), Mid(vgRs!Fec_Pago, 7, 2))
                vlVar7 = Mid(vgRs!Fec_Pago, 1, 6)
            Else
                vlVar4 = Space(10)
                vlVar7 = Space(6)
            End If
            If (Len(Trim(vgRs!Cod_Moneda)) <= 5) Then
                vlVar8 = Trim(vgRs!Cod_Moneda) & Space(5 - Len(Trim(vgRs!Cod_Moneda)))
            Else
                vlVar8 = Mid(Trim(vgRs!Cod_Moneda), 1, 5)
            End If
            vlVar9 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlCodPen = Trim(vgRs!Cod_TipPension)
            
            Select Case vgRs!Cod_TipPension
             Case "04"
                vlVar10 = Format(Trim("76"), "00000")
             Case "05"
                vlVar10 = Format(Trim("76"), "00000")
             Case "06"
                vlVar10 = Format(Trim("94"), "00000")
             Case "07"
                vlVar10 = Format(Trim("94"), "00000")
             Case "08"
                vlVar10 = Format(Trim("95"), "00000")
            End Select
            
            'vlVar10 = Format(Trim(clRamoContPP), "00000")
            ''vlVar11 = Format(Trim(vgRs!cod_tipoidenafi), "00") & Format(Trim(vgRs!num_idenafi), "000000000000")
            ''vlVar12 = Trim(vgRs!Gls_NomReceptor) & IIf(IsNull(vgRs!Gls_NomSegReceptor), "", Trim(vgRs!Gls_NomSegReceptor)) & Trim(vgRs!Gls_PatReceptor) & Trim(vgRs!Gls_MatReceptor)
            ''vlVar12 = vlVar12 & Space(60 - Len(Trim(vlVar12)))
            vlVar11 = vlRutContrat
            vlVar12 = vlNomContrat
            vlVar13 = Format(Trim(clFrecPagPP), "00000")
            
            If vgRs!Cod_ViaPago = "02" Then
                If Len(vgRs!Num_Cuenta) <> 0 Then
                    vlVar17 = vgRs!Num_Cuenta
                    vlVar17 = vlVar17 & Space(60 - Len(Trim(vlVar17)))
                Else
                    vlVar17 = Space(60)
                End If
            Else
                If Len(vgRs!num_cuenta_cci) <> 0 Then
                    vlVar17 = vgRs!num_cuenta_cci
                    vlVar17 = vlVar17 & Space(60 - Len(Trim(vlVar17)))
                Else
                    vlVar17 = Space(60)
                End If
            End If
            
            
            vlVar19_Pen = Format(Trim(vgRs!Cod_AFP), "00000")
            vlVar20 = Trim(vgRs!gls_banco)
            vlVar20 = vlVar20 & Space(60 - Len(Trim(vlVar20)))
            If Trim(vgRs!Cod_Banco) = "00" Or Len(Trim(vgRs!Cod_Banco)) <> 0 Then
                vlVar22 = Format(Trim(vgRs!Cod_Banco), "00000")
            Else
                vlVar22 = String(5, "0")
            End If
            
            
            vlafp = Trim(vgRs!Cod_AFP)
            vlVar23 = Format(Trim(clReaPP), "0000000000")
            vlVar25 = Format(Trim(vgRs!Cod_TipPension), "00000")
            vlVar61 = Format(Trim(clTipMovSinLiqPP), "00000")
            vlVar62 = IIf(IsNull(vgRs!Mto_LiqPagar), 0, Format(vgRs!Mto_LiqPagar, "#0.00"))
            ''vlVar62 = String(18 - Len(Trim(vlVar62)), "0") & Trim(vlVar62)
            vlVar62 = flFormatNum18_2(vlVar62)
            vlVar67 = Trim(clTipPerNatPP) & Space(1 - Len(Trim(clTipPerNatPP)))
            
            vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                      (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                      (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                      (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19_Pen) & (vlVar20) & _
                      (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                      (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                      (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                      (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                      (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                      (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                      (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                      (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                      (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                      (vlVar66) & (vlVar67)
    
            vlLinea = Replace(vlLinea, ",", ".")
            Print #1, vlLinea
            
            'Guarda el detalle de Pensión de la póliza informada
            vlMtoPen = IIf(IsNull(vgRs!Mto_LiqPagar), 0, Format(vgRs!Mto_LiqPagar, "#0.00"))
            Call flGrabaDetPriPago(clTipMovSinLiqPP, vlVar2, vlNumOrd, vlFecPago, vlCodPen, vlafp, vlVar12, vlMtoPen)
            
            'Contador Primeros Pagos de Pensión
            vlNumCasosPPPension = vlNumCasosPPPension + 1
            vlMtoPPPension = vlMtoPPPension + vlMtoPen

    
            '****************** Movimiento de Salud ****************************
            vlVar19_Sal = Format(Trim(vgRs!Cod_InsSalud), "00000")
            vlVar61 = Format(Trim(clTipMovSinSalPP), "00000")
            vlVar62 = IIf(IsNull(vgRs!Mto_PlanSalud), 0, Format(vgRs!Mto_PlanSalud, "#0.00"))
            ''vlVar62 = String(18 - Len(Trim(vlVar62)), "0") & Trim(vlVar62)
            vlVar62 = flFormatNum18_2(vlVar62)
            vlVar67 = Trim(clTipPerJurPP) & Space(1 - Len(Trim(clTipPerJurPP)))
            If (vlVar62 <> 0) Then
            
                vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                          (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                          (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                          (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19_Sal) & (vlVar20) & _
                          (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                          (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                          (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                          (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                          (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                          (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                          (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                          (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                          (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                          (vlVar66) & (vlVar67)
    
                vlLinea = Replace(vlLinea, ",", ".")
                Print #1, vlLinea
                
                'Guarda el detalle de Salud de la póliza informada
                vlMtoSal = IIf(IsNull(vgRs!Mto_PlanSalud), 0, Format(vgRs!Mto_PlanSalud, "#0.00"))
                Call flGrabaDetPriPago(clTipMovSinSalPP, vlVar2, vlNumOrd, vlFecPago, vlCodPen, vlafp, vlVar12, vlMtoSal)

                'Contador Primeros Pagos de Salud
                vlNumCasosPPSalud = vlNumCasosPPSalud + 1
                vlMtoPPSalud = vlMtoPPSalud + vlMtoSal
                
            End If
            
                        
            '***************** RESUMEN DE MOVIMIENTO (por persona) ********************
            vlVar19_Pen = Format(Trim(vgRs!Cod_AFP), "00000") 'cod_afp
            vlVar61 = Format(Trim(clTipMovSinResPP), "00000")
            vlVar62 = IIf(IsNull(vgRs!Mto_PensionTot), 0, Format(vgRs!Mto_PensionTot, "#0.00"))
            vlMtoPenTot = vlVar62
            vlVar62 = flFormatNum18_2(vlVar62)
            
            vlVar63 = IIf(IsNull(vgRs!Mto_LiqPagar), 0, Format(vgRs!Mto_LiqPagar, "#0.00"))
            vlVar63 = flFormatNum18_2(vlVar63)
            
            vlVar66 = IIf(IsNull(vgRs!Mto_PlanSalud), 0, Format(vgRs!Mto_PlanSalud, "#0.00"))
            vlVar66 = flFormatNum18_2(vlVar66)
            
            vlVar67 = Trim(clTipPerNatPP) & Space(1 - Len(Trim(clTipPerNatPP)))
        
            vlLinea = (vlVar1) & (vlVar2) & (vlVar3) & (vlVar4) & (vlVar5) & _
                      (vlVar6) & (vlVar7) & (vlVar8) & (vlVar9) & (vlVar10) & _
                      (vlVar11) & (vlVar12) & (vlVar13) & (vlVar14) & (vlVar15) & _
                      (vlVar16) & (vlVar17) & (vlVar18) & (vlVar19_Pen) & (vlVar20) & _
                      (vlVar21) & (vlVar22) & (vlVar23) & (vlVar24) & (vlVar25) & _
                      (vlVar26) & (vlVar27) & (vlVar28) & (vlVar29) & (vlVar30) & _
                      (vlVar31) & (vlVar32) & (vlVar33) & (vlVar34) & (vlVar35) & _
                      (vlVar36) & (vlVar37) & (vlVar38) & (vlVar39) & (vlVar40) & _
                      (vlVar41) & (vlVar42) & (vlVar43) & (vlVar44) & (vlVar45) & _
                      (vlVar47) & (vlVar48) & (vlVar49) & (vlVar50) & _
                      (vlVar51) & (vlVar52) & (vlVar53) & (vlVar54) & (vlVar55) & _
                      (vlVar56) & (vlVar57) & (vlVar58) & (vlVar59) & (vlVar60) & _
                      (vlVar61) & (vlVar62) & (vlVar63) & (vlVar64) & (vlVar65) & _
                      (vlVar66) & (vlVar67)

            vlLinea = Replace(vlLinea, ",", ".")
            Print #1, vlLinea
            
            'update al registro de Mov de Pensión para Crear el registro 1 (detalle montos)
            Call flActDetallePriPagos(clTipRegPP, clTipMovSinLiqPP, vlNumPol, vlNumOrd, vlMtoPenTot, vlMtoSal)
                       
            'Limpia las variables de resumen ya que no se informan en pension ni salud
            vlVar63 = String(18, "0")
            vlVar64 = String(5, "0")
            vlVar66 = String(18, "0")
            vgRs.MoveNext
        Wend
        vgRs.Close
        
        'Actualiza la cantidad de casos y mtos informados Pension
        Call flActEstadisticaPriPago(clTipRegPP, clTipMovSinLiqPP, vlNumCasosPPPension, vlMtoPPPension)
        'Actualiza la cantidad de casos y mtos informados Salud
        Call flActEstadisticaPriPago(clTipRegPP, clTipMovSinSalPP, vlNumCasosPPSalud, vlMtoPPSalud)
        
        vgRs1.MoveNext
    Wend
    vgRs1.Close
    
Close #1
vlOpen = False

flExportarPriPagos = True

Exit Function
Errores:
Screen.MousePointer = vbDefault
If Err.Number <> 0 Then
    If vlOpen Then
        Close #1
''        Unload Frm_BarraProg
    End If
    MsgBox "Se ha producido el siguiente error : " & Err.Description, vbCritical, "Error"
End If
End Function

Private Sub Txt_Desde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_Hasta.SetFocus
End If
End Sub

Private Sub Txt_Desde_LostFocus()
    If Txt_Desde <> "" Then
        If (flValidaFecha(Txt_Desde) = False) Then
            Txt_Desde = ""
            Exit Sub
        End If
        Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
        Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
    End If
End Sub

Private Sub Txt_Hasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdContable.SetFocus
End If
End Sub

Private Sub Txt_Hasta_LostFocus()
    If Txt_Hasta <> "" Then
        If (flValidaFecha(Txt_Hasta) = False) Then
            Txt_Hasta = ""
            Exit Sub
        End If
        Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
        Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
    End If
End Sub

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

Private Function flEstadisticaPriUnica()

    vgSql = "INSERT INTO PD_TMAE_CONTABLEPRIUNI ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "cod_moneda,fec_desde,fec_hasta,"
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(clTipRegPU) & "',"
    vgSql = vgSql & "'" & Trim(clTipMovPU) & "',"
    vgSql = vgSql & "'" & Trim(vgMonedaCodOfi) & "',"
    vgSql = vgSql & "'" & vlFecDesde & "',"
    vgSql = vgSql & "'" & vlFecHasta & "',"
    vgSql = vgSql & "'" & (vgUsuario) & "',"
    vgSql = vgSql & "'" & vlFecCrea & "',"
    vgSql = vgSql & "'" & vlHorCrea & "')"
    vgConexionBD.Execute (vgSql)
       
    Lbl_NumArchivo.Caption = vlNumArchivo
    
End Function

Private Function flEstadisticaPriPago(iTipMov As String)

    vgSql = "INSERT INTO PD_TMAE_CONTABLEPRIPAGO ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "cod_moneda,fec_desde,fec_hasta,"
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(clTipRegPP) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & Trim(vlMoneda) & "',"
    vgSql = vgSql & "'" & vlFecDesde & "',"
    vgSql = vgSql & "'" & vlFecHasta & "',"
    vgSql = vgSql & "'" & (vgUsuario) & "',"
    vgSql = vgSql & "'" & vlFecCrea & "',"
    vgSql = vgSql & "'" & vlHorCrea & "')"
    vgConexionBD.Execute (vgSql)
    
    Lbl_NumArchivo.Caption = vlNumArchivo
   
End Function

Private Function flActEstadisticaPriUnica()

    vgSql = "UPDATE PD_TMAE_CONTABLEPRIUNI set "
    vgSql = vgSql & "num_casos = " & vlNumCasosPriUni & ","
    vgSql = vgSql & "mto_primas = " & str(vlMtoPrimas) & " "
    vgSql = vgSql & "WHERE num_archivo = " & vlNumArchivo & " "
    vgSql = vgSql & "and cod_moneda='" & vgMonedaCodOfi & "' "
    vgConexionBD.Execute (vgSql)
          
End Function

Private Function flActEstadisticaPriPago(iTipReg As String, iTipMov As String, iNumCasos As Long, iMtoPago As Double)

    vgSql = "UPDATE PD_TMAE_CONTABLEPRIPAGO set "
    vgSql = vgSql & "num_casos = " & iNumCasos & ","
    vgSql = vgSql & "mto_pago = " & str(iMtoPago) & " "
    vgSql = vgSql & "WHERE num_archivo = " & vlNumArchivo & " "
    vgSql = vgSql & "and cod_tipreg ='" & Trim(iTipReg) & "' "
    vgSql = vgSql & "and cod_tipmov ='" & Trim(iTipMov) & "' "
    vgSql = vgSql & "and cod_moneda ='" & vlMoneda & "' "
    vgConexionBD.Execute (vgSql)
          
End Function

'FUNCION QUE GUARDA EL NOMBRE DEL FORMULARIO DEL QUE SE LLAMO A LA FUNCION
Function flInicio(iNomForm)
    vgNomForm = iNomForm
    Call Form_Load
End Function

Private Function flInicializaVar()

    vlVar1 = String(5, "0")     'Tipo de Registro
    vlVar2 = String(10, "0")    'POLIZA
    vlVar3 = String(5, "0")     'SUCURSAL
    vlVar4 = Space(10)          'VIGENCIA "DESDE" DE LA POLIZA
    vlVar5 = Space(10)          '--VIGENCIA "HASTA" DE LA POLIZA
    vlVar6 = Space(10)          '--VIGENCIA 'DESDE' ORIGINAL
    vlVar7 = String(6, "0")     'FECHA CONTABLE(MES / ANO)
    vlVar8 = Space(5)           'MONEDA DEL MOVIMIENTO
    vlVar9 = String(5, "0")     'COBERTURA
    vlVar10 = String(5, "0")    'RAMO CONTABLE
    vlVar11 = Space(14)         'CONTRATANTE RUT
    vlVar12 = Space(60)         'CONTRATANTE NOMBRE
    vlVar13 = String(5, "0")    'FRECUENCIA DE PAGO (INMEDIATA)
    vlVar14 = Space(14)         '--INTERMERDIARIO GTO DE COBRANZA
    vlVar15 = String(5, "0")    '--REGISTRO EN TRATAMIENTO
    vlVar16 = Space(10)         '--FECHA EFECTO REGISTRO (MOV)
    vlVar17 = Space(60)         'NOMBRE
    vlVar18 = Space(14)         'RUT
    vlVar19_Pen = String(5, "0")   'SUCURSAL
    vlVar19_Sal = String(5, "0")   'SUCURSAL
    vlVar20 = Space(60)         '--NOMBRE INTERMEDIARIO
    vlVar21 = Space(14)         '--RUT INNTERMEDIARIO
    vlVar22 = String(5, "0")    'TIPO DE INTERMEDIARIO
    vlVar23 = String(10, "0")   '--REASEGURADOR
    vlVar24 = Space(40)         'NACIONALIDAD REASEGURADOR
    vlVar25 = String(5, "0")    'CONTRATO DE REASEGURO
    vlVar26 = Space(6)          'TIPO DE REASEGURO
    vlVar27 = String(10, "0")   'NUMERO DE SINIESTRO
    vlVar28 = Space(6)          'ESTADO DEL SINIESTRO
    vlVar29 = String(5, "0")    'TIPO MOVIMIENTO
    vlVar30 = String(10, "0")   'NUMERO DE MOVIMIENTO
    vlVar31 = String(18, "0")   'MONTO EXENTO PRIMA
    vlVar32 = String(18, "0")   '--MONTO AFECTO PRIMA
    vlVar33 = String(18, "0")   '--MONTO IGV PRIMA
    vlVar34 = String(18, "0")   '--MONTO BRUTO PRIMA
    vlVar35 = String(18, "0")   '--MONTO NETO PRIMA DEVENGADA
    vlVar36 = String(18, "0")   '--CAPITALES ASEGURADOS
    vlVar37 = String(5, "0")    '--ORIGEN DEL RECIBO
    vlVar38 = String(5, "0")    '--TIPO DE MOVIMIENTO
    vlVar39 = String(18, "0")   '--MONTO PRIMA CEDIDA ANTES DSCTO
    vlVar40 = String(18, "0")   '--MONTO DESC. POR PRIMA CEDIDA
    vlVar41 = String(4, "0")    '--MONTO IMPUESTO 2%
    vlVar42 = String(18, "0")   '--MONTO EXCESO DE PERDIDA
    vlVar43 = String(18, "0")   '--CAPITALES CEDIDOS
    vlVar44 = Space(6)          '--TIPO DE RESERVA
    vlVar45 = String(18, "0")   '--MONTO RESERVA MATEMATICA
    vlVar47 = String(5, "0")    '-- % DE COMSION SOBRE LA PRIMA
    vlVar48 = Space(6)          '--TIPO DE COMISION
    vlVar49 = String(18, "0")   '--MONTO COMISION NETA
    vlVar50 = String(18, "0")   '--MONTO IGV COMISION
    vlVar51 = String(18, "0")   '--MONTO BRUTO COMISION
    vlVar52 = String(5, "0")    '--PERIODO DE GRACIA
    vlVar53 = String(18, "0")   '--MONTO NETO COMISION
    vlVar54 = Space(6)          '--ESQUEMA DE PAGO
    vlVar55 = String(10, "0")   '--fecha desde
    vlVar56 = String(10, "0")   '--fecha Hasta
    vlVar57 = String(5, "0")    '--RAMO
    vlVar58 = String(5, "0")    '--PRODUCTO
    vlVar59 = String(10, "0")   '--POLIZA
    vlVar60 = Space(14)         '--RUT DEL CLIENTE
    vlVar61 = String(5, "0")    '--TIPO DE MOVIMIENTO SINIESTRO
    vlVar62 = String(18, "0")   '--MONTO
    vlVar63 = String(18, "0")   '--MONTO CEDIDO EN EL MES
    vlVar64 = String(5, "0")    '-- % DE COMISION DE GASTOS DE COB
    vlVar65 = String(18, "0")   '--MTO. GASTOS DE COB. PRIMA REC.
    vlVar66 = String(18, "0")   '--MTO. GASTOS DE COB. PRIMA DEV.
    vlVar67 = Space(1)          '--Tipo de persona (jurídico / natural)

End Function

Private Function flFormatNum18_2(iNumero As String) As String
Dim vlNum As String
    
    vlNum = Format(iNumero, "#00000000000000.00")

    If (CDbl(vlNum) < 0) Then
        vlNum = Mid(vlNum, 1, 1) & "0" & Mid(vlNum, 2, 14) & Mid(vlNum, 17, 2)
    Else
        vlNum = "00" & Mid(vlNum, 1, 14) & Mid(vlNum, 16, 2)
    End If
    
    flFormatNum18_2 = vlNum

End Function
Private Function flFormatNum6_2(iNumero As String) As String
Dim vlNum As String
Dim x As Integer
    
    vlNum = Format(iNumero, "#00.00")

'    If (CDbl(vlNum) < 0) Then
'        vlNum = Mid(vlNum, 1, 1) & "0" & Mid(vlNum, 2, 2) & Mid(vlNum, 5, 2)
'    Else
'        vlNum = Mid(vlNum, 1, 2) & Mid(vlNum, 4, 2)
'
'    End If
    
    x = InStr(1, iNumero, ".")
    vlNum = Mid(iNumero, 1, x - 1) & Mid(iNumero, x + 1, 2)
    
    flFormatNum6_2 = Format(CDbl(vlNum), "000000")

End Function
Private Function flFormatNum5_2(iNumero As String) As String
Dim vlNum As String
Dim varPunto As Integer
'    vlNum = Format(iNumero, "#000.00")

'    If (CDbl(vlNum) < 0) Then
'        ''vlNum = Mid(vlNum, 1, 1) & Mid(vlNum, 2, 3) & Mid(vlNum, 6, 2)
'        vlNum = Mid(vlNum, 2, 3) & Mid(vlNum, 6, 2)
'    Else
'        ''vlNum = "00" & Mid(vlNum, 1, 3) & Mid(vlNum, 5, 2)
'        vlNum = Mid(vlNum, 1, 3) & Mid(vlNum, 5, 2)
'    End If
    iNumero = Replace(iNumero, ".", "")
'    vlNum = Format(iNumero, "#0000.00")
'    If (CDbl(vlNum) < 0) Then
'        ''vlNum = Mid(vlNum, 1, 1) & Mid(vlNum, 2, 3) & Mid(vlNum, 6, 2)
'        vlNum = Mid(vlNum, 2, 3) & Mid(vlNum, 6, 2)
'    Else
'        ''vlNum = "00" & Mid(vlNum, 1, 3) & Mid(vlNum, 5, 2)
'        vlNum = Mid(vlNum, 1, 5) & Mid(vlNum, 6, 2)
'    End If

'    varPunto = InStr(1, vlNum, ".")
'    varPunto = Replace(vlNum, ".", "")
    
'    flFormatNum5_2 = vlNum
flFormatNum5_2 = CDbl(iNumero)
End Function


Private Function flGrabaDetPriUnica(iNumPol As String, iNumOrd As Integer, iFecTrasp As String, iPension As String, iAfp As String, iLiq As String, iNombre As String, iMtoPrima As Double)

    vgSql = "INSERT INTO PD_TMAE_CONTABLEDETPRIUNI ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "num_poliza,num_orden,fec_traspaso,cod_moneda,"
    vgSql = vgSql & "cod_tippension,cod_afp,cod_liquidacion,"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "gls_nombre,"
    vgSql = vgSql & "mto_prima "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(clTipRegPU) & "',"
    vgSql = vgSql & "'" & Trim(clTipMovPU) & "',"
    vgSql = vgSql & "'" & iNumPol & "',"
    vgSql = vgSql & " " & iNumOrd & ","
    vgSql = vgSql & "'" & iFecTrasp & "',"
    vgSql = vgSql & "'" & Trim(vlMoneda) & "',"
    vgSql = vgSql & "'" & Trim(iPension) & "',"
    vgSql = vgSql & "'" & Trim(iAfp) & "',"
    vgSql = vgSql & "'" & Trim(iLiq) & "',"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "'" & Trim(iNombre) & "',"
    vgSql = vgSql & " " & str(iMtoPrima) & ")"
    vgConexionBD.Execute (vgSql)

End Function

Private Function flGrabaDetPriPago(iTipMov As String, iNumPol As String, iNumOrd As Integer, iFecPago As String, iPension As String, iAfp As String, iNombre As String, iMtoPago As Double)

    vgSql = "INSERT INTO PD_TMAE_CONTABLEDETPRIPAGO ("
    vgSql = vgSql & "num_archivo,cod_tipreg,cod_tipmov,"
    vgSql = vgSql & "num_poliza,num_orden,fec_pago,cod_moneda,"
    vgSql = vgSql & "cod_tippension,cod_afp,"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "gls_nombre,"
    vgSql = vgSql & "mto_pago "
    vgSql = vgSql & ") VALUES ("
    vgSql = vgSql & " " & vlNumArchivo & ","
    vgSql = vgSql & "'" & Trim(clTipRegPP) & "',"
    vgSql = vgSql & "'" & Trim(iTipMov) & "',"
    vgSql = vgSql & "'" & iNumPol & "',"
    vgSql = vgSql & " " & iNumOrd & ","
    vgSql = vgSql & "'" & iFecPago & "',"
    vgSql = vgSql & "'" & Trim(vlMoneda) & "',"
    vgSql = vgSql & "'" & Trim(iPension) & "',"
    vgSql = vgSql & "'" & Trim(iAfp) & "',"
    If (Trim(iNombre) <> "") Then vgSql = vgSql & "'" & Trim(iNombre) & "',"
    vgSql = vgSql & " " & str(iMtoPago) & ")"
    vgConexionBD.Execute (vgSql)

End Function

Private Function flActDetallePriPagos(iTipReg As String, iTipMov As String, iNumPol As String, iNumOrd As Integer, iMtoPenTotal As Double, iMtoSalud As Double)

    vgSql = "UPDATE PD_TMAE_CONTABLEDETPRIPAGO set "
    vgSql = vgSql & "mto_pentot = " & str(iMtoPenTotal) & ","
    vgSql = vgSql & "mto_salud = " & str(iMtoSalud) & " "
    vgSql = vgSql & "WHERE num_archivo = " & vlNumArchivo & " "
    vgSql = vgSql & "and cod_tipreg ='" & Trim(iTipReg) & "' "
    vgSql = vgSql & "and cod_tipmov ='" & Trim(iTipMov) & "' "
    vgSql = vgSql & "and num_poliza ='" & iNumPol & "' "
    vgSql = vgSql & "and num_orden = " & iNumOrd & " "
    vgConexionBD.Execute (vgSql)
          
End Function

Private Sub flImpDetallePen()
Dim vlArchivo As String
Dim objRep As New ClsReporte
Err.Clear
On Error GoTo Errores1
   
    Screen.MousePointer = 11
    
    If (Trim(Lbl_NumArchivo) = "") Then
        MsgBox "Debe seleccionar un Periodo a Imprimir.", vbInformation, "Falta Información"
        Screen.MousePointer = 0
        Exit Sub
    End If
        
    If Not fgExiste(vlArchivo) Then
        MsgBox "Archivo de Reporte de Detalle de Archivo Contable no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
        Screen.MousePointer = 0
        Exit Sub
    End If
'    If (vgNomForm = "PriPag") Then
'        vlArchivo = strRpt & "PD_Rpt_ContableDetPriPagosPen.rpt"   '\Reportes
'        vgQuery = "{PD_TMAE_CONTABLEDETPRIPAGO.NUM_ARCHIVO} = " & Trim(Lbl_NumArchivo) & " AND "
'        vgQuery = vgQuery & "{PD_TMAE_CONTABLEDETPRIPAGO.COD_TIPMOV} = '" & Trim(clTipMovSinLiqPP) & "' "
'    Else
'        Exit Sub
'    End If
'
'    If Not fgExiste(vlArchivo) Then     ', vbNormal
'        MsgBox "Archivo de Reporte de Detalle de Archivo Contable no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'
'    Rpt_General.Reset
'    Rpt_General.WindowState = crptMaximized
'    Rpt_General.ReportFileName = vlArchivo
'    Rpt_General.Connect = vgRutaDataBase
'    Rpt_General.Destination = crptToWindow
'    Rpt_General.SelectionFormula = ""
'    Rpt_General.SelectionFormula = vgQuery
'
'    Rpt_General.Formulas(0) = ""
'    Rpt_General.Formulas(1) = ""
'    Rpt_General.Formulas(2) = ""
'    Rpt_General.Formulas(3) = ""
'
'    Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
'    Rpt_General.Formulas(1) = "NombreSistema = '" & vgNombreSistema & "'"
'    Rpt_General.Formulas(2) = "NombreSubSistema = '" & vgNombreSubSistema & "'"
'
'    Rpt_General.WindowTitle = "Informe Detalle Archivo Contable"
'    Rpt_General.Action = 1
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    Dim LNGa As Long
    
    rs.Open "PD_LISTA_ARCH_CONT_PAGOS_PEN.LISTAR('" & Lbl_NumArchivo.Caption & "','" & Trim(clTipMovSinLiqPP) & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly

    LNGa = CreateFieldDefFile(rs, Replace(UCase(strRpt & "Estructura\PD_Rpt_ContableDetPriPagosPen.rpt"), ".RPT", ".TTX"), 1)
    If objRep.CargaReporte(strRpt & "", "PD_Rpt_ContableDetPriPagosPen.rpt", "Informe Detalle Archivo Contable", rs, True, _
                        ArrFormulas("NombreCompania", vgNombreCompania), _
                        ArrFormulas("NombreSistema", vgNombreSistema), _
                        ArrFormulas("NombreSubSistema", vgNombreSubSistema)) = False Then
        
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
        
        
    Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Private Function flMostrarOption()
    Opt_Resumen.Left = 2160
    Opt_DetMov.Left = 3840
    Opt_Resumen.Visible = True
    Opt_DetMov.Visible = True
    Opt_DetPen.Visible = True
End Function

Private Function flOcultarOption()
    Opt_Resumen.Left = 2760
    Opt_DetMov.Left = 5040
    Opt_Resumen.Visible = True
    Opt_DetMov.Visible = True
    Opt_DetPen.Visible = False
End Function

