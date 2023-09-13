VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_CalInfSBS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes para la SBS."
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8070
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   7815
      Begin VB.CommandButton Cmd_Imprimir_reporte 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   1680
         Picture         =   "Frm_CalInfSBS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Refrescar 
         Caption         =   "&Refrescar"
         Height          =   675
         Left            =   3600
         Picture         =   "Frm_CalInfSBS.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Restaurar o refrescar los datos desde la BD"
         Top             =   240
         Width           =   790
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   240
         Picture         =   "Frm_CalInfSBS.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4680
         Picture         =   "Frm_CalInfSBS.frx":12E6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_CalInfSBS.frx":19A0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Archivo"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_CalInfSBS.frx":1A9A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exportar a Archivo"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Fra_Resumen 
      Caption         =   "  Resumen de Selección  "
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
      TabIndex        =   15
      Top             =   1200
      Width           =   7815
      Begin VB.Label Lbl_MtoTotPri 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Lbl_PolSinRecPri 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Lbl_PolRecPri 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Lbl_TotPol 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Número de Pólizas Sin Recepción de Primas"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   19
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Número de Pólizas Con Recepción de Primas"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Monto Total Primas"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   17
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Número Total de Pólizas"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   3375
      End
   End
   Begin Crystal.CrystalReport Rpt_SBS 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   6000
         Picture         =   "Frm_CalInfSBS.frx":22BC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Efectuar Consulta"
         Top             =   360
         Width           =   855
      End
      Begin MSComDlg.CommonDialog ComDialogo 
         Left            =   0
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rango de Fechas   :"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   360
         Width           =   1695
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
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Incorporación de Pólizas"
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
         Left            =   2280
         TabIndex        =   12
         Top             =   0
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Frm_CalInfSBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const clEstadoSinTraspasoPrima As String * 3 = "PNT"
Const clEstadoConTraspasoPrima As String * 2 = "PT"
Const clCoberturaSin As String * 2 = "SC"
Const clCoberturaCon As String * 2 = "CC"
Const clCodTipRecR As String * 1 = "FOR0067_"

Dim vlFechaDesde As String, vlFechaHasta As String
Dim vlDesde      As String, vlHasta      As String
Dim vlNumeroPoliza As String
Dim vlArchivo As String
Dim vlSwExiste As Boolean
Dim vlValor As Double

Const clCodFormato As String * 2 = "67"
Const clAnexo As String * 1 = "1"
Const clCodExpresionMtos As String * 2 = "12"
Const clDatosControl As String * 1 = "0"

Const clCodArchSucave_01 As String = "01" '17/06/2008 Código del archivo sucave segun sbs

Function flLimpiar()
On Error GoTo Err_Limpiar

    Txt_Desde.Text = ""
    Txt_Hasta.Text = ""
    Lbl_TotPol.Caption = ""
    Lbl_PolRecPri.Caption = ""
    Lbl_PolSinRecPri.Caption = ""
    Lbl_MtoTotPri.Caption = ""
    Txt_Desde.SetFocus
   
Exit Function
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Sub plCalcularResumen()

    vlValor = 0
    'Obtiene la cantidad de polizas, a partir de un rango de fechas
    Sql = "SELECT count(num_poliza) as NumeroPoliza FROM pd_ttmp_sbstasacot "
    Sql = Sql & " where cod_usuario = '" & vgUsuario & "'"
    Set vgRs = vgConexionBD.Execute(Sql)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!numeropoliza) Then vlValor = vgRs!numeropoliza
    End If
    vgRs.Close
    Lbl_TotPol = Format(vlValor, "#,#0")
      
    vlValor = 0
    'Obtiene la cantidad de polizas CON RECEPCION PRIMAS, a partir de un rango de fechas
    Sql = "SELECT count(num_poliza) as NumeroPoliza FROM pd_ttmp_sbstasacot "
    Sql = Sql & " where cod_usuario = '" & vgUsuario & "'"
    Sql = Sql & " AND gls_estado = '" & clEstadoConTraspasoPrima & "'"
    Set vgRs = vgConexionBD.Execute(Sql)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!numeropoliza) Then vlValor = vgRs!numeropoliza
    End If
    vgRs.Close
    Lbl_PolRecPri = Format(vlValor, "#,#0")
      
    vlValor = 0
    'Obtiene la cantidad de polizas SIN RECEPCION PRIMAS, a partir de un rango de fechas
    Sql = "SELECT count(num_poliza) as NumeroPoliza FROM pd_ttmp_sbstasacot "
    Sql = Sql & " where cod_usuario = '" & vgUsuario & "'"
    Sql = Sql & " AND gls_estado = '" & clEstadoSinTraspasoPrima & "'"
    Set vgRs = vgConexionBD.Execute(Sql)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!numeropoliza) Then vlValor = vgRs!numeropoliza
    End If
    vgRs.Close
    Lbl_PolSinRecPri = Format(vlValor, "#,#0")
    
    vlValor = 0
    'Obtiene la suma de las primas, a partir de un rango de fechas
    Sql = "select sum(mto_priuni) as SumaPrima from pd_ttmp_sbstasacot "
    Sql = Sql & " where cod_usuario = '" & vgUsuario & "'"
    Set vgRs = vgConexionBD.Execute(Sql)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!sumaprima) Then vlValor = vgRs!sumaprima
    End If
    vgRs.Close
    Lbl_MtoTotPri = Format(vlValor, "#,#0.00")
    
End Sub

Function flValidarExistenciaPeriodo(oFechaDesde As String, oFechaHasta As String) As Boolean
'Función: Permite validar la Existencia de un Periodo para el Usuario Conectado
'Parámetros de Entrada:
'Parámetros de Salida:
'- Retorna un verdadero si Existe Periodo, y un Falso en caso contrario
'- oFechaDesde  => Fecha de Inicio del periodo generado
'- oFechaHasta  => Fecha de Fin del periodo generado
'------------------------------------------------------------
'Fecha de Creación     : 17/08/2007
'Fecha de Modificación :
'------------------------------------------------------------

    flValidarExistenciaPeriodo = False
    oFechaDesde = ""
    oFechaHasta = ""
    
    'Determinar la Existencia de Información
    vgQuery = " SELECT DISTINCT fec_ini, fec_fin "
    vgQuery = vgQuery & "FROM pd_ttmp_sbstasacot "
    vgQuery = vgQuery & "WHERE cod_usuario = '" & vgUsuario & "'"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not vgRs.EOF Then
        oFechaDesde = vgRs!fec_ini
        oFechaHasta = vgRs!fec_fin
        flValidarExistenciaPeriodo = True
    End If
    vgRs.Close

End Function

Function flRefrescarPeriodo() As Boolean
'Función: Permite obtener los datos registrados en la BD
'Parámetros de Entrada:
'Parámetros de Salida:
'- Retorna un verdadero cuando existe información, y un falso en caso contrario
'--------------------------------------------------------------
'- Fecha de Creación     : 17/08/2007
'- Fecha de Modificación :
'--------------------------------------------------------------

    flRefrescarPeriodo = False
    vlFechaDesde = ""
    vlFechaHasta = ""
    
    If (flValidarExistenciaPeriodo(vlFechaDesde, vlFechaHasta) = True) Then
        Txt_Desde = DateSerial(Mid(vlFechaDesde, 1, 4), Mid(vlFechaDesde, 5, 2), Mid(vlFechaDesde, 7, 2))
        Txt_Hasta = DateSerial(Mid(vlFechaHasta, 1, 4), Mid(vlFechaHasta, 5, 2), Mid(vlFechaHasta, 7, 2))
        
        Call plCalcularResumen
    End If
    
    flRefrescarPeriodo = True
End Function

Function flArchivoAnexoSVS()
Dim vlNumRegistros As Integer
Dim vlAumento As Double
Dim vlFiller1 As String, vlFiller2 As String
Dim vlOpen As Boolean
'Variables del Encabezado
Dim vlCodformCons As String, vlAnexoCons As String
Dim vlEspCodIdenForm As String, vlCodSbs As String
Dim vlFechaRep As String, vlExpMontos As String
Dim vlDatCon As String
'Variables del Detalle
Dim vlCorrelativo As String
Dim vlNumPoliza As String
Dim vlCodCuspp As String, vlEspCuspp As String
Dim vlafp As String, vlFecDev As String
Dim vlGlsTipPen As String, vlEspTipPen  As String
Dim vlGlsTipRen As String, vlEspModPen  As String
Dim vlMtoPension As String
Dim vlNumMesDif As String, vlNumMesGar As String
Dim vlFecAcepta As String, vlFecPriPago As String
Dim vlMtoPriUni As String
Dim vlGlsMoneda As String
Dim vlPrcTasaVta As String
Dim vlA  As String, vlB  As String
Dim vlEspacios As String
Dim vlLinea As String
Dim vlTipoCambio As String
Dim vlEstado As String

On Error GoTo Err_flInformeAnexoSVS

    vlFiller1 = ""
    
    vlNumRegistros = 0

    Open vlArchivo For Output As #1
    vlOpen = True

'1. ENCABEZADO
'--------------
    'vlCodSbs = vgCodigoSBSCompania
    vlCodSbs = vgCodigoSucaveCompañia

    'Código de Formato
    vlCodformCons = clCodFormato
    vlCodformCons = Format(vlCodformCons, "0000")

    'Anexo
    vlAnexoCons = clAnexo
    vlAnexoCons = Format(vlAnexoCons, "00")

    'Entidad
'I--- ABV 23/11/2007 ---
'    vlEspCodIdenForm = Space(5 - Len(vlCodSbs))
'    vlCodSbs = (vlCodSbs & vlEspCodIdenForm)
    vlCodSbs = String(5 - Len(vlCodSbs), "0") & vlCodSbs
'F--- ABV 23/11/2007 ---
    
    'Fecha
'    vlFechaRep = vlFechaDesde
    vlFechaRep = vlFechaHasta

    'Código de Expresión de Montos
    vlExpMontos = clCodExpresionMtos
    vlExpMontos = Format(vlExpMontos, "000")
    
    'Datos de Control
    vlDatCon = clDatosControl
    vlDatCon = Format(vlDatCon, "000000000000000")

    'Registro Tipo 1
    vlLinea = vlCodformCons & vlAnexoCons & vlCodSbs & _
    vlFechaRep & vlExpMontos & vlDatCon
    Print #1, vlLinea

'2. DETALLE
'--------------
    vgI = 1
    
    'Registro Tipo 2
    'Números de Pólizas que se vendieron en el Mes anterior al Mes indicado
    vgSql = "SELECT "
    'I - MC 18/06/2008 Se informa el código de la AFP no el nombre
    'vgSql = vgSql & "num_poliza,cod_cuspp,gls_afp as cod_afp,fec_dev,gls_tippension, "
    vgSql = vgSql & "num_poliza,cod_cuspp,cod_afp as cod_afp,fec_dev,gls_tippension, "
    'F - 18/06/2008
    vgSql = vgSql & "gls_tipren,mto_pension,num_mesdif,num_mesgar, "
    vgSql = vgSql & "fec_acepta,fec_pripago,mto_priuni,gls_moneda,prc_tasavta "
    vgSql = vgSql & ",mto_valmoneda,cod_tippension,gls_estado "
    vgSql = vgSql & "FROM pd_ttmp_sbstasacot "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "cod_usuario = '" & vgUsuario & "'"
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    While Not vlRegistro.EOF
            
        'Código de Fila
        vlCorrelativo = Format(vgI, "0000")
            
        'Número de Póliza
        vlNumPoliza = vlRegistro!Num_Poliza
        'vlNumPoliza = Format(vlRegistro!Num_Poliza, "000000000000000")
        vlNumPoliza = String(15 - Len(vlRegistro!Num_Poliza), "0") & vlRegistro!Num_Poliza
        
        'CUSPP del afiliado
        vlCodCuspp = vlRegistro!Cod_Cuspp
        vlEspCuspp = Space(12 - Len(vlCodCuspp))
        vlCodCuspp = (vlCodCuspp & vlEspCuspp)
        
        'AFP
        'DC 17/06/2009
        vlafp = Format(Trim(vlRegistro!Cod_AFP), "00000")
'I--- 23/11/2007 ---
        'vlafp = Trim(vlRegistro!Cod_AFP)
        'vlEspacios = Space(5 - Len(vlafp))
        'vlafp = vlafp & vlEspacios
'        vlAfp = Trim(vlRegistro!Cod_AFP)
'        vlAfp = String(5 - Len(vlAfp), "0") & vlAfp
'F--- 23/11/2007 ---
        
        'Fecha de Devengue
        vlFecDev = vlRegistro!Fec_Dev
        
        '----------------------------------------------------------
        'TIPOS DE PENSIÓN
        '----------------------------------------------------------
        'Tipo de Pensión Sin Cobertura
        If (vlRegistro!Cod_TipPension = "04") Or (vlRegistro!Cod_TipPension = "05") Then
            vlGlsTipPen = Trim(vlRegistro!GLS_TIPPENSION)
        Else
            'I - MC 18/06/2008 'Agrega la cobertura al tipo de pensión
            'vlGlsTipPen = Mid(Trim(vlRegistro!GLS_TIPPENSION), 1, 1)
            vlGlsTipPen = Mid(Trim(vlRegistro!GLS_TIPPENSION), 1, 4)
            'F - MC 18/06/2008
        End If
        vlEspTipPen = Space(4 - Len(vlGlsTipPen))
        vlGlsTipPen = (vlGlsTipPen & vlEspTipPen)
    
        'Modalidad de Pensión
        vlGlsTipRen = (vlRegistro!GLS_TIPREN)
        vlEspModPen = Space(4 - Len(vlGlsTipRen))
        vlGlsTipRen = (vlGlsTipRen & vlEspModPen)
        
        'Monto de Pensión
        vlMtoPension = Format(vlRegistro!Mto_Pension, "0000000000000.00")
        vlA = Mid(vlMtoPension, 1, 13)
        vlB = Mid(vlMtoPension, 15, 2)
        vlMtoPension = vlA & vlB

        'Periodo Diferido
        If (vlRegistro!Num_MesDif = 0) Then
            vlNumMesDif = vlRegistro!Num_MesDif
        Else
            vlNumMesDif = vlRegistro!Num_MesDif / 12
        End If
        vlNumMesDif = Format(vlNumMesDif, "00")
        
        'Periodo Garantizado
        If (vlRegistro!Num_MesGar = 0) Then
            vlNumMesGar = vlRegistro!Num_MesGar
        Else
            vlNumMesGar = (vlRegistro!Num_MesGar / 12)
        End If
        vlNumMesGar = Format(vlNumMesGar, "00")
    
        'Fecha de Incorporación
        vlFecAcepta = vlRegistro!Fec_Acepta
        
        'Fecha de Primer Pago
'I--- ABV 23/11/2007 ---
'        vlFecPriPago = vlRegistro!fec_pripago
        If (vlRegistro!Num_MesDif = 0) Then
            vlFecPriPago = vlRegistro!Fec_Dev
        Else
            'vlFecPriPago = vlRegistro!fec_pripago
            'DC 17/06/2009
            vlFecPriPago = DateAdd("yyyy", Val(vlNumMesDif), Mid(vlRegistro!Fec_Dev, 7, 2) & "/" & Mid(vlRegistro!Fec_Dev, 5, 2) & "/" & Mid(vlRegistro!Fec_Dev, 1, 4))
            vlFecPriPago = Format(vlFecPriPago, "YYYYMMDD")
        End If
'F--- ABV 23/11/2007 ---
    
        'Prima Única - Transformar de acuerdo al Tipo de Moneda de la Pensión
        If (vlRegistro!mto_priuni <> 0) Then
            vlMtoPriUni = vlRegistro!mto_priuni / vlRegistro!Mto_ValMoneda
        Else
            vlMtoPriUni = 0
        End If
        vlMtoPriUni = Format(vlMtoPriUni, "0000000000000.00")
        vlA = Mid(vlMtoPriUni, 1, 13)
        vlB = Mid(vlMtoPriUni, 15, 2)
        vlMtoPriUni = vlA & vlB
        
        'Moneda
        
        Dim mon As String
        Select Case vlRegistro!Gls_Moneda
            Case "01"
                 mon = "01"
            Case "02"
                 mon = "03"
            Case "03"
                 mon = "02"
            Case "04"
                 mon = "04"
        End Select
        
        vlEspacios = Space(2 - Len(mon))
        vlGlsMoneda = mon & vlEspacios
        
        'Tipo de Cambio
        vlTipoCambio = Format(vlRegistro!Mto_ValMoneda, "0000.0000")
        vlA = Mid(vlTipoCambio, 1, 4)
        vlB = Mid(vlTipoCambio, 6, 4)
        vlTipoCambio = vlA & vlB
        
        'Estado
        vlEspacios = Space(4 - Len(vlRegistro!GLS_ESTADO))
        vlEstado = (vlRegistro!GLS_ESTADO) & vlEspacios
       
        'Tasa de Cotización
        vlPrcTasaVta = Format((vlRegistro!Prc_TasaVta) / 100, "0.0000")
        vlA = Mid(vlPrcTasaVta, 1, 1)
        vlB = Mid(vlPrcTasaVta, 3, 4)
        vlPrcTasaVta = vlA & vlB

        vlLinea = vlCorrelativo & vlNumPoliza & vlCodCuspp & vlafp & _
        vlFecDev & vlGlsTipPen & vlGlsTipRen & vlMtoPension & _
        vlNumMesDif & vlNumMesGar & vlFecAcepta & vlFecPriPago & _
        vlMtoPriUni & vlGlsMoneda & vlTipoCambio & vlEstado & vlPrcTasaVta & _
        vlFiller1
        
        Print #1, vlLinea
        
        vgI = vgI + 1
        vlRegistro.MoveNext
    Wend
    vlRegistro.Close

    Close #1

    vlOpen = False
    MsgBox "La Exportación de datos al Archivo ha sido finalizada exitosamente.", vbInformation, "Estado de Generación Archivo"

    Screen.MousePointer = 0

Exit Function
Err_flInformeAnexoSVS:
    Screen.MousePointer = 0
    If vlOpen Then
        Close #1
    End If
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_Buscar
    
    'Validar el Ingreso de Fechas
    If fgValidaFecha(Trim(Txt_Desde)) = False Then
        Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
        Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
    Else
        'MsgBox "Debe ingresar una fecha válida para la Fecha Desde.", vbCritical, "Error de Datos"
        Txt_Desde.SetFocus
        Exit Sub
    End If
   
'I--- ABV 22/10/2007 ---
'    If fgValidaFecha(Trim(Txt_Hasta)) = False Then
'        Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
'        Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
'    Else
'        'MsgBox "Debe ingresar una fecha válida para la Fecha Hasta.", vbCritical, "Error de Datos"
'        Txt_Hasta.SetFocus
'        Exit Sub
'    End If
    If (Trim(Txt_Hasta) = "") Then
        MsgBox "Falta Ingresar Fecha", vbExclamation, "Error de Datos"
        Txt_Hasta.SetFocus
        Exit Sub
    End If
    If Not IsDate(Txt_Hasta) Then
        MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
        Txt_Hasta.SetFocus
        Exit Sub
    End If
    If (Year(CDate(Txt_Hasta)) < 1900) Then
        MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
        Txt_Hasta.SetFocus
        Exit Sub
    End If
'F--- ABV 22/10/2007 ---
    
    vlFechaDesde = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaHasta = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    vlSwExiste = False
    
    'Validación Lógica entre las Fechas
    If (vlFechaDesde > vlFechaHasta) Then
        MsgBox "La Fecha Desde es mayor a la Fecha Hasta. Favor volver a ingresarlas.", vbCritical, "Error de Datos"
        Exit Sub
    End If
    
    vgI = MsgBox("Desea generar la información para este Periodo ?", 4 + 64, "Generación de Periodo")
    If (vgI <> 6) Then
        Exit Sub
    End If
    
    'Determinar la Existencia de Información a las Fechas Indicadas
    vgQuery = "SELECT num_poliza FROM pd_tmae_oripoliza "
    vgQuery = vgQuery & "WHERE fec_acepta BETWEEN "
    vgQuery = vgQuery & "'" & vlFechaDesde & "' AND '" & vlFechaHasta & "' "
    vgQuery = vgQuery & "UNION "
    vgQuery = vgQuery & "SELECT num_poliza FROM pd_tmae_poliza "
    vgQuery = vgQuery & "WHERE fec_acepta BETWEEN "
    vgQuery = vgQuery & "'" & vlFechaDesde & "' AND '" & vlFechaHasta & "' "
    'RRR 07052013
    vgQuery = vgQuery & "and num_endoso=1 "
    'RRR
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not vgRs.EOF Then
        vlSwExiste = True
    End If
    vgRs.Close
    
    If (vlSwExiste = False) Then
        MsgBox "No Existe información de Pólizas Aceptadas o Incorporadas para el Rango de Fechas.", vbExclamation, "Inexistencia de Pólizas"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    'Borrar el Contenido de la Tabla Temporal por Usuario
    Sql = "DELETE FROM PD_TTMP_SBSTASACOT WHERE cod_usuario = '" & vgUsuario & "'"
    vgConexionBD.Execute Sql
    
    'Cargar la Tabla Temporal con los casos a Visualizar
    Sql = "SELECT num_poliza,cod_cuspp,num_cot,num_correlativo,num_operacion,"
    Sql = Sql & "ind_cob,cod_cobercon,cod_dercre,cod_dergra,"
    Sql = Sql & "cod_afp,af.cod_adicional AS gls_afp,"
    Sql = Sql & "fec_solicitud,fec_vigencia,fec_pripago,fec_dev,fec_acepta,fec_inipencia,"
    Sql = Sql & "cod_tippension,tp.cod_adicional AS gls_tippension,"
    Sql = Sql & "cod_moneda,decode(cod_moneda||cod_tipreajuste,'NS1', '01','US0', '03', 'NS2','02', '04') AS gls_moneda,mto_valmoneda,"
    Sql = Sql & "cod_tipren,tr.cod_adicional AS gls_tipren,"
    Sql = Sql & "cod_modalidad,num_mesdif,num_mesgar,prc_tasace,"
    Sql = Sql & "prc_tasavta,mto_priuni, "
    Sql = Sql & "DECODE(o.cod_tippension,'08',(select max(z.mto_pension) from pd_tmae_oripolben z where z.num_poliza=o.num_poliza and cod_par<>'99'), mto_pension) as mto_pension,"
    Sql = Sql & "'" & clEstadoSinTraspasoPrima & "' AS cod_estado "
    Sql = Sql & ",al.cod_adicional AS gls_modalidad "
    Sql = Sql & "FROM "
    Sql = Sql & "pd_tmae_oripoliza o, ma_tpar_tabcod af,"
    Sql = Sql & "ma_tpar_tabcod tp, ma_tpar_tabcod tr, ma_tpar_tabcod tm "
    Sql = Sql & ",ma_tpar_tabcod al "
    Sql = Sql & "WHERE "
    Sql = Sql & "fec_acepta BETWEEN '" & vlFechaDesde & "' AND '" & vlFechaHasta & "' AND "
    Sql = Sql & "(af.cod_tabla(+) = 'AF' AND o.cod_afp = af.cod_elemento(+)) AND "
    Sql = Sql & "(tp.cod_tabla(+) = 'TP' AND o.cod_tippension = tp.cod_elemento(+)) AND "
    Sql = Sql & "(tr.cod_tabla(+) = 'TR' AND o.cod_tipren = tr.cod_elemento(+)) AND "
    Sql = Sql & "(tm.cod_tabla(+) = 'TM' AND o.cod_moneda = tm.cod_elemento(+)) "
    Sql = Sql & "AND (al.cod_tabla(+) = 'AL' AND o.cod_modalidad = al.cod_elemento(+)) "
    Sql = Sql & "UNION  "
    Sql = Sql & "SELECT num_poliza,cod_cuspp,num_cot,num_correlativo,num_operacion,"
    Sql = Sql & "ind_cob,cod_cobercon,cod_dercre,cod_dergra,"
    Sql = Sql & "cod_afp,af.cod_adicional AS gls_afp,"
    Sql = Sql & "fec_solicitud,fec_vigencia,fec_pripago,fec_dev,fec_acepta,fec_inipencia,"
    Sql = Sql & "cod_tippension,tp.cod_adicional AS gls_tippension,"
    Sql = Sql & "cod_moneda,decode(cod_moneda||cod_tipreajuste,'NS1', '01','US0', '03', 'NS2','02', '04') AS gls_moneda,mto_valmoneda,"
    Sql = Sql & "cod_tipren,tr.cod_adicional AS gls_tipren,"
    Sql = Sql & "cod_modalidad,num_mesdif,num_mesgar,prc_tasace,"
    Sql = Sql & "prc_tasavta,mto_priuni, "
    Sql = Sql & "DECODE(o.cod_tippension,'08',(select max(z.mto_pension) from pd_tmae_polben z where z.num_poliza=o.num_poliza and cod_par<>'99'), mto_pension) as mto_pension,"
    Sql = Sql & "'" & clEstadoConTraspasoPrima & "' AS cod_estado "
    Sql = Sql & ",al.cod_adicional AS gls_modalidad "
    Sql = Sql & "FROM "
    Sql = Sql & "pd_tmae_poliza o, ma_tpar_tabcod af, "
    Sql = Sql & "ma_tpar_tabcod tp, ma_tpar_tabcod tr, ma_tpar_tabcod tm "
    Sql = Sql & ",ma_tpar_tabcod al "
    Sql = Sql & "WHERE "
    Sql = Sql & "fec_acepta BETWEEN '" & vlFechaDesde & "' AND '" & vlFechaHasta & "' AND "
    Sql = Sql & "(af.cod_tabla(+) = 'AF' AND o.cod_afp = af.cod_elemento(+)) AND "
    Sql = Sql & "(tp.cod_tabla(+) = 'TP' AND o.cod_tippension = tp.cod_elemento(+)) AND "
    Sql = Sql & "(tr.cod_tabla(+) = 'TR' AND o.cod_tipren = tr.cod_elemento(+)) AND "
    Sql = Sql & "(tm.cod_tabla(+) = 'TM' AND o.cod_moneda = tm.cod_elemento(+)) "
    Sql = Sql & "AND (al.cod_tabla(+) = 'AL' AND o.cod_modalidad = al.cod_elemento(+)) "
    Sql = Sql & "AND num_endoso=1 "
    Sql = Sql & "ORDER BY num_poliza"
    Set vgRs = vgConexionBD.Execute(Sql)
    While Not vgRs.EOF
        
        Sql = "INSERT INTO PD_TTMP_SBSTASACOT ("
        Sql = Sql & "COD_USUARIO,NUM_POLIZA,NUM_ENDOSO,"
        Sql = Sql & "NUM_COT,NUM_CORRELATIVO,NUM_OPERACION,"
        Sql = Sql & "COD_AFP,COD_TIPPENSION,COD_CUSPP,FEC_INI,FEC_FIN,"
        Sql = Sql & "FEC_SOLICITUD,FEC_VIGENCIA,FEC_DEV,FEC_ACEPTA,FEC_PRIPAGO,"
        Sql = Sql & "MTO_PRIUNI,IND_COB,COD_MONEDA,MTO_VALMONEDA,"
        Sql = Sql & "COD_TIPREN,NUM_MESDIF,COD_MODALIDAD,NUM_MESGAR,"
        Sql = Sql & "COD_COBERCON,COD_DERCRE,COD_DERGRA,"
        Sql = Sql & "PRC_TASACE,PRC_TASAVTA,MTO_PENSION,"
        Sql = Sql & "GLS_AFP,GLS_TIPPENSION,GLS_TIPREN,GLS_MONEDA,GLS_ESTADO"
        Sql = Sql & ")VALUES("
        Sql = Sql & "'" & vgUsuario & "',"
        Sql = Sql & "'" & vgRs!Num_Poliza & "',"
        Sql = Sql & " " & "1" & ","
        Sql = Sql & "'" & vgRs!Num_Cot & "',"
        Sql = Sql & " " & vgRs!Num_Correlativo & ","
        Sql = Sql & "'" & vgRs!Num_Operacion & "',"
        Sql = Sql & "'" & vgRs!Cod_AFP & "',"
        Sql = Sql & "'" & vgRs!Cod_TipPension & "',"
        Sql = Sql & "'" & vgRs!Cod_Cuspp & "',"
        Sql = Sql & "'" & vlFechaDesde & "',"
        Sql = Sql & "'" & vlFechaHasta & "',"
        Sql = Sql & "'" & vgRs!Fec_Solicitud & "',"
        Sql = Sql & "'" & vgRs!Fec_Vigencia & "',"
        Sql = Sql & "'" & vgRs!Fec_Dev & "',"
        Sql = Sql & "'" & vgRs!Fec_Acepta & "',"
'I--- ABV 23/11/2007 ---
'        Sql = Sql & "'" & vgRs!fec_pripago & "',"
        Sql = Sql & "'" & vgRs!fec_inipencia & "',"
'F--- ABV 23/11/2007 ---
        Sql = Sql & " " & Str(vgRs!mto_priuni) & ","
        Sql = Sql & "'" & vgRs!Ind_Cob & "',"
        Sql = Sql & "'" & vgRs!Gls_Moneda & "',"
        Sql = Sql & " " & Str(vgRs!Mto_ValMoneda) & ","
        Sql = Sql & "'" & vgRs!Cod_TipRen & "',"
        Sql = Sql & " " & vgRs!Num_MesDif & ","
        Sql = Sql & "'" & vgRs!Cod_Modalidad & "',"
        Sql = Sql & " " & vgRs!Num_MesGar & ","
        Sql = Sql & "'" & vgRs!Cod_CoberCon & "',"
        Sql = Sql & "'" & vgRs!Cod_DerCre & "',"
        Sql = Sql & "'" & vgRs!Cod_DerGra & "',"
        Sql = Sql & " " & Str(vgRs!Prc_TasaCe) & ","
        Sql = Sql & " " & Str(vgRs!Prc_TasaVta) & ","
        Sql = Sql & " " & Str(vgRs!Mto_Pension) & ","
        Sql = Sql & "'" & Trim(vgRs!GLS_AFP) & "',"
        If (vgRs!Cod_TipPension = "04") Or (vgRs!Cod_TipPension = "05") Then
            Sql = Sql & "'" & Trim(vgRs!GLS_TIPPENSION) & "',"
        Else
            If (vgRs!Ind_Cob = "S") Then
                Sql = Sql & "'" & Trim(vgRs!GLS_TIPPENSION) & clCoberturaCon & "',"
            Else
                Sql = Sql & "'" & Trim(vgRs!GLS_TIPPENSION) & clCoberturaSin & "',"
            End If
        End If
        If IsNull(vgRs!Gls_Modalidad) Then
            Sql = Sql & "'" & Trim(vgRs!GLS_TIPREN) & "',"
        Else
            Sql = Sql & "'" & Trim(vgRs!GLS_TIPREN) & Trim(vgRs!Gls_Modalidad) & "',"
        End If
        Sql = Sql & "'" & Trim(vgRs!Gls_Moneda) & "',"
        Sql = Sql & "'" & Trim(vgRs!Cod_Estado) & "'"
        Sql = Sql & ")"
        vgConexionBD.Execute Sql
        vgRs.MoveNext
    Wend
    vgRs.Close
    
    'Calcular RESUMEN de PANTALLA
    Call plCalcularResumen
    
    MsgBox "El Proceso de Generación de la Información ha finalizado Exitosamente.", vbInformation, "Estado del Proceso"
    
    Cmd_Imprimir_reporte.SetFocus
    
    Screen.MousePointer = 0

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]" & vgRs!Num_Poliza, vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Cargar_Click()
On Error GoTo Err_Cargar

    'Validar la Existencia de un Periodo
    If (flValidarExistenciaPeriodo(vlFechaDesde, vlFechaHasta) = False) Then
        MsgBox "No se ha generado el proceso de Selección o Búsqueda de Pólizas.", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    'Refrescar la Pantalla
    If (flRefrescarPeriodo = False) Then
        MsgBox "Ha ocurrido un error en la Actualización de la Información a Pantalla.", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    vlDesde = Trim(Txt_Desde)
    vlHasta = Trim(Txt_Hasta)

    'Selección del Archivo de Resumen de Reservas
    ComDialogo.CancelError = True
    'I - MC 17/06/2008 Se modifica el nombre del archivo Sucave
    'ComDialogo.FileName = "FOR0067_" & vlFechaDesde & ".067"
    ComDialogo.FileName = clCodArchSucave_01 & Mid(vlFechaHasta, 3, 6) & ".067"
    'F - MC 17/06/2008
    ComDialogo.DialogTitle = "Guardar Pólizas de Rtas. Vit. del Mes como"
    ComDialogo.Filter = "*.067"
    ComDialogo.FilterIndex = 1
    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    ComDialogo.ShowSave
    vlArchivo = ComDialogo.FileName
    If vlArchivo = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    'Permite generar el Archivo con la Opción Indicada a través del Menú
    Call flArchivoAnexoSVS

    Screen.MousePointer = 0

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    If Err.Number = 32755 Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
        End If
    End If
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir

    'Validar la Existencia de un Periodo
    If (flValidarExistenciaPeriodo(vlFechaDesde, vlFechaHasta) = False) Then
        MsgBox "No se ha generado el proceso de Selección o Búsqueda de Pólizas.", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    'Refrescar la Pantalla
    If (flRefrescarPeriodo = False) Then
        MsgBox "Ha ocurrido un error en la Actualización de la Información a Pantalla.", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    vlDesde = Trim(Txt_Desde)
    vlHasta = Trim(Txt_Hasta)

    vlArchivo = strRpt & "PD_Rpt_SBSTasaCot.rpt"   '\Reportes
    If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Consulta de Pólizas Traspasadas no se encuentra en el Directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
    End If

'    vgQuery = "{pd_TTMP_SBSTASACOT.FEC_INI} >= '" & Trim(vlFechaDesde) & "' AND "
'    vgQuery = vgQuery & "{pd_TTMP_SBSTASACOT.FEC_FIN} <= '" & Trim(vlFechaHasta) & "'"

    vgQuery = "{PD_TTMP_SBSTASACOT.COD_USUARIO} = '" & vgUsuario & "' "
   
    Rpt_SBS.Reset
    Rpt_SBS.ReportFileName = vlArchivo     'App.Path & "\rpt_Areas.rpt"
    Rpt_SBS.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
    Rpt_SBS.SelectionFormula = vgQuery
    Rpt_SBS.Formulas(0) = ""
    Rpt_SBS.Formulas(1) = ""
    Rpt_SBS.Formulas(0) = "FechaDesde = '" & vlDesde & "'"
    Rpt_SBS.Formulas(1) = "FechaHasta = '" & vlHasta & "'"
    Rpt_SBS.WindowState = crptMaximized
    Rpt_SBS.Destination = crptToWindow
    Rpt_SBS.WindowTitle = "Informe de Tasas para la Superintendencia"
    Rpt_SBS.Action = 1
    Screen.MousePointer = 0
   
Exit Sub
Err_Imprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Imprimir_reporte_Click()

On Error GoTo Err_Imprimir

Dim r_temp As ADODB.Recordset

    'Validar la Existencia de un Periodo
    If (flValidarExistenciaPeriodo(vlFechaDesde, vlFechaHasta) = False) Then
        MsgBox "No se ha generado el proceso de Selección o Búsqueda de Pólizas.", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    'Refrescar la Pantalla
    If (flRefrescarPeriodo = False) Then
        MsgBox "Ha ocurrido un error en la Actualización de la Información a Pantalla.", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    vlDesde = Trim(Txt_Desde)
    vlHasta = Trim(Txt_Hasta)

    vlArchivo = strRpt & "PD_Rpt_SBSTasaCot.rpt"   '\Reportes
    If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Consulta de Pólizas Traspasadas no se encuentra en el Directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
    End If

    Call p_crear_rs

    sSql = "select * from PD_TTMP_SBSTASACOT where COD_USUARIO = '" & vgUsuario & "' "
   
    Set r_temp = New ADODB.Recordset
    Set r_temp = vgConexionBD.Execute(sSql)
    If Not r_temp.EOF Then
        Do Until r_temp.EOF
            objRsRpt.AddNew
            objRsRpt.Fields("COD_USUARIO").Value = Trim(r_temp!COD_USUARIO)
            objRsRpt.Fields("NUM_POLIZA").Value = Trim(r_temp!Num_Poliza)
            objRsRpt.Fields("COD_CUSPP").Value = Trim(r_temp!Cod_Cuspp)
            objRsRpt.Fields("FEC_DEV").Value = Trim(r_temp!Fec_Dev)
            objRsRpt.Fields("FEC_ACEPTA").Value = Trim(r_temp!Fec_Acepta)
            objRsRpt.Fields("MTO_PRIUNI").Value = r_temp!mto_priuni
            objRsRpt.Fields("MTO_VALMONEDA").Value = r_temp!Mto_ValMoneda
            objRsRpt.Fields("NUM_MESDIF").Value = Trim(r_temp!Num_MesDif)
            objRsRpt.Fields("NUM_MESGAR").Value = Trim(r_temp!Num_MesGar)
            objRsRpt.Fields("PRC_TASAVTA").Value = Trim(r_temp!Prc_TasaVta)
            objRsRpt.Fields("MTO_PENSION").Value = Trim(r_temp!Mto_Pension)
            objRsRpt.Fields("GLS_AFP").Value = r_temp!GLS_AFP
            objRsRpt.Fields("GLS_TIPPENSION").Value = r_temp!GLS_TIPPENSION
            objRsRpt.Fields("GLS_TIPREN").Value = r_temp!GLS_TIPREN
            
            
            'objRsRpt.Fields("GLS_MONEDA").Value = r_temp!Gls_Moneda
            
            Select Case r_temp!Gls_Moneda
            Case "01"
                objRsRpt.Fields("GLS_MONEDA").Value = "01"
            Case "02"
                objRsRpt.Fields("GLS_MONEDA").Value = "03"
            Case "03"
                objRsRpt.Fields("GLS_MONEDA").Value = "02"
            Case "04"
                objRsRpt.Fields("GLS_MONEDA").Value = "04"
            End Select
            
            
            objRsRpt.Fields("GLS_ESTADO").Value = r_temp!GLS_ESTADO
            objRsRpt.Update
            r_temp.MoveNext
        Loop
        sName_Reporte = "PD_Rpt_SBSTasaCot.rpt"
        frm_plantilla.Show
    End If
   
    Screen.MousePointer = 0
   
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

    Call flLimpiar

Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Refrescar_Click()
On Error GoTo Err_Descargar

    If (flRefrescarPeriodo = True) Then
        MsgBox "La Información ha sido Actualizada.", vbInformation, "Estado del Proceso"
    Else
        Call flLimpiar
    End If

Exit Sub
Err_Descargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
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

Private Sub Form_Load()
On Error GoTo Err_Cargar
    
    Me.Top = 0
    Me.Left = 0
    
    If (flRefrescarPeriodo = True) Then
        'MsgBox "La Información ha sido Actualizada.", vbInformation, "Estado del Proceso"
    Else
        Call flLimpiar
    End If
    
Exit Sub
Err_Cargar:
Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
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

'    vlDesde = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
'    vlHasta = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")
'
'    'Obtiene la cantidad de polizas, a partir de un rango de fechas
'    Sql = "select count(num_poliza)as NumeroPoliza from pd_ttmp_sbstasacot"
'    Sql = Sql & " where fec_ini> = '" & vlDesde & "'"
'    Sql = Sql & " and fec_fin< = '" & vlHasta & "'"
'
'    Set vgRs = vgConexionBD.Execute(Sql)
'    If Not vgRs.EOF Then
'        Lbl_TotPol = Trim(vgRs!numeropoliza)
'    End If
'
'    'Obtiene la suma de las primas, a partir de un rango de fechas
'    Sql = "select sum(mto_priuni)as SumaPrima from pd_ttmp_sbstasacot"
'    Sql = Sql & " where fec_ini>= '" & vlDesde & "'"
'    Sql = Sql & " and fec_fin<= '" & vlHasta & "'"
'
'    Set vgRs = vgConexionBD.Execute(Sql)
'    If Not vgRs.EOF Then
'        Lbl_MtoTotPri = Format(vgRs!sumaprima, "#,#0.00")
'    End If
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

Private Sub p_crear_rs()
   
   Set objRsRpt = New ADODB.Recordset
   
   objRsRpt.Fields.Append "COD_USUARIO", adVarChar, 10
   objRsRpt.Fields.Append "NUM_POLIZA", adVarChar, 10
   objRsRpt.Fields.Append "COD_CUSPP", adVarChar, 12
   objRsRpt.Fields.Append "FEC_DEV", adVarChar, 8
   objRsRpt.Fields.Append "FEC_ACEPTA", adVarChar, 8
   objRsRpt.Fields.Append "MTO_PRIUNI", adDouble
   objRsRpt.Fields.Append "MTO_VALMONEDA", adDouble
   objRsRpt.Fields.Append "NUM_MESDIF", adInteger
   objRsRpt.Fields.Append "NUM_MESGAR", adInteger
   objRsRpt.Fields.Append "PRC_TASAVTA", adDouble
   objRsRpt.Fields.Append "MTO_PENSION", adDouble
   objRsRpt.Fields.Append "GLS_AFP", adVarChar, 10
   objRsRpt.Fields.Append "GLS_TIPPENSION", adVarChar, 10
   objRsRpt.Fields.Append "GLS_TIPREN", adVarChar, 10
   objRsRpt.Fields.Append "GLS_MONEDA", adVarChar, 10
   objRsRpt.Fields.Append "GLS_ESTADO", adVarChar, 10
      
   objRsRpt.Open
  

End Sub

