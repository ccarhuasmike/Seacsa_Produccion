VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_CalInfInt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Interno para Compañía."
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10335
   Begin Crystal.CrystalReport Rpt_CIA 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   10095
      Begin VB.CommandButton Cmd_Refrescar 
         Caption         =   "&Refrescar"
         Height          =   675
         Left            =   3600
         Picture         =   "Frm_CalInfInt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Restaurar o refrescar los datos desde la BD"
         Top             =   240
         Width           =   790
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2520
         Picture         =   "Frm_CalInfInt.frx":0572
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4800
         Picture         =   "Frm_CalInfInt.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5880
         Picture         =   "Frm_CalInfInt.frx":12E6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
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
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   10095
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   6720
         Picture         =   "Frm_CalInfInt.frx":13E0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Efectuar Consulta"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Opt_Todos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   1320
         Width           =   2775
      End
      Begin VB.OptionButton Opt_ConPrimaRec 
         Caption         =   "Con Prima Recepcionada"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton Opt_SinPrimaRec 
         Caption         =   "Sin Prima Recepcionada"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   840
         Value           =   -1  'True
         Width           =   2775
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
         Left            =   4440
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo Selección            :"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Periodo de Incorporación de Pólizas"
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
         Left            =   2880
         TabIndex        =   12
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rango de Fechas        :"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   3495
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   1
      Cols            =   18
      BackColor       =   14745599
      FormatString    =   $"Frm_CalInfInt.frx":14E2
   End
End
Attribute VB_Name = "Frm_CalInfInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const clEstadoSinTraspasoPrima As String * 3 = "PNT"
Const clEstadoConTraspasoPrima As String * 2 = "PT"

Dim vlRegistro As ADODB.Recordset

Dim vlFechaDesde As String, vlFechaHasta As String
Dim vlDesde As String, vlHasta As String
Dim vlTipPension As String
Dim vlFechaAcepta As String, vlFechaDev As String
Dim vlSexoCausante As String, vlFechaNacCausante As String
Dim vlSexoConyuge As String, vlFechaNacConyuge As String
Dim vlPrimaUnica As String, vlTipoCambio As String, vlPensionMes As String
Dim vlTaCtoEqui As String, vlTasaVta As String, vlComision As String
Dim vlPerDif As String, vlPerGar As String
Dim vlElemento As String

Dim vlGlosaMoneda As String 'I--- ABV 05/02/2011 ---

Function flLimpiar()
    Txt_Desde.Text = ""
    Txt_Hasta.Text = ""
    Msf_Grilla.Rows = 1
End Function

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

'FUNCION QUE INICIA LA GRILLA
Function flIniciaGrilla()
On Error GoTo Err_IniGrilla
    
    Msf_Grilla.Clear
    Msf_Grilla.Rows = 1
    Msf_Grilla.Cols = 18
    
    Msf_Grilla.Row = 0
    
    Msf_Grilla.Col = 0
    Msf_Grilla.ColWidth(0) = 800
    Msf_Grilla.Text = "Moneda"
 
    Msf_Grilla.Col = 1
    Msf_Grilla.ColWidth(1) = 1200
    Msf_Grilla.Text = "N° Póliza"
    
    Msf_Grilla.Col = 2
    Msf_Grilla.ColWidth(2) = 1500
    Msf_Grilla.Text = "CUSPP"
    
    Msf_Grilla.Col = 3
    Msf_Grilla.ColWidth(3) = 2500
    Msf_Grilla.Text = "Tipo Renta"
    
    Msf_Grilla.Col = 4
    Msf_Grilla.ColWidth(4) = 1100
    Msf_Grilla.Text = "F.Nac.Cau."
    
    Msf_Grilla.Col = 5
    Msf_Grilla.ColWidth(5) = 1100
    Msf_Grilla.Text = "F.Incorp."
    
    Msf_Grilla.Col = 6
    Msf_Grilla.ColWidth(6) = 1000
    Msf_Grilla.Text = "Sexo Cau."
    
    Msf_Grilla.Col = 7
    Msf_Grilla.ColWidth(7) = 1100
    Msf_Grilla.Text = "F.Nac.Cóny."
    
    Msf_Grilla.Col = 8
    Msf_Grilla.ColWidth(8) = 1000
    Msf_Grilla.Text = "Sexo Cóny."
    
    Msf_Grilla.Col = 9
    Msf_Grilla.ColWidth(9) = 1200
    Msf_Grilla.Text = "Prima (S/.)"
    
    Msf_Grilla.Col = 10
    Msf_Grilla.ColWidth(10) = 1000
    Msf_Grilla.Text = "Tipo Cambio"
    
    Msf_Grilla.Col = 11
    Msf_Grilla.ColWidth(11) = 1000
    Msf_Grilla.Text = "Pensión"
    
    Msf_Grilla.Col = 12
    Msf_Grilla.ColWidth(12) = 800
    Msf_Grilla.Text = "Per.Dif."
    
    Msf_Grilla.Col = 13
    Msf_Grilla.ColWidth(13) = 800
    Msf_Grilla.Text = "Per.Gar."
    
    Msf_Grilla.Col = 14
    Msf_Grilla.ColWidth(14) = 800
    Msf_Grilla.Text = "Tasa CE"
    
    Msf_Grilla.Col = 15
    Msf_Grilla.ColWidth(15) = 800
    Msf_Grilla.Text = "Tasa Vta"
    
    Msf_Grilla.Col = 16
    Msf_Grilla.ColWidth(16) = 800
    Msf_Grilla.Text = "Comisión"
    
    Msf_Grilla.Col = 17
    Msf_Grilla.ColWidth(17) = 1100
    Msf_Grilla.Text = "F.Devengue"

Exit Function
Err_IniGrilla:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Function

Function flCargarGrilla()
    
    Msf_Grilla.Rows = 1
    
'I--- ABV 05/02/2011 ---
'    Sql = "SELECT cod_moneda,num_poliza,cod_cuspp,"
    Sql = "SELECT gls_moneda as cod_moneda,num_poliza,cod_cuspp,"
'F--- ABV 05/02/2011 ---
    Sql = Sql & "cod_tippension,fec_naccau,fec_acepta,"
    Sql = Sql & "cod_sexocau,fec_naccon,cod_sexocon,mto_priuni, "
    Sql = Sql & "mto_valmoneda,mto_pension,num_mesdif,num_mesgar, "
    Sql = Sql & "prc_tasace,prc_tasavta,prc_corcom,fec_dev "
    Sql = Sql & "FROM pd_ttmp_ciatasacot "
    Sql = Sql & "WHERE "
    Sql = Sql & "cod_usuario = '" & vgUsuario & "' "
    If Opt_SinPrimaRec.Value = True Then
        Sql = Sql & "AND gls_estado = '" & clEstadoSinTraspasoPrima & "' "
    End If
    If Opt_ConPrimaRec.Value = True Then
        Sql = Sql & "AND gls_estado = '" & clEstadoConTraspasoPrima & "' "
    End If
    Sql = Sql & "ORDER BY num_poliza "
    Set vgRs = vgConexionBD.Execute(Sql)
    While Not vgRs.EOF
        
        vlFechaAcepta = DateSerial(Mid((vgRs!Fec_Acepta), 1, 4), Mid((vgRs!Fec_Acepta), 5, 2), Mid(vgRs!Fec_Acepta, 7, 2))
        vlFechaDev = DateSerial(Mid((vgRs!Fec_Dev), 1, 4), Mid((vgRs!Fec_Dev), 5, 2), Mid((vgRs!Fec_Dev), 7, 2))
        vlFechaNacCausante = DateSerial(Mid((vgRs!fec_naccau), 1, 4), Mid((vgRs!fec_naccau), 5, 2), Mid((vgRs!fec_naccau), 7, 2))
        vlSexoCausante = Trim(vgRs!cod_Sexocau)
        
        If Not IsNull(vgRs!fec_naccon) Then
            vlFechaNacConyuge = DateSerial(Mid((vgRs!fec_naccon), 1, 4), Mid((vgRs!fec_naccon), 5, 2), Mid((vgRs!fec_naccon), 7, 2))
            vlSexoConyuge = Trim(vgRs!cod_Sexocon)
        Else
            vlFechaNacConyuge = ""
            vlSexoConyuge = ""
        End If
        
        vlTipPension = fgBuscarGlosaElemento(vgCodTabla_TipPen, vgRs!Cod_TipPension)
        vlPrimaUnica = Format(vgRs!mto_priuni, "#,#0.00")
        vlTipoCambio = Format(vgRs!Mto_ValMoneda, "#,#0.00")
        vlPensionMes = Format(vgRs!Mto_Pension, "#,#0.00")
        vlTaCtoEqui = Format(vgRs!Prc_TasaCe, "#0.00")
        vlTasaVta = Format(vgRs!Prc_TasaVta, "#0.00")
        vlComision = Format(vgRs!Prc_CorCom, "#0.00")
        If (vgRs!Num_MesDif = 0) Then
            vlPerDif = 0
        Else
            vlPerDif = Format((vgRs!Num_MesDif) / 12, "#0")
        End If
        If (vgRs!Num_MesGar = 0) Then
            vlPerGar = 0
        Else
            vlPerGar = Format((vgRs!Num_MesGar) / 12, "#0")
        End If
        
        Msf_Grilla.AddItem (vgRs!Cod_Moneda) & vbTab & _
            (vgRs!Num_Poliza) & vbTab & _
            " " & (vgRs!Cod_Cuspp) & vbTab & _
            (vlTipPension) & vbTab & _
            (vlFechaNacCausante) & vbTab & _
            (vlFechaAcepta) & vbTab & _
            (vlSexoCausante) & vbTab & _
            (vlFechaNacConyuge) & vbTab & _
            (vlSexoConyuge) & vbTab & _
            (vlPrimaUnica) & vbTab & _
            (vlTipoCambio) & vbTab & _
            (vlPensionMes) & vbTab & _
            (vlPerDif) & vbTab & _
            (vlPerGar) & vbTab & _
            (vlTaCtoEqui) & vbTab & _
            (vlTasaVta) & vbTab & _
            (vlComision) & vbTab & _
            (vlFechaDev)
        vgRs.MoveNext
    
    Wend
    vgRs.Close
    
End Function

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
    vgQuery = vgQuery & "FROM pd_ttmp_ciatasacot "
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
        
        Call flCargarGrilla
    End If
    
    flRefrescarPeriodo = True
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
    If fgValidaFecha(Trim(Txt_Hasta)) = False Then
        Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
        Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
    Else
        'MsgBox "Debe ingresar una fecha válida para la Fecha Hasta.", vbCritical, "Error de Datos"
        Txt_Hasta.SetFocus
        Exit Sub
    End If
    
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
    Sql = "DELETE FROM PD_TTMP_CIATASACOT WHERE cod_usuario = '" & vgUsuario & "'"
    vgConexionBD.Execute Sql
    
    'Cargar la Tabla Temporal con los casos a Visualizar
    Sql = " SELECT o.num_poliza,o.cod_cuspp,o.num_cot,o.num_correlativo,"
    Sql = Sql & "o.num_operacion,o.ind_cob,o.cod_cobercon,o.cod_dercre,o.cod_dergra, "
    Sql = Sql & "o.cod_afp,o.fec_solicitud,o.fec_vigencia,o.fec_pripago,o.fec_dev,"
    Sql = Sql & "o.fec_acepta,o.fec_inipencia,o.cod_tippension, "
'I--- ABV 05/02/2011 ---
'    Sql = Sql & "o.cod_moneda,tm.cod_scomp AS gls_moneda,"
    Sql = Sql & "o.cod_moneda,tm.cod_scomp AS gls_monedapol, "
    Sql = Sql & "mtr.cod_scomp as gls_moneda, "
'F--- ABV 05/02/2011 ---
    Sql = Sql & "o.mto_valmoneda, "
    Sql = Sql & "o.cod_tipren, o.cod_modalidad,o.num_mesdif,o.num_mesgar,"
    Sql = Sql & "o.prc_tasace,o.prc_tasavta,o.mto_priuni,o.mto_pension, "
    Sql = Sql & "'" & clEstadoSinTraspasoPrima & "' AS cod_estado ,o.prc_corcom,"
    Sql = Sql & "b.fec_nacben,b.cod_sexo, 'O' as TipoTabla "
    Sql = Sql & "FROM "
    Sql = Sql & "pd_tmae_oripoliza o, pd_tmae_oripolben b, "
    Sql = Sql & "ma_tpar_tabcod tm "
'I--- ABV 05/02/2011 ---
    Sql = Sql & ",ma_tpar_monedatiporeaju mtr "
'F--- ABV 05/02/2011 ---
    Sql = Sql & "WHERE "
    Sql = Sql & "o.num_poliza = b.num_poliza AND "
    Sql = Sql & "fec_acepta BETWEEN '" & vlFechaDesde & "' AND '" & vlFechaHasta & "' AND "
    Sql = Sql & "(tm.cod_tabla(+) = 'TM' AND o.cod_moneda = tm.cod_elemento(+)) "
    Sql = Sql & "AND b.cod_par = '99' "
'I--- ABV 05/02/2011 ---
    Sql = Sql & "AND o.cod_tipreajuste = mtr.cod_tipreajuste(+) AND "
    Sql = Sql & "o.cod_moneda = mtr.cod_moneda(+) "
'F--- ABV 05/02/2011 ---
    Sql = Sql & "UNION "
    Sql = Sql & "SELECT o.num_poliza,o.cod_cuspp,o.num_cot,o.num_correlativo,"
    Sql = Sql & "o.num_operacion,o.ind_cob,o.cod_cobercon,o.cod_dercre,o.cod_dergra, "
    Sql = Sql & "o.cod_afp,o.fec_solicitud,o.fec_vigencia,o.fec_pripago,o.fec_dev,"
    Sql = Sql & "o.fec_acepta,o.fec_inipencia,o.cod_tippension, "
'I--- ABV 05/02/2011 ---
'    Sql = Sql & "o.cod_moneda,tm.cod_scomp AS gls_moneda,"
    Sql = Sql & "o.cod_moneda,tm.cod_scomp AS gls_monedapol, "
    Sql = Sql & "mtr.cod_scomp as gls_moneda, "
'F--- ABV 05/02/2011 ---
    Sql = Sql & "o.mto_valmoneda, "
    Sql = Sql & "o.cod_tipren, o.cod_modalidad,o.num_mesdif,o.num_mesgar,"
    Sql = Sql & "o.prc_tasace,o.prc_tasavta,o.mto_priuni,o.mto_pension, "
    Sql = Sql & "'" & clEstadoConTraspasoPrima & "' AS cod_estado ,o.prc_corcom,"
    Sql = Sql & "b.fec_nacben,b.cod_sexo, 'P' as TipoTabla "
    Sql = Sql & "FROM "
    Sql = Sql & "pd_tmae_poliza o, pd_tmae_polben b, "
    Sql = Sql & "ma_tpar_tabcod tm "
'I--- ABV 05/02/2011 ---
    Sql = Sql & ",ma_tpar_monedatiporeaju mtr "
'F--- ABV 05/02/2011 ---
    Sql = Sql & "WHERE "
    Sql = Sql & "o.num_poliza = b.num_poliza AND "
    Sql = Sql & "fec_acepta BETWEEN '" & vlFechaDesde & "' AND '" & vlFechaHasta & "' AND "
    Sql = Sql & "(tm.cod_tabla(+) = 'TM' AND o.cod_moneda = tm.cod_elemento(+)) "
    Sql = Sql & "AND b.cod_par = '99' "
'I--- ABV 05/02/2011 ---
    Sql = Sql & "AND o.cod_tipreajuste = mtr.cod_tipreajuste(+) AND "
    Sql = Sql & "o.cod_moneda = mtr.cod_moneda(+) "
'F--- ABV 05/02/2011 ---
    Sql = Sql & "ORDER BY num_poliza"
    Set vgRs = vgConexionBD.Execute(Sql)
    While Not vgRs.EOF
        
        vlSexoConyuge = ""
        vlFechaNacConyuge = ""
        
        'Buscar la Existencia de una Cónyuge o Madre
        Sql = "SELECT cod_sexo,fec_nacben "
        If (vgRs!TipoTabla = "O") Then
            Sql = Sql & "FROM pd_tmae_oripolben "
        Else
            Sql = Sql & "FROM pd_tmae_polben "
        End If
        Sql = Sql & "WHERE "
        Sql = Sql & "num_poliza = '" & vgRs!Num_Poliza & "' AND "
        Sql = Sql & "cod_par in (10,11,20,21) "
        Set vlRegistro = vgConexionBD.Execute(Sql)
        If Not vlRegistro.EOF Then
            vlSexoConyuge = Trim(vlRegistro!Cod_Sexo)
            vlFechaNacConyuge = vlRegistro!Fec_NacBen
        End If
        vlRegistro.Close

'I--- ABV 05/02/2011 ---
        If IsNull(vgRs!Gls_Moneda) Then
            vlGlosaMoneda = Trim(vgRs!gls_monedapol)
        Else
            vlGlosaMoneda = Trim(vgRs!Gls_Moneda)
        End If
'F--- ABV 05/02/2011 ---
        
        Sql = " INSERT INTO PD_TTMP_CIATASACOT "
        Sql = Sql & "("
        Sql = Sql & "COD_USUARIO,NUM_POLIZA,NUM_ENDOSO,NUM_COT,NUM_CORRELATIVO,"
        Sql = Sql & "NUM_OPERACION,COD_AFP,COD_TIPPENSION,COD_CUSPP,"
        Sql = Sql & "FEC_INI,FEC_FIN,FEC_SOLICITUD,FEC_VIGENCIA,FEC_DEV,"
        Sql = Sql & "FEC_ACEPTA,FEC_PRIPAGO,MTO_PRIUNI,"
        Sql = Sql & "IND_COB,COD_MONEDA,MTO_VALMONEDA,"
        Sql = Sql & "COD_TIPREN,NUM_MESDIF,COD_MODALIDAD,NUM_MESGAR,"
        Sql = Sql & "COD_COBERCON,COD_DERCRE,COD_DERGRA,"
        Sql = Sql & "PRC_TASACE,PRC_TASAVTA,MTO_PENSION,PRC_CORCOM,"
        Sql = Sql & "FEC_NACCAU,COD_SEXOCAU"
        If (vlSexoConyuge <> "") Then
            Sql = Sql & ",FEC_NACCON,COD_SEXOCON"
        End If
        Sql = Sql & ",GLS_ESTADO"
'I--- ABV 05/02/2011 ---
        Sql = Sql & ",GLS_MONEDA "
'F--- ABV 05/02/2011 ---
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
        Sql = Sql & "'" & vgRs!fec_pripago & "',"
        Sql = Sql & " " & Str(vgRs!mto_priuni) & ","
        Sql = Sql & "'" & vgRs!Ind_Cob & "',"
        Sql = Sql & "'" & vgRs!Cod_Moneda & "',"
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
        Sql = Sql & " " & Str(vgRs!Prc_CorCom) & ","
        Sql = Sql & "'" & vgRs!Fec_NacBen & "',"
        Sql = Sql & "'" & vgRs!Cod_Sexo & "'"
        If (vlSexoConyuge <> "") Then
            Sql = Sql & ",'" & vlFechaNacConyuge & "'"
            Sql = Sql & ",'" & vlSexoConyuge & "'"
        End If
        Sql = Sql & ",'" & Trim(vgRs!Cod_Estado) & "'"
'I--- ABV 05/02/2011 ---
        Sql = Sql & ",'" & vlGlosaMoneda & "'"
'F--- ABV 05/02/2011 ---
        Sql = Sql & ")"
        vgConexionBD.Execute Sql
        vgRs.MoveNext
    Wend
    vgRs.Close
    
    MsgBox "El Proceso de Generación de la Información ha finalizado Exitosamente.", vbInformation, "Estado del Proceso"
    
    'Actualizar la Grilla
    Call flCargarGrilla
    
    
    Cmd_Imprimir.SetFocus
    
    Screen.MousePointer = 0

Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
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
    
    vlArchivo = strRpt & "PD_Rpt_CIATasaCot.rpt"   '\Reportes
    If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Consulta de Pólizas Traspasadas no se encuentra en el Directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
    End If
 
    vgQuery = "{PD_TTMP_CIATASACOT.COD_USUARIO} = '" & vgUsuario & "' "
    If (Opt_SinPrimaRec.Value = True) Then
        vgQuery = vgQuery & " AND {PD_TTMP_CIATASACOT.GLS_ESTADO} = '" & clEstadoSinTraspasoPrima & "'"
    End If
    If (Opt_ConPrimaRec.Value = True) Then
        vgQuery = vgQuery & " AND {PD_TTMP_CIATASACOT.GLS_ESTADO} = '" & clEstadoConTraspasoPrima & "'"
    End If

    Rpt_CIA.Reset
    Rpt_CIA.ReportFileName = vlArchivo     'App.Path & "\rpt_Areas.rpt"
    Rpt_CIA.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
    Rpt_CIA.SelectionFormula = vgQuery
    Rpt_CIA.Formulas(0) = ""
    Rpt_CIA.Formulas(1) = ""
    Rpt_CIA.Formulas(2) = ""
    Rpt_CIA.Formulas(3) = ""
    Rpt_CIA.Formulas(4) = ""
    Rpt_CIA.Formulas(0) = "FechaDesde = '" & vlDesde & "'"
    Rpt_CIA.Formulas(1) = "FechaHasta = '" & vlHasta & "'"
    Rpt_CIA.Formulas(2) = "NombreCompania = '" & vgNombreCompania & "'"
    Rpt_CIA.Formulas(3) = "NombreSistema= '" & vgNombreSistema & "'"
    Rpt_CIA.Formulas(4) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
    Rpt_CIA.WindowState = crptMaximized
    Rpt_CIA.Destination = crptToWindow
    Rpt_CIA.WindowTitle = "Informe Interno de las Tasas de Cotización"
    Rpt_CIA.Action = 1
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
    Txt_Desde.SetFocus
    
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
    
    Call flIniciaGrilla
    
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

Private Sub Opt_ConPrimaRec_Click()
If (Opt_ConPrimaRec.Value = True) Then
    Call flCargarGrilla
End If
End Sub

Private Sub Opt_SinPrimaRec_Click()
If (Opt_SinPrimaRec.Value = True) Then
    Call flCargarGrilla
End If
End Sub

Private Sub Opt_Todos_Click()
If (Opt_Todos.Value = True) Then
    Call flCargarGrilla
End If
End Sub

Private Sub Txt_Desde_GotFocus()
    Txt_Desde.SelStart = 0
    Txt_Desde.SelLength = Len(Txt_Desde)
End Sub

Private Sub Txt_Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And (Trim(Txt_Desde) <> "") Then
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
    If KeyAscii = 13 And (Trim(Txt_Hasta) <> "") Then
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
'    flCargarGrilla
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
