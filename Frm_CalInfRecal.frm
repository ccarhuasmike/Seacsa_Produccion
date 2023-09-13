VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_CalInfRecal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Casos Recalculados."
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9165
   Begin VB.Frame Fra_Fechas 
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
      TabIndex        =   7
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   6840
         Picture         =   "Frm_CalInfRecal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Efectuar Busqueda de Tasas"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Recálculo realizado a Pólizas"
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
         Left            =   2760
         TabIndex        =   10
         Top             =   0
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
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rango de Fechas de Emisión  :"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   8895
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   2880
         Picture         =   "Frm_CalInfRecal.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4080
         Picture         =   "Frm_CalInfRecal.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5280
         Picture         =   "Frm_CalInfRecal.frx":0E76
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   4095
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      BackColor       =   14745599
      FormatString    =   "Nº Póliza     | CUSPP                     | Fecha Incorp. | Tipo Pensión | Usuario          | Fecha        | Hora         "
   End
   Begin Crystal.CrystalReport Rpt_General 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Frm_CalInfRecal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function flInicializaGrilla()

    Msf_Grilla.Clear
    Msf_Grilla.Cols = 7
    Msf_Grilla.Rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0
        
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Nº Póliza"
    Msf_Grilla.ColWidth(0) = 1000
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "CUSPP"
    Msf_Grilla.ColWidth(1) = 1500
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Fec.Emisión"
    Msf_Grilla.ColWidth(2) = 1000
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Tipo Pensión"
    Msf_Grilla.ColWidth(3) = 2200
    
    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Usuario"
    Msf_Grilla.ColWidth(4) = 1200
    
    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "Fecha"
    Msf_Grilla.ColWidth(5) = 1000

    Msf_Grilla.Col = 6
    Msf_Grilla.Text = "Hora"
    Msf_Grilla.ColWidth(6) = 800
    
End Function

Function flCargaGrilla()
On Error GoTo Err_Carga
Dim vlFechaEmision As String, vlTipoPension As String
Dim vlFechaCreacion As String, vlHoraCreacion As String

    Call flInicializaGrilla
    While Not vgRs.EOF
    
        vlFechaEmision = DateSerial(Mid((vgRs!Fec_Emision), 1, 4), Mid(vgRs!Fec_Emision, 5, 2), Mid(vgRs!Fec_Emision, 7, 2))
        vlTipoPension = fgBuscarGlosaElemento(vgCodTabla_TipPen, (vgRs!Cod_TipPension))
        vlTipoPension = " " + Trim(vgRs!Cod_TipPension) + " - " + Trim(vlTipoPension)
        vlFechaCreacion = DateSerial(Mid((vgRs!Fec_Crea), 1, 4), Mid(vgRs!Fec_Crea, 5, 2), Mid(vgRs!Fec_Crea, 7, 2))
        vlHoraCreacion = TimeSerial(Mid((vgRs!Hor_Crea), 1, 2), Mid(vgRs!Hor_Crea, 3, 2), Mid(vgRs!Hor_Crea, 5, 2))
        
        Msf_Grilla.AddItem Trim(vgRs!Num_Poliza) & vbTab _
                & " " & (vgRs!Cod_Cuspp) & vbTab _
                & Trim(vlFechaEmision) & vbTab _
                & (vlTipoPension) & vbTab _
                & (vgRs!Cod_UsuarioCrea) & vbTab _
                & Trim(vlFechaCreacion) & vbTab _
                & Trim(vlHoraCreacion)
                        
        vgRs.MoveNext
    Wend

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub Cmd_Buscar_Click()
On Error GoTo Err_Buscar

    If (Trim(Txt_Desde) = "") Then
       MsgBox "Debe Ingresar una Fecha para el Valor Desde", vbCritical, "Error de Datos"
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_Desde.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_Desde) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If (Year(Txt_Desde) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_Desde.SetFocus
       Exit Sub
    End If
    
    Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")

    If (Trim(Txt_Hasta) = "") Then
       MsgBox "Debe Ingresar una Fecha para el Valor Desde", vbCritical, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_Hasta.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    If (Year(Txt_Hasta) < 1900) Then
       MsgBox "La Fecha ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    
    Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    
    If Trim(Txt_Desde.Text) > Trim(Txt_Hasta.Text) Then
       MsgBox "Fecha Hasta, Debe ser Mayor o Igual a Fecha Desde", vbCritical, "Error de Datos"
       Txt_Hasta.Text = ""
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    
    Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
    Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
   
    vlFechaDesde = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
    vlFechaHasta = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")
        
    Sql = " SELECT num_poliza,fec_emision,fec_vigencia,"
    Sql = Sql & " cod_cuspp,cod_tippension,mto_pension, "
    Sql = Sql & " cod_usuariocrea,fec_crea,hor_crea "
    Sql = Sql & " FROM pd_tmae_oripoliza "
    Sql = Sql & " WHERE (fec_emision >= '" & Trim(vlFechaDesde) & "') AND "
    Sql = Sql & " (fec_emision <= '" & Trim(vlFechaHasta) & "') AND "
    Sql = Sql & " ind_recalculo = 'S' "
    Sql = Sql & " UNION "
    Sql = Sql & " SELECT num_poliza,fec_emision,fec_vigencia,"
    Sql = Sql & " cod_cuspp,cod_tippension,mto_pension, "
    Sql = Sql & " cod_usuariocrea,fec_crea,hor_crea "
    Sql = Sql & " FROM pd_tmae_poliza "
    Sql = Sql & " WHERE (fec_emision >= '" & Trim(vlFechaDesde) & "') AND "
    Sql = Sql & " (fec_emision <= '" & Trim(vlFechaHasta) & "') AND "
    Sql = Sql & " ind_recalculo = 'S' "
    Sql = Sql & " ORDER BY num_poliza "
    Set vgRs = vgConexionBD.Execute(Sql)
    If Not vgRs.EOF Then
       Call flCargaGrilla
       Fra_Fechas.Enabled = False
    Else
        Screen.MousePointer = 11
        MsgBox "No Existen Polizas Recalculadas en el Período Ingresado", vbInformation, "Información"
        Screen.MousePointer = 0
        Exit Sub
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
On Error GoTo Err_CmdImprimir

    If (Trim(Txt_Desde) = "") Then
       MsgBox "Debe Ingresar una Fecha para el Valor Desde", vbCritical, "Error de Datos"
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_Desde.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If (CDate(Txt_Desde) > CDate(Date)) Then
       MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
       Txt_Desde.SetFocus
       Exit Sub
    End If
    If (Year(Txt_Desde) < 1900) Then
       MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_Desde.SetFocus
       Exit Sub
    End If
    
    Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")

    If (Trim(Txt_Hasta) = "") Then
       MsgBox "Debe Ingresar una Fecha para el Valor Desde", vbCritical, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    If Not IsDate(Txt_Hasta.Text) Then
       MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    If (Year(Txt_Hasta) < 1900) Then
       MsgBox "La Fecha ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    
    Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    
    If Trim(Txt_Desde.Text) > Trim(Txt_Hasta.Text) Then
       MsgBox "Fecha Hasta, Debe ser Mayor o Igual a Fecha Desde", vbCritical, "Error de Datos"
       Txt_Hasta.Text = ""
       Txt_Hasta.SetFocus
       Exit Sub
    End If
    
    Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
    Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
   
    Screen.MousePointer = 11
    
    vlArchivo = strRpt & "PD_Rpt_CalRecalculo.rpt"   '\Reportes
    If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Consulta de Pólizas Traspasadas no se encuentra en el Directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
    End If
 
    vlFechaDesde = Format(CDate(Trim(Txt_Desde.Text)), "yyyymmdd")
    vlFechaHasta = Format(CDate(Trim(Txt_Hasta.Text)), "yyyymmdd")

    vgQuery = ""
    vgQuery = vgQuery & "{pd_TMAE_POLIZA.fec_emision} >= '" & Trim(vlFechaDesde) & "' AND "
    vgQuery = vgQuery & "{pd_TMAE_POLIZA.fec_emision} <= '" & Trim(vlFechaHasta) & "' AND "
    vgQuery = vgQuery & "{pd_TMAE_POLIZA.ind_recalculo} = 'S' "

    Rpt_General.Reset
    Rpt_General.ReportFileName = vlArchivo     'App.Path & "\rpt_Areas.rpt"
    Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
    Rpt_General.SelectionFormula = vgQuery
   
    Rpt_General.Formulas(0) = ""
    Rpt_General.Formulas(1) = ""
    Rpt_General.Formulas(2) = ""
    Rpt_General.Formulas(3) = ""
    Rpt_General.Formulas(4) = ""
 
    vgPalabra = Txt_Desde.Text & "  *  " & Txt_Hasta.Text
    Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
    Rpt_General.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
    Rpt_General.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
    Rpt_General.Formulas(3) = "Periodo = '" & vgPalabra & "'"

    Rpt_General.SubreportToChange = "PD_Rpt_CalRecalculo.rpt"     'App.Path & "\rpt_Areas.rpt"
    Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
    Rpt_General.SelectionFormula = vgQuery

    vgQuery = "{pd_TMAE_ORIPOLIZA.fec_emision} >= '" & Trim(vlFechaDesde) & "' AND "
    vgQuery = vgQuery & "{pd_TMAE_ORIPOLIZA.fec_emision} <= '" & Trim(vlFechaHasta) & "' AND "
    vgQuery = vgQuery & "{pd_TMAE_ORIPOLIZA.ind_recalculo} = 'S' "
     
    Rpt_General.SubreportToChange = "PD_Rpt_CalRecalculo_Ori.rpt"     'App.Path & "\rpt_Areas.rpt"
    Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
    Rpt_General.SelectionFormula = vgQuery
    
    Rpt_General.WindowState = crptMaximized
    Rpt_General.Destination = crptToWindow
    Rpt_General.WindowTitle = "Informe de Consulta de Pólizas Recalculadas"
    Rpt_General.Action = 1
    
    Rpt_General.SubreportToChange = ""
    
    Screen.MousePointer = 0
   
Exit Sub
Err_CmdImprimir:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Limpiar_Click()
On Error GoTo Err_CmdLimpiar

    Fra_Fechas.Enabled = True
    Txt_Desde.Text = ""
    Txt_Hasta.Text = ""
    Txt_Desde.SetFocus

    Call flInicializaGrilla

Exit Sub
Err_CmdLimpiar:
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
    
    Call flInicializaGrilla

Exit Sub
Err_Cargar:
Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If (Trim(Txt_Desde) = "") Then
         MsgBox "Debe Ingresar una Fecha para el Valor Desde", vbCritical, "Error de Datos"
         Txt_Desde.SetFocus
         Exit Sub
      End If
      If Not IsDate(Txt_Desde.Text) Then
         MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
         Txt_Desde.SetFocus
         Exit Sub
      End If
      If (CDate(Txt_Desde) > CDate(Date)) Then
         MsgBox "La Fecha Ingresada es Mayor a la Fecha Actual", vbCritical, "Error de Datos"
         Txt_Desde.SetFocus
         Exit Sub
      End If
      If (Year(Txt_Desde) < 1900) Then
         MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
         Txt_Desde.SetFocus
         Exit Sub
      End If
      Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
      Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))
      Txt_Hasta.SetFocus
    End If
End Sub

Private Sub Txt_Desde_LostFocus()

    If (Trim(Txt_Desde) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_Desde.Text) Then
       Exit Sub
    End If
    If (CDate(Txt_Desde) > CDate(Date)) Then
       Exit Sub
    End If
    If (Year(Txt_Desde) < 1900) Then
       Exit Sub
    End If
    Txt_Desde.Text = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    Txt_Desde.Text = DateSerial(Mid((Txt_Desde.Text), 1, 4), Mid((Txt_Desde.Text), 5, 2), Mid((Txt_Desde.Text), 7, 2))

End Sub

Private Sub Txt_Hasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
      If (Trim(Txt_Hasta) = "") Then
         MsgBox "Debe Ingresar una Fecha para el Valor Hasta", vbCritical, "Error de Datos"
         Txt_Hasta.SetFocus
         Exit Sub
      End If
      If Not IsDate(Txt_Hasta.Text) Then
         MsgBox "La Fecha Ingresada No es una Fecha Válida.", vbCritical, "Error de Datos"
         Txt_Hasta.SetFocus
         Exit Sub
      End If
      If (Year(Txt_Hasta) < 1900) Then
         MsgBox "La Fecha Ingresada es Inferior a la Fecha Mínima de Ingreso (1900).", vbCritical, "Error de Datos"
         Txt_Hasta.SetFocus
         Exit Sub
      End If
      
      Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
      
      If Trim(Format(CDate(Trim(Txt_Desde)), "yyyymmdd")) > Trim(Txt_Hasta.Text) Then
         MsgBox "Fecha Hasta, Debe ser Mayor o Igual a Fecha Desde", vbCritical, "Error de Datos"
         Txt_Hasta.Text = ""
         Txt_Hasta.SetFocus
         Exit Sub
      End If
      Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
      Cmd_Buscar.SetFocus
    End If
    
End Sub

Private Sub Txt_Hasta_LostFocus()

    If (Trim(Txt_Hasta) = "") Then
       Exit Sub
    End If
    If Not IsDate(Txt_Hasta.Text) Then
       Exit Sub
    End If
    If (CDate(Txt_Hasta) > CDate(Date)) Then
       Exit Sub
    End If
    If (Year(Txt_Hasta) < 1900) Then
       Exit Sub
    End If

    Txt_Hasta.Text = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    
    If Trim(Format(CDate(Trim(Txt_Desde)), "yyyymmdd")) > Trim(Txt_Hasta.Text) Then
       Exit Sub
    End If
    
    Txt_Hasta.Text = DateSerial(Mid((Txt_Hasta.Text), 1, 4), Mid((Txt_Hasta.Text), 5, 2), Mid((Txt_Hasta.Text), 7, 2))
    
End Sub
