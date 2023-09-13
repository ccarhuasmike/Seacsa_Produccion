VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_CalPrimaArc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Archivo de Confirmación de Primas."
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9150
   Begin VB.Frame Frame1 
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
      TabIndex        =   26
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton Cmd_BuscarDir 
         Height          =   375
         Left            =   7800
         Picture         =   "Frm_CalPrimaArc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Buscar Directorio del Archivo"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Archivo 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   525
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   7455
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Ubicación del Archivo"
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
         Index           =   9
         Left            =   240
         TabIndex        =   27
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Frame Fra_Resumen 
      Caption         =   "  Resumen de Archivo de Confirmación "
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
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   8895
      Begin VB.Label Lbl_HorCrea 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Lbl_FecCrea 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Lbl_Usuario 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Lbl_MtoTotPri 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Lbl_TotPol 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Hora de Creación"
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   25
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha de Creación"
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   24
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Usuario Creación"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   23
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Número Total de Pólizas"
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   22
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Monto Total Primas"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   21
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   8895
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Resumen"
         Height          =   675
         Index           =   1
         Left            =   3120
         Picture         =   "Frm_CalPrimaArc.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton Cmd_BuscarArch 
         Caption         =   "Archivo"
         Height          =   675
         Left            =   5160
         Picture         =   "Frm_CalPrimaArc.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Buscar Archivo Generado"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Cargar 
         Caption         =   "&Cargar"
         Height          =   675
         Left            =   2160
         Picture         =   "Frm_CalPrimaArc.frx":08BE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Exportar a Archivo"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6120
         Picture         =   "Frm_CalPrimaArc.frx":10E0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4200
         Picture         =   "Frm_CalPrimaArc.frx":11DA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   8280
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8895
      Begin VB.CommandButton Cmd_BuscarRan 
         Height          =   375
         Left            =   6360
         Picture         =   "Frm_CalPrimaArc.frx":1894
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Efectuar Consulta"
         Top             =   310
         Width           =   855
      End
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog ComDialogo 
         Left            =   8280
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rango de Fechas   :"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Recepción de Primas"
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
         TabIndex        =   17
         Top             =   0
         Width           =   3015
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
         Index           =   4
         Left            =   4200
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "Frm_CalPrimaArc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlArchivo As String, vlArchOri As String, linea As String, vlLargoArchivo As Long
Dim Query As String
Dim vlFechaDesde As String, vlFechaHasta As String
Dim vlContPolizas As Long, vlSumPrima As Double
Dim vlNumArch As Long

'Constantes utilizadas para etiquetas xml
Const clXmlCargaConf = "cargaConfirmaciones"
Const clXmlConf = "confirmacion"
Const clXmlOperacion = "operacion"
Const clXmlCuspp = "CUSPP"
Const clXmlPensionES = "pensionEESS"
Const clXmlNumPol = "numeroPoliza"
Const clXmlPriPenRV = "primeraPensionRV"
Const clXmlPriPenRT = "primeraPensionRT"
Const clXmlPriPenRVD = "primeraPensionRVD"
'I--- ABV 09/12/2009 ---
'Const clXmlMtoTransf = "montoTransferido"
Const clXmlMtoTransf = "primaUnicaEESS"
'I--- ABV 09/12/2009 ---
Const clXmlIniVig = "inicioVigencia"
'I--- ABV 09/12/2009 ---
Const clXmlMtoTransfAFP = "primaUnicaAFPEESS"
'F--- ABV 09/12/2009 ---

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

Private Sub Cmd_BuscarArch_Click()
On Error GoTo Err_CmdBuscarClick

    Frm_BuscaArchivo.flInicio ("Frm_CalPrimaArc")
    
Exit Sub
Err_CmdBuscarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarRan_Click()

    If Trim(Txt_Desde) = "" Or Trim(Txt_Hasta) = "" Then
        MsgBox "Debe ingresar el Rango de Fechas a Buscar", vbInformation, "Falta Información"
        Txt_Desde.SetFocus
        Exit Sub
    Else
       If flValidaFecha(Txt_Desde) <> True Then
          MsgBox "Debe ingresar Rango de Fecha Válido", vbInformation, "Falta Información"
          Txt_Desde.SetFocus
          Exit Sub
       End If
       If flValidaFecha(Txt_Hasta) <> True Then
          MsgBox "Debe ingresar Rango de Fecha Válido", vbInformation, "Falta Información"
          Txt_Hasta.SetFocus
          Exit Sub
       End If
       vlFechaDesde = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
       vlFechaHasta = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    End If

    Call flBuscaRango
    
    Cmd_Cargar.SetFocus
    
End Sub

Private Sub Cmd_BuscarDir_Click()
Dim ilargo As Long
On Error GoTo Err_Cmd

    vlArchivo = ""
    ComDialogo.CancelError = True
    ComDialogo.FileName = "ArcConfirmacion" & Format(Date, "yyyymmdd") & ".xml"
    ComDialogo.DialogTitle = "Archivo de Confirmación de Primas"
    ComDialogo.Filter = "*.xml"
    ComDialogo.FilterIndex = 1
    ComDialogo.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    ComDialogo.ShowOpen
    
    vlArchivo = ComDialogo.FileName
    Lbl_Archivo.Caption = vlArchivo
    If (Len(vlArchivo) > 100) Then
        While Len(Lbl_Archivo) > 100
            ilargo = InStr(1, Lbl_Archivo, "\")
            Lbl_Archivo = Mid(Lbl_Archivo, ilargo + 1, Len(Lbl_Archivo))
        Wend
        Lbl_Archivo.Caption = "\\" & Lbl_Archivo
    End If
   
    Txt_Desde.SetFocus
Exit Sub
Err_Cmd:
    If Err.Number = 32755 Then
       Exit Sub
    End If
    Screen.MousePointer = 0
    MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
End Sub

Private Sub Cmd_Cargar_Click()

    Screen.MousePointer = 11
    
    'Validar Directorio
    If Lbl_Archivo = "" Then
        MsgBox "Debe seleccionar Archivo a Generar.", vbCritical, "Operación Cancelada"
        Screen.MousePointer = 0
        Cmd_BuscarDir.SetFocus
        Exit Sub
    End If
    'validacion fechas
    If Not IsDate(Txt_Desde) Then
        MsgBox "La Fecha Desde ingresada no es válida.", vbCritical, "Operación Cancelada"
        Screen.MousePointer = 0
        Txt_Desde.SetFocus
        Exit Sub
    End If
    If Not IsDate(Txt_Hasta) Then
        MsgBox "La Fecha Hasta ingresada no es válida.", vbCritical, "Operación Cancelada"
        Screen.MousePointer = 0
        Txt_Hasta.SetFocus
        Exit Sub
    End If

    vgRes = MsgBox("¿ Está seguro que desea Generar el Archivo de Confirmación?", 4 + 32 + 256, "Proceso de Generación")
    If vgRes <> 6 Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    vlFechaDesde = Format(CDate(Trim(Txt_Desde)), "yyyymmdd")
    vlFechaHasta = Format(CDate(Trim(Txt_Hasta)), "yyyymmdd")
    
    'Obtiene el nombre del archivo sin extensión
    Dim pos As Integer
    pos = InStrRev(vlArchivo, ".xml")
    If (pos = 0) Then
        vlArchOri = vlArchivo
    Else
        vlArchOri = Mid(vlArchivo, 1, pos - 1)
    End If
        
    If (vlArchivo <> "") Then
        'Elimina el archivo .xml si existe
        If fgExiste(vlArchOri & ".xml") Then Kill vlArchOri & ".xml"
        'Elimina el archivo .txt si existe
        If fgExiste(vlArchOri & ".txt") Then Kill vlArchOri & ".txt"
    End If
 
    vlArchivo = vlArchOri & ".txt"
    
    'Genera archivo xml
    If (flGeneraXml(vlArchivo, vlFechaDesde, vlFechaHasta) = False) Then
        MsgBox "Se ha Producido un error en la Generación del Archivo", vbCritical, "Proceso Cancelado"
        Exit Sub
    End If
    
    pos = InStrRev(vlArchivo, "\") + 1
    
    vlArchivo = Mid(vlArchivo, pos, Len(vlArchivo))
    
    'Guarda estadistica
    Call flEstadistica
    
    Lbl_TotPol.Caption = Format(vlContPolizas, "#,#0")
    Lbl_MtoTotPri.Caption = Format(vlSumPrima, "#,#0.00")
    
    MsgBox "El Proceso de Generación de Archivo ha Finalizado Exitosamente", vbInformation, "Proceso Finalizado"
    
    Screen.MousePointer = 0
    
End Sub

Private Sub Cmd_Imprimir_Click(Index As Integer)
Dim vlArchivo As String
Err.Clear
On Error GoTo Errores1
   
    'Validar el Ingreso del Rango de Fechas
    If Txt_Desde = "" Then
        MsgBox "Debe ingresar el Rango de Inicio de la Consulta de Primas.", vbCritical, "Error de Datos"
        Txt_Desde.SetFocus
        Exit Sub
    Else
        vlFechaDesde = Trim(Txt_Desde)
        If (flValidaFecha(vlFechaDesde) = False) Then
            Txt_Desde = ""
            Txt_Desde.SetFocus
            Exit Sub
        End If
        vlFechaDesde = Format(vlFechaDesde, "yyyyMMdd")
    End If
    
    If Txt_Hasta = "" Then
       MsgBox "Debe ingresar el Rango de Inicio de la Consulta de Primas.", vbCritical, "Error de Datos"
       Txt_Hasta.SetFocus
       Exit Sub
    Else
        vlFechaHasta = Trim(Txt_Hasta)
        If (flValidaFecha(vlFechaHasta) = False) Then
            Txt_Hasta = ""
            Txt_Hasta.SetFocus
            Exit Sub
        End If
        vlFechaHasta = Format(vlFechaHasta, "yyyyMMdd")
    End If
    
    If (Lbl_Archivo.Caption = "") Then
       MsgBox "Debe ingresar Archivo de Confirmación.", vbCritical, "Error de Datos"
       Cmd_BuscarDir.SetFocus
       Exit Sub
    End If
    
    Screen.MousePointer = 11

    vgSql = "SELECT 1 from pd_tmae_estcarconpri "
    vgSql = vgSql & " WHERE "
    vgSql = vgSql & " fec_desde= '" & vlFechaDesde & "' "
    vgSql = vgSql & " and fec_hasta= '" & vlFechaHasta & "' "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If vlRegistro.EOF Then
       vlRegistro.Close
       MsgBox "No existen el Archivo Seleccionado en la BD.", vbInformation, "Inexistencia de Datos"
       Screen.MousePointer = 0
       Exit Sub
    End If
    vlRegistro.Close
     
     vlArchivo = strRpt & "pd_rpt_EstConPri.rpt"   '\Reportes
     If Not fgExiste(vlArchivo) Then     ', vbNormal
         MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
         Screen.MousePointer = 0
         Exit Sub
     End If
     
     vgQuery = "{PD_TMAE_ESTCARCONPRI.FEC_DESDE} = '" & vlFechaDesde & "' AND "
     vgQuery = vgQuery & "{PD_TMAE_ESTCARCONPRI.FEC_HASTA} = '" & vlFechaHasta & "' AND "
     vgQuery = vgQuery & "{PD_TMAE_ESTCARCONPRI.NUM_ARCHIVO} = " & vlNumArch & ""
 
    Rpt_Reporte.Reset
    Rpt_Reporte.ReportFileName = vlArchivo
    Rpt_Reporte.Connect = vgRutaDataBase
    Rpt_Reporte.SelectionFormula = vgQuery
    Rpt_Reporte.Formulas(0) = ""
    Rpt_Reporte.Formulas(1) = ""
    Rpt_Reporte.Formulas(2) = ""
   
    Rpt_Reporte.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
    Rpt_Reporte.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
    Rpt_Reporte.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"

    Rpt_Reporte.Destination = crptToWindow
    Rpt_Reporte.WindowState = crptMaximized
    Rpt_Reporte.WindowTitle = "Informe de Archivo de Confirmación de Primas"
    Rpt_Reporte.Action = 1
   
    Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Cmd_Limpiar_Click()
    Txt_Desde = ""
    Txt_Hasta = ""
    Lbl_Archivo = ""
    Lbl_TotPol.Caption = ""
    Lbl_MtoTotPri.Caption = ""
    Lbl_Usuario.Caption = ""
    Lbl_FecCrea.Caption = ""
    Lbl_HorCrea.Caption = ""
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
          Cmd_BuscarRan.SetFocus
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
End Sub

Private Sub Cmd_Salir_Click()
On Error GoTo Err_CmdSalir
    
    Screen.MousePointer = 11
    Unload Me
    Screen.MousePointer = 0
    
Exit Sub
Err_CmdSalir:
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
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Function flGeneraXml(strNombreArch As String, iFecDesde As String, iFecHasta As String) As Boolean
On Error GoTo Err_CrearXml
Dim vlNumOpe, vlcuspp, vlNumPol, vlIniVig As String
Dim vlMtoPension As String, vlMtoPriRec As String
Dim i As Integer, vlMesDif As Integer
Dim vlMtoSumPension, vlMtoSumPensionAfp, vlMtoPensionRT As String
Dim vlTipPension As String
Dim vlMtoRtaTmpAfp As String
Dim vlMtoValMoneda As Double
Dim linea As String
Dim fn
'I--- ABV 09/12/2009 ---
Dim vlMtoPriRecAFP As String
'F--- ABV 09/12/2009 ---

    flGeneraXml = False

    vlContPolizas = 0
    vlSumPrima = 0
    
    fn = FreeFile
    
    Open strNombreArch For Output As fn
    'se abre el archivo txt
        
    'tag de descripción del xml
    Print #fn, "<?xml version='1.0' encoding='ISO-8859-1'?> "
    Print #fn, "<!-- edited with XMLSpy v2006 rel. 3 sp1 (http://www.altova.com) by newbe (EMBRACE) -->"
    '*Print #fn, "<descargaSolicitudesEESS xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance' xsi:noNamespaceSchemaLocation='http://extranet.sbs.gob.pe/xsd/descargaSolicitudesEESS12.xsd'>"
    'Segun CMelo van doble comilla
    Print #fn, "<" & clXmlCargaConf & ">"
    
    Query = ""
    Query = "SELECT P.NUM_OPERACION,P.COD_CUSPP,P.NUM_POLIZA,R.MTO_PENSION,"
    Query = Query & "R.MTO_PRIREC,P.FEC_INIPENCIA,P.NUM_MESDIF,"
    Query = Query & "R.MTO_SUMPENSION,R.MTO_SUMPENSIONAFP," '09/11/2007
'    Query = Query & "COD_TIPPENSION,MTO_RENTATMPAFP,MTO_VALMONEDA " '07/11/2007
    Query = Query & "COD_TIPPENSION,R.MTO_PENSIONAFP as MTO_RENTATMPAFP,MTO_VALMONEDA " '10/11/2007
'I--- ABV 09/12/2009 ---
    Query = Query & ",R.MTO_PRIAFP as MTO_PRIMAAFP "
'F--- ABV 09/12/2009 ---
    Query = Query & "FROM PD_TMAE_POLIZA P, PD_TMAE_POLPRIREC R "
    Query = Query & "WHERE P.NUM_POLIZA=R.NUM_POLIZA AND "
    Query = Query & "R.FEC_TRASPASO BETWEEN '" & iFecDesde & "' AND '" & iFecHasta & "' "
    Query = Query & "ORDER BY P.NUM_OPERACION "
    Set vgRs = vgConexionBD.Execute(Query)
    If Not (vgRs.EOF) Then
        While Not vgRs.EOF
            
            vlNumOpe = Trim(vgRs!Num_Operacion)
            vlcuspp = Trim(vgRs!Cod_Cuspp)
            vlNumPol = Trim(vgRs!Num_Poliza)
            vlMtoPension = Replace(Format(vgRs!Mto_Pension, "#0.00"), ",", ".")
            vlMtoPriRec = Replace(Format(vgRs!MTO_PRIREC, "#0.00"), ",", ".")
            vlIniVig = Mid(vgRs!fec_inipencia, 1, 4) & "-" & Mid(vgRs!fec_inipencia, 5, 2) & "-" & Mid(vgRs!fec_inipencia, 7, 2)
            vlMesDif = CInt(vgRs!Num_MesDif)
            vlMtoSumPension = Replace(Format(vgRs!Mto_SumPension, "#0.00"), ",", ".")
            vlMtoSumPensionAfp = Format(vgRs!Mto_SumPensionAfp, "#0.00")
            vlTipPension = Trim(vgRs!Cod_TipPension)
            vlMtoValMoneda = Format(vgRs!Mto_ValMoneda, "#0.00")
            vlMtoRtaTmpAfp = Format(vgRs!Mto_RentaTMPAFP, "#0.00")
            
'I--- ABV 09/12/2009 ---
            If Not IsNull(vgRs!Mto_PrimaAFP) Then
                vlMtoPriRecAFP = Replace(Format(vgRs!Mto_PrimaAFP, "#0.00"), ",", ".")
            Else
                vlMtoPriRecAFP = Replace(Format(0, "#0.00"), ",", ".")
            End If
'F--- ABV 09/12/2009 ---
            
            If (vlTipPension = "08") Then
                vlMtoPensionRT = Replace(Format((vlMtoSumPensionAfp), "#0.00"), ",", ".")
            Else
                vlMtoPensionRT = Replace(Format((vlMtoRtaTmpAfp), "#0.00"), ",", ".")
            End If

            'Escribe en el archivo
            Print #fn, "<" & clXmlConf & ">"
            
            Print #fn, "<" & clXmlOperacion & ">" & vlNumOpe & "</" & clXmlOperacion & ">"
            
            Print #fn, "<" & clXmlCuspp & ">" & vlcuspp & "</" & clXmlCuspp & ">"
            
            Print #fn, "<" & clXmlPensionES & ">"
            
            Print #fn, "<" & clXmlNumPol & ">" & vlNumPol & "</" & clXmlNumPol & ">"
            
            If (vlTipPension = "08") Then
                If (vlMesDif = 0) Then
                    Print #fn, "<" & clXmlPriPenRV & ">" & vlMtoSumPension & "</" & clXmlPriPenRV & ">"
                Else
                    Print #fn, "<" & clXmlPriPenRT & ">" & vlMtoPensionRT & "</" & clXmlPriPenRT & ">"
                    Print #fn, "<" & clXmlPriPenRVD & ">" & vlMtoSumPension & "</" & clXmlPriPenRVD & ">"
                End If
            Else
                If (vlMesDif = 0) Then
                    Print #fn, "<" & clXmlPriPenRV & ">" & vlMtoPension & "</" & clXmlPriPenRV & ">"
                Else
                    Print #fn, "<" & clXmlPriPenRT & ">" & vlMtoPensionRT & "</" & clXmlPriPenRT & ">"
                    Print #fn, "<" & clXmlPriPenRVD & ">" & vlMtoPension & "</" & clXmlPriPenRVD & ">"
                End If
            End If
            
'I--- ABV 09/12/2009 ---
            If (vlMesDif > 0) Then
                Print #fn, "<" & clXmlMtoTransfAFP & ">" & vlMtoPriRecAFP & "</" & clXmlMtoTransfAFP & ">"
            End If
'F--- ABV 09/12/2009 ---
            
            Print #fn, "<" & clXmlMtoTransf & ">" & vlMtoPriRec & "</" & clXmlMtoTransf & ">"
            
            Print #fn, "<" & clXmlIniVig & ">" & vlIniVig & "</" & clXmlIniVig & ">"
            
            Print #fn, "</" & clXmlPensionES & ">"

            Print #fn, "</" & clXmlConf & ">"
            
            vlContPolizas = vlContPolizas + 1
            vlSumPrima = vlSumPrima + Val(vlMtoPriRec) 'hqr 19/02/2011 Por problema en decimales
            
            vgRs.MoveNext
        Wend
    End If
    vgRs.Close
    
    Print #fn, "</" & clXmlCargaConf & ">"
        
    'cerramos el archivo txt
    Close #fn
    
    'Renombra el archivo txt a xml
    Name (vlArchOri & ".txt") As (vlArchOri & ".xml")

    'Copia el archivo txt como xml y luego borra el txt
    ''FileCopy (strNombreArch & ".txt"), (strNombreArch & ".xml")
    ''Kill (strNombreArch & ".txt")

    flGeneraXml = True
       
Exit Function
Err_CrearXml:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Private Sub flBuscaRango()
On Error GoTo Err_BuscaRan

    Query = ""
    Query = Query & "SELECT COUNT(P.NUM_POLIZA) AS NUMPOL ,SUM(R.MTO_PRIREC) AS MTOPRIMA "
    Query = Query & "FROM PD_TMAE_POLIZA P, PD_TMAE_POLPRIREC R "
    Query = Query & "WHERE P.NUM_POLIZA=R.NUM_POLIZA AND "
    Query = Query & "R.FEC_TRASPASO BETWEEN '" & vlFechaDesde & "' AND '" & vlFechaHasta & "' "
    Set vgRs = vgConexionBD.Execute(Query)
    If Not (vgRs.EOF) Then
        vlContPolizas = IIf(IsNull(vgRs!numpol), 0, vgRs!numpol)
        vlSumPrima = IIf(IsNull(vgRs!MTOPRIMA), 0, vgRs!MTOPRIMA)
    End If
    vgRs.Close
    
    Lbl_TotPol.Caption = Format(vlContPolizas, "#,#0")
    Lbl_MtoTotPri.Caption = Format(vlSumPrima, "#,#0.00")
    Lbl_Usuario.Caption = vgUsuario
    Lbl_FecCrea.Caption = Date
    Lbl_HorCrea.Caption = Time
    
Exit Sub
Err_BuscaRan:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

 '----- RESCATA EL NUMERO DE ENTRADA CORRESPONDIENTE AL ARCHIVO QUE SE ESTA CARGANDO -----
Private Function flNumArchivo() As Integer
    Query = ""
    Query = "SELECT NUM_ARCHIVO FROM PD_TMAE_ESTCARCONPRI "
    Query = Query & " ORDER BY NUM_ARCHIVO DESC"
    Set vgRs = vgConexionBD.Execute(Query)
    If Not (vgRs.EOF) Then
        flNumArchivo = CInt(vgRs!Num_Archivo) + 1
    Else
        flNumArchivo = 1
    End If
End Function

Private Sub flEstadistica()
'Guarda en la información de la carga realizada
On Error GoTo Err_Estadistica

    'Obtiene el nº de archivo a crear
    vlNumArch = flNumArchivo
    
    Query = ""
    Query = "INSERT INTO PD_TMAE_ESTCARCONPRI (NUM_ARCHIVO,GLS_NOMARCH,"
    Query = Query & "FEC_DESDE,FEC_HASTA,NUM_POLIZAS,MTO_PRIMAS,"
    Query = Query & "COD_USUARIOCREA,FEC_CREA,HOR_CREA "
    Query = Query & ") VALUES ("
    Query = Query & " " & vlNumArch & ","
    Query = Query & "'" & vlArchivo & "',"
    Query = Query & "'" & vlFechaDesde & "',"
    Query = Query & "'" & vlFechaHasta & "',"
    Query = Query & " " & vlContPolizas & ","
    Query = Query & " " & Str(vlSumPrima) & ","
    Query = Query & "'" & vgUsuario & "',"
    Query = Query & "'" & (Format(Now(), "yyyyMMdd")) & "',"
    Query = Query & "'" & (Format(Now(), "hhmmss")) & "')"
    vgConexionBD.Execute (Query)

Exit Sub
Err_Estadistica:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

'-------------------------------------------------------
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA
'-------------------------------------------------------
Function flRecibeArchivo(iNumArchivo)
    
    vlNumArch = iNumArchivo
    Call flCargarDatos(vlNumArch)
    
End Function


Private Sub flCargarDatos(iNumArc)
On Error GoTo Err_CargaD

    vgSql = ""
    vgSql = "select gls_nomarch,fec_desde,fec_hasta,num_polizas,mto_primas,"
    vgSql = vgSql & "cod_usuariocrea,fec_crea,hor_crea "
    vgSql = vgSql & "from pd_tmae_estcarconpri "
    vgSql = vgSql & "where num_archivo = " & Trim(iNumArc) & " "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        Lbl_Archivo = vgRs!gls_nomarch
        Txt_Desde = DateSerial(Mid((vgRs!Fec_desde), 1, 4), Mid((vgRs!Fec_desde), 5, 2), Mid((vgRs!Fec_desde), 7, 2))
        Txt_Hasta = DateSerial(Mid((vgRs!fec_hasta), 1, 4), Mid((vgRs!fec_hasta), 5, 2), Mid((vgRs!fec_hasta), 7, 2))
        Lbl_TotPol = Format(vgRs!NUM_POLIZAS, "#,#0")
        Lbl_MtoTotPri = Format(vgRs!mto_primas, "#,#0.00")
        Lbl_Usuario = vgRs!Cod_UsuarioCrea
        Lbl_FecCrea = DateSerial(Mid((vgRs!Fec_Crea), 1, 4), Mid((vgRs!Fec_Crea), 5, 2), Mid((vgRs!Fec_Crea), 7, 2))
        Lbl_HorCrea = Mid(vgRs!Hor_Crea, 1, 2) & ":" & Mid(vgRs!Hor_Crea, 3, 2) & ":" & Mid(vgRs!Hor_Crea, 5, 2)
    Else
        MsgBox "La Póliza Ingresada no contiene Información", vbCritical, "Operación Cancelada"
        Exit Sub
    End If

Exit Sub
Err_CargaD:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
