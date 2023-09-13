VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Frm_CalTraConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Pólizas Traspasadas."
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9135
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
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   7455
      Begin VB.TextBox Txt_Desde 
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Txt_Hasta 
         Height          =   285
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Cmd_Buscar 
         Height          =   375
         Left            =   6000
         Picture         =   "Frm_CalTraConsulta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Efectuar Busqueda"
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Rango de Fechas   :"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
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
         Left            =   3800
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Lbl_Nombre 
         Alignment       =   2  'Center
         Caption         =   "Traspasos Efectuados a Sistema de Pago de Pensiones"
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
         Left            =   1200
         TabIndex        =   8
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   8895
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5160
         Picture         =   "Frm_CalTraConsulta.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3000
         Picture         =   "Frm_CalTraConsulta.frx":01FC
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
         Picture         =   "Frm_CalTraConsulta.frx":08B6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   720
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   4335
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   7646
      _Version        =   393216
      BackColor       =   14745599
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Frm_CalTraConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlFechaDesde As String
Dim vlFechaHasta As String
Dim vlFecTraspaso As String
Dim vlMtoPriRec As String
Dim vlFecTraPagoPen As String
Dim vlFecVigencia As String
Dim vlRutAfiliado As String
Dim vlGlsNombre As String
Dim vlTipoPension As String
Dim vlTipoRenta As String
Dim vlAnnosDif As String
Dim vlModalidad As String
Dim vlArchivo As String
Dim vlafp As String
Dim vlSql As String
Dim vlCodScomp As String
Dim vlNombrePrimero As String
Dim vlNombreSegundo As String
Dim vlApellidoPrimero As String
Dim vlApellidoSegundo As String

Dim i As Integer, vlLen As Integer, n As Integer
Dim icodtabla As String, icanlin As Integer, inom As String
Dim vlStr As Variant
Dim vlNom As String


Function flInicializaGrilla()

    Msf_Grilla.Clear
    Msf_Grilla.Cols = 15
    Msf_Grilla.rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0
        
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Nº Póliza"
    Msf_Grilla.ColWidth(0) = 1200
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Fec. Trasp. Prima"
    Msf_Grilla.ColWidth(1) = 1400
    
    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Mto. Prima Rec."
    Msf_Grilla.ColWidth(2) = 1250
    
    Msf_Grilla.Col = 3
    Msf_Grilla.Text = "Fec. Trasp. PP"
    Msf_Grilla.ColWidth(3) = 1300
    
    Msf_Grilla.Col = 4
    Msf_Grilla.Text = "Fec. Vigencia"
    Msf_Grilla.ColWidth(4) = 1200
    
    Msf_Grilla.Col = 5
    Msf_Grilla.Text = "CUSPP Afiliado"
    Msf_Grilla.ColWidth(5) = 1450

    Msf_Grilla.Col = 6
    Msf_Grilla.Text = "Nombre"
    Msf_Grilla.ColWidth(6) = 3000
    
    Msf_Grilla.Col = 7
    Msf_Grilla.Text = "Tipo Pensión"
    Msf_Grilla.ColWidth(7) = 2500
    
    Msf_Grilla.Col = 8
    Msf_Grilla.Text = "Tipo Renta"
    Msf_Grilla.ColWidth(8) = 4000
    
    Msf_Grilla.Col = 9
    Msf_Grilla.Text = "Años Dif."
    Msf_Grilla.ColWidth(9) = 900
    
    Msf_Grilla.Col = 10
    Msf_Grilla.Text = "Modalidad"
    Msf_Grilla.ColWidth(10) = 1800
    
    Msf_Grilla.Col = 11
    Msf_Grilla.Text = "Meses Gar."
    Msf_Grilla.ColWidth(11) = 1000

    Msf_Grilla.Col = 12
    Msf_Grilla.Text = "A.F.P."
    Msf_Grilla.ColWidth(12) = 1500
    
    Msf_Grilla.Col = 13
    Msf_Grilla.Text = "Moneda"
    Msf_Grilla.ColWidth(13) = 700
    
    Msf_Grilla.Col = 14
    Msf_Grilla.Text = "Pensión"
    Msf_Grilla.ColWidth(14) = 1000
    
End Function

Function flCargaGrilla()
On Error GoTo Err_Carga
    
    Call flInicializaGrilla
    While Not vgRs.EOF
    
          vgSql = ""
          vgSql = "SELECT gls_nomben,gls_nomsegben,gls_patben,gls_matben "
          vgSql = vgSql & "FROM pd_tmae_polben "
          vgSql = vgSql & "WHERE (num_poliza = '" & (vgRs!Num_Poliza) & "') AND "
          vgSql = vgSql & "(cod_par = '99') "
          Set vgRs2 = vgConexionBD.Execute(vgSql)
          If Not vgRs2.EOF Then
             vgSql = ""
             vgSql = "SELECT fec_traspaso,mto_prirec "
             vgSql = vgSql & "FROM pd_tmae_polprirec "
             vgSql = vgSql & "WHERE num_poliza = '" & (vgRs!Num_Poliza) & "' "
             Set vgRs3 = vgConexionBD.Execute(vgSql)
             If Not vgRs3.EOF Then
              
                vlFecTraspaso = DateSerial(Mid((vgRs3!FEC_TRASPASO), 1, 4), Mid(vgRs3!FEC_TRASPASO, 5, 2), Mid(vgRs3!FEC_TRASPASO, 7, 2))
                vlMtoPriRec = Format((vgRs3!MTO_PRIREC), "###,###,##0.00")
                vlFecTraPagoPen = DateSerial(Mid((vgRs!fec_trapagopen), 1, 4), Mid(vgRs!fec_trapagopen, 5, 2), Mid(vgRs!fec_trapagopen, 7, 2))
                vlFecVigencia = DateSerial(Mid((vgRs!Fec_Vigencia), 1, 4), Mid(vgRs!Fec_Vigencia, 5, 2), Mid(vgRs!Fec_Vigencia, 7, 2))
                       
                'controla la existencia del primer nombre
                If Not IsNull(vgRs2!Gls_NomBen) Then
                    vlNombrePrimero = Trim(vgRs2!Gls_NomBen) & " "
                Else: vlNombrePrimero = ""
                End If
                                
                'Controla la existencia del segundo nombre
                If Not IsNull(vgRs2!Gls_NomSegBen) Then
                    vlNombreSegundo = Trim(vgRs2!Gls_NomSegBen) & " "
                Else: vlNombreSegundo = ""
                End If
                
                'Controla la existencia del primer apellido
                If Not IsNull(vgRs2!Gls_PatBen) Then
                    vlApellidoPrimero = Trim(vgRs2!Gls_PatBen) & " "
                Else: vlApellidoPrimero = ""
                End If
                
                'Controla la existencia del segundo apellido
                If Not IsNull(vgRs2!Gls_MatBen) Then
                    vlApellidoSegundo = Trim(vgRs2!Gls_MatBen) & " "
                Else: vlApellidoSegundo = ""
                End If
                
                vlGlsNombre = (vlNombrePrimero) + (vlNombreSegundo) + (vlApellidoPrimero) + (vlApellidoSegundo)
                'vlGlsNombre = Trim(vgRs2!gls_nomben) + " " + Trim(vgRs2!gls_nomsegben) + " " + Trim(vgRs2!gls_patben) + " " + Trim(vgRs2!gls_matben)

'I--- ABV 05/02/2011 ---
'                vlCodScomp = fgObtenerCodMonedaScomp(egTablaMoneda(), vgNumeroTotalTablasMoneda, vgRs!Cod_Moneda)
                If IsNull(vgRs!Cod_Moneda) Then
                    vlCodScomp = vgRs!Cod_Monedapol
                Else
                    vlCodScomp = vgRs!Cod_Moneda
                End If
'F--- ABV 05/02/2011 ---
              
                vlTipoPension = fgBuscarGlosaElemento(vgCodTabla_TipPen, (vgRs!Cod_TipPension))
                vlTipoPension = " " + Trim(vgRs!Cod_TipPension) + " - " + Trim(vlTipoPension)
                vlTipoRenta = fgBuscarGlosaElemento(vgCodTabla_TipRen, (vgRs!Cod_TipRen))
                vlTipoRenta = " " + Trim(vgRs!Cod_TipRen) + " - " + Trim(vlTipoRenta)
                vlAnnosDif = ((vgRs!Num_MesDif) / 12)
                vlModalidad = fgBuscarGlosaElemento(vgCodTabla_AltPen, (vgRs!Cod_Modalidad))
                vlModalidad = " " + Trim(vgRs!Cod_Modalidad) + " - " + Trim(vlModalidad)
                vlafp = fgBuscarGlosaElemento(vgCodTabla_AFP, (vgRs!Cod_AFP))
                vlTipoPension = fgBuscarGlosaElemento(vgCodTabla_TipPen, (vgRs!Cod_TipPension))
                                                    
                Msf_Grilla.AddItem Trim(vgRs!Num_Poliza) & vbTab _
                        & Trim(vlFecTraspaso) & vbTab _
                        & Trim(vlMtoPriRec) & vbTab _
                        & Trim(vlFecTraPagoPen) & vbTab _
                        & Trim(vlFecVigencia) & vbTab _
                        & " " & (vgRs!Cod_Cuspp) & vbTab _
                        & Trim(vlGlsNombre) & vbTab _
                        & (vlTipoPension) & vbTab _
                        & (vlTipoRenta) & vbTab _
                        & Trim(vlAnnosDif) & vbTab _
                        & (vlModalidad) & vbTab _
                        & Trim(vgRs!Num_MesGar) & vbTab _
                        & Trim(vlafp) & vbTab _
                        & Trim(vlCodScomp) & vbTab _
                        & Trim(vgRs!Mto_Pension)
                        
                        vgRs.MoveNext
             Else
                 Screen.MousePointer = 11
                 MsgBox "Datos de la Prima de la Póliza Recibida No se Encuentran Registrados", vbInformation, "Información"
                 Screen.MousePointer = 0
                 Exit Function
             End If
          Else
              Screen.MousePointer = 11
              MsgBox "El Causante de la Póliza '" & Trim(vgRs!Num_Poliza) & "' traspasada al Pago de Pensiones, No se Encuentra Registrada", vbInformation, "Información"
              Screen.MousePointer = 0
              Exit Function
          End If
    Wend

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flCarCodRep(icodtabla, icanlin, inom)

On Error GoTo Err_Codigos
    
    Sql = ""
    Sql = Sql & "SELECT cod_elemento, gls_elemento FROM ma_tpar_tabcod "
    Sql = Sql & "WHERE cod_tabla = '" & icodtabla & "' "
    Sql = Sql & " ORDER BY cod_elemento ASC"
    Set vgRs = vgConexionBD.Execute(Sql)
    For n = 1 To icanlin
        vlLen = 0
        vlStr = ""
        Do While Not vgRs.EOF And vlLen <= 200
           If Trim(icodtabla) = Trim(vgCodTabla_TipPen) Then
              If ((vgRs!cod_elemento) >= 4) And ((vgRs!cod_elemento) <= 8) Then
                  vlStr = vlStr + (vgRs!cod_elemento) + " - " + (vgRs!GLS_ELEMENTO) + " / "
                  vlLen = Len(vlStr)
              End If
           Else
               vlStr = vlStr + (vgRs!cod_elemento) + " - " + (vgRs!GLS_ELEMENTO) + " / "
               vlLen = Len(vlStr)
           End If
           vgRs.MoveNext
        Loop
        If vlLen <> 0 Then
            vlStr = Mid(vlStr, 1, vlLen - 3)
            Rpt_General.Formulas(i) = inom & "= '" & Trim(vlStr) & "'"
            i = i + 1
        End If
    Next n
Exit Function
Err_Codigos:
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
        
    vgSql = ""
    vgSql = "SELECT p.num_poliza,p.fec_trapagopen,p.fec_vigencia,"
    vgSql = vgSql & "p.cod_cuspp,p.cod_tippension,p.cod_tipren,p.num_mesdif,"
    vgSql = vgSql & "p.cod_modalidad,p.num_mesgar,p.cod_afp," 'p.cod_moneda," 'I--- ABV 05/02/2011 ---
    vgSql = vgSql & "p.mto_pension "
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & ",p.cod_moneda as cod_monedapol "
    vgSql = vgSql & ",mtr.cod_scomp as cod_moneda "
'F--- ABV 05/02/2011 ---
    vgSql = vgSql & "FROM pd_tmae_poliza p "
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & ",ma_tpar_monedatiporeaju mtr "
'F--- ABV 05/02/2011 ---
    vgSql = vgSql & "WHERE (p.fec_trapagopen >= '" & Trim(vlFechaDesde) & "') AND "
    vgSql = vgSql & "(p.fec_trapagopen <= '" & Trim(vlFechaHasta) & "') AND "
    vgSql = vgSql & "(p.cod_trapagopen = 'S') "
'I--- ABV 05/02/2011 ---
    vgSql = vgSql & "AND p.cod_tipreajuste = mtr.cod_tipreajuste(+) AND "
    vgSql = vgSql & "p.cod_moneda = mtr.cod_moneda(+) "
'F--- ABV 05/02/2011 ---
'RRR
    vgSql = vgSql & "and num_endoso=1 "
'RRR

    vgSql = vgSql & " ORDER BY num_poliza "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       Call flCargaGrilla
       Fra_Fechas.Enabled = False
    Else
        Screen.MousePointer = 11
        MsgBox "No Existen Polizas Traspasadas en el Período Ingresado", vbInformation, "Información"
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
Dim objRep As New ClsReporte 'RRR1305/2013
On Error GoTo Err_CmdImprimir

    vgPalabra = "Periodo del " & Txt_Desde.Text & " al " & Txt_Hasta.Text

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
   
    Screen.MousePointer = 11
    
    vlArchivo = strRpt & "pd_rpt_CalTraConsulta.rpt"   '\Reportes
    If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Consulta de Pólizas Traspasadas no se encuentra en el Directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
    End If
 
'    vgQuery = "select distinct A.NUM_POLIZA, A.FEC_TRAPAGOPEN, G.GLS_ELEMENTO COD_TIPPENSION, A.COD_CUSPP, A.FEC_VIGENCIA, E.GLS_ELEMENTO COD_TIPREN, A.NUM_MESDIF, F.GLS_ELEMENTO COD_MODALIDAD, A.NUM_MESGAR, A.MTO_PENSION,"
'    vgQuery = vgQuery & " A.MTO_VALREAJUSTEMEN, A.MTO_VALREAJUSTETRI,"
'    vgQuery = vgQuery & " B.Num_Orden , Gls_NomBen, Gls_NomSegBen, Gls_PatBen, Gls_MatBen, c.FEC_TRASPASO, c.MTO_PRIREC, D.COD_SCOMP, H.GLS_ELEMENTO AFP, I.GLS_NOMCOR, I.GLS_PATCOR, I.GLS_MATCOR, A.PRC_CORCOMREAL"
'    vgQuery = vgQuery & " FROM PD_TMAE_POLIZA A"
'    vgQuery = vgQuery & " JOIN PD_TMAE_POLBEN B ON A.NUM_POLIZA=B.NUM_POLIZA"
'    vgQuery = vgQuery & " JOIN PD_TMAE_POLPRIREC C ON B.NUM_POLIZA=C.NUM_POLIZA"
'    vgQuery = vgQuery & " JOIN MA_TPAR_MONEDATIPOREAJU D ON A.COD_MONEDA=D.COD_MONEDA AND A.COD_TIPREAJUSTE=D.COD_TIPREAJUSTE"
'    vgQuery = vgQuery & " JOIN MA_TPAR_TABCOD E ON A.COD_TIPREN=E.COD_ELEMENTO AND E.COD_TABLA='TR'"
'    vgQuery = vgQuery & " JOIN MA_TPAR_TABCOD F ON A.COD_MODALIDAD=F.COD_ELEMENTO AND F.COD_TABLA='AL'"
'    vgQuery = vgQuery & " JOIN MA_TPAR_TABCOD G ON A.COD_TIPPENSION=G.COD_ELEMENTO AND G.COD_TABLA='TP'"
'    vgQuery = vgQuery & " JOIN MA_TPAR_TABCOD H ON A.COD_AFP=H.COD_ELEMENTO AND H.COD_TABLA='AF'"
'    vgQuery = vgQuery & " JOIN PT_TMAE_CORREDOR I ON A.NUM_IDENCOR=I.NUM_IDENCOR"
'    vgQuery = vgQuery & " WHERE A.COD_TRAPAGOPEN='S' AND FEC_TRAPAGOPEN >= '" & Trim(vlFechaDesde) & "' AND FEC_TRAPAGOPEN <= '" & Trim(vlFechaHasta) & "' AND B.COD_PAR=99 ORDER BY 1"
    
    vgQuery = "select distinct A.NUM_POLIZA, A.FEC_TRAPAGOPEN, G.GLS_ELEMENTO COD_TIPPENSION, A.COD_CUSPP, A.FEC_VIGENCIA,"
    vgQuery = vgQuery & " E.GLS_ELEMENTO COD_TIPREN, A.NUM_MESDIF, F.GLS_ELEMENTO, A.COD_MODALIDAD, A.NUM_MESGAR, A.MTO_PENSION, A.MTO_VALREAJUSTEMEN,"
    vgQuery = vgQuery & " A.MTO_VALREAJUSTETRI, b.Num_Orden , Gls_NomBen, Gls_NomSegBen, Gls_PatBen, Gls_MatBen, c.FEC_TRASPASO, c.MTO_PRIREC,"
    vgQuery = vgQuery & " D.COD_SCOMP, H.GLS_ELEMENTO AFP, I.GLS_NOMCOR, I.GLS_PATCOR, I.GLS_MATCOR, A.PRC_CORCOMREAL"
    vgQuery = vgQuery & " FROM PD_TMAE_POLIZA A"
    vgQuery = vgQuery & " JOIN PD_TMAE_POLBEN B ON A.NUM_POLIZA=B.NUM_POLIZA"
    vgQuery = vgQuery & " JOIN PD_TMAE_POLPRIREC C ON B.NUM_POLIZA=C.NUM_POLIZA"
    vgQuery = vgQuery & " JOIN MA_TPAR_MONEDATIPOREAJU D ON A.COD_MONEDA=D.COD_MONEDA AND A.COD_TIPREAJUSTE=D.COD_TIPREAJUSTE"
    vgQuery = vgQuery & " JOIN MA_TPAR_TABCOD E ON A.COD_TIPREN=E.COD_ELEMENTO AND E.COD_TABLA='TR'"
    vgQuery = vgQuery & " JOIN MA_TPAR_TABCOD F ON A.COD_MODALIDAD=F.COD_ELEMENTO AND F.COD_TABLA='AL'"
    vgQuery = vgQuery & " JOIN MA_TPAR_TABCOD G ON A.COD_TIPPENSION=G.COD_ELEMENTO AND G.COD_TABLA='TP'"
    vgQuery = vgQuery & " JOIN MA_TPAR_TABCOD H ON A.COD_AFP=H.COD_ELEMENTO AND H.COD_TABLA='AF'"
    vgQuery = vgQuery & " JOIN PT_TMAE_CORREDOR I ON A.NUM_IDENCOR=I.NUM_IDENCOR"
    vgQuery = vgQuery & " WHERE A.COD_TRAPAGOPEN='S' AND FEC_TRAPAGOPEN >= '" & Trim(vlFechaDesde) & "'  AND FEC_TRAPAGOPEN <= '" & Trim(vlFechaHasta) & "' AND b.COD_PAR=99 ORDER BY 1"
    
    
    Set vgRs = vgConexionBD.Execute(vgQuery)
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(vgRs, Replace(UCase(strRpt & "Estructura\pd_rpt_CalTraConsulta.rpt"), ".RPT", ".TTX"), 1)

    If objRep.CargaReporte(strRpt & "", "pd_rpt_CalTraConsulta.rpt", "PrePoliza", vgRs, True, _
                            ArrFormulas("NombreCompania", vgNombreCompania), _
                            ArrFormulas("NombreSistema", vgNombreSistema), _
                            ArrFormulas("NombreSubSistema", "vgNombreSubSistema"), _
                            ArrFormulas("Periodo", vgPalabra), _
                            ArrFormulas("Moneda", vlCodScomp)) = False Then
            
      'If objRep.CargaReporte(strRpt & "", "PD_Rpt_Poliza.rpt", "PrePoliza", vgRs, True) = False Then
            
       MsgBox "No se pudo abrir el reporte", vbInformation
       Exit Sub
    End If
   
    
    
  

    'vgQuery = ""
    'vgQuery = vgQuery & "{pd_TMAE_POLIZA.fec_trapagopen} >= '" & Trim(vlFechaDesde) & "' AND "
    'vgQuery = vgQuery & "{pd_TMAE_POLIZA.fec_trapagopen} <= '" & Trim(vlFechaHasta) & "' AND "
    'vgQuery = vgQuery & "{pd_TMAE_POLIZA.cod_trapagopen} = 'S' "

    'Rpt_General.Reset
    'Rpt_General.ReportFileName = vlArchivo     'App.Path & "\rpt_Areas.rpt"
    'Rpt_General.Connect = vgRutaDataBase       ' o App.Path & "\Nestle.mdb"
    'Rpt_General.SelectionFormula = vgQuery
  '
   ' Rpt_General.Formulas(0) = ""
   ' Rpt_General.Formulas(1) = ""
   ' Rpt_General.Formulas(2) = ""
   ' Rpt_General.Formulas(3) = ""
   ' Rpt_General.Formulas(4) = ""
'
 '   vgPalabra = Txt_Desde.Text & "  *  " & Txt_Hasta.Text
 '   Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
 '   Rpt_General.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
 '   Rpt_General.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
 '   Rpt_General.Formulas(3) = "Periodo = '" & vgPalabra & "'"
 '   Rpt_General.Formulas(4) = "Moneda = '" & vlCodScomp & "'"

 '   i = 4
 '   vlLen = 0
     
    'Busca los códigos de Tipo de Pensión y su glosa
  '  Call flCarCodRep(vgCodTabla_TipPen, 1, "TipoPension")
    'Busca los códigos de Tipo de Renta y su glosa
  '  Call flCarCodRep(vgCodTabla_TipRen, 1, "TipoRenta")
    'Busca los códigos de Modalidad y su glosa
  '  Call flCarCodRep(vgCodTabla_AltPen, 1, "Modalidad")
   
  '  Rpt_General.WindowState = crptMaximized
  '  Rpt_General.Destination = crptToWindow
  '  Rpt_General.WindowTitle = "Informe de Consulta de Pólizas Traspasadas"
  '  Rpt_General.Action = 1
  '  Screen.MousePointer = 0
   
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
On Error GoTo Err_CmdSalir

    Unload Me
    
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

    Frm_CalTraConsulta.Top = 0
    Frm_CalTraConsulta.Left = 0
    Call flInicializaGrilla
    
    Call fgCargarTablaMoneda(vgCodTabla_TipMon, egTablaMoneda, vgNumeroTotalTablasMoneda)
    
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
