VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frm_BuscaCoti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscador de Cotización"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   0
      TabIndex        =   18
      Top             =   5280
      Width           =   11775
      Begin VB.CommandButton cmdExport 
         Caption         =   "Exportar"
         Height          =   675
         Left            =   4560
         Picture         =   "Frm_BuscaCoti.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_BuscarCotizaciones 
         Caption         =   "Todas"
         Height          =   675
         Left            =   3480
         Picture         =   "Frm_BuscaCoti.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Buscar Todas las Póliza"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   5520
         Picture         =   "Frm_BuscaCoti.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6960
         Picture         =   "Frm_BuscaCoti.frx":0BFE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   11775
      Begin VB.TextBox Txt_ApellMat 
         Height          =   285
         Left            =   10080
         TabIndex        =   7
         Top             =   600
         Width           =   1600
      End
      Begin VB.TextBox Txt_ApellPat 
         Height          =   285
         Left            =   8440
         TabIndex        =   6
         Top             =   600
         Width           =   1650
      End
      Begin VB.TextBox Txt_Nombre 
         Height          =   285
         Left            =   6800
         TabIndex        =   5
         Top             =   600
         Width           =   1660
      End
      Begin VB.TextBox Txt_NumIden 
         Height          =   285
         Left            =   5560
         TabIndex        =   4
         Top             =   600
         Width           =   1255
      End
      Begin VB.TextBox Txt_TipoIden 
         Height          =   285
         Left            =   4320
         TabIndex        =   3
         Top             =   600
         Width           =   1255
      End
      Begin VB.TextBox Txt_Cuspp 
         Height          =   285
         Left            =   2480
         TabIndex        =   2
         Top             =   600
         Width           =   1855
      End
      Begin VB.TextBox Txt_SolOfe 
         Height          =   285
         Left            =   1350
         TabIndex        =   1
         Top             =   600
         Width           =   1135
      End
      Begin VB.TextBox Txt_cotiz 
         Height          =   285
         Left            =   120
         MaxLength       =   12
         TabIndex        =   0
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "CUSPP"
         Height          =   255
         Index           =   8
         Left            =   2480
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nº Indent."
         Height          =   255
         Index           =   7
         Left            =   5560
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nº Operación"
         Height          =   255
         Index           =   6
         Left            =   1350
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Ap. Materno"
         Height          =   255
         Index           =   4
         Left            =   10080
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Ap. Paterno"
         Height          =   255
         Index           =   3
         Left            =   8440
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   6800
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Tipo Indent."
         Height          =   255
         Index           =   1
         Left            =   4395
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nº Cotización"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_GrillaBuscaCot 
      Height          =   4275
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1080
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7541
      _Version        =   393216
      BackColor       =   14745599
   End
End
Attribute VB_Name = "Frm_BuscaCoti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlNumCot As String
Dim vlCodSolOfe As String
Dim vlcuspp As String
Dim vlNomAfi As String
Dim vlApellPat As String
Dim vlApellMat As String
Dim vlTipoIdenAfi As String
Dim vlNumIdenAfi As String
Dim vlCodTipCot As String
Dim Fila As Long

Dim vlTablaDetCotizacion As String
Dim vlSql As String

Const clCodTipCotOfe As String * 1 = "O"
Const clCodTipCotExt As String * 1 = "E"
Const clCodTipCotRmt As String * 1 = "R"

Function flLimpiar()
On Error GoTo Err_Limpia

    Txt_cotiz.Text = ""
    Txt_SolOfe.Text = ""
    Txt_Cuspp.Text = ""
    Txt_TipoIden.Text = ""
    Txt_NumIden.Text = ""
    Txt_Nombre.Text = ""
    Txt_ApellPat.Text = ""
    Txt_ApellMat.Text = ""
    
Exit Function
Err_Limpia:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrilla()
Dim vlCuenta As Integer
Dim vlColumna As Integer
Dim vlNumIdent As String
Dim vlTipoIdenAfi As String
Dim vlNombre As String
Dim vlPaterno As String
Dim vlMaterno As String

'On Error GoTo Err_Carga
    Msf_GrillaBuscaCot.rows = 1
    
    Msf_GrillaBuscaCot.Enabled = True
    Msf_GrillaBuscaCot.Cols = 10
    Msf_GrillaBuscaCot.rows = 1
    Msf_GrillaBuscaCot.Row = 0
    
    Msf_GrillaBuscaCot.Col = 0
    Msf_GrillaBuscaCot.ColWidth(0) = 0
    
    Msf_GrillaBuscaCot.Col = 1
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.ColWidth(1) = 1300
    Msf_GrillaBuscaCot.Text = "Nº Cotización"
    Msf_GrillaBuscaCot.CellFontBold = True
        
    Msf_GrillaBuscaCot.Col = 2
    Msf_GrillaBuscaCot.ColWidth(2) = 1200
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Nº Operación"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 3
    Msf_GrillaBuscaCot.ColWidth(3) = 1750
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "CUSPP"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 4
    Msf_GrillaBuscaCot.ColWidth(4) = 1300
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Tipo Ident."
    Msf_GrillaBuscaCot.CellFontBold = True

    Msf_GrillaBuscaCot.Col = 5
    Msf_GrillaBuscaCot.ColWidth(5) = 1250
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Nº Ident."
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 6
    Msf_GrillaBuscaCot.ColWidth(6) = 1600
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Nombre"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 7
    Msf_GrillaBuscaCot.ColWidth(7) = 1600
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.CellFontBold = True
    Msf_GrillaBuscaCot.Text = "Ap. Paterno"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 8
    Msf_GrillaBuscaCot.ColWidth(8) = 1350
    Msf_GrillaBuscaCot.CellFontBold = True
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Ap. Materno"
    Msf_GrillaBuscaCot.CellFontBold = True

    Msf_GrillaBuscaCot.Col = 9
    Msf_GrillaBuscaCot.ColWidth(9) = 0
    Msf_GrillaBuscaCot.CellFontBold = True
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "CodTipCot"
    Msf_GrillaBuscaCot.CellFontBold = True

Exit Function
Err_Carga:
      Screen.MousePointer = 0
      Select Case Err
        Case Else
          MsgBox "Error grave [" & Err & Space(4) & Err.Description & "]", vbCritical
      End Select
End Function

Private Sub Btn_Salir_Click()
On Error GoTo Err_Volver
    If vgFormulario = "C" Then
        Frm_CalCotizacion.flLimpiarAfiliado ("D")
        Frm_CalCotizacion.Fra_Inter.Enabled = False
        Frm_CalCotizacion.SSTab_Cotiz.Enabled = False
    Else
        If vgFormulario = "P" Then
            Frm_CalPoliza.Enabled = True
        End If
    End If
    Unload Me
Exit Sub
Err_Volver:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Btn_Limpiar_Click()
On Error GoTo Err_Limpiar

    Msf_GrillaBuscaCot.rows = 1
    Call flLimpiar
    Txt_cotiz.SetFocus

Exit Sub
Err_Limpiar:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Cmd_BuscarCotizaciones_Click()
On Error GoTo Err_BuscarCotizaciones
    
    Call plGenerarConsulta(True)
    
Exit Sub
Err_BuscarCotizaciones:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub cmdExport_Click()


    
        Dim objCmd As ADODB.Command
        Dim RS As ADODB.Recordset
        Dim conn As ADODB.Connection
            
        Set RS = New ADODB.Recordset
        Set conn = New ADODB.Connection
        
        Dim Texto As String
        
        
        Set conn = New ADODB.Connection
        Set RS = New ADODB.Recordset
        Set objCmd = New ADODB.Command
        
        conn.Provider = "OraOLEDB.Oracle"
        conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
        conn.CursorLocation = adUseClient
        conn.Open
  
                       
        Set objCmd = CreateObject("ADODB.Command")
        Set objCmd.ActiveConnection = conn
        
        objCmd.CommandText = "SP_EXPORCOTIZACION"
        objCmd.CommandType = adCmdStoredProc
        
            
        Set RS = objCmd.Execute
        
        Exportar RS
                       
                   
        conn.Close
        Set RS = Nothing
        Set conn = Nothing

    


End Sub

Private Sub Exportar(ByRef RS As ADODB.Recordset)

    Dim Obj_Excel As Object
    Dim Obj_Libro As Object
    Dim Obj_Hoja As Object
    Dim vFila As Integer
    Dim vColumna As Integer
    
    Set Obj_Excel = CreateObject("Excel.Application")
    
    Set Obj_Libro = Obj_Excel.Workbooks.Add
    Set Obj_Hoja = Obj_Libro.Worksheets.Add
    
    vFila = 0
    vColumna = 0
    
    Dim Titulos(10) As String
      
    Titulos(0) = "NUMCOT"
    Titulos(1) = "NOMBEN"
    Titulos(2) = "APABEN"
    Titulos(3) = "AMABEN"
    Titulos(4) = "CUSPP"
    Titulos(5) = "TIPOIDEN"
    Titulos(6) = "NUMIDEN"
    Titulos(7) = "COD_ESTCOT"
    Titulos(8) = "COD OPERACION"
    Titulos(9) = "COD_TIPCOT"
    
    vTotaCampos = 10
      
    vFila = 1
    Obj_Hoja.Cells(vFila, 1) = "Exportacion de Cotizaciones"
    Obj_Hoja.rows(vFila).Font.Bold = True
    vFila = 2
    'Obj_Hoja.Cells(vFila, 1) = "Periodo: " & CmbMesExtrae.Text
    'Obj_Hoja.rows(vFila).Font.Bold = True
   
    vFila = 3
    
    For i = 0 To UBound(Titulos)
    
          Obj_Hoja.Cells(vFila, i + 1) = Titulos(i)
    
    Next
    
     Obj_Hoja.Columns("A:J").NumberFormat = "@"
    
     Obj_Hoja.rows(3).Font.Bold = True
     Obj_Hoja.Range("A3:J3").Borders.LineStyle = xlContinuous
     Obj_Hoja.Range("A3:J3").VerticalAlignment = xlVAlignCenter

          
      Do While Not RS.EOF
               vFila = vFila + 1
              For vColumna = 0 To vTotaCampos - 1
                Obj_Hoja.Cells(vFila, vColumna + 1) = RS.Fields(vColumna).Value
                
              Next
    
         
            Me.Refresh
            
        RS.MoveNext
        
      Loop
      
    Obj_Hoja.Columns("A:J").AutoFit
      
    Obj_Excel.Visible = True
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing
    
    
    
         
      


End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Me.Top = 0
    Me.Left = 0
    
    Frm_BuscaCoti.Caption = "Buscador de Cotización"
    Call flCargaGrilla
    
    Call flLimpiar
    
Exit Sub
Err_Cargar:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Unload
    If vgFormulario = "C" Then
        Frm_CalCotizacion.Enabled = True
    End If
    If vgFormulario = "P" Then
        Frm_CalPoliza.Enabled = True
    End If

Exit Sub
Err_Unload:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_GrillaBuscaCot_DblClick()
On Error GoTo Err_Seleccionar

    Msf_GrillaBuscaCot.Col = 1
    Msf_GrillaBuscaCot.Row = Msf_GrillaBuscaCot.RowSel
    If (Not (Msf_GrillaBuscaCot.Text = "") And (Msf_GrillaBuscaCot.Row <> 0)) Then
        Msf_GrillaBuscaCot.Col = 9
        vlCodTipCot = Msf_GrillaBuscaCot.Text
        Msf_GrillaBuscaCot.Col = 1
        If vgFormulario = "C" Then
            vlNumCot = Msf_GrillaBuscaCot.Text
            Call Frm_CalCotizacion.flBuscarDatosCot(vlNumCot, vlCodTipCot)
            Unload Me
        Else
            If vgFormulario = "P" Then
                vlNumCot = Msf_GrillaBuscaCot.Text
                Call Frm_CalPoliza.flBuscaCotizacion(vlNumCot, vlCodTipCot)
                Unload Me
            End If
        End If
    Else
       MsgBox "No Hay Datos Para Modificar", vbInformation, "No Hay Datos "
    End If

Exit Sub
Err_Seleccionar:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_ApellMat_KeyPress(KeyAscii As Integer)
On Error GoTo Err_KeyPress

    If KeyAscii = 13 Then
        
        Txt_ApellMat = UCase(Trim(Txt_ApellMat))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_SolOfe) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        
        End If
        Cmd_BuscarCotizaciones.SetFocus
    End If
    
Exit Sub
Err_KeyPress:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
'On Error GoTo Err_KeyPress
'
'    If KeyAscii = 13 Then
'        Txt_Nombre = Trim(UCase(Txt_Nombre))
'        Txt_ApellPat = Trim(UCase(Txt_ApellPat))
'        Txt_ApellMat = Trim(UCase(Txt_ApellMat))
'        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_SolOfe) <> "") Or (Trim(Txt_Cuspp) <> "") Or (Trim(Txt_Nombre) <> "") Or (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
'            vlSql = ""
'            vlSql = "SELECT c.num_cot,c.cod_tipcot "
'            vlSql = vlSql & " FROM pt_tmae_cotizacion c "
'            vlSql = vlSql & " ORDER BY c.num_cot "
'            Set vgRs1 = vgConexionBD.Execute(vlSql)
'            If Not vgRs1.EOF Then
'                vlCodTipCot = (vgRs1!cod_tipcot)
'                If Trim(vgRs1!cod_tipcot) = clCodTipCotOfe Then
'                    vlTablaDetCotizacion = "pt_tmae_detcotizacion"
'                End If
'                If Trim(vgRs1!cod_tipcot) = clCodTipCotExt Then
'                    vlTablaDetCotizacion = "pt_tmae_detcotexterna"
'                End If
'                If Trim(vgRs1!cod_tipcot) = clCodTipCotRmt Then
'                    vlTablaDetCotizacion = "pt_tmae_detcotremate"
'                End If
'
'                Fila = 1
'                Msf_GrillaBuscaCot.Rows = 1
'                While Not vgRs1.EOF
'
''I--- ABV 22/06/2006 ---
'                    vlCodTipCot = (vgRs1!cod_tipcot)
'                    If Trim(vgRs1!cod_tipcot) = clCodTipCotOfe Then
'                        vlTablaDetCotizacion = "pt_tmae_detcotizacion"
'                    End If
'                    If Trim(vgRs1!cod_tipcot) = clCodTipCotExt Then
'                        vlTablaDetCotizacion = "pt_tmae_detcotexterna"
'                    End If
'                    If Trim(vgRs1!cod_tipcot) = clCodTipCotRmt Then
'                        vlTablaDetCotizacion = "pt_tmae_detcotremate"
'                    End If
''F--- ABV 22/06/2006 ---
'
'                    vlSql = ""
'                    vlSql = "SELECT b.num_cot as numcot,"
'                    vlSql = vlSql & "b.gls_nomben as nomben,"
'                    vlSql = vlSql & "b.gls_patben as apaben,"
'                    vlSql = vlSql & "b.gls_matben as amaben,"
'                    vlSql = vlSql & "b.gls_matben as amaben,"
'                    vlSql = vlSql & "c.rut_afi as rutafi, d.cod_estcot, "
'                    vlSql = vlSql & "d.cod_solofe as codigo "
'                    vlSql = vlSql & " FROM pt_tmae_cotizacion c, pt_tmae_cotben b, " & vlTablaDetCotizacion & " d "
'                    vlSql = vlSql & " WHERE b.gls_matben LIKE '" & Trim(UCase(Txt_ApellMat)) & "%'"
'                    If Trim(UCase(Txt_Nombre)) <> "" Then vlSql = vlSql & " AND b.gls_nomben LIKE '" & Trim(UCase(Txt_Nombre)) & "%'"
'                    If Trim(UCase(Txt_ApellPat)) <> "" Then vlSql = vlSql & " AND b.gls_patben LIKE '" & Trim(UCase(Txt_ApellPat)) & "%'"
'                    If Trim(Txt_cotiz) <> "" Then vlSql = vlSql & " AND b.num_cot LIKE '" & Trim(Txt_cotiz) & "%'"
'                    If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND c.rut_afiliado LIKE '" & Trim(Txt_Cuspp) & "%'"
'                    If Trim(Txt_SolOfe) <> "" Then vlSql = vlSql & " AND d.cod_solofe LIKE '" & Trim(Txt_SolOfe) & "%'"
'                    vlSql = vlSql & " AND c.num_cot = '" & Trim(vgRs1!num_cot) & "' "
'                    vlSql = vlSql & " AND b.num_cot = c.num_cot "
'                    vlSql = vlSql & " AND d.num_cot = c.num_cot "
'                    vlSql = vlSql & " AND d.cod_estcot = 'A' "
'                    vlSql = vlSql & " AND b.cod_par = '99' "
'                    vlSql = vlSql & " ORDER BY b.num_cot "
'
'                    Set vgRs = vgConexionBD.Execute(vlSql)
'                    If Not vgRs.EOF Then
'
'                        vlNumCot = ""
'                        vlCodSolOfe = ""
'                        vlTipoIdenAfi = ""
'                        vlNomAfi = ""
'                        vlApellPat = ""
'                        vlApellMat = ""
'
'                        If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
'                        If Not IsNull(vgRs!Codigo) Then vlCodSolOfe = Trim(vgRs!Codigo)
'                        If Not IsNull(vgRs!rutafi) Then vlTipoIdenAfi = Trim(vgRs!rutafi)
'                        If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
'                        If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
'                        If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)
'
'                        Msf_GrillaBuscaCot.AddItem Fila & vbTab & vlNumCot & vbTab & vlCodSolOfe & vbTab & vlTipoIdenAfi & vbTab & vlNomAfi & _
'                                              vbTab & vlApellPat & vbTab & vlApellMat & vbTab & vlCodTipCot
'
'                        Fila = Fila + 1
'                    End If
'                    vgRs1.MoveNext
'                Wend
'            End If
'            vgRs1.Close
'            vgRs.Close
'        Else
'            flCargaGrilla
'        End If
'        Btn_Limpiar.SetFocus
'    End If
'
'Exit Sub
'Err_KeyPress:
'  Screen.MousePointer = 0
'  Select Case Err
'    Case Else
'      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
'  End Select
End Sub

Private Sub Txt_ApellMat_LostFocus()

    Txt_ApellMat = Trim(UCase(Txt_ApellMat))

End Sub

Private Sub Txt_ApellPat_KeyPress(KeyAscii As Integer)
On Error GoTo Err_KeyPress

    If KeyAscii = 13 Then
        
        Txt_ApellPat = UCase(Trim(Txt_ApellPat))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_SolOfe) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        
        End If
        Txt_ApellMat.SetFocus
    End If
    
Exit Sub
Err_KeyPress:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_ApellPat_LostFocus()

    Txt_ApellPat = Trim(UCase(Txt_ApellPat))

End Sub

Private Sub Txt_Cuspp_KeyPress(KeyAscii As Integer)
On Error GoTo Err_KeyPress

    If KeyAscii = 13 Then
        
        Txt_Cuspp = UCase(Trim(Txt_Cuspp))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_SolOfe) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        
        End If
        Txt_TipoIden.SetFocus
    End If
    
Exit Sub
Err_KeyPress:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_nombre_KeyPress(KeyAscii As Integer)
On Error GoTo Err_KeyPress

    If KeyAscii = 13 Then
        
        Txt_Nombre = UCase(Trim(Txt_Nombre))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_SolOfe) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        
        End If
        Txt_ApellPat.SetFocus
    End If
    
Exit Sub
Err_KeyPress:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_Cotiz_KeyPress(KeyAscii As Integer)
On Error GoTo Err_KeyPress

    If KeyAscii = 13 Then

        Txt_cotiz = UCase(Trim(Txt_cotiz))

        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_SolOfe) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then

'            vlSql = ""
'            vlSql = "SELECT c.num_cot,c.cod_tipcot "
'            vlSql = vlSql & " FROM pt_tmae_cotizacion c "
'            vlSql = vlSql & " ORDER BY c.num_cot "
'            Set vgRs1 = vgConexionBD.Execute(vlSql)
'            If Not vgRs1.EOF Then
'                vlCodTipCot = (vgRs1!cod_tipcot)
'                If Trim(vgRs1!cod_tipcot) = clCodTipCotOfe Then
'                    vlTablaDetCotizacion = "pt_tmae_detcotizacion"
'                End If
'                If Trim(vgRs1!cod_tipcot) = clCodTipCotExt Then
'                    vlTablaDetCotizacion = "pt_tmae_detcotexterna"
'                End If
'                If Trim(vgRs1!cod_tipcot) = clCodTipCotRmt Then
'                    vlTablaDetCotizacion = "pt_tmae_detcotremate"
'                End If
'
'                fila = 1
'                Msf_GrillaBuscaCot.Rows = 1
'                While Not vgRs1.EOF
'
''I--- ABV 22/06/2006 ---
'                    vlCodTipCot = (vgRs1!cod_tipcot)
'                    If Trim(vgRs1!cod_tipcot) = clCodTipCotOfe Then
'                        vlTablaDetCotizacion = "pt_tmae_detcotizacion"
'                    End If
'                    If Trim(vgRs1!cod_tipcot) = clCodTipCotExt Then
'                        vlTablaDetCotizacion = "pt_tmae_detcotexterna"
'                    End If
'                    If Trim(vgRs1!cod_tipcot) = clCodTipCotRmt Then
'                        vlTablaDetCotizacion = "pt_tmae_detcotremate"
'                    End If
''F--- ABV 22/06/2006 ---
'
'                    vlSql = ""
'                    vlSql = "SELECT b.num_cot as numcot,b.gls_nomben as nomben,"
'                    vlSql = vlSql & " b.gls_patben as apaben,"
'                    vlSql = vlSql & " b.gls_matben as amaben,"
'                    vlSql = vlSql & "c.rut_afi as rutafi, d.cod_estcot, "
'                    vlSql = vlSql & "d.cod_solofe as codigo "
'                    vlSql = vlSql & " FROM pt_tmae_cotizacion c, pt_tmae_cotben b," & vlTablaDetCotizacion & " d "
'                    vlSql = vlSql & " WHERE "
'                    vlSql = vlSql & " b.num_cot LIKE '" & Txt_cotiz & "%'"
'                    If Trim(UCase(Txt_Nombre)) <> "" Then vlSql = vlSql & " AND b.gls_nomben LIKE '" & Trim(UCase(Txt_Nombre)) & "%'"
'                    If Trim(UCase(Txt_ApellPat)) <> "" Then vlSql = vlSql & " AND b.gls_patben LIKE '" & Trim(UCase(Txt_ApellPat)) & "%'"
'                    If Trim(UCase(Txt_ApellMat)) <> "" Then vlSql = vlSql & " AND b.gls_matben LIKE '" & Trim(UCase(Txt_ApellMat)) & "%'"
'                    If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  c.rut_afiliado LIKE '" & Trim(Txt_Cuspp) & "%'"
'                    If Trim(Txt_SolOfe) <> "" Then vlSql = vlSql & " AND d.cod_solofe LIKE '" & Trim(Txt_SolOfe) & "%'"
'                    vlSql = vlSql & " AND c.num_cot = '" & Trim(vgRs1!num_cot) & "' "
'                    vlSql = vlSql & " AND b.num_cot = c.num_cot "
'                    vlSql = vlSql & " AND d.num_cot = c.num_cot "
'                    vlSql = vlSql & " AND d.cod_estcot = 'A' "
'                    vlSql = vlSql & " AND b.cod_par = '99' "
'                    vlSql = vlSql & " ORDER BY b.num_cot "
                    
            Call plGenerarConsulta
        End If
        Txt_SolOfe.SetFocus
    End If
    
Exit Sub
Err_KeyPress:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_Nombre_LostFocus()

    Txt_Nombre = Trim(UCase(Txt_Nombre))

End Sub

Private Sub Txt_NumIden_KeyPress(KeyAscii As Integer)
On Error GoTo Err_KeyPress

    If KeyAscii = 13 Then
        
        Txt_NumIden = UCase(Trim(Txt_NumIden))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_SolOfe) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        
        End If
        Txt_Nombre.SetFocus
    End If
    
Exit Sub
Err_KeyPress:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_SolOfe_Change()
If Not IsNumeric(Txt_SolOfe) Then
    Txt_SolOfe = ""
End If
End Sub

Private Sub Txt_SolOfe_KeyPress(KeyAscii As Integer)
On Error GoTo Err_KeyPress

    If KeyAscii = 13 Then
        
        Txt_SolOfe = UCase(Trim(Txt_SolOfe))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_SolOfe) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        
        End If
        Txt_Cuspp.SetFocus
    End If
    
Exit Sub
Err_KeyPress:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_TipoIden_KeyPress(KeyAscii As Integer)
On Error GoTo Err_KeyPress

    If KeyAscii = 13 Then
        
        Txt_TipoIden = UCase(Trim(Txt_TipoIden))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_SolOfe) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        
        End If
        Txt_Cuspp.SetFocus
    End If
    
Exit Sub
Err_KeyPress:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Sub plGenerarConsulta(Optional iTodas As Boolean)

    Txt_cotiz = Trim(Txt_cotiz)
    Txt_SolOfe = Trim(Txt_SolOfe)
    Txt_Cuspp = UCase(Trim(Txt_Cuspp))
    Txt_TipoIden = UCase(Trim(Txt_TipoIden))
    Txt_NumIden = Trim(Txt_NumIden)
    Txt_Nombre = UCase(Trim(Txt_Nombre))
    Txt_ApellPat = UCase(Trim(Txt_ApellPat))
    Txt_ApellMat = UCase(Trim(Txt_ApellMat))

'    vlCodTipCot = (vgRs1!cod_tipcot)
'    If Trim(vgRs1!cod_tipcot) = clCodTipCotOfe Then
'        vlTablaDetCotizacion = "pt_tmae_detcotizacion"
'    End If
'    If Trim(vgRs1!cod_tipcot) = clCodTipCotExt Then
'        vlTablaDetCotizacion = "pt_tmae_detcotexterna"
'    End If
'    If Trim(vgRs1!cod_tipcot) = clCodTipCotRmt Then
'        vlTablaDetCotizacion = "pt_tmae_detcotremate"
'    End If
    
    Msf_GrillaBuscaCot.rows = 1
    Fila = 1
    
    vlTablaDetCotizacion = "pt_tmae_detcotizacion"
            
    vlSql = "SELECT b.num_cot as numcot,b.gls_nomben as nomben,"
    vlSql = vlSql & "b.gls_patben as apaben,b.gls_matben as amaben,"
    vlSql = vlSql & "c.cod_cuspp as cuspp,"
    vlSql = vlSql & "i.gls_tipoidencor as tipoiden,c.num_iden as numiden, "
    vlSql = vlSql & "d.cod_estcot, d.num_operacion as codigo,c.cod_tipcot "
    vlSql = vlSql & "FROM pt_tmae_cotizacion c, pt_tmae_cotben b," & vlTablaDetCotizacion & " d "
    vlSql = vlSql & ",ma_tpar_tipoiden i "
    vlSql = vlSql & " WHERE "
    vlSql = vlSql & " c.num_cot = b.num_cot "
    vlSql = vlSql & " AND c.num_cot = d.num_cot "
    vlSql = vlSql & " AND c.cod_tipoiden = i.cod_tipoiden "
    vlSql = vlSql & " AND d.cod_estcot = 'A' "
    vlSql = vlSql & " AND b.cod_par = '99' "
    If (iTodas = False) Then
        If Txt_cotiz <> "" Then vlSql = vlSql & " AND c.num_cot LIKE '" & Txt_cotiz & "%'"
        If Txt_SolOfe <> "" Then vlSql = vlSql & " AND d.num_operacion LIKE '" & Txt_SolOfe & "%'"
        If Txt_Cuspp <> "" Then vlSql = vlSql & " AND c.cod_cuspp LIKE '" & Txt_Cuspp & "%'"
        If Txt_TipoIden <> "" Then vlSql = vlSql & " AND i.gls_tipoidencor LIKE '" & Txt_TipoIden & "%'"
        If Txt_NumIden <> "" Then vlSql = vlSql & " AND c.num_iden LIKE '" & Txt_NumIden & "%'"
        If Txt_Nombre <> "" Then vlSql = vlSql & " AND b.gls_nomben LIKE '" & Txt_Nombre & "%'"
        If Txt_ApellPat <> "" Then vlSql = vlSql & " AND b.gls_patben LIKE '" & Txt_ApellPat & "%'"
        If Txt_ApellMat <> "" Then vlSql = vlSql & " AND b.gls_matben LIKE '" & Txt_ApellMat & "%'"
    End If
    vlSql = vlSql & " ORDER BY b.num_cot "
    Set vgRs = vgConexionBD.Execute(vlSql)
    While Not vgRs.EOF
        vlNumCot = ""
        vlCodSolOfe = ""
        vlcuspp = ""
        vlTipoIdenAfi = ""
        vlNumIdenAfi = ""
        vlNomAfi = ""
        vlApellPat = ""
        vlApellMat = ""
        vlCodTipCot = (vgRs!cod_tipcot)
                
        If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
        If Not IsNull(vgRs!Codigo) Then vlCodSolOfe = Trim(vgRs!Codigo)
        If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
        If Not IsNull(vgRs!tipoiden) Then vlTipoIdenAfi = Trim(vgRs!tipoiden)
        If Not IsNull(vgRs!numiden) Then vlNumIdenAfi = Trim(vgRs!numiden)
        If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
        If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
        If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)

        Msf_GrillaBuscaCot.AddItem Fila & vbTab & _
        vlNumCot & vbTab & vlCodSolOfe & vbTab & vlcuspp & vbTab & _
        vlTipoIdenAfi & vbTab & vlNumIdenAfi & vbTab & vlNomAfi & vbTab & _
        vlApellPat & vbTab & vlApellMat & vbTab & vlCodTipCot
        Fila = Fila + 1
        
        vgRs.MoveNext
    Wend
    vgRs.Close

End Sub

