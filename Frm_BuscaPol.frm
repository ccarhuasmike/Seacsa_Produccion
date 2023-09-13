VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frm_BuscaPol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscador de Póliza"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   13170
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
      TabIndex        =   9
      Top             =   0
      Width           =   12975
      Begin VB.TextBox Txt_ApellMat 
         Height          =   285
         Left            =   10980
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Txt_ApellPat 
         Height          =   285
         Left            =   9180
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Txt_Nombre 
         Height          =   285
         Left            =   7380
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Txt_NumIden 
         Height          =   285
         Left            =   5940
         TabIndex        =   18
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_TipoIden 
         Height          =   285
         Left            =   4640
         TabIndex        =   2
         Top             =   600
         Width           =   1320
      End
      Begin VB.TextBox Txt_Cuspp 
         Height          =   285
         Left            =   3140
         TabIndex        =   20
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox Txt_NumCot 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   1595
      End
      Begin VB.TextBox Txt_cotiz 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "CUSPP"
         Height          =   195
         Index           =   8
         Left            =   3140
         TabIndex        =   21
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "Nº Ident."
         Height          =   195
         Index           =   7
         Left            =   6000
         TabIndex        =   19
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nº Cotización"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nº de Póliza"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Ident."
         Height          =   195
         Index           =   1
         Left            =   4635
         TabIndex        =   13
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   7380
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Ap. Paterno"
         Height          =   255
         Index           =   3
         Left            =   9180
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Ap. Materno"
         Height          =   255
         Index           =   4
         Left            =   10980
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   12975
      Begin VB.CommandButton Cmd_BuscarPolizas 
         Caption         =   "Todas"
         Height          =   675
         Left            =   4830
         Picture         =   "Frm_BuscaPol.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Buscar Todas las Pólizas"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   7680
         Picture         =   "Frm_BuscaPol.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   6240
         Picture         =   "Frm_BuscaPol.frx":01FC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_GrillaBuscaCot 
      Height          =   4575
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8070
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   14745599
   End
End
Attribute VB_Name = "Frm_BuscaPol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlNumPol As String
Dim vlNumCot As String
Dim vlcuspp As String
Dim vlTipoIdenAfi As String
Dim vlNumIdenAfi As String
Dim vlNomAfi As String
Dim vlApellPat As String
Dim vlApellMat As String
Dim vlNumOrd As Integer

Function flLimpiar()
On Error GoTo Err_Limpiar

    Txt_Nombre.Text = ""
    Txt_ApellPat.Text = ""
    Txt_ApellMat.Text = ""
    Txt_cotiz.Text = ""
    Txt_Cuspp.Text = ""
    Txt_cotiz.SetFocus

Exit Function
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrillaPol()
Dim vlCuenta As Integer
Dim vlColumna As Integer
Dim vlNumIdent As String
Dim vlRutAfi As String
Dim vlNombre As String
Dim vlPaterno As String
Dim vlMaterno As String

On Error GoTo Err_Carga
    Msf_GrillaBuscaCot.Clear
    Msf_GrillaBuscaCot.rows = 1
    
    Msf_GrillaBuscaCot.Enabled = True
    Msf_GrillaBuscaCot.Cols = 9
    Msf_GrillaBuscaCot.rows = 1
    Msf_GrillaBuscaCot.Row = 0
    
    Msf_GrillaBuscaCot.Col = 0
    Msf_GrillaBuscaCot.ColWidth(0) = 0
    
    Msf_GrillaBuscaCot.Col = 1
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.ColWidth(1) = 1550
    Msf_GrillaBuscaCot.Text = "Nº de Póliza"
    Msf_GrillaBuscaCot.CellFontBold = True
        
    Msf_GrillaBuscaCot.Col = 2
    Msf_GrillaBuscaCot.ColWidth(2) = 1550
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Nº Cotización"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 3
    Msf_GrillaBuscaCot.ColWidth(3) = 1550
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "CUSPP"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 4
    Msf_GrillaBuscaCot.ColWidth(4) = 1280
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Tipo Ident."
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 5
    Msf_GrillaBuscaCot.ColWidth(5) = 1470
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Nº Ident."
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 6
    Msf_GrillaBuscaCot.ColWidth(6) = 1800
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Nombre"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 7
    Msf_GrillaBuscaCot.ColWidth(7) = 1750
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.CellFontBold = True
    Msf_GrillaBuscaCot.Text = "Ap. Paterno"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 8
    Msf_GrillaBuscaCot.ColWidth(8) = 1650
    Msf_GrillaBuscaCot.CellFontBold = True
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Ap. Materno"
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

    Unload Me
    Frm_CalPoliza.Enabled = True

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
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_BuscarPolizas_Click()
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

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Me.Top = 0
    Me.Left = 0
    flCargaGrillaPol
    Txt_Nombre.Text = ""
    Txt_ApellPat.Text = ""
    Txt_ApellMat.Text = ""
    Txt_cotiz.Text = ""
    Txt_Cuspp.Text = ""
    
Exit Sub
Err_Cargar:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Frm_CalPoliza.Enabled = True
End Sub

Private Sub Msf_GrillaBuscaCot_DblClick()
On Error GoTo Err_Seleccionar

    Msf_GrillaBuscaCot.Col = 1
    Msf_GrillaBuscaCot.Row = Msf_GrillaBuscaCot.RowSel
    
    If (Not (Msf_GrillaBuscaCot.Text = "") And (Msf_GrillaBuscaCot.Row <> 0)) Then
        vlNumPol = Msf_GrillaBuscaCot.Text
        Call Frm_CalPoliza.flBuscaPoliza(vlNumPol)
        Unload Me
    Else
       MsgBox "No existen Pólizas Para Modificar", vbInformation, "Información "
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
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_ApellMat = UCase(Trim(Txt_ApellMat))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_NumCot) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        End If
        Cmd_BuscarPolizas.SetFocus
    End If
 
Exit Sub
Err_Ingreso:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_ApellMat_LostFocus()
    Txt_ApellMat = Trim(UCase(Txt_ApellMat))
End Sub

Private Sub Txt_ApellPat_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_ApellPat = UCase(Trim(Txt_ApellPat))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_NumCot) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        End If
        Txt_ApellMat.SetFocus
    End If
 
Exit Sub
Err_Ingreso:
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
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_Cuspp = UCase(Trim(Txt_Cuspp))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_NumCot) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        End If
        Txt_TipoIden.SetFocus
    End If
 
Exit Sub
Err_Ingreso:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_nombre_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_Nombre = UCase(Trim(Txt_Nombre))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_NumCot) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        End If
        Txt_ApellPat.SetFocus
    End If
 
Exit Sub
Err_Ingreso:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_Cotiz_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Change

    If KeyAscii = 13 Then
        Txt_cotiz = UCase(Trim(Txt_cotiz))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_NumCot) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
'        vlSql = ""
'        vlSql = "SELECT b.num_poliza as numpol,"
'        vlSql = vlSql & "b.num_orden as numord,"
'        vlSql = vlSql & "b.gls_nomben as nomben,"
'        vlSql = vlSql & "b.gls_patben as apaben,"
'        vlSql = vlSql & "b.gls_matben as amaben,"
'        vlSql = vlSql & "p.rut_afi as rutafi,"
'        vlSql = vlSql & "p.dgv_afi as dgvafi,"
'        vlSql = vlSql & "p.num_cot as numcot "
'        vlSql = vlSql & " FROM pd_tmae_oripoliza p, pd_tmae_oripolben b"
'        vlSql = vlSql & " WHERE b.num_poliza LIKE '" & Trim(Txt_cotiz) & "%'"
'
'        If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  p.rut_afi LIKE '" & Trim(Txt_Cuspp) & "%'"
'        If Trim(UCase(Txt_Nombre)) <> "" Then vlSql = vlSql & " AND  b.gls_nomben LIKE '" & Trim(UCase(Txt_Nombre)) & "%'"
'        If Trim(UCase(Txt_ApellPat)) <> "" Then vlSql = vlSql & " AND b.gls_patben LIKE '" & Trim(UCase(Txt_ApellPat)) & "%'"
'        If Trim(UCase(Txt_ApellMat)) <> "" Then vlSql = vlSql & " AND b.gls_matben LIKE '" & Trim(UCase(Txt_ApellMat)) & "%'"
'        If Trim(Txt_NumCot) <> "" Then vlSql = vlSql & " AND  p.num_cot LIKE '" & Trim(Txt_NumCot) & "%'"
'
'        vlSql = vlSql & " AND b.num_poliza = p.num_poliza"
'        vlSql = vlSql & " AND b.cod_par = '99'"
'        vlSql = vlSql & " ORDER BY b.num_poliza "
'        Set vgRs = vgConexionBD.Execute(vlSql)
'        Fila = 1
'        Msf_GrillaBuscaCot.Rows = 1
'        While Not vgRs.EOF
'
'            vlNumPol = ""
'            vlNumCot = ""
'            vlRutAfi = ""
'            vlNomAfi = ""
'            vlApellPat = ""
'            vlApellMat = ""
'            vlDig = ""
'            vlNumOrd = 0
'            If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
'            If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
'            If Not IsNull(vgRs!rutafi) Then vlRutAfi = Trim(vgRs!rutafi)
'            If Not IsNull(vgRs!dgvafi) Then vlDig = Trim(vgRs!dgvafi)
'            If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
'            If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
'            If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)
'            If Not IsNull(vgRs!numord) Then vlNumOrd = Trim(vgRs!numord)
'            Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & vlNumCot & vbTab & _
'                                vbTab & (vlRutAfi + "- " + vlDig) & vbTab & vlNomAfi & _
'                                vbTab & vlApellPat & vbTab & vlApellMat
'            Fila = Fila + 1
'            vgRs.MoveNext
'        Wend
'    vgRs.Close
'    Else
'
'         Call flCargaGrillaPol
'
        
            Call plGenerarConsulta
        End If
        Txt_NumCot.SetFocus
    End If
 
Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_Nombre_LostFocus()

    Txt_Nombre = Trim(UCase(Txt_Nombre))

End Sub

Private Sub Txt_NumCot_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_NumCot = UCase(Trim(Txt_NumCot))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_NumCot) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        End If
        Txt_Cuspp.SetFocus
    End If
 
Exit Sub
Err_Ingreso:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Sub plGenerarConsulta(Optional iTodas As Boolean)

    Txt_cotiz = Trim(Txt_cotiz)
    Txt_NumCot = Trim(Txt_NumCot)
    Txt_Cuspp = UCase(Trim(Txt_Cuspp))
    Txt_TipoIden = UCase(Trim(Txt_TipoIden))
    Txt_NumIden = Trim(Txt_NumIden)
    Txt_Nombre = UCase(Trim(Txt_Nombre))
    Txt_ApellPat = UCase(Trim(Txt_ApellPat))
    Txt_ApellMat = UCase(Trim(Txt_ApellMat))

    Msf_GrillaBuscaCot.rows = 1

    vlSql = "SELECT b.num_poliza as numpol,b.num_orden as numord,"
    vlSql = vlSql & "p.num_cot as numcot,b.gls_nomben as nomben,"
    vlSql = vlSql & "b.gls_patben as apaben,b.gls_matben as amaben,"
    vlSql = vlSql & "p.cod_cuspp as cuspp,"
    vlSql = vlSql & "i.gls_tipoidencor as tipoiden,p.num_idenafi as numiden "
    vlSql = vlSql & "FROM pd_tmae_oripoliza p, pd_tmae_oripolben b "
    vlSql = vlSql & ",ma_tpar_tipoiden i "
    vlSql = vlSql & " WHERE "
    vlSql = vlSql & " p.num_poliza = b.num_poliza "
    vlSql = vlSql & " AND p.cod_tipoidenafi = i.cod_tipoiden "
    vlSql = vlSql & " AND b.cod_par = '99' "
    If (iTodas = False) Then
        If Txt_cotiz <> "" Then vlSql = vlSql & " AND p.num_poliza LIKE '%" & Txt_cotiz & "%'"
        If Txt_NumCot <> "" Then vlSql = vlSql & " AND p.num_cot LIKE '" & Txt_NumCot & "%'"
        If Txt_Cuspp <> "" Then vlSql = vlSql & " AND p.cod_cuspp LIKE '" & Txt_Cuspp & "%'"
        If Txt_TipoIden <> "" Then vlSql = vlSql & " AND i.gls_tipoidencor LIKE '" & Txt_TipoIden & "%'"
        If Txt_NumIden <> "" Then vlSql = vlSql & " AND p.num_idenafi LIKE '" & Txt_NumIden & "%'"
        If Txt_Nombre <> "" Then vlSql = vlSql & " AND b.gls_nomben LIKE '" & Txt_Nombre & "%'"
        If Txt_ApellPat <> "" Then vlSql = vlSql & " AND b.gls_patben LIKE '" & Txt_ApellPat & "%'"
        If Txt_ApellMat <> "" Then vlSql = vlSql & " AND b.gls_matben LIKE '" & Txt_ApellMat & "%'"
    End If
    vlSql = vlSql & " ORDER BY p.num_poliza "
    Set vgRs = vgConexionBD.Execute(vlSql)
    While Not vgRs.EOF
    
        vlNumPol = ""
        vlNumCot = ""
        vlcuspp = ""
        vlTipoIdenAfi = ""
        vlNumIdenAfi = ""
        vlNomAfi = ""
        vlApellPat = ""
        vlApellMat = ""
        vlNumOrd = 0
                
        If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
        If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
        If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
        If Not IsNull(vgRs!tipoiden) Then vlTipoIdenAfi = Trim(vgRs!tipoiden)
        If Not IsNull(vgRs!numiden) Then vlNumIdenAfi = Trim(vgRs!numiden)
        If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
        If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
        If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)

        Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & _
        vlNumCot & vbTab & vlcuspp & vbTab & _
        vlTipoIdenAfi & vbTab & vlNumIdenAfi & vbTab & vlNomAfi & vbTab & _
        vlApellPat & vbTab & vlApellMat
        
        Fila = Fila + 1
        
        vgRs.MoveNext
    Wend
    vgRs.Close

End Sub

Private Sub Txt_NumIden_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_NumIden = UCase(Trim(Txt_NumIden))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_NumCot) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        End If
        Txt_Nombre.SetFocus
    End If
 
Exit Sub
Err_Ingreso:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_TipoIden_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_TipoIden = UCase(Trim(Txt_TipoIden))
        
        If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_NumCot) <> "") Or _
        (Trim(Txt_Cuspp) <> "") Or (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        End If
        Txt_NumIden.SetFocus
    End If
 
Exit Sub
Err_Ingreso:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub
