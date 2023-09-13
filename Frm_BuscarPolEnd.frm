VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_BuscarPolEnd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11910
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   5040
      Width           =   11655
      Begin VB.CommandButton Btn_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4440
         Picture         =   "Frm_BuscarPolEnd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5880
         Picture         =   "Frm_BuscarPolEnd.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   8
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
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   11685
      Begin VB.TextBox Txt_ApellMat 
         Height          =   285
         Left            =   10095
         TabIndex        =   6
         Top             =   600
         Width           =   1530
      End
      Begin VB.TextBox Txt_ApellPat 
         Height          =   285
         Left            =   8520
         TabIndex        =   5
         Top             =   600
         Width           =   1600
      End
      Begin VB.TextBox Txt_Nombre 
         Height          =   285
         Left            =   7080
         TabIndex        =   4
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox Txt_NumIden 
         Height          =   285
         Left            =   5760
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Txt_TipoIden 
         Height          =   285
         Left            =   4680
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Txt_Cuspp 
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Txt_NumCot 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Txt_Endoso 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Txt_cotiz 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "CUSPP"
         Height          =   195
         Index           =   8
         Left            =   3360
         TabIndex        =   22
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "Nº Ident."
         Height          =   195
         Index           =   5
         Left            =   5760
         TabIndex        =   20
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nº Cotización"
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "End"
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Ap. Materno"
         Height          =   255
         Index           =   4
         Left            =   10095
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Ap. Paterno"
         Height          =   255
         Index           =   3
         Left            =   8520
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   7080
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Ident."
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   11
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nº de Póliza"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_GrillaBuscaCot 
      Height          =   3855
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6800
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   14745599
      Enabled         =   -1  'True
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "Frm_BuscarPolEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vlNomAfi As String
Dim vlApellPat As String
Dim vlApellMat As String
Dim vlRutAfi As String
Dim vlNumPol As String
Dim vlDig As String
Dim vlNumOrd As Integer
Dim vlEndoso As Integer
Dim vlNumCot As String
Dim vlcuspp As String
Dim vlRut As String

Dim vlSql As String

'FUNCION QUE GUARDA EL NOMBRE DEL FORMULARIO DEL QUE SE LLAMO A LA FUNCION
Function flInicio(iNomForm)
    vgNomForm = iNomForm
    Call Form_Load
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
    Msf_GrillaBuscaCot.Rows = 1
    
    Msf_GrillaBuscaCot.Enabled = True
    Msf_GrillaBuscaCot.Cols = 10
    Msf_GrillaBuscaCot.Rows = 1
    Msf_GrillaBuscaCot.Row = 0
    
    Msf_GrillaBuscaCot.Col = 0
    Msf_GrillaBuscaCot.ColWidth(0) = 0
    
    Msf_GrillaBuscaCot.Col = 1
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.ColWidth(1) = 1400
    Msf_GrillaBuscaCot.Text = "Nº de Póliza"
    Msf_GrillaBuscaCot.CellFontBold = True
        
    Msf_GrillaBuscaCot.Col = 2
    Msf_GrillaBuscaCot.ColWidth(2) = 400
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "End"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 3
    Msf_GrillaBuscaCot.ColWidth(3) = 1500
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Nº Cotización"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 4
    Msf_GrillaBuscaCot.ColWidth(4) = 1300
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "CUSPP"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 5
    Msf_GrillaBuscaCot.ColWidth(5) = 1100
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.Text = "Tipo Ident."
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 6
    Msf_GrillaBuscaCot.ColWidth(6) = 1400
    Msf_GrillaBuscaCot.CellAlignment = 5
    Msf_GrillaBuscaCot.Text = "Nº Ident."
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 7
    Msf_GrillaBuscaCot.ColWidth(7) = 1400
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.CellFontBold = True
    Msf_GrillaBuscaCot.Text = "Nombre"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 8
    Msf_GrillaBuscaCot.ColWidth(8) = 1600
    Msf_GrillaBuscaCot.CellAlignment = 4
    Msf_GrillaBuscaCot.CellFontBold = True
    Msf_GrillaBuscaCot.Text = "Ap. Paterno"
    Msf_GrillaBuscaCot.CellFontBold = True
    
    Msf_GrillaBuscaCot.Col = 9
    Msf_GrillaBuscaCot.ColWidth(9) = 1600
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
      
Private Sub Btn_Limpiar_Click()
On Error GoTo Err_Limpiar
    
    Txt_cotiz.Text = ""
    Txt_Endoso.Text = ""
    Txt_NumCot.Text = ""
    Txt_Cuspp.Text = ""
    Txt_TipoIden.Text = ""
    Txt_NumIden.Text = ""
    Txt_Nombre.Text = ""
    Txt_ApellPat.Text = ""
    Txt_ApellMat.Text = ""
    Txt_cotiz.SetFocus
   
Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Btn_Salir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
On Error GoTo Err_Cargar

    Me.Top = 0
    Me.Left = 0
    
    flCargaGrillaPol
    Txt_cotiz.Text = ""
    Txt_Endoso.Text = ""
    Txt_NumCot.Text = ""
    Txt_Cuspp.Text = ""
    Txt_TipoIden.Text = ""
    Txt_NumIden.Text = ""
    Txt_Nombre.Text = ""
    Txt_ApellPat.Text = ""
    Txt_ApellMat.Text = ""
    
Exit Sub
Err_Cargar:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Msf_GrillaBuscaCot_DblClick()
On Error GoTo Err_Seleccionar

    Msf_GrillaBuscaCot.Col = 1
    Msf_GrillaBuscaCot.Row = Msf_GrillaBuscaCot.RowSel
    If (Not (Msf_GrillaBuscaCot.Text = "") And (Msf_GrillaBuscaCot.Row <> 0)) Then
        vlNumPol = Msf_GrillaBuscaCot.Text
        Msf_GrillaBuscaCot.Col = 2
        vlEndoso = Msf_GrillaBuscaCot.Text
        Call flAsignaFormulario(vgNomForm, vlNumPol, vlEndoso)
        'Call Frm_CalConsulta.flBuscaPoliza(vlNumPol, vlEndoso)
        Unload Me
    Else
       MsgBox "No existen Pólizas Para Mostrar", vbInformation, "Información "
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
On Error GoTo Err_Change

If KeyAscii = 13 Then
    
    If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_Endoso) <> "") Or (Trim(Txt_NumCot) <> "") Or _
       (Trim(Txt_Cuspp) <> "") Or (Trim(Txt_TipoIden) <> "") Or (Trim(Txt_NumIden) <> "") Or _
       (Trim(Txt_Nombre) <> "") Or (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
              
        Txt_ApellMat = UCase(Trim(Txt_ApellMat))
        
        vlSql = ""
        vlSql = "SELECT p.num_poliza as numpol,"
        vlSql = vlSql & "b.num_orden as numord,"
        vlSql = vlSql & "b.gls_nomben as nomben,"
        vlSql = vlSql & "b.gls_patben as apaben,"
        vlSql = vlSql & "b.gls_matben as amaben,"
        vlSql = vlSql & "i.gls_tipoidencor as rutafi,"
        vlSql = vlSql & "b.num_idenben as dgvafi,"
        vlSql = vlSql & "p.num_cot as numcot,"
        vlSql = vlSql & "p.cod_cuspp as cuspp,"
        vlSql = vlSql & "p.num_endoso as numend "
        vlSql = vlSql & " FROM pd_tmae_poliza p, pd_tmae_polben b, ma_tpar_tipoiden i "
        vlSql = vlSql & " WHERE "
        
        vlSql = vlSql & " b.num_poliza = p.num_poliza"
        vlSql = vlSql & " AND b.num_endoso = p.num_endoso"
        vlSql = vlSql & " AND b.cod_par = '99'"
        vlSql = vlSql & " AND b.cod_tipoidenben = i.cod_tipoiden "
        
        If Trim(Txt_cotiz) <> "" Then vlSql = vlSql & " AND p.num_poliza like '%" & Trim(Txt_cotiz) & "%'"
        If Trim(Txt_Endoso) <> "" Then vlSql = vlSql & " AND p.num_endoso like '" & Trim(Txt_Endoso) & "%'"
        If Trim(Txt_NumCot) <> "" Then vlSql = vlSql & " AND p.num_cot like '" & Trim(Txt_NumCot) & "%'"
        If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  p.cod_cuspp like '" & Trim(Txt_Cuspp) & "%'"
        If Trim(Txt_TipoIden) <> "" Then vlSql = vlSql & " AND i.gls_cod_tipoidencor like '" & Trim(Txt_TipoIden) & "%'"
        If Trim(Txt_NumIden) <> "" Then vlSql = vlSql & " AND  b.num_idenben like '" & Trim(Txt_NumIden) & "%'"
        If Trim(Txt_Nombre) <> "" Then vlSql = vlSql & " AND  b.gls_nomben like '" & Trim(Txt_Nombre) & "%'"
        If Trim(Txt_ApellPat) <> "" Then vlSql = vlSql & " AND b.gls_patben like '" & Trim(Txt_ApellPat) & "%'"
        If Trim(Txt_ApellMat) <> "" Then vlSql = vlSql & " AND b.gls_matben like '" & Trim(Txt_ApellMat) & "%'"

        vlSql = vlSql & " ORDER BY b.num_poliza "
                      
        Set vgRs = vgConexionBD.Execute(vlSql)
        fila = 1
        Msf_GrillaBuscaCot.Rows = 1
        While Not vgRs.EOF
         
            vlNumPol = ""
            vlNumCot = ""
            vlRutAfi = ""
            vlcuspp = ""
            vlNomAfi = ""
            vlApellPat = ""
            vlApellMat = ""
            vlDig = ""
            vlNumOrd = 0
            vlEndoso = 0
                      
            If Not IsNull(vgRs!numord) Then vlNumOrd = Trim(vgRs!numord)
            If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
            If Not IsNull(vgRs!numend) Then vlEndoso = Trim(vgRs!numend)
            If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
            If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
            If Not IsNull(vgRs!rutafi) Then vlRutAfi = Trim(vgRs!rutafi)
            If Not IsNull(vgRs!dgvafi) Then vlDig = Trim(vgRs!dgvafi)
            If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
            If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
            If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)

            Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & vlEndoso & vbTab & _
                                vlNumCot & vbTab & vlcuspp & vbTab & vlRutAfi & vbTab & vlDig & vbTab & vlNomAfi & _
                                vbTab & vlApellPat & vbTab & vlApellMat
            fila = fila + 1
            vgRs.MoveNext
        Wend
    vgRs.Close
    Else
            
            Call flCargaGrillaPol
            
    End If
    Txt_ApellMat.SetFocus
End If

Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub
 
Private Sub Txt_ApellPat_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Change

If KeyAscii = 13 Then
    
    If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_Endoso) <> "") Or (Trim(Txt_NumCot) <> "") Or _
       (Trim(Txt_Cuspp) <> "") Or (Trim(Txt_TipoIden) <> "") Or (Trim(Txt_NumIden) <> "") Or _
       (Trim(Txt_Nombre) <> "") Or (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
              
        Txt_ApellPat = UCase(Trim(Txt_ApellPat))
        
        vlSql = ""
        vlSql = "SELECT p.num_poliza as numpol,"
        vlSql = vlSql & "b.num_orden as numord,"
        vlSql = vlSql & "b.gls_nomben as nomben,"
        vlSql = vlSql & "b.gls_patben as apaben,"
        vlSql = vlSql & "b.gls_matben as amaben,"
        vlSql = vlSql & "i.gls_tipoidencor as rutafi,"
        vlSql = vlSql & "b.num_idenben as dgvafi,"
        vlSql = vlSql & "p.num_cot as numcot,"
        vlSql = vlSql & "p.cod_cuspp as cuspp,"
        vlSql = vlSql & "p.num_endoso as numend "
        vlSql = vlSql & " FROM pd_tmae_poliza p, pd_tmae_polben b, ma_tpar_tipoiden i "
        vlSql = vlSql & " WHERE "
        
        vlSql = vlSql & " b.num_poliza = p.num_poliza"
        vlSql = vlSql & " AND b.num_endoso = p.num_endoso"
        vlSql = vlSql & " AND b.cod_par = '99'"
        vlSql = vlSql & " AND b.cod_tipoidenben = i.cod_tipoiden "
        
        If Trim(Txt_cotiz) <> "" Then vlSql = vlSql & " AND p.num_poliza like '%" & Trim(Txt_cotiz) & "%'"
        If Trim(Txt_Endoso) <> "" Then vlSql = vlSql & " AND p.num_endoso like '" & Trim(Txt_Endoso) & "%'"
        If Trim(Txt_NumCot) <> "" Then vlSql = vlSql & " AND p.num_cot like '" & Trim(Txt_NumCot) & "%'"
        If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  p.cod_cuspp like '" & Trim(Txt_Cuspp) & "%'"
        If Trim(Txt_TipoIden) <> "" Then vlSql = vlSql & " AND i.gls_cod_tipoidencor like '" & Trim(Txt_TipoIden) & "%'"
        If Trim(Txt_NumIden) <> "" Then vlSql = vlSql & " AND  b.num_idenben like '" & Trim(Txt_NumIden) & "%'"
        If Trim(Txt_Nombre) <> "" Then vlSql = vlSql & " AND  b.gls_nomben like '" & Trim(Txt_Nombre) & "%'"
        If Trim(Txt_ApellPat) <> "" Then vlSql = vlSql & " AND b.gls_patben like '" & Trim(Txt_ApellPat) & "%'"
        If Trim(Txt_ApellMat) <> "" Then vlSql = vlSql & " AND b.gls_matben like '" & Trim(Txt_ApellMat) & "%'"

        vlSql = vlSql & " ORDER BY b.num_poliza "
                      
        Set vgRs = vgConexionBD.Execute(vlSql)
        fila = 1
        Msf_GrillaBuscaCot.Rows = 1
        While Not vgRs.EOF
         
            vlNumPol = ""
            vlNumCot = ""
            vlRutAfi = ""
            vlcuspp = ""
            vlNomAfi = ""
            vlApellPat = ""
            vlApellMat = ""
            vlDig = ""
            vlNumOrd = 0
            vlEndoso = 0
                      
            If Not IsNull(vgRs!numord) Then vlNumOrd = Trim(vgRs!numord)
            If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
            If Not IsNull(vgRs!numend) Then vlEndoso = Trim(vgRs!numend)
            If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
            If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
            If Not IsNull(vgRs!rutafi) Then vlRutAfi = Trim(vgRs!rutafi)
            If Not IsNull(vgRs!dgvafi) Then vlDig = Trim(vgRs!dgvafi)
            If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
            If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
            If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)

            Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & vlEndoso & vbTab & _
                                vlNumCot & vbTab & vlcuspp & vbTab & vlRutAfi & vbTab & vlDig & vbTab & vlNomAfi & _
                                vbTab & vlApellPat & vbTab & vlApellMat
            fila = fila + 1
            vgRs.MoveNext
        Wend
    vgRs.Close
    Else
            
            Call flCargaGrillaPol
            
    End If
    Txt_ApellMat.SetFocus
End If

Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_Cotiz_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Change

If KeyAscii = 13 Then
    
    If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_Endoso) <> "") Or (Trim(Txt_NumCot) <> "") Or _
       (Trim(Txt_Cuspp) <> "") Or (Trim(Txt_TipoIden) <> "") Or (Trim(Txt_NumIden) <> "") Or _
       (Trim(Txt_Nombre) <> "") Or (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
              
'        Txt_cotiz = Format(Trim(Txt_cotiz), "0000000000")
        Txt_cotiz = UCase(Trim(Txt_cotiz))
        
        vlSql = ""
        vlSql = "SELECT p.num_poliza as numpol,"
        vlSql = vlSql & "b.num_orden as numord,"
        vlSql = vlSql & "b.gls_nomben as nomben,"
        vlSql = vlSql & "b.gls_patben as apaben,"
        vlSql = vlSql & "b.gls_matben as amaben,"
        vlSql = vlSql & "i.gls_tipoidencor as rutafi,"
        vlSql = vlSql & "b.num_idenben as dgvafi,"
        vlSql = vlSql & "p.num_cot as numcot,"
        vlSql = vlSql & "p.cod_cuspp as cuspp,"
        vlSql = vlSql & "p.num_endoso as numend "
        vlSql = vlSql & " FROM pd_tmae_poliza p, pd_tmae_polben b, ma_tpar_tipoiden i "
        vlSql = vlSql & " WHERE "
        
        vlSql = vlSql & " b.num_poliza = p.num_poliza"
        vlSql = vlSql & " AND b.num_endoso = p.num_endoso"
        vlSql = vlSql & " AND b.cod_par = '99'"
        vlSql = vlSql & " AND b.cod_tipoidenben = i.cod_tipoiden "
        
        If Trim(Txt_cotiz) <> "" Then vlSql = vlSql & " AND p.num_poliza like '%" & Trim(Txt_cotiz) & "%'"
        If Trim(Txt_Endoso) <> "" Then vlSql = vlSql & " AND p.num_endoso like '" & Trim(Txt_Endoso) & "%'"
        If Trim(Txt_NumCot) <> "" Then vlSql = vlSql & " AND p.num_cot like '" & Trim(Txt_NumCot) & "%'"
        If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  p.cod_cuspp like '" & Trim(Txt_Cuspp) & "%'"
        If Trim(Txt_TipoIden) <> "" Then vlSql = vlSql & " AND i.gls_cod_tipoidencor like '" & Trim(Txt_TipoIden) & "%'"
        If Trim(Txt_NumIden) <> "" Then vlSql = vlSql & " AND  b.num_idenben like '" & Trim(Txt_NumIden) & "%'"
        If Trim(Txt_Nombre) <> "" Then vlSql = vlSql & " AND  b.gls_nomben like '" & Trim(Txt_Nombre) & "%'"
        If Trim(Txt_ApellPat) <> "" Then vlSql = vlSql & " AND b.gls_patben like '" & Trim(Txt_ApellPat) & "%'"
        If Trim(Txt_ApellMat) <> "" Then vlSql = vlSql & " AND b.gls_matben like '" & Trim(Txt_ApellMat) & "%'"

        vlSql = vlSql & " ORDER BY b.num_poliza "
                      
        Set vgRs = vgConexionBD.Execute(vlSql)
        fila = 1
        Msf_GrillaBuscaCot.Rows = 1
        While Not vgRs.EOF
         
            vlNumPol = ""
            vlNumCot = ""
            vlRutAfi = ""
            vlcuspp = ""
            vlNomAfi = ""
            vlApellPat = ""
            vlApellMat = ""
            vlDig = ""
            vlNumOrd = 0
            vlEndoso = 0
                      
            If Not IsNull(vgRs!numord) Then vlNumOrd = Trim(vgRs!numord)
            If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
            If Not IsNull(vgRs!numend) Then vlEndoso = Trim(vgRs!numend)
            If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
            If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
            If Not IsNull(vgRs!rutafi) Then vlRutAfi = Trim(vgRs!rutafi)
            If Not IsNull(vgRs!dgvafi) Then vlDig = Trim(vgRs!dgvafi)
            If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
            If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
            If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)

            Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & vlEndoso & vbTab & _
                                vlNumCot & vbTab & vlcuspp & vbTab & vlRutAfi & vbTab & vlDig & vbTab & vlNomAfi & _
                                vbTab & vlApellPat & vbTab & vlApellMat
            fila = fila + 1
            vgRs.MoveNext
        Wend
    vgRs.Close
    Else
            
            Call flCargaGrillaPol
            
    End If
    Txt_Endoso.SetFocus
End If

Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_Cuspp_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Change

If KeyAscii = 13 Then
    
   If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_Endoso) <> "") Or (Trim(Txt_NumCot) <> "") Or _
       (Trim(Txt_Cuspp) <> "") Or (Trim(Txt_TipoIden) <> "") Or (Trim(Txt_NumIden) <> "") Or _
       (Trim(Txt_Nombre) <> "") Or (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
              
        Txt_Cuspp = UCase(Trim(Txt_Cuspp))
        
        vlSql = ""
        vlSql = "SELECT p.num_poliza as numpol,"
        vlSql = vlSql & "b.num_orden as numord,"
        vlSql = vlSql & "b.gls_nomben as nomben,"
        vlSql = vlSql & "b.gls_patben as apaben,"
        vlSql = vlSql & "b.gls_matben as amaben,"
        vlSql = vlSql & "i.gls_tipoidencor as rutafi,"
        vlSql = vlSql & "b.num_idenben as dgvafi,"
        vlSql = vlSql & "p.num_cot as numcot,"
        vlSql = vlSql & "p.cod_cuspp as cuspp,"
        vlSql = vlSql & "p.num_endoso as numend "
        vlSql = vlSql & " FROM pd_tmae_poliza p, pd_tmae_polben b, ma_tpar_tipoiden i "
        vlSql = vlSql & " WHERE "
        
        vlSql = vlSql & " b.num_poliza = p.num_poliza"
        vlSql = vlSql & " AND b.num_endoso = p.num_endoso"
        vlSql = vlSql & " AND b.cod_par = '99'"
        vlSql = vlSql & " AND b.cod_tipoidenben = i.cod_tipoiden "
        
        If Trim(Txt_cotiz) <> "" Then vlSql = vlSql & " AND p.num_poliza like '%" & Trim(Txt_cotiz) & "%'"
        If Trim(Txt_Endoso) <> "" Then vlSql = vlSql & " AND p.num_endoso like '" & Trim(Txt_Endoso) & "%'"
        If Trim(Txt_NumCot) <> "" Then vlSql = vlSql & " AND p.num_cot like '" & Trim(Txt_NumCot) & "%'"
        If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  p.cod_cuspp like '" & Trim(Txt_Cuspp) & "%'"
        If Trim(Txt_TipoIden) <> "" Then vlSql = vlSql & " AND i.gls_cod_tipoidencor like '" & Trim(Txt_TipoIden) & "%'"
        If Trim(Txt_NumIden) <> "" Then vlSql = vlSql & " AND  b.num_idenben like '" & Trim(Txt_NumIden) & "%'"
        If Trim(Txt_Nombre) <> "" Then vlSql = vlSql & " AND  b.gls_nomben like '" & Trim(Txt_Nombre) & "%'"
        If Trim(Txt_ApellPat) <> "" Then vlSql = vlSql & " AND b.gls_patben like '" & Trim(Txt_ApellPat) & "%'"
        If Trim(Txt_ApellMat) <> "" Then vlSql = vlSql & " AND b.gls_matben like '" & Trim(Txt_ApellMat) & "%'"

        vlSql = vlSql & " ORDER BY b.num_poliza "
                      
        Set vgRs = vgConexionBD.Execute(vlSql)
        fila = 1
        Msf_GrillaBuscaCot.Rows = 1
        While Not vgRs.EOF
         
            vlNumPol = ""
            vlNumCot = ""
            vlRutAfi = ""
            vlcuspp = ""
            vlNomAfi = ""
            vlApellPat = ""
            vlApellMat = ""
            vlDig = ""
            vlNumOrd = 0
            vlEndoso = 0
                      
            If Not IsNull(vgRs!numord) Then vlNumOrd = Trim(vgRs!numord)
            If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
            If Not IsNull(vgRs!numend) Then vlEndoso = Trim(vgRs!numend)
            If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
            If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
            If Not IsNull(vgRs!rutafi) Then vlRutAfi = Trim(vgRs!rutafi)
            If Not IsNull(vgRs!dgvafi) Then vlDig = Trim(vgRs!dgvafi)
            If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
            If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
            If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)

            Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & vlEndoso & vbTab & _
                                vlNumCot & vbTab & vlcuspp & vbTab & vlRutAfi & vbTab & vlDig & vbTab & vlNomAfi & _
                                vbTab & vlApellPat & vbTab & vlApellMat
            fila = fila + 1
            vgRs.MoveNext
        Wend
    vgRs.Close
    Else
            
            Call flCargaGrillaPol
            
    End If
    Txt_TipoIden.SetFocus
End If

Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_Endoso_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Change
If KeyAscii = 13 Then
   
    If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_Endoso) <> "") Or (Trim(Txt_NumCot) <> "") Or _
       (Trim(Txt_Cuspp) <> "") Or (Trim(Txt_TipoIden) <> "") Or (Trim(Txt_NumIden) <> "") Or _
       (Trim(Txt_Nombre) <> "") Or (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
                       
        Txt_Endoso = UCase(Trim(Txt_Endoso))
        
        vlSql = ""
        vlSql = "SELECT p.num_poliza as numpol,"
        vlSql = vlSql & "b.num_orden as numord,"
        vlSql = vlSql & "b.gls_nomben as nomben,"
        vlSql = vlSql & "b.gls_patben as apaben,"
        vlSql = vlSql & "b.gls_matben as amaben,"
        vlSql = vlSql & "i.gls_tipoidencor as rutafi,"
        vlSql = vlSql & "b.num_idenben as dgvafi,"
        vlSql = vlSql & "p.num_cot as numcot,"
        vlSql = vlSql & "p.cod_cuspp as cuspp,"
        vlSql = vlSql & "p.num_endoso as numend "
        vlSql = vlSql & " FROM pd_tmae_poliza p, pd_tmae_polben b, ma_tpar_tipoiden i "
        vlSql = vlSql & " WHERE "
        
        vlSql = vlSql & " b.num_poliza = p.num_poliza"
        vlSql = vlSql & " AND b.num_endoso = p.num_endoso"
        vlSql = vlSql & " AND b.cod_par = '99'"
        vlSql = vlSql & " AND b.cod_tipoidenben = i.cod_tipoiden "
            
        If Trim(Txt_cotiz) <> "" Then vlSql = vlSql & " AND p.num_poliza like '%" & Trim(Txt_cotiz) & "%'"
        If Trim(Txt_Endoso) <> "" Then vlSql = vlSql & " AND p.num_endoso like '" & Trim(Txt_Endoso) & "%'"
        If Trim(Txt_NumCot) <> "" Then vlSql = vlSql & " AND p.num_cot like '" & Trim(Txt_NumCot) & "%'"
        If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  p.cod_cuspp like '" & Trim(Txt_Cuspp) & "%'"
        If Trim(Txt_TipoIden) <> "" Then vlSql = vlSql & " AND i.gls_cod_tipoidencor like '" & Trim(Txt_TipoIden) & "%'"
        If Trim(Txt_NumIden) <> "" Then vlSql = vlSql & " AND  b.num_idenben like '" & Trim(Txt_NumIden) & "%'"
        If Trim(Txt_Nombre) <> "" Then vlSql = vlSql & " AND  b.gls_nomben like '" & Trim(Txt_Nombre) & "%'"
        If Trim(Txt_ApellPat) <> "" Then vlSql = vlSql & " AND b.gls_patben like '" & Trim(Txt_ApellPat) & "%'"
        If Trim(Txt_ApellMat) <> "" Then vlSql = vlSql & " AND b.gls_matben like '" & Trim(Txt_ApellMat) & "%'"
     
        vlSql = vlSql & " ORDER BY b.num_poliza "
                      
        Set vgRs = vgConexionBD.Execute(vlSql)
        fila = 1
        Msf_GrillaBuscaCot.Rows = 1
        While Not vgRs.EOF
         
            vlNumPol = ""
            vlNumCot = ""
            vlRutAfi = ""
            vlcuspp = ""
            vlNomAfi = ""
            vlApellPat = ""
            vlApellMat = ""
            vlDig = ""
            vlNumOrd = 0
            vlEndoso = 0
                      
            If Not IsNull(vgRs!numord) Then vlNumOrd = Trim(vgRs!numord)
            If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
            If Not IsNull(vgRs!numend) Then vlEndoso = Trim(vgRs!numend)
            If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
            If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
            If Not IsNull(vgRs!rutafi) Then vlRutAfi = Trim(vgRs!rutafi)
            If Not IsNull(vgRs!dgvafi) Then vlDig = Trim(vgRs!dgvafi)
            If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
            If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
            If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)

            Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & vlEndoso & vbTab & _
                                vlNumCot & vbTab & vlcuspp & vbTab & vlRutAfi & vbTab & vlDig & vbTab & vlNomAfi & _
                                vbTab & vlApellPat & vbTab & vlApellMat
            fila = fila + 1
            vgRs.MoveNext
        Wend
    vgRs.Close
    Else
            
            Call flCargaGrillaPol
            
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

Private Sub Txt_nombre_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Change

If KeyAscii = 13 Then
    
    If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_Endoso) <> "") Or (Trim(Txt_NumCot) <> "") Or _
       (Trim(Txt_Cuspp) <> "") Or (Trim(Txt_TipoIden) <> "") Or (Trim(Txt_NumIden) <> "") Or _
       (Trim(Txt_Nombre) <> "") Or (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
              
        Txt_Nombre = UCase(Trim(Txt_Nombre))
        
        vlSql = ""
        vlSql = "SELECT p.num_poliza as numpol,"
        vlSql = vlSql & "b.num_orden as numord,"
        vlSql = vlSql & "b.gls_nomben as nomben,"
        vlSql = vlSql & "b.gls_patben as apaben,"
        vlSql = vlSql & "b.gls_matben as amaben,"
        vlSql = vlSql & "i.gls_tipoidencor as rutafi,"
        vlSql = vlSql & "b.num_idenben as dgvafi,"
        vlSql = vlSql & "p.num_cot as numcot,"
        vlSql = vlSql & "p.cod_cuspp as cuspp,"
        vlSql = vlSql & "p.num_endoso as numend "
        vlSql = vlSql & " FROM pd_tmae_poliza p, pd_tmae_polben b, ma_tpar_tipoiden i "
        vlSql = vlSql & " WHERE "
        
        vlSql = vlSql & " b.num_poliza = p.num_poliza"
        vlSql = vlSql & " AND b.num_endoso = p.num_endoso"
        vlSql = vlSql & " AND b.cod_par = '99'"
        vlSql = vlSql & " AND b.cod_tipoidenben = i.cod_tipoiden "
        
        If Trim(Txt_cotiz) <> "" Then vlSql = vlSql & " AND p.num_poliza like '%" & Trim(Txt_cotiz) & "%'"
        If Trim(Txt_Endoso) <> "" Then vlSql = vlSql & " AND p.num_endoso like '" & Trim(Txt_Endoso) & "%'"
        If Trim(Txt_NumCot) <> "" Then vlSql = vlSql & " AND p.num_cot like '" & Trim(Txt_NumCot) & "%'"
        If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  p.cod_cuspp like '" & Trim(Txt_Cuspp) & "%'"
        If Trim(Txt_TipoIden) <> "" Then vlSql = vlSql & " AND i.gls_cod_tipoidencor like '" & Trim(Txt_TipoIden) & "%'"
        If Trim(Txt_NumIden) <> "" Then vlSql = vlSql & " AND  b.num_idenben like '" & Trim(Txt_NumIden) & "%'"
        If Trim(Txt_Nombre) <> "" Then vlSql = vlSql & " AND  b.gls_nomben like '" & Trim(Txt_Nombre) & "%'"
        If Trim(Txt_ApellPat) <> "" Then vlSql = vlSql & " AND b.gls_patben like '" & Trim(Txt_ApellPat) & "%'"
        If Trim(Txt_ApellMat) <> "" Then vlSql = vlSql & " AND b.gls_matben like '" & Trim(Txt_ApellMat) & "%'"

        vlSql = vlSql & " ORDER BY b.num_poliza "
                      
        Set vgRs = vgConexionBD.Execute(vlSql)
        fila = 1
        Msf_GrillaBuscaCot.Rows = 1
        While Not vgRs.EOF
         
            vlNumPol = ""
            vlNumCot = ""
            vlRutAfi = ""
            vlcuspp = ""
            vlNomAfi = ""
            vlApellPat = ""
            vlApellMat = ""
            vlDig = ""
            vlNumOrd = 0
            vlEndoso = 0
                      
            If Not IsNull(vgRs!numord) Then vlNumOrd = Trim(vgRs!numord)
            If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
            If Not IsNull(vgRs!numend) Then vlEndoso = Trim(vgRs!numend)
            If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
            If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
            If Not IsNull(vgRs!rutafi) Then vlRutAfi = Trim(vgRs!rutafi)
            If Not IsNull(vgRs!dgvafi) Then vlDig = Trim(vgRs!dgvafi)
            If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
            If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
            If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)

            Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & vlEndoso & vbTab & _
                                vlNumCot & vbTab & vlcuspp & vbTab & vlRutAfi & vbTab & vlDig & vbTab & vlNomAfi & _
                                vbTab & vlApellPat & vbTab & vlApellMat
            fila = fila + 1
            vgRs.MoveNext
        Wend
    vgRs.Close
    Else
            
            Call flCargaGrillaPol
            
    End If
    Txt_ApellPat.SetFocus
End If

Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_NumCot_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Change

If KeyAscii = 13 Then
    
    If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_Endoso) <> "") Or (Trim(Txt_NumCot) <> "") Or _
       (Trim(Txt_Cuspp) <> "") Or (Trim(Txt_TipoIden) <> "") Or (Trim(Txt_NumIden) <> "") Or _
       (Trim(Txt_Nombre) <> "") Or (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
              
        Txt_NumCot = UCase(Trim(Txt_NumCot))
        
        vlSql = ""
        vlSql = "SELECT p.num_poliza as numpol,"
        vlSql = vlSql & "b.num_orden as numord,"
        vlSql = vlSql & "b.gls_nomben as nomben,"
        vlSql = vlSql & "b.gls_patben as apaben,"
        vlSql = vlSql & "b.gls_matben as amaben,"
        vlSql = vlSql & "i.gls_tipoidencor as rutafi,"
        vlSql = vlSql & "b.num_idenben as dgvafi,"
        vlSql = vlSql & "p.num_cot as numcot,"
        vlSql = vlSql & "p.cod_cuspp as cuspp,"
        vlSql = vlSql & "p.num_endoso as numend "
        vlSql = vlSql & " FROM pd_tmae_poliza p, pd_tmae_polben b, ma_tpar_tipoiden i "
        vlSql = vlSql & " WHERE "
        
        vlSql = vlSql & " b.num_poliza = p.num_poliza"
        vlSql = vlSql & " AND b.num_endoso = p.num_endoso"
        vlSql = vlSql & " AND b.cod_par = '99'"
        vlSql = vlSql & " AND b.cod_tipoidenben = i.cod_tipoiden "
        
        If Trim(Txt_cotiz) <> "" Then vlSql = vlSql & " AND p.num_poliza like '%" & Trim(Txt_cotiz) & "%'"
        If Trim(Txt_Endoso) <> "" Then vlSql = vlSql & " AND p.num_endoso like '" & Trim(Txt_Endoso) & "%'"
        If Trim(Txt_NumCot) <> "" Then vlSql = vlSql & " AND p.num_cot like '" & Trim(Txt_NumCot) & "%'"
        If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  p.cod_cuspp like '" & Trim(Txt_Cuspp) & "%'"
        If Trim(Txt_TipoIden) <> "" Then vlSql = vlSql & " AND i.gls_cod_tipoidencor like '" & Trim(Txt_TipoIden) & "%'"
        If Trim(Txt_NumIden) <> "" Then vlSql = vlSql & " AND  b.num_idenben like '" & Trim(Txt_NumIden) & "%'"
        If Trim(Txt_Nombre) <> "" Then vlSql = vlSql & " AND  b.gls_nomben like '" & Trim(Txt_Nombre) & "%'"
        If Trim(Txt_ApellPat) <> "" Then vlSql = vlSql & " AND b.gls_patben like '" & Trim(Txt_ApellPat) & "%'"
        If Trim(Txt_ApellMat) <> "" Then vlSql = vlSql & " AND b.gls_matben like '" & Trim(Txt_ApellMat) & "%'"

        vlSql = vlSql & " ORDER BY b.num_poliza "
                      
        Set vgRs = vgConexionBD.Execute(vlSql)
        fila = 1
        Msf_GrillaBuscaCot.Rows = 1
        While Not vgRs.EOF
         
            vlNumPol = ""
            vlNumCot = ""
            vlRutAfi = ""
            vlcuspp = ""
            vlNomAfi = ""
            vlApellPat = ""
            vlApellMat = ""
            vlDig = ""
            vlNumOrd = 0
            vlEndoso = 0
                      
            If Not IsNull(vgRs!numord) Then vlNumOrd = Trim(vgRs!numord)
            If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
            If Not IsNull(vgRs!numend) Then vlEndoso = Trim(vgRs!numend)
            If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
            If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
            If Not IsNull(vgRs!rutafi) Then vlRutAfi = Trim(vgRs!rutafi)
            If Not IsNull(vgRs!dgvafi) Then vlDig = Trim(vgRs!dgvafi)
            If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
            If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
            If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)

            Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & vlEndoso & vbTab & _
                                vlNumCot & vbTab & vlcuspp & vbTab & vlRutAfi & vbTab & vlDig & vbTab & vlNomAfi & _
                                vbTab & vlApellPat & vbTab & vlApellMat
            fila = fila + 1
            vgRs.MoveNext
        Wend
    vgRs.Close
    Else
            
            Call flCargaGrillaPol
            
    End If
    Txt_Cuspp.SetFocus
End If

Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_TipoIden_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Change

If KeyAscii = 13 Then
    
    If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_Endoso) <> "") Or (Trim(Txt_NumCot) <> "") Or _
       (Trim(Txt_Cuspp) <> "") Or (Trim(Txt_TipoIden) <> "") Or (Trim(Txt_NumIden) <> "") Or _
       (Trim(Txt_Nombre) <> "") Or (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
              
        Txt_TipoIden = UCase(Trim(Txt_TipoIden))
        
        vlSql = ""
        vlSql = "SELECT p.num_poliza as numpol,"
        vlSql = vlSql & "b.num_orden as numord,"
        vlSql = vlSql & "b.gls_nomben as nomben,"
        vlSql = vlSql & "b.gls_patben as apaben,"
        vlSql = vlSql & "b.gls_matben as amaben,"
        vlSql = vlSql & "i.gls_tipoidencor as rutafi,"
        vlSql = vlSql & "b.num_idenben as dgvafi,"
        vlSql = vlSql & "p.num_cot as numcot,"
        vlSql = vlSql & "p.cod_cuspp as cuspp,"
        vlSql = vlSql & "p.num_endoso as numend "
        vlSql = vlSql & " FROM pd_tmae_poliza p, pd_tmae_polben b, ma_tpar_tipoiden i "
        vlSql = vlSql & " WHERE "
        
        vlSql = vlSql & " b.num_poliza = p.num_poliza"
        vlSql = vlSql & " AND b.num_endoso = p.num_endoso"
        vlSql = vlSql & " AND b.cod_par = '99'"
        vlSql = vlSql & " AND b.cod_tipoidenben = i.cod_tipoiden "
        
        If Trim(Txt_cotiz) <> "" Then vlSql = vlSql & " AND p.num_poliza like '%" & Trim(Txt_cotiz) & "%'"
        If Trim(Txt_Endoso) <> "" Then vlSql = vlSql & " AND p.num_endoso like '" & Trim(Txt_Endoso) & "%'"
        If Trim(Txt_NumCot) <> "" Then vlSql = vlSql & " AND p.num_cot like '" & Trim(Txt_NumCot) & "%'"
        If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  p.cod_cuspp like '" & Trim(Txt_Cuspp) & "%'"
        If Trim(Txt_TipoIden) <> "" Then vlSql = vlSql & " AND i.gls_tipoidencor like '" & Trim(Txt_TipoIden) & "%'"
        If Trim(Txt_NumIden) <> "" Then vlSql = vlSql & " AND  b.num_idenben like '" & Trim(Txt_NumIden) & "%'"
        If Trim(Txt_Nombre) <> "" Then vlSql = vlSql & " AND  b.gls_nomben like '" & Trim(Txt_Nombre) & "%'"
        If Trim(Txt_ApellPat) <> "" Then vlSql = vlSql & " AND b.gls_patben like '" & Trim(Txt_ApellPat) & "%'"
        If Trim(Txt_ApellMat) <> "" Then vlSql = vlSql & " AND b.gls_matben like '" & Trim(Txt_ApellMat) & "%'"

        vlSql = vlSql & " ORDER BY b.num_poliza "
                      
        Set vgRs = vgConexionBD.Execute(vlSql)
        fila = 1
        Msf_GrillaBuscaCot.Rows = 1
        While Not vgRs.EOF
         
            vlNumPol = ""
            vlNumCot = ""
            vlRutAfi = ""
            vlcuspp = ""
            vlNomAfi = ""
            vlApellPat = ""
            vlApellMat = ""
            vlDig = ""
            vlNumOrd = 0
            vlEndoso = 0
                      
            If Not IsNull(vgRs!numord) Then vlNumOrd = Trim(vgRs!numord)
            If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
            If Not IsNull(vgRs!numend) Then vlEndoso = Trim(vgRs!numend)
            If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
            If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
            If Not IsNull(vgRs!rutafi) Then vlRutAfi = Trim(vgRs!rutafi)
            If Not IsNull(vgRs!dgvafi) Then vlDig = Trim(vgRs!dgvafi)
            If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
            If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
            If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)
        
            Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & vlEndoso & vbTab & _
                                vlNumCot & vbTab & vlcuspp & vbTab & vlRutAfi & vbTab & vlDig & vbTab & vlNomAfi & _
                                vbTab & vlApellPat & vbTab & vlApellMat
            fila = fila + 1
            vgRs.MoveNext
        Wend
    vgRs.Close
    Else
            
            Call flCargaGrillaPol
            
    End If
    Txt_Nombre.SetFocus
End If

Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub
         
Private Sub Txt_NumIden_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Change

If KeyAscii = 13 Then
    
    If (Trim(Txt_cotiz) <> "") Or (Trim(Txt_Endoso) <> "") Or (Trim(Txt_NumCot) <> "") Or _
       (Trim(Txt_Cuspp) <> "") Or (Trim(Txt_TipoIden) <> "") Or (Trim(Txt_NumIden) <> "") Or _
       (Trim(Txt_Nombre) <> "") Or (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
              
        Txt_NumIden = Trim(Txt_NumIden)
        
        vlSql = ""
        vlSql = "SELECT p.num_poliza as numpol,"
        vlSql = vlSql & "b.num_orden as numord,"
        vlSql = vlSql & "b.gls_nomben as nomben,"
        vlSql = vlSql & "b.gls_patben as apaben,"
        vlSql = vlSql & "b.gls_matben as amaben,"
        vlSql = vlSql & "i.gls_tipoidencor as rutafi,"
        vlSql = vlSql & "b.num_idenben as dgvafi,"
        vlSql = vlSql & "p.num_cot as numcot,"
        vlSql = vlSql & "p.cod_cuspp as cuspp,"
        vlSql = vlSql & "p.num_endoso as numend "
        vlSql = vlSql & " FROM pd_tmae_poliza p, pd_tmae_polben b, ma_tpar_tipoiden i "
        vlSql = vlSql & " WHERE "
        
        vlSql = vlSql & " b.num_poliza = p.num_poliza"
        vlSql = vlSql & " AND b.num_endoso = p.num_endoso"
        vlSql = vlSql & " AND b.cod_par = '99'"
        vlSql = vlSql & " AND b.cod_tipoidenben = i.cod_tipoiden "
        
        If Trim(Txt_cotiz) <> "" Then vlSql = vlSql & " AND p.num_poliza like '%" & Trim(Txt_cotiz) & "%'"
        If Trim(Txt_Endoso) <> "" Then vlSql = vlSql & " AND p.num_endoso like '" & Trim(Txt_Endoso) & "%'"
        If Trim(Txt_NumCot) <> "" Then vlSql = vlSql & " AND p.num_cot like '" & Trim(Txt_NumCot) & "%'"
        If Trim(Txt_Cuspp) <> "" Then vlSql = vlSql & " AND  p.cod_cuspp like '" & Trim(Txt_Cuspp) & "%'"
        If Trim(Txt_TipoIden) <> "" Then vlSql = vlSql & " AND i.gls_tipoidencor like '" & Trim(Txt_TipoIden) & "%'"
        If Trim(Txt_NumIden) <> "" Then vlSql = vlSql & " AND  b.num_idenben like '" & Trim(Txt_NumIden) & "%'"
        If Trim(Txt_Nombre) <> "" Then vlSql = vlSql & " AND  b.gls_nomben like '" & Trim(Txt_Nombre) & "%'"
        If Trim(Txt_ApellPat) <> "" Then vlSql = vlSql & " AND b.gls_patben like '" & Trim(Txt_ApellPat) & "%'"
        If Trim(Txt_ApellMat) <> "" Then vlSql = vlSql & " AND b.gls_matben like '" & Trim(Txt_ApellMat) & "%'"

        vlSql = vlSql & " ORDER BY b.num_poliza "
                      
        Set vgRs = vgConexionBD.Execute(vlSql)
        fila = 1
        Msf_GrillaBuscaCot.Rows = 1
        While Not vgRs.EOF
         
            vlNumPol = ""
            vlNumCot = ""
            vlRutAfi = ""
            vlcuspp = ""
            vlNomAfi = ""
            vlApellPat = ""
            vlApellMat = ""
            vlDig = ""
            vlNumOrd = 0
            vlEndoso = 0
                      
            If Not IsNull(vgRs!numord) Then vlNumOrd = Trim(vgRs!numord)
            If Not IsNull(vgRs!numpol) Then vlNumPol = Trim(vgRs!numpol)
            If Not IsNull(vgRs!numend) Then vlEndoso = Trim(vgRs!numend)
            If Not IsNull(vgRs!numcot) Then vlNumCot = Trim(vgRs!numcot)
            If Not IsNull(vgRs!cuspp) Then vlcuspp = Trim(vgRs!cuspp)
            If Not IsNull(vgRs!rutafi) Then vlRutAfi = Trim(vgRs!rutafi)
            If Not IsNull(vgRs!dgvafi) Then vlDig = Trim(vgRs!dgvafi)
            If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
            If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
            If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)
        
            Msf_GrillaBuscaCot.AddItem vlNumOrd & vbTab & vlNumPol & vbTab & vlEndoso & vbTab & _
                                vlNumCot & vbTab & vlcuspp & vbTab & vlRutAfi & vbTab & vlDig & vbTab & vlNomAfi & _
                                vbTab & vlApellPat & vbTab & vlApellMat
            fila = fila + 1
            vgRs.MoveNext
        Wend
    vgRs.Close
    Else
            
            Call flCargaGrillaPol
            
    End If
    Txt_Nombre.SetFocus
End If

Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub
         
Function flAsignaFormulario(vgNomForm, vlNumPol As String, vlNumEnd As Integer)

    If vgNomForm = "Frm_CalConsulta" Then
       Call Frm_CalConsulta.flBuscaPoliza(vlNumPol, vlNumEnd)
    End If
    If vgNomForm = "Frm_CalPrimaInf" Then
       Call Frm_CalPrimaInf.flRecibe(vlNumPol, vlNumEnd)
    End If
End Function

