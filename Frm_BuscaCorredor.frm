VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_BuscaCorredor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Corredor"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   10650
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
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   10455
      Begin VB.TextBox Txt_ApellMat 
         Height          =   285
         Left            =   8160
         TabIndex        =   4
         Top             =   480
         Width           =   2000
      End
      Begin VB.TextBox Txt_ApellPat 
         Height          =   285
         Left            =   6000
         TabIndex        =   3
         Top             =   480
         Width           =   2235
      End
      Begin VB.TextBox Txt_Nombre 
         Height          =   285
         Left            =   3840
         TabIndex        =   2
         Top             =   480
         Width           =   2235
      End
      Begin VB.TextBox Txt_NumIden 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   480
         Width           =   1905
      End
      Begin VB.TextBox Txt_TipoIden 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "Nº Identificación"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Ap. Materno"
         Height          =   255
         Index           =   4
         Left            =   8160
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Ap. Paterno"
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Identificación"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   10455
      Begin VB.CommandButton Btn_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4200
         Picture         =   "Frm_BuscaCorredor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_BuscaCorredor.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_GrillaBuscaCorredor 
      Height          =   2595
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4577
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   14745599
   End
End
Attribute VB_Name = "Frm_BuscaCorredor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlSql As String

Dim vlTipoIdenAfi As String
Dim vlNumIdenAfi As String
Dim vlNomAfi As String
Dim vlApellPat As String
Dim vlApellMat As String
Dim vlCodTipoIden As String, vlBenSocial As String
Dim vlPrcComision As Double

Dim vlFila As Long
Dim vlPos As Integer

'FUNCION QUE GUARDA EL NOMBRE DEL FORMULARIO DEL QUE SE LLAMO A LA FUNCION
Function flInicio(iNomForm)
    vgNomForm = iNomForm
    Call Form_Load
End Function

Function flAsignaFormulario(vgNomForm As String, iNomTipoIden As String, iNumIden As String, iCodTipoIden As String, iBenSocial As String, iPrcComision As Double)

    If vgNomForm = "Frm_CalPoliza" Then
        Call Frm_CalPoliza.flRecibeCorredor(iNomTipoIden, iNumIden, iCodTipoIden, iBenSocial, iPrcComision)
    End If
    If vgNomForm = "Frm_CalPolizaRec" Then
        Call Frm_CalPolizaRec.flRecibeCorredor(iNomTipoIden, iNumIden, iCodTipoIden, iBenSocial, iPrcComision)
    End If
    
End Function

Function flLimpiar()
On Error GoTo Err_Limpiar

    Txt_Nombre.Text = ""
    Txt_ApellPat.Text = ""
    Txt_ApellMat.Text = ""
    Txt_NumIden.Text = ""
    Txt_TipoIden.Text = ""

Exit Function
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrillaCor()
Dim vlCuenta As Integer
Dim vlColumna As Integer
Dim vlNumIdent As String
Dim vlRutAfi As String
Dim vlNombre As String
Dim vlPaterno As String
Dim vlMaterno As String

On Error GoTo Err_Carga

    Msf_GrillaBuscaCorredor.Clear
    Msf_GrillaBuscaCorredor.Rows = 1
    
    Msf_GrillaBuscaCorredor.Enabled = True
    Msf_GrillaBuscaCorredor.Cols = 8
    Msf_GrillaBuscaCorredor.Rows = 1
    Msf_GrillaBuscaCorredor.Row = 0
    
    Msf_GrillaBuscaCorredor.Col = 0
    Msf_GrillaBuscaCorredor.CellAlignment = 4
    Msf_GrillaBuscaCorredor.ColWidth(0) = 1850
    Msf_GrillaBuscaCorredor.Text = "Tipo Identificación"
    Msf_GrillaBuscaCorredor.CellFontBold = True
        
    Msf_GrillaBuscaCorredor.Col = 1
    Msf_GrillaBuscaCorredor.CellAlignment = 4
    Msf_GrillaBuscaCorredor.ColWidth(1) = 1900
    Msf_GrillaBuscaCorredor.Text = "Nº Identificación"
    Msf_GrillaBuscaCorredor.CellFontBold = True
    
    Msf_GrillaBuscaCorredor.Col = 2
    Msf_GrillaBuscaCorredor.ColWidth(2) = 2150
    Msf_GrillaBuscaCorredor.CellAlignment = 4
    Msf_GrillaBuscaCorredor.Text = "Nombre"
    Msf_GrillaBuscaCorredor.CellFontBold = True
    
    Msf_GrillaBuscaCorredor.Col = 3
    Msf_GrillaBuscaCorredor.ColWidth(3) = 2150
    Msf_GrillaBuscaCorredor.CellAlignment = 4
    Msf_GrillaBuscaCorredor.Text = "Ap. Paterno"
    Msf_GrillaBuscaCorredor.CellFontBold = True
    
    Msf_GrillaBuscaCorredor.Col = 4
    Msf_GrillaBuscaCorredor.ColWidth(4) = 2200
    Msf_GrillaBuscaCorredor.CellAlignment = 4
    Msf_GrillaBuscaCorredor.CellFontBold = True
    Msf_GrillaBuscaCorredor.Text = "Ap. Materno"
    Msf_GrillaBuscaCorredor.CellFontBold = True
            
    Msf_GrillaBuscaCorredor.Col = 5
    Msf_GrillaBuscaCorredor.ColWidth(5) = 0
    Msf_GrillaBuscaCorredor.Text = "Cod Tipo Ident"

    Msf_GrillaBuscaCorredor.Col = 6
    Msf_GrillaBuscaCorredor.ColWidth(6) = 0
    Msf_GrillaBuscaCorredor.Text = "Ben Social"

    Msf_GrillaBuscaCorredor.Col = 7
    Msf_GrillaBuscaCorredor.ColWidth(7) = 0
    Msf_GrillaBuscaCorredor.Text = "Prc.Comisión"

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

    Call flLimpiar
    Msf_GrillaBuscaCorredor.Rows = 1

Exit Sub
Err_Limpiar:
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
    Call flCargaGrillaCor
    
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

Private Sub Msf_GrillaBuscaCorredor_DblClick()
On Error GoTo Err_Seleccionar

    Msf_GrillaBuscaCorredor.Col = 0
    Msf_GrillaBuscaCorredor.Row = Msf_GrillaBuscaCorredor.RowSel
    If Msf_GrillaBuscaCorredor.Text = "" Or (Msf_GrillaBuscaCorredor.Row = 0) Then
        Exit Sub
    End If

    Msf_GrillaBuscaCorredor.Col = 0
    vlTipoIdenAfi = Trim(Msf_GrillaBuscaCorredor.Text)

    Msf_GrillaBuscaCorredor.Col = 1
    vlNumIdenAfi = Trim(Msf_GrillaBuscaCorredor.Text)
    
    Msf_GrillaBuscaCorredor.Col = 5
    vlCodTipoIden = Trim(Msf_GrillaBuscaCorredor.Text)
    
    Msf_GrillaBuscaCorredor.Col = 6
    vlBenSocial = Trim(Msf_GrillaBuscaCorredor.Text)
    
    Msf_GrillaBuscaCorredor.Col = 7
    vlPrcComision = Trim(Msf_GrillaBuscaCorredor.Text)
    
    Call flAsignaFormulario(vgNomForm, vlTipoIdenAfi, vlNumIdenAfi, vlCodTipoIden, vlBenSocial, vlPrcComision)
    
    Unload Me
    
Exit Sub
Err_Seleccionar:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Sub plGenerarConsulta(Optional iTodas As Boolean)

    Txt_TipoIden = UCase(Trim(Txt_TipoIden))
    Txt_NumIden = UCase(Trim(Txt_NumIden))
    Txt_Nombre = UCase(Trim(Txt_Nombre))
    Txt_ApellPat = UCase(Trim(Txt_ApellPat))
    Txt_ApellMat = UCase(Trim(Txt_ApellMat))

    Msf_GrillaBuscaCorredor.Rows = 1
    vlFila = 1
    
    vlSql = "SELECT c.gls_nomcor as nomben,c.cod_tipoidencor,c.ind_bensocial,c.prc_comcor, "
    vlSql = vlSql & "c.gls_patcor as apaben,c.gls_matcor as amaben,"
    vlSql = vlSql & "i.gls_tipoidencor as tipoiden,c.num_idencor as numiden "
    vlSql = vlSql & "FROM pt_tmae_corredor c "
    vlSql = vlSql & ",ma_tpar_tipoiden i "
    vlSql = vlSql & " WHERE "
    vlSql = vlSql & " c.cod_tipoidencor = i.cod_tipoiden "
    If (iTodas = False) Then
        If Txt_TipoIden <> "" Then vlSql = vlSql & " AND i.gls_tipoidencor LIKE '" & Txt_TipoIden & "%'"
        If Txt_NumIden <> "" Then vlSql = vlSql & " AND c.num_idencor LIKE '" & Txt_NumIden & "%'"
        If Txt_Nombre <> "" Then vlSql = vlSql & " AND c.gls_nomcor LIKE '" & Txt_Nombre & "%'"
        If Txt_ApellPat <> "" Then vlSql = vlSql & " AND c.gls_patcor LIKE '" & Txt_ApellPat & "%'"
        If Txt_ApellMat <> "" Then vlSql = vlSql & " AND c.gls_matcor LIKE '" & Txt_ApellMat & "%'"
    End If
    vlSql = vlSql & " ORDER BY c.cod_tipoidencor,c.num_idencor "
    Set vgRs = vgConexionBD.Execute(vlSql)
    While Not vgRs.EOF
        
        vlTipoIdenAfi = ""
        vlNumIdenAfi = ""
        vlNomAfi = ""
        vlApellPat = ""
        vlApellMat = ""
        vlCodTipoIden = ""
        vlBenSocial = ""
        vlPrcComision = 0
        
        vlCodTipoIden = Trim(vgRs!cod_tipoidencor)
        vlBenSocial = Trim(vgRs!ind_bensocial)
        vlPrcComision = IIf(IsNull(vgRs!prc_comcor), 0, vgRs!prc_comcor)
        
        If Not IsNull(vgRs!tipoiden) Then vlTipoIdenAfi = Trim(vgRs!tipoiden)
        If Not IsNull(vgRs!numiden) Then vlNumIdenAfi = Trim(vgRs!numiden)
        If Not IsNull(vgRs!nomben) Then vlNomAfi = Trim(vgRs!nomben)
        If Not IsNull(vgRs!apaben) Then vlApellPat = Trim(vgRs!apaben)
        If Not IsNull(vgRs!amaben) Then vlApellMat = Trim(vgRs!amaben)

        Msf_GrillaBuscaCorredor.AddItem vlTipoIdenAfi & vbTab & _
        vlNumIdenAfi & vbTab & vlNomAfi & vbTab & _
        vlApellPat & vbTab & vlApellMat & vbTab & vlCodTipoIden & vbTab & _
        vlBenSocial & vbTab & vlPrcComision
        
        vlFila = vlFila + 1
        
        vgRs.MoveNext
    Wend
    vgRs.Close

End Sub

Private Sub Txt_ApellMat_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_ApellMat = UCase(Trim(Txt_ApellMat))
        
        If (Txt_TipoIden <> "") Or _
        (Txt_NumIden <> "") Or (Trim(Txt_Nombre) <> "") Or _
        (Trim(Txt_ApellPat) <> "") Or (Trim(Txt_ApellMat) <> "") Then
        
            Call plGenerarConsulta
        End If
        Btn_Limpiar.SetFocus
    End If
 
Exit Sub
Err_Ingreso:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_ApellPat_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_ApellPat = UCase(Trim(Txt_ApellPat))
        
        If (Txt_TipoIden <> "") Or _
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

Private Sub Txt_nombre_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_Nombre = UCase(Trim(Txt_Nombre))
        
        If (Txt_TipoIden <> "") Or _
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

Private Sub Txt_NumIden_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_NumIden = UCase(Trim(Txt_NumIden))
        
        If (Txt_TipoIden <> "") Or _
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
        
        If (Txt_TipoIden <> "") Or _
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
