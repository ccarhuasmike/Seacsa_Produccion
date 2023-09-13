VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_BuscaCauInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Causal de Invalidez"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7890
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   7695
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4080
         Picture         =   "Frm_BuscaCauInv.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   2640
         Picture         =   "Frm_BuscaCauInv.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpiar Formulario"
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.TextBox Txt_CodCauInv 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Txt_DescCodCauInv 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Código"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Descripción"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_GrillaBuscaCauInv 
      Height          =   2235
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3942
      _Version        =   393216
      BackColor       =   14745599
   End
   Begin VB.Label Lbl_Buscador 
      Caption         =   "Resultado Búsqueda"
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
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Frm_BuscaCauInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlSql As String

Dim vlCodigo As String
Dim vlDescripcion As String
Dim vlCauInv As String
Dim vlFila As Integer

Function flLimpiar()
On Error GoTo Err_Limpia

    Txt_CodCauInv.Text = ""
    Txt_DescCodCauInv.Text = ""
    Txt_CodCauInv.SetFocus
    
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
Dim vlCodigo As String
Dim vlDescripcion As String

'On Error GoTo Err_Carga
    Msf_GrillaBuscaCauInv.Rows = 1
    
    Msf_GrillaBuscaCauInv.Enabled = True
    Msf_GrillaBuscaCauInv.Cols = 3
    Msf_GrillaBuscaCauInv.Rows = 1
    Msf_GrillaBuscaCauInv.Row = 0
    
    Msf_GrillaBuscaCauInv.Col = 0
    Msf_GrillaBuscaCauInv.ColWidth(0) = 0
    
    Msf_GrillaBuscaCauInv.Col = 1
    Msf_GrillaBuscaCauInv.CellAlignment = 4
    Msf_GrillaBuscaCauInv.ColWidth(1) = 1300
    Msf_GrillaBuscaCauInv.Text = "Código"
    Msf_GrillaBuscaCauInv.CellFontBold = True
        
    Msf_GrillaBuscaCauInv.Col = 2
    Msf_GrillaBuscaCauInv.ColWidth(2) = 6000
    Msf_GrillaBuscaCauInv.CellAlignment = 4
    Msf_GrillaBuscaCauInv.Text = "Descripción"
    Msf_GrillaBuscaCauInv.CellFontBold = True
    
    If Not fgConexionBaseDatos(vgConectarBD) Then
        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
        Exit Function
    End If
    
    Sql = ""
    Sql = "SELECT p.cod_patologia,p.gls_patologia "
    Sql = Sql & " FROM ma_tpar_patologia p "
    Sql = Sql & " ORDER BY p.cod_patologia "
    Set vgRs = vgConectarBD.Execute(Sql)
        
    vlCuenta = 1
    Msf_GrillaBuscaCauInv.Rows = 1
        Do While Not (vgRs.EOF)
            vlCodigo = vgRs!cod_patologia
            vlDescripcion = vgRs!gls_patologia
            
            Msf_GrillaBuscaCauInv.AddItem vlCuenta & vbTab & vlCodigo & vbTab & _
                                    vlDescripcion
        
            vlCuenta = vlCuenta + 1
            vgRs.MoveNext
        Loop
    vgRs.Close
    vgConectarBD.Close
    
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
    If vgFormulario = "P" Then
        Frm_CalPoliza.Enabled = True
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
    flLimpiar
End Sub
Private Sub Form_Load()
On Error GoTo Err_Cargar

    Me.Top = 0
    Me.Left = 0
    
    Call flCargaGrilla
    
    Txt_CodCauInv.Text = ""
    Txt_DescCodCauInv.Text = ""
    
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

Private Sub Msf_GrillaBuscaCauInv_DblClick()
On Error GoTo Err_Seleccionar

    Msf_GrillaBuscaCauInv.Col = 1
    Msf_GrillaBuscaCauInv.Row = Msf_GrillaBuscaCauInv.RowSel
    If (Not (Msf_GrillaBuscaCauInv.Text = "") And (Msf_GrillaBuscaCauInv.Row <> 0)) Then
        If vgFormulario = "P" Then
            vlCauInv = Msf_GrillaBuscaCauInv.Text
            Msf_GrillaBuscaCauInv.Col = 2
            vlCauInv = Trim(vlCauInv) & " - " & Trim(Msf_GrillaBuscaCauInv.Text)
            If (vgFormularioCarpeta = "A") Then
                Frm_CalPoliza.Lbl_CauInv = vlCauInv
            Else
                Frm_CalPoliza.Lbl_CauInvBen = vlCauInv
            End If
            Frm_CalPoliza.flDatosCompletos
            
            Frm_CalPoliza.Enabled = True
            
            Unload Me
        End If
        If vgFormulario = "R" Then
            vlCauInv = Msf_GrillaBuscaCauInv.Text
            Msf_GrillaBuscaCauInv.Col = 2
            vlCauInv = Trim(vlCauInv) & " - " & Trim(Msf_GrillaBuscaCauInv.Text)
            If (vgFormularioCarpeta = "A") Then
                Frm_CalPolizaRec.Lbl_CauInv = vlCauInv
            Else
                Frm_CalPolizaRec.Lbl_CauInvBen = vlCauInv
            End If
            Frm_CalPolizaRec.flDatosCompletos
            
            Frm_CalPolizaRec.Enabled = True
            Unload Me
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

Private Sub Txt_CodCauInv_Change()
On Error GoTo Err_Change

    vlSql = ""
    vlSql = "SELECT p.cod_patologia as codigo,"
    vlSql = vlSql & " p.gls_patologia as glosa"
    vlSql = vlSql & " FROM ma_tpar_patologia p "
    vlSql = vlSql & " WHERE "
    vlSql = vlSql & " p.cod_patologia LIKE '" & Txt_CodCauInv & "%'"
    If Trim(Txt_DescCodCauInv) <> "" Then vlSql = vlSql & " AND p.gls_patologia LIKE '" & Trim(Txt_DescCodCauInv) & "%'"
    vlSql = vlSql & " ORDER BY p.cod_patologia "
    Set vgRs = vgConexionBD.Execute(vlSql)
    
    vlFila = 1
    Msf_GrillaBuscaCauInv.Rows = 1
    While Not vgRs.EOF
     
        vlCodigo = ""
        vlDescripcion = ""
                
        If Not IsNull(vgRs!Codigo) Then vlCodigo = Trim(vgRs!Codigo)
        If Not IsNull(vgRs!glosa) Then vlDescripcion = Trim(vgRs!glosa)
                
        Msf_GrillaBuscaCauInv.AddItem vlFila & vbTab & vlCodigo & vbTab & vlDescripcion
        vlFila = vlFila + 1
        vgRs.MoveNext
    Wend
    vgRs.Close

Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_CodCauInv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txt_DescCodCauInv.SetFocus
End If

End Sub

Private Sub Txt_DescCodCauInv_Change()
On Error GoTo Err_Change

    vlSql = ""
    vlSql = "SELECT p.cod_patologia as codigo,"
    vlSql = vlSql & "p.gls_patologia as glosa "
    vlSql = vlSql & " FROM ma_tpar_patologia p "
    vlSql = vlSql & " WHERE p.gls_patologia LIKE '" & Trim(Txt_DescCodCauInv) & "%'"
    If Trim(Txt_CodCauInv) <> "" Then vlSql = vlSql & " AND  p.cod_patologia LIKE '" & Trim(Txt_CodCauInv) & "%'"
    vlSql = vlSql & " ORDER BY p.cod_patologia "
    Set vgRs2 = vgConexionBD.Execute(vlSql)
    
    vlFila = 1
    Msf_GrillaBuscaCauInv.Rows = 1
    While Not vgRs2.EOF
             
        vlCodigo = ""
        vlDescripcion = ""
        
        If Not IsNull(vgRs2!Codigo) Then vlCodigo = Trim(vgRs2!Codigo)
        If Not IsNull(vgRs2!glosa) Then vlDescripcion = Trim(vgRs2!glosa)
        
        Msf_GrillaBuscaCauInv.AddItem vlFila & vbTab & vlCodigo & vbTab & vlDescripcion
        vlFila = vlFila + 1
        vgRs2.MoveNext
    Wend
vgRs2.Close

Exit Sub
Err_Change:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Txt_DescCodCauInv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Btn_Limpiar.SetFocus
    End If
End Sub
