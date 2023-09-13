VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_BuscaArchivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Archivo de Confirmación de Primas"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8505
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
      TabIndex        =   5
      Top             =   0
      Width           =   8295
      Begin VB.TextBox Txt_FecCarga 
         Height          =   285
         Left            =   6240
         TabIndex        =   2
         Top             =   480
         Width           =   1875
      End
      Begin VB.TextBox Txt_NomArch 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   480
         Width           =   4545
      End
      Begin VB.TextBox Txt_NumArch 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Archivo"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Fecha Carga"
         Height          =   255
         Index           =   2
         Left            =   6240
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "Nº Archivo"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   8295
      Begin VB.CommandButton Btn_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   3000
         Picture         =   "Frm_BuscaArchivo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4440
         Picture         =   "Frm_BuscaArchivo.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_GrillaBuscaArchivo 
      Height          =   2595
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4577
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   14745599
   End
End
Attribute VB_Name = "Frm_BuscaArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlSql As String

Dim vlNumArchivo As String
Dim vlNomArchivo As String
Dim vlFecCarga As String, strFecCarga As String

Dim vlFila As Long
Dim vlPos As Integer

'FUNCION QUE GUARDA EL NOMBRE DEL FORMULARIO DEL QUE SE LLAMO A LA FUNCION
Function flInicio(iNomForm)
    vgNomForm = iNomForm
    Call Form_Load
End Function

Function flAsignaFormulario(iNumArchivo As String)

    If vgNomForm = "Frm_CalPrimaArc" Then
        Call Frm_CalPrimaArc.flRecibeArchivo(iNumArchivo)
    End If
     
End Function

Function flLimpiar()
On Error GoTo Err_Limpiar

    Txt_NumArch.Text = ""
    Txt_NomArch.Text = ""
    Txt_FecCarga.Text = ""
    
Exit Function
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flCargaGrilla()
On Error GoTo Err_Carga

    Msf_GrillaBuscaArchivo.Clear
    Msf_GrillaBuscaArchivo.Enabled = True
    Msf_GrillaBuscaArchivo.Cols = 3
    Msf_GrillaBuscaArchivo.Rows = 1
    Msf_GrillaBuscaArchivo.Row = 0
    
    Msf_GrillaBuscaArchivo.Col = 0
    Msf_GrillaBuscaArchivo.CellAlignment = 4
    Msf_GrillaBuscaArchivo.ColWidth(0) = 1500
    Msf_GrillaBuscaArchivo.Text = "Nº Archivo"
    'Msf_GrillaBuscaArchivo.CellFontBold = True
        
    Msf_GrillaBuscaArchivo.Col = 1
    Msf_GrillaBuscaArchivo.CellAlignment = 4
    Msf_GrillaBuscaArchivo.ColWidth(1) = 4650
    Msf_GrillaBuscaArchivo.Text = "Nombre Archivo"
    'Msf_GrillaBuscaArchivo.CellFontBold = True
    
    Msf_GrillaBuscaArchivo.Col = 2
    Msf_GrillaBuscaArchivo.ColWidth(2) = 1850
    Msf_GrillaBuscaArchivo.CellAlignment = 4
    Msf_GrillaBuscaArchivo.Text = "Fecha Carga"
    'Msf_GrillaBuscaArchivo.CellFontBold = True
    
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
    Frm_CalPrimaArc.Enabled = True
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
    Msf_GrillaBuscaArchivo.Rows = 1

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
    Call flCargaGrilla
    
Exit Sub
Err_Cargar:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Frm_CalPrimaArc.Enabled = True
End Sub

Private Sub Msf_GrillaBuscaArchivo_DblClick()
On Error GoTo Err_Seleccionar

    Msf_GrillaBuscaArchivo.Col = 0
    Msf_GrillaBuscaArchivo.Row = Msf_GrillaBuscaArchivo.RowSel
    If Msf_GrillaBuscaArchivo.Text = "" Or (Msf_GrillaBuscaArchivo.Row = 0) Then
        Exit Sub
    End If

    Msf_GrillaBuscaArchivo.Col = 0
    vlNumArchivo = Trim(Msf_GrillaBuscaArchivo.Text)

    Msf_GrillaBuscaArchivo.Col = 1
    vlNomArchivo = Trim(Msf_GrillaBuscaArchivo.Text)
    
    Msf_GrillaBuscaArchivo.Col = 2
    vlFecCarga = Trim(Msf_GrillaBuscaArchivo.Text)
    
    Call flAsignaFormulario(vlNumArchivo)
    
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

    Txt_NumArch = Trim(Txt_NumArch)
    Txt_NomArch = UCase(Trim(Txt_NomArch))
    Txt_FecCarga = UCase(Trim(Txt_FecCarga))
    
    Msf_GrillaBuscaArchivo.Rows = 1
    vlFila = 1
    
    vlSql = "SELECT num_archivo,gls_nomarch,fec_crea "
    vlSql = vlSql & " FROM pd_tmae_estcarconpri"
    vlSql = vlSql & " WHERE num_archivo=num_archivo "
    If (iTodas = False) Then
        If Txt_NumArch <> "" Then vlSql = vlSql & " AND num_archivo LIKE '" & Txt_NumArch & "%'"
        If Txt_NomArch <> "" Then vlSql = vlSql & " AND gls_nomarch LIKE '" & Txt_NomArch & "%'"
        If Txt_FecCarga <> "" Then
            strFecCarga = Format(CDate(Trim(Txt_FecCarga)), "yyyyMMdd")
            vlSql = vlSql & " AND fec_crea LIKE '" & strFecCarga & "' "
        End If
    End If
    vlSql = vlSql & " ORDER BY num_archivo,gls_nomarch "
    Set vgRs = vgConexionBD.Execute(vlSql)
    While Not vgRs.EOF
        
        vlNumArchivo = ""
        vlNomArchivo = ""
        vlFecCarga = ""
        
        vlNumArchivo = Trim(vgRs!Num_Archivo)
        vlNomArchivo = Trim(vgRs!gls_nomarch)
        vlFecCarga = Mid(DateSerial(Mid(vgRs!Fec_Crea, 1, 4), Mid(vgRs!Fec_Crea, 5, 2), Mid(vgRs!Fec_Crea, 7, 2)), 1, 10)

        Msf_GrillaBuscaArchivo.AddItem vlNumArchivo & vbTab & _
                                        vlNomArchivo & vbTab & vlFecCarga
        
        vlFila = vlFila + 1
        
        vgRs.MoveNext
    Wend
    vgRs.Close

End Sub

Private Sub Txt_FecCarga_LostFocus()
    If Txt_FecCarga = "" Then
       Exit Sub
    End If
    If Not IsDate(Txt_FecCarga) Then
       Txt_FecCarga = ""
       Exit Sub
    End If
    If Txt_FecCarga <> "" Then
       vlFecCarga = Format(CDate(Trim(Txt_FecCarga)), "yyyymmdd")
       Txt_FecCarga = DateSerial(Mid((vlFecCarga), 1, 4), Mid((vlFecCarga), 5, 2), Mid((vlFecCarga), 7, 2))
    End If
End Sub

Private Sub Txt_NumArch_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_NumArch = UCase(Trim(Txt_NumArch))
        
        If (Txt_NumArch <> "") Or (Txt_NomArch <> "") Or (Trim(Txt_FecCarga) <> "") Then
            Call plGenerarConsulta
        Else
            Msf_GrillaBuscaArchivo.Rows = 1
        End If
        Txt_NomArch.SetFocus
    End If
 
Exit Sub
Err_Ingreso:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub
Private Sub Txt_NomArch_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        Txt_NomArch = UCase(Trim(Txt_NomArch))
        
        If (Txt_NumArch <> "") Or (Txt_NomArch <> "") Or (Trim(Txt_FecCarga) <> "") Then
            Call plGenerarConsulta
        Else
            Msf_GrillaBuscaArchivo.Rows = 1
        End If
        Txt_FecCarga.SetFocus
    End If
 
Exit Sub
Err_Ingreso:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
End Sub
Private Sub Txt_FecCarga_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Ingreso

    If KeyAscii = 13 Then
        If (Txt_FecCarga <> "") Then
            If flValidaFecha(Trim(Txt_FecCarga)) = False Then
                Txt_FecCarga.SetFocus
                Exit Sub
            Else
                vlFecCarga = Format(CDate(Trim(Txt_FecCarga)), "yyyymmdd")
                Txt_FecCarga = DateSerial(Mid((vlFecCarga), 1, 4), Mid((vlFecCarga), 5, 2), Mid((vlFecCarga), 7, 2))
            End If
            If (Txt_NumArch <> "") Or (Txt_NomArch <> "") Or (Trim(Txt_FecCarga) <> "") Then
                Call plGenerarConsulta
            End If
        Else
            Txt_NumArch.SetFocus
        End If
    End If
 
Exit Sub
Err_Ingreso:
  Screen.MousePointer = 0
  Select Case Err
    Case Else
      MsgBox "Error Grave [" & Err & Space(4) & Err.Description & "]", vbCritical
  End Select
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
