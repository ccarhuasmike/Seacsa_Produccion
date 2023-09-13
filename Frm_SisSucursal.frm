VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "Crystl32.OCX"
Begin VB.Form Frm_SisSucursal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Sucursal."
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7455
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   7215
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   825
         Picture         =   "Frm_SisSucursal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Grabar Datos"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Eliminar 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   1995
         Picture         =   "Frm_SisSucursal.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Eliminar Registro"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Limpiar 
         Caption         =   "&Limpiar"
         Height          =   675
         Left            =   4440
         Picture         =   "Frm_SisSucursal.frx":09FC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpiar Formulario"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   3240
         Picture         =   "Frm_SisSucursal.frx":10B6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Reporte"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5640
         Picture         =   "Frm_SisSucursal.frx":1770
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Fra_Operacion 
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
      Height          =   1335
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   7215
      Begin VB.TextBox Txt_Descripcion 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox Txt_Codigo 
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin Crystal.CrystalReport Rpt_General 
         Left            =   5040
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Descripción          :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Código                  :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf_Grilla 
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   14745599
      AllowBigSelection=   0   'False
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
Attribute VB_Name = "Frm_SisSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DECLARACION DE VARIABLES
Const clLargoCodigo As Integer = 5

Dim vlRegistro   As ADODB.Recordset
Dim vlContador   As Integer
Dim vlResp       As Integer
Dim vlPos        As Integer
Dim vlAccion     As String
Dim vlOperacion  As String
Dim vlSw         As Boolean
Dim vlTablaSeleccionada As String
Dim vlGlosaTabla        As String

'--------------------------------------
'Permite Limpiar los Casilleros de Información
'--------------------------------------
Function flLimpia()
    Txt_Codigo.Text = ""
    Txt_Descripcion.Text = ""
    Txt_Codigo.Enabled = True
End Function

'--------------------------------------
'Permite Limpiar la Grilla desde la cual se muestran los Datos
'--------------------------------------
Function flLmpGrilla()
    Msf_Grilla.Clear
    Msf_Grilla.Rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0
    
    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Código"
    Msf_Grilla.ColWidth(0) = 1200
    Msf_Grilla.ColAlignment(0) = 1  'centrado
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Descripción"
    Msf_Grilla.ColWidth(1) = 5300

End Function

'--------------------------------------
'Permite cargar los Datos en la Grilla para desplegarlos por pantalla
'--------------------------------------
Function flCargaGrilla()
On Error GoTo Err_Carga

    vgSql = "SELECT cod_sucursal as codigo,gls_sucursal as nombre "
    vgSql = vgSql & "FROM pd_tpar_sucursal "
    vgSql = vgSql & "order by cod_sucursal "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    If Not vlRegistro.EOF Then
        While Not vlRegistro.EOF
        
            Msf_Grilla.AddItem CStr(Trim(vlRegistro!Codigo)) & vbTab _
            & Trim(vlRegistro!Nombre)
            vlRegistro.MoveNext
        Wend
    End If
    vlRegistro.Close

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flRegistrar()
Dim vlEncuentra As String
Dim vlValor As Double
On Error GoTo Err_Registrar

    'Validar Código del elemento
    If Trim(Txt_Codigo = "") Then
        MsgBox "Debe ingresar un Código representativo para el Parámetro.", vbCritical, "Error de Datos"
        Txt_Codigo.SetFocus
        Exit Function
    End If
    'Validar Descripción
    If Trim(Txt_Descripcion) = "" Then
        MsgBox "Debe ingresar un nombre descriptivo para el Parámetro.", vbCritical, "Error de Datos"
        Txt_Descripcion.SetFocus
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    Txt_Codigo = UCase(Trim(Txt_Codigo))
    Txt_Codigo = String(clLargoCodigo - Len(Txt_Codigo), "0") & Txt_Codigo
    Txt_Descripcion = UCase(Trim(Txt_Descripcion))
        
    vlOperacion = ""
    vlEncuentra = ""

    'Valida la Existencia de los Códigos en TPAR_TABCOD
    'Códigos en la Tabla de Detalles
    'A : Actualización de los Datos
    'I : Insertar los nuevos Datos
    vgSql = "SELECT cod_sucursal FROM pd_tpar_sucursal WHERE "
    vgSql = vgSql & "cod_sucursal = '" & Txt_Codigo & "'"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not (vgRs.EOF) Then
        vlOperacion = "A"
    Else
        vlOperacion = "I"
    End If
    vgRs.Close
    
    If vlOperacion = "A" Then
        vlResp = MsgBox(" ¿ Está seguro que desea Modificar los Datos ?", 4 + 32 + 256, "Actualización")
        If vlResp <> 6 Then
            Screen.MousePointer = 0
            Exit Function
        End If
            
        'Actualiza los Datos en la tabla TPAR_TABCOD
        Sql = "update pd_tpar_sucursal set"
        Sql = Sql & " gls_sucursal = '" & (Txt_Descripcion) & "', "
        Sql = Sql & " cod_usuariomodi = '" & (vgUsuario) & "',"
        Sql = Sql & " fec_modi = '" & Format(Date, "yyyymmdd") & "',"
        Sql = Sql & " hor_modi = '" & Format(Time, "hhmmss") & "'"
        Sql = Sql & " where "
        Sql = Sql & " cod_sucursal = '" & Txt_Codigo & "' "
        vgConexionBD.Execute (Sql)
            
    Else
       'Inserta los Datos en la Tabla TPAR_TABCOD
        Sql = "insert into pd_tpar_sucursal ("
        Sql = Sql & "cod_sucursal,gls_sucursal,"
        Sql = Sql & "cod_usuariocrea,fec_crea,hor_crea"
        Sql = Sql & ") values ("
        Sql = Sql & "'" & Txt_Codigo & "',"
        Sql = Sql & "'" & (Txt_Descripcion) & "',"
        Sql = Sql & "'" & (vgUsuario) & "',"
        Sql = Sql & "'" & Format(Date, "yyyymmdd") & "',"
        Sql = Sql & "'" & Format(Time, "hhmmss") & "'"
        Sql = Sql & ")"
        vgConexionBD.Execute (Sql)
    End If
    
    If (vlOperacion <> "") Then
        'Limpia los Datos de la Pantalla
        flLimpia
        
        flLmpGrilla
        flCargaGrilla
        
        Txt_Codigo.SetFocus
    End If
    Screen.MousePointer = 0

Exit Function
Err_Registrar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flEliminar()
On Error GoTo Err_Eliminar

    'Validación de Datos
    If (Trim(Txt_Codigo) = "") Then
        MsgBox "Debe ingresar el Código del Parámetro a eliminar.", vbCritical, "Error de Datos"
        Txt_Codigo.SetFocus
        Exit Function
    End If

    Txt_Codigo = UCase(Trim(Txt_Codigo))
    vlOperacion = ""
    vlSw = False

    Screen.MousePointer = 11

    vgQuery = "SELECT cod_sucursal FROM pd_TPAR_sucursal WHERE "
    vgQuery = vgQuery & "cod_sucursal = '" & Txt_Codigo & "'"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not (vgRs.EOF) Then
        vlOperacion = "E"
    End If
    vgRs.Close

    If (vlOperacion = "E") Then
        vgRes = MsgBox(" ¿ Esta seguro que desea Eliminar los Datos ? ", vbQuestion + vbYesNo + 256, "Operación de Eliminación")
        If vgRes <> 6 Then
            Screen.MousePointer = 0
            
            Exit Function
        End If
        
        vgQuery = "DELETE FROM pd_tpar_sucursal WHERE "
        vgQuery = vgQuery & "cod_sucursal = '" & Txt_Codigo & "'"
        vgConexionBD.Execute (vgQuery)
        
        Call flLimpia
        Call flLmpGrilla
        Call flCargaGrilla
        
        Txt_Codigo.SetFocus
    End If
    Screen.MousePointer = 0
    
Exit Function
Err_Eliminar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Sub flImpresion()
Dim vlArchivo As String

Err.Clear
On Error GoTo Errores1
   
   Screen.MousePointer = 11
   
   vlArchivo = strRpt & "PD_Rpt_ParSucursal.rpt"   '\Reportes
   If Not fgExiste(vlArchivo) Then     ', vbNormal
       MsgBox "Archivo de Reporte de Listados no se encuentra en el directorio de la Aplicación.", 16, "Archivo no encontrado"
       Screen.MousePointer = 0
       Exit Sub
   End If
   
   vgQuery = ""
   
   Rpt_General.Reset
   Rpt_General.ReportFileName = vlArchivo     'App.Path & "\Rpt_Areas.rpt"
   Rpt_General.Connect = vgRutaDataBase
   Rpt_General.SelectionFormula = ""
   Rpt_General.Formulas(0) = ""
   Rpt_General.Formulas(1) = ""
   Rpt_General.Formulas(2) = ""
   
   Rpt_General.Formulas(0) = "NombreCompania = '" & vgNombreCompania & "'"
   Rpt_General.Formulas(1) = "NombreSistema= '" & vgNombreSistema & "'"
   Rpt_General.Formulas(2) = "NombreSubSistema= '" & vgNombreSubSistema & "'"
   
   Rpt_General.Destination = crptToWindow
   Rpt_General.WindowState = crptMaximized
   Rpt_General.WindowTitle = "Informe de Sucursales de Producción"
   Rpt_General.Action = 1
   
   Screen.MousePointer = 0
   
Exit Sub
Errores1:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & "Numero de Error : " & Err.Number & Err.HelpFile & Err.HelpContext, vbExclamation, "Mensaje de error ..."
        Screen.MousePointer = 0
    End If
End Sub
'-------------------------------------------------------------
'PROCEDIMIENTOS DE LOS OBJETOS
'-------------------------------------------------------------

Private Sub Cmd_Eliminar_Click()
On Error GoTo Err_Eliminar

    Call flEliminar

Exit Sub
Err_Eliminar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Grabar_Click()
On Error GoTo Err_Grabar
    
    Call flRegistrar
    
Exit Sub
Err_Grabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Err_Imprimir

    'Limpia los Datos del Formulario
    flLimpia
    'Imprime el Reporte de Variables
    flImpresion

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

    Call flLimpia
    Txt_Codigo.SetFocus

Exit Sub
Err_Limpiar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Salir_Click()
On Error GoTo Err_Salir

    Screen.MousePointer = 11
    Unload Me
    Screen.MousePointer = 0
            
Exit Sub
Err_Salir:
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
    
    'Limpiar la Grilla
    flLmpGrilla
    'Actualizar los Datos en la Grilla
    flCargaGrilla
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_Grilla_Click()
On Error GoTo Err_Grilla

    Msf_Grilla.Col = 0
    vlPos = Msf_Grilla.RowSel
    If (Msf_Grilla.Text = "") Or (Msf_Grilla.Row = 0) Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    
    Msf_Grilla.Row = vlPos
    Msf_Grilla.Col = 0
    Txt_Codigo = Trim(Msf_Grilla.Text)
    Msf_Grilla.Col = 1
    Txt_Descripcion = Trim(Msf_Grilla.Text)
    
    Txt_Codigo.Enabled = False
    Txt_Descripcion.SetFocus
    Screen.MousePointer = 0

Exit Sub
Err_Grilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Codigo_GotFocus()
    Txt_Codigo.SelStart = 0
    Txt_Codigo.SelLength = Len(Txt_Codigo)
End Sub

Private Sub Txt_Codigo_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Busqueda

If (KeyAscii = 13) Then
    If (Trim(Txt_Codigo) <> "") Then
        Txt_Codigo = UCase(Trim(Txt_Codigo))
        Txt_Codigo = String(clLargoCodigo - Len(Txt_Codigo), "0") & Txt_Codigo
        
        'Busqueda de la existencia del Código
        vgSql = "SELECT gls_sucursal as descripcion "
        vgSql = vgSql & "FROM pd_tpar_sucursal WHERE "
        vgSql = vgSql & "cod_sucursal = '" & Txt_Codigo & "' "
        Set vgRs = vgConexionBD.Execute(vgSql)
        If Not vgRs.EOF Then
            Txt_Descripcion = UCase(Trim(vgRs!Descripcion))
            Txt_Codigo.Enabled = False
            'Txt_Descripcion.SetFocus
        Else
            Txt_Codigo.Enabled = True
        End If
        vgRs.Close
        
        Txt_Descripcion.SetFocus
    End If
End If

Exit Sub
Err_Busqueda:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub txt_codigo_LostFocus()
On Error GoTo Err_Busqueda
    
    If (Trim(Txt_Codigo) <> "") Then
        Txt_Codigo = UCase(Trim(Txt_Codigo))
        Txt_Codigo = String(clLargoCodigo - Len(Txt_Codigo), "0") & Txt_Codigo
        
        'Busqueda de la existencia del Código
        vgSql = "SELECT gls_sucursal as descripcion FROM pd_tpar_sucursal WHERE "
        vgSql = vgSql & "cod_sucursal = '" & Txt_Codigo & "' "
        Set vgRs = vgConexionBD.Execute(vgSql)
        If Not vgRs.EOF Then
            Txt_Descripcion = UCase(Trim(vgRs!Descripcion))
            Txt_Codigo.Enabled = False
            'Txt_Descripcion.SetFocus
        Else
            Txt_Codigo.Enabled = True
        End If
        vgRs.Close
    End If
    
Exit Sub
Err_Busqueda:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_Descripcion_GotFocus()
    Txt_Descripcion.SelStart = 0
    Txt_Descripcion.SelLength = Len(Txt_Descripcion)
End Sub

Private Sub Txt_Descripcion_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Trim(Txt_Descripcion) <> "") Then
        Txt_Descripcion = UCase(Trim(Txt_Descripcion))
        Cmd_Grabar.SetFocus
    End If
End If
End Sub

Private Sub txt_descripcion_LostFocus()
    If (Trim(Txt_Descripcion) <> "") Then
        Txt_Descripcion = UCase(Trim(Txt_Descripcion))
    End If
End Sub
