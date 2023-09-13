VERSION 5.00
Begin VB.Form Frm_SisApoderado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Antecedentes del Analista."
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "Frm_SisApoderado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6630
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   6375
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3360
         Picture         =   "Frm_SisApoderado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmd_grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2280
         Picture         =   "Frm_SisApoderado.frx":053C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1935
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6375
      Begin VB.TextBox Txt_nombre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   35
         TabIndex        =   0
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox Txt_cargo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   40
         TabIndex        =   1
         Top             =   1320
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "Los Informes irán certificados por :"
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
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Cuyo cargo actual es:"
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
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Frm_SisApoderado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Grabar_Click()
On Error GoTo Err_Grabar
    
    Txt_cargo = (Trim(Txt_cargo))
    Txt_Nombre = (Trim(Txt_Nombre))
    
    If Txt_Nombre = "" Or IsNull(Txt_Nombre) Then
        MsgBox "Debe ingresar el Nombre del Representante o Apoderado.", vbCritical, "Error de Datos"
        Txt_Nombre.SetFocus
        Exit Sub
    End If
    If Txt_cargo = "" Or IsNull(Txt_cargo) Then
        MsgBox "Debe ingresar el Cargo del Representante o Apoderado.", vbCritical, "Error de Datos"
        Txt_cargo.SetFocus
        Exit Sub
    End If
    
    'Call AbrirBaseDeDatos(vgRutaBasedeDatos)
'''    If Not AbrirBaseDeDatos(vgConexionBD) Then
'''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
'''        Exit Sub
'''    End If
    
    vgSql = "SELECT * FROM pd_TMAE_APODERADO "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        vgPregunta = MsgBox("¿ Esta seguro de modificar los datos ?", vbQuestion + vbYesNo + 256, "Verificar Operación")
        If vgPregunta = 6 Then
            vgGra = "UPDATE pd_TMAE_APODERADO SET "
            vgGra = vgGra & "GLS_nomApo = '" & Txt_Nombre & "',"
            vgGra = vgGra & "GLS_carApo = '" & Txt_cargo & "',"
            'I---- KVR 06/09/2004 ----
            vgGra = vgGra & "cod_usuariomodi = '" & vgUsuario & "',"
            vgGra = vgGra & "fec_modi = '" & Format(Date, "yyyymmdd") & "',"
            vgGra = vgGra & "hor_modi = '" & Format(Time, "hhmmss") & "'"
            'F---- KVR 06/09/2004 ----
            vgConexionBD.Execute vgGra
        Else
'''            Call CerrarBaseDeDatos(vgConexionBD)
            cmd_salir.SetFocus
            Exit Sub
        End If
        'Exit Sub
    Else
        vgGra = "INSERT INTO pd_TMAE_APODERADO (GLS_nomApo,"
        vgGra = vgGra & "GLS_carApo"
        'I---- KVR 06/09/2004 ----
        vgGra = vgGra & ",cod_usuariocrea,fec_crea,hor_crea"
        'F---- KVR 06/09/2004 ----
        vgGra = vgGra & ") VALUES ("
        vgGra = vgGra & "'" & Txt_Nombre & "',"
        vgGra = vgGra & "'" & Txt_cargo & "',"
        'I---- KVR 06/09/2004 ----
        vgGra = vgGra & "'" & vgUsuario & "',"
        vgGra = vgGra & "'" & Format(Date, "yyyymmdd") & "',"
        vgGra = vgGra & "'" & Format(Time, "hhmmss") & "')"
        'F---- KVR 06/09/2004 ----
        vgConexionBD.Execute vgGra
    End If
    vgRs.Close
    
'''    Call CerrarBaseDeDatos(vgConexionBD)
    
    Call fgApoderado
    
    MsgBox "Los datos han sido grabados Satisfactoriamente.", vbInformation, "Estado de la Actualización"
    cmd_salir.SetFocus
    
Exit Sub
Err_Grabar:
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

    Frm_SisApoderado.Top = 0
    Frm_SisApoderado.Left = 0
    
    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)
'    If Not AbrirBaseDeDatos(vgConexionBD) Then
'        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
'        Exit Sub
'    End If

    
    vgSql = "SELECT * FROM pd_TMAE_APODERADO"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       Txt_Nombre = Trim(vgRs!gls_nomApo)
       Txt_cargo = Trim(vgRs!gls_carApo)
    End If
    vgRs.Close
    
'''    Call CerrarBaseDeDatos(vgConectarBD)

Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Txt_cargo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Trim(Txt_cargo) <> "") Then
        Txt_cargo = (Trim(Txt_cargo))
        cmd_grabar.SetFocus
    End If
End If
End Sub

'Private Sub Txt_cargo_LostFocus()
'Txt_cargo = UCase(Txt_cargo)
'End Sub

Private Sub Txt_nombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Trim(Txt_Nombre) <> "") Then
        Txt_Nombre = (Trim(Txt_Nombre))
        Txt_cargo.SetFocus
    End If
End If
End Sub

'Private Sub Txt_nombre_LostFocus()
'Txt_nombre = UCase(Txt_nombre)
'End Sub
