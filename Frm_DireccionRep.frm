VERSION 5.00
Begin VB.Form Frm_DireccionRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direccion del Representante"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6870
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtdirec 
      Height          =   405
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   200
      TabIndex        =   20
      Top             =   5280
      Width           =   6375
   End
   Begin VB.ComboBox cboconj 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "Frm_DireccionRep.frx":0000
      Left            =   1080
      List            =   "Frm_DireccionRep.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame Direccion 
      Caption         =   "Direccion"
      Height          =   3495
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   6615
      Begin VB.CommandButton Cmd_BuscarDir 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Efectuar Busqueda de Dirección"
         Top             =   2280
         Width           =   300
      End
      Begin VB.TextBox Txt_Referencia 
         Height          =   285
         Left            =   960
         MaxLength       =   30
         TabIndex        =   19
         Top             =   3000
         Width           =   4335
      End
      Begin VB.TextBox txtnombreconj 
         Height          =   285
         Left            =   4440
         TabIndex        =   15
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtnumero 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   8
         Top             =   810
         Width           =   735
      End
      Begin VB.TextBox txtdireccion 
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox cbotipovia 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "Frm_DireccionRep.frx":0004
         Left            =   960
         List            =   "Frm_DireccionRep.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Txt_Interior 
         Height          =   285
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cbodepart 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtmanzana 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtlote 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtetapa 
         Height          =   285
         Left            =   5640
         TabIndex        =   13
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cbobloque 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "Frm_DireccionRep.frx":0008
         Left            =   960
         List            =   "Frm_DireccionRep.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtbloque 
         Height          =   285
         Left            =   5520
         TabIndex        =   17
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Prov.:"
         Height          =   255
         Left            =   3720
         TabIndex        =   45
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Ref.:"
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   44
         Top             =   3000
         Width           =   405
      End
      Begin VB.Label Lbl_DistritoEdit 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   43
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Lbl_ProvinciaEdit 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4320
         TabIndex        =   42
         Top             =   2280
         Width           =   1725
      End
      Begin VB.Label Lbl_DepartamentoEdit 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   41
         Top             =   2280
         Width           =   2490
      End
      Begin VB.Label Lbl_Afiliado 
         Caption         =   "Distrito:"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   40
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Lbl_Afiliado 
         Caption         =   "Dpto.:"
         Height          =   375
         Index           =   15
         Left            =   360
         TabIndex        =   39
         Top             =   2280
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre Conj.Habit:"
         Height          =   255
         Left            =   2760
         TabIndex        =   38
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Conj.Habit:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Direccion :"
         Height          =   375
         Left            =   2880
         TabIndex        =   35
         Top             =   405
         Width           =   855
      End
      Begin VB.Label lblTipoVia 
         Caption         =   "Via :"
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Interior:"
         Height          =   255
         Index           =   13
         Left            =   4920
         TabIndex        =   33
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblmanzana 
         Caption         =   "Manzana:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lbllote 
         Caption         =   "Lote:"
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbletapa 
         Caption         =   "Etapa:"
         Height          =   255
         Left            =   5040
         TabIndex        =   30
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Num/Letra"
         Height          =   255
         Left            =   4320
         TabIndex        =   29
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Pref.Depart:"
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblbloque 
         Caption         =   "Bloque/ Chalet"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Telefono"
      Height          =   1455
      Left            =   240
      TabIndex        =   22
      Top             =   120
      Width           =   6375
      Begin VB.TextBox Txt_celular 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cboCiudad2 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboTipoTelefono2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_DireccionRep.frx":000C
         Left            =   3600
         List            =   "Frm_DireccionRep.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cboTipoTelefono 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_DireccionRep.frx":0010
         Left            =   1320
         List            =   "Frm_DireccionRep.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cboCiudad 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Txt_Fono 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lbltipotelefono 
         Caption         =   "Tipo Telefono : "
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblciudad 
         Caption         =   "Ciudad : "
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblnumerotelf 
         Caption         =   "Numero :"
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
   End
End
Attribute VB_Name = "Frm_DireccionRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodLoad As Integer




Private Sub cboTipoTelefono_Click()
Dim valor As String
valor = fgObtenerCodigo_TextoCompuesto(cboTipoTelefono)
If valor <> "2" Then
 cboCiudad.Visible = True
 lblciudad.Visible = True
 Else
 cboCiudad.Visible = False
 lblciudad.Visible = False
End If
End Sub

Private Sub Cmd_BuscarDir_Click()
    On Error GoTo Err_Buscar
  Frm_BusDireccion.flInicio ("Frm_DireccionRep")
Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
 
End Sub
Function flRecibeDireccionEdit(iNomDepartamento As String, iNomProvincia As String, iNomDistrito As String, iCodDir As String)
'FUNCION QUE RECIBE LOS DATOS DEL FORMULARIO DE BUSQUEDA de Dirección
 Lbl_DepartamentoEdit = Trim(iNomDepartamento)
    Lbl_ProvinciaEdit = Trim(iNomProvincia)
    Lbl_DistritoEdit = Trim(iNomDistrito)
       
    
    DirRep.vCodDireccion = iCodDir
    DirRep.vNomDepartamento = Trim(iNomDepartamento)
    DirRep.vNomProvincia = Trim(iNomProvincia)
    DirRep.vNomDistrito = Trim(iNomDistrito)
   
   
    
    Frm_DireccionRep.Enabled = True
    Frm_DireccionRep.Refresh
    
End Function

Private Sub Command1_Click()

 Call GetVariables
If DirRep.vTipoTelefono = "2" Then
 DirRep.vCodigoTelefono = ""
End If
Call Frm_CalPoliza.RecibeDireccionRepr(DirRep)
Unload Me
End Sub
Private Sub ConcatenarDireccionRep()
Dim vlRegistroDir As ADODB.Recordset
On Error GoTo Err_Buscar
 Dim tipo_via As String
 Dim tipo_bloque As String
 Dim tipo_interior As String
 Dim tipo_cjht As String
 Dim vDireccionConcat As String
 
Call fgBuscarNombreComunaProvinciaRegion(DirRep.vCodDireccion)

     vgSql = "SELECT"
     vgSql = vgSql + "(SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_VIA T "
     vgSql = vgSql + " WHERE T.COD_DIRE_VIA = '" + DirRep.vTipoVia + "') as TIPO_VIA,"
     vgSql = vgSql + " (SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_BLOQUE T"
     vgSql = vgSql + " WHERE T.COD_BLOCKCHALET =  '" + DirRep.vTipoBlock + "') AS TIPO_BLOQUE,"
     vgSql = vgSql + " (SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_INTERIOR T"
     vgSql = vgSql + " WHERE T.COD_INTERIOR = '" + DirRep.vTipoPref + "') AS TIPO_INTERIOR,"
     vgSql = vgSql + "(SELECT TRIM(GLS_DESCRIPCION)"
     vgSql = vgSql + " FROM MA_TPAR_TIPO_CJHT T"
     vgSql = vgSql + " WHERE T.COD_CJHT = '" + DirRep.vTipoConj + "' ) AS TIPO_CJHT"
     vgSql = vgSql + " FROM DUAL"
     Set vlRegistroDir = vgConexionBD.Execute(vgSql)
     If Not vlRegistroDir.EOF Then
        tipo_via = IIf(IsNull(vlRegistroDir!tipo_via), "", vlRegistroDir!tipo_via)
        tipo_bloque = IIf(IsNull(vlRegistroDir!tipo_bloque), "", vlRegistroDir!tipo_bloque)
        tipo_interior = IIf(IsNull(vlRegistroDir!tipo_interior), "", vlRegistroDir!tipo_interior)
        tipo_cjht = IIf(IsNull(vlRegistroDir!tipo_cjht), "", vlRegistroDir!tipo_cjht)
     End If
     vlRegistroDir.Close
     Dim Strmanzana As String
     If Trim(DirRep.vManzana) <> "" Then
       Strmanzana = "Manzana " & DirRep.vManzana & " "
     Else
       Strmanzana = ""
     End If
     
     Dim StrLote As String
     If Trim(DirRep.vLote) <> "" Then
       StrLote = "Lote " & DirRep.vLote & " "
     Else
       StrLote = ""
     End If
     Dim StrEtapa As String
     If Trim(DirRep.vEtapa) <> "" Then
       StrEtapa = "Etapa " & DirRep.vEtapa & " "
     Else
       StrEtapa = ""
     End If
     Dim strBloque As String
     If Trim(tipo_bloque) <> "" And Trim(DirRep.vNumBlock) <> "" Then
     strBloque = tipo_bloque & " " & Trim(DirRep.vNumBlock) & " "
     Else
     strBloque = ""
     End If
     Dim StrInterior As String
     If Trim(tipo_interior) <> "" And Trim(DirRep.vInterior) <> "" Then
     StrInterior = tipo_interior & " " & Trim(DirRep.vInterior) & " "
     Else
     StrInterior = ""
     End If
     Dim StrCjht As String
     If DirRep.vTipoConj = "99" Then
      StrCjht = DirRep.vTipoConj & " "
     Else
        If Trim(DirRep.vTipoConj) <> "" And Trim(DirRep.vConjHabit) <> "" Then
        StrCjht = tipo_cjht & " " & DirRep.vConjHabit & " "
        Else
        StrCjht = ""
        End If
     End If
     Dim StrDireccion As String
     If DirRep.vTipoVia = "99" Or DirRep.vTipoVia = "88" Then
     StrDireccion = "" & DirRep.vDireccion
     
    Else
        If Trim(tipo_via) <> "" And Trim(DirRep.vDireccion) <> "" Then
        StrDireccion = tipo_via & " " & DirRep.vDireccion
        Else
        StrDireccion = ""
        End If
     
     End If
     
    vDireccionConcat = Trim(UCase(StrDireccion & " " & IIf(Trim(DirRep.vNumero) = "", "", Trim(DirRep.vNumero) & " ") _
        & strBloque & StrInterior _
        & StrCjht _
        & StrEtapa & Strmanzana & StrLote _
        & " " & IIf(Trim(DirRep.vReferencia) = "", "", Trim(DirRep.vReferencia) & " ") _
        & IIf(Trim(DirRep.vNomDepartamento) = "", "", Trim(DirRep.vNomDepartamento) & " ") _
        & IIf(Trim(DirRep.vNomProvincia) = "", "", Trim(DirRep.vNomProvincia) & " ") _
        & IIf(Trim(DirRep.vNomDistrito) = "", "", Trim(DirRep.vNomDistrito) & " ")))
        
      
      
      DirRep.vgls_desdirebusq = vDireccionConcat

        
 Exit Sub
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
Private Sub Form_Load()

    Call limpiar
    Call llenarTipoInterior(cbodepart)
    Call llenarTipoBloque(cbobloque)
    Call llenarTipoConj(cboconj)
    Call llenarTipoVia(cbotipovia)
    Call llenarcomboTipoTelefono(cboTipoTelefono)
    Call llenarcomboTipoTelefono(cboTipoTelefono2)
    Call llenarCodigoTelefono(cboCiudad)
    Call setVariables
    If DirRep.vTipoTelefono = "" Then
    Call llenarComboValue(cboTipoTelefono, "4")
    End If
End Sub
Private Sub GetVariables()
    DirRep.vTipoTelefono = UCase(CStr(fgObtenerCodigo_TextoCompuesto(cboTipoTelefono.Text)))
    DirRep.vTipoTelefono2 = UCase(CStr(fgObtenerCodigo_TextoCompuesto(cboTipoTelefono2.Text)))
    DirRep.vNumTelefono = UCase(Txt_Fono.Text)
    DirRep.vNumTelefono2 = UCase(Txt_celular.Text)
    DirRep.vCodigoTelefono = UCase(fgObtenerCodigo_TextoCompuesto(cboCiudad.Text))
    DirRep.vCodigoTelefono2 = UCase(fgObtenerCodigo_TextoCompuesto(cboCiudad2.Text))
    DirRep.vTipoVia = UCase(fgObtenerCodigo_TextoCompuesto(cbotipovia.Text))
    DirRep.vDireccion = UCase(txtdireccion.Text)
    DirRep.vNumero = UCase(txtnumero.Text)
    DirRep.vTipoPref = UCase(fgObtenerCodigo_TextoCompuesto(cbodepart.Text))
    DirRep.vInterior = UCase(Txt_Interior.Text)
    DirRep.vManzana = UCase(txtmanzana.Text)
    DirRep.vLote = UCase(txtlote.Text)
    DirRep.vEtapa = UCase(txtetapa.Text)
    DirRep.vTipoConj = UCase(fgObtenerCodigo_TextoCompuesto(cboconj.Text))
    DirRep.vConjHabit = UCase(txtnombreconj.Text)
    DirRep.vTipoBlock = UCase(fgObtenerCodigo_TextoCompuesto(cbobloque.Text))
    DirRep.vNumBlock = UCase(txtbloque.Text)
    DirRep.vReferencia = UCase(Txt_Referencia.Text)
    Call ConcatenarDireccionRep


End Sub

Private Sub setVariables()
vCodLoad = 1
Txt_Fono.Text = DirRep.vNumTelefono
txtdireccion.Text = DirRep.vDireccion
txtnumero.Text = DirRep.vNumero
Txt_Interior.Text = DirRep.vInterior
txtmanzana.Text = DirRep.vManzana
txtlote.Text = DirRep.vLote
txtetapa.Text = DirRep.vEtapa
txtnombreconj.Text = DirRep.vConjHabit
txtbloque.Text = DirRep.vNumBlock
Txt_Referencia.Text = DirRep.vReferencia
Txt_celular.Text = DirRep.vNumTelefono2
txtdirec.Text = DirRep.vgls_desdirebusq
If DirRep.vCodDireccion <> "" Then
    Call fgBuscarNombreComunaProvinciaRegion(DirRep.vCodDireccion)
End If
DirRep.vTipoTelefono = "4"
DirRep.vTipoTelefono2 = "2"
Call llenarComboValue(cboTipoTelefono, DirRep.vTipoTelefono)
Call llenarComboValue(cboTipoTelefono2, DirRep.vTipoTelefono2)
Call llenarComboValue(cboCiudad, DirRep.vCodigoTelefono)
Call llenarComboValue(cbotipovia, DirRep.vTipoVia)
Call llenarComboValue(cbodepart, DirRep.vTipoPref)
Call llenarComboValue(cboconj, DirRep.vTipoConj)
Call llenarComboValue(cbobloque, DirRep.vTipoBlock)
vCodLoad = 0
End Sub

Private Sub llenarComboValue(combo As ComboBox, Value As String)
If Value <> "" Then
Call fgBuscaPos(combo, Value)
Else
combo.ListIndex = -1
End If
End Sub
Public Sub llenarCodigoTelefono(icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    icombo.Clear
    
    icombo.AddItem ""
  
    vgSql = "SELECT COD_AREA,GLS_REGION FROM MA_TPAR_TIPO_AREA MTTA , MA_TPAR_REGION MTR "
    vgSql = vgSql & "Where MTTA.COD_REGION = MTR.COD_REGION ORDER BY COD_AREA ASC"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!COD_AREA) & " - " & Trim((vlRsCombo!gls_region))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        icombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If icombo.ListCount <> 0 Then
        icombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub
Public Sub llenarcomboTipoTelefono(icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    icombo.Clear
    
    icombo.AddItem ""
    vgSql = "select cod_tipo_telefono,gls_descripcion  from ma_tpar_tipo_telefono where "
    vgSql = vgSql & "cod_estado = 1"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!COD_TIPO_TELEFONO) & " - " & Trim((vlRsCombo!gls_descripcion))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        icombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If icombo.ListCount <> 0 Then
        icombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Sub
Public Sub llenarTipoVia(icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    icombo.Clear
    icombo.AddItem ""
    vgSql = "select COD_DIRE_VIA ,GLS_DESCRIPCION  from MA_TPAR_TIPO_VIA where "
    vgSql = vgSql & "COD_ESTADO = 1 "
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!cod_dire_via) & " - " & Trim((vlRsCombo!gls_descripcion))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        icombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If icombo.ListCount <> 0 Then
        icombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Public Sub llenarTipoConj(icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    icombo.Clear
    
    icombo.AddItem ""
    vgSql = "select COD_CJHT ,GLS_DESCRIPCION  from MA_TPAR_TIPO_CJHT where "
    vgSql = vgSql & "COD_ESTADO = 1 AND FLG_USO_RV = 1"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!cod_cjht) & " - " & Trim((vlRsCombo!gls_descripcion))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        icombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If icombo.ListCount <> 0 Then
        icombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
Public Sub llenarTipoBloque(icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    icombo.Clear
    
    icombo.AddItem ""
    vgSql = "select cod_blockchalet ,gls_descripcion  from MA_TPAR_TIPO_BLOQUE where "
    vgSql = vgSql & "cod_Estado = 1"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!cod_blockchalet) & " - " & Trim((vlRsCombo!gls_descripcion))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        icombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If icombo.ListCount <> 0 Then
        icombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
Private Sub limpiar()
    Txt_Fono.Text = ""
    txtdireccion.Text = ""
    txtnumero.Text = ""
    Txt_Interior.Text = ""
    txtmanzana.Text = ""
    txtlote.Text = ""
    txtetapa.Text = ""
    txtnombreconj.Text = ""
    txtbloque.Text = ""
    Txt_Referencia.Text = ""
    Txt_celular.Text = ""
End Sub
Public Sub llenarTipoInterior(icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    icombo.Clear
    
    icombo.AddItem ""
    vgSql = "select cod_interior,gls_descripcion as  from ma_tpar_tipo_interior where "
    vgSql = vgSql & "cod_estado = 1"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!cod_interior) & " - " & Trim((vlRsCombo!gls_descripcion))
        'iCombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
        icombo.AddItem vlDescripcion
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    
    If icombo.ListCount <> 0 Then
        icombo.ListIndex = 0
    End If

Exit Sub
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub
Function fgBuscarNombreComunaProvinciaRegion(vlCodDir As String)
Dim vlRegistroDir As ADODB.Recordset
On Error GoTo Err_Buscar

     vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,c.Cod_Comuna,c.Gls_Comuna"
     vgSql = vgSql & " FROM MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r"
     vgSql = vgSql & " Where c.Cod_Direccion = '" & DirRep.vCodDireccion & "' and  "
     vgSql = vgSql & " c.cod_region = p.cod_region and"
     vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
     vgSql = vgSql & " p.cod_region = r.cod_region"
     Set vlRegistroDir = vgConexionBD.Execute(vgSql)
     If Not vlRegistroDir.EOF Then
        Lbl_DepartamentoEdit = (vlRegistroDir!gls_region)
        Lbl_ProvinciaEdit = (vlRegistroDir!gls_provincia)
        Lbl_DistritoEdit = (vlRegistroDir!gls_comuna)
        
        DirRep.vNomDepartamento = (vlRegistroDir!gls_region)
        DirRep.vNomProvincia = vlRegistroDir!gls_provincia
        DirRep.vNomDistrito = vlRegistroDir!gls_comuna
        
        
        
     End If
     vlRegistroDir.Close

Exit Function
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
