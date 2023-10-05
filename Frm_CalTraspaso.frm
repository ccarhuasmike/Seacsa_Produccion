VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_CalTraspaso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Pólizas recibidas a Pago de Pensiones."
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9735
   Begin VB.CommandButton Cmd_RestarTodos 
      Caption         =   "<<"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Cmd_AgregarTodos 
      Caption         =   ">>"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   3840
      Width           =   615
   End
   Begin VB.Frame Fra_Formulario 
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   9495
      Begin VB.Label Lbl_ClickGrilla 
         Caption         =   "Label1"
         Height          =   255
         Left            =   6720
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Lbl_FechaOpera 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Efectiva del Traspaso  :"
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
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   9495
      Begin VB.CommandButton Cmd_Grabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   3645
         Picture         =   "Frm_CalTraspaso.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Generar el Traspaso de Póliza"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5040
         Picture         =   "Frm_CalTraspaso.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir del Formulario"
         Top             =   240
         Width           =   720
      End
      Begin Crystal.CrystalReport Rpt_Poliza 
         Left            =   720
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.CommandButton Cmd_Restar 
      Caption         =   "<"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Cmd_Agregar 
      Caption         =   ">"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   2640
      Width           =   615
   End
   Begin VB.Frame Fra_Formulario 
      Caption         =   "  Pólizas Traspasadas a Pago de Pensiones  "
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
      Height          =   4815
      Index           =   2
      Left            =   5400
      TabIndex        =   7
      Top             =   720
      Width           =   4215
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaTraspasadas 
         Height          =   4455
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   12632256
         BackColorBkg    =   -2147483632
         GridColor       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Fra_Formulario 
      Caption         =   "  Pólizas Recibidas  "
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
      Height          =   4815
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4215
      Begin MSFlexGridLib.MSFlexGrid Msf_GrillaRecibidas 
         Height          =   4455
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   12632256
         BackColorBkg    =   -2147483632
         GridColor       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
      End
   End
End
Attribute VB_Name = "Frm_CalTraspaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vlRegistro As ADODB.Recordset
Dim vlPos As Integer
Dim vlNumPoliza As String
Dim vlFecTraspaso As String
Dim vlPrimaRec As Double
Dim vlFecTerPagoPenGar As String
Dim vlAnno As String
Dim vlMes As String
Dim vlDia As String
Dim vlEstadoTrans As Boolean
Dim vlGlsUsuarioCrea As Variant
Dim vlFecCrea As Variant
Dim vlHorCrea As Variant
Dim vlGlsUsuarioModi As Variant
Dim vlFecModi As Variant
Dim vlHorModi As Variant
Dim vlSaludCod As String, vlSaludMto As Double, vlSaludMod As String
Dim vlViaPagoCod As String, vlViaPagoSuc As String, vlViaPagoTC As String
Dim vlViaPagoBco As String, vlViaPagoNumCta As String
Dim vlNumCtaCCI As String
Dim vlMoncta As String
Dim vlSw As Boolean

'Variables y Constantes para determinar derecho a pensión del beneficiario a crear.
Dim vlCodTipPension As String
Dim vlCodEstPension As String
Const clCodTipPensionSob As String * 2 = "08"
Const clCodParCausus As String * 2 = "99"
Const clCodSinDerPen As String * 2 = "10"
Const clCodConDerPen As String * 2 = "99"

Const clCodEstado6 As String * 1 = "6"
Const clFechaTopeTer As String * 8 = "99991231"
Const clTasaCtoRea0 As Integer = 0
Const clModalPagoSalud As String = "PORCE"
Const clAsignarPagoPensionDefCia As Boolean = True
Dim vlFactorEsc As Double  'RRR 05/08/2016

Function flInicializaGrilla(Msf_Grilla As MSFlexGrid)
'Permite Limpiar e inicializar la grilla que contendra las polizas recibidas

On Error GoTo Err_flInicializaGrilla

    Msf_Grilla.Clear
    Msf_Grilla.Cols = 3
    Msf_Grilla.rows = 1
    Msf_Grilla.RowHeight(0) = 250
    Msf_Grilla.Row = 0

    Msf_Grilla.Col = 0
    Msf_Grilla.Text = "Nº Poliza"
    Msf_Grilla.ColWidth(0) = 1200
    Msf_Grilla.ColAlignment(0) = 1  'centrado
    
    Msf_Grilla.Col = 1
    Msf_Grilla.Text = "Fecha Trasp."
    Msf_Grilla.ColWidth(1) = 1200

    Msf_Grilla.Col = 2
    Msf_Grilla.Text = "Prima Recaudada"
    Msf_Grilla.ColWidth(2) = 1500
    
Exit Function
Err_flInicializaGrilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flCargaGrilla()
On Error GoTo Err_Carga
    
    vgSql = ""
    vgSql = "SELECT p.cod_trapagopen,p.num_poliza,r.fec_traspaso,r.mto_prirec "
    vgSql = vgSql & "FROM pd_tmae_poliza p, pd_tmae_polprirec r "
    vgSql = vgSql & "WHERE p.num_poliza = r.num_poliza "
    vgSql = vgSql & "ORDER BY p.num_poliza "
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
       Call flInicializaGrilla(Msf_GrillaRecibidas)
       While Not vgRs.EOF
          If Trim(vgRs!cod_trapagopen) = "N" Then
             Msf_GrillaRecibidas.AddItem CStr(Trim(vgRs!Num_Poliza)) & vbTab _
             & DateSerial(Mid((vgRs!fec_traspaso), 1, 4), Mid((vgRs!fec_traspaso), 5, 2), Mid((vgRs!fec_traspaso), 7, 2)) & vbTab _
             & Format((vgRs!MTO_PRIREC), "###,###,##0.00")
          End If
          vgRs.MoveNext
       Wend
    End If
    vgRs.Close

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flAgregarDetalle()
On Error GoTo Err_flAgregarDetalle

    vlPos = Msf_GrillaRecibidas.Row
    If vlPos = 0 Then
       MsgBox "No Existen Polizas para Traspasar", vbInformation, "Información"
       Exit Function
    End If
    If IsNumeric(Lbl_ClickGrilla.Caption) Then
    
       Msf_GrillaRecibidas.Col = 0
       vlNumPoliza = Msf_GrillaRecibidas.Text
       Msf_GrillaRecibidas.Col = 1
       vlFecTraspaso = Msf_GrillaRecibidas.Text
       Msf_GrillaRecibidas.Col = 2
       vlPrimaRec = Msf_GrillaRecibidas.Text
          
       Msf_GrillaTraspasadas.AddItem vlNumPoliza & vbTab _
       & vlFecTraspaso & vbTab _
       & Format(vlPrimaRec, "###,###,##0.00")
        
       If vlPos = (Msf_GrillaRecibidas.rows - 1) Then
          If Msf_GrillaRecibidas.Row = 1 Then
             Call flInicializaGrilla(Msf_GrillaRecibidas)
             Lbl_ClickGrilla.Caption = ""
             Exit Function
          End If
       End If
       
       Msf_GrillaRecibidas.RemoveItem vlPos
       Lbl_ClickGrilla.Caption = ""
    
   Else
       MsgBox "Debe Seleccionar Póliza a Traspasar", vbInformation, "Información"
   End If
    
Exit Function
Err_flAgregarDetalle:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flQuitarDetalle()

On Error GoTo Err_flQuitarDetalle

    vlPos = Msf_GrillaTraspasadas.Row
    If vlPos = 0 Then
       MsgBox "No Existen Polizas para Traspasar", vbInformation, "Información"
       Exit Function
    End If
    
    If IsNumeric(Lbl_ClickGrilla.Caption) Then
    
       Msf_GrillaTraspasadas.Col = 0
       vlNumPoliza = Msf_GrillaTraspasadas.Text
       Msf_GrillaTraspasadas.Col = 1
       vlFecTraspaso = Msf_GrillaTraspasadas.Text
       Msf_GrillaTraspasadas.Col = 2
       vlPrimaRec = Msf_GrillaTraspasadas.Text
           
       Msf_GrillaRecibidas.AddItem vlNumPoliza & vbTab _
       & vlFecTraspaso & vbTab _
       & Format(vlPrimaRec, "###,###,##0.00")
        
       If vlPos = (Msf_GrillaTraspasadas.rows - 1) Then
          If Msf_GrillaTraspasadas.Row = 1 Then
             Call flInicializaGrilla(Msf_GrillaTraspasadas)
             Lbl_ClickGrilla.Caption = ""
             Exit Function
          End If
       End If
       
       Msf_GrillaTraspasadas.RemoveItem vlPos
       Lbl_ClickGrilla.Caption = ""
   Else
       MsgBox "Debe Seleccionar Póliza a Traspasar", vbInformation, "Información"
   End If

Exit Function
Err_flQuitarDetalle:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function flMoverTodos(Msf_GrillaDesde As MSFlexGrid, Msf_GrillaHasta As MSFlexGrid)

On Error GoTo Err_flMoverTodos
             
    vlPos = Msf_GrillaDesde.Row
    If vlPos = 0 Then
       MsgBox "No Existen Polizas para Traspasar", vbInformation, "Información"
       Exit Function
    End If
    
    vlPos = 1
    Msf_GrillaDesde.Row = 1
    
    While Msf_GrillaDesde.Row <> 0
       
          Msf_GrillaDesde.Col = 0
          vlNumPoliza = Msf_GrillaDesde.Text
          Msf_GrillaDesde.Col = 1
          vlFecTraspaso = Msf_GrillaDesde.Text
          Msf_GrillaDesde.Col = 2
          vlPrimaRec = Msf_GrillaDesde.Text
       
          Msf_GrillaHasta.AddItem vlNumPoliza & vbTab _
          & vlFecTraspaso & vbTab _
          & Format(vlPrimaRec, "#,#0.00")
    
          If vlPos = (Msf_GrillaDesde.rows - 1) Then
             If Msf_GrillaDesde.Row = 1 Then
                Call flInicializaGrilla(Msf_GrillaDesde)
                Lbl_ClickGrilla.Caption = ""
                Exit Function
             End If
          End If
          
          Msf_GrillaDesde.RemoveItem vlPos
        
    Wend
    Lbl_ClickGrilla.Caption = ""
    
    
Exit Function
Err_flMoverTodos:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function flInsertarBeneficiario()

    vlGlsUsuarioCrea = vgUsuario
    vlFecCrea = Format(Date, "yyyymmdd")
    vlHorCrea = Format(Time, "hhmmss")

    vlSaludCod = ""
    vlSaludMto = 0
    vlViaPagoCod = ""
    vlViaPagoSuc = ""
    vlViaPagoTC = ""
    vlViaPagoBco = ""
    vlViaPagoNumCta = ""
    vlNumCtaCCI = ""
    
   
    
    'salud
    Sql = ""
    Sql = "SELECT cod_inssalud,cod_modsalud,cod_viapago,"
    Sql = Sql & "cod_tipcuenta, cod_banco, num_cuenta,cod_sucursal "
    Sql = Sql & "FROM ma_tcod_general"
    Set vgRs4 = vgConexionBD.Execute(Sql)
    If Not vgRs4.EOF Then
        If (vgRs!cod_isapre = "00") Then
            vlSaludCod = vgRs4!Cod_InsSalud
        Else
            vlSaludCod = vgRs!cod_isapre
        End If
        vlSaludMod = vgRs4!Cod_ModSalud
        
        'se saca el mto de salud
        Sql = ""
        Sql = "SELECT mto_elemento FROM ma_tpar_tabcodvig "
        Sql = Sql & "WHERE cod_tabla= '" & vgCodTabla_PrcSal & "' AND "
        Sql = Sql & "cod_elemento = 'PSM' AND "
        Sql = Sql & "fec_inivig <= '" & (vgRs!Fec_Vigencia) & "' AND "
        Sql = Sql & "fec_tervig >= '" & (vgRs!Fec_Vigencia) & "'"
        Set vlRegistro = vgConexionBD.Execute(Sql)
        If Not vlRegistro.EOF Then
            vlSaludMto = Format(vlRegistro!mto_elemento, "#0.00")
        Else
            vlSaludMto = 0
        End If
        vlRegistro.Close
        
        'via pago
        'El Usuario indistintamente quiere pasar el Tipo de Pago de Transferencia AFP - 04
        'como el Tipo de Pago estándar, por ende se deberá enviar en la Sucursal el Código de la AFP del Afiliado
        If (clAsignarPagoPensionDefCia = True) Then
        
             'INICIO GCP-FRACTAL 04042019
   
'            vlViaPagoCod = "04"
'            vlViaPagoSuc = vgRs!Cod_AFP 'En la sucursal asignar la AFP de la Póliza
'            vlViaPagoTC = "00"
'            vlViaPagoBco = "00"
'            vlViaPagoNumCta = ""

            If vgRs2!Num_Orden = "1" Then
                Sql = ""
                Sql = "SELECT COD_VIAPAGO, COD_SUCURSAL, COD_TIPCUENTA, COD_BANCO, NUM_CUENTA, NUM_CUENTA_CCI, COD_MONCTA FROM PD_TMAE_POLIZA WHERE NUM_POLIZA='" & vgRs2!Num_Poliza & "'"
                Set vlRegistro = vgConexionBD.Execute(Sql)
                If Not vlRegistro.EOF Then
                    vlViaPagoCod = vlRegistro!Cod_ViaPago
                    vlViaPagoSuc = vlRegistro!Cod_Sucursal
                    vlViaPagoTC = IIf(IsNull(vlRegistro!Cod_TipCuenta), "", vlRegistro!Cod_TipCuenta)
                    vlViaPagoBco = IIf(IsNull(vlRegistro!Cod_Banco), "", vlRegistro!Cod_Banco)
                    vlViaPagoNumCta = IIf(IsNull(vlRegistro!Num_Cuenta), "", vlRegistro!Num_Cuenta)
                    vlNumCtaCCI = IIf(IsNull(vlRegistro!num_cuenta_cci), "", vlRegistro!num_cuenta_cci)
                    vlMoncta = IIf(IsNull(vlRegistro!COD_MONCTA), "", vlRegistro!COD_MONCTA)
                Else
                    vlViaPagoCod = vgRs!Cod_ViaPago
                    vlViaPagoSuc = vgRs!Cod_Sucursal
                    vlViaPagoTC = IIf(IsNull(vgRs2!cod_tipcta), "", vgRs2!cod_tipcta)
                    vlViaPagoBco = IIf(IsNull(vgRs2!Cod_Banco), "", vgRs2!Cod_Banco)
                    vlViaPagoNumCta = IIf(IsNull(vgRs2!num_ctabco), "", vgRs2!num_ctabco)
                    vlNumCtaCCI = IIf(IsNull(vgRs2!num_cuenta_cci), "", vgRs2!num_cuenta_cci)
                    vlMoncta = IIf(IsNull(vgRs2!cod_monbco), "", vgRs2!cod_monbco)
                End If
                vlRegistro.Close
            Else
                    vlViaPagoCod = vgRs!Cod_ViaPago
                    vlViaPagoSuc = vgRs!Cod_Sucursal
                    vlViaPagoTC = IIf(IsNull(vgRs2!cod_tipcta), "", vgRs2!cod_tipcta)
                    vlViaPagoBco = IIf(IsNull(vgRs2!Cod_Banco), "", vgRs2!Cod_Banco)
                    vlViaPagoNumCta = IIf(IsNull(vgRs2!num_ctabco), "", vgRs2!num_ctabco)
                    vlNumCtaCCI = IIf(IsNull(vgRs2!num_cuenta_cci), "", vgRs2!num_cuenta_cci)
                    vlMoncta = IIf(IsNull(vgRs2!cod_monbco), "", vgRs2!cod_monbco)
            End If
            
        
            
            
            'FIN GCP-FRACTAL 04042019
            
                     
        Else
            If vgRs2!Cod_Par = "99" Then
            'Datos por defecto de la tabla ma_tcod_general
                If vgRs!Cod_ViaPago = "00" Then
                    vlViaPagoCod = vgRs4!Cod_ViaPago
                    vlViaPagoSuc = vgRs4!Cod_Sucursal
                    vlViaPagoTC = vgRs4!Cod_TipCuenta
                    vlViaPagoBco = vgRs4!Cod_Banco
                    If Not IsNull(vgRs4!Num_Cuenta) Then
                        vlViaPagoNumCta = vgRs4!Num_Cuenta
                    End If
              
                    
                Else
                    vlViaPagoCod = vgRs!Cod_ViaPago
                    vlViaPagoSuc = vgRs!Cod_Sucursal
                    vlViaPagoTC = vgRs!Cod_TipCuenta
                    vlViaPagoBco = vgRs!Cod_Banco
                    
                    If Not IsNull(vgRs!Num_Cuenta) Then
                        vlViaPagoNumCta = vgRs!Num_Cuenta
                    End If
                End If
            Else
           
                 vlViaPagoCod = vgRs4!Cod_ViaPago
                    vlViaPagoSuc = vgRs4!Cod_Sucursal
                    vlViaPagoTC = vgRs4!Cod_TipCuenta
                    vlViaPagoBco = vgRs4!Cod_Banco
                    If Not IsNull(vgRs4!Num_Cuenta) Then
                        vlViaPagoNumCta = vgRs4!Num_Cuenta
                    End If
              
                
      
            End If
        End If
    End If
    vgRs4.Close
    
'    'Permite Determinar el Derecho a Pensión o No del Beneficiario
'    If vlCodTipPension = clCodTipPensionSob Then
'       If (vgRs2!Cod_Par) = clCodParCausus Then
'          vlCodEstPension = clCodSinDerPen
'       Else
'           vlCodEstPension = clCodConDerPen
'       End If
'    Else
'        If (vgRs2!Cod_Par) = clCodParCausus Then
'           vlCodEstPension = clCodConDerPen
'        Else
'            vlCodEstPension = clCodSinDerPen
'        End If
'    End If
    'Solo se debe traspasar, ya que ya fue calculado para el Primer Pago
    If Not IsNull(vgRs2!Cod_EstPension) Then
        vlCodEstPension = vgRs2!Cod_EstPension
    Else
        vlCodEstPension = clCodSinDerPen
    End If
    
    Sql = ""
    Sql = "INSERT INTO pp_tmae_ben "
    Sql = Sql & " (num_poliza,num_endoso,num_orden,fec_ingreso, "
    Sql = Sql & " gls_nomben,gls_patben,gls_matben,gls_dirben, "
    Sql = Sql & " cod_direccion,gls_fonoben,gls_correoben,cod_grufam, "
    Sql = Sql & " cod_par,cod_sexo,cod_sitinv,cod_dercre,cod_derpen, "
    Sql = Sql & " cod_cauinv,fec_nacben,fec_nachm,fec_invben,cod_motreqpen, "
    Sql = Sql & " mto_pension,mto_pensiongar,prc_pension,cod_inssalud, "
    Sql = Sql & " cod_modsalud,mto_plansalud,cod_estpension, "
    Sql = Sql & " cod_viapago,cod_banco,cod_tipcuenta,num_cuenta,cod_sucursal, "
    Sql = Sql & " fec_fallben,cod_caususben,fec_susben,fec_inipagopen,"
    Sql = Sql & " fec_terpagopengar,"
    Sql = Sql & " cod_tipoidenben,gls_nomsegben,num_idenben , "
    Sql = Sql & " cod_usuariocrea,fec_crea,hor_crea "
    Sql = Sql & " ,prc_pensionleg "
    Sql = Sql & " ,prc_pensiongar "
    Sql = Sql & " ,ind_bolelec "
   'INICIO GCP-FRACTAL 04042019
    Sql = Sql & " ,NUM_CUENTA_CCI "
    Sql = Sql & " ,cod_tipcta, cod_monbco, num_ctabco "
    Sql = Sql & " ,gls_telben2 "
    'Inicio 02/10/2023-SMCCB
    Sql = Sql & " ,COD_MODTIPOCUENTA_MANC , COD_TIPODOC_MANC, NUM_DOC_MANC, NOMBRE_MANC, APELLIDO_MANC "
    'Fin 02/10/2023-SMCCB
     Sql = Sql & " ) VALUES ( "
    Sql = Sql & "'" & Trim(vgRs2!Num_Poliza) & "' , "
    Sql = Sql & " " & cgNumeroEndosoInicial & ", " '1
    Sql = Sql & " " & str(vgRs2!Num_Orden) & ","
    Sql = Sql & "'" & Format(CDate(Trim(Lbl_FechaOpera.Caption)), "yyyymmdd") & "',"
    Sql = Sql & "'" & (vgRs2!Gls_NomBen) & "',"
    Sql = Sql & "'" & (vgRs2!Gls_PatBen) & "',"
    If IsNull(vgRs2!Gls_MatBen) Then
       Sql = Sql & " NULL ,"
    Else
        Sql = Sql & "'" & (vgRs2!Gls_MatBen) & "',"
    End If
    Sql = Sql & "'" & (vgRs!Gls_Direccion) & "',"
    Sql = Sql & " " & str(vgRs!Cod_Direccion) & ","
    Sql = Sql & "'" & (vgRs2!GLS_FONO) & "',"
    Sql = Sql & "'" & (vgRs2!GLS_CORREO) & "',"
    Sql = Sql & "'" & (vgRs2!Cod_GruFam) & "',"
    Sql = Sql & "'" & (vgRs2!Cod_Par) & "',"
    Sql = Sql & "'" & (vgRs2!Cod_Sexo) & "',"
    Sql = Sql & "'" & (vgRs2!Cod_SitInv) & "',"
    Sql = Sql & "'" & (vgRs2!Cod_DerCre) & "',"
    Sql = Sql & "'" & (vgRs2!Cod_DerPen) & "',"
    Sql = Sql & "'" & (vgRs2!Cod_CauInv) & "',"
    Sql = Sql & "'" & (vgRs2!Fec_NacBen) & "',"
    If IsNull(vgRs2!Fec_NacHM) Then
       Sql = Sql & " NULL ,"
    Else
        Sql = Sql & "'" & (vgRs2!Fec_NacHM) & "',"
    End If
    If IsNull(vgRs2!Fec_InvBen) Then
       Sql = Sql & " NULL ,"
    Else
        Sql = Sql & "'" & (vgRs2!Fec_InvBen) & "',"
    End If
    Sql = Sql & "'1',"
    Sql = Sql & " " & str(vgRs2!Mto_Pension) & ","
'I--- ABV 19/10/2007 ---
    Sql = Sql & " " & str(vgRs2!Mto_PensionGar) & ","
    
    'Sql = Sql & " 0,"
'F--- ABV 19/10/2007 ---
    Sql = Sql & " " & str(vgRs2!Prc_Pension) & ","
    Sql = Sql & "'" & vlSaludCod & "',"
    Sql = Sql & "'" & vlSaludMod & "',"
    Sql = Sql & " " & str(vlSaludMto) & ","
    Sql = Sql & "'" & Trim(vlCodEstPension) & "',"
    Sql = Sql & "'" & vlViaPagoCod & "',"
    Sql = Sql & "'" & vlViaPagoBco & "',"
    Sql = Sql & "'" & vlViaPagoTC & "',"
    If vlViaPagoNumCta <> "" Then
        Sql = Sql & "'" & vlViaPagoNumCta & "',"
    Else
        Sql = Sql & " NULL ,"
    End If
    Sql = Sql & "'" & vlViaPagoSuc & "',"
    If IsNull(vgRs2!Fec_FallBen) Then
       Sql = Sql & " NULL ,"
    Else
        Sql = Sql & "'" & (vgRs2!Fec_FallBen) & "',"
    End If
    Sql = Sql & " NULL ,"
    Sql = Sql & " NULL ,"
    'Sql = Sql & "'" & (vgRs!Fec_IniPagoPen) & "',"
    Sql = Sql & "'" & (vgRs!fec_inipencia) & "'," 'hqr 08/09/2007 Debe ser la fecha de inicio de pagos al beneficiario
'I--- ABV 19/10/2007 ---
'    If (vlFecTerPagoPenGar <> "") Then
'        Sql = Sql & "'" & Trim(vlFecTerPagoPenGar) & "',"
'    Else
        Sql = Sql & " NULL ,"
'    End If
'F--- ABV 19/10/2007 ---
    Sql = Sql & " " & (vgRs2!Cod_TipoIdenBen) & ","
    If IsNull(vgRs2!Gls_NomSegBen) Then
        Sql = Sql & " NULL ,"
    Else
        Sql = Sql & "'" & (vgRs2!Gls_NomSegBen) & "',"
    End If
    Sql = Sql & "'" & (vgRs2!Num_IdenBen) & "',"
    Sql = Sql & "'" & vlGlsUsuarioCrea & "',"
    Sql = Sql & "'" & vlFecCrea & "',"
    Sql = Sql & "'" & vlHorCrea & "' "
    Sql = Sql & "," & str(vgRs2!Prc_PensionLeg) & " "
'I--- ABV 19/10/2007 ---
    Sql = Sql & "," & str(vgRs2!Prc_PensionGar) & " "
    'Sql = Sql & ",0 "
'F--- ABV 19/10/2007 ---
'mvg 20170904
    Sql = Sql & ",'" & IIf(IsNull(vgRs2!ind_bolelec), "N", vgRs2!ind_bolelec) & "' "
    'INICIO GCP-FRACTAL 04042019
    Sql = Sql & ",'" & vlNumCtaCCI & "' "
    Sql = Sql & ",'" & vlViaPagoTC & "' "
    If vlMoncta <> "" Then
       Sql = Sql & ", '" & vlMoncta & "' "
    Else
        Sql = Sql & ", NULL "
    End If
     If vlViaPagoNumCta <> "" Then
        Sql = Sql & ", '" & vlViaPagoNumCta & "'"
    Else
        Sql = Sql & ", NULL "
    End If
   'FIN GCP-FRACTAL 04042019
   Sql = Sql & ",'" & (vgRs2!gls_fono2) & "'"
   'Inicio 02/10/2023-SMCCB
    Sql = Sql & ",'" & (vgRs2!COD_MODTIPOCUENTA_MANC) & "'"
    Sql = Sql & ",'" & (vgRs2!COD_TIPODOC_MANC) & "'"
    Sql = Sql & ",'" & (vgRs2!NUM_DOC_MANC) & "'"
    Sql = Sql & ",'" & (vgRs2!NOMBRE_MANC) & "'"
    Sql = Sql & ",'" & (vgRs2!APELLIDO_MANC) & "'"
    'Inicio 02/10/2023-SMCCB
    Sql = Sql & ") "
    vgConectarBD.Execute Sql

  
    'hqr 14/09/2007 Se inserta certificado ficticio
    If vlCodEstPension <> clCodSinDerPen Then
        If vgRs2!Cod_Par >= 30 And vgRs2!Cod_Par <= 40 Then
            Sql = "INSERT INTO PP_TMAE_CERTIFICADO"
            Sql = Sql & "(NUM_POLIZA, NUM_ENDOSO, NUM_ORDEN, FEC_INICER, "
            Sql = Sql & "COD_TIPO, FEC_TERCER, COD_FRECUENCIA, GLS_NOMINSTITUCION, "
            Sql = Sql & "FEC_RECCIA, FEC_INGCIA, FEC_EFECTO, FEC_INIRETROPEN, "
            Sql = Sql & "FEC_TERRETROPEN, COD_USUARIOCREA, FEC_CREA, HOR_CREA, "
            Sql = Sql & "COD_USUARIOMODI, FEC_MODI, HOR_MODI, EST_ACT " ', COD_INDRELIQUIDAR, NUM_RELIQ"
            Sql = Sql & ")"
            Sql = Sql & "VALUES ("
            Sql = Sql & "'" & Trim(vgRs2!Num_Poliza) & "' , "
            Sql = Sql & " " & cgNumeroEndosoInicial & ", " '1
            Sql = Sql & " " & str(vgRs2!Num_Orden) & ","
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            'HQR 13/10/2007 Se deja la Fecha de termino del certificado de estudio como el último dia del mes
            'Sql = Sql & " 'SUP','" & Format(DateAdd("m", 6, DateSerial(Mid(vgRs!Fec_IniPagoPen, 1, 4), Mid(vgRs!Fec_IniPagoPen, 5, 2), Mid(vgRs!Fec_IniPagoPen, 7, 2))), "yyyymmdd") & "',"
            Sql = Sql & " 'SUP','" & Format(DateAdd("m", 6, DateSerial(Mid(vgRs!Fec_IniPagoPen, 1, 4), Mid(vgRs!Fec_IniPagoPen, 5, 2), Mid(vgRs!Fec_IniPagoPen, 7, 2) - 1)), "yyyymmdd") & "',"
            Sql = Sql & " 'S', 'Certificado Estudios Inicial',"
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            Sql = Sql & " NULL, NULL,"
            Sql = Sql & "'" & vlGlsUsuarioCrea & "',"
            Sql = Sql & "'" & vlFecCrea & "',"
            Sql = Sql & "'" & vlHorCrea & "', "
            Sql = Sql & " NULL, NULL, NULL, '0'"
            'Sql = Sql & ",NULL,NULL"
            Sql = Sql & ")"
            vgConectarBD.Execute Sql
            
            ''DEBE CREAR EL CERTIFICADO DE ESTUDIOS
            Sql = "INSERT INTO PP_TMAE_CERTIFICADO"
            Sql = Sql & "(NUM_POLIZA, NUM_ENDOSO, NUM_ORDEN, FEC_INICER, "
            Sql = Sql & "COD_TIPO, FEC_TERCER, COD_FRECUENCIA, GLS_NOMINSTITUCION, "
            Sql = Sql & "FEC_RECCIA, FEC_INGCIA, FEC_EFECTO, FEC_INIRETROPEN, "
            Sql = Sql & "FEC_TERRETROPEN, COD_USUARIOCREA, FEC_CREA, HOR_CREA, "
            Sql = Sql & "COD_USUARIOMODI, FEC_MODI, HOR_MODI, EST_ACT " ', COD_INDRELIQUIDAR, NUM_RELIQ"
            Sql = Sql & ")"
            Sql = Sql & "VALUES ("
            Sql = Sql & "'" & Trim(vgRs2!Num_Poliza) & "' , "
            Sql = Sql & " " & cgNumeroEndosoInicial & ", " '1
            Sql = Sql & " " & str(vgRs2!Num_Orden) & ","
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
                'HQR 13/10/2007 Se deja la Fecha de termino del certificado de estudio como el último dia del mes
                'Sql = Sql & " 'SUP','" & Format(DateAdd("m", 6, DateSerial(Mid(vgRs!Fec_IniPagoPen, 1, 4), Mid(vgRs!Fec_IniPagoPen, 5, 2), Mid(vgRs!Fec_IniPagoPen, 7, 2))), "yyyymmdd") & "',"
            Sql = Sql & " 'EST','" & Format(DateAdd("m", 6, DateSerial(Mid(vgRs!Fec_IniPagoPen, 1, 4), Mid(vgRs!Fec_IniPagoPen, 5, 2), Mid(vgRs!Fec_IniPagoPen, 7, 2) - 1)), "yyyymmdd") & "',"
            Sql = Sql & " 'S', 'Certificado de estudios inicial.',"
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            Sql = Sql & " NULL, NULL,"
            Sql = Sql & "'" & vlGlsUsuarioCrea & "',"
            Sql = Sql & "'" & vlFecCrea & "',"
            Sql = Sql & "'" & vlHorCrea & "', "
            Sql = Sql & " NULL, NULL, NULL, '1'"
                'Sql = Sql & ",NULL,NULL"
            Sql = Sql & ")"
            vgConectarBD.Execute Sql
 
        Else
             Sql = "INSERT INTO PP_TMAE_CERTIFICADO"
            Sql = Sql & "(NUM_POLIZA, NUM_ENDOSO, NUM_ORDEN, FEC_INICER, "
            Sql = Sql & "COD_TIPO, FEC_TERCER, COD_FRECUENCIA, GLS_NOMINSTITUCION, "
            Sql = Sql & "FEC_RECCIA, FEC_INGCIA, FEC_EFECTO, FEC_INIRETROPEN, "
            Sql = Sql & "FEC_TERRETROPEN, COD_USUARIOCREA, FEC_CREA, HOR_CREA, "
            Sql = Sql & "COD_USUARIOMODI, FEC_MODI, HOR_MODI, EST_ACT " ', COD_INDRELIQUIDAR, NUM_RELIQ"
            Sql = Sql & ")"
            Sql = Sql & "VALUES ("
            Sql = Sql & "'" & Trim(vgRs2!Num_Poliza) & "' , "
            Sql = Sql & " " & cgNumeroEndosoInicial & ", " '1
            Sql = Sql & " " & str(vgRs2!Num_Orden) & ","
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            'HQR 13/10/2007 Se deja la Fecha de termino del certificado de estudio como el último dia del mes
            'Sql = Sql & " 'SUP','" & Format(DateAdd("m", 6, DateSerial(Mid(vgRs!Fec_IniPagoPen, 1, 4), Mid(vgRs!Fec_IniPagoPen, 5, 2), Mid(vgRs!Fec_IniPagoPen, 7, 2))), "yyyymmdd") & "',"
            Sql = Sql & " 'SUP','" & Format(DateAdd("m", 6, DateSerial(Mid(vgRs!Fec_IniPagoPen, 1, 4), Mid(vgRs!Fec_IniPagoPen, 5, 2), Mid(vgRs!Fec_IniPagoPen, 7, 2) - 1)), "yyyymmdd") & "',"
            Sql = Sql & " 'S', 'Certificado de Supervivencia Inicial',"
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
            Sql = Sql & " NULL, NULL,"
            Sql = Sql & "'" & vlGlsUsuarioCrea & "',"
            Sql = Sql & "'" & vlFecCrea & "',"
            Sql = Sql & "'" & vlHorCrea & "', "
            Sql = Sql & " NULL, NULL, NULL, '1'"
            'Sql = Sql & ",NULL,NULL"
            Sql = Sql & ")"
            vgConectarBD.Execute Sql
        End If
    End If
    
       
'    ''DEBE CREAR EL CERTIFICADO DE ESTUDIOS
'
'    'If vlCodEstPension <> clCodSinDerPen Then
'    If vgRs2!Cod_Par >= 30 And vgRs2!Cod_Par <= 40 Then
'        Sql = "INSERT INTO PP_TMAE_CERTIFICADO"
'        Sql = Sql & "(NUM_POLIZA, NUM_ENDOSO, NUM_ORDEN, FEC_INICER, "
'        Sql = Sql & "COD_TIPO, FEC_TERCER, COD_FRECUENCIA, GLS_NOMINSTITUCION, "
'        Sql = Sql & "FEC_RECCIA, FEC_INGCIA, FEC_EFECTO, FEC_INIRETROPEN, "
'        Sql = Sql & "FEC_TERRETROPEN, COD_USUARIOCREA, FEC_CREA, HOR_CREA, "
'        Sql = Sql & "COD_USUARIOMODI, FEC_MODI, HOR_MODI, EST_ACT " ', COD_INDRELIQUIDAR, NUM_RELIQ"
'        Sql = Sql & ")"
'        Sql = Sql & "VALUES ("
'        Sql = Sql & "'" & Trim(vgRs2!Num_Poliza) & "' , "
'        Sql = Sql & " " & cgNumeroEndosoInicial & ", " '1
'        Sql = Sql & " " & Str(vgRs2!Num_Orden) & ","
'        Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
'            'HQR 13/10/2007 Se deja la Fecha de termino del certificado de estudio como el último dia del mes
'            'Sql = Sql & " 'SUP','" & Format(DateAdd("m", 6, DateSerial(Mid(vgRs!Fec_IniPagoPen, 1, 4), Mid(vgRs!Fec_IniPagoPen, 5, 2), Mid(vgRs!Fec_IniPagoPen, 7, 2))), "yyyymmdd") & "',"
'        Sql = Sql & " 'EST','" & Format(DateAdd("m", 6, DateSerial(Mid(vgRs!Fec_IniPagoPen, 1, 4), Mid(vgRs!Fec_IniPagoPen, 5, 2), Mid(vgRs!Fec_IniPagoPen, 7, 2) - 1)), "yyyymmdd") & "',"
'        Sql = Sql & " 'S', 'Certificado de estudios inicial.',"
'        Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
'        Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
'        Sql = Sql & " '" & vgRs!Fec_IniPagoPen & "',"
'        Sql = Sql & " NULL, NULL,"
'        Sql = Sql & "'" & vlGlsUsuarioCrea & "',"
'        Sql = Sql & "'" & vlFecCrea & "',"
'        Sql = Sql & "'" & vlHorCrea & "', "
'        Sql = Sql & " NULL, NULL, NULL, '1'"
'            'Sql = Sql & ",NULL,NULL"
'        Sql = Sql & ")"
'        vgConectarBD.Execute Sql
'    End If
'    'End If
    
 
End Function

Function flInsertarPoliza()

    vlGlsUsuarioCrea = vgUsuario
    vlFecCrea = Format(Date, "yyyymmdd")
    vlHorCrea = Format(Time, "hhmmss")
    
    Sql = ""
    Sql = "INSERT INTO PP_TMAE_POLIZA "
    Sql = Sql & " (num_poliza,num_endoso,cod_afp,cod_tippension, "
    Sql = Sql & " cod_estado,cod_tipren,cod_modalidad,num_cargas, "
    Sql = Sql & " fec_vigencia,fec_tervigencia,mto_prima,mto_pension, "
    Sql = Sql & " mto_pensiongar,"
    Sql = Sql & " num_mesdif,num_mesgar,prc_tasace,prc_tasavta,fec_dev, "
    Sql = Sql & " prc_tasactorea,prc_tasaintpergar,fec_inipagopen, "
    Sql = Sql & " cod_cobercon,cod_cuspp,cod_dercre,cod_dergra, "
    Sql = Sql & " cod_moneda,fec_emision,ind_cob,mto_facpenella, "
    Sql = Sql & " mto_valmoneda,prc_facpenella,prc_tasatir, "
    Sql = Sql & " fec_inipencia,fec_pripago, "
    Sql = Sql & " cod_usuariocrea,fec_crea,hor_crea "
    If (vgRs!Num_MesDif > 0) Then Sql = Sql & ",fec_finperdif "
    If (vgRs!Num_MesGar > 0) Then Sql = Sql & ",fec_finpergar "
'I--- ABV 05/02/2011 ---
'mvg 20170904
    Sql = Sql & ",cod_tipreajuste,mto_valreajustetri,mto_valreajustemen, fec_devsol, ind_bendes,ind_bolelec "
'F--- ABV 05/02/2011 ---
    Sql = Sql & " ) VALUES ( "
    Sql = Sql & "'" & (vgRs!Num_Poliza) & "', "
    Sql = Sql & " " & cgNumeroEndosoInicial & ", " '1
    Sql = Sql & "'" & (vgRs!Cod_AFP) & "' , "
    Sql = Sql & "'" & (vgRs!Cod_TipPension) & "' , "
    Sql = Sql & "'" & clCodEstado6 & "' , "  'Vigente
    Sql = Sql & "'" & (vgRs!Cod_TipRen) & "' , "
    Sql = Sql & "'" & (vgRs!Cod_Modalidad) & "' , "
    Sql = Sql & " " & (vgRs!Num_Cargas) & " , "
    Sql = Sql & "'" & (vgRs!Fec_Vigencia) & "' , "
    Sql = Sql & "'" & clFechaTopeTer & "' , " '99991231
    Sql = Sql & " " & str(vgRs!MTO_PRIUNI) & " , "
    Sql = Sql & " " & str(vgRs!Mto_Pension) & " , "
    Sql = Sql & " " & str(vgRs!Mto_PensionGar) & ", "
    Sql = Sql & " " & str(vgRs!Num_MesDif) & ", "
    Sql = Sql & " " & str(vgRs!Num_MesGar) & ", "
    Sql = Sql & " " & str(vgRs!Prc_TasaCe) & ", "
    Sql = Sql & " " & str(vgRs!Prc_TasaVta) & ", "
    Sql = Sql & "'" & (vgRs!Fec_Dev) & "', "
    Sql = Sql & " " & str(vlFactorEsc) & " , " '0 Hasta que se calcule el reaseguro posteriormente
    Sql = Sql & " " & str(vgRs!prc_tasapergar) & " , "
    Sql = Sql & "'" & (vgRs!Fec_IniPagoPen) & "' , "
    Sql = Sql & "'" & (vgRs!Cod_CoberCon) & "' , "
    Sql = Sql & "'" & (vgRs!Cod_Cuspp) & "' , "
    Sql = Sql & "'" & (vgRs!Cod_DerCre) & "' , "
    Sql = Sql & "'" & (vgRs!Cod_DerGra) & "' , "
    Sql = Sql & "'" & (vgRs!Cod_Moneda) & "' , "
    Sql = Sql & "'" & (vgRs!Fec_Emision) & "' , "
    Sql = Sql & "'" & (vgRs!Ind_Cob) & "' , "
    Sql = Sql & " " & str(vgRs!Mto_FacPenElla) & ", "
    Sql = Sql & " " & str(vgRs!Mto_ValMoneda) & ", "
    Sql = Sql & " " & str(vgRs!Prc_FacPenElla) & ", "
    Sql = Sql & " " & str(vgRs!Prc_TasaTir) & ", "
    Sql = Sql & "'" & (vgRs!fec_inipencia) & "', "
    Sql = Sql & "'" & (vgRs!fec_pripago) & "' , "
    Sql = Sql & "'" & vlGlsUsuarioCrea & "', "
    Sql = Sql & "'" & vlFecCrea & "', "
    Sql = Sql & "'" & vlHorCrea & "' "
    If (vgRs!Num_MesDif > 0) Then Sql = Sql & ",'" & vgRs!fec_finperdif & "' "
    If (vgRs!Num_MesGar > 0) Then Sql = Sql & ",'" & vgRs!fec_finpergar & "' "
'I--- ABV 05/02/2011 ---
    Sql = Sql & ",'" & (vgRs!Cod_TipReajuste) & "',"
    Sql = Sql & " " & str(vgRs!Mto_ValReajusteTri) & ","
    Sql = Sql & " " & str(vgRs!Mto_ValReajusteMen) & ","
    ''RRR 18/9/13
    Sql = Sql & " " & str(vgRs!fec_devsol) & ", "
    Sql = Sql & " '" & vgRs!ind_bendes & "', "
    'mvg 20170904
    Sql = Sql & " '" & IIf(IsNull(vgRs!ind_bolelec), "N", vgRs!ind_bolelec) & "' "
'F--- ABV 05/02/2011 ---
    Sql = Sql & ") "
    vgConectarBD.Execute Sql

End Function

Function flMarcarPolizaTraspasada()

    vlGlsUsuarioModi = vgUsuario
    vlFecModi = Format(Date, "yyyymmdd")
    vlHorModi = Format(Time, "hhmmss")

    Sql = ""
    Sql = " UPDATE pd_tmae_poliza SET "
    Sql = Sql & " cod_trapagopen = 'S', "
    Sql = Sql & " fec_trapagopen =  '" & Format(CDate(Trim(Lbl_FechaOpera.Caption)), "yyyymmdd") & "', "
    Sql = Sql & " cod_usuariomodi = '" & vlGlsUsuarioModi & "', "
    Sql = Sql & " fec_modi = '" & vlFecModi & "', "
    Sql = Sql & " hor_modi = '" & vlHorModi & "' "
    Sql = Sql & " WHERE "
    Sql = Sql & " num_poliza = '" & Trim(vlNumPoliza) & "' "
    vgConectarBD.Execute Sql

End Function

Function flInsertarLiquidacion()

    'Select a la Tabla por Nº Poliza y ordenado por periodo
    Sql = "select num_poliza,cod_banco,cod_direccion,cod_inssalud,"
    Sql = Sql & "cod_moneda,cod_sucursal,cod_tipcuenta,cod_tipoidenreceptor,"
    Sql = Sql & "cod_tipopago,cod_tippension,cod_tipreceptor,cod_viapago,"
    Sql = Sql & "fec_pago,gls_direccion,gls_matreceptor,gls_nomreceptor,"
    Sql = Sql & "gls_nomsegreceptor,gls_patreceptor,mto_baseimp,mto_basetri,"
    Sql = Sql & "mto_descuento,mto_haber,mto_liqpagar,mto_pension,mto_plansalud,"
    Sql = Sql & "num_cuenta,num_idenreceptor,num_orden,num_perpago,cod_modsalud "
    Sql = Sql & "FROM PD_TMAE_LIQPAGOPEN "
    Sql = Sql & "WHERE num_poliza = '" & Trim(vlNumPoliza) & "' "
    Sql = Sql & "order by num_perpago"
    Set vgRs4 = vgConexionBD.Execute(Sql)
    If Not vgRs4.EOF Then
        vlSw = True
        While Not vgRs4.EOF
            Sql = ""
            Sql = "INSERT INTO PP_TMAE_LIQPAGOPENDEF "
            Sql = Sql & " (num_poliza,cod_banco,cod_direccion,cod_inssalud,cod_modsalud, "
            Sql = Sql & " cod_moneda,cod_sucursal,cod_tipcuenta,cod_tipoidenreceptor, "
            Sql = Sql & " cod_tipopago,cod_tippension,cod_tipreceptor,cod_viapago, "
            Sql = Sql & " fec_pago,gls_direccion,gls_matreceptor,gls_nomreceptor, "
            Sql = Sql & " gls_nomsegreceptor,gls_patreceptor,mto_baseimp,mto_basetri, "
            Sql = Sql & " mto_descuento,mto_haber,mto_liqpagar,mto_pension, "
            Sql = Sql & " mto_plansalud,num_cargas,num_cuenta,num_endoso, "
            Sql = Sql & " num_idenreceptor,num_orden,num_perpago "
            Sql = Sql & " ) VALUES ( "
            Sql = Sql & "'" & (vgRs4!Num_Poliza) & "', "
            Sql = Sql & "'" & (vgRs4!Cod_Banco) & "', "
            Sql = Sql & " " & (vgRs4!Cod_Direccion) & ", "
            Sql = Sql & "'" & (vgRs4!Cod_InsSalud) & "', "
            Sql = Sql & "'" & (vgRs4!Cod_ModSalud) & "', "
            Sql = Sql & "'" & (vgRs4!Cod_Moneda) & "', "
            Sql = Sql & "'" & (vgRs4!Cod_Sucursal) & "', "
            Sql = Sql & "'" & (vgRs4!Cod_TipCuenta) & "', "
            Sql = Sql & " " & (vgRs4!Cod_TipoIdenReceptor) & ", "
            Sql = Sql & "'" & (vgRs4!Cod_TipoPago) & "', "
            Sql = Sql & "'" & (vgRs4!Cod_TipPension) & "', "
            Sql = Sql & "'" & (vgRs4!Cod_TipReceptor) & "', "
            Sql = Sql & "'" & (vgRs4!Cod_ViaPago) & "', "
            Sql = Sql & "'" & (vgRs4!Fec_Pago) & "', "
            Sql = Sql & "'" & (vgRs4!Gls_Direccion) & "', "
            If IsNull(vgRs4!Gls_MatReceptor) Then
                Sql = Sql & " NULL, "
            Else
                Sql = Sql & "'" & (vgRs4!Gls_MatReceptor) & "', "
            End If
            Sql = Sql & "'" & (vgRs4!Gls_NomReceptor) & "', "
            If IsNull(vgRs4!Gls_NomSegReceptor) Then
                Sql = Sql & " NULL, "
            Else
                Sql = Sql & "'" & (vgRs4!Gls_NomSegReceptor) & "', "
            End If
            Sql = Sql & "'" & (vgRs4!Gls_PatReceptor) & "', "
            Sql = Sql & " " & str(vgRs4!Mto_BaseImp) & ", "
            Sql = Sql & " " & str(vgRs4!Mto_BaseTri) & ", "
            Sql = Sql & " " & str(vgRs4!Mto_Descuento) & ", "
            Sql = Sql & " " & str(vgRs4!Mto_Haber) & ", "
            Sql = Sql & " " & str(vgRs4!Mto_LiqPagar) & ", "
            Sql = Sql & " " & str(vgRs4!Mto_Pension) & ", "
            Sql = Sql & " " & str(vgRs4!Mto_PlanSalud) & ", "
            Sql = Sql & "'0', "
            If IsNull(vgRs4!Num_Cuenta) Then
                Sql = Sql & " NULL, "
            Else
                Sql = Sql & "'" & (vgRs4!Num_Cuenta) & "', "
            End If
            Sql = Sql & " " & cgNumeroEndosoInicial & ","
            Sql = Sql & "'" & (vgRs4!Num_IdenReceptor) & "', "
            Sql = Sql & " " & (vgRs4!Num_Orden) & ", "
            Sql = Sql & "'" & (vgRs4!Num_PerPago) & "'"
            Sql = Sql & ")"
            vgConectarBD.Execute Sql
        
            vgRs4.MoveNext
        Wend
    End If
    vgRs4.Close
    
End Function

Function flInsertarPagos()

    'Select a la Tabla por Nº Poliza y ordenado por periodo
    Sql = "select num_poliza,cod_conhabdes,cod_tipoidenreceptor,"
    Sql = Sql & "cod_tipreceptor,fec_inipago,fec_terpago,mto_conhabdes,"
    Sql = Sql & "num_idenreceptor,num_orden,num_perpago "
    Sql = Sql & "FROM pd_tmae_pagopen "
    Sql = Sql & "WHERE num_poliza = '" & Trim(vlNumPoliza) & "' "
    Sql = Sql & "order by num_perpago"
    Set vgRs4 = vgConexionBD.Execute(Sql)
    If Not vgRs4.EOF Then
        vlSw = True
        While Not vgRs4.EOF
            Sql = ""
            Sql = "INSERT INTO PP_TMAE_PAGOPENDEF "
            Sql = Sql & " (num_poliza,cod_conhabdes,cod_tipoidenreceptor,"
            Sql = Sql & " cod_tipreceptor,fec_inipago,fec_terpago,mto_conhabdes,"
            Sql = Sql & " num_endoso,num_idenreceptor,num_orden,num_perpago "
            Sql = Sql & " ) VALUES ( "
            Sql = Sql & "'" & (vgRs4!Num_Poliza) & "', "
            Sql = Sql & "'" & (vgRs4!Cod_ConHabDes) & "', "
            Sql = Sql & " " & (vgRs4!Cod_TipoIdenReceptor) & ", "
            Sql = Sql & "'" & (vgRs4!Cod_TipReceptor) & "', "
            Sql = Sql & "'" & (vgRs4!Fec_IniPago) & "', "
            Sql = Sql & "'" & (vgRs4!Fec_TerPago) & "', "
            Sql = Sql & " " & str(vgRs4!Mto_ConHabDes) & ", "
            Sql = Sql & " " & cgNumeroEndosoInicial & ", "
            Sql = Sql & "'" & (vgRs4!Num_IdenReceptor) & "', "
            Sql = Sql & " " & (vgRs4!Num_Orden) & ", "
            Sql = Sql & "'" & (vgRs4!Num_PerPago) & "' "
            Sql = Sql & ")"
            vgConectarBD.Execute Sql
        
            vgRs4.MoveNext
        Wend
    End If
    vgRs4.Close
    
End Function

Function flInsertarPensionActualizada()
    
    'Select a la Tabla por Nº Poliza y ordenado por periodo
    Sql = "select num_poliza,fec_desde,mto_pension, mto_pensiongar, prc_fatorajus "
    Sql = Sql & "FROM pd_tmae_pensionact "
    Sql = Sql & "WHERE num_poliza = '" & Trim(vlNumPoliza) & "' "
    Sql = Sql & "order by num_poliza,fec_desde "
    Set vgRs4 = vgConexionBD.Execute(Sql)
    If Not vgRs4.EOF Then
        While Not vgRs4.EOF
            
            Sql = "INSERT INTO PP_TMAE_PENSIONACT "
            Sql = Sql & "(NUM_POLIZA, FEC_DESDE, MTO_PENSION, NUM_ENDOSO, MTO_PENSIONGAR, PRC_FATORAJUS "
            Sql = Sql & ") VALUES ("
            Sql = Sql & "'" & vgRs4!Num_Poliza & "',"
            Sql = Sql & "'" & vgRs4!Fec_desde & "',"
            Sql = Sql & " " & str(vgRs4!Mto_Pension) & ","
            Sql = Sql & " " & cgNumeroEndosoInicial & ","
            Sql = Sql & " " & str(vgRs4!Mto_PensionGar) & ","
            Sql = Sql & " " & str(vgRs4!prc_fatorajus) & " "
            Sql = Sql & ")"
            vgConectarBD.Execute Sql
        
            vgRs4.MoveNext
        Wend
    End If
    vgRs4.Close

End Function

Function flTraspaso()
On Error GoTo Err_Traspaso

    vlEstadoTrans = False
    
    If Not fgConexionBaseDatos(vgConectarBD) Then
        MsgBox "Falló la Conexión con la Base de Datos.", vbCritical, "Error de Conexión"
        Exit Function
    End If
        
    'Comenzar Transacción
    vgConectarBD.BeginTrans
    
    Sql = ""
    Sql = "SELECT num_poliza "
    Sql = Sql & "FROM pp_tmae_poliza "
    Sql = Sql & "WHERE num_poliza = '" & Trim(vlNumPoliza) & "' "
    Set vgRs3 = vgConexionBD.Execute(Sql)
    If vgRs3.EOF Then
        vgSql = "SELECT num_poliza,num_endoso,cod_afp,cod_cobercon,cod_cuspp, "
        vgSql = vgSql & "cod_dercre,cod_dergra,cod_direccion,cod_isapre,cod_modalidad, "
        vgSql = vgSql & "cod_moneda,cod_tippension,cod_tipren,ind_cob,"
        vgSql = vgSql & "cod_viapago,cod_sucursal,cod_tipcuenta,cod_banco,num_cuenta,"
        vgSql = vgSql & "fec_dev,fec_emision,fec_finperdif,fec_finpergar,"
        vgSql = vgSql & "fec_inipagopen,fec_inipencia,fec_vigencia, "
        vgSql = vgSql & "gls_direccion,gls_fono,gls_correo,mto_facpenella, "
        vgSql = vgSql & "mto_pension,mto_pensiongar,mto_priuni,mto_valmoneda,"
        vgSql = vgSql & "num_cargas,num_mesdif,num_mesgar,prc_facpenella, "
        vgSql = vgSql & "prc_tasace,prc_tasapergar,prc_tasatir,Prc_TasaVta,"
        vgSql = vgSql & "fec_inipencia, fec_pripago "
        vgSql = vgSql & ",fec_finperdif, fec_finpergar "
'I--- ABV 05/02/2011 ---
        vgSql = vgSql & ",cod_tipreajuste,mto_valreajustetri,mto_valreajustemen, fec_devsol, mto_rentatmpafp, ind_bendes, ind_bolelec "
'F--- ABV 05/02/2011 ---
        vgSql = vgSql & "FROM pd_tmae_poliza "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' AND "
        vgSql = vgSql & "cod_trapagopen = 'N' "
        Set vgRs = vgConexionBD.Execute(vgSql)
        If Not vgRs.EOF Then
            vlCodTipPension = vgRs!Cod_TipPension
            vgSql = ""
            vgSql = "SELECT * "
            vgSql = vgSql & "FROM pd_tmae_polben "
            vgSql = vgSql & "WHERE "
            vgSql = vgSql & "num_poliza = '" & Trim(vlNumPoliza) & "' "
            Set vgRs2 = vgConexionBD.Execute(vgSql)
            If Not vgRs2.EOF Then
                vlAnno = Mid(Trim(vgRs!fec_inipencia), 1, 4)
                vlMes = Mid(Trim(vgRs!fec_inipencia), 5, 2)
                vlDia = Mid(Trim(vgRs!fec_inipencia), 7, 2)
                vlFactorEsc = 0
                If vgRs!Cod_TipRen = "6" Then
                    vlFactorEsc = Math.Round(CDbl(vgRs!Mto_RentaTMPAFP) / CDbl(vgRs!Mto_Pension), 5)
                End If 'RRR 05/08/2016
                
'                If (vgRs!Num_MesGar > 0) And (vlCodTipPension <> clCodTipPensionSob) Then
'                    'vlFecTerPagoPenGar = DateSerial(vlAnno, vlMes + (vgRs!Num_MesGar), vlDia - 1)
'                    'vlFecTerPagoPenGar = Format(CDate(vlFecTerPagoPenGar), "yyyymmdd")
'                    vlFecTerPagoPenGar = IIf(IsNull(vgRs!fec_finpergar), "", vgRs!fec_finpergar)
'                Else
                    vlFecTerPagoPenGar = ""
'                End If
                  
                Call flInsertarPoliza
                                
                While Not vgRs2.EOF
                    Call flInsertarBeneficiario
                    vgRs2.MoveNext
                Wend
                        
                'Traspaso de la Liquidación
                vlSw = False
                Call flInsertarLiquidacion

                'Traspaso de Pagos de Pensión - Conceptos
                If (vlSw = True) Then
                    Call flInsertarPagos
                    Call flInsertarPensionActualizada
                    
                End If

                Call flMarcarPolizaTraspasada
            End If
            vgRs2.Close
        End If
    End If
        
    'para de pd_tmae_poltutor a pp_tmae_tutor
    Call TraspasoTutores(Trim(vlNumPoliza))
    
    'Ejecutar Transacción
    vgConectarBD.CommitTrans
        
    'Cerrar Transacción
    vgConectarBD.Close
    
    vlEstadoTrans = True

Exit Function
Err_Traspaso:
    
    'Deshacer Transacción
    vgConectarBD.RollbackTrans
    'Cerrar Transacción
    vgConectarBD.Close
    
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
Private Sub TraspasoTutores(ByVal poliza As String)

    vlGlsUsuarioModi = vgUsuario
    vlFecModi = Format(ObtenerFechaServer, "yyyymmdd")
    vlHorModi = Format(Time, "hhmmss")

    Sql = ""
    Sql = "Insert into pp_tmae_tutor(NUM_POLIZA,NUM_ENDOSO, NUM_ORDEN, NUM_IDENTUT, COD_TIPOIDENTUT, GLS_NOMTUT, GLS_NOMSEGTUT, "
    Sql = Sql & " GLS_PATTUT, GLS_MATTUT, GLS_DIRTUT, COD_DIRECCION, GLS_FONOTUT, NUM_MESPODNOT, FEC_INIPODNOT, FEC_TERPODNOT,"
    Sql = Sql & " COD_VIAPAGO, COD_TIPCUENTA, COD_BANCO, NUM_CUENTA, COD_SUCURSAL, FEC_EFECTO, FEC_RECCIA, COD_USUARIOCREA, FEC_CREA,"
    Sql = Sql & " HOR_CREA, COD_USUARIOMODI, FEC_MODI, HOR_MODI, gls_correotut)"
    Sql = Sql & " select NUM_POLIZA,NUM_ENDOSO, NUM_ORDEN, NUM_IDENTUT, COD_TIPOIDENTUT, GLS_NOMTUT, GLS_NOMSEGTUT, GLS_PATTUT, GLS_MATTUT,"
    Sql = Sql & " GLS_DIRTUT, COD_DIRECCION, GLS_FONOTUT, NUM_MESPODNOT, FEC_INIPODNOT, FEC_TERPODNOT, COD_VIAPAGO, COD_TIPCUENTA,"
    Sql = Sql & " COD_BANCO, NUM_CUENTA, COD_SUCURSAL, FEC_EFECTO, FEC_RECCIA, '" & vlGlsUsuarioModi & "', '" & vlFecModi & "', '" & vlHorModi & "', '" & vlGlsUsuarioModi & "',"
    Sql = Sql & " FEC_MODI, HOR_MODI, gls_correotut from PD_TMAE_poltutor where NUM_POLIZA='" & poliza & "'"
    
    
    vgConectarBD.Execute Sql

End Sub
Private Sub Cmd_Agregar_Click()
On Error GoTo Err_CmdAgregarClick

    Call flAgregarDetalle
    
Exit Sub
Err_CmdAgregarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_AgregarTodos_Click()
On Error GoTo Err_CmdAgregarTodosClick

    Call flMoverTodos(Msf_GrillaRecibidas, Msf_GrillaTraspasadas)
    
Exit Sub
Err_CmdAgregarTodosClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Grabar_Click()
Dim vlResp As Long
On Error GoTo Err_CmdGrabar
    
    If Msf_GrillaTraspasadas.Row = 0 Then
       MsgBox "No Existen Polizas Traspasadas para Grabar", vbInformation, "Información"
       Exit Sub
    End If
    
    vlResp = MsgBox(" ¿ Está seguro que desea realizar el Traspaso de Datos ?", 4 + 32 + 256, "Proceso de Ingreso de Datos")
    If vlResp <> 6 Then
        Cmd_Salir.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If Msf_GrillaTraspasadas.rows > 1 Then
       vlPos = 1
       While vlPos <= (Msf_GrillaTraspasadas.rows - 1)
       
             Msf_GrillaTraspasadas.Row = vlPos
             Msf_GrillaTraspasadas.Col = 0
             vlNumPoliza = Trim(Msf_GrillaTraspasadas.Text)
             
             Call flTraspaso
            
             If vlEstadoTrans = False Then
                MsgBox "El Registro de los Datos No Fue Realizado", vbCritical, "Error de Datos"
                Exit Sub
             End If
                           
             vlPos = vlPos + 1
       Wend
       Screen.MousePointer = 11
       MsgBox "El Traspaso de los Datos fue realizado Satisfactoriamente.", vbInformation, "Estado Operación"
       Screen.MousePointer = 0
       Call flInicializaGrilla(Msf_GrillaTraspasadas)
    End If

Exit Sub
Err_CmdGrabar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_Restar_Click()
On Error GoTo Err_CmdRestarClick

    Call flQuitarDetalle
    
Exit Sub
Err_CmdRestarClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Cmd_RestarTodos_Click()
On Error GoTo Err_CmdRestarTodosClick
             
    Call flMoverTodos(Msf_GrillaTraspasadas, Msf_GrillaRecibidas)
    
Exit Sub
Err_CmdRestarTodosClick:
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

    Frm_CalTraspaso.Top = 0
    Frm_CalTraspaso.Left = 0
    
    Lbl_FechaOpera.Caption = fgBuscaFecServ
    
    Call flInicializaGrilla(Msf_GrillaRecibidas)
    Call flInicializaGrilla(Msf_GrillaTraspasadas)
    Call flCargaGrilla
    
Exit Sub
Err_Cargar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_GrillaRecibidas_Click()
On Error GoTo Err_MsfGrillaRecibidasClick

    Lbl_ClickGrilla.Caption = Msf_GrillaRecibidas.Text
    
Exit Sub
Err_MsfGrillaRecibidasClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_GrillaRecibidas_dblClick()
On Error GoTo Err_MsfGrillaRecibidasDblClick

    Msf_GrillaRecibidas.Col = 0
    Lbl_ClickGrilla.Caption = Msf_GrillaRecibidas.Text
    Call flAgregarDetalle
    Lbl_ClickGrilla.Caption = ""
    
Exit Sub
Err_MsfGrillaRecibidasDblClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_GrillaTraspasadas_Click()
On Error GoTo Err_MsfGrillaTraspasadasClick

    Lbl_ClickGrilla.Caption = Msf_GrillaTraspasadas.Text
    
Exit Sub
Err_MsfGrillaTraspasadasClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub

Private Sub Msf_GrillaTraspasadas_dblClick()
On Error GoTo Err_MsfGrillaTraspasadasDblClick

    Msf_GrillaTraspasadas.Col = 0
    Lbl_ClickGrilla.Caption = Msf_GrillaTraspasadas.Text
    Call flQuitarDetalle
    Lbl_ClickGrilla.Caption = ""
    
Exit Sub
Err_MsfGrillaTraspasadasDblClick:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Sub


