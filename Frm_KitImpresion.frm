VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_KitImpresion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de KiT"
   ClientHeight    =   5295
   ClientLeft      =   5190
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8865
   Begin VB.CommandButton Cmd_Imprimir 
      Caption         =   "&AFP"
      Height          =   1035
      Index           =   2
      Left            =   10800
      Picture         =   "Frm_KitImpresion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Imprimir Reporte"
      Top             =   7080
      Width           =   885
   End
   Begin VB.Frame Frame5 
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
      Height          =   1185
      Left            =   180
      TabIndex        =   41
      Top             =   0
      Width           =   8565
      Begin MSComCtl2.DTPicker DTPDesde 
         Height          =   345
         Left            =   3390
         TabIndex        =   47
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   43122689
         CurrentDate     =   43573
      End
      Begin VB.CommandButton Cmd_Buscarpolizas 
         Caption         =   "&Buscar"
         Height          =   675
         Left            =   7770
         Picture         =   "Frm_KitImpresion.frx":53E2
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Buscar Todas las Póliza"
         Top             =   450
         Width           =   720
      End
      Begin VB.OptionButton OptHoy 
         Caption         =   "Hoy"
         Height          =   255
         Left            =   180
         TabIndex        =   43
         Top             =   300
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton OptRangoFechas 
         Caption         =   "Por Rango de Fechas"
         Height          =   405
         Left            =   180
         TabIndex        =   42
         Top             =   690
         Width           =   2505
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   345
         Left            =   5100
         TabIndex        =   48
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   43122689
         CurrentDate     =   43573
      End
      Begin VB.Label Lbl_Buscador 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   46
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Lbl_Buscador 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   3390
         TabIndex        =   45
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.Frame Fra_AntGral 
      Caption         =   "  Antecedentes Generales  "
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
      Height          =   3255
      Left            =   480
      TabIndex        =   7
      Top             =   5850
      Width           =   8655
      Begin VB.Label Lbl_ReajusteDescripcion 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Lbl_ReajusteValorMen 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   39
         Top             =   915
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Lbl_ReajusteValor 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   38
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   29
         Left            =   240
         TabIndex        =   37
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Lbl_ReajusteTipo 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6120
         TabIndex        =   36
         Top             =   915
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Reajuste Trim."
         Height          =   255
         Index           =   28
         Left            =   6000
         TabIndex        =   35
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Lbl_NumLiquidacion 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6600
         TabIndex        =   34
         Top             =   2805
         Width           =   1575
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "N° Boleta de Venta"
         Height          =   255
         Index           =   9
         Left            =   4680
         TabIndex        =   33
         Top             =   2805
         Width           =   1815
      End
      Begin VB.Label Lbl_FechaRec 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   2805
         Width           =   1695
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Fecha Traspaso Prima"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   31
         Top             =   2805
         Width           =   1815
      End
      Begin VB.Label Lbl_Diferidos 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   30
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Lbl_Meses 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   29
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label Lbl_PensionDef 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6600
         TabIndex        =   28
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Lbl_PrimaDef 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Lbl_NumIdent 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4920
         TabIndex        =   26
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Lbl_TipoIdent 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Lbl_Modalidad 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   1755
         Width           =   3615
      End
      Begin VB.Label Lbl_TipoRenta 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   1470
         Width           =   3615
      End
      Begin VB.Label Lbl_TipoPension 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   1185
         Width           =   6135
      End
      Begin VB.Label Lbl_NomAfiliado 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   630
         Width           =   6135
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nº Ident."
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Meses Garant."
         Height          =   255
         Index           =   6
         Left            =   6000
         TabIndex        =   19
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Años Diferidos"
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   18
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Modalidad"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   1755
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Renta"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Tipo de Pensión"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1185
         Width           =   1215
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   630
         Width           =   1215
      End
      Begin VB.Line Lin_Separar 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   8160
         Y1              =   2385
         Y2              =   2385
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Prima Definitiva"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "Pensión Definitiva"
         Height          =   195
         Index           =   8
         Left            =   4680
         TabIndex        =   12
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label Lbl_Nombre 
         Caption         =   "CUSPP"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   11
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label Lbl_CUSPP 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   915
         Width           =   2535
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   9
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbl_Moneda 
         Alignment       =   1  'Right Justify
         Caption         =   "(TM)"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   8
         Top             =   2520
         Width           =   375
      End
   End
   Begin VB.Frame Fra_Operacion 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   4020
      Width           =   8655
      Begin VB.CommandButton cmdBoleta 
         Caption         =   "&Boleta"
         Height          =   1035
         Left            =   3960
         Picture         =   "Frm_KitImpresion.frx":54E4
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Prima"
         Height          =   1035
         Index           =   5
         Left            =   3000
         Picture         =   "Frm_KitImpresion.frx":A8C6
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Imprimir Reporte"
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton cmdReporteBienvenida 
         Caption         =   "&Bienvenida"
         Height          =   1035
         Left            =   120
         Picture         =   "Frm_KitImpresion.frx":FCA8
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton cmd_afp 
         Caption         =   "&AFP"
         Height          =   1035
         Left            =   2040
         Picture         =   "Frm_KitImpresion.frx":1508A
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton cmdPoliza 
         Caption         =   "&Póliza"
         Height          =   1035
         Left            =   1080
         Picture         =   "Frm_KitImpresion.frx":1A46C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   885
      End
      Begin VB.CommandButton Cmd_Salir 
         Caption         =   "&Salir"
         Height          =   1035
         Left            =   4920
         Picture         =   "Frm_KitImpresion.frx":1F84E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir del Formulario"
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.CommandButton Cmd_Imprimir 
      Caption         =   "&Bienvenida"
      Height          =   675
      Index           =   0
      Left            =   9750
      Picture         =   "Frm_KitImpresion.frx":1F948
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Reporte"
      Top             =   6480
      Width           =   960
   End
   Begin VB.TextBox Txt_Poliza 
      Height          =   285
      Left            =   9420
      MaxLength       =   10
      TabIndex        =   2
      Top             =   6030
      Width           =   1545
   End
   Begin VB.TextBox Txt_Endoso 
      Height          =   285
      Left            =   10980
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "1"
      Top             =   6030
      Width           =   465
   End
   Begin MSFlexGridLib.MSFlexGrid GrdPol 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4789
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      BackColor       =   14745599
      GridColor       =   0
      AllowBigSelection=   0   'False
      HighLight       =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport Rpt_Reporte 
      Left            =   9810
      Top             =   7290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Frm_KitImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type DatosPoliza
    Num_Poliza As String
    tippension  As String
    tiporenta As String
    Num_MesDif As Integer
    modalidad  As String
    Num_MesGar As Integer
    nombres As String
    Reportes As String
    Num_Endoso As Integer
    Tipo_Renta As String
End Type


Dim dPol As DatosPoliza
Dim NumPolSel As String

Private Sub ObtenerPolEmitidas()

                Dim Mensaje As String
         
                Dim conn    As ADODB.Connection
                Set conn = New ADODB.Connection
                Dim objCmd As ADODB.Command
                Dim RS As ADODB.Recordset
                
                Set RS = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
                Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
                
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMARRVV.sp_ObtenerPolEmitidas"
                objCmd.CommandType = adCmdStoredProc
                
                Dim param1 As ADODB.Parameter
                Dim param2 As ADODB.Parameter
                Dim param3 As ADODB.Parameter
                
                Set param1 = objCmd.CreateParameter("pfechaDesde", adVarChar, adParamInput, 10, Format(Me.DTPDesde.Value, "yyyyMMDD"))
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("pfechaHasta", adVarChar, adParamInput, 10, Format(Me.DTPHasta.Value, "yyyyMMDD"))
                objCmd.Parameters.Append param2
                                     
                Set param3 = objCmd.CreateParameter("mensajeError", adVarChar, adParamOutput, 100)
                objCmd.Parameters.Append param3
                                   
                Set RS = objCmd.Execute
                
                If Not IsNull(objCmd.Parameters.Item("mensajeError").Value) Then
                    Mensaje = objCmd.Parameters.Item("mensajeError").Value
                Else
                    Mensaje = ""
                End If
               
                If Len(Trim(Mensaje)) = 0 Then
                     If RS.RecordCount > 0 Then
                     
                        While Not RS.EOF
                    
                            dPol.Num_Poliza = IIf(IsNull(RS!Num_Poliza), "", RS!Num_Poliza)
                            dPol.tippension = IIf(IsNull(RS!tippension), "", RS!tippension)
                            dPol.tiporenta = IIf(IsNull(RS!tiporenta), "", RS!tiporenta)
                            dPol.Num_MesDif = IIf(IsNull(RS!Num_MesDif), 0, RS!Num_MesDif)
                            dPol.modalidad = IIf(IsNull(RS!modalidad), "", RS!modalidad)
                            dPol.Num_MesGar = IIf(IsNull(RS!Num_MesGar), 0, RS!Num_MesGar)
                            dPol.nombres = IIf(IsNull(RS!nombres), "", RS!nombres)
                      
                             GrdPol.AddItem dPol.Num_Poliza & vbTab _
                             & dPol.nombres & vbTab _
                             & dPol.tippension & vbTab _
                             & dPol.tiporenta & vbTab _
                             & dPol.modalidad & vbTab _
                             & dPol.Num_MesGar
                           
                            RS.MoveNext
                        Wend
                    Else
                          MsgBox "No existen Datos para el rango de fechas indicado.", vbExclamation, "Información"
                    End If
                
        
                Else
                
                    MsgBox Mensaje, vbCritical, "Error"
                
                End If
                
  conn.Close
  Set objCmd = Nothing
  Set RS = Nothing
  Set conn = Nothing
  
  GrdPol.ColAlignment(1) = 1
  GrdPol.ColAlignment(2) = 1
  GrdPol.ColAlignment(3) = 1
  GrdPol.ColAlignment(4) = 1
  GrdPol.ColAlignment(5) = 1
  
             
End Sub
Private Sub cargaDatos_poliza(ByVal pnum_poliza As String, _
                                 ByRef pnumError As Integer, _
                                 ByRef pnumMsg As String)
                                 


                Dim Mensaje As String
                Dim RS As ADODB.Recordset
                Dim objCmd As ADODB.Command
                
                Dim param1 As ADODB.Parameter
                Dim param2 As ADODB.Parameter
                Dim param3 As ADODB.Parameter
                Dim param4 As ADODB.Parameter
          
                Dim conn    As ADODB.Connection
                Set conn = New ADODB.Connection
                Set RS = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
                Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
                
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMARRVV.sp_datos_poliza"
                objCmd.CommandType = adCmdStoredProc
                                       
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 10, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("pnum_orden", adInteger, adParamInput)
                param2.Value = 1
                objCmd.Parameters.Append param2
                         
                Set param3 = objCmd.CreateParameter("p_outNumError", adInteger, adParamOutput)
                objCmd.Parameters.Append param3
                
                Set param4 = objCmd.CreateParameter("p_outMsgError", adVarChar, adParamOutput, 300)
                objCmd.Parameters.Append param4
                                   
                Set RS = objCmd.Execute
                
                If Not IsNull(objCmd.Parameters.Item("p_outNumError").Value) Then
                    pnumError = Trim(objCmd.Parameters.Item("p_outNumError").Value)
                Else
                    pnumError = 0
                End If
                                  
            
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    pnumMsg = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    pnumMsg = ""
                End If
    
                
                If Len(Trim(Mensaje)) = 0 Then
                
                  If RS.RecordCount > 0 Then
                     
                        While Not RS.EOF
                        
                        dPol.Num_Poliza = IIf(IsNull(RS!Num_Poliza), "", RS!Num_Poliza)
                        dPol.tippension = IIf(IsNull(RS!tippension), "", "e " & RS!tippension)
                        dPol.tiporenta = IIf(IsNull(RS!tiporenta), "", RS!tiporenta)
                        dPol.Num_MesDif = IIf(IsNull(RS!Num_MesDif), 0, RS!Num_MesDif)
                        dPol.modalidad = IIf(IsNull(RS!modalidad), "", RS!modalidad)
                        dPol.Num_MesGar = IIf(IsNull(RS!Num_MesGar), 0, RS!Num_MesGar)
                        dPol.nombres = IIf(IsNull(RS!nombres), "", RS!nombres)
                        dPol.Reportes = IIf(IsNull(RS!Reportes), "", RS!Reportes)
                        dPol.Num_Endoso = IIf(IsNull(RS!Num_Endoso), 0, RS!Num_Endoso)
                        dPol.Tipo_Renta = IIf(IsNull(RS!Tipo_Renta), "", RS!Tipo_Renta)
                        
                        
                        Wend
                        
                 End If
                        
        
                Else
                
                    MsgBox Mensaje, vbCritical, "Error"
                
                End If
                
  conn.Close
  Set objCmd = Nothing
  Set RS = Nothing
  Set conn = Nothing
  
End Sub

Private Sub Cmd_BuscarPolizas_Click()
    Call ObtenerPolEmitidas
End Sub

Private Sub cmdPoliza_Click()
    
    'NumPolSel = GrdPol.TextMatrix(GrdPol.Row, 0)
    NumPolSel = "0000008698"
    Dim MensajeError As String
    Dim SeguirSinCorreo As Boolean
    Dim NumError As Integer
    Dim ListaPDFS As String
    Dim ArrPdf() As String
    Dim sNomPdfFInal As String
    
    sNomPdfFInal = "PdfNum_" & NumPolSel
        
    Call sp_kitPolizaEmitidas(NumPolSel, ListaPDFS, NumError, MensajeError)
     
     If NumError <> 0 Then
        MsgBox "Error=>" & MensajeError
     Else
        MsgBox "Lista=>" & ListaPDFS
        
        If Len(Trim(ListaPDFS)) > 0 Then
        
                 EliminarArchivo ("C:\PDFS\" & sNomPdfFInal & ".pdf")
                 ArrPdf = Split(ListaPDFS, ";")
                 
                 Dim ObjOdf As ClsMergePDF.MergeTool
                 Dim sArchivoFinal As String
                 Dim x As Integer
                 
                 Set ObjOdf = New ClsMergePDF.MergeTool
                                
                 
                 For x = 0 To UBound(ArrPdf) - 1
                     
                      If ArrPdf(x + 1) <> "" Then
                      
                        If x = 0 Then
                           sArchivoFinal = ObjOdf.UnirPDF(strRpt & ArrPdf(x), strRpt & ArrPdf(x + 1), "C:\PDFS\" & sNomPdfFInal & "_part" & str(x) & ".pdf")
                        Else
                            sArchivoFinal = ObjOdf.UnirPDF(sArchivoFinal, strRpt & ArrPdf(x + 1), "C:\PDFS\" & sNomPdfFInal & "_part" & str(x) & ".pdf")
                        End If
                     
                     End If
                   
                 Next
                 
                 Name sArchivoFinal As "C:\PDFS\" & sNomPdfFInal & ".pdf"
                 Kill "C:\PDFS\*part*.*"
           
       End If
   End If
     
     
     
End Sub
Private Sub RptPoliza()
     
    Dim RS As ADODB.Recordset
    Dim vlFecTras As String
    Dim objRep As New ClsReporte
    Dim NomReporte As String
  
   NumPolSel = "0000008698"
  
    Set RS = New ADODB.Recordset
    
      If dPol.tiporenta <> "6" Then
        NomReporte = "PD_Rpt_PolizaDef.rpt"
    Else
        NomReporte = "PD_Rpt_PolizaDefEsc.rpt"
    End If
     
    RS.CursorLocation = adUseClient
    RS.Open "PP_LISTA_BIENVENIDA.LISTAR('" & NumPolSel & "','" & dPol.tippension & "','" & "10" & "','" & "99" & "', '" & str(dPol.Num_Endoso) & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaBien.rpt"), ".RPT", ".TTX"), 1)
    
    
      If objRep.CargaReporte_toPdf(strRpt & "", NomReporte, "Póliza", "", "0", RS, True, "c:\PDFS\hola.pdf", _
                           ArrFormulas("NombreAfi", dPol.nombres), _
                            ArrFormulas("TipoPension", dPol.tippension), _
                            ArrFormulas("MesGar", "2"), _
                            ArrFormulas("NombreCompania", "vgNombreCompania"), _
                            ArrFormulas("Concatenar", "vlCobertura"), _
                            ArrFormulas("Sucursal", "Surquillo"), _
                            ArrFormulas("RepresentanteNom", "vlRepresentante"), _
                            ArrFormulas("RepresentanteDoc", "vlDocum"), _
                            ArrFormulas("CodTipPen", Left(Trim(dPol.tippension), 2)), _
                            ArrFormulas("TipoDocTit", "Trim(Lbl_TipoIdent.Caption)"), _
                            ArrFormulas("NumDocTit", "Trim(Lbl_NumIdent.Caption)"), _
                            ArrFormulas("fec_trasp", "vlFecTras"), _
                            ArrFormulas("TipoRenta", dPol.tiporenta)) = False Then

                            
'
'     If objRep.CargaReporte(strRpt & "", NomReporte, "Póliza", rs, True, _
'                            ArrFormulas("NombreAfi", dPol.nombres), _
'                            ArrFormulas("TipoPension", dPol.tippension), _
'                            ArrFormulas("MesGar", dPol.Num_MesGar), _
'                            ArrFormulas("NombreCompania", "vgNombreCompania"), _
'                            ArrFormulas("Concatenar", "vlCobertura"), _
'                            ArrFormulas("Sucursal", "Surquillo"), _
'                            ArrFormulas("RepresentanteNom", "vlRepresentante"), _
'                            ArrFormulas("RepresentanteDoc", "vlDocum"), _
'                            ArrFormulas("CodTipPen", Left(Trim(dPol.tippension), 2)), _
'                            ArrFormulas("TipoDocTit", "Trim(Lbl_TipoIdent.Caption)"), _
'                            ArrFormulas("NumDocTit", "Trim(Lbl_NumIdent.Caption)"), _
'                            ArrFormulas("fec_trasp", "vlFecTras"), _
'                            ArrFormulas("TipoRenta", dPol.tiporenta)) = False Then
                            
                            MsgBox "No se pudo abrir el reporte", vbInformation
                            Exit Sub
     End If

End Sub
Private Sub RptBienvenida()
       'NumPolSel = GrdPol.TextMatrix(GrdPol.Row, 0)
    NumPolSel = "0000008698"
    Dim NumError As Integer
    Dim Mensaje As String
    
    
    Call cargaDatos_poliza(NumPolSel, NumError, Mensaje)
  
    
    Dim RS As ADODB.Recordset
    Dim vlFecTras As String
    Dim objRep As New ClsReporte
    
    
    Set RS = New ADODB.Recordset
    
    vlFecTras = "99991231"
    
    
    RS.CursorLocation = adUseClient
      RS.Open "PP_LISTA_BIENVENIDA.LISTAR('" & NumPolSel & "','" & dPol.tippension & "','" & "10" & "','" & "99" & "', '" & str(dPol.Num_Endoso) & "')", vgConexionBD, adOpenForwardOnly, adLockReadOnly
    
    Dim LNGa As Long
    LNGa = CreateFieldDefFile(RS, Replace(UCase(strRpt & "Estructura\PD_Rpt_PolizaBien.rpt"), ".RPT", ".TTX"), 1)
    
    
    
       If objRep.CargaReporte_toPdf(strRpt & "", "PD_Rpt_PolizaBien.rpt", "Informe de Liquidación de Rentas Vitalicias", "", 0, RS, True, "c:\PDFS\hola.pdf", _
                                 ArrFormulas("NombreCompaniaCorto", vgNombreCortoCompania), _
                            ArrFormulas("Nombre", vgNombreApoderado), _
                            ArrFormulas("Cargo", vgCargoApoderado), _
                            ArrFormulas("Sucursal", "Surquillo"), _
                            ArrFormulas("fec_trasp", vlFecTras)) = False Then
                            
                            

            
        MsgBox "No se pudo abrir el reporte", vbInformation
        Exit Sub
    End If
  
End Sub
Private Sub cmdReporteBienvenida_Click()
   ' Call RptBienvenida
   
   Call RptPoliza
 
  
End Sub

Private Sub Form_Load()
    DTPDesde.Value = Date
    Me.DTPHasta.Value = Date
    
GrdPol.Row = 0: GrdPol.Col = 0: GrdPol.ColWidth(0) = 1000: GrdPol.Text = "Poliza"
GrdPol.Row = 0: GrdPol.Col = 1: GrdPol.ColWidth(1) = 3000: GrdPol.Text = "Nombres"
GrdPol.Row = 0: GrdPol.Col = 2: GrdPol.ColWidth(2) = 2550: GrdPol.Text = "Tipo Pensión"
GrdPol.Row = 0: GrdPol.Col = 3: GrdPol.ColWidth(3) = 2550: GrdPol.Text = "Tipo Renta"
GrdPol.Row = 0: GrdPol.Col = 4: GrdPol.ColWidth(4) = 1550: GrdPol.Text = "Modalidad"
GrdPol.Row = 0: GrdPol.Col = 5: GrdPol.ColWidth(5) = 1500: GrdPol.Text = "Meses Garantizados"

End Sub

Private Sub OptRangoFechas_Click()

    If OptHoy.Value Then
        DTPDesde.Value = Date
        DTPHasta.Value = Date
        
        DTPDesde.Enabled = False
        DTPHasta.Enabled = False
    Else
    
      DTPDesde.Enabled = True
      DTPHasta.Enabled = True
    
    End If

End Sub
Private Sub sp_kitPolizaEmitidas(ByVal pnum_poliza As String, _
                                 ByRef pListaPdf As String, _
                                 ByRef pnumError As Integer, _
                                 ByRef pnumMsg As String)
                                   
                                   
                Dim conn    As ADODB.Connection
                Dim RS As ADODB.Recordset
                Dim objCmd As ADODB.Command
                
                Set conn = New ADODB.Connection
                Set RS = New ADODB.Recordset
                Set objCmd = New ADODB.Command
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
                
                Dim param1 As ADODB.Parameter
                Dim param2 As ADODB.Parameter
                Dim param3 As ADODB.Parameter
                Dim param4 As ADODB.Parameter
                
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMARRVV.sp_kitPolizaEmitidas"
                objCmd.CommandType = adCmdStoredProc
                
               
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 10, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("p_outLstPDF", adVarChar, adParamOutput, 300)
                objCmd.Parameters.Append param2
                                            
                Set param3 = objCmd.CreateParameter("p_outNumError", adInteger, adParamOutput)
                objCmd.Parameters.Append param3
                
                Set param4 = objCmd.CreateParameter("p_outMsgError", adVarChar, adParamOutput, 300)
                objCmd.Parameters.Append param4
     
                Set RS = objCmd.Execute
                
                If Not IsNull(objCmd.Parameters.Item("p_outLstPDF").Value) Then
                    pListaPdf = Trim(objCmd.Parameters.Item("p_outLstPDF").Value)
                Else
                    pListaPdf = ""
                End If
                  
                
                If Not IsNull(objCmd.Parameters.Item("p_outNumError").Value) Then
                    pnumError = Trim(objCmd.Parameters.Item("p_outNumError").Value)
                Else
                    pnumError = 0
                End If
                                  
            
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    pnumMsg = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    pnumMsg = ""
                End If
                
                conn.Close
                Set objCmd = Nothing
                Set RS = Nothing
                Set conn = Nothing
  
End Sub
Public Sub EliminarArchivo(ByVal Archivo As String)
Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If (fso.FileExists(Archivo)) Then
        Kill Archivo
    End If

End Sub

