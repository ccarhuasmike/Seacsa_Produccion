VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Frm_StockFirmas 
   Caption         =   "Firmas Stock"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar barra 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdIniciarProceso 
      Caption         =   "Iniciar Proceso"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6615
      Begin VB.CheckBox chjJson 
         Caption         =   "Generar Archivos"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label lblmensaje 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Env�o de Stock de Firmas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Frm_StockFirmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type Firmantes_Stock
    nacionalidad As String
    tipodocumento As String
    numerodocumento As String
    nombre As String
    apellidos As String
    correo As String
    celular As String
    Direccion As String
    genero As String
    departamento As String
    provincia As String
    distrito As String
    ID_FIRMANTE As String
    TIPO_FIRMA As String
    FIRMA_TIPO As String
    PARENTESCO As String
    tipo As String
    MENOREDAD As String
   
End Type

Private Type DatosPoli
    tiporenta As String
    fechacontrato As String
    Num_Poliza As String

End Type

Dim LST_Firmastock() As Firmantes_Stock
     



Private Sub cmdIniciarProceso_Click()

  Dim glob As New ChilkatGlobal
  Dim success As Long
  Dim vTokenDavidCloud As String
  Dim VidDocumento As String
  Dim vDatosPoli As DatosPoli
  Dim pToken As String
  Dim vNum_poliza As String
  Dim vTipoRenta As String
  
  Dim rs As ADODB.Recordset
  Dim conn As ADODB.Connection
  Dim v_id_transac As Integer
  
  Dim x As Integer
  Dim TotalRegistros As Integer
  Dim TotalProcesados As Integer
  
   cmdIniciarProceso.Enabled = False
  
  
  Set rs = New ADODB.Recordset
  Set conn = New ADODB.Connection
  
  
   success = glob.UnlockBundle("GVNCRZ.CB1032023_x4BpcXzLDR4D")
    If (success <> 1) Then
        Debug.Print glob.LastErrorText
        Exit Sub
    End If
        
        
     Call DataPrincipal("1", rs, conn)
     Screen.MousePointer = 11
     
     TotalRegistros = rs.RecordCount
     TotalProcesados = 0
     barra.Value = 0
     barra.Max = TotalRegistros
  
     While Not rs.EOF()
          Open App.Path & "\Json_poliza_" & rs!Num_Poliza & ".txt" For Output As #1
          Dim FlgError As Boolean
          
          FlgError = False
                   
           
           Me.Refresh
           
           vNum_poliza = rs!Num_Poliza
           vTipoRenta = rs!Tipo_Renta
           
           vTokenDavidCloud = TokenDavicloud(v_id_transac, vNum_poliza)
     
           FlgError = Get_Firmantes(vNum_poliza, vTipoRenta, v_id_transac)  'LLENA LOS FIRMANTES EN LST_Firmastock
   
            
          If Not FlgError Then
          
                For x = 0 To UBound(LST_Firmastock)
                    Call Reg_Firmantes_stock(vTokenDavidCloud, LST_Firmastock(x), v_id_transac)
                 Next
          
                  vDatosPoli = datosRegDocumentos(vNum_poliza)
                  VidDocumento = RegistraDoc(vDatosPoli, vTokenDavidCloud, v_id_transac)
                  
                  
                  Call IniciarProcesoFirma(VidDocumento, v_id_transac, vTokenDavidCloud, LST_Firmastock)
                  TotalProcesados = TotalProcesados + 1
              
          End If
          
        
        barra.Value = TotalProcesados
        lblmensaje = "Procesados " & TotalProcesados & " de " & TotalRegistros
                 
                 Me.Refresh
          
          rs.MoveNext
     
        Close #1
     Wend
     
  
     Screen.MousePointer = 0
     Me.Refresh
     MsgBox "Proceso Terminado", vbInformation, "Stock de Firmas"
     
   
End Sub

Private Function TokenDavicloud(ByRef v_id_transac As Integer, ByVal vNum_poliza As String) As String

        Dim rest As New ChilkatRest
        Dim success As Long
        
        ' URL: https://test36.davicloud.com/API/sign/v1/api_rest.php/access_token
        Dim bTls As Long
        bTls = 1
        Dim port As Long
        port = 443
        Dim bAutoReconnect As Long
        bAutoReconnect = 1
        'PRUEBAS
        'success = rest.Connect("test36.davicloud.com", port, bTls, bAutoReconnect)
        'PROD
        success = rest.Connect("protecta.davicloud.com", port, bTls, bAutoReconnect)
        
        If (success <> 1) Then
            Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
            Debug.Print rest.LastErrorText
            Exit Function
        End If
        
        success = rest.AddHeader("Id-Organizacion", "PROTECTA")
        'PRUEBAS
        ' success = rest.AddHeader("Authorization", "Basic QVBJUFJPVEVDVEE6UHJvdGVjdGEuMjAyMSM=")
        'PROD
        success = rest.AddHeader("Authorization", "Basic QVBJUFJPVEVDVEE6UHJvdGVjdGFQcmQuMjAyMSM=")
        
        'PRUEBAS
        ' success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "bca9b2ead50d08bbc075e7ba661c6d3f")
        'PROD
        success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "c12b12d9d8b815adc2733e72aee4540b")
        
        Dim sbResponseBody As New ChilkatStringBuilder
        success = rest.FullRequestNoBodySb("POST", "/API/sign/v1/api_rest.php/access_token", sbResponseBody)
        If (success <> 1) Then
            Debug.Print rest.LastErrorText
            Exit Function
        End If
        
        Dim respStatusCode As Long
        respStatusCode = rest.ResponseStatusCode
        Debug.Print "response status code = " & respStatusCode
        If (respStatusCode >= 400) Then
            v_id_transac = SP_LOG_API_DOC(vgUsuario, vNum_poliza, "protecta.davicloud.com/API/sign/v1/api_rest.php/access_token", "1", sbResponseBody.GetAsString)
        '    Debug.Print "Response Status Code = " & respStatusCode
        '    Debug.Print "Response Header:"
        '    Debug.Print rest.ResponseHeader
        '    Debug.Print "Response Body:"
        '    MsgBox sbResponseBody.GetAsString()
            Exit Function
        Else
          v_id_transac = SP_LOG_API_DOC(vgUsuario, vNum_poliza, "protecta.davicloud.com/API/sign/v1/api_rest.php/access_token", "0", "Tocken davidCloud obtenido correctamente")
        End If
        
        Dim Rpta() As String
        Dim vTokenDavidCloud As String
        
        Dim p As Object
        Set p = json.parse(sbResponseBody.GetAsString())
        vTokenDavidCloud = p.Item("access_token")
                
        
        TokenDavicloud = vTokenDavidCloud
        
        
        End Function
 
Private Sub Reg_Firmantes_stock(ByVal pToken As String, ByRef vFirmante As Firmantes_Stock, ByVal v_id_transac As Integer)
        
            Dim rest As New ChilkatRest

            Dim success As Long
            
            ' URL: https://test36.davicloud.com/API/sign/v1/api_rest.php/registra_firmante
            Dim bTls As Long
            bTls = 1
            Dim port As Long
            port = 443
            Dim bAutoReconnect As Long
            bAutoReconnect = 1
            'PRUEBAS
            'success = rest.Connect("test36.davicloud.com", port, bTls, bAutoReconnect)
            'PROD
            success = rest.Connect("protecta.davicloud.com", port, bTls, bAutoReconnect)
            If (success <> 1) Then
                Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
                Debug.Print rest.LastErrorText
                Exit Sub
            End If
                 
            Dim jsonStock As New ChilkatJsonObject
            success = jsonStock.UpdateString("nacionalidad", vFirmante.nacionalidad)
            success = jsonStock.UpdateString("tipodocumento", vFirmante.tipodocumento)
            success = jsonStock.UpdateString("numerodocumento", vFirmante.numerodocumento)
            success = jsonStock.UpdateString("nombre", vFirmante.nombre)
            success = jsonStock.UpdateString("apellidos", vFirmante.apellidos)
            success = jsonStock.UpdateString("correo", vFirmante.correo)
            success = jsonStock.UpdateString("celular", vFirmante.celular)
            success = jsonStock.UpdateString("direccion", vFirmante.Direccion)
            success = jsonStock.UpdateString("genero", vFirmante.genero)
            success = jsonStock.UpdateString("departamento", vFirmante.departamento)
            success = jsonStock.UpdateString("provincia", vFirmante.provincia)
            success = jsonStock.UpdateString("distrito", vFirmante.distrito)
            
            success = rest.AddHeader("Id-Organizacion", "PROTECTA")
            success = rest.AddHeader("Content-Type", "application/json")
            success = rest.AddHeader("Authorization", "Bearer " & pToken)
            'PRUEBAS
            'success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "bca9b2ead50d08bbc075e7ba661c6d3f")
            'PROD
            success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "c12b12d9d8b815adc2733e72aee4540b")
            
            Dim sbRequestBody As New ChilkatStringBuilder
            success = jsonStock.EmitSb(sbRequestBody)
            Dim sbResponseBody As New ChilkatStringBuilder
            
            If chjJson.Value = "1" Then
              Print #1, sbRequestBody.GetAsString
              Print #1, "******************************"
            
            Else
                success = rest.FullRequestSb("POST", "/API/sign/v1/api_rest.php/registra_firmante", sbRequestBody, sbResponseBody)
                
                
                
              
                
                If (success <> 1) Then
                    Debug.Print rest.LastErrorText
                    Exit Sub
                End If
                
                 Dim p As Object
                 Dim vID_FIRMANTE As String
                 
                 Set p = json.parse(sbResponseBody.GetAsString)
          
            
            
                Dim respStatusCode As Long
                respStatusCode = rest.ResponseStatusCode
                Debug.Print "response status code = " & respStatusCode
                If (respStatusCode >= 400) Then
                       SP_LOG_API_DOC2_STOCK vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/registra_firmante", "1", sbResponseBody.GetAsString
                    Debug.Print "Response Status Code = " & respStatusCode
                    Debug.Print "Response Header:"
                    Debug.Print rest.ResponseHeader
                    Debug.Print "Response Body:"
                    Debug.Print sbResponseBody.GetAsString()
                    Exit Sub
                Else
                                         vFirmante.ID_FIRMANTE = p.Item("idfirmante")
                     SP_LOG_API_DOC2_STOCK vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/registra_firmante", "0", p.Item("message") + " - " + vFirmante.tipo + " - " + vFirmante.numerodocumento + " - " + vFirmante.nombre + " " + vFirmante.apellidos + " - " + vFirmante.correo
                End If
                
        End If
        

  End Sub
  

Private Function Get_Firmantes(ByVal pnum_poliza As String, ByVal ptipo_renta As String, ByVal pid_transac As String) As Boolean

                                   
                Dim conn    As ADODB.Connection
                Set conn = New ADODB.Connection
                Dim rs As ADODB.Recordset
                Dim objCmd As ADODB.Command
                Dim MensajeError As String
                
                
                Dim param1 As ADODB.Parameter
                Dim param2 As ADODB.Parameter
                Dim param3 As ADODB.Parameter
    
                
                Set rs = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
                Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")
                
                On Error GoTo ManejoError
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
                
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMAS_STOCK.sp_datos_poliza_stock"
                objCmd.CommandType = adCmdStoredProc
                
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 10, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("p_outNumError", adDouble, adParamOutput)
                objCmd.Parameters.Append param2
                
                Set param3 = objCmd.CreateParameter("p_outMsgError", adVarChar, adParamOutput, 200)
                objCmd.Parameters.Append param3
                
                       
                Set rs = objCmd.Execute
   
                
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                    Err.Raise "-2001", , MensajeError
                Else
                
                    If rs.RecordCount = 0 Then
                        MensajeError = "No se encontraron registros"
                        Err.Raise "-2001", , "No se encontraron registros"
                        
                    End If
                    
                    MensajeError = ""
                    Dim i As Integer
                    i = 0
                               
                  While Not rs.EOF()
                    i = i + 1
                    ReDim Preserve LST_Firmastock(i - 1)
                    
                     '  , TIPO, NUM_IDENBEN, GLS_NOMBEN, GLS_PATBEN, GLS_MATBEN, GLS_CORREOBEN, CELULAR, COD_SEXO, COD_TIPOIDENBEN, MENOREDAD
                    
                    
                    LST_Firmastock(i - 1).nacionalidad = "PE"
                    LST_Firmastock(i - 1).tipodocumento = IIf(IsNull(rs!Cod_TipoIdenBen), "", rs!Cod_TipoIdenBen)
                    LST_Firmastock(i - 1).numerodocumento = IIf(IsNull(rs!Num_IdenBen), "", rs!Num_IdenBen)
                    LST_Firmastock(i - 1).nombre = IIf(IsNull(rs!Gls_NomBen), "", rs!Gls_NomBen)
                    LST_Firmastock(i - 1).apellidos = IIf(IsNull(rs!Gls_PatBen), "", rs!Gls_PatBen) & " " & IIf(IsNull(rs!Gls_MatBen), "", rs!Gls_MatBen)
                    LST_Firmastock(i - 1).correo = IIf(IsNull(rs!Gls_CorreoBen), "", rs!Gls_CorreoBen)
                    LST_Firmastock(i - 1).celular = IIf(IsNull(rs!celular), "", rs!celular)
                    LST_Firmastock(i - 1).Direccion = IIf(IsNull(rs!Direccion), "", rs!Direccion)
                    LST_Firmastock(i - 1).genero = IIf(IsNull(rs!genero), "", rs!genero)
                    LST_Firmastock(i - 1).departamento = IIf(IsNull(rs!departamento), "", rs!departamento)
                    LST_Firmastock(i - 1).provincia = IIf(IsNull(rs!provincia), "", rs!provincia)
                    LST_Firmastock(i - 1).distrito = IIf(IsNull(rs!distrito), "", rs!distrito)
                    LST_Firmastock(i - 1).tipo = IIf(IsNull(rs!tipo), "", rs!tipo)
                    LST_Firmastock(i - 1).MENOREDAD = IIf(IsNull(rs!MENOREDAD), "", rs!MENOREDAD)
                    
                    
                    
                    Select Case LST_Firmastock(i - 1).tipo
                        Case "CONT"
                                  LST_Firmastock(i - 1).FIRMA_TIPO = "A"
                                 
                                  
                                  If ptipo_renta = "SOBREVIVENCIA" Then
                                    LST_Firmastock(i - 1).TIPO_FIRMA = "NF"
                                   Else
                                    LST_Firmastock(i - 1).TIPO_FIRMA = "FE"
                                  End If
            
                                  
                        Case "REP"
                                  LST_Firmastock(i - 1).FIRMA_TIPO = "R"
                                  LST_Firmastock(i - 1).PARENTESCO = "OTROS"
                                  
                                  If ptipo_renta = "SOBREVIVENCIA" Then
                                    LST_Firmastock(i - 1).TIPO_FIRMA = "FE"
                                   Else
                                    LST_Firmastock(i - 1).TIPO_FIRMA = "NF"
                                  End If
                                  
                        Case "BEN"
                                  LST_Firmastock(i - 1).FIRMA_TIPO = "B"
                                  LST_Firmastock(i - 1).TIPO_FIRMA = "NF"
                      
                    End Select
        
                    rs.MoveNext
   
    
                  Wend
   
               
                End If
                   
                             
                conn.Close
                Set objCmd = Nothing
                Set rs = Nothing
                Set conn = Nothing
                
                Get_Firmantes = False
                Exit Function
                
                
ManejoError:
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                    SP_LOG_API_DOC2_STOCK vgUsuario, pid_transac, "", "1", MensajeError
                    Get_Firmantes = True
                Else
                    MensajeError = ""
                End If
                'MsgBox Err.Description + MensajeError, vbCritical
                        
End Function

Private Function datosRegDocumentos(ByVal pnum_poliza As String) As DatosPoli

                Dim objCmd As ADODB.Command
                Dim rs As ADODB.Recordset
                Dim conn As ADODB.Connection
                Dim oDatosPoli As DatosPoli
                Dim MensajeError As String
                
                        
                Dim param1 As ADODB.Parameter
                Dim param2 As ADODB.Parameter
                Dim param3 As ADODB.Parameter
                    
                Set rs = New ADODB.Recordset
                Set conn = New ADODB.Connection
                
                Dim Texto As String
                
                
                Set conn = New ADODB.Connection
                Set rs = New ADODB.Recordset
                Set objCmd = New ADODB.Command
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
          
                               
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMAS_STOCK.datos_regDoc_stock"
                objCmd.CommandType = adCmdStoredProc
                

                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 10, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("p_outNumError", adDouble, adParamOutput)
                objCmd.Parameters.Append param2
                
                Set param3 = objCmd.CreateParameter("p_outMsgError", adVarChar, adParamOutput, 200)
                objCmd.Parameters.Append param3
                
                       
                Set rs = objCmd.Execute
   
                
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                     
                   While Not rs.EOF()
                  
                        oDatosPoli.fechacontrato = IIf(IsNull(rs!Tipo_Renta), "", rs!fechacontrato)
                        oDatosPoli.tiporenta = IIf(IsNull(rs!fechacontrato), "", rs!Tipo_Renta)
                        oDatosPoli.Num_Poliza = IIf(IsNull(rs!Num_Poliza), "", rs!Num_Poliza)
                                
                        rs.MoveNext
    
                   Wend
    
               
                End If
                
                conn.Close
                Set objCmd = Nothing
                Set rs = Nothing
                Set conn = Nothing
                datosRegDocumentos = oDatosPoli
                Exit Function
                
                
ManejoError:
                If Not IsNull(objCmd.Parameters.Item("p_outMsgError").Value) Then
                    MensajeError = Trim(objCmd.Parameters.Item("p_outMsgError").Value)
                Else
                    MensajeError = ""
                End If
                MsgBox Err.Description + MensajeError, vbCritical



End Function

Private Function RegistraDoc(ByRef oDatosPoli As DatosPoli, ByVal pToken As String, ByVal v_id_transac As Integer) As String

    Dim rest As New ChilkatRest
    Dim success As Long
    
    ' URL: https://test36.davicloud.com/API/sign/v1/api_rest.php/registra_documento_vitalicia_stock
    Dim bTls As Long
    bTls = 1
    Dim port As Long
    port = 443
    Dim bAutoReconnect As Long
    bAutoReconnect = 1
    'PRUEBAS
    'success = rest.Connect("test36.davicloud.com", port, bTls, bAutoReconnect)
    'PROD
    success = rest.Connect("protecta.davicloud.com", port, bTls, bAutoReconnect)
    If (success <> 1) Then
        Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
        Debug.Print rest.LastErrorText
        Exit Function
    End If
  
    
    Dim jsonCh As New ChilkatJsonObject
    success = jsonCh.UpdateString("documentotipo", "19")
    success = jsonCh.UpdateString("documentocreadorfirmacorreo", "juan.davila@bigdavi.com")
    success = jsonCh.UpdateString("documentodescripcion", "Stock RRVV")
    success = jsonCh.UpdateString("documentonombre", "")
    success = jsonCh.UpdateString("documentoclavepdf", "")
    success = jsonCh.UpdateString("documentocodigoqr", "0")
    success = jsonCh.UpdateString("documentonumpoliza", oDatosPoli.Num_Poliza)
    success = jsonCh.UpdateString("documentoformatoact", oDatosPoli.tiporenta)
    success = jsonCh.UpdateString("documentofechacontrato", oDatosPoli.fechacontrato)
    success = jsonCh.UpdateString("documentoproducto", "RV")
    
    success = rest.AddHeader("Id-Organizacion", "PROTECTA")
    success = rest.AddHeader("Content-Type", "application/json")
    success = rest.AddHeader("Authorization", "Bearer " & pToken)
    'PRUEBAS
    'success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "bca9b2ead50d08bbc075e7ba661c6d3f")
    'PROD
    success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "c12b12d9d8b815adc2733e72aee4540b")
    
    
    Dim sbRequestBody As New ChilkatStringBuilder
    success = jsonCh.EmitSb(sbRequestBody)
    Dim sbResponseBody As New ChilkatStringBuilder
    
    If chjJson.Value = "1" Then
      Print #1, sbRequestBody.GetAsString
      Print #1, "******************************"
    
    Else
    
                success = rest.FullRequestSb("POST", "/API/sign/v1/api_rest.php/registra_documento_vitalicia_stock", sbRequestBody, sbResponseBody)
                
            
                
                Dim VidDocumento As String
                Dim p As Object
                
                Set p = json.parse(sbResponseBody.GetAsString)
                
                If (success <> 1) Then
                    Debug.Print rest.LastErrorText
                    Exit Function
                
                End If
                     Dim respStatusCode As Long
                    respStatusCode = rest.ResponseStatusCode
                     
                     If (respStatusCode >= 400) Then
                      SP_LOG_API_DOC2_STOCK vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/registra_documento_vitalicia_stock", "1", sbResponseBody.GetAsString
            '            Debug.Print "Response Status Code = " & respStatusCode
            '            Debug.Print "Response Header:"
            '            Debug.Print rest.ResponseHeader
            '            Debug.Print "Response Body:"
            '            Debug.Print sbResponseBody.GetAsString()
                        Exit Function
                    Else
                        VidDocumento = p.Item("iddocumento")
                        RegistraDoc = VidDocumento
                        
                        SP_LOG_API_DOC2_STOCK vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/registra_documento_vitalicia_stock", "0", p.Item("message")
                    End If
        End If
        
                    
   
End Function

Private Sub DataPrincipal(ByVal tipo As String, ByRef rs As ADODB.Recordset, ByRef Con As ADODB.Connection)
            
              
                Dim objCmd As ADODB.Command
                
                Dim Texto As String
                
                Dim conn As ADODB.Connection
                                              
                Set conn = New ADODB.Connection
                Set rs = New ADODB.Recordset
                Set objCmd = New ADODB.Command
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
                
                               
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "PKG_API_FIRMAS_STOCK.sp_obtPolizasStock"
                objCmd.CommandType = adCmdStoredProc
               
        
                Set rs = objCmd.Execute
        

End Sub
 Private Function SP_LOG_API_DOC(ByVal p_usuario As String, _
                            ByVal vNum_poliza As String, _
                           ByVal p_urlapi As String, _
                           ByVal p_error As String, _
                           ByVal p_mensaje As String) As Integer

        Dim conn    As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim objCmd As ADODB.Command
        
        Dim param1 As ADODB.Parameter
        Dim param2 As ADODB.Parameter
        Dim param3 As ADODB.Parameter
        Dim param4 As ADODB.Parameter
        Dim param5 As ADODB.Parameter
        Dim param6 As ADODB.Parameter
        
        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Set objCmd = New ADODB.Command
        
           
        conn.Provider = "OraOLEDB.Oracle"
        conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
        conn.CursorLocation = adUseClient
        conn.Open
        
        Set objCmd = CreateObject("ADODB.Command")
        objCmd.ActiveConnection = conn
     
        objCmd.CommandText = "PKG_API_FIRMAS_STOCK.SP_LOG_API_DOC_STOCK"
        objCmd.CommandType = adCmdStoredProc
 
        Set param1 = objCmd.CreateParameter("p_usuario", adVarChar, adParamInput, 10, p_usuario)
        objCmd.Parameters.Append param1
        
        Set param2 = objCmd.CreateParameter("p_num_poliza", adVarChar, adParamInput, 12, vNum_poliza)
        objCmd.Parameters.Append param2
        
        Set param3 = objCmd.CreateParameter("p_id_transac", adDouble, adParamOutput)
        objCmd.Parameters.Append param3
        
        Set param4 = objCmd.CreateParameter("p_urlapi", adVarChar, adParamInput, 255, p_urlapi)
        objCmd.Parameters.Append param4
        
        Set param5 = objCmd.CreateParameter("p_error", adVarChar, adParamInput, 1, p_error)
        objCmd.Parameters.Append param5
        
        Set param6 = objCmd.CreateParameter("p_mensaje", adVarChar, adParamInput, 300, p_mensaje)
        objCmd.Parameters.Append param6
        
            
        objCmd.Execute
        
        SP_LOG_API_DOC = Trim(objCmd.Parameters.Item("p_id_transac").Value)
        
        
          
        conn.Close
        Set objCmd = Nothing
        Set rs = Nothing
        Set conn = Nothing
 

End Function
Private Sub SP_LOG_API_DOC2_STOCK(ByVal p_usuario As String, _
                          ByVal p_id_transac As Integer, _
                           ByVal p_urlapi As String, _
                           ByVal p_error As String, _
                           ByVal p_mensaje As String)

        Dim conn    As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim objCmd As ADODB.Command
        
        Dim param1 As ADODB.Parameter
        Dim param2 As ADODB.Parameter
        Dim param3 As ADODB.Parameter
        Dim param4 As ADODB.Parameter
        Dim param5 As ADODB.Parameter
        Dim param6 As ADODB.Parameter
        
        
        Set conn = New ADODB.Connection
        
        
        Set rs = New ADODB.Recordset
        Set objCmd = New ADODB.Command
        
        conn.Provider = "OraOLEDB.Oracle"
        conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
        conn.CursorLocation = adUseClient
        conn.Open
        
        Set objCmd = CreateObject("ADODB.Command")
        Set objCmd.ActiveConnection = conn
        
        objCmd.CommandText = "PKG_API_FIRMAS_STOCK.SP_LOG_API_DOC2_STOCK"
        objCmd.CommandType = adCmdStoredProc
   
        Set param1 = objCmd.CreateParameter("p_usuario", adVarChar, adParamInput, 10, p_usuario)
        objCmd.Parameters.Append param1
        
        Set param2 = objCmd.CreateParameter("p_id_transac", adDouble, adParamInput)
        param2.Value = p_id_transac
        objCmd.Parameters.Append param2
        
        Set param4 = objCmd.CreateParameter("p_urlapi", adVarChar, adParamInput, 255, p_urlapi)
        objCmd.Parameters.Append param4
        
        Set param5 = objCmd.CreateParameter("p_error", adVarChar, adParamInput, 1, p_error)
        objCmd.Parameters.Append param5
        
        Set param6 = objCmd.CreateParameter("p_mensaje", adVarChar, adParamInput, 300, p_mensaje)
        objCmd.Parameters.Append param6
        
        objCmd.Execute
    
        
        conn.Close
        Set objCmd = Nothing
        Set rs = Nothing
        Set conn = Nothing
 
End Sub
Private Sub IniciarProcesoFirma(ByVal VidDocumento As String, ByVal v_id_transac As String, ByVal vToken As String, ByRef lstFirmantes() As Firmantes_Stock)

        Dim rest As New ChilkatRest
        Dim success As Long
        Dim x As Integer
        
        
        ' URL: https://test36.davicloud.com/API/sign/v1/api_rest.php/proceso_firma_vitalicia
        Dim bTls As Long
        bTls = 1
        Dim port As Long
        port = 443
        Dim bAutoReconnect As Long
        bAutoReconnect = 1
        'PRUEBAS
        'success = rest.Connect("test36.davicloud.com", port, bTls, bAutoReconnect)
        'PROD
        success = rest.Connect("protecta.davicloud.com", port, bTls, bAutoReconnect)
        
        If (success <> 1) Then
            Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
            Debug.Print rest.LastErrorText
            Exit Sub
        End If
        
        Dim JSONchk As New ChilkatJsonObject
        
        success = JSONchk.UpdateString("iddocumento", VidDocumento)
        Dim NumFirmante As String
       
        For x = 0 To UBound(lstFirmantes)
            NumFirmante = "firmantes.firmante" & x + 1
      
               
            success = JSONchk.UpdateString(NumFirmante & ".idfirmante", lstFirmantes(x).ID_FIRMANTE)
            success = JSONchk.UpdateString(NumFirmante & ".tipofirma", lstFirmantes(x).TIPO_FIRMA)
            success = JSONchk.UpdateString(NumFirmante & ".firmatipo", lstFirmantes(x).FIRMA_TIPO)
            success = JSONchk.UpdateString(NumFirmante & ".parentesco", lstFirmantes(x).PARENTESCO)
            success = JSONchk.UpdateString(NumFirmante & ".menordeedad", lstFirmantes(x).MENOREDAD)
            
        Next
         
        success = rest.AddHeader("Id-Organizacion", "PROTECTA")
        success = rest.AddHeader("Content-Type", "application/json")
        success = rest.AddHeader("Authorization", "Bearer " & vToken)
        'PRUEBAS
        'success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "bca9b2ead50d08bbc075e7ba661c6d3f")
        'PROD
        success = rest.AddHeader("Jcdf-Apib-Subscription-Key", "c12b12d9d8b815adc2733e72aee4540b")
            
        Dim sbRequestBody As New ChilkatStringBuilder
        success = JSONchk.EmitSb(sbRequestBody)
        
        
        Dim sbResponseBody As New ChilkatStringBuilder
        
       If chjJson.Value = "1" Then
            Print #1, sbRequestBody.GetAsString
            Print #1, "******************************"
    
    Else
    
        success = rest.FullRequestSb("POST", "/API/sign/v1/api_rest.php/proceso_firma_vitalicia", sbRequestBody, sbResponseBody)

        
        If (success <> 1) Then
            Debug.Print rest.LastErrorText
            Exit Sub
        End If
        
        Dim respStatusCode As Long
        respStatusCode = rest.ResponseStatusCode
        
        Dim p As Object
        Set p = json.parse(sbResponseBody.GetAsString)
        
        Debug.Print "response status code = " & respStatusCode
        If (respStatusCode >= 400) Then
        SP_LOG_API_DOC2_STOCK vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/proceso_firma_vitalicia", "1", sbResponseBody.GetAsString
'            Debug.Print "Response Status Code = " & respStatusCode
'            Debug.Print "Response Header:"
'            Debug.Print rest.ResponseHeader
'            Debug.Print "Response Body:"
'            Debug.Print sbResponseBody.GetAsString()
            Exit Sub
        Else
             SP_LOG_API_DOC2_STOCK vgUsuario, v_id_transac, "protecta.davicloud.com/API/sign/v1/api_rest.php/proceso_firma_vitalicia", "0", p.Item("message")
        End If
        
   End If
        

End Sub

Private Sub Form_Load()
    lblmensaje = ""
    
End Sub


