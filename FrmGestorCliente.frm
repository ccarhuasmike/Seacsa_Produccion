VERSION 5.00
Begin VB.Form FrmGestorCliente 
   Caption         =   "Gestor Cliente"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Iniciar Proceso"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblmensaje 
      Caption         =   "Label1"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
End
Attribute VB_Name = "FrmGestorCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EnviaDatosGestionCliente()


   Dim rs As ADODB.Recordset
   Dim conn As ADODB.Connection
   
   Set rs = New ADODB.Recordset
   Set conn = New ADODB.Connection
      
   Dim RS2 As ADODB.Recordset
   Dim conn2 As ADODB.Connection
  
    Dim Texto  As String
    Dim cadena As String
  
   
   Call DataPrincipal("1", rs, conn)
                     Screen.MousePointer = 11
    
   
               While Not rs.EOF
                    lblmensaje.Caption = "Procesando..." & rs!Num_Poliza & " Orden:" & rs!Num_Orden
                    Me.Refresh
                    
                
                            Texto = "{" + Chr(13) + _
                            "p_CodAplicacion : 'SEACSA', " + Chr(13) + _
                            "p_TipOper : 'INS', " + Chr(13) + _
                            "P_NUSERCODE : '" + rs!P_NUSERCODE + "', " + Chr(13) + _
                            "P_NIDDOC_TYPE : '" + rs!P_NIDDOC_TYPE + "', " + Chr(13) + _
                            "P_SIDDOC : '" + rs!P_SIDDOC + "', " + Chr(13) + _
                            "P_SFIRSTNAME : '" + rs!P_SFIRSTNAME + "', " + Chr(13) + _
                            "P_SLASTNAME : '" + rs!P_SLASTNAME + "', " + Chr(13) + _
                            "P_SLASTNAME2 : '" + rs!P_SLASTNAME2 + "', " + Chr(13) + _
                            "P_SLEGALNAME : '" + rs!P_SLEGALNAME + "', " + Chr(13) + _
                            "P_SSEXCLIEN : '" + rs!P_SSEXCLIEN + "', " + Chr(13) + _
                            "P_NINCAPACITY : '" + rs!P_NINCAPACITY + "', " + Chr(13) + _
                            "P_DBIRTHDAT : '" + rs!P_DBIRTHDAT + "', " + Chr(13)
          
                            Texto = Texto + _
                            "p_DINCAPACITY : '" + CambiarVacioxNull(rs!p_DINCAPACITY) + "', " + Chr(13) + _
                            "P_NSPECIALITY : '99', " + Chr(13) + _
                            "P_NCIVILSTA : '" + CambiarVacioxNull(rs!P_NCIVILSTA) + "', " + Chr(13) + _
                            "P_NTITLE : '99', " + Chr(13) + _
                            "P_NAFP : '" + rs!P_NAFP + "', " + Chr(13) + _
                            "P_NNATIONALITY : '" + CambiarVacioxNull(rs!P_NNATIONALITY) + "', " + Chr(13) + _
                            "P_SBAJAMAIL_IND  : '" + CambiarVacioxNull(rs!P_SBAJAMAIL_IND) + "', " + Chr(13) + _
                            "P_SPROTEG_DATOS_IND  : '" + CambiarVacioxNull(rs!P_SPROTEG_DATOS_IND) + "', " + Chr(13) + _
                            "P_SISCLIENT_IND : '1', " + Chr(13) + _
                            "P_SISRENIEC_IND : '2' , " + Chr(13)
                            
                            Texto = Texto + "EListAddresClient : [ "
                   
                                  If Len(Trim(rs!pDireccionConcat)) > 0 Then
                                
                                    Texto = Texto + _
                                      "{" + Chr(13) + _
                                      "P_SRECTYPE : '2', " + Chr(13) + _
                                      "P_NCOUNTRY : '604', " + Chr(13) + _
                                      "P_NPROVINCE : '" + rs!P_NPROVINCE + "', " + Chr(13) + _
                                      "P_NLOCAL : '" + rs!P_NLOCAL + "', " + Chr(13) + _
                                      "P_NMUNICIPALITY : '" + rs!P_NMUNICIPALITY + "', " + Chr(13) + _
                                      "P_STI_DIRE : '" + rs!P_STI_DIRE + "', " + Chr(13) + _
                                      "P_SNOM_DIRECCION : '" + rs!P_SNOM_DIRECCION + "', " + Chr(13) + _
                                      "P_SNUM_DIRECCION : '" + rs!P_SNUM_DIRECCION + "', " + Chr(13) + _
                                      "P_STI_BLOCKCHALET : '" + rs!P_STI_BLOCKCHALET + "', " + Chr(13) + _
                                      "P_SBLOCKCHALET : '" + rs!P_SBLOCKCHALET + "', " + Chr(13) + _
                                      "P_STI_INTERIOR : '" + rs!P_STI_INTERIOR + "', " + Chr(13) + _
                                      "P_SNUM_INTERIOR : '" + rs!P_SNUM_INTERIOR + "', " + Chr(13) + _
                                      "P_STI_CJHT : '" + rs!P_STI_CJHT + "', " + Chr(13) + _
                                      "P_SNOM_CJHT : '" + rs!P_SNOM_CJHT + "', " + Chr(13) + _
                                      "P_SETAPA  : '" + rs!P_SETAPA + "', " + Chr(13) + _
                                      "P_SMANZANA : '" + rs!P_SMANZANA + "', " + Chr(13) + _
                                      "P_SLOTE : '" + rs!P_SLOTE + "', " + Chr(13) + _
                                      "P_SREFERENCIA : '" + rs!P_SREFERENCIA + "'" + Chr(13) + _
                                      "} "
                                
                                End If
                                  
                            
                            
                 Texto = Mid(Texto, 1, Len(Texto) - 1)
                 Texto = Texto + "],EListPhoneClient : [ "
                 
                If Len(Trim(rs!P_SPHONE)) > 0 Then
                
                     Texto = Texto + _
                     "{" + Chr(13) + _
                     "P_NAREA_CODE : '" + Trim(rs!P_NAREA_CODE) + "', " + Chr(13) + _
                     "P_SPHONE : '" + rs!P_SPHONE + "', " + Chr(13) + _
                     "P_NPHONE_TYPE : '" + rs!P_NPHONE_TYPE + "'" + Chr(13) + _
                     "},"
                     
                End If
                
                If Len(Trim(rs!P_SPHONE2)) > 0 Then
                     Texto = Texto + _
                     "{" + Chr(13) + _
                     "P_NAREA_CODE : '" + Trim(rs!P_NAREA_CODE2) + "', " + Chr(13) + _
                     "P_SPHONE : '" + rs!P_SPHONE2 + "', " + Chr(13) + _
                     "P_NPHONE_TYPE : '" + rs!P_NPHONE_TYPE2 + "'" + Chr(13) + _
                     "} "
                                
                End If
             
                        
               Texto = Mid(Texto, 1, Len(Texto) - 1)
               'JSON DATOS EMAIL_

                Texto = Texto + "],EListEmailClient : [ "
                
                If Len(Trim(rs!P_SE_MAIL)) > 0 Then
        
                   Texto = Texto + _
                   "{" + Chr(13) + _
                   "P_NROW : '1', " + Chr(13) + _
                   "P_SRECTYPE : '4', " + Chr(13) + _
                   "P_SE_MAIL : '" + rs!P_SE_MAIL + "'  " + Chr(13) + _
                   "} "
                   
                End If
        
     
               Texto = Mid(Texto, 1, Len(Texto) - 1)
               'JSON DATOS CONTACTO_
               Texto = Texto + "],EListContactClient : [ "
               
               
               If rs!P_NAFP = "99" Then
               
               Call DataBeneficiarios("2", rs!Num_Poliza, RS2, conn2)
                
                   While Not RS2.EOF
                      Texto = Texto + _
                       "{" + Chr(13) + _
                       "P_NTIPCONT : '" & RS2!P_NAFP & "', " + Chr(13) + _
                       "P_NIDDOC_TYPE :'" & RS2!P_NIDDOC_TYPE & "', " + Chr(13) + _
                       "P_SIDDOC :'" & RS2!P_SIDDOC & "', " + Chr(13) + _
                       "P_SNOMBRES :'" & RS2!P_SFIRSTNAME & "', " + Chr(13) + _
                       "P_SAPEPAT :'" & RS2!P_SLASTNAME & "', " + Chr(13) + _
                       "P_SAPEMAT :'" & RS2!P_SLASTNAME2 & "', " + Chr(13) + _
                       "P_SPHONE :'" & RS2!P_SPHONE & "'" + Chr(13) + _
                       "},"
                
                       RS2.MoveNext
                   Wend
                   
                    
                    Set RS2 = Nothing
                    Set conn2 = Nothing
                   
                   Texto = Mid(Texto, 1, Len(Texto) - 1)
                
              End If
              
             
              Texto = Texto + _
                "]," + Chr(13) + _
                "EListCIIUClient : Null " + Chr(13) + _
                "}" + Chr(13)
            
        
             'cadena = "http://10.10.1.51/WSGClientesDesarrollo/Api/Cliente/ValidarCliente"
             cadena = "https://soatservicios.protectasecurity.pe/WSGestorCliente/Api/Cliente/ValidarCliente"
             
           If EnviarGestorCliente1(cadena, Texto, rs!Num_Poliza, rs!Num_Orden) = True Then
             'cadena = "http://10.10.1.51/WSGClientesDesarrollo/Api/Cliente/GestionarCliente"
             cadena = "https://soatservicios.protectasecurity.pe/WSGestorCliente/Api/Cliente/GestionarCliente"
             
              If EnviarGestorCliente1(cadena, Texto, rs!Num_Poliza, rs!Num_Orden) = False Then
                
                'MsgBox ("Error en Poliza" & RS!Num_Poliza)
               
              End If
          
           End If
          
          rs.MoveNext

        Wend
        
         Screen.MousePointer = 0
    
    Set rs = Nothing
    Set conn = Nothing
            
    MsgBox ("Proceso Finalizado")
End Sub
Function EnviarGestorCliente1(ByVal cadCnx As String, ByVal cadJson As String, ByVal Num_Poliza As String, ByVal Num_Orden As String) As Boolean

            Dim sInputJson As String
            Dim httpURL As Object
            Dim response As String
             Dim valor As Object
             Set valor = json.parse(cadJson)
             sInputJson = json.toString(valor)
             
             Set httpURL = CreateObject("WinHttp.WinHttpRequest.5.1")

             httpURL.Open "POST", cadCnx, False
             httpURL.SetRequestHeader "Content-Type", "application/json"
             httpURL.Send sInputJson
             response = httpURL.ResponseText
              
              
            Dim p As Object
            Dim CodReturn As String
            Dim RspMensaje As String
            Dim pflag As String
            Set p = json.parse(response)
            RspMensaje = p.Item("P_SMESSAGE")
            CodReturn = p.Item("P_NCODE")
            If CodReturn = "1" Then
                Mensaje = RspMensaje
                EnviarGestorCliente1 = False
                pflag = "N"
            Else
                EnviarGestorCliente1 = True
                pflag = "S"
            End If
            
            Call RegistraLog(Num_Poliza, Num_Orden, pflag, response)
            
End Function


Private Sub DataPrincipal(ByVal tipo As String, ByRef rs As ADODB.Recordset, ByRef Con As ADODB.Connection)
            
              
                Dim objCmd As ADODB.Command
                
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
                
                objCmd.CommandText = "PKG_ActualizacionGestor.SP_ActGestorPrincipal"
                objCmd.CommandType = adCmdStoredProc
                
                            
        
                Set rs = objCmd.Execute
                
        

End Sub

Private Sub DataBeneficiarios(ByVal tipo As String, ByVal poliza As String, rs As ADODB.Recordset, ByRef Con As ADODB.Connection)
            
              
                Dim objCmd As ADODB.Command
                
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
                
                objCmd.CommandText = "PKG_ActualizacionGestor.SP_ActGestorBeneficiarios"
                objCmd.CommandType = adCmdStoredProc
                
                    
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 10, poliza)
                objCmd.Parameters.Append param1
        
                Set rs = objCmd.Execute
         
End Sub
Private Function CambiarVacioxNull(ByVal valor As String) As String
If Trim(valor) = "" Then
 valor = ""
End If
CambiarVacioxNull = valor
End Function
Private Sub RegistraLog(ByVal poliza As String, _
                        ByVal Num_Orden As Integer, _
                        ByVal pflag As String, _
                        ByVal pmensaje As String)
            
              
                Dim objCmd As ADODB.Command
                Dim rs As ADODB.Recordset
                Dim conn As ADODB.Connection
                    
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
                
                objCmd.CommandText = "PKG_ActualizacionGestor.RegistraResultado"
                objCmd.CommandType = adCmdStoredProc
                
                    
                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 15, poliza)
                objCmd.Parameters.Append param1
                
                Set param2 = objCmd.CreateParameter("pnum_orden", adInteger, adParamInput, 10, Num_Orden)
                objCmd.Parameters.Append param2
                
                Set param3 = objCmd.CreateParameter("pflag", adVarChar, adParamInput, 1, pflag)
                objCmd.Parameters.Append param3
                
                Set param4 = objCmd.CreateParameter("pmensaje", adVarChar, adParamInput, 1000, pmensaje)
                objCmd.Parameters.Append param4
        
                Set rs = objCmd.Execute
                
                   
         conn.Close
        Set rs = Nothing
        Set conn = Nothing

    
    
    
                
         
End Sub


Private Sub cmdProcesar_Click()
    Call EnviaDatosGestionCliente
End Sub

Private Sub Command1_Click()
Dim rest As New ChilkatRest
Dim success As Long

' URL: https://soatservicios.protectasecurity.pe/WSGestorCliente/Api/Cliente/ValidarCliente
Dim bTls As Long
bTls = 1
Dim port As Long
port = 443
Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.Connect("soatservicios.protectasecurity.pe", port, bTls, bAutoReconnect)
If (success <> 1) Then
    Debug.Print "ConnectFailReason: " & rest.ConnectFailReason
    Debug.Print rest.LastErrorText
    Exit Sub
End If


Dim json As New ChilkatJsonObject
success = json.UpdateString("p_CodAplicacion", "SEACSA")
success = json.UpdateString("p_TipOper", "INS")
success = json.UpdateString("P_NUSERCODE", vgUsuario)
success = json.UpdateString("P_NIDDOC_TYPE", "1")
success = json.UpdateString("P_SIDDOC", "09996107")
success = json.UpdateString("P_SFIRSTNAME", "MIGUEL EDUARDO")
success = json.UpdateString("P_SLASTNAME", "ALEGRE")
success = json.UpdateString("P_SLASTNAME2", "ROSAS")
success = json.UpdateString("P_SLEGALNAME", "MIGUEL EDUARDO")
success = json.UpdateString("P_SSEXCLIEN", "M")
success = json.UpdateString("P_NINCAPACITY", "N")
success = json.UpdateString("P_DBIRTHDAT", "12-06-1967")
success = json.UpdateString("p_DINCAPACITY", "")
success = json.UpdateString("P_NSPECIALITY", "99")
success = json.UpdateString("P_NCIVILSTA", "S")
success = json.UpdateString("P_NTITLE", "99")
success = json.UpdateString("P_NAFP", "99")
success = json.UpdateString("P_NNATIONALITY", "1")
success = json.UpdateString("P_SBAJAMAIL_IND", "2")
success = json.UpdateString("P_SPROTEG_DATOS_IND", "2")
success = json.UpdateString("P_SISCLIENT_IND", "1")
success = json.UpdateString("P_SISRENIEC_IND", "2")
success = json.UpdateString("EListAddresClient[0].P_SRECTYPE", "2")
success = json.UpdateString("EListAddresClient[0].P_NCOUNTRY", "604")
success = json.UpdateString("EListAddresClient[0].P_NPROVINCE", "14")
success = json.UpdateString("EListAddresClient[0].P_NLOCAL", "1401")
success = json.UpdateString("EListAddresClient[0].P_NMUNICIPALITY", "140103")
success = json.UpdateString("EListAddresClient[0].P_STI_DIRE", "03")
success = json.UpdateString("EListAddresClient[0].P_SNOM_DIRECCION", "GUADALAJARA")
success = json.UpdateString("EListAddresClient[0].P_SNUM_DIRECCION", "292")
success = json.UpdateString("EListAddresClient[0].P_STI_BLOCKCHALET", "")
success = json.UpdateString("EListAddresClient[0].P_SBLOCKCHALET", "")
success = json.UpdateString("EListAddresClient[0].P_STI_INTERIOR", "")
success = json.UpdateString("EListAddresClient[0].P_SNUM_INTERIOR", "")
success = json.UpdateString("EListAddresClient[0].P_STI_CJHT", "01")
success = json.UpdateString("EListAddresClient[0].P_SNOM_CJHT", "MAYORAZGO")
success = json.UpdateString("EListAddresClient[0].P_SETAPA", "3")
success = json.UpdateString("EListAddresClient[0].P_SMANZANA", "")
success = json.UpdateString("EListAddresClient[0].P_SLOTE", "")
success = json.UpdateString("EListAddresClient[0].P_SREFERENCIA", "")
success = json.UpdateString("EListPhoneClient[0].P_TIPOPER", "DEL")
success = json.UpdateString("EListPhoneClient[0].P_NAREA_CODE", "1")
success = json.UpdateString("EListPhoneClient[0].P_SPHONE", "")
success = json.UpdateString("EListPhoneClient[0].P_NPHONE_TYPE", "4")
success = json.UpdateString("EListPhoneClient[1].P_NAREA_CODE", "1")
success = json.UpdateString("EListPhoneClient[1].P_SPHONE", "992367123")
success = json.UpdateString("EListPhoneClient[1].P_NPHONE_TYPE", "4")
success = json.UpdateString("EListPhoneClient[2].P_NAREA_CODE", "")
success = json.UpdateString("EListPhoneClient[2].P_SPHONE", "998703878")
success = json.UpdateString("EListPhoneClient[2].P_NPHONE_TYPE", "2")
success = json.UpdateString("EListEmailClient[0].P_NROW", "1")
success = json.UpdateString("EListEmailClient[0].P_SRECTYPE", "4")
success = json.UpdateString("EListEmailClient[0].P_SE_MAIL", "guicella.rubina1@gmail.com")
success = json.UpdateString("EListContactClient[0].P_TIPOPER", "DEL")
success = json.UpdateString("EListContactClient[0].P_NTIPCONT", "10")
success = json.UpdateString("EListContactClient[0].P_NIDDOC_TYPE", "1")
success = json.UpdateString("EListContactClient[0].P_SIDDOC", "10137644")
success = json.UpdateString("EListContactClient[1].P_NTIPCONT", "10")
success = json.UpdateString("EListContactClient[1].P_NIDDOC_TYPE", "1")
success = json.UpdateString("EListContactClient[1].P_SIDDOC", "10137644")
success = json.UpdateString("EListContactClient[1].P_SNOMBRES", "GUICELLA HARLETTY")
success = json.UpdateString("EListContactClient[1].P_SAPEPAT", "RUBINA")
success = json.UpdateString("EListContactClient[1].P_SAPEMAT", "BOSSIO")
success = json.UpdateString("EListContactClient[1].P_SPHONE", "2723652")

success = rest.AddHeader("Content-Type", "application/json")

Dim sbRequestBody As New ChilkatStringBuilder
success = json.EmitSb(sbRequestBody)
Dim sbResponseBody As New ChilkatStringBuilder
success = rest.FullRequestSb("POST", "/WSGestorCliente/Api/Cliente/ValidarCliente", sbRequestBody, sbResponseBody)
If (success <> 1) Then
    Debug.Print rest.LastErrorText
    Exit Sub
End If

Dim respStatusCode As Long
respStatusCode = rest.ResponseStatusCode
Debug.Print "response status code = " & respStatusCode
If (respStatusCode >= 400) Then
    Debug.Print "Response Status Code = " & respStatusCode
    Debug.Print "Response Header:"
    Debug.Print rest.ResponseHeader
    Debug.Print "Response Body:"
    Debug.Print sbResponseBody.GetAsString()
    Exit Sub
End If
End Sub

