VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmGeneraJson 
   Caption         =   "Json Envío Electrónico"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdJson 
      Caption         =   "Generar"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      Begin MSComctlLib.ProgressBar barProgres 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Generacion de Json Envio Electronico RRVV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Width           =   5655
      End
   End
End
Attribute VB_Name = "FrmGeneraJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataPrincipal(ByRef rs As ADODB.Recordset, ByRef Con As ADODB.Connection)
            
              
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
        
     
          
        conn.Close
        Set objCmd = Nothing
        Set rs = Nothing
        Set conn = Nothing
 

End Function


Private Sub CmdJson_Click()
                Dim rsPol As ADODB.Recordset
                Dim connPol As ADODB.Connection
                
                Set rsPol = New ADODB.Recordset
                Set connPol = New ADODB.Connection
  
        
                Call DataPrincipal(rsPol, connPol)
                
                barProgres.Max = rsPol.RecordCount
                barProgres.Value = 0
                
                While Not rsPol.EOF
                    GeneraJson (rsPol!Num_Poliza)
                    barProgres.Value = barProgres.Value + 1
                                  
                    rsPol.MoveNext
                 
                Wend
                
                MsgBox "Proceso Concluido, Los archivos fueron creados en " & App.Path, vbInformation, "Generacion Json RRVV"
                
                
End Sub

Private Sub GeneraJson(ByVal pnum_poliza As String)

                Dim Mensaje As String
                Dim StrDocument As String
                Dim StrFirmantes As String
        
                Dim conn    As ADODB.Connection
                Set conn = New ADODB.Connection
                Set rs = New ADODB.Recordset ' CreateObject("ADODB.Recordset")
                Set objCmd = New ADODB.Command ' CreateObject("ADODB.Command")
                
                conn.Provider = "OraOLEDB.Oracle"
                conn.ConnectionString = "PLSQLRSet=1;Data Source=" & vgNombreBaseDatos & ";" & "User ID=" & vgNombreUsuario & ";Password=" & vgPassWord & ";"
                conn.CursorLocation = adUseClient
                conn.Open
                
                Set objCmd = CreateObject("ADODB.Command")
                Set objCmd.ActiveConnection = conn
                
                objCmd.CommandText = "SP_generaJsonRRVV"
                objCmd.CommandType = adCmdStoredProc
                

                Set param1 = objCmd.CreateParameter("pnum_poliza", adVarChar, adParamInput, 10, pnum_poliza)
                objCmd.Parameters.Append param1
                
                Set rs = objCmd.Execute
                
    
                While Not rs.EOF
                
                    Open App.Path & "\Json_poliza_" & rs!Num_Poliza & ".txt" For Output As #1
    
                  StrDocument = IIf(IsNull(rs!Json_Documentos), "", rs!Json_Documentos)
                  StrFirmantes = IIf(IsNull(rs!json_Firmantes), "", rs!json_Firmantes)
                  
                  Print #1, StrDocument
                  Print #1, "******************************"
                  Print #1, StrFirmantes
                  
                  Close #1
                      
                      
                  rs.MoveNext
                Wend
   
  conn.Close
  Set objCmd = Nothing
  Set rs = Nothing
  Set conn = Nothing
  
End Sub
  

