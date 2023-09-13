Attribute VB_Name = "Module1"
Global Const vgTopeEdadHombre = "65"
Global Const vgTopeEdadMujer = "60"

Global vgNombreProvincia As String
Global vgNombreRegion As String
Global Const vgCodTabla_AFP = "AF"         'Administradora de Fondos de Pensiones
Global Const vgCodTabla_AltPen = "AL"      'Alternativas de Pensión
Global Const vgCodTabla_Bco = "BCO"        'Entidades Bancarias
Global Const vgCodTabla_CCAF = "CC"        'Caja de Compensación
Global Const vgCodTabla_CauEnd = "CE"      'Causa de Generación de Endosos
Global Const vgCodTabla_CobPol = "CO"      'Cobertura de Pólizas (Trad)
Global Const vgCodTabla_ConPagTer = "COP"  'Concepto de Pago a Terceros
Global Const vgCodTabla_ComRea = "CR"      'Compañías de Reaseguro
Global Const vgCodTabla_CauSupAsiFam = "CSA" 'Causa Suspensión de la Asignación Familiar
Global Const vgCodTabla_CauSusGarEst = "CSG" 'Causa Suspensión de la Garantía Estatal
Global Const vgCodTabla_DerAcre = "DC"     'Derecho a Acrecer
Global Const vgCodTabla_DerPen = "DE"      'Derecho a Pensión
Global Const vgCodTabla_DerGarEst = "DEG"  'Derecho a Garantía Estatal
Global Const vgCodTabla_EstCiv = "EC"      'Estado Civil
Global Const vgCodTabla_EstPol = "EP"      'Estado de la Póliza
Global Const vgCodTabla_EstVigAsiFam = "EV" 'Estado de Vigencia de la Asignación Familiar
Global Const vgCodTabla_FacCas = "FC"      'Porcentaje Disminución Pensión por Quiebra Cía.
Global Const vgCodTabla_FrePago = "FP"     'Forma o Frecuencia de Pago
Global Const vgCodTabla_FacQui = "FQ"      'Porc. Pensión Garantizada por el Estado
Global Const vgCodTabla_GarEst = "GE"      'Estado de Garantía Estatal
Global Const vgCodTabla_GruFam = "GF"      'Grupo Familiar
Global Const vgCodTabla_InsSal = "IS"      'Institución de Salud
Global Const vgCodTabla_LimEdad = "LI"     'Límite de Edad para la Tabla de Mortalidad
Global Const vgCodTabla_ModRen = "MO"      'Modalidad de Rentabilidad
Global Const vgCodTabla_ModOriHabDes = "MOR" 'Modalidad de Origen del Haber o Descuento
Global Const vgCodTabla_MovHabDes = "MOV"  'Tipo de Movimiento del Haber o Descuento
Global Const vgCodTabla_ModPago = "MP"     'Modalidad de Pago
Global Const vgCodTabla_ModRea = "MR"      'Modalidad de Reaseguro
Global Const vgCodTabla_ModVejAnt = "MVA"  'Modalidad de Vejez Anticipada para Gar.Est.
Global Const vgCodTabla_OpeRea = "OR"      'Operación de Reaseguro
Global Const vgCodTabla_Par = "PA"         'Parentesco
Global Const vgCodTabla_ParNoBen = "PNB"   'Estado de Parentesco de los No Beneficiarios
Global Const vgCodTabla_Plan = "PL"        'Plan (Trad)
Global Const vgCodTabla_ReqPen = "RPE"     'Requisitos de Pensión
Global Const vgCodTabla_SitCor = "SC"      'Situación del Corredor
Global Const vgCodTabla_Sexo = "SE"        'Sexo
Global Const vgCodTabla_SitInv = "SI"      'Situación de Invalidez
Global Const vgCodTabla_TipCor = "TC"      'Tipo de Corredor
Global Const vgCodTabla_TipCta = "TCT"     'Tipo de Cuenta de Depósito
Global Const vgCodTabla_TipDoc = "TD"      'Tipo de Documento
Global Const vgCodTabla_TipPagTem = "TE"   'Tipo de Pago de Pensiones Temporales
Global Const vgCodTabla_TipEnd = "TEN"     'Tipo de Endoso
Global Const vgCodTabla_TipPri = "TI"      'Tipo de Prima
Global Const vgCodTabla_TipIngMen = "TIM"  'Tipo de Ingreso de Mensaje
Global Const vgCodTabla_TipMon = "TM"      'Tipo de Moneda
Global Const vgCodTabla_TipPen = "TP"      'Tipo de Pensión
Global Const vgCodTabla_TipPer = "TPE"     'Tipo de Persona
Global Const vgCodTabla_TipRen = "TR"      'Tipo de Rentabilidad
Global Const vgCodTabla_TipReso = "TRE"    'Tipo de Resolución
Global Const vgCodTabla_TipRetJud = "TRJ"  'Tipo de Retención Judicial
Global Const vgCodTabla_TipVej = "TV"      'Tipo de Vejez
Global Const vgCodTabla_TipVig = "VI"      'Estado de Vigencia
Global Const vgCodTabla_TipVigPol = "VP"   'Estado de Vigencia de la Póliza
Global Const vgCodTabla_ViaPago = "VPG"    'Vía de Pago de la Pensión

'------------------------------------------------------------
'Permite Cargar los Distintos Combos existentes en el sistema
'De acuerdo a los parámetros de llamada
'------------------------------------------------------------
Function fgComboGeneral(icodigo, icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_ComboGeneral

    icombo.Clear
    vgSql = "select cod_elemento,gls_elemento "
    vgSql = vgSql & "from MA_TPAR_TABCOD where "
    vgSql = vgSql & "cod_tabla = '" & icodigo & "'"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        If icodigo = "TP" Then
            If vlRsCombo!cod_elemento = "04" Or vlRsCombo!cod_elemento = "05" _
                Or vlRsCombo!cod_elemento = "06" Or vlRsCombo!cod_elemento = "07" Or _
                vlRsCombo!cod_elemento = "08" Then
                icombo.AddItem ((Trim(vlRsCombo!cod_elemento) & " - " & Trim(vlRsCombo!gls_elemento)))
                
            End If
        Else
            icombo.AddItem ((Trim(vlRsCombo!cod_elemento) & " - " & Trim(vlRsCombo!gls_elemento)))
        End If
        vlRsCombo.MoveNext
        
        
    Loop
    vlRsCombo.Close
    
    If icombo.ListCount <> 0 Then
        icombo.ListIndex = 0
    End If

Exit Function
Err_ComboGeneral:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'-----------------------------------------------------------------------
'PERMITE CARGAR TODAS LAS COMUNAS
'-----------------------------------------------------------------------
Function fgComboComuna(vlCombo As ComboBox)
On Error GoTo Err_Carga
     
     vlCombo.Clear
     vgSql = ""
     vlCont = 0
     vgSql = "Select Gls_Comuna,Cod_Direccion from MA_TPAR_COMUNA ORDER BY GLS_Comuna"
     Set vgCmb = vgConexionBD.Execute(vgSql)
         If Not (vgCmb.EOF) Then
            While Not (vgCmb.EOF)
                  vlCombo.AddItem (Trim(vgCmb!gls_comuna))
                  vlCont = vlCombo.ListCount - 1
                  vlCombo.ItemData(vlCont) = (vgCmb!cod_direccion)
                  vgCmb.MoveNext
            Wend
         End If
         vgCmb.Close
        
     If vlCombo.ListCount <> 0 Then
        vlCombo.ListIndex = 0
     End If

Exit Function
Err_Carga:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'-----------------------------------------------------------------------
'BUSCA EL NOMBRE DE LA PROVINCIA
'-----------------------------------------------------------------------
Function fgBuscarNombreProvinciaRegion(vlCodDir)
On Error GoTo Err_Buscar

     vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,c.Cod_Comuna,c.Gls_Comuna"
     vgSql = vgSql & " FROM MA_TPAR_COMUNA as c, MA_TPAR_PROVINCIA as p, MA_TPAR_REGION as r"
     vgSql = vgSql & " Where c.Cod_Direccion = '" & vlCodDir & "' and  "
     vgSql = vgSql & " c.cod_region = p.cod_region and"
     vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
     vgSql = vgSql & " p.cod_region = r.cod_region"
     Set vgRs = vgConexionBD.Execute(vgSql)
     If Not vgRs.EOF Then
        vgNombreRegion = (vgRs!gls_region)
        vgNombreProvincia = (vgRs!gls_provincia)
     End If

Exit Function
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

'------------------------------------------------------------
'Permite Cargar el Combo de Sucursales del Sistema
'------------------------------------------------------------
Function fgComboSucursal(icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_ComboSucursal
    
    icombo.Clear
    vgSql = "select cod_sucursal,gls_sucursal "
    vgSql = vgSql & "from MA_TPAR_SUCURSAL "
    vgSql = vgSql & "order by cod_sucursal "
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        icombo.AddItem ((Trim(vlRsCombo!cod_sucursal) & " - " & Trim(vlRsCombo!gls_sucursal)))
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    If icombo.ListCount <> 0 Then
        icombo.ListIndex = 0
    End If
    
Exit Function
Err_ComboSucursal:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'------------------------------------------------------------
'Permite Cargar el Combo de Causa de Invalidez del Sistema
'------------------------------------------------------------
Function fgComboCauInvalidez(icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_CombocauInvalidez
    
    icombo.Clear
    vgSql = "select cod_patologia,gls_patologia "
    vgSql = vgSql & "from MA_TPAR_PATOLOGIA "
    vgSql = vgSql & "order by cod_patologia "
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        icombo.AddItem ((Trim(vlRsCombo!cod_patologia) & " - " & Trim(vlRsCombo!gls_patologia)))
        vlRsCombo.MoveNext
    Loop
    vlRsCombo.Close
    If icombo.ListCount <> 0 Then
        icombo.ListIndex = 0
    End If
Exit Function
Err_CombocauInvalidez:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'----------------------------------------------
'FUNCION BUSCA POSICION EN UN COMBO
'----------------------------------------------
Function fgBuscaPos(icombo As ComboBox, icodigo)
        vgI = 0
        icombo.ListIndex = 0
        Do While vgI < icombo.ListCount
            If (Trim(icombo) <> "") Then
                If (Trim(icodigo) = Trim(Mid(icombo.Text, 1, (InStr(1, icombo, "-") - 1)))) Then
                    Exit Do
                End If
            End If
            vgI = vgI + 1
            If (vgI = icombo.ListCount) Then
                icombo.ListIndex = 0
                Exit Do
            End If
                icombo.ListIndex = vgI
        Loop
End Function
'----------------------------------------
'FUNCION VALIDA FECHAS
'----------------------------------------
Function fgValidaFecha(ifecha As String)
    If (Trim(ifecha) = "") Then
        MsgBox "Falta Ingresar Fecha", vbExclamation, "Error de Datos"
        fgValidaFecha = True
        Exit Function
    End If
    If Not IsDate(ifecha) Then
        MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
        fgValidaFecha = True
        Exit Function
    End If
    If (CDate(ifecha) > CDate(Date)) Then
        MsgBox "La Fecha es mayor a la fecha actual", vbCritical, "Error de Datos"
        fgValidaFecha = True
        Exit Function
    End If
    If (Year(CDate(ifecha)) < 1900) Then
        MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
        fgValidaFecha = True
        Exit Function
    End If
        
End Function
