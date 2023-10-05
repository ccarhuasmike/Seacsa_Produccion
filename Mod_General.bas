Attribute VB_Name = "Mod_General"


'Global Wrk As Workspace
'dim C As Double
'Global vgPath As String
'Global vgRuta As String

'Definición de Constantes
Global Const vgTipoSistema = "PD"
Global Const vgNombreSistema = "Sistema Previsional"

Global Const cgTipoNacionalidad As String = "PERUANA"
Global Const cgTipoIdenSinInformacion As String * 1 = "0"
Global Const cgTipoRentaInmediata As String * 1 = "1"

Global vgNombreHombre As String
Global vgNombreMujer  As String

'Nombres del Cliente o la Compañía adquisidora
Global vgNombreCortoCompania As String
Global vgNombreCompania      As String
Global vgCodigoSBSCompania   As String
Global vgCodigoSucaveCompañia As String
Global vgNombreSubSistema    As String
Global vgTipoIdenCompania    As String
Global vgNumIdenCompania     As String

'Variables utilizadas en la Conexión al Sistema
Global vgMensaje            As String
Global vgNombreServidor     As String
Global vgNombreBaseDatos    As String
Global vgNombreUsuario      As String
Global vgPassWord     As String
Global vgUsuario      As String
Global vgUsuarioSuc   As String
Global vgRutUsuario   As String
Global vgDsn          As String
Global vgRutaDataBase As String
Global vgRutaBasedeDatos As String
Global vgRutaBasedeDatos_Aux As String
Global vgRutaArchivo  As String

Global lpAppName As String
Global lpKeyName As String
Global lpDefault As String
Global lpReturnString As String
Global Size As Integer
Global lpFileName As String

Global vgConexionBD As ADODB.Connection
Global vgConectarBD As ADODB.Connection
Global Db_Laptop As ADODB.Connection
'Global vgConexionTrans As ADODB.Connection
'Global vgConexionParam As ADODB.Connection

Global vgRs          As ADODB.Recordset
Global vgRs1         As ADODB.Recordset
Global vgRs2         As ADODB.Recordset
Global vgRs3         As ADODB.Recordset
Global vgRs4         As ADODB.Recordset

'Datos correspondientes al Usuario que accesa el programa
Global vgLogin   As String
Global vgContraseña As String
Global vgNivel      As Integer

Global vgSql        As String
Global Sql          As String
Global vgQuery      As String
Global vgGra        As String
Global vgPalabra    As String
Global vgPalabraAux As String
Global vgRes        As Long

Global vgSw         As Boolean
Global vgI          As Long
Global vgX          As Integer
Global vgJ          As Long

Global vgNombreApoderado As String
Global vgCargoApoderado  As String

Global vgPolizaProceso As String

'Global vgPolizaProceso As String
Global Frase As String
Global vgModalPrueba As Boolean

Global vgDolar       As Double
Global vgValorMoneda As Double
Global vgValDol      As Double
'*******
Global vgRutGrilla As String
Global vgReiniciar As Boolean
Global vgNumCot As String
Global vgEdadActual As Integer
Global vgEdadJubila As Integer
Global vgClicEnGrilla As Boolean
Global vripc As Double
Global vgFormulario As String
Global vgNomForm As String
Global vgFormularioCarpeta As String
Global vlTipoPension As String
Global vgCalculo As String
Global vgNumeroCotizacion As String

Global vgTipoPension As String

Global Tb_Difpol     As ADODB.Recordset
Global Tb_Difben     As ADODB.Recordset
Global tb            As ADODB.Recordset
Global vgCmb         As ADODB.Recordset
Global CtaDifpol     As ADODB.Recordset
Global Tb_EvaparUS   As ADODB.Recordset
Global Tb_EvapasUS   As ADODB.Recordset
Global Tb_EvaproUS   As ADODB.Recordset

'Nuevas Variables
Global vgTipoTablaGeneral  As String
Global vgGlosaTablaGeneral As String
Global vgTituloGeneral     As String
Global vgAfp               As String
Global Arr_Tasas(112)      As Double

'Definición de Códigos de Tablas Virtuales
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
Global Const vgCodTabla_TipBono = "TB"      'Tipo de Bono
Global Const vgCodTabla_TipCor = "TC"      'Tipo de Corredor
Global Const vgCodTabla_TipCta = "TCT"     'Tipo de Cuenta de Depósito
Global Const vgCodTabla_ModTipCta = "TCC"     'Modalidad Tipo de Cuenta de Depósito
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
Global Const vgCodTabla_TipReajuste = "TTR" 'Tipo de Reajuste (2%) 'I--- ABV 05/02/2011 ---
Global Const vgCodTabla_TipVej = "TV"      'Tipo de Vejez
Global Const vgCodTabla_TipVig = "VI"      'Estado de Vigencia
Global Const vgCodTabla_TipVigPol = "VP"   'Estado de Vigencia de la Póliza
Global Const vgCodTabla_ViaPago = "VPG"    'Vía de Pago de la Pensión
Global Const vgCodTabla_PrcSal = "PS"      'Porcentaje de Salud Minima

Global vgNombreRegion     As String
Global vgNombreProvincia  As String
Global vgNombreComuna     As String
Global vgCodigoRegion   As String
Global vgCodigoProvincia As String
Global vgCodigoComuna As String

Public DatabaseName As String
Public ServerName   As String
Public ProviderName As String
Public UserName     As String
Public PasswordName As String

Global Const vgTipoBase = "ORACLE"
'Global Const vgTipoBase = "SQL"

Global Const cgCodTipMonedaUF As String * 2 = "NS"
Global Const cgNomTipMonedaUF As String * 2 = "S/."
Global vgMonedaCodOfi  As String
Global vgMonedaCodTran As String

'variables para visualizar las tasas
Global vgNivelIndicadorVer   As String 'visualiza las tasas desde la consulta
Global vgNivelIndicadorBoton As String 'activa el botón de recalculo de la póliza

Global Const cgIndicadorSi As String * 2 = "Si"
Global Const cgIndicadorNo As String * 2 = "No"
Global Const cgTipoSucursalSuc As String * 1 = "S"
Global Const cgTipoSucursalAfp As String * 1 = "A"
Global Const cgMonedaValorNS  As String * 1 = "1"
Global Const cgTipoDocSinInformacion As String * 1 = "0"

Global vgTipoSucursal As String
Global vgPrcBenSocial As Double

'Estructura para los Tipos de Moneda
Public Type TypeTablaMoneda
    Codigo      As String
    Descripcion As String
    Scomp       As String
End Type
Global egTablaMoneda()           As TypeTablaMoneda
Global vgNumeroTotalTablasMoneda As Long

Global Const clCodTipPensionInvTot As String * 2 = "06"
Global Const clCodTipPensionInvPar As String * 2 = "07"
Global Const clCodTipPensionSob As String * 2 = "08"

Public sName_Reporte As String
Public objRsRpt As ADODB.Recordset

'marco agrego de la fuente RV
Global vgFechaEfecto       As String

'hqr 12/01/2012
Global Const cgSINAJUSTE As String * 1 = "0"
Global Const cgAJUSTESOLES As String * 1 = "1"
Global Const cgAJUSTETASAFIJA As String * 1 = "2"
'fin hqr 03/12/2010

Global Const cgCodParentescoCau As String * 2 = "99"  'I--- ABV 05/02/2011 ---


'RRR18/01/2012

Public tammin As Integer
Public cantclvant As Integer
Public canmincaralf As Integer
Public freccambio As Integer
Public canantclv As Integer
Public FechaIni As String
Public FechaFin As String
Public fechaant As String
Public fecfinDa As Date
Public vlPassword As String
Global vgIntentos As Integer
Global vgChkdiaant As Integer
Global vgDiasFaltan As Integer
Global vgValorAr As Integer
Dim balfanum As Integer
Global strRpt As String
Global strID As String
'rrr

'Integracion GestorCliente_
Global stPolizaBenDirec() As TyBeneficiariosEst
Global stPolizaBenDirecMod() As TyBeneficiariosEst

Public Type TyBeneficiariosEst
    Num_Poliza As String
    Num_Endoso As Integer
    Num_Orden As Integer
    Fec_Ingreso As String
    Cod_TipoIdenBen As String
    Num_IdenBen As String
    Cod_Direccion As String
    cod_tip_fonoben As String
    cod_area_fonoben As String
    Gls_FonoBen As String
    cod_tipo_telben2 As String
    cod_area_telben2 As String
    gls_telben2 As String
    pTipoVia As String
    pDireccion As String
    pNumero As String
    pTipoPref As String
    pInterior As String
    pManzana As String
    pLote As String
    pEtapa As String
    pTipoConj As String
    pConjHabit As String
    pTipoBlock As String
    pNumBlock As String
    pReferencia As String
    pConcatDirec As String
    pGlsCorreo As String
    pDireccionConcat As String
    codComuna As String
    codRegion As String
    codProvincia As String
    pvalEndosoGS As String
    cod_pais As String
End Type
'Fin Integracion GestorCliente_
Public Type DireccionRep
    vTipoTelefono As String
    vNumTelefono As String
    vCodigoTelefono As String
    vTipoTelefono2 As String
    vNumTelefono2 As String
    vCodigoTelefono2 As String
    vTipoVia As String
    vDireccion As String
    vNumero As String
    vTipoPref As String
    vInterior As String
    vManzana As String
    vLote As String
    vEtapa As String
    vTipoConj As String
    vConjHabit As String
    vTipoBlock As String
    vNumBlock As String
    vReferencia As String
    vcodeDepar As Integer
    vcodeProv As Integer
    vCodeDistr As Integer
    vCodLoad As Integer
    vNomDepartamento As String
    vNomProvincia As String
    vNomDistrito As String
    vCodDireccion As String
    vgls_desdirebusq As String
End Type
'MVG
Global strBolElec As String
Global DirRep As DireccionRep
'************************ F U N C I O N E S RRR ******************
'JEVC CORPTEC 24/07/2017
Global num_session As Double
'JEVC CORPTEC 24/07/2017
Function fgLogIn_Pro() As Boolean
    Dim sistema, modulo As String
    sistema = "SEACSA"
    modulo = "PRODUCCION"
    vlSql = "INSERT INTO LOG_SESSION (COD_USUARIO, FEC_INI, HOR_INI, COD_ESTADO, "
    vlSql = vlSql & "FEC_FIN , HOR_FIN,  GLS_MODULO, GLS_SISTEMA) VALUES ("
    vlSql = vlSql & "'" & vgLogin & "', TO_CHAR(SYSDATE, 'YYYYMMDD'), TO_CHAR(SYSDATE, 'HH24:MI'), 'A', "
    vlSql = vlSql & "NULL, NULL, '" & modulo & "', '" & sistema & "')"
    vgConexionBD.Execute (vlSql)
    
    vgSql = "SELECT LOG_SESSION_SEQ.CURRVAL AS LSID FROM DUAL"
    Set vgRs = vgConexionBD.Execute(vgSql)
    If Not vgRs.EOF Then
        num_session = vgRs!LSID
    End If
End Function
'JEVC CORPTEC 24/07/2017
Function fgLogOut_Pro() As Boolean
    vlSql = "UPDATE LOG_SESSION SET FEC_FIN = TO_CHAR(SYSDATE, 'YYYYMMDD'), "
    vlSql = vlSql & "HOR_FIN = TO_CHAR(SYSDATE, 'HH24:MI'), COD_ESTADO = 'I' "
    vlSql = vlSql & "WHERE NUM_SESSION = " & num_session
    vgConexionBD.Execute (vlSql)
End Function
'RRR 18/01/2012
Public Function fIaplicavalidacion(usuario As String, txt_password As TextBox, txt_passwordcomfir As TextBox) As Integer

    vgSql = "SELECT * FROM MA_TMAE_ADMINCUENTAS WHERE "
    vgSql = vgSql & "cod_cliente = '1' "
    Set vgRs = vgConexionBD.Execute(vgSql)

    If Not vgRs.EOF Then
        tammin = vgRs!ntamañomin
        cantclvant = vgRs!ncanclvant
        canmincaralf = vgRs!ncaracmin
        freccambio = vgRs!nfrecuencia
        canantclv = vgRs!ncanclvant
        balfanum = vgRs!balfanum
    End If
    
    If (txt_password <> txt_passwordcomfir) Then
            MsgBox "Las Contraseñas registradas son distintas, vuelva a registrarlas.", vbExclamation, "Error de Contraseña"
            txt_password = ""
            txt_passwordcomfir = ""
            txt_password.SetFocus
            fIaplicavalidacion = 0
        Exit Function
    End If
    
    If Len(txt_password) < tammin Then
        MsgBox "Password debe ser minimo de " & CStr(tammin) & " caracteres ", vbCritical, "Error de Datos"
        txt_password.SetFocus
        fIaplicavalidacion = 0
        Exit Function
    End If
   
    If Len(txt_passwordcomfir) < tammin Then
        MsgBox "Password debe ser minimo de " & CStr(tammin) & " caracteres ", vbCritical, "Error de Datos"
        txt_passwordcomfir.SetFocus
        fIaplicavalidacion = 0
        Exit Function
    End If
    
    vgRs.Close
    
       
    vgSql = " select nro_usupass, gls_password from MA_TMAE_USUPASSWORD "
    vgSql = vgSql & " where cod_usuario='" & usuario & "' "
    vgSql = vgSql & " and nro_usupass > (select count(*) from MA_TMAE_USUPASSWORD where cod_usuario='" & usuario & "') - " & cantclvant
    vgSql = vgSql & " and nro_usupass <= (select count(*) from MA_TMAE_USUPASSWORD where cod_usuario='" & usuario & "')"
    vgSql = vgSql & " order by 1 desc"


    Set vgRs = vgConexionBD.Execute(vgSql)

    Dim strclave As String
    
    If Not vgRs.EOF Then
        Do While Not vgRs.EOF
            strclave = fgDesPassword(vgRs!gls_password)
            
            If UCase(Trim(txt_password)) = strclave Then
                MsgBox "No puede utilizar un password anterior, Por favor elegir otro.", vbCritical, "Error de Datos"
                txt_password.SetFocus
                fIaplicavalidacion = 0
                Exit Function
            End If
            vgRs.MoveNext
        Loop
    End If
    
    Dim i, l, n, a As Integer
    Dim car As String

    
    For i = 1 To Len(txt_password)
    
        car = Mid(txt_password, i, 1)
    
        If VLetras(Asc(car)) <> 0 Then l = l + 1
        If Numeros(Asc(car)) <> 0 Then n = n + 1
        If VAlfanumerico(Asc(car)) <> 0 Then a = a + 1
        
    Next
    
    If balfanum = 1 Then
        If a < canmincaralf Then
            MsgBox "La clave debe contener como minimo " & canmincaralf & " caracteres alfanumericos.", vbCritical, "Error de Datos"
            txt_password.SetFocus
            aplicavalidacion = 0
            Exit Function
        End If
    Else
        If a > 0 Then
            MsgBox "La clave no debe contener caracteres alfanumericos.", vbCritical, "Error de Datos"
            aplicavalidacion = 0
            Exit Function
        End If
    End If
    
    FechaIni = Mid(CStr(Now), 7, 4) & Mid(CStr(Now), 4, 2) & Mid(CStr(Now), 1, 2)
    fecfinDa = DateAdd("d", freccambio, Now)
    FechaFin = Mid(CStr(DateAdd("d", freccambio, Now)), 7, 4) & Mid(CStr(DateAdd("d", freccambio, Now)), 4, 2) & Mid(CStr(DateAdd("d", freccambio, Now)), 1, 2)
    fechaant = Mid(CStr(DateAdd("d", -CInt(canantclv), fecfinDa)), 7, 4) & Mid(CStr(DateAdd("d", -CInt(canantclv), fecfinDa)), 4, 2) & Mid(CStr(DateAdd("d", -CInt(canantclv), fecfinDa)), 1, 2)
    
    fIaplicavalidacion = 1
End Function

Public Function VLetras(Tecla As Integer) As Integer
Dim strValido As String
'letras no validas: .*-}¿'!%&/()=?¡]¨*[Ññ;:_ áéíó
strValido = "qwertyuioplkjhgfdsazxcvbnmQWERTYUIOPASDFGHJKLZXCV BNM, "
If Tecla > 26 Then
If InStr(strValido, Chr(Tecla)) = 0 Then
Tecla = 0
End If
End If
VLetras = Tecla
End Function
Public Function Numeros(Tecla As Integer) As Integer
Dim strValido As String
strValido = "0123456789"
If Tecla > 26 Then
'compara los numeros ke hay en la variable strValido _
con el numero ingresado(Tecla)
'si el numero ingresado(Tecla) no esta en la variable strValido entonces _
Tecla = 0, la funcion Chr convierte el numero a ascii
If InStr(strValido, Chr(Tecla)) = 0 Then
Tecla = 0
End If
End If
Numeros = Tecla
End Function

Public Function VAlfanumerico(Tecla As Integer) As Integer
Dim strValido As String
'letras no validas: .*-}¿'!%&/()=?¡]¨*[Ññ;:_ áéíó
strValido = "!#$%&/()=?¡'¿{}^`[]*\-+.,;:_ "
If Tecla > 26 Then
If InStr(strValido, Chr(Tecla)) = 0 Then
Tecla = 0
End If
End If
VAlfanumerico = Tecla
End Function
'RRR 18/01/2012


'----------------------------------------------
'FUNCION BUSCA POSICION EN UN COMBO
'----------------------------------------------
Function fgBuscaPos(icombo As ComboBox, iCodigo)
        vgI = 0
        icombo.ListIndex = 0
        Do While vgI < icombo.ListCount
            If (Trim(icombo) <> "") Then
                If (Trim(iCodigo) = Trim(Mid(icombo.Text, 1, (InStr(1, icombo, "-") - 1)))) Then
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
Function fgValidaFecha(iFecha As String)
    If (Trim(iFecha) = "") Then
        MsgBox "Falta Ingresar Fecha", vbExclamation, "Error de Datos"
        fgValidaFecha = True
        Exit Function
    End If
    If Not IsDate(iFecha) Then
        MsgBox "La Fecha ingresada no es una fecha válida.", vbCritical, "Error de Datos"
        fgValidaFecha = True
        Exit Function
    End If
    If (CDate(iFecha) > CDate(Date)) Then
        MsgBox "La Fecha es mayor a la fecha actual", vbCritical, "Error de Datos"
        fgValidaFecha = True
        Exit Function
    End If
    If (Year(CDate(iFecha)) < 1900) Then
        MsgBox "Error en la Fecha ingresada es menor a la mínima fecha que se puede ingresar (1900).", vbCritical, "Error de Datos"
        fgValidaFecha = True
        Exit Function
    End If
End Function

'------------------------------------------------------------
'Permite Cargar los Distintos Combos existentes en el sistema
'De acuerdo a los parámetros de llamada
'------------------------------------------------------------
Function fgComboGeneral(iCodigo, icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_ComboGeneral

    icombo.Clear
    vgSql = "select cod_elemento,gls_elemento "
    vgSql = vgSql & "from MA_TPAR_TABCOD where "
    vgSql = vgSql & "cod_tabla = '" & iCodigo & "'"
    If (iCodigo = vgCodTabla_TipVej) Then
        vgSql = vgSql & " ORDER BY cod_elemento DESC "
    Else
        vgSql = vgSql & " ORDER BY cod_elemento ASC "
    End If
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        If iCodigo = "TP" Then
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

'Carga los combos para las direcciones
Function fgComboGeneralDirec(iCodigo, icombo As ComboBox)
Dim vlRsCombo As ADODB.Recordset
Dim vlDescripcion As String * 50

On Error GoTo Err_ComboGeneral

    icombo.Clear
    icombo.AddItem ""
    vgSql = "select cod_elemento,gls_elemento,cod_scomp from MA_TPAR_TABCOD where "
    vgSql = vgSql & "cod_tabla = '" & iCodigo & "' ORDER BY cod_elemento"
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        vlDescripcion = (vlRsCombo!cod_elemento) & " - " & (vlRsCombo!gls_elemento)
        icombo.AddItem vlDescripcion & Trim(vlRsCombo!cod_scomp)
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

'PERMITE CARGAR TODAS LAS COMUNAS
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
                  vlCombo.ItemData(vlCont) = (vgCmb!Cod_Direccion)
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

Function fgBuscarNombreComunaProvinciaRegion(vlCodDir As String)
Dim vlRegistroDir As ADODB.Recordset
On Error GoTo Err_Buscar

     vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,c.Cod_Comuna,c.Gls_Comuna"
     vgSql = vgSql & " FROM MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r"
     vgSql = vgSql & " Where c.Cod_Direccion = '" & vlCodDir & "' and  "
     vgSql = vgSql & " c.cod_region = p.cod_region and"
     vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
     vgSql = vgSql & " p.cod_region = r.cod_region"
     Set vlRegistroDir = vgConexionBD.Execute(vgSql)
     If Not vlRegistroDir.EOF Then
        vgNombreRegion = (vlRegistroDir!gls_region)
        vgNombreProvincia = (vlRegistroDir!gls_provincia)
        vgNombreComuna = (vlRegistroDir!gls_comuna)
        vgCodigoRegion = (vlRegistroDir!cod_region)
        vgCodigoProvincia = (vlRegistroDir!COD_PROVINCIA)
        vgCodigoComuna = (vlRegistroDir!cod_comuna)
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

'Function fgBonoReconoc(vld_FNac, vld_FEmision, vld_FBalance, vld_ValorUF, vld_VNBon, vld_TDBon, Sexo, vli_EdadCobroAnt) As Double
''LAN,LMN,LDN : FECHA DE NACIMIENTO DEL AFILIADO
''LAB,LMB,LDB : FECHA DE BALANCE
''LAE,LME,LDE : FECHA DE EMISION DEL DOCUMENTO
''vld_VNBon       : VALOR NOMINAL A LA FECHA DE EMISION
''vld_TDBon       : TASA DE INTERES DE DESCUENTO DEL BONO
''SEXO        : SEXO DEL AFILIADO
''vli_EdadCobroAnt        : EDAD COBRO DE BONO ANTICIPADAMENTE
''vld_ValorUF        : Valor UF
'Dim vripc, vruf, agnos As Double
'Dim lanoa As Double
'Dim im1, im2, in1, in2, lan, lmn, ldn, lae, lme As Integer
'Dim lde, lab, lmb, ldb, lacobro, ldc, lmc  As Integer
'Dim lmesc As Integer
'Dim lano As Integer, lmes As Integer, lacc As Integer
'Dim vl_ipcbalance As Double, vl_ipcemision As Double
'
'',LAN,LMN,LDN,LAE,LME,LDE,LAB,LMB,LDB,vld_ValorUF,vld_VNBon,vld_TDBon,SEXO,vli_EdadCobroAnt,CODSUC
''vld_FNac,vld_FEmision,vld_FBalance,vld_ValorUF,vld_VNBon,vld_TDBon,Sexo,vli_EdadCobroAnt
'    'CALCULO DE LA FECHA DE COBRO DEL BONO
'
'    vl_EdadJubLegalMasc = 65
'    vl_EdadJubLegalFem = 60
'    vl_tasa = 0.04
'
'    vl_ipcbalance = 0
'    vl_ipcemision = 0
'    vld_TDBon = vld_TDBon / 100
'    lac = lan + vl_EdadJubLegalMasc
'    If Sexo = "F" Then lac = lan + vl_EdadJubLegalFem
'    If (vli_EdadCobroAnt <> 0) Then lac = lan + vli_EdadCobroAnt
'    lacobro = lac
'    lmc = lmn
'    ldc = ldn
'    'CALCULO DEL VALOR DEL IPC DE ACTUALIZACION
'    im1 = lae '- 1978
'    im2 = lme - 1
'    If im2 = 0 Then
'        im2 = 12
'        im1 = im1 - 1
'    End If
'    vgSql = "Select * from MA_TVAL_IPC where num_anno = " & im1 & " "
'    Set vlRegistro = vgConexionBD.Execute(vgSql)
'    If Not (vlRegistro.EOF) Then
'        Select Case (im2)
'            Case 1:     vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes1), 0, vlRegistro!prc_mes1)), "#0.000000")
'            Case 2:     vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes2), 0, vlRegistro!prc_mes2)), "#0.000000")
'            Case 3:     vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes3), 0, vlRegistro!prc_mes3)), "#0.000000")
'            Case 4:     vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes4), 0, vlRegistro!prc_mes4)), "#0.000000")
'            Case 5:     vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes5), 0, vlRegistro!prc_mes5)), "#0.000000")
'            Case 6:     vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes6), 0, vlRegistro!prc_mes6)), "#0.000000")
'            Case 7:     vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes7), 0, vlRegistro!prc_mes7)), "#0.000000")
'            Case 8:     vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes8), 0, vlRegistro!prc_mes8)), "#0.000000")
'            Case 9:     vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes9), 0, vlRegistro!prc_mes9)), "#0.000000")
'            Case 10:    vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes10), 0, vlRegistro!prc_mes10)), "#0.000000")
'            Case 11:    vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes11), 0, vlRegistro!prc_mes11)), "#0.000000")
'            Case 12:    vl_ipcemision = Format((IIf(IsNull(vlRegistro!prc_mes12), 0, vlRegistro!prc_mes12)), "#0.000000")
'            Case Else
'                'error
'                fgBonoReconoc = 0
'                Exit Function
'        End Select
'    Else
'        'Error
'        fgBonoReconoc = 0
'        Exit Function
'    End If
'
'    in1 = lab '- 1978
'    in2 = lmb - 1
'    If in2 <= 0 Then
'        If in2 = 0 Then in2 = 12
'        in1 = in1 - 1
'    End If
'    vgSql = "Select * from MA_TVAL_IPC where num_anno = " & iAño & " "
'    Set vlRegistro = vgConexionBD.Execute(vgSql)
'    If Not (vlRegistro.EOF) Then
'        Select Case (im2)
'            Case 1:     vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes1), 0, vlRegistro!prc_mes1)), "#0.000000")
'            Case 2:     vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes2), 0, vlRegistro!prc_mes2)), "#0.000000")
'            Case 3:     vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes3), 0, vlRegistro!prc_mes3)), "#0.000000")
'            Case 4:     vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes4), 0, vlRegistro!prc_mes4)), "#0.000000")
'            Case 5:     vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes5), 0, vlRegistro!prc_mes5)), "#0.000000")
'            Case 6:     vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes6), 0, vlRegistro!prc_mes6)), "#0.000000")
'            Case 7:     vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes7), 0, vlRegistro!prc_mes7)), "#0.000000")
'            Case 8:     vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes8), 0, vlRegistro!prc_mes8)), "#0.000000")
'            Case 9:     vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes9), 0, vlRegistro!prc_mes9)), "#0.000000")
'            Case 10:    vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes10), 0, vlRegistro!prc_mes10)), "#0.000000")
'            Case 11:    vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes11), 0, vlRegistro!prc_mes11)), "#0.000000")
'            Case 12:    vl_ipcbalance = Format((IIf(IsNull(vlRegistro!prc_mes12), 0, vlRegistro!prc_mes12)), "#0.000000")
'            Case Else
'                'error
'                fgBonoReconoc = 0
'                Exit Function
'        End Select
'    Else
'        'Error
'        fgBonoReconoc = 0
'        Exit Function
'    End If
'
'    'If in1 < 1 Or im1 < 1 Then GoTo 804
'    'If (in2 < 1 Or in2 > 12 Or im2 < 1 Or im2 > 12) Then GoTo 804
'    'vripc = tabipca(in1, in2) / tabipca(im1, im2)
'
'    vripc = vl_ipcbalance / vl_ipcemision
'
'    'CALCULO DEL PERIODO ENTRE FECHAS DE COBRO Y BALANCE
'    lano = lac - lab
'    lmes = lmc - lmb
'    If (lmes < 0) Then
'        lmes = lmes + 12
'        lano = lano - 1
'    End If
'    'If (lano < 0) Then lano = lano + 100
'    agnos = lmes / 12 + lano
'
'    'CALCULO DEL PERIODO ENTRE EMISION Y COBRO
'    lacc = lac - lae
'    lmesc = lmc - lme
'    If (lmesc < 0) Then
'        lmesc = lmesc + 12
'        lacc = lacc - 1
'    End If
'    'CALCULO DEL PERIODO ENTRE FECHAS DE BALANCE Y EMISION
'    lanoa = lab - lae
'    lmesa = lmb - lme
'    If (lmesa < 0) Then
'        lmesa = lmesa + 12
'        lanoa = lanoa - 1
'    End If
'    If lanoa < 0 Then lanoa = lanoa + 100
'    AGNOA = 1# * lmesa / 12# + lanoa
'    'VALOR BONO FECHA BALANCE
'
'    FACTBB = (1.04 ^ lanoa) * (1 + (0.04 * lmesa / 12#))
'    VALBB = vld_VNBon * vripc * FACTBB
'
'    'ACTUALIZACION BONO NOMINAL  Y
'    'DETERMINACION DEL VALOR CASTIGADO
'
'    FACTB = ((1.04 ^ lacc) * (1 + (0.04 * lmesc / 12#)))
'    VALFB = vld_VNBon * vripc * FACTB
'
'    FDES = (1# + vld_TDBon) ^ agnos
'    VPB = VALFB / FDES
'    FACDE = (1.04 ^ agnos)
'    FACDAC = FACDE / FDES
'
'    vruf = VPB / vld_ValorUF
'
'    'Impresion
'    'NRCOT,CODSUC,VRUF,LDC,LMC,LACOBRO
'    Exit Function
'
'End Function

Function fgGetPrivateIni(section, key$, FnameIni)
Dim retVal As String
Dim AppName As String
Dim worked As Integer

retVal = String$(255, 0)
worked = GetPrivateProfileString(section, key, "", retVal, Len(retVal), FnameIni)
If (worked = 0) Then
    fgGetPrivateIni = "DESCONOCIDO"
Else
    fgGetPrivateIni = Left(retVal, InStr(retVal, Chr(0)) - 1)
End If
End Function

Function fgConexionBaseDatos_aux(oConBD As ADODB.Connection)
       
    'Por defecto supone que falla la Conexión
    fgConexionBaseDatos_aux = False

On Local Error GoTo Err_ConsultaBD

    Set oConBD = New ADODB.Connection
    'oConBD.ConnectionString = "driver={Sql Server}; server=" & vgNombreServidor & ";UID=" & vgNombreUsuario & ";PWD=" & vgPassWord & ";database=" & vgNombreBaseDatos
    oConBD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & vgRutaBasedeDatos_Aux & ";Persist Security Info=False"
    oConBD.ConnectionTimeout = 1800
    oConBD.CommandTimeout = 1800
    oConBD.Open
    'oConBD.BeginTrans
    'La Conexión fue realizada
    fgConexionBaseDatos_aux = True
   ' oConBD.CommitTrans
    Exit Function

Err_ConsultaBD:
    'MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : " & vbCrLf & Err.Description, vbCritical, "Error de Consulta a Base de Datos"
    'Err.Clear
    'Text3.Text = "Errores en la Impresión de Solicitudes"
    MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : [ " & Err & Space(4) & Err.Description & " ]"
End Function

'Function fgConexionBaseDatos(oConBD As ADODB.Connection)
'
'    'Por defecto supone que falla la Conexión
'    fgConexionBaseDatos = False
'
'On Local Error GoTo Err_ConsultaBD
'
'    Set oConBD = New ADODB.Connection
'    oConBD.ConnectionString = "driver={Sql Server}; server=" & vgNombreServidor & ";UID=" & vgNombreUsuario & ";PWD=" & vgPassWord & ";database=" & vgNombreBaseDatos
'    'oConBD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & vgRutaBasedeDatos & ";Persist Security Info=False"
'    'oConBD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vgRutaBasedeDatos & ";Persist Security Info=False"
'    oConBD.ConnectionTimeout = 1800
'    oConBD.CommandTimeout = 1800
'    oConBD.Open
'    'oConBD.BeginTrans
'    'La Conexión fue realizada
'    fgConexionBaseDatos = True
'   ' oConBD.CommitTrans
'    Exit Function
'
'Err_ConsultaBD:
'    'MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : " & vbCrLf & Err.Description, vbCritical, "Error de Consulta a Base de Datos"
'    'Err.Clear
'    'Text3.Text = "Errores en la Impresión de Solicitudes"
'    MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : [ " & Err & Space(4) & Err.Description & " ]"
'End Function

Function fgConexionBaseDatos(oConBD As ADODB.Connection)
    Dim StringConexion As String
    'Por defecto supone que falla la Conexión
    fgConexionBaseDatos = False

On Local Error GoTo Err_ConsultaBD

    Set oConBD = New ADODB.Connection
    'If Not (True) Then 'SQL
    If vgTipoBase = "ORACLE" Then 'Oracle
        StringConexion = "Provider=" & ProviderName & ";"
        StringConexion = StringConexion & "Server= " & vgNombreServidor & " ;"
        StringConexion = StringConexion & "User ID= " & vgNombreUsuario & " ;"
        StringConexion = StringConexion & "Password= " & LCase(vgPassWord) & ";"
        StringConexion = StringConexion & "Data Source=" & vgNombreBaseDatos & " "
    Else
        StringConexion = "driver={Sql Server}; server=" & vgNombreServidor & ";UID=" & vgNombreUsuario & ";PWD=" & vgPassWord & ";database=" & vgNombreBaseDatos
    End If

    oConBD.ConnectionString = StringConexion
   
    'oConBD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & vgRutaBasedeDatos & ";Persist Security Info=False"
    'oConBD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vgRutaBasedeDatos & ";Persist Security Info=False"
    oConBD.ConnectionTimeout = 1800
    oConBD.CommandTimeout = 1800
    oConBD.Open
    'oConBD.BeginTrans
    'La Conexión fue realizada
    fgConexionBaseDatos = True
   ' oConBD.CommitTrans
    Exit Function

Err_ConsultaBD:
    'MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : " & vbCrLf & Err.Description, vbCritical, "Error de Consulta a Base de Datos"
    'Err.Clear
    'Text3.Text = "Errores en la Impresión de Solicitudes"
    MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : [ " & Err & Space(4) & Err.Description & " ]"
End Function

'Function amin1(arg3, arg4) As Double
'Dim xx As Double
'    If arg3 <= arg4 Then
'       xx = arg3
'    Else
'       xx = arg4
'    End If
'    amin1 = xx
'End Function

'Function amin0(arg1, arg2) As Integer
'Dim xx As Integer
'    If arg1 <= arg2 Then
'       xx = arg1
'    Else
'       xx = arg2
'    End If
'    amin0 = xx
'End Function

'Function amax1(arg3, arg4) As Double
'Dim xx As Double
'    If arg3 >= arg4 Then
'       xx = arg3
'    Else
'       xx = arg4
'    End If
'    amax1 = xx
'End Function

'Function amax0(arg1, arg2) As Integer
'Dim xx As Integer
'    If arg1 >= arg2 Then
'       xx = arg1
'    Else
'       xx = arg2
'    End If
'    amax0 = xx
'End Function

'Function Cambia_Dec(campo As Variant) As String
'  Dim ok As Boolean
'  Dim POS As Integer
'
'  ok = campo Like "*,*"
'
'  If ok Then
'    POS = InStr(campo, ",")
'    Mid(campo, POS) = "."
'  End If
'  Cambia_Dec = campo
'
'End Function

'Function fgValorPorcentaje(cont_param, h, tipo_per, iSitInv As String, iSexo As String, iFecha) As Double
'
'Dim Tb_Por As ADODB.Recordset
'
'    Sql = "select prc_pension as valor_porcentaje "
'    Sql = Sql & "from MA_TVAL_PORPAR where "
'    Sql = Sql & "Cod_par = '" & tipo_per & "' and "
'    Sql = Sql & "Cod_sitinv = '" & iSitInv & "' and "
'    Sql = Sql & "Cod_sexo = '" & iSexo & "' and "
'    Sql = Sql & "fec_inivigpor <= '" & iFecha & "' and "
'    Sql = Sql & "fec_tervigpor >= '" & iFecha & "'"
'
'    Set Tb_Por = vgConexionBD.Execute(Sql)
'    If Not Tb_Por.EOF Then
'        'Valor_Porcentaje = tb_por!Valor_Porcentaje / cont_param
'        fgValorPorcentaje = Tb_Por!valor_porcentaje
'    Else
'        fgValorPorcentaje = -1
'        MsgBox "No existen datos en la tabla de Porcentajes de Parentesco", vbCritical, "Falta Información"
'    End If
'    Tb_Por.Close
    
'Function fgValorPorcentaje(cont_param, h, tipo_per) As Double
'Dim Tb_Por As ADODB.Recordset
'
'    Sql = "select mto_elemento as valor_porcentaje "
'    Sql = Sql & "from MA_TPAR_TABCOD where "
'    Sql = Sql & "Cod_Tabla = 'PA' and "
'    Sql = Sql & "Cod_Elemento = '" & tipo_per & "' "
'    'Sql = Sql & "and fec_inipar <= '" & Format(vg_fecsin, "yyyy/mm/dd") & "' And "
'    'Sql = Sql & "fec_terpar >= '" & Format(vg_fecsin, "yyyy/mm/dd") & "'"
'    Set Tb_Por = vgConexionBD.Execute(Sql)
'    If Not Tb_Por.EOF Then
'        'Valor_Porcentaje = tb_por!Valor_Porcentaje / cont_param
'        fgValorPorcentaje = Tb_Por!Valor_Porcentaje
'    Else
'        fgValorPorcentaje = -1
'        MsgBox "No existen datos en la tabla de Porcentajes de Parentesco", vbCritical
'    End If
'    Tb_Por.Close
    
'End Function

''Permite obtener el valor del Gasto de Sepelio de desde Mant. de Parámetros
'Function fgBuscaSepelio(iDia, iMes, iAño) As Double
''On Error GoTo Err_BusSep
'
'    fgBuscaSepelio = 0
'
'    vlSql = "SELECT mto_cuomor FROM MA_TVAL_CUOMOR WHERE "
'    vlSql = vlSql & "cod_moneda = '" & vgTipoMoneda & "' AND "
'    vlSql = vlSql & "fec_inicuomor <= '" & Format(DateSerial(iAño, iMes, iDia), "yyyymmdd") & "' and "
'    vlSql = vlSql & "fec_tercuomor >= '" & Format(DateSerial(iAño, iMes, iDia), "yyyymmdd") & "' "
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    If Not vgRs.EOF Then
'        If Not IsNull(vgRs!MTO_CUOMOR) Then
'            fgBuscaSepelio = vgRs!MTO_CUOMOR
'        End If
'    End If
'    vgRs.Close
'
'Exit Function
'Err_BusSep:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Function

'Function Calcula_Diferida(Cotizacion, Cod_afp, Comision, indicador, iRentabilidad)
'Dim Add_porc_be As Double, Vpptem As Double, Tasa_afp As Double, Prima_unica As Double
'Dim Rete_sim As Double, Prun_sim As Double, Sald_sim As Double, Mesga2 As Double
'Dim vlNumCoti As String
'Dim vlCorrCot As Integer
'
'Dim Dif As ADODB.Recordset
'
'    Add_porc_ben = 0
'    Vpptem = 0
'    Tasa_afp = 0
'    Prima_unica = 0
'    Rete_sim = 0
'    Prun_sim = 0
'    Sald_sim = 0
'    Mesga2 = 0
'    gto_sepelio = 0
'    Query = "SELECT * FROM PT_TMAE_PROPUESTA WHERE "
'    Query = Query & " num_cot = '" & Cotizacion & "' and "
'    Query = Query & " num_pro = " & indicador & ""
'    Set Dif = vgConectarBD.Execute(Query)
'    If Not Dif.EOF Then
'        If Dif!num_mesdif > 0 Then
'            Prima_unica = Dif!mto_priuni
'            gto_sepelio = Dif!mto_pengar
'            Tasa_afp = iRentabilidad / 100
''            'Determina la Suma Total de Porcentajes del Grupo Familiar
''            Paso1 = "SELECT sum(prc_legal) as porcen "
''            Paso1 = Paso1 & "FROM PT_TMAE_BENPRO WHERE "
''            Paso1 = Paso1 & "num_cot = '" & Cotizacion & "'"
''            Set Q = vgConectarBD.Execute(Paso1)
''            If Not Q.EOF Then
''                If Not IsNull(Q!porcen) Then
''                    Add_porc_be = Q!porcen
''                Else
''                    Add_porc_be = 0
''                End If
''            End If
''            Q.Close
'            Vpptem = ((1 - 1 / ((1 + Tasa_afp) ^ Dif!num_mesdif)) / Tasa_afp) * (1 + Tasa_afp) * 12
'            Vpptem = CDbl(Format(CDbl(Vpptem), "##0.000000"))
'            vlCorrCot = Dif!num_correlcot
'            vlNumCoti = Mid(Cotizacion, 1, 10) & Format(vlCorrCot, "00")
'            vlSql = "Update PT_TMAE_PROPUESTA set "
'            vlSql = vlSql & "mto_valprepentmp = " & Str(Vpptem) & " "
'            vlSql = vlSql & "WHERE "
'            vlSql = vlSql & "num_cot = '" & Cotizacion & "' and "
'            vlSql = vlSql & "num_pro = " & indicador & " "
'            vgConectarBD.Execute (vlSql)
'            If (vlCorrCot <> 0) Then
'                vlSql = "Update PT_TMAE_COTIZACION set "
'                vlSql = vlSql & "mto_valprepentmp = " & Str(Vpptem) & " "
'                vlSql = vlSql & "WHERE "
'                vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                vgConectarBD.Execute (vlSql)
'            End If
''            Add_porc_be = Add_porc_be / 100
'            If Dif!mto_priunisim > 0 Then
'                Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
'            End If
'            Sald_sim = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
'            If (Dif!mto_priunisim > 0) Then
'                Prun_sim = CDbl(Format(CDbl(Prima_unica - Sald_sim), "#,#0.00"))
'                'Renta Vitalicia Compañia
'                vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / Dif!mto_priunisim), "#,#0.00"))
'            End If
'            vlSql = "Update PT_TMAE_PROPUESTA set "
'            vlSql = vlSql & "mto_pensim = 0 "
'            vlSql = vlSql & "WHERE "
'            vlSql = vlSql & "num_cot = '" & Cotizacion & "' and "
'            vlSql = vlSql & "num_pro = " & indicador & " "
'            vgConectarBD.Execute (vlSql)
'
'            If (vlCorrCot <> 0) Then
'                vlSql = "Update PT_TMAE_COTIZACION set "
'                vlSql = vlSql & "mto_pensim = 0 "
'                vlSql = vlSql & "WHERE "
'                vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                vgConectarBD.Execute (vlSql)
'            End If
'            If (Dif!mto_priunisim > 0) Then
'                vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / Dif!mto_priunisim), "#,#0.00"))
'                vlSql = "Update PT_TMAE_PROPUESTA set "
'                vlSql = vlSql & "mto_pensim = " & Str(vlPensim) & " "
'                vlSql = vlSql & "WHERE "
'                vlSql = vlSql & "num_cot = '" & Cotizacion & "' and "
'                vlSql = vlSql & "num_pro = " & indicador & " "
'                vgConectarBD.Execute (vlSql)
'
'                If (vlCorrCot <> 0) Then
'                    vlSql = "Update PT_TMAE_COTIZACION set "
'                    vlSql = vlSql & "mto_pensim = " & Str(vlPensim) & " "
'                    vlSql = vlSql & "WHERE "
'                    vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                    vgConectarBD.Execute (vlSql)
'                End If
'            End If
'            vlSql = "Update PT_TMAE_PROPUESTA set "
'            vlSql = vlSql & "mto_priunidif = " & Str(Prun_sim) & ", "
'            vlSql = vlSql & "mto_ctaindafp = " & Str(Sald_sim) & ", "
'            vlSql = vlSql & "mto_rentatmpafp = " & Str(Rete_sim) & " "
'            vlSql = vlSql & "WHERE "
'            vlSql = vlSql & "num_cot = '" & Cotizacion & "' and "
'            vlSql = vlSql & "num_pro = " & indicador & " "
'            vgConectarBD.Execute (vlSql)
'            If (vlCorrCot <> 0) Then
'                vlSql = "Update PT_TMAE_COTIZACION set "
'                vlSql = vlSql & "mto_priunidif = " & Str(Prun_sim) & ","
'                vlSql = vlSql & "mto_ctaindafp = " & Str(Sald_sim) & ","
'                vlSql = vlSql & "mto_rentatmpafp = " & Str(Rete_sim) & " "
'                vlSql = vlSql & "WHERE "
'                vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                vgConectarBD.Execute (vlSql)
'            End If
'        End If
'    End If
'   Dif.Close
'End Function

'Function Tarifa_Todo(Coti, iOperacion) As Boolean
'Dim Tb_Evapar As ADODB.Recordset
'Dim Tb_Evapas As ADODB.Recordset
'Dim Tb_Evapro As ADODB.Recordset
'Dim vlFecha As String
'
'Dim Cp()     As Double, Prodin() As Double
'Dim Flupen() As Double, Flucm() As Double, Exced() As Double
'Dim impres(9, 110) As Double
'Dim Ncorbe(1 To 20) As Integer
'Dim Penben(1 To 20) As Double, Porcbe(1 To 20) As Double
'Dim Coinbe(1 To 20) As String, Codcbe(1 To 20) As String, Sexobe(1 To 20) As String
'Dim Nanbe(1 To 20)  As Integer, Nmnbe(1 To 20) As Integer, Ndnbe(1 To 20) As Integer
'Dim Ijam(1 To 20)   As Integer, Ijmn(1 To 20)  As Integer, Ijdn(1 To 20) As Integer
'Dim Npolbe(1 To 20) As String, derpen(1 To 20) As Integer
'Dim i As Integer
'Dim Totpor As Double
'Dim cob(5) As String, alt1(3) As String, tip(2) As String
'
'Dim Npolca  As String, Mone As String
'Dim Cober   As String, Alt  As String, Indi As Integer, cplan As String
'Dim Nben    As Long
'Dim Nap     As Integer, Nmp As Integer
'Dim Fechan  As Long, Fechap As Long
'Dim Mesdif  As Long, Mesgar As Long
'Dim Bono_Sol1 As Double, Bono_Pesos1 As Double, GtoFun As Double
'Dim CtaInd    As Double, SalCta As Double, Salcta_Sol As Double
'Dim Ffam    As Double
'Dim Prc_Tasa_Afp As Double, Prc_Pension_Afp As Double
'Dim vgs_Coti As String
'
'Dim edbedi As Long, mdif As Long
'Dim large  As Integer
'Dim edaca  As Long, edalca   As Long, edacai  As Long, edacas As Long, edabe As Long, edalbe As Long
'Dim Fasolp As Long, Fmsolp   As Long, Fdsolp  As Long, pergar As Long, numrec As Long, numrep As Long
'Dim nrel   As Long, nmdif    As Long, numbep As Long, ni As Long, ns As Long, nibe As Long, nsbe As Long, limite As Long
'Dim limite1 As Long, limite2 As Long, limite3 As Long, limite4 As Long, imas1 As Long, kdif As Long, nt  As Long
'Dim nmax   As Integer, j As Integer
'Dim rmpol As Double, px As Double, py As Double, qx As Double, relres As Double
'Dim comisi As Double, tasac As Double, timp As Double, tm As Double, tmm As Double
'Dim gasemi As Double
'Dim tm3 As Double, sumapx As Double, sumaqx As Double, actual As Double, actua1 As Double, PenBase As Double, tce As Double
'Dim vpte As Double, difres As Double, difre1 As Double, tir As Double, tinc As Double
'Dim Tasa As Double
'Dim tastce As Double
'Dim tirvta As Double
'Dim tvmax As Double
'Dim vppen As Double, vpcm As Double, penanu As Double, reserva As Double, gastos As Double, rdeuda As Double
'Dim resfin As Double, rend As Double, varrm As Double, resant As Double, flupag As Double, gto As Double
'Dim sumaex As Double, sumaex1 As Double, tirmax As Double
'Dim Sql, Numero As String
'Dim Linea1 As String
'Dim Inserta As String
'Dim Var, Nombre As String
'Dim cuenta As Integer
'Dim nom_moneda As String
'Dim nom_alt As String
'Dim nom_plan As String
'Dim nom_modalidad As String
'Dim vlMargenDespuesImpuesto As Double
'Dim facfam As Double
'Dim fprob As Double
'Dim vlI As Long
'Dim tirmax_ori As Double
'Dim tasac_mod As Double
'Dim vp_tasac As Double
'Dim vlContarMaximo As Long
'Dim Tb_Difben As ADODB.Recordset
'
'    Screen.MousePointer = 11
'    'Inicializacion de variables
'    Tarifa_Todo = False
'
'    '----------------------------------------------------------------
'    'Las Tablas de Mortalidad a utilizar en esta función son ANUALES
'    '----------------------------------------------------------------
'    'Determinar los Finales de Tablas de Mortalidad para cada Tipo
'
'    'Validar Tablas de Mujeres
'    vgFinTabVit_F_A = fgFinTab_Mortal(vgMortalVit_F_A)
'    If (vgFinTabVit_F_A = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Anual de Rtas. Vitalicias de Mujeres.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    vgFinTabTot_F_A = fgFinTab_Mortal(vgMortalTot_F_A)
'    If (vgFinTabTot_F_A = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Anual de Inv. Total de Mujeres.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    vgFinTabPar_F_A = fgFinTab_Mortal(vgMortalPar_F_A)
'    If (vgFinTabPar_F_A = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Anual de Inv. Parcial de Mujeres.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    vgFinTabBen_F_A = fgFinTab_Mortal(vgMortalBen_F_A)
'    If (vgFinTabBen_F_A = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Anual de Beneficiarios de Mujeres.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    'Validar Tablas de Hombres
'    vgFinTabVit_M_A = fgFinTab_Mortal(vgMortalVit_M_A)
'    If (vgFinTabVit_M_A = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Anual de Rtas. Vitalicias de Hombres.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    vgFinTabTot_M_A = fgFinTab_Mortal(vgMortalTot_M_A)
'    If (vgFinTabTot_M_A = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Anual de Inv. Total de Hombres.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    vgFinTabPar_M_A = fgFinTab_Mortal(vgMortalPar_M_A)
'    If (vgFinTabPar_M_A = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Anual de Inv. Parcial de Hombres.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    vgFinTabBen_M_A = fgFinTab_Mortal(vgMortalBen_M_A)
'    If (vgFinTabBen_M_A = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Anual de Beneficiarios de Hombres.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'
'    'Falta validar que el FinTab tome el mayor valor de Tablas de Mortalidad
'    'Tomar el Mayor valor para el Término de la Tabla de Mortalidad
'    vlValor = amax1(vgFinTabVit_F_A, vgFinTabTot_F_A)
'    vlValor = amax1(vlValor, vgFinTabPar_F_A)
'    vlValor = amax1(vlValor, vgFinTabBen_F_A)
'    vlValor = amax1(vlValor, vgFinTabVit_M_A)
'    vlValor = amax1(vlValor, vgFinTabTot_M_A)
'    vlValor = amax1(vlValor, vgFinTabPar_M_A)
'    vlValor = amax1(vlValor, vgFinTabBen_M_A)
'    Fintab = vlValor
'
'    ReDim lx(1 To 2, 1 To 3, 1 To Fintab) As Double
'    ReDim Ly(1 To 2, 1 To 3, 1 To Fintab) As Double
'
'    ReDim Cp(Fintab) As Double
'    ReDim Prodin(Fintab) As Double
'    ReDim Flupen(Fintab) As Double
'    ReDim Flucm(Fintab) As Double
'    ReDim Exced(Fintab) As Double
'
'    'Validar los Topes de Edad de Pago de Pensiones
'    L24 = fgCarga_Param("LI", "L24", vlFecha)
'    If (L24 = (-1000)) Then
'        MsgBox "No existe Edad de tope para los 24 años.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    L21 = fgCarga_Param("LI", "L21", vlFecha)
'    If (L21 = (-1000)) Then
'        MsgBox "No existe Edad de tope para los 21 años.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    L18 = fgCarga_Param("LI", "L18", vlFecha)
'    If (L18 = (-1000)) Then
'        MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    'Las Edades de Tope se encuentran como Anuales, por lo que hay que Mensualizarlas
'
'    L24 = L24 * 12
'    L21 = L21 * 12
'    L18 = L18 * 12
'    cuenta = 0
'    numrec = 0
'    numrep = 0
'    numrec = -1
'
'    '-------------------------------------------------
'    'Leer Tabla de Mortalidad
'    '-------------------------------------------------
'    If (fgBuscarMortalidad(vgMortalVit_F_A, vgMortalTot_F_A, vgMortalPar_F_A, vgMortalBen_F_A, _
'    vgMortalVit_M_A, vgMortalTot_M_A, vgMortalPar_M_A, vgMortalBen_M_A) = False) Then
'        'Se produjeron errores en la obtención de una de las Tablas de Mortalidad
'        Tarifa_Todo = False
'        Exit Function
'    End If
'
'    vgX = 0
'    'Permite determinar el Número de Propuestas existentes
'    Sql = "SELECT count(num_cot) as numero from PT_TMAE_PROPUESTA "
'    Sql = Sql & "where num_cot = '" & Coti & "'"
'    Set vgRs = vgConectarBD.Execute(Sql)
'    If Not (vgRs.EOF) Then
'        vgX = IIf(IsNull(vgRs!Numero), 0, vgRs!Numero)
'    Else
'        MsgBox "Proceso de Cálculo Abortado por inexistencia de Datos.", vbCritical, "Proceso de cálculo Abortado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    vgRs.Close
'
'    If vgX = 0 Then
'        MsgBox "No existen Propuestas de cotizaciones en la Base de Datos.", vbCritical, "Proceso de cálculo Abortado"
'        Tarifa_Todo = False
'        Exit Function
'    Else
'        vlAumento = 100 / vgX 'PORQUE NO TOMA EL VALOR REAL Y DICE DIVISION POR CERO ?????
'    End If
'
'    Frm_Progress.Show
'    Frm_Progress.Caption = "Pogreso del cálulo"
'    Frm_Progress.Refresh
'    Frm_Progress.ProgressBar1.Value = 0
'    Frm_Progress.lbl_progress = "Realizando Cálculos de Tarifa..."
'    Frm_Progress.Refresh
'    Frm_Progress.ProgressBar1.Visible = True
'    Frm_Progress.Refresh
'
'    'Borrar el contenido de la Tabla que registra los valores de la Evaluación
'    vlSql = "delete from PT_TMAE_EVAPRO where num_cot = '" & Coti & "'"
'    vgConectarBD.Execute (vlSql)
'
'    'Borrar el contenido de la Tabla de Evaluaciones de Cotizaciones
'    vlSql = "delete from PT_TMAE_EVACOT where "
'    vlSql = vlSql & "mid(num_cot,1,10) = '" & Mid(Coti, 1, 10) & "' "
'    vgConectarBD.Execute (vlSql)
'
'    For indicador = 1 To vgX
'
'        If Frm_Progress.ProgressBar1.Value + vlAumento < 100 Then
'            Frm_Progress.ProgressBar1.Value = Frm_Progress.ProgressBar1.Value + vlAumento
'            Frm_Progress.Refresh
'        End If
'
'        Sql = "select "
'        Sql = Sql & "num_cot,num_ben,cod_plan,cod_modalidad,"
'        Sql = Sql & "cod_alternativa,cod_indicador,num_mesgar,fec_ingcot,"
'        Sql = Sql & "mto_bonact,mto_bonactpesos,mto_ctaind,mto_priuni,"
'        Sql = Sql & "mto_facpenella,num_mesdif,fec_dev,mto_gassep,"
'        Sql = Sql & "prc_rentaafpori,prc_rentatmp,num_correlcot "
'        Sql = Sql & " from PT_TMAE_PROPUESTA where "
'        Sql = Sql & " num_cot = '" & Coti & "' and "
'        Sql = Sql & " num_pro = " & indicador
'        Set CtaDifpol = vgConectarBD.Execute(Sql)
'        If Not CtaDifpol.EOF Then
'            cuenta = 1
'            vgd_tasa_vta = 0
'            Npolca = CtaDifpol!num_cot
'            Nben = CtaDifpol!num_ben
'            If vgTipoPension = "S" Then Nben = Nben - 1
'            Cober = CtaDifpol!cod_plan
'            Indi = CtaDifpol!cod_indicador
'            Alt = CtaDifpol!cod_alternativa
'            Mesgar = CtaDifpol!num_mesgar
'            Mone = vgMonedaOficial
'            Nap = Mid(CtaDifpol!fec_ingcot, 1, 4)
'            Nmp = Mid(CtaDifpol!fec_ingcot, 5, 2)
'            Bono_Sol1 = CtaDifpol!mto_bonact
'            Bono_Pesos1 = CtaDifpol!mto_bonactpesos
'            CtaInd = CtaDifpol!mto_ctaind
'            Salcta_Sol = CtaDifpol!mto_priuni
'            Ffam = CtaDifpol!mto_facpenella
'            Mesdif = CtaDifpol!num_mesdif
'            Fasolp = Mid(CtaDifpol!fec_dev, 1, 4)   'a_sol_pen
'            Fmsolp = Mid(CtaDifpol!fec_dev, 5, 2)  'm_sol_pen
'            Fdsolp = Mid(CtaDifpol!fec_dev, 7, 2)    'd_sol_pen
'            GtoFun = IIf(IsNull(CtaDifpol!mto_gassep), 0, CtaDifpol!mto_gassep)
'            Prc_Tasa_Afp = CtaDifpol!prc_rentaafpori / 100
'            Prc_Pension_Afp = CtaDifpol!prc_rentatmp / 100
'            vlCorrCot = CtaDifpol!num_correlcot
'            'La conversión de estos códigos debe ser corregida a la Oficial
'            If Cober = "08" Then Cober = "S"
'            If Cober = "06" Then Cober = "I"
'            If Cober = "07" Then Cober = "P"
'            If Cober = "04" Or Cober = "05" Then Cober = "V"
'            SalCta = Salcta_Sol
'
'            If indicador = 1 Then
'                Totpor = 0
'                vgs_Coti = CtaDifpol!num_cot
'
'                'Obtiene los Datos de los Beneficiarios
'                Sql = "select * from PT_TMAE_BENPRO "
'                Sql = Sql & "where num_cot = '" & vgs_Coti & "' "
'                If (vgTipoPension = "S") Then
'                    Sql = Sql & " and num_orden <> 1 "  'el orden empieza de 1 no de 0
'                End If
'                Sql = Sql & "order by num_orden" 'cod_par"
'                Set Tb_Difben = vgConectarBD.Execute(Sql)
'                If (Tb_Difben.EOF) Then
'                    MsgBox "Falta de antecedentes de Beneficiarios en Propuestas de Cotizaciones para realización de cálculo.", vbCritical, "Proceso de Cálculo Abortado"
'                    Tarifa_Todo = False
'                    Exit Function
'                End If
'                i = 1
'                While Not (Tb_Difben.EOF)
'                    Ncorbe(i) = Tb_Difben!cod_par
'                    Porcbe(i) = Tb_Difben!prc_legal
'                    Nanbe(i) = Mid(Tb_Difben!fec_nacben, 1, 4)      'aa_nac
'                    Nmnbe(i) = Mid(Tb_Difben!fec_nacben, 5, 2)      'mm_nac
'                    Ndnbe(i) = Mid(Tb_Difben!fec_nacben, 7, 2)      'mm_nac
'                    derpen(i) = Tb_Difben!cod_derpen
'                    Sexobe(i) = Tb_Difben!cod_sexo
'                    Coinbe(i) = Tb_Difben!cod_sitinv
'                    Codcbe(i) = Tb_Difben!cod_dercre
'                    If Not IsNull(Tb_Difben!Fec_NacHM) Then
'                        Ijam(i) = Mid(Tb_Difben!Fec_NacHM, 1, 4)    'aa_hijom
'                        Ijmn(i) = Mid(Tb_Difben!Fec_NacHM, 5, 2)    'mm_hijom
'                        Ijdn(i) = Mid(Tb_Difben!Fec_NacHM, 7, 2)    'mm_hijom
'                    Else
'                        Ijam(i) = "0000"                            'Year(tb_difben!fec_nachm)   'aa_hijom
'                        Ijmn(i) = "00"                              'Month(tb_difben!fec_nachm)     'mm_hijom
'                        Ijdn(i) = "00"                              'Month(tb_difben!fec_nachm)     'mm_hijom
'                    End If
'                    Npolbe(i) = Tb_Difben!num_cot
'                    Porcbe(i) = Porcbe(i) / 100
'                    If derpen(i) <> 10 Then
'                        If vgTipoPension <> "S" Then
'                            If Ncorbe(i) <> 99 Then
'                                Totpor = Totpor + Porcbe(i)
'                            End If
'                        Else
'                            Totpor = Totpor + Porcbe(i)
'                        End If
'                    End If
'                    Tb_Difben.MoveNext
'                    i = i + 1
'                Wend
'                Tb_Difben.Close
'                Nben = i - 1
'
'                'Determina los Valores de Gastos
'                Sql = "select * from PT_TVAL_GASTO where "
'                Sql = Sql & "cod_moneda = '" & Mone & "' "
'                Set Tb_Evapar = vgConectarBD.Execute(Sql)
'                If Not (Tb_Evapar.EOF) Then
'                    comisi = Format(Tb_Evapar!prc_gasint, "#0.00000")   'Gastos de Intermediación
'                    tasac = Format(Tb_Evapar!prc_ctocap, "#0.00000")    'Tasa de Costo Capital
'                    gastos = Format(Tb_Evapar!mto_gasadm, "#0.00000")   'Gastos de Administración
'                    rdeuda = Format(Tb_Evapar!prc_endeuda, "#0.00000")  'Indice de Endeudamiento
'                    timp = Format(Tb_Evapar!prc_impuesto, "#0.00000")   'Impuesto
'                    gasemi = Format(Tb_Evapar!mto_gasemi, "#0.00000")   'Gastos de Emisión
'                    'I---- ABV 08/06/2004 ---
'                    comisi = comisi / 100
'                    tasac = tasac / 100
'                    timp = timp / 100
'                    'F---- ABV 08/06/2004 ---
'
'                Else
'                    MsgBox "Inexistencia de datos de Parámetros de Evaluación en Dólares.", vbCritical, "Proceso de Cálculo Abortado"
'                    Tarifa_Todo = False
'                    Exit Function
'                End If
'                Tb_Evapar.Close
'
'                'Determina los Valores de Calce
'                Sql = "select * from PT_TVAL_CALCE where "
'                Sql = Sql & "cod_moneda = '" & Mone & "' "
'                Sql = Sql & "order by num_anno "
'                Set Tb_Evapas = vgConectarBD.Execute(Sql)
'                If Not (Tb_Evapas.EOF) Then
'                    vlI = 1
'                    tm = Tb_Evapas!prc_tasamer
'                    While Not (Tb_Evapas.EOF)
'                        vlI = Tb_Evapas!num_anno
'                        Cp(vlI) = Tb_Evapas!prc_cpk
'                        Tb_Evapas.MoveNext
'                    Wend
'                Else
'                    MsgBox "Inexistencia de Datos de Parámetros de Tabla de Calce en Dólares.", vbCritical, "Proceso de cálculo Abortado"
'                    Tarifa_Todo = False
'                    Exit Function
'                End If
'                Tb_Evapas.Close
'
'                'Determina las Tasas de Rentabilidad
'                Sql = "select * from PT_TVAL_RENTABILIDAD where "
'                Sql = Sql & "cod_moneda = '" & Mone & "' "
'                Sql = Sql & "order by num_anno "
'                Set Tb_Evapro = vgConectarBD.Execute(Sql)
'                If Not (Tb_Evapro.EOF) Then
'                    While Not (Tb_Evapro.EOF)
'                        vlI = Tb_Evapro!num_anno
'                        Prodin(vlI) = Tb_Evapro!prc_tasatip / 100
'                        Tb_Evapro.MoveNext
'                    Wend
'                Else
'                    MsgBox "Inexistencia de datos de Parámetros de Productos de Inversiones en Dólares.", vbCritical, "Proceso de Cálculo Abortado"
'                    Tarifa_Todo = False
'
'                    Exit Function
'                End If
'                Tb_Evapro.Close
'            End If
'
'            'I---- ABV 08/06/2004 ---
'            'Inicialización de variables
'            tirvta = 0
'            tinc = 0
'            sumaex = 0
'            sumaex1 = 0
'            tirmax = 0
'            comisi = comisi '/ 100
'            tasac = tasac '/ 100
'            timp = timp '/ 100
'            'F---- ABV 08/06/2004 ---
'
'            tmm = (1 + tm / 100)
'            tm3 = (1.03)
'
'            'Inicializacion de variables
'            For i = 1 To Fintab
'                Exced(i) = 0
'                Flupen(i) = 0
'                Flucm(i) = 0
'            Next i
'            nmax = 0
'            cplan = Cober
'            Fechap = Nap * 12 + Nmp
'            If Indi = 2 Then
'                Mesdif = Mesdif
'            Else
'                Mesdif = 0
'            End If
'            Mesgar = Mesgar / 12
'            ltot = Mesgar + Mesdif
'            If Cober <> "S" Then
'                'Calculo de flujos de vejez e invalidez
'                facfam = Ffam
'                For j = 1 To Nben
'                    Penben(j) = Porcbe(j)
'                    numbep = numbep + 1
'                    If Ncorbe(j) = 99 And j = 1 Then
'                        ni = 0
'                        If Coinbe(j) = "T" Then ni = 1
'                        If Coinbe(j) = "N" Then ni = 2
'                        If Coinbe(j) = "P" Then ni = 3
'                        If ni = 0 Then
'                            X = MsgBox("Error en código de invalidez", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                        ns = 0
'                        If Sexobe(j) = "M" Then ns = 1
'                        If Sexobe(j) = "F" Then ns = 2
'                        If ns = 0 Then
'                            X = MsgBox("Error en código de sexo", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                        Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                        edaca = Fechap - Fechan
'                        difdia = idp - Ndnbe(j)
'                        If difdia > 15 Then edaca = edaca + 1
'                        If edaca < 780 And ns = 1 And ni = 2 Then cplan = "A"
'                        If edaca < 720 And ns = 2 And ni = 2 Then cplan = "A"
'                        edaca = CInt(edaca / 12)
'                        If edaca <= 0 Or edaca > Fintab Then
'                            X = MsgBox("Error en edad de causante ", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                        limite1 = Fintab - edaca - 1
'                        nmax = limite1
'                        For i = 0 To limite1
'                            imas1 = i + 1
'                            edacai = edaca + i
'                            px = lx(ns, ni, edacai) / lx(ns, ni, edaca)
'                            edacas = edacai + 1
'                            qx = ((lx(ns, ni, edacai) - lx(ns, ni, edacas))) / lx(ns, ni, edaca)
'                            Flupen(imas1) = Flupen(imas1) + px * Penben(j)
'                            Flucm(imas1) = Flucm(imas1) + GtoFun * qx
'                        Next i
'                    End If
'                    If Ncorbe(j) <> 99 Then
'                        nibe = 0
'                        If Coinbe(j) = "T" Then nibe = 1
'                        If Coinbe(j) = "N" Then nibe = 2
'                        If Coinbe(j) = "P" Then nibe = 2
'                        If nibe = 0 Then
'                            X = MsgBox("Error en codificacion de invalidez beneficiario", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                        nsbe = 0
'                        If Sexobe(j) = "M" Then nsbe = 1
'                        If Sexobe(j) = "F" Then nsbe = 2
'                        If nsbe = 0 Then
'                            X = MsgBox("Error en codificacion de sexo beneficiario", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                        edabe = Fechap - (Nanbe(j) * 12 + Nmnbe(j))
'                        difdia = idp - Ndnbe(j)
'                        If difdia > 15 Then edabe = edabe + 1
'                        edabe = CInt(edabe / 12)
'                        If edabe < 1 Then edabe = 1
'                        If edabe > Fintab Then
'                            X = MsgBox("Error en Edad del beneficiario es mayor al limite de la tabla de mortalidad", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                        'Calculo de rentas vitalicias
'                        If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
'                           Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
'                           Ncorbe(j) = 41 Or Ncorbe(j) = 42 Or _
'                           (Ncorbe(j) >= 30 And Ncorbe(j) < 40 And _
'                           (Coinbe(j) <> "N" And edabe > L24)) Then
'                            limite1 = Fintab - edabe - 1
'                            nmax = amax0(nmax, CInt(limite1))
'                            For i = 0 To limite1
'                                imas1 = i + 1
'                                edalbe = edabe + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                Flupen(imas1) = Flupen(imas1) + py * Penben(j) * facfam
'                            Next i
'                            limite2 = Fintab - edaca - 1
'                            limite = amin0(limite1, CInt(limite2))
'                            nmax = amax0(nmax, CInt(limite))
'                            For i = 0 To limite
'                                imas1 = i + 1
'                                edalca = edaca + i
'                                edalbe = edabe + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                px = lx(ns, ni, edalca) / lx(ns, ni, edaca)
'                                Flupen(imas1) = Flupen(imas1) - (py * px * Penben(j) * facfam)
'                            Next i
'                            If Ncorbe(j) = 11 Or Ncorbe(j) = 21 Then
'                                'FLUJO POR DERECHO A ACRECER
'                                edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
'                                difdia = idp - Ijdn(j)
'                                If difdia > 15 Then edhm = edhm + 1
'                                If Codcbe(j) <> "N" Then
'                                    If edhm > L24 Then
'                                        nmdif = 0
'                                    Else
'                                        If edhm >= L18 Then
'                                            nmdif = L24 - edhm
'                                        Else
'                                            nmdif = L21 - edhm
'                                        End If
'                                    End If
'                                    Ebedif = edabe + nmdif
'                                    'Probabilidad del beneficiario solo
'                                    limite1 = Fintab - Ebedif - 1
'                                    For i = 0 To limite1
'                                        edalbe = Ebedif + i
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        nmdifi = nmdif + i
'                                        imas1 = nmdifi + 1
'                                        If (nmdifi < nmdiga) Then py = 0
'                                        If nmdifi < perdif Then py = 0
'                                        Flupen(imas1) = Flupen(imas1) + (py * Penben(j) * 0.2)
'                                    Next i
'                                    'Probabilidad conjunta del causante y beneficiario
'                                    Ecadif = edaca + nmdif
'                                    limite4 = Fintab - Ecadif - 1
'                                    limite = amin0(limite1, limite4)
'                                    For i = 0 To limite
'                                        edalca = Ecadif + i
'                                        edalbe = Ebedif + i
'                                        nmdifi = nmdif + i
'                                        imas1 = nmdifi + 1
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        px = lx(ns, ni, edalca) / lx(ns, ni, edaca)
'                                        If (nmdifi < nmdiga) Then py = 0
'                                        If nmdifi < perdif Then py = 0
'                                        Flupen(1, imas1) = Flupen(1, imas1) - (py * px * Penben(j) * 0.2)
'                                    Next i
'                                End If
'                            End If
'                        Else
'                            'Calculo de rentas temporales
'                            If Ncorbe(j) >= 30 And Ncorbe(j) < 40 Then
'                                If edabe <= L24 Then
'                                    If edabe >= L18 Then
'                                        mdif = L24 - edabe
'                                    Else
'                                        mdif = L21 - edabe
'                                    End If
'                                    nmdif = mdif
'                                    limite2 = Fintab - edaca
'                                    limite = amin0(nmdif, CInt(limite2)) - 1
'                                    nmax = amax0(nmax, CInt(limite))
'                                    For i = 0 To limite
'                                        imas1 = i + 1
'                                        edalca = edaca + i
'                                        edalbe = edabe + i
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        px = lx(ns, ni, edalca) / lx(ns, ni, edaca)
'                                        Flupen(imas1) = Flupen(imas1) + (py * Penben(j) - py * px * Penben(j)) * facfam
'                                    Next i
'                                End If
'                                If Coinbe(j) <> "N" Then
'                                    edbedi = edabe + nmdif
'                                    limite3 = Fintab - edbedi - 1
'                                    limite4 = Fintab - (edaca + nmdif) - 1
'                                    nmax = amax0(nmax, CInt(limite3))
'                                    For i = 0 To limite3
'                                        imas1 = nmdif + i + 1
'                                        edalca = edaca + nmdif + i
'                                        edalbe = edbedi + i
'                                        edalca = amin0(edalca, CInt(Fintab))
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        px = lx(ns, ni, edalca) / lx(ns, ni, edaca)
'                                        Flupen(imas1) = Flupen(imas1) + (py - py * px) * Porcbe(j) * facfam
'                                    Next i
'                                End If
'                            'End If
'                            End If
'                        End If
'                    End If
'                Next j
'            Else
'                'Calculo de flujos de Sobrevivencia
'                For j = 1 To Nben
'                    Penben(j) = Porcbe(j)
'                    numbep = numbep + 1
'                    nrel = 0
'                    nibe = 0
'                    If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
'                    If Coinbe(j) = "N" Then nibe = 2
'                    If Coinbe(j) = "P" Then nibe = 2
'                    If nibe = 0 Then
'                        X = MsgBox("Error en codificacion de invalidez beneficiario", vbCritical, "Proceso de cálculo Abortado")
'                        Tarifa_Todo = False
'                        Exit Function
'                    End If
'                    nsbe = 0
'                    If (Sexobe(j) = "M") Then nsbe = 1
'                    If (Sexobe(j) = "F") Then nsbe = 2
'                    If nsbe = 0 Then
'                        X = MsgBox("Error en codificacion de sexo beneficiario", vbCritical, "Proceso de cálculo Abortado")
'                        Tarifa_Todo = False
'                        Exit Function
'                    End If
'                    Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                    edabe = Fechap - Fechan
'                    difdia = idp - Ndnbe(j)
'                    If difdia > 15 Then edabe = edabe + 1
'                    edabe = CInt(edabe / 12)
'                    If edabe > Fintab Then
'                        X = MsgBox("Error en Edad del beneficiario es mayor al limite de la tabla de mortalidad", vbCritical, "Proceso de cálculo Abortado")
'                        Tarifa_Todo = False
'                        Exit Function
'                    End If
'                    If edabe < 1 Then edabe = 1
'                    If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Then
'                        nrel = 1
'                    End If
'                    If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
'                       Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
'                       Ncorbe(j) = 41 Or Ncorbe(j) = 42 Or _
'                       (Ncorbe(j) >= 30 And Ncorbe(j) < 40 And _
'                       (Coinbe(j) <> "N" And edabe > L24)) Then
'                        limite1 = Fintab - edabe - 1
'                        nmax = amax0(nmax, CInt(limite1))
'                        For i = 0 To limite1
'                            imas1 = i + 1
'                            edalbe = edabe + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            If i <= ltot And nrel = 1 Then py = 1
'                            Flupen(imas1) = Flupen(imas1) + py * Penben(j)
'                        Next i
'                        If Ncorbe(j) = 11 Or Ncorbe(j) = 21 Then
'                            If Codcbe(j) = "S" Then
'                                If edhm > L24 Then
'                                    mdif = 0
'                                Else
'                                    If edhm >= L18 Then
'                                        mdif = L24 - edhm
'                                    Else
'                                        mdif = L21 - edhm
'                                    End If
'                                End If
'                                nmdif = mdif
'                                If Cober = 3 Or Cober = 9 Or Cober = 10 Or _
'                                    Cober = 11 Or Cober = 12 Or Cober = 15 Then
'                                    nmdif = amax0(mdif, nmdiga)
'                                End If
'                                Ecadif = edabe + nmdif
'                                limite1 = Fintab - Ecadif - 1
'                                pension = Penben(j) * 0.2
'                                If Cober = 3 Or Cober = 9 Or Cober = 10 Or _
'                                    Cober = 11 Or Cober = 12 Or Cober = 15 Then
'                                    If nmdiga > 0 Then pension = 0
'                                End If
'                                For i = 0 To limite1
'                                    edcadi = Ecadif + i
'                                    py = Ly(nsbe, nibe, edcadi) / Ly(nsbe, nibe, edabe)
'                                    nmdifi = nmdif + i
'                                    imas1 = nmdifi + 1
'                                    If nmdifi < nmdiga And nrel = 2 Then py = 1
'                                    If nmdifi < perdif Then py = 0
'                                    If nmdifi >= nmdiga Then pension = PenBase * Porcbe(j) * 0.2
'                                    Flupen(imas1) = Flupen(imas1) + py * pension
'                                Next i
'                            End If  'Fin de Dacrecer="S"
'                        End If
'                    Else
'                        If Ncorbe(j) >= 30 And Ncorbe(j) < 40 Then
'                            If edabe <= L24 Then
'                                If edabe >= L18 Then
'                                    mdif = L24 - edabe
'                                Else
'                                    mdif = L21 - edabe
'                                End If
'                                nmdif = mdif - 1
'                                nmax = amax0(nmax, CInt(nmdif))
'                                For i = 0 To nmdif
'                                    imas1 = i + 1
'                                    edalbe = edabe + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    Flupen(imas1) = Flupen(imas1) + py * Penben(j)
'                                Next i
'                            End If
'                            If Coinbe(j) <> "N" Then
'                                kdif = mdif
'                                edbedi = edabe + kdif
'                                limite3 = Fintab - edbedi - 1
'                                nmax = amax0(nmax, CInt(limite3))
'                                For i = 0 To limite3
'                                    edalbe = edbedi + i
'                                    imas1 = kdif + i + 1
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    Flupen(imas1) = Flupen(imas1) + py * Penben(j)
'                                Next i
'                            End If
'                        Else
'                            X = MsgBox("Error en Código de relación del beneficiario", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                    End If
'                Next j
'            End If
'
'            '**************************
'            '**************************
'            '
'            'Evaluación de Cotizaciones
'            '
'            '**************************
'            '**************************
'
'            If Cober = "I" Or Cober = "V" Or Cober = "A" Or Cober = "P" Then
'                    For i = 1 To nmax
'                        'I---- ABV 21/02/2004 ---
'                        'If i <= ltot Then
'                        If i <= (ltot + 1) Then
'                        'I---- ABV 21/02/2004 ---
'                            Flupen(i) = amax1(Flupen(i), 1)
'                            'I---- ABV 21/02/2004 ---
'                            'If i <= mesdif Then
'                            If i <= (Mesdif + 1) Then
'                            'F---- ABV 21/02/2004 ---
'                                Flupen(i) = 0
'                                Flucm(i) = 0
'                            End If
'                        End If
'                    Next i
'                End If
'                'I---- DAJ 29/02/2004 ---
'                If Cober = "S" And Mesdif > 0 Then
'                    For i = 1 To nmax
'                        If i <= (Mesdif + 1) Then
'                            Flupen(i) = 0
'                            Flucm(i) = 0
'                        End If
'                    Next i
'                End If
'                'I---- DAJ 29/02/2004 ---
'
'                rmpol = 0
'                nmax = nmax + 1
'                sumapx = 0
'                sumaqx = 0
'                For i = 1 To nmax
'                    actual = (0.8 * Cp(i) / tmm ^ (i - 1)) + ((1 - 0.8 * Cp(i)) / tm3 ^ (i - 1))
'                    sumapx = sumapx + Flupen(i) * actual
'                    actua1 = (0.8 * Cp(i) / tmm ^ (i - 0.5)) + ((1 - 0.8 * Cp(i)) / tm3 ^ (i - 0.5))
'                    sumaqx = sumaqx + Flucm(i) * actua1
'                Next i
'                If sumapx <= 0 Then
'                    PenBase = 0
'                Else
'                    PenBase = (SalCta - sumaqx) / sumapx
'                End If
'                If sumapx <= 0 Then
'                    rmpol = 0
'                Else
'                    rmpol = sumapx * PenBase + sumaqx
'                End If
'                tce = 0
'                vpte = 0
'                difres = 0
'                difre1 = 0
'                tir = 3
'                tinc = 0.1
'225:
'                Tasa = (1 + tir / 100)
'                i = 1
'                For i = 1 To nmax
'                    vpte = vpte + (Flupen(i) * PenBase / Tasa ^ (i - 1)) + (Flucm(i) / Tasa ^ (i - 0.5))
'                Next i
'
'                difres = vpte - rmpol
'                If CDbl(Format(difres, "#0.00000")) >= 0 Then
'                    tir = tir + tinc
'                    If tir > 100 Then
'                        X = MsgBox("TASA TIR MAYOR A 100%", vbCritical)
'                        Tarifa_Todo = False
'                        Exit Function
'                        'GoTo 202 debe salir del programa
'                    End If
'                    difre1 = difres
'                    vpte = 0
'                    GoTo 225
'                End If
'                tce = tir + tinc * (difres / (difre1 - difres))
'                tastce = (1 + tce / 100)
'                tirvta = 3
'222:
'                tvmax = tirvta / 100
'                salcta_eva = SalCta
'                vppen = 0
'                vpcm = 0
'                For i = 1 To nmax
'                    vppen = vppen + Flupen(i) / (1 + tvmax) ^ (i - 1)
'                    vpcm = vpcm + Flucm(i) / (1 + tvmax) ^ (i - 0.5)
'                Next i
'                penanu = (salcta_eva - vpcm) / vppen
'                'I---- DAJ 09/03/2004 ---
'                If Indi = 2 Then
'                    Vpptem = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ Mesdif)) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp)
'                    Add_porc_be = Totpor
'                    If vppen > 0 Then
'                        Rete_sim = CDbl(Format(CDbl((salcta_eva * (1 / (Vpptem + (Prc_Pension_Afp) * vppen)))), "##0.00"))
'                    End If
'                    vld_saldo = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
'                    If vppen > 0 Then
'                        salcta_eva = CDbl(Format(CDbl(salcta_eva - vld_saldo), "#,#0.00"))
'                    End If
'                    pensim = 0
'                    If vppen > 0 Then
'                        penanu = CDbl(Format(CDbl((salcta_eva - vpcm) / vppen), "#,#0.00"))
'                    End If
'                End If
'                reserva = penanu * sumapx + sumaqx
'                Exced(1) = (salcta_eva * (1 - comisi) - gastos - reserva - gasemi) * (1 - timp) - 1 / rdeuda * reserva
'                'I---- ABV 13/03/2004 ---
'                'vld_comision = (salcta_eva * (1 - comisi))
'                vld_comision = (salcta_eva * (comisi))
'                'F---- ABV 13/03/2004 ---
'                'vld_gtosbs = (salcta_eva * facdec)
'                vld_impuesto = (salcta_eva * (1 - comisi) - gastos - reserva - gasemi) * (1 - timp)
'                vld_puesta = 1 / rdeuda * reserva
'
'                'I--- ABV 21/02/2004 ---
'                flupag = penanu * Flupen(1) + Flucm(1)
'                'If indi = "D" And mesdif <> 0 Then
'                '    penanu = 0
'                '    flupag = 0
'                '    Flucm(1) = 0
'''''            '    Else
'''''            '        penanu = (salcta_eva - vpcm) / vppen
'                'End If
'                'F--- ABV 21/02/2004 ---
'
'                relres = 1
'                'I--- ABV 23/02/2004 ---
'                ''I--- ABV 21/02/2004 ---
'                ''resfin = (reserva - penanu - Flucm(1)) * tastce
'                'resfin = (reserva - flupag - flucm(1)) * tastce
'                ''F--- ABV 21/02/2004 ---
'                'F--- ABV 23/02/2004 ---
'
'                resfin = resfin * relres
'                rend = ((reserva + resfin) / 2) * (1 + 1 / rdeuda) * Prodin(1)
'
'                'I---- ABV 23/02/2004 ---
'                'varrm = resfin - reserva
'                'resant = resfin
'                varrm = reserva
'                resant = reserva
'                'F---- ABV 23/02/2004 ---
'
'                'I---- ABV 27/11/2003 ---
'                vlContarMaximo = nmax
'                'I---- ABV 27/11/2003 ---
'
'                For i = 2 To nmax
'                    'I--- ABV 21/02/2004 ---
'                    ''I---- ABV 21/02/2004 ---
'                    ''If indi = "D" And mesdif >= i Then
'                    'If indi = "D" And (mesdif + 1) >= i Then
'                    ''F---- ABV 21/02/2004 ---
'                    '    penanu = 0
'                    '    flupag = 0
'                    'Else
'                    '    penanu = (salcta_eva - vpcm) / vppen
'                    '    flupag = penanu * Flupen(i) + Flucm(i)
'                    'End If
'                    flupag = penanu * Flupen(i) + Flucm(i)
'                    'I---- ABV 23/02/2004 ---
'                    relres = 1
'                    resfin = (resant - flupag) * tastce
'                    resfin = resfin * relres
'                    resfin = amax1(resfin, 0)
'                    'If resfin = 0 Then GoTo 131
'                    varrm = resfin - resant
'                    'F--- ABV 21/02/2004 ---
'                    gto = (gastos * Flupen(i))
'                    Exced(i) = (-flupag - gto - varrm + rend) * (1 - timp) - 1 / rdeuda * varrm
'
'                    'relres = 1
'                    'resfin = (resant - flupag) * tastce
'                    'resfin = resfin * relres
'                    'resfin = amax1(resfin, 0)
'                    'F---- ABV 23/02/2004 ---
'                    'I---- ABV 21/02/2004 ---
'                    'If resfin = 0 Then GoTo 131
'                    If resfin <= 0 Then GoTo 131
'                    'I---- ABV 21/02/2004 ---
'                    'I---- ABV 23/02/2004 ---
'                    'varrm = resfin - resant
'                    'F---- ABV 23/02/2004 ---
'                    rend = ((resant + resfin) / 2) * (1 + 1 / rdeuda) * Prodin(i)
'                    resant = resfin
'                    vld_comision = 0
'                    vld_gtosbs = gto
'                    vld_impuesto = (-flupag - gto - varrm + rend) * (1 - timp)
'                    vld_puesta = 1 / rdeuda * varrm
'
'                    'I---- ABV 27/11/2003 ---
'                    'Se debe cortar la evaluacion del Flujo cuando la
'                    'Rentabilidad se haga Negativa
'                    If rend <= 0 Then
'                        vlContarMaximo = i
'                        Exit For
'                    End If
'                    'F---- ABV 27/11/2003 ---
'                    'I---- ABV 08/03/2004 ---
'                    'Se debe cortar la impresión del Flujo cuando el
'                    'Ajuste de Reservas se haga positivo
'                    If (varrm >= 0 And i > (Mesdif + 1)) Then
'                        vlContarMaximo = i - 1
'                        Exit For
'                    End If
'                    'F---- ABV 08/03/2004 ---
'
'                Next i
'131:
'                sumaex = 0
'                'I---- ABV 27/11/2003 ---
'                'For i = 1 To nmax
'                For i = 1 To vlContarMaximo
'                'F---- ABV 27/11/2003 ---
'                    sumaex = sumaex + Exced(i) / (1 + tasac) ^ i
'                Next i
'                If sumaex >= 0 Then
'                    tirvta = tirvta + tinc
'                    If tirvta > 100 Then
'                        X = MsgBox("TASA TIR MAYOR A 100%", vbCritical)
'                        Tarifa_Todo = False
'                        Exit Function
'                        'GoTo 202 ' debe salir del programa
'                    End If
'                    sumaex1 = sumaex
'                    sumaex = 0
'                    GoTo 222
'                End If
'
'                tirmax = tirvta + tinc * (sumaex / (sumaex1 - sumaex))
'                'Inicio Daniela 20/11/2003
'                tirmax_ori = tirmax
'                'Fin Daniela 20/11/2003
'                tirmax = Format(CDbl(tirmax), "###0.00")
'                tce = Format(tce, "###0.00")
'
'                'Grabar nmax en tabla difpol en campo tasa_simple
'                vgd_tasa_vta = tirmax
'                vgd_tce = tce
'
'                'Se registra la Tasa de Venta = TirMax
'                'o se debe mostrar por Pantalla
'                act = "UPDATE PT_TMAE_PROPUESTA SET"
'                act = act & " prc_tasavta = " & Str(tirmax) & ", "
'                act = act & " mto_resmat = 0, "
'                act = act & " prc_tasatce = " & Str(vgd_tce) & ""
'                act = act & " WHERE num_cot = '" & Coti & "' and "
'                act = act & " num_pro = " & indicador
'                vgConectarBD.Execute (act)
'
'                'vlCorrCot  variable que indica que cotizacion corresponde a la propuesta
'                'vlNumCoti = Mid(Coti, 1, 13) & Format(vlCorrCot, "00") & Mid(Coti, 16, 15)
'                vlNumCoti = Mid(Coti, 1, 10) & Format(vlCorrCot, "00")
'
'                act = "UPDATE PT_TMAE_COTIZACION SET"
'                act = act & " prc_tasavta = " & Str(tirmax) & ", "
'                act = act & " mto_resmat = 0, "
'                act = act & " prc_tasatce = " & Str(vgd_tce) & ""
'                act = act & " WHERE num_cot = '" & vlNumCoti & "'"
'                vgConectarBD.Execute (act)
'
'                'Impresion informe tarifa.lis
'                If Alt = 1 Then ia = 1
'                If Alt = 3 Then ia = 2
'                If cplan = "A" Then IC = 1
'                If cplan = "V" Then IC = 2
'                If cplan = "I" Or cplan = "P" Then IC = 3
'                If cplan = "S" Then IC = 4
'                vppen = 0
'                vpcm = 0
'                For i = 1 To nmax
'                    vppen = vppen + Flupen(i) / (1 + tirmax / 100) ^ (i - 1)
'                    vpcm = vpcm + Flucm(i) / (1 + tirmax / 100) ^ (i - 0.5)
'                Next i
'
'                'Primer periodo
'                salcta_eva = SalCta
'                penanu = (salcta_eva - vpcm) / vppen
'                'Daniela
'                If Indi = 2 Then
'                    Vpptem = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ Mesdif)) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp)
'                    Add_porc_be = Totpor
'                    If vppen > 0 Then
'                        Rete_sim = CDbl(Format(CDbl((salcta_eva * (1 / (Vpptem + (Prc_Pension_Afp) * vppen)))), "##0.00"))
'                    End If
'                    vld_saldo = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
'                    If vppen > 0 Then
'                        salcta_eva = CDbl(Format(CDbl(salcta_eva - vld_saldo), "#,#0.00"))
'                    End If
'                    pensim = 0
'                    If vppen > 0 Then
'                        penanu = CDbl(Format(CDbl((salcta_eva - vpcm) / vppen), "#,#0.00"))
'                    End If
'                End If
'                reserva = penanu * sumapx + sumaqx
'                'I---- ABV 21/02/2004 ---
'                flupag = penanu * Flupen(1) + Flucm(1)
'                'If indi = "D" And mesdif <> 0 Then
'                '    penanu = 0
'                '    flupag = 0
'                'Else
'                '    penanu = (salcta_eva - vpcm) / vppen
'                '    flupag = penanu + flucm(1)
'                'End If
'                'I---- ABV 21/02/2004 ---
'
'                'Guardar Monto de la Reserva
'                act = "UPDATE PT_TMAE_PROPUESTA SET "
'                act = act & " mto_resmat = '" & reserva & "' "
'                act = act & " WHERE num_cot = '" & Coti & "' and "
'                act = act & " num_pro = " & indicador & ""
'                vgConectarBD.Execute (act)
'
'                act = "UPDATE PT_TMAE_COTIZACION SET "
'                act = act & " mto_resmat = '" & reserva & "'"
'                act = act & " WHERE num_cot = '" & vlNumCoti & "'"
'                vgConectarBD.Execute (act)
'
'                'I---- ABV 20/11/2003 ---
'                'If (iOperacion = "I") Then
'
'                    For j = 1 To nmax
'                        impres(1, j) = 0
'                        impres(2, j) = 0
'                        impres(3, j) = 0
'                        impres(4, j) = 0
'                        impres(5, j) = 0
'                        impres(6, j) = 0
'                        impres(7, j) = 0
'                        impres(8, j) = 0
'                        impres(9, j) = 0
'                    Next j
'
'                    gto_inicial = gasemi
'                    Exced(1) = (salcta_eva * (1 - comisi) - gastos - reserva - gasemi) * (1 - timp) - 1 / rdeuda * reserva
'                    Comision = salcta_eva * comisi
'                    margen = salcta_eva - Comision - gastos - gto_inicial - reserva
'                    vlMargenDespuesImpuesto = margen * (1 - timp)
'                    rend = 0
'
'                    impres(1, 1) = salcta_eva
'                    impres(2, 1) = Comision
'                    'I---- ABV 20/11/2003 ---
'                    'impres(3, 1) = penanu
'                    impres(3, 1) = 0
'                    'F---- ABV 20/11/2003 ---
'                    impres(4, 1) = gastos + gto_inicial
'                    impres(5, 1) = reserva
'                    impres(6, 1) = rend
'                    impres(7, 1) = margen
'                    impres(8, 1) = Exced(1)
'                    impres(9, 1) = vlMargenDespuesImpuesto
'
'                    relres = 1
'                    'I--- ABV 23/02/2004 ---
'                    ''I--- ABV 21/02/2004 ---
'                    ''resfin = (reserva - penanu - flucm(1)) * tastce
'                    'resfin = (reserva - flupag - flucm(1)) * tastce
'                    ''F--- ABV 21/02/2004 ---
'                    'F--- ABV 23/02/2004 ---
'
'                    'resfin = resfin * relres
'                    rend = ((reserva + resfin) / 2) * (1 + 1 / rdeuda) * Prodin(1)
'
'                    'I--- ABV 23/02/2004 ---
'                    'varrm = resfin - reserva
'                    'resant = resfin
'                    varrm = reserva
'                    resant = reserva
'                    'F--- ABV 23/02/2004 ---
'
'                    'I---- ABV 27/11/2003 ---
'                    vlContarMaximo = nmax
'                    'I---- ABV 27/11/2003 ---
'
'                    For i = 2 To nmax
'                        flupag = penanu * Flupen(i) + Flucm(i)
'                        'I---- ABV 21/02/2004 ---
'                        ''I---- ABV 21/02/2004 ---
'                        ''If indi = "D" And mesdif >= i Then
'                        'If indi = "D" And (mesdif + 1) >= i Then
'                        ''F---- ABV 21/02/2004 ---
'                        '    penanu = 0
'                        '    flupag = 0
'                        'Else
'                        '    penanu = (salcta_eva - vpcm) / vppen
'                        '    flupag = penanu * Flupen(i) + flucm(i)
'                        'End If
'                        'F---- ABV 21/02/2004 ---
'                        gto = (gastos * Flupen(i))
'
'                        'I---- ABV 23/11/2003 ---
'                        ''I---- ABV 21/11/2003 ---
'                        ''Lo cambie de posición, ya que antes se encontraba donde dice (A)
'                        relres = 1
'                        resfin = (resant - flupag) * tastce
'                        resfin = resfin * relres
'                        varrm = resfin - resant
'                        rend = ((resant + resfin) / 2) * (1 + 1 / rdeuda) * Prodin(i)
'                        ''F---- ABV 21/11/2003 ---
'                        'F---- ABV 23/11/2003 ---
'
'                        Exced(i) = (-flupag - gto - varrm + rend) * (1 - timp) - 1 / rdeuda * varrm
'                        margen = (-flupag - gto - varrm + rend)
'                        impres(1, i) = 0
'                        impres(2, i) = 0
'                        impres(3, i) = flupag
'                        impres(4, i) = gto
'                        impres(5, i) = varrm
'                        impres(6, i) = rend
'                        impres(7, i) = margen
'                        impres(8, i) = Exced(i)
'                        vlMargenDespuesImpuesto = margen * (1 - timp)
'                        impres(9, i) = vlMargenDespuesImpuesto
'
'                        'I---- ABV 23/11/2003 ---
'                        ''I---- ABV 21/11/2003 ---
'                        ''(A)
'                        'relres = 1
'                        'resfin = (resant - flupag) * tastce
'                        'resfin = resfin * relres
'                        'I---- ABV 21/02/2004 ---
'                        If (resfin <= 0) Then
'                            vlContarMaximo = i
'                            Exit For
'                        End If
'                        ''F---- ABV 21/02/2004 ---
'                        'varrm = resfin - resant
'                        'rend = ((resant + resfin) / 2) * (1 + 1 / rdeuda) * prodin(i)
'                        'F---- ABV 23/11/2003 ---
'
'
'                        'I---- ABV 27/11/2003 ---
'                        'Se debe cortar la impresión del Flujo cuando la
'                        'Rentabilidad se haga Negativa
'                        If rend <= 0 Then
'                            vlContarMaximo = i
'                            Exit For
'                        End If
'                        'F---- ABV 27/11/2003 ---
'                        'F---- ABV 21/11/2003 ---
'                        resant = resfin
'
'                        'I---- ABV 08/03/2004 ---
'                        'Se debe cortar la impresión del Flujo cuando el
'                        'Ajuste de Reservas se haga positivo
'                        If (varrm >= 0 And i > (Mesdif + 1)) Then
'                            vlContarMaximo = i - 1
'                            Exit For
'                        End If
'                        'F---- ABV 08/03/2004 ---
'
'                    Next i
'                    'Inicio Daniela 20/11/2003
'                    'If tirmax_ori > CDbl(tirmax) Or tirmax_ori < CDbl(tirmax) Then
''                    If tirmax_ori > penmax Or tirmax_ori < penmin Then
''                        tasac_mod = 0
''                        vp_tasac = 0
''                        difres = 0
''                        difre1 = 0
''                        tasac_mod = 3
''                        tinc = 0.2
''300:
''                        Tasa = tasac_mod
''                        i = 1
''                        'I---- ABV 27/11/2003 ---
''                        'For i = 1 To nmax
''                        For i = 1 To vlContarMaximo
''                        'I---- ABV 27/11/2003 ---
''                            vp_tasac = vp_tasac + Exced(i) / (1 + Tasa / 100) ^ (i - 1)
''                        Next i
''                        difres = vp_tasac
''                        If difres >= 0 Then
''                            tasac_mod = tasac_mod + tinc
''                            If tasac_mod > 100 Then
''                                GoTo 301
''                                tasac_fin = tasac_mod
''                            End If
''                            difre1 = difres
''                            vp_tasac = 0
''                            GoTo 300
''                        End If
''                        tasac_fin = Tasa + tinc * (difres / (difre1 - difres))
''301:
''                        tasac = tasac_fin / 100
''                    End If
'
'                    tasa_tir = Format(CDbl(tasac * 100), "#0.00")
'
'                    vlSql = "update PT_TMAE_PROPUESTA set "
'                    vlSql = vlSql & "prc_tasatir = " & Str(tasa_tir) & " "
'                    vlSql = vlSql & " WHERE num_cot = '" & Coti & "' and "
'                    vlSql = vlSql & " num_pro = " & indicador & ""
'                    vgConectarBD.Execute (vlSql)
'
'                    vlSql = "update PT_TMAE_COTIZACION set "
'                    vlSql = vlSql & "prc_tasatir = " & Str(tasa_tir) & " where "
'                    vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                    vgConectarBD.Execute (vlSql)
'
'                    'Fin Daniela 20/11/2003
'
'                    '--------------------------------------------------------
'                    'ABV : Este número máximo se puede reemplazar por la sgte.
'                    'Fórmula = fintab - edaca -1,
'                    'para lo cual se modificaría el número máximo de impresión
'                    '--------------------------------------------------------
'
'                    'Impresion
'                    '---------
''                    Sql = ""
'                    'Sql = "delete * from rpt_detalle where "
'                    'Sql = Sql & "numcoti = '" & Coti & "'"
'                    'DB_LOC.Execute Sql, dbFailOnError
'
'                    'Set rsreporte = DB_LOC.OpenRecordset("rpt_detalle")
''''                    For i = 1 To nmax
'                        'I---- ABV 18/11/2003 ---
'                        'Realizar modificación de la forma de impresión del
'                        'Detalle de la Evaluación
'                        '1. Comenzar la numeración de los años desde Cero
'                        '2. Mover las pensiones una posición, es decir, comienzan en el año 1
'                        '3. Cuando la Rentabilidad sea negativa, asignarle un Cero
'
''''                        rsreporte.AddNew
''''                        rsreporte.Fields("numcoti") = Coti
''''                        rsreporte.Fields("correla") = indicador
''''                        'I---- ABV 18/11/2003 ---
''''                        'rsreporte.Fields("agno") = i
''''                        rsreporte.Fields("agno") = i - 1
''''                        'F---- ABV 18/11/2003 ---
''''                        rsreporte.Fields("prima") = Format(impres(1, i), "#,#0.00")
''''                        rsreporte.Fields("comision") = Format(impres(2, i), "#,#0.00")
''''                        'I---- ABV 18/11/2003 ---
''''                        'If (i = 1) Then
''''                        '    rsreporte.Fields("pensiones") = Format(0, "#,#0.00")
''''                        'Else
''''                        '    rsreporte.Fields("pensiones") = Format(impres(3, i - 1), "#,#0.00")
''''                            rsreporte.Fields("pensiones") = Format(impres(3, i), "#,#0.00")
''''                        'End If
''''                        'F---- ABV 18/11/2003 ---
''''                        rsreporte.Fields("gastos") = Format(impres(4, i), "#,#0.00")
''''                        rsreporte.Fields("reserva") = Format(impres(5, i), "#,#0.00")
''''                        'I---- ABV 18/11/2003 ---
''''                        If (impres(6, i) < 0) Then
''''                            rsreporte.Fields("ajuste") = Format(0, "#,#0.00")
''''                        Else
''''                            rsreporte.Fields("ajuste") = Format(impres(6, i), "#,#0.00")
''''                        End If
''''                        'I---- ABV 18/11/2003 ---
''''                        rsreporte.Fields("margen") = Format(impres(7, i), "#,#0.00")
''''                        rsreporte.Fields("margendesimp") = Format(impres(9, i), "#,#0.00")
''''                        rsreporte.Fields("excedente") = Format(impres(8, i), "#,#0.00")
''''                        rsreporte.Update
''''                    Next i
'
'                    'vlCorrCot  variable que indica que cotizacion corresponde a la propuesta
'                    'vlNumCoti = Mid(Coti, 1, 13) & Format(vlCorrCot, "00") & Mid(Coti, 16, 15)
'                    vlNumCoti = Mid(Coti, 1, 10) & Format(vlCorrCot, "00")
'
'                    For i = 1 To vlContarMaximo
'                        'Evaluaciones de la Propuesta
'                        vlSql = "insert into PT_TMAE_EVAPRO ("
'                        vlSql = vlSql & "num_cot,num_pro,num_anno,mto_prima,mto_comision,"
'                        vlSql = vlSql & "mto_pension,mto_gasto,mto_renta,"
'                        vlSql = vlSql & "mto_ajuste,mto_margen,mto_margenimp,mto_excedente,"
'                        vlSql = vlSql & "mto_yiecurve, mto_cpk"
'                        vlSql = vlSql & ") values ("
'                        vlSql = vlSql & "'" & Coti & "',"
'                        vlSql = vlSql & "" & indicador & ","
'                        'I---- ABV 18/11/2004 ---
'                        'vlSql = vlSql & "" & i & ","
'                        vlSql = vlSql & "" & i - 1 & ","
'                        'F---- ABV 18/11/2004 ---
'                        vlSql = vlSql & "" & Str(Format(impres(1, i), "#,#0.00")) & "," 'Prima
'                        vlSql = vlSql & "" & Str(Format(impres(2, i), "#,#0.00")) & "," 'Comisión
'                        vlSql = vlSql & "" & Str(Format(impres(3, i), "#,#0.00")) & "," 'Pensiones
'                        vlSql = vlSql & "" & Str(Format(impres(4, i), "#,#0.00")) & "," 'Gastos
'                        If (impres(6, i) < 0) Then 'Reserva
'                            vlSql = vlSql & "" & Str(Format(0, "#,#0.00")) & ","
'                        Else
'                            vlSql = vlSql & "" & Str(Format(impres(6, i), "#,#0.00")) & ","
'                        End If
'                        vlSql = vlSql & "" & Str(Format(impres(5, i), "#,#0.00")) & "," 'Ajuste
'                        vlSql = vlSql & "" & Str(Format(impres(7, i), "#,#0.00")) & "," 'Margen
'                        vlSql = vlSql & "" & Str(Format(impres(9, i), "#,#0.00")) & "," 'Margen Des. Impuesto
'                        vlSql = vlSql & "" & Str(Format(impres(8, i), "#,#0.00")) & "," 'Excedente
'                        vlSql = vlSql & "" & Str(Format(Prodin(i) * 100, "#,#0.00")) & "," 'yiecurve
'                        vlSql = vlSql & "" & Str(Format(Cp(i), "#,#0.000")) & " " 'cpk
'                        vlSql = vlSql & ")"
'                        vgConectarBD.Execute (vlSql)
'
'                        'Evaluaciones de la Cotización
'                        vlSql = "insert into PT_TMAE_EVACOT ("
'                        vlSql = vlSql & "num_cot,num_pro,num_anno,mto_prima,mto_comision,"
'                        vlSql = vlSql & "mto_pension,mto_gasto,mto_ajuste,"
'                        vlSql = vlSql & "mto_renta,mto_margen,mto_margenimp,mto_excedente,"
'                        vlSql = vlSql & "mto_yiecurve, mto_cpk"
'                        vlSql = vlSql & ") values ("
'                        vlSql = vlSql & "'" & vlNumCoti & "',"
'                        vlSql = vlSql & "" & indicador & ","
'                        vlSql = vlSql & "" & i - 1 & ","
'                        vlSql = vlSql & "" & Str(Format(impres(1, i), "#,#0.00")) & ","
'                        vlSql = vlSql & "" & Str(Format(impres(2, i), "#,#0.00")) & ","
'                        vlSql = vlSql & "" & Str(Format(impres(3, i), "#,#0.00")) & ","
'                        vlSql = vlSql & "" & Str(Format(impres(4, i), "#,#0.00")) & ","
'                        If (impres(6, i) < 0) Then
'                            vlSql = vlSql & "" & Str(Format(0, "#,#0.00")) & ","
'                        Else
'                            vlSql = vlSql & "" & Str(Format(impres(6, i), "#,#0.00")) & ","
'                        End If
'                        vlSql = vlSql & "" & Str(Format(impres(5, i), "#,#0.00")) & ","
'                        vlSql = vlSql & "" & Str(Format(impres(7, i), "#,#0.00")) & ","
'                        vlSql = vlSql & "" & Str(Format(impres(9, i), "#,#0.00")) & ","
'                        vlSql = vlSql & "" & Str(Format(impres(8, i), "#,#0.00")) & ","
'                        vlSql = vlSql & "" & Str(Format(Prodin(i) * 100, "#,#0.00")) & "," 'mto_yiecurve
'                        vlSql = vlSql & "" & Str(Format(Cp(i), "#,#0.000")) & "" 'mto_cpk
'                        vlSql = vlSql & ")"
'                        vgConectarBD.Execute (vlSql)
'
'                    Next i
'
'               ' End If
'            'End If
'        End If
'    Next indicador
'    Tarifa_Todo = True
'
'    If (Frm_Progress.ProgressBar1.Value < 100) Then
'       Frm_Progress.ProgressBar1.Value = 100
'    End If
'
'End Function



'Function Renta_Vitalicia(Coti, Codigo_Afp, iRentaAFP, iComision) As Boolean
'Dim vlFecha As String
'
'Dim Npolca As String
'Dim Nap As Integer, Nmp As Integer
'Dim Mesgar As Long, Mesdif As Integer
'Dim Cober As String, Alt As String
'Dim Ffam  As Double, SalCta As Double
'Dim GtoFun  As Double
'Dim swg As String
'
'Dim Npolbe(1 To 20) As String
'Dim Ijam(1 To 20) As Integer, Ijmn(1 To 20) As Integer, Ijdn(1 To 20) As Integer
'Dim Ncorbe(1 To 20) As Integer, Nanbe(1 To 20) As Integer, Nmnbe(1 To 20) As Integer, Ndnbe(1 To 20) As Integer
'Dim Coinbe(1 To 20) As String, Sexobe(1 To 20) As String, Codcbe(1 To 20) As String
'Dim Porcbe(1 To 20) As Double, Penben(1 To 20) As Double
'Dim Nben As Long
'Dim Totpor  As Double
'Dim i As Integer, j As Integer, k As Integer, ll As Integer, ij As Integer
'Dim derpen(1 To 20) As Integer
'Dim Fechan As Long, Fechap As Long, Fechas As Long
'Dim limcic As Long, fapag As Long
'Dim fmpag As Integer, Fasolp As Long
'Dim X As Long
'Dim Sql As String
'Dim perdif As Long, fdpag As Integer
'Dim pergar As Long
'Dim numrec As Integer
'Dim sumaqx As Double, Tasa   As Double
'Dim flumax As Double, qx  As Double, tmtce  As Double, px  As Double, py  As Double, pension  As Double
'Dim renta  As Double, Tarifa As Double
'Dim iciclo As Long, nmdiga As Long, nrel As Long
'Dim nibe As Integer, nsbe As Integer
'Dim limite1 As Long, limite3 As Long
'Dim limite2 As Long, limite As Long
'Dim limite4 As Long
'Dim nmax   As Long, imas1    As Long, mdif As Long, nmdif As Long
'Dim edaca  As Long, edabe    As Long, edalca As Long, edalbe As Long
'Dim edhm   As Long, edacas   As Long, edacai As Long, edbedi As Integer
'Dim m      As Long, kdif     As Long
'Dim nmdifi As Long, Fmsolp   As Long
'Dim ax     As Double
'Dim ni     As Integer, ns As Integer
'Dim facfam As Double
'
'Dim Flupen() As Double
'
'    Renta_Vitalicia = False
'
''   Inicializacion de variables
'    '----------------------------------------------------------------
'    'Las Tablas de Mortalidad a utilizar en esta función son MENSUALES
'    '----------------------------------------------------------------
'    'Determinar los Finales de Tablas de Mortalidad para cada Tipo
'    'FinTab = fgCarga_Param("PC", "LIMM")
'    'Validar Tablas de Mujeres
'    vgFinTabVit_F = fgFinTab_Mortal(vgMortalVit_F)
'    If (vgFinTabVit_F = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Mensual de Rtas. Vitalicias de Mujeres.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    vgFinTabTot_F = fgFinTab_Mortal(vgMortalTot_F)
'    If (vgFinTabTot_F = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Mensual de Inv. Total de Mujeres.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    vgFinTabPar_F = fgFinTab_Mortal(vgMortalPar_F)
'    If (vgFinTabPar_F = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Mensual de Inv. Parcial de Mujeres.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    vgFinTabBen_F = fgFinTab_Mortal(vgMortalBen_F)
'    If (vgFinTabBen_F = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Mensual de Beneficiarios de Mujeres.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    'Validar Tablas de Hombres
'    vgFinTabVit_M = fgFinTab_Mortal(vgMortalVit_M)
'    If (vgFinTabVit_M = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Mensual de Rtas. Vitalicias de Hombres.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    vgFinTabTot_M = fgFinTab_Mortal(vgMortalTot_M)
'    If (vgFinTabTot_M = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Mensual de Inv. Total de Hombres.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    vgFinTabPar_M = fgFinTab_Mortal(vgMortalPar_M)
'    If (vgFinTabPar_M = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Mensual de Inv. Parcial de Hombres.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    vgFinTabBen_M = fgFinTab_Mortal(vgMortalBen_M)
'    If (vgFinTabBen_M = -1) Then
'        MsgBox "No existe la Edad Final de la Tabla de Mortalidad Mensual de Beneficiarios de Hombres.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'
'    'Falta validar que el FinTab tome el mayor valor de Tablas de Mortalidad
'    'Tomar el Mayor valor para el Término de la Tabla de Mortalidad
'    vlValor = amax1(vgFinTabVit_F, vgFinTabTot_F)
'    vlValor = amax1(vlValor, vgFinTabPar_F)
'    vlValor = amax1(vlValor, vgFinTabBen_F)
'    vlValor = amax1(vlValor, vgFinTabVit_M)
'    vlValor = amax1(vlValor, vgFinTabTot_M)
'    vlValor = amax1(vlValor, vgFinTabPar_M)
'    vlValor = amax1(vlValor, vgFinTabBen_M)
'    Fintab = vlValor
'
'    ReDim lx(1 To 2, 1 To 3, 1 To Fintab) As Double
'    ReDim Ly(1 To 2, 1 To 3, 1 To Fintab) As Double
'
'    ReDim Flupen(Fintab) As Double
'
'    'Validar los Topes de Edad de Pago de Pensiones
'    L24 = fgCarga_Param("LI", "L24", vlFecha)
'    If (L24 = (-1000)) Then
'        MsgBox "No existe Edad de tope para los 24 años.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    L21 = fgCarga_Param("LI", "L21", vlFecha)
'    If (L21 = (-1000)) Then
'        MsgBox "No existe Edad de tope para los 21 años.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    L18 = fgCarga_Param("LI", "L18", vlFecha)
'    If (L18 = (-1000)) Then
'        MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    'Las Edades de Tope se encuentran como Anuales, por lo que hay que Mensualizarlas
'    L24 = L24 * 12
'    L21 = L21 * 12
'    L18 = L18 * 12
'
'    cuenta = 0
'    numrec = -1
'    comint = iComision
'
'    '-------------------------------------------------
'    'Leer Tabla de Mortalidad
'    '-------------------------------------------------
'    If (fgBuscarMortalidad(vgMortalVit_F, vgMortalTot_F, vgMortalPar_F, vgMortalBen_F, _
'    vgMortalVit_M, vgMortalTot_M, vgMortalPar_M, vgMortalBen_M) = False) Then
'        'Se produjeron errores en la obtención de una de las Tablas de Mortalidad
'        Renta_Vitalicia = False
'        Exit Function
'    End If
''--------------------------------------------------
'
'    vgX = 0
'    'Inicializacion de variables
'    Sql = "SELECT count(1) as numero FROM PT_TMAE_PROPUESTA "
'    Sql = Sql & " WHERE num_cot= '" & Coti & "'"
'    Set vgRs = vgConectarBD.Execute(Sql)
'    If Not vgRs.EOF And (Not IsNull(vgRs!Numero)) Then
'        vgX = IIf(IsNull(vgRs!Numero), 0, vgRs!Numero)
'    Else
'        MsgBox "Falta de antecedentes en Propuestas de Cotizaciones.", vbCritical, "Proceso de Cálculo Abortado"
'        Renta_Vitalicia = False
'        Exit Function
'    End If
'    vgRs.Close
'
'    If (vgX = 0) Then
'        MsgBox "Inexistencia de antecedentes en Propuesta de Cotizaciones.", vbCritical, "Proceso de Cálculo Abortado"
'        Renta_Vitalicia = False
'        Exit Function
'    Else
'        vlAumento = 100 / vgX
'    End If
'    Frm_Progress.Show
'    Frm_Progress.Refresh
'    Frm_Progress.ProgressBar1.Value = 0
'    Frm_Progress.lbl_progress = "Realizando Cálculo de Renta Vitalicia ..."
'    Frm_Progress.Refresh
'    Frm_Progress.ProgressBar1.Visible = True
'    Frm_Progress.Refresh
'
'    For indicador = 1 To vgX 'numero de cotizaciones
'
'        Sql = "select "
'        Sql = Sql & "num_cot,num_ben,cod_plan,cod_modalidad,"
'        Sql = Sql & "cod_alternativa,cod_indicador,num_mesgar,fec_ingcot,"
'        Sql = Sql & "mto_bonact,mto_bonactpesos,mto_ctaind,mto_priuni,"
'        Sql = Sql & "mto_facpenella,num_mesdif,fec_dev,mto_gassep,"
'        Sql = Sql & "prc_tasavta,num_correlcot "
'        Sql = Sql & " from PT_TMAE_PROPUESTA where "
'        Sql = Sql & " num_cot = '" & Coti & "' and "
'        Sql = Sql & " num_pro = " & indicador & " "
'        Set Tb_Difpol = vgConectarBD.Execute(Sql)
'        If (Tb_Difpol.EOF) Then
'            MsgBox "Inexistencia de antecedentes de Propuestas de Cotizaciones.", vbCritical, "Proceso de Cálculo Abortado"
'            Renta_Vitalicia = False
'            Exit Function
'        Else
'            cuenta = 1
'            Npolca = Tb_Difpol!num_cot
'            Nben = Tb_Difpol!num_ben
'            If vgTipoPension = "S" Then
'                Nben = Nben - 1
'            End If
'            Cober = Tb_Difpol!cod_plan
'
'            Indi = Tb_Difpol!cod_indicador
'            Alt = Tb_Difpol!cod_alternativa
'            pergar = Tb_Difpol!num_mesgar
'            Mone = vgMonedaOficial
'            Nap = Mid(Tb_Difpol!fec_ingcot, 1, 4)  'aa_proceso
'            Nmp = Mid(Tb_Difpol!fec_ingcot, 5, 2)  'mm_proceso
'            Bono_Sol1 = Tb_Difpol!mto_bonact       'bono_sol
'            Bono_Pesos1 = Tb_Difpol!mto_bonactpesos
'            CtaInd = Tb_Difpol!mto_ctaind          'cuenta
'            Salcta_Sol = Tb_Difpol!mto_priuni      'prima_sol
'            Ffam = Tb_Difpol!mto_facpenella        'fac_pella
'            Mesdif = Tb_Difpol!num_mesdif          'meses_dif
'            Fasolp = Mid(Tb_Difpol!fec_dev, 1, 4)  'a_sol_pen
'            Fmsolp = Mid(Tb_Difpol!fec_dev, 5, 2)  'm_sol_pen
'            Fdsolp = Mid(Tb_Difpol!fec_dev, 7, 2)  'd_sol_pen
'            GtoFun = Tb_Difpol!mto_gassep          'sepelio
'            vlCorrCot = Tb_Difpol!num_correlcot
'            If IsNull(Tb_Difpol!prc_tasavta) Then
'                Tasa = 0
'                X = MsgBox("La Tasa de Venta no esta calculada, se abortara el cálculo de esta Cotización.", vbCritical, "Tasa de Venta")
'                Exit For
'            Else
'                Tasa = CDbl(Format(Tb_Difpol!prc_tasavta, "0.00"))
'            End If
'            SalCta = Salcta_Sol
'            ' ABV ---------------------------------------------------------
'            'La conversión de estos códigos debe ser corregida a la Oficial
'            ' ABV ---------------------------------------------------------
'            'If Indi = "1" Then Indi = "I"
'            'If Indi = "2" Then Indi = "D"
'
'            'If Alt = "1" Then Alt = "S"
'            'If Alt = "3" Then Alt = "G"
'
'            'If Cober = "08" Then Cober = "S"
'            'If Cober = "06" Then Cober = "I"
'            'If Cober = "07" Then Cober = "P"
'            'If (Cober = "04" Or Cober = "05") Then Cober = "V"
'            Totpor = 0
'            If indicador = 1 Then
'                'Determina los Datos de los Beneficiarios
'                Sql = "select * from PT_TMAE_BENPRO where "
'                Sql = Sql & " num_cot = '" & Npolca & "' "
'                If vgTipoPension = "S" Then
'                    Sql = Sql & " and num_orden <> 1 "
'                End If
'                Sql = Sql & " order by num_orden "
'                Set Tb_Difben = vgConectarBD.Execute(Sql)
'                If (Tb_Difben.EOF) Then
'                    MsgBox "Inexistencia de Beneficiarios de Propuestas de Cotizaciones.", vbCritical, "Proceso de Cálculo Abortado"
'                    Renta_Vitalicia = False
'                    Exit Function
'                Else
'                    i = 1
'                    While Not (Tb_Difben.EOF)
'                        Ncorbe(i) = CInt(Tb_Difben!cod_par)
'                        Porcbe(i) = CDbl(Tb_Difben!prc_legal)
'                        Nanbe(i) = Mid(Tb_Difben!fec_nacben, 1, 4)
'                        Nmnbe(i) = Mid(Tb_Difben!fec_nacben, 5, 2)
'                        Ndnbe(i) = Mid(Tb_Difben!fec_nacben, 7, 2)
'                        Sexobe(i) = Tb_Difben!cod_sexo
'                        Coinbe(i) = Tb_Difben!cod_sitinv
'                        Codcbe(i) = Tb_Difben!cod_dercre
'                        derpen(i) = Tb_Difben!cod_derpen
'                        If Not IsNull(Tb_Difben!Fec_NacHM) Then
'                            Ijam(i) = Mid(Tb_Difben!Fec_NacHM, 1, 4)
'                            Ijmn(i) = Mid(Tb_Difben!Fec_NacHM, 5, 2)
'                            Ijdn(i) = Mid(Tb_Difben!Fec_NacHM, 7, 2)
'                        Else
'                            Ijam(i) = "0000"                     'Year(tb_difben!fec_nachm)   'aa_hijom
'                            Ijmn(i) = "00"                       'Month(tb_difben!fec_nachm)  'mm_hijom
'                            Ijdn(i) = "00"                       'Month(tb_difben!fec_nachm)  'mm_hijom
'                        End If
'                        Npolbe(i) = Tb_Difben!num_cot
'                        Porcbe(i) = Porcbe(i) / 100
'                        Penben(i) = Porcbe(i)
'
'                        If derpen(i) <> 10 Then
'                            If vgTipoPension = "S" Then
'                                If Ncorbe(i) <> 99 Then
'                                    Totpor = Totpor + Porcbe(i)
'                                End If
'                            Else
'                                Totpor = Totpor + Porcbe(i)
'                            End If
'                        End If
'                        Tb_Difben.MoveNext
'                        i = i + 1
'                    Wend
'                End If
'                Tb_Difben.Close
'                Nben = i - 1
'            End If
'        End If
'
'        'Inicializacion de variables
'        tmtce = (1 + Tasa / 100) ^ (1 / 12)
'        Fechap = Nap * 12 + Nmp
'        perdif = 0
'        If Indi = 2 Then
'            Mesdif = Mesdif * 12
'            perdif = Mesdif
'        End If
'        rmpol = 0
'        For ij = 1 To Fintab
'            Flupen(ij) = 0
'        Next ij
'        Mesgar = pergar
'        If Alt = 1 Then Mesgar = 0
'        If Alt = 3 Or (Alt = 4 And pergar > 0) Then Mesgar = pergar
'        nmdiga = Mesgar + perdif
'        If Cober = 8 Or Cober = 9 Or Cober = 10 Or Cober = 11 Or Cober = 12 Then
'            'PRIMA DE SOBREVIVENCIA
'            For j = 1 To Nben
'                If derpen(j) <> 10 And Ncorbe(j) <> 99 Then
'                    pension = Porcbe(j)
'                    swg = "N"
'                    nrel = 0
'                    nibe = 0
'                    If Coinbe(j) = "S" Or Coinbe(j) = "T" Then nibe = 1
'                    If Coinbe(j) = "N" Then nibe = 2
'                    If Coinbe(j) = "P" Then nibe = 2
'                    If nibe = 0 Then
'                        X = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                        Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    nsbe = 0
'                    If Sexobe(j) = "M" Then nsbe = 1
'                    If Sexobe(j) = "F" Then nsbe = 2
'                    If nsbe = 0 Then
'                        X = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                        Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    'Calculo de la edad de los beneficiarios
'                    Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                    edabe = Fechap - Fechan
'                    difdia = idp - Ndnbe(j)
'                    If difdia > 15 Then edabe = edabe + 1
'                    If edabe > Fintab Then
'                        X = MsgBox("Error edad del beneficiario es mayor a la tabla de mortalidad.", vbCritical, "Proceso de cálculo Abortado")
'                        Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    If edabe < 1 Then edabe = 1
'                    If Ncorbe(j) = 10 Or Ncorbe(j) = 20 Then nrel = 1
'                    If Ncorbe(j) = 11 Or Ncorbe(j) = 21 Then nrel = 2
'                    'PRIMA SOBREVIVENCIA VITALICIA
'                    If Ncorbe(j) = 10 Or Ncorbe(j) = 10 = 20 Or _
'                         Ncorbe(j) = 41 Or Ncorbe(j) = 42 Or _
'                        (Ncorbe(j) >= 30 And Ncorbe(j) < 40 And _
'                        (Coinbe(j) <> "N" And edabe > 24)) Then
'                        pension = Porcbe(j)
'                        limite1 = Fintab - edabe - 1
'                        nmax = amax0(nmax, limite1)
'                        For i = 0 To limite1
'                            imas1 = i + 1
'                            edalbe = edabe + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            If i < nmdiga And nrel = 1 Then py = 1
'                            If i < perdif Then py = 0
'                            Flupen(imas1) = Flupen(imas1) + py * pension
'                        Next i
'                    Else
'                        If Ncorbe(j) = 11 Or Ncorbe(j) = 21 Then
'                            'BENEFICIARIOS PENSION VITALICIA
'                            edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
'                            difdia = idp - Ijdn(j)
'                            If difdia > 15 Then edhm = edhm + 1
'                            limite1 = Fintab - edabe - 1
'                            pension = Porcbe(j)
'                            nmax = amax0(nmax, limite1)
'                            For i = 0 To limite1
'                                imas1 = i + 1
'                                edalbe = edabe + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                If i < nmdiga And nrel = 2 Then py = 1
'                                If i < perdif Then py = 0
'                                Flupen(imas1) = Flupen(imas1) + py * pension
'                            Next i
'                            'DERECHO A ACRECER DE PENSIONES VITALICIAS
'                            If Codcbe(j) <> "N" Then
'                                If edhm > L24 Then
'                                    mdif = 0
'                                Else
'                                    If edhm >= L18 Then
'                                        mdif = L24 - edhm
'                                    Else
'                                        mdif = L21 - edhm
'                                    End If
'                                End If
'                                nmdif = mdif
'                                Ecadif = edabe + nmdif
'                                limite1 = Fintab - Ecadif - 1
'                                nmax = amax0(nmax, limite1)
'                                For i = 0 To limite1
'                                    edcadi = Ecadif + i
'                                    py = Ly(nsbe, nibe, edcadi) / Ly(nsbe, nibe, edabe)
'                                    nmdifi = nmdif + i
'                                    imas1 = nmdifi + 1
'                                    If nmdifi < nmdiga And nrel = 2 Then py = 1
'                                    If nmdifi < perdif Then py = 0
'                                    Flupen(imas1) = Flupen(imas1) + py * pension * 0.2
'                                Next i
'                            End If
'                        Else
'                            If Ncorbe(j) >= 30 And Ncorbe(j) < 40 Then
'                                'PRIMA DE PENSIONES TEMPORALES
'                                If edabe <= L24 Then
'                                    If edabe >= L18 Then
'                                        mdif = L24 - edabe
'                                    Else
'                                        mdif = L21 - edabe
'                                    End If
'                                    nmdif = mdif - 1
'                                    nmax = amax0(nmax, nmdif)
'                                    For i = 0 To nmdif
'                                        imas1 = i + 1
'                                        edalbe = edabe + i
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        If (swg = "S" And i < nmdiga) Then py = 1
'                                        If i < perdif Then py = 0
'                                        Flupen(imas1) = Flupen(imas1) + py * pension
'                                    Next i
'                                    'PRIMA DE HIJOS INVALIDOS
'                                    If (Coinbe(j) <> "N") Then
'                                        kdif = mdif
'                                        edbedi = edabe + kdif
'                                        limite3 = Fintab - edbedi - 1
'                                        pension = Porcbe(j)
'                                        nmax = amax0(nmax, limite3)
'                                        For i = 0 To limite3
'                                            edalbe = edbedi + i
'                                            nmdifi = i + kdif
'                                            imas1 = nmdifi + 1
'                                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                            If nmdifi < perdif Then py = 0
'                                            Flupen(imas1) = Flupen(imas1) + py * pension
'                                        Next i
'                                    End If
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'            Next j
'            ax = 0
'            nmdiga = perdif + Mesgar
''            If Alt = 3 Or (Alt = 4 And Mesgar > 0) Then
''                For LL = 1 To nmdiga
''                    flumax = amax1(1, Flupen(LL))
''                    If LL <= perdif Then flumax = 0
''                    ax = ax + flumax / tmtce ^ (LL - 1)
''                Next LL
''                For LL = nmdiga + 1 To nmax
''                    ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
''                Next LL
''            Else
'                For ll = 1 To nmax
'                    flumax = Flupen(ll)
'                    If ll <= perdif Then flumax = 0
'                    ax = ax + flumax / tmtce ^ (ll - 1)
'                Next ll
''            End If
'
'            'renta = SalCta / ax
'            'Tarifa = ax
'            If Indi = 1 Then
'                If ax <= 0 Then
'                    renta = 0
'                    ax = 0
'                Else
'                    sumaqx = CDbl(Format(sumaqx, "#,#0.000000"))
'                    ax = CDbl(Format(ax, "#,#0.000000"))
'                    renta = CDbl(Format((SalCta / ax), "#,#0.00"))
'                End If
'            Else
'                ax = CDbl(Format(ax, "#,#0.000000"))
'                renta = 0
'            End If
'            'Grabacion de resultados a difpol
'            vlSql = "update PT_TMAE_PROPUESTA set "
'            vlSql = vlSql & "mto_priunisim = " & Str(Format(ax, "##0.000000")) & ","
'            vlSql = vlSql & "mto_pensim = " & Str(Format(renta, "##0.00")) & ","
'            vlSql = vlSql & "mto_pengar = " & Str(Format(0, "##0.00")) & " "
'            vlSql = vlSql & "WHERE "
'            vlSql = vlSql & "num_cot = '" & Coti & "' and "
'            vlSql = vlSql & "num_pro = " & indicador & ""
'            vgConectarBD.Execute (vlSql)
'            vlNumCoti = Mid(Coti, 1, 10) & Format(vlCorrCot, "00")
'            If (vlCorrCot <> 0) Then
'                vlSql = "update PT_TMAE_COTIZACION set "
'                vlSql = vlSql & "mto_priunisim = " & Str(Format(ax, "##0.000000")) & ","
'                vlSql = vlSql & "mto_pensim = " & Str(Format(renta, "##0.00")) & ","
'                vlSql = vlSql & "mto_pengar = " & Format(0, "##0.00") & " "
'                vlSql = vlSql & "WHERE "
'                vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                vgConectarBD.Execute (vlSql)
'            End If
'            If Indi = 2 Then
'                Call Calcula_Diferida(Coti, Codigo_Afp, comint, indicador, iRentaAFP)
'            Else
'                vlSql = "update PT_TMAE_PROPUESTA set "
'                vlSql = vlSql & "mto_valprepentmp = 0,"
'                vlSql = vlSql & "mto_priunidif = 0,"
'                vlSql = vlSql & "mto_ctaindafp = 0,"
'                vlSql = vlSql & "mto_rentatmpafp = 0 "
'                vlSql = vlSql & "WHERE "
'                vlSql = vlSql & "num_cot = '" & Coti & "' and "
'                vlSql = vlSql & "num_pro = " & indicador & ""
'                vgConectarBD.Execute (vlSql)
'                If (vlCorrCot <> 0) Then
'                    'si la propuesta no es oficial, la cotizacion no existe ,por lo tanto no modifica la cotizacion
'                    vlSql = "update PT_TMAE_COTIZACION set "
'                    vlSql = vlSql & "mto_valprepentmp = 0,"
'                    vlSql = vlSql & "mto_priunidif = 0,"
'                    vlSql = vlSql & "mto_ctaindafp = 0,"
'                    vlSql = vlSql & "mto_rentatmpafp = 0 "
'                    vlSql = vlSql & "WHERE "
'                    vlSql = vlSql & "num_cot = '" & vlNumCoti & "' "
'                    vgConectarBD.Execute (vlSql)
'                End If
'            End If
'            'Exit Function
'        Else
'            'FLUJOS DE RENTAS VITALICIAS DE VEJEZ E INVALIDEZ
'            For ij = 1 To Fintab
'                Flupen(ij) = 0
'            Next ij
'            facfam = Ffam
'            'Definicion del periodo garantizado en 0 para las
'            'alternativas Simple o pensiones con distinto porcentaje legal
'            For j = 1 To Nben
'                If derpen(j) <> 10 Then
'                    pension = Porcbe(j)
'                    'CALCULO DE LA PRIMA DEL AFILIADO
'                    If (Ncorbe(j) = 0 Or Ncorbe(j) = 99) And j = 1 Then
'                        ni = 0
'                        If Coinbe(j) = "S" Or Coinbe(j) = "T" Then ni = 1
'                        If Coinbe(j) = "N" Then ni = 2
'                        If Coinbe(j) = "P" Then ni = 3
'                        If ni = 0 Then
'                            X = MsgBox("Error de códificación de tipo de inavlidez", vbCritical, "Proceso de Cálculo Abortado")
'                            Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                        ns = 0
'                        If Sexobe(j) = "M" Then ns = 1
'                        If Sexobe(j) = "F" Then ns = 2
'                        If ns = 0 Then
'                            X = MsgBox("Error de códificación de sexo.", vbCritical, "Proceso de Cálculo Abortado")
'                            Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                        'Calculo de edad del causante
'                        Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                        edaca = Fechap - Fechan
'                        difdia = idp - Ndnbe(j)
'                        If difdia > 15 Then edaca = edaca + 1
'                        'If (edaca <= 216) Then edaca = 216
'                        If edaca > Fintab Then
'                            X = MsgBox("Error en edad del beneficiario mayor a tabla de mortalidad", vbCritical, "Proceso de cálculo Abortado")
'                            Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                        sumaqx = 0
'                        limite1 = Fintab - edaca - 1
'                        nmax = limite1
'                        For i = 0 To limite1
'                            imas1 = i + 1
'                            edacai = edaca + i
'                            px = lx(ns, ni, edacai) / lx(ns, ni, edaca)
'                            edacas = edacai + 1
'                            qx = ((lx(ns, ni, edacai) - lx(ns, ni, edacas))) / lx(ns, ni, edaca)
'                            Flupen(imas1) = Flupen(imas1) + px * pension
'                            sumaqx = sumaqx + GtoFun * qx / tmtce ^ (i + 0.5)
'                        Next i
'                    End If
'                    If Nben > 1 And (Ncorbe(j) <> 0 And Ncorbe(j) <> 99) Then
'                        'PRIMA DE LOS BENEFICIARIOS
'                        nibe = 0
'                        If Coinbe(j) = "S" Or Coinbe(j) = "T" Then nibe = 1
'                        If Coinbe(j) = "N" Then nibe = 2
'                        If Coinbe(j) = "P" Then nibe = 3
'                        If nibe = 0 Then
'                            X = MsgBox("Error de códificación de tipo de invalidez.", vbCritical, "Proceso de cálculo Abortado")
'                            Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                        nsbe = 0
'                        If Sexobe(j) = "M" Then nsbe = 1
'                        If Sexobe(j) = "F" Then nsbe = 2
'                        If nsbe = 0 Then
'                            X = MsgBox("Error de códificación de sexo.", vbCritical, "Proceso de cálculo Abortado")
'                            Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                        'Calculo de la edad del beneficiario
'                        edabe = Fechap - (Nanbe(j) * 12 + Nmnbe(j))
'                        difdia = idp - Ndnbe(j)
'                        If difdia > 15 Then edabe = edabe + 1
'                        If edabe < 1 Then edabe = 1
'                        If edabe > Fintab Then
'                            X = MsgBox("Error Edad del beneficario es mayor a la tabla.", vbCritical, "Proceso de cálculo Abortado")
'                            Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                        If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
'                            Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
'                            Ncorbe(j) = 41 Or Ncorbe(j) = 42 Or _
'                            (Ncorbe(j) >= 30 And Ncorbe(j) < 40 And _
'                            (Coinbe(j) <> "N" And edabe > 24)) Then
'                            'FLUJOS DE VIDAS CONJUNTAS VITALICIAS
'                            'Probabilidad del beneficiario solo
'                            limite1 = Fintab - edabe - 1
'                            nmax = amax0(nmax, limite1)
'                            For i = 0 To limite1
'                                imas1 = i + 1
'                                edalbe = edabe + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                Flupen(imas1) = Flupen(imas1) + py * pension * facfam
'                            Next i
'                            'Probabilidad conjunta de causante y beneficiario
'                            limite2 = Fintab - edaca - 1
'                            limite = amin0(limite1, limite2)
'                            nmax = amax0(nmax, limite)
'                            For i = 0 To limite
'                                imas1 = i + 1
'                                edalca = edaca + i
'                                edalbe = edabe + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                px = lx(ns, ni, edalca) / lx(ns, ni, edaca)
'                                Flupen(imas1) = Flupen(imas1) - (py * px * pension * facfam)
'                            Next i
'                            If Ncorbe(j) = 11 Or Ncorbe(j) = 21 Then
'                                'FLUJO POR DERECHO A ACRECER
'                                edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
'                                difdia = idp - Ijdn(j)
'                                If difdia > 15 Then edhm = edhm + 1
'                                If Codcbe(j) <> "N" Then
'                                    If edhm > L24 Then
'                                        nmdif = 0
'                                    Else
'                                        If edhm >= L18 Then
'                                            nmdif = L24 - edhm
'                                        Else
'                                            nmdif = L21 - edhm
'                                        End If
'                                    End If
'                                    Ebedif = edabe + nmdif
'                                    'Probabilidad del beneficiario solo
'                                    limite1 = Fintab - Ebedif - 1
'                                    nmax = amax0(nmax, limite1)
'                                    For i = 0 To limite1
'                                        edalbe = Ebedif + i
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        nmdifi = nmdif + i
'                                        imas1 = nmdifi + 1
'                                        Flupen(imas1) = Flupen(imas1) + (py * pension * 0.2 * facfam)
'                                    Next i
'                                    'Probabilidad conjunta del causante y beneficiario
'                                    Ecadif = edaca + nmdif
'                                    limite4 = Fintab - Ecadif - 1
'                                    limite = amin0(limite1, limite4)
'                                    nmax = amax0(nmax, limite)
'                                    For i = 0 To limite
'                                        edalca = Ecadif + i
'                                        edalbe = Ebedif + i
'                                        nmdifi = nmdif + i
'                                        imas1 = nmdifi + 1
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        px = lx(ns, ni, edalca) / lx(ns, ni, edaca)
'                                        Flupen(imas1) = Flupen(imas1) - (py * px * pension * 0.2 * facfam)
'                                    Next i
'                                End If
'                            End If
'                        Else
'                            'PRIMA RENTAS TEMPORALES
'                            If Ncorbe(j) >= 30 And Ncorbe(j) < 40 And _
'                                edabe < L24 Then
'                                If edabe >= L18 Then
'                                    mdif = L24 - edabe
'                                Else
'                                    mdif = L21 - edabe
'                                End If
'                                nmdif = mdif
'                                'Probabilidad conjunta del causante y beneficiario
'                                limite2 = Fintab - edaca
'                                limite = amin0(nmdif, limite2) - 1
'                                nmax = amax0(nmax, limite)
'                                For i = 0 To limite
'                                    imas1 = i + 1
'                                    edalca = edaca + i
'                                    edalbe = edabe + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    px = lx(ns, ni, edalca) / lx(ns, ni, edaca)
'                                    Flupen(imas1) = Flupen(imas1) + (py * pension - py * px * pension) * facfam
'                                Next i
'                                'PRIMA DEL HIJO INVALIDO
'                                If Coinbe(j) <> "N" Then
'                                    'Probabilidad conjunta del causante y beneficiario
'                                    edbedi = edabe + nmdif
'                                    limite3 = Fintab - edbedi - 1
'                                    limite4 = Fintab - (edaca + nmdif) - 1
'                                    nmax = amax0(nmax, limite3)
'                                    For i = 0 To limite3
'                                        nmdifi = nmdif + i
'                                        imas1 = nmdifi + 1
'                                        edalca = edaca + nmdif + i
'                                        edalbe = edbedi + i
'                                        edalca = amin0(edalca, Fintab)
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        px = lx(ns, ni, edalca) / lx(ns, ni, edaca)
'                                        Flupen(imas1) = Flupen(imas1) + (py - py * px) * pension * facfam
'                                    Next i
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'            Next j
'
'            ax = 0
'            flumax = 0
'            If Alt = 3 Or (Alt = 4 And pergar > 0) Then
'                ax = 0
'                For ll = 1 To nmdiga
'                    flumax = amax1(1, Flupen(ll))
'                    If ll <= perdif Then flumax = 0
'                    ax = ax + flumax / tmtce ^ (ll - 1)
'                Next ll
'                For ll = nmdiga + 1 To nmax
'                    ax = ax + Flupen(ll) / tmtce ^ (ll - 1)
'                Next ll
'            Else
'                For ll = 1 To nmax
'                    If ll <= perdif Then flumax = 0
'                    ax = ax + Flupen(ll) / tmtce ^ (ll - 1)
'                Next ll
'            End If
'            If Indi = 1 Then
'                If ax <= 0 Then
'                    renta = 0
'                    ax = 0
'                Else
'                    sumaqx = CDbl(Format(sumaqx, "#,#0.000000"))
'                    ax = CDbl(Format(ax, "#,#0.000000"))
'                    renta = CDbl(Format(((SalCta - sumaqx) / ax), "#,#0.00"))
'                End If
'            Else
'                ax = CDbl(Format(ax, "#,#0.000000"))
'                renta = 0
'            End If
'            act = "UPDATE PT_TMAE_PROPUESTA SET "
'            act = act & "mto_priunisim = " & Str(Format(ax, "##0.000000")) & ", "
'            act = act & "mto_pensim = " & Str(Format(renta, "##0.00")) & ", "
'            act = act & "mto_pengar = " & Str(Format(sumaqx, "##0.00")) & " "
'            act = act & "WHERE "
'            act = act & "num_cot = '" & Coti & "' and "
'            act = act & "num_pro = " & indicador & ""
'            vgConectarBD.Execute (act)
'
'            vlNumCoti = Mid(Coti, 1, 10) & Format(vlCorrCot, "00")
'            If (vlCorrCot <> 0) Then
'                act = "UPDATE PT_TMAE_COTIZACION SET "
'                act = act & "mto_priunisim = " & Str(Format(ax, "#,#0.000000")) & ", "
'                act = act & "mto_pensim = " & Str(Format(renta, "#,#0.00")) & ", "
'                act = act & "mto_pengar = " & Str(Format(sumaqx, "#,#0.00")) & " "
'                act = act & "WHERE "
'                act = act & "num_cot = '" & vlNumCoti & "'"
'                vgConectarBD.Execute (act)
'            End If
'            If Indi = 2 Then
'                Call Calcula_Diferida(Coti, Codigo_Afp, comint, indicador, iRentaAFP)
'            Else
'                vlSql = "update PT_TMAE_PROPUESTA set "
'                vlSql = vlSql & "mto_valprepentmp = 0,"
'                vlSql = vlSql & "mto_priunidif = 0,"
'                vlSql = vlSql & "mto_ctaindafp = 0,"
'                vlSql = vlSql & "mto_rentatmpafp = 0 "
'                vlSql = vlSql & "WHERE "
'                vlSql = vlSql & "num_cot = '" & Coti & "' and "
'                vlSql = vlSql & "num_pro = " & indicador & ""
'                vgConectarBD.Execute (vlSql)
'
'                If (vlCorrCot <> 0) Then
'                    vlSql = "update PT_TMAE_COTIZACION set "
'                    vlSql = vlSql & "mto_valprepentmp = 0,"
'                    vlSql = vlSql & "mto_priunidif = 0,"
'                    vlSql = vlSql & "mto_ctaindafp = 0,"
'                    vlSql = vlSql & "mto_rentatmpafp = 0 "
'                    vlSql = vlSql & "WHERE "
'                    vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                    vgConectarBD.Execute (vlSql)
'                End If
'            End If
'        End If
'        If Frm_Progress.ProgressBar1.Value + vlAumento < 100 Then
'            Frm_Progress.ProgressBar1.Value = Frm_Progress.ProgressBar1.Value + vlAumento
'            Frm_Progress.Refresh
'        End If
'    Next indicador
'    Renta_Vitalicia = True
'    If Frm_Progress.ProgressBar1.Value < 100 Then
'       Frm_Progress.ProgressBar1.Value = 100
'       Frm_Progress.Refresh
'    End If
'    Unload Frm_Progress
'End Function

Function AbrirBaseDeDatos(oConBD As ADODB.Connection)

    'Por defecto supone que falla la Conexión
    AbrirBaseDeDatos = False

On Local Error GoTo Err_ConsultaBD

    Set oConBD = New ADODB.Connection
    oConBD.ConnectionString = "driver={Sql Server}; server=" & vgNombreServidor & ";UID=" & vgNombreUsuario & ";PWD=" & vgPassWord & ";database=" & vgNombreBaseDatos
    'oConBD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & vgRutaBasedeDatos & ";Persist Security Info=False"
    'oConBD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vgRutaBasedeDatos & ";Persist Security Info=False"
    oConBD.ConnectionTimeout = 1800
    oConBD.CommandTimeout = 1800
    oConBD.Open
    'oConBD.BeginTrans
    'La Conexión fue realizada
    AbrirBaseDeDatos = True
   ' oConBD.CommitTrans
    Exit Function

Err_ConsultaBD:
    'MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : " & vbCrLf & Err.Description, vbCritical, "Error de Consulta a Base de Datos"
    'Err.Clear
    'Text3.Text = "Errores en la Impresión de Solicitudes"
    MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : [ " & Err & Space(4) & Err.Description & " ]"
End Function

'Function AbrirBaseDeDatos(RutaBasedeDatos As String)
'On Error GoTo ErrorAbrirBaseAccess
'
'    Set BaseDeDatos = Wrk.OpenDatabase(RutaBasedeDatos, False, False)
'    'Set vgConectarBD = Wrk.OpenDatabase(RutaBasedeDatos, False, False)
'
'    'Set BaseDeDatos = OpenDatabase(RutaBasedeDatos, False, False)
'    'Set vgConectarBD = OpenDatabase(RutaBasedeDatos, False, False)
'
'    'Set vgDb = New ADODB.Connection
'    'Set vgRs = New ADODB.Recordset
'    'vgDb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & RutaBasedeDatos & ";Persist Security Info=False"
'    'vgDb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaBasedeDatos & ";Persist Security Info=False"
'    'VGDb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & RutaBasedeDatos
'    'vgDb.Open
'
'Exit Function
'ErrorAbrirBaseAccess:
'     MsgBox "No Se Logro Conexión Con La Base de Datos", vbCritical, "Error de Conexión"
'     End
'End Function

Function AbrirBaseDeDatos_Aux(RutaBasedeDatos As String)

    'Por defecto supone que falla la Conexión
    AbrirBaseDeDatos_Aux = False

On Local Error GoTo Err_ConsultaBD

    Set oConBD = New ADODB.Connection
    'oConBD.ConnectionString = "driver={Sql Server}; server=" & vgNombreServidor & ";UID=" & vgNombreUsuario & ";PWD=" & vgPassWord & ";database=" & vgNombreBaseDatos
    oConBD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & vgRutaBasedeDatos & ";Persist Security Info=False"
    oConBD.ConnectionTimeout = 1800
    oConBD.CommandTimeout = 1800
    oConBD.Open
    'oConBD.BeginTrans
    'La Conexión fue realizada
    AbrirBaseDeDatos_Aux = True
   ' oConBD.CommitTrans
    Exit Function

Err_ConsultaBD:
    'MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : " & vbCrLf & Err.Description, vbCritical, "Error de Consulta a Base de Datos"
    'Err.Clear
    'Text3.Text = "Errores en la Impresión de Solicitudes"
    MsgBox "La consulta a la Base de Datos no fue llevada a cabo. Error : [ " & Err & Space(4) & Err.Description & " ]"
End Function

'Function CerrarBaseDeDatos()
'    BaseDeDatos.Close
'End Function
Function CerrarBaseDeDatos(iConexion)
    iConexion.Close
End Function

'Function CerrarBaseDeDatos_Aux()
''    vgConectarBD.Close
'End Function

Function LeeArchivoIni(ByVal Seccion$, ByVal Item$, ByVal Default$, ByVal NombreArchivo$) As String
Dim temp As String
Dim x   As Integer

    temp = String$(2048, 32)
    x = GetPrivateProfileString(Seccion$, Item$, Default$, temp, Len(temp), NombreArchivo$)
    LeeArchivoIni = Mid$(temp, 1, x)

End Function

Function EscribeArchivoIni(ByVal Seccion$, ByVal Item$, ByVal Default$, ByVal NombreArchivo$) As Integer
    EscribeArchivoIni = WritePrivateProfileString(Seccion$, Item$, Default$, NombreArchivo$)
End Function

'--------------------------------------------------------
'Permite determinar si el Archivo Existe
'--------------------------------------------------------
Function fgExiste(iArchivo) As Boolean
On Local Error GoTo Err_NoExiste
    
    If Dir$(iArchivo) = "" Then
        fgExiste = False
    Else
        fgExiste = True
    End If
    Exit Function
    
Err_NoExiste:
    fgExiste = False
    Exit Function
End Function

'--------------------------------------------------------
'Permite Encriptar la Password del Usuario que es registrada
'en la Base de Datos
'--------------------------------------------------------
Function fgEncPassword(iContraseña) As String
Dim iPassword As String
Dim iOculto As String
On Error GoTo Err_Encriptar
    
    iPassword = ""
    iOculto = ""
    For vgI = 1 To Len(UCase(iContraseña))
        iOculto = Chr(255 - Asc(UCase(Mid(iContraseña, vgI, 1))))
        Asc (iOculto)
        iPassword = iPassword + iOculto
        Chr (15)
    Next
    fgEncPassword = iPassword

Exit Function
Err_Encriptar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'--------------------------------------------------------
'Permite Desencriptar la Password del Usuario registrada
'en la Base de Datos
'--------------------------------------------------------
Function fgDesPassword(iContraseña) As String
Dim iPassword As String
Dim iOculto As String
On Error GoTo Err_Desencriptar
    
    iPassword = ""
    iOculto = ""
    For vgI = 1 To Len(UCase(iContraseña))
        iOculto = Chr(255 - Asc(UCase(Mid(iContraseña, vgI, 1))))
        Asc (iOculto)
        iPassword = iPassword + iOculto
        Chr (15)
    Next
    fgDesPassword = iPassword

Exit Function
Err_Desencriptar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'----------------------------------------------------------
'Función que permite centralizar el Formulario o Pantalla
'----------------------------------------------------------
Sub Center(Frm As Form)
    Frm.Left = (Screen.Width - Frm.Width) \ 2
    Frm.Top = (Screen.Height - (Frm.Height + 2000)) \ 2
End Sub

'----------------------------------------------------------
' Determinar el Nombre y Cargo del Auditor de la Reserva
'----------------------------------------------------------
Function fgApoderado()
On Error GoTo Err_BuscarApoderado

    vgNombreApoderado = ""
    vgCargoApoderado = ""

    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)

    'Abrir la Conexión a la Base de Datos
'''    If Not AbrirBaseDeDatos(vgConectarBD) Then
'''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
'''        Exit Function
'''    End If

    'Determinar Nombre y Cargo del Auditor
    vgQuery = "Select * from PD_TMAE_APODERADO "
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!gls_nomApo) Then vgNombreApoderado = Trim(vgRs!gls_nomApo)
        If Not IsNull(vgRs!gls_carApo) Then vgCargoApoderado = Trim(vgRs!gls_carApo)
    End If
    vgRs.Close

    'call CerrarBaseDeDatos (vgconectarbd)

    'Cerrar Conexión
'''    Call CerrarBaseDeDatos(vgConectarBD)

Exit Function
Err_BuscarApoderado:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

''----------------------------------------------------------
'' Determinar el Valor del Dolar para una Fecha dada
''----------------------------------------------------------
'Function fgBuscarMoneda(iMes, iAño, iMoneda)
'On Error GoTo Err_BuscarMoneda
'
'    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)
''''    If Not AbrirBaseDeDatos(vgConexionBD) Then
''''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
''''        Exit Function
''''    End If
'
'    vgValorMoneda = 0
'    vgError = 0
'    'Buscar el valor del Dolar a la Fecha de Proceso
'    vgQuery = "select mto_mes" & Format(iMes, "#0") & " as valor "
'    vgQuery = vgQuery & "FROM MA_TVAL_MONEDA WHERE "
'    vgQuery = vgQuery & "num_anno = " & iAño & " AND "
'    vgQuery = vgQuery & "cod_moneda = '" & iMoneda & "'"
'    Set vgRs = vgConexionBD.Execute(vgQuery)
'    If Not (vgRs.EOF) Then
'        If Not IsNull(vgRs!valor) Then vgValorMoneda = vgRs!valor
'        If (vgValorMoneda = 0) Then vgError = 5000
'    Else
'        vgError = 5000
'    End If
'    vgRs.Close
'
''''    Call CerrarBaseDeDatos(vgConectarBD)
'
'Exit Function
'Err_BuscarMoneda:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Function

'Function fgBuscarComisionIntermediario() As Double
'On Error GoTo Err_Combo
'    'Call AbrirBaseDeDatos_Aux(vgRutaBasedeDatos)
''''    If Not AbrirBaseDeDatos(vgConectarBD) Then
''''        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
''''        Exit Function
''''    End If
'
'    'vlCombo.Clear
'    vgQuery = "SELECT MTO_ELEMENTO from PT_TPAR_TABCOD WHERE "
'    vgQuery = vgQuery & "cod_tabla = 'CI' and "
'    vgQuery = vgQuery & "cod_elemento = 'I' "
'    vgQuery = vgQuery & "ORDER BY cod_elemento "
'    Set vgRs = vgConexionBD.Execute(vgQuery)
'    If Not (vgRs.EOF) Then
'        fgBuscarComisionIntermediario = vgRs.Fields("MTO_ELEMENTO")
'    Else
'        fgBuscarComisionIntermediario = 0
'    End If
'    vgRs.Close
'
''''    Call CerrarBaseDeDatos(vgConectarBD)
'Exit Function
'Err_Combo:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'
'End Function

'------------------------------------------------------
' Función de Inicio del Sistema
'------------------------------------------------------
Public Sub p_Actualiza_Version(ByVal g_strPathExe As String, ByVal sExeUpgade As String)
    
    Dim FechaAppS As Date
    Dim FechaAppM As Date
    Dim r_FlagOk As Variant
    'Hora y fecha del exe en el servidor
    FechaAppS = Format$(FileDateTime(g_strPathExe & "Produccion.exe"), "dd/mm/yyyy hh:mm")
    'Hora y fecha del exe local
    'FechaAppM = Format$(FileDateTime(g_strPathExe & "\Produccion.exe"), "dd/mm/yyyy hh:mm")
    FechaAppM = Format$(FileDateTime(App.Path & "\Produccion.exe"), "dd/mm/yyyy hh:mm")
        
    If FechaAppS <> FechaAppM Then
        If MsgBox("La versión del sistema en el servidor, difiere a la del computador. Desea Actualizar la versión ?", vbQuestion + vbYesNo, "") = vbYes Then
           r_FlagOk = Shell(sExeUpgade & " RutaO=" & g_strPathExe & "Produccion.exe" & ",RutaD=" & App.Path & ",NomEx=Produccion.exe", vbNormalFocus)
           End
        Else
           End
        End If
    End If
        
End Sub

Sub Main()
Dim inises As String
Dim usua As String
Dim pas As String

Dim strExes As String
Dim strUpgrade As String

On Error GoTo Err_Main

    inises = "sincn"
    Do Until inises = "concn"
        usua = LeeArchivoIni("Conexion", "Usuario", "", App.Path & "\AdmPrevBD.Ini")
        pas = LeeArchivoIni("Conexion", "Password", "", App.Path & "\AdmPrevBD.Ini")
        strRpt = LeeArchivoIni("REPORTES", "strPath", "", App.Path & "\Rutas.ini")
        strExes = LeeArchivoIni("EXES", "strPath", "", App.Path & "\Rutas.ini")
        strUpgrade = LeeArchivoIni("EXES", "strUpgrade", "", App.Path & "\Rutas.ini")
        
        If (usua = "" Or usua = Empty) Or (pas = "" Or pas = Empty) Then
            'Frm_InicioSesion.Show 1
            If inises = "cancelar" Then End
        Else
            inises = "concn"
        End If
    Loop
    
    If strExes <> "" Then
        'Call p_Actualiza_Version(strExes, strUpgrade)
    End If
    
    vgDsn = ""
    vgNombreServidor = ""
    vgNombreBaseDatos = ""
    vgNombreUsuario = ""
    vgPassWord = ""
    vgMensaje = ""
    vgRutaArchivo = ""
    
    'Valida Si Existe Archivo de AdmBasDat.Inicio
    'vgRutaArchivo = App.Path & "\AdmPrevBD.Ini"
    vgRutaArchivo = App.Path & "\AdmPrevBD_Prod.Ini"
    If Not fgExiste(vgRutaArchivo) Then
        MsgBox "No existe el Archivo de Parámetros para ejecutar la Aplicación.", vbCritical, "Ejecución Cancelada"
        End
    End If

    lpFileName = vgRutaArchivo
    lpAppName = "Conexion"
    lpDefault = ""
    lpReturnString = Space$(128)
    Size = Len(lpReturnString)
    lpKeyName = ""
    
    'Valida Si Existe Nombre de la entrada para definir al Proveedor
    ProviderName = fgGetPrivateIni(lpAppName, "Proveedor", lpFileName)
    If (ProviderName = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'Proveedor', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    End If
    
    'Valida Si Existe Nombre de la entrada para definir al Servidor
    vgNombreServidor = fgGetPrivateIni(lpAppName, "Servidor", lpFileName)
    If (vgNombreServidor = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'Servidor', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    End If

    'Valida Si Existe Nombre de la entrada para definir Base de Datos SisSin
    vgNombreBaseDatos = fgGetPrivateIni(lpAppName, "BaseDatos", lpFileName)
    If (vgNombreBaseDatos = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'Base de Datos', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    End If
    
    'Valida Si Existe Nombre de la entrada para definir Usuario de SisSin
    vgNombreUsuario = fgGetPrivateIni(lpAppName, "Usuario", lpFileName)
    If (vgNombreUsuario = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'Usuario', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    End If
    vgNombreUsuario = "prorv_user" ' fgDesPassword(vgNombreUsuario)
    
    'Valida Si Existe Nombre de la entrada para definir Password de SisSin
    vgPassWord = fgGetPrivateIni(lpAppName, "Password", lpFileName)
    'vgPassWord = ""
    If (vgPassWord = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'PassWord', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    Else
        If (vgPassWord = "DESCONOCIDO") And (UCase(vgNombreUsuario) = "SA") Then
            vgPassWord = ""
        End If
    End If
    'vgPassWord = fgDesPassword(vgPassWord)
    'vgPassWord = "rentcalidad64"
    vgPassWord = "rentcalidad64" '"protecta"
    
    'Valida Si Existe Nombre de la entrada para definir DSN de SisSin
    vgDsn = fgGetPrivateIni(lpAppName, "DSN", lpFileName)
    If (vgDsn = "DESCONOCIDO") Then
        vgMensaje = "La Entrada 'DSN', no está definida en el Archivo AdmPrevBD.Ini" & vbCrLf
    End If

    If (vgMensaje <> "") Then
        MsgBox "Status de los Datos de Inicio" & vbCrLf & vbCrLf & vgMensaje & vbCrLf & vbCrLf & "Proceso Cancelado." & vbCrLf & "Se deben Ingresar todos los datos Básicos."
        'Exit Sub
        End
    End If

    'vgRutaBasedeDatos = LeeArchivoIni("Conexion", "Ruta", "", App.Path & "\AdmPrevBD.Ini")
    'vgRutaBasedeDatos = vgRutaBasedeDatos & LeeArchivoIni("Conexion", "BasedeDatos", "", App.Path & "\AdmPrevBD.Ini")
    ''AbrirBaseDeDatos (vgRutaBasedeDatos)

    If Not fgConexionBaseDatos(vgConexionBD) Then
        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
        'Exit Sub
        End
    End If
    'vgConBD.Close
    
    'Call CerrarBaseDeDatos

    'Ruta de la Base de Datos
    vgRutaDataBase = "ODBC;DSN=" & vgDsn & ";UID=" & vgNombreUsuario & ";PWD=" & vgPassWord & ";DATABASE=" & vgNombreBaseDatos & ";"
    vgRutaBasedeDatos = vgRutaDataBase

'    If Not fgConexionBaseDatos(vgConexionBD) Then
'        MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
'        End
'    End If
    
    'If Not fgConexionBaseDatos(vgConectarBD) Then
    '    MsgBox "Falló la conexión con la Base de Datos.", vbCritical, "Error de Conexión"
    '    End
    'End If
    
    '*******************************************************
    'Sacar luego cuando sean ingresados los datos que faltan
    '*******************************************************
    'Determinar los Datos del Cliente
    vgQuery = "select * from ma_tmae_cliente "
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!gls_nomcli) Then vgNombreCompania = Trim(vgRs!gls_nomcli)
        If Not IsNull(vgRs!gls_nomcorcli) Then vgNombreCortoCompania = Trim(vgRs!gls_nomcorcli)
        If Not IsNull(vgRs!num_idencli) Then vgNumIdenCliente = Trim(vgRs!num_idencli)
        If Not IsNull(vgRs!cod_interno) Then vgCodigoSBSCompania = Trim(vgRs!cod_interno)
        If Not IsNull(vgRs!cod_sucave) Then vgCodigoSucaveCompañia = Trim(vgRs!cod_sucave)
        If Not IsNull(vgRs!cod_tipoidencli) Then vgTipoIdenCompania = Trim(vgRs!cod_tipoidencli)
        If Not IsNull(vgRs!num_idencli) Then vgNumIdenCompania = Trim(vgRs!num_idencli)
    Else
        vgNombreCompania = ""
        vgNombreCortoCompania = ""
        vgNumIdenCliente = ""
        vgTipoIdenCompania = ""
        vgNumIdenCompania = ""
        vgCodigoSucaveCompañia = ""
    End If
    vgRs.Close
        
    'Determinar los Datos del Sistema
    vgQuery = "select * from ma_tpar_sistema where "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "'"
    Set vgRs = vgConexionBD.Execute(vgQuery)
    If Not vgRs.EOF Then
        If Not IsNull(vgRs!gls_sistema) Then vgNombreSubSistema = Trim(vgRs!gls_sistema)
    Else
        vgNombreSubSistema = ""
    End If
    vgRs.Close
    '*******************************************************

    
    'RRR
        vgSql = "SELECT * FROM MA_TMAE_ADMINCUENTAS WHERE "
        vgSql = vgSql & "cod_cliente = '1' "
        Set vgRs = vgConexionBD.Execute(vgSql)

                    If Not vgRs.EOF Then
                        vgIntentos = vgRs!nintentos
                        vgChkdiaant = vgRs!bdiasvence
                        vgDiasFaltan = vgRs!ndiasvence
                    End If
        vgRs.Close
    'RRR

    'Cerrar la Conexión
    vgConexionBD.Close
    
    'Ruta de la Base de Datos
    'vgRutaDataBase = "ODBC;DSN=" & vgDsn & ";UID=" & vgNombreUsuario & ";PWD=" & vgPassWord & ";DATABASE=" & vgNombreBaseDatos & ";"
    'vgRutaDataBase = "ODBC;DSN=" & vgDsn & ";DATABASE=" & vgNombreBaseDatos & ";"
    'vgRutaDataBase = "ODBC;UID="";PWD="";DATABASE=" & vgNombreBaseDatos & ";DSN=" & vgDsn & ";"
    'vgRutaBasedeDatos = vgRutaDataBase

    '****************************
    'Inicio: Sacar Comentarios
    '****************************
    Screen.MousePointer = 11
    Frm_Menu.Show
    'I----- Debe validarse ABV 21/01/2004 ---
    'Activarlo cuando se encuentre el Menú corregido y activado
    Frm_Menu.Mnu_AdmSistema.Enabled = False
    Frm_Menu.Mnu_AdmParametros.Enabled = False
    Frm_Menu.Mnu_ProcGeneracion.Enabled = False
    Frm_Menu.Mnu_ProcConsulta.Enabled = False 'hqr 15/07/2007
    'Frm_Menu.Mnu_Notebook.Enabled = False
    Frm_Menu.Mnu_Acerca.Enabled = False
    Frm_Menu.Mnu_Salir.Enabled = False
    'I----- Debe validarse ABV 21/01/2004 ---
    Screen.MousePointer = 0
    
    Screen.MousePointer = 11
    Frm_Password.Show
    Screen.MousePointer = 0
    '****************************
    'Fin  : Sacar Comentarios
    '****************************
    fgCargarVariablesGlobales
Exit Sub
Err_Main:
    Screen.MousePointer = 0
    Select Case Err
    Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
        End
    End Select
End Sub

'********************************************************
'NUEVAS FUNCIONES CREADAS
'********************************************************
'**************A G R E G A D O S ********************

'------------------------------------------------------------
'Permite Cargar el Combo de Sucursales del Sistema
'------------------------------------------------------------
Function fgComboSucursal(icombo As ComboBox, iTipo As String)
Dim vlRsCombo As ADODB.Recordset
On Error GoTo Err_ComboSucursal

    icombo.Clear
    vgSql = "SELECT cod_sucursal,gls_sucursal "
    vgSql = vgSql & "FROM MA_TPAR_SUCURSAL "
    vgSql = vgSql & "WHERE cod_tipo = '" & iTipo & "' "
    vgSql = vgSql & "ORDER BY cod_sucursal "
    Set vlRsCombo = vgConexionBD.Execute(vgSql)
    Do While Not vlRsCombo.EOF
        icombo.AddItem ((Trim(vlRsCombo!Cod_Sucursal) & " - " & Trim(vlRsCombo!gls_sucursal)))
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

Function fgComboNivel(vlCombo As ComboBox)
On Error GoTo Err_Combo

    vlCombo.Clear
    vgQuery = "SELECT cod_nivel as codigo from "
    vgQuery = vgQuery & "MA_TPAR_NIVEL WHERE "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' "
    vgQuery = vgQuery & "ORDER BY cod_nivel "
    Set vgCmb = vgConexionBD.Execute(vgQuery)
    If Not (vgCmb.EOF) Then
        'vgCmb.MoveFirst
        While Not (vgCmb.EOF)
            vlCombo.AddItem (Trim(vgCmb!Codigo))
            vgCmb.MoveNext
        Wend
        If (vlCombo.ListCount <> 0) Then
            vlCombo.ListIndex = 0
        End If
    End If
    vgCmb.Close
    
Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function ValiRut(Rut As String, DgV As String) As Boolean
Dim vlDigitos As String
Dim Resultado, y As Integer
    
    Rut = Format(Rut, "#0")
    vlDigitos = "32765432"
    Resultado = 0
    Select Case Len(Rut)
           Case 7: Rut = "0" + Rut
           Case 6: Rut = "00" + Rut
           Case 5: Rut = "000" + Rut
           Case 4: Rut = "0000" + Rut
           Case 3: Rut = "00000" + Rut
           Case 2: Rut = "000000" + Rut
           Case 1: Rut = "0000000" + Rut
    End Select
    For y = 1 To Len(Trim(Rut))
        Resultado = Resultado + Val(Mid(Rut, Len(Trim(Rut)) - y + 1, 1)) * Val(Mid(vlDigitos, Len(Trim(Rut)) - y + 1, 1))
    Next y
    vlDiv = Int(Resultado / 11)
    vlMul = vlDiv * 11
    vlDiv = Resultado - vlMul
    vlDgv = 11 - vlDiv
    If vlDgv = 11 Then
       vlDgv = 0
    End If
    If vlDgv = 10 Then
       vlDgv = "K"
    End If
    If vlDgv = DgV Then
       ValiRut = True
    Else
       ValiRut = False
    End If
End Function

'Sub Carga_Arreglo_Tasas()
'
''vlSql = "SELECT * FROM PT_TVAL_CALCE WHERE "
''vlSql = vlSql & "cod_moneda = '" & vgMonedaOficial & "'"
''vlSql = vlSql & "order by num_anno "
''Set vgRs = vgConexionBD.Execute(vlSql)
''fila = 1
''While Not vgRs.EOF
''    Arr_Tasas(fila) = vgRs!prc_cpk
''    vgRs.MoveNext
''    fila = fila + 1
''Wend
''vgRs.Close
'
'    vlSql = "SELECT * FROM PT_TVAL_RENTABILIDAD WHERE "
'    vlSql = vlSql & "cod_moneda = '" & vgMonedaOficial & "'"
'    vlSql = vlSql & "order by num_anno "
'    Set vgRs = vgConexionBD.Execute(vlSql)
'    fila = 1
'    While Not vgRs.EOF
'        Arr_Tasas(fila) = vgRs!prc_tasatip
'        vgRs.MoveNext
'        fila = fila + 1
'    Wend
'    vgRs.Close
'End Sub

Function fgObtieneConversion(iFecha, iMoneda, oValor) As Boolean
Dim vlFechaAnterior As String
Dim vlFechaMaxima As String
Dim vlFeriadosMaximo As Long

    'Obtiene Valor de la Moneda a la Fecha ingresada como Parámetro
    Dim vlSql As String
    Dim vlTb As ADODB.Recordset
    
    fgObtieneConversion = False
    vlFeriadosMaximo = 4
    
    If (iMoneda <> vgMonedaCodOfi) Then
        
        vlSql = "SELECT MTO_MONEDA FROM MA_TVAL_MONEDA "
        vlSql = vlSql & " WHERE COD_MONEDA = '" & iMoneda & "' "
        vlSql = vlSql & " AND FEC_MONEDA = '" & iFecha & "' "
        Set vlTb = vgConexionBD.Execute(vlSql)
        If Not vlTb.EOF Then
            oValor = vlTb!mto_moneda
            fgObtieneConversion = True
        End If
        vlTb.Close
        
        If (fgObtieneConversion = False) Then
            vlFechaAnterior = ""
            vlFechaMaxima = Format(DateSerial(Mid(iFecha, 1, 4), Mid(iFecha, 5, 2), CInt(Mid(iFecha, 7, 2)) - vlFeriadosMaximo), "yyyymmdd")
            
            'Debe obtener el primero que encuentre anterior a la fecha consultada
            vlSql = "SELECT FEC_MONEDA, MTO_MONEDA FROM MA_TVAL_MONEDA "
            vlSql = vlSql & " WHERE COD_MONEDA = '" & iMoneda & "' "
            vlSql = vlSql & " AND FEC_MONEDA < '" & iFecha & "' "
            vlSql = vlSql & " ORDER BY FEC_MONEDA DESC "
            Set vlTb = vgConexionBD.Execute(vlSql)
            If Not vlTb.EOF Then
                vlFechaAnterior = vlTb!FEC_MONEDA
                If (vlFechaAnterior > vlFechaMaxima) Then
                    oValor = vlTb!mto_moneda
                    fgObtieneConversion = True
                End If
            End If
            vlTb.Close
        End If
        
    Else
        fgObtieneConversion = True
        oValor = 1
    End If

End Function

'Function fgBuscarGlosaElemento(iTabla, iElemento) As String
'Dim vlRsDescripcion As ADODB.Recordset
'On Error GoTo Err_BuscarGlosa
'
'    fgBuscarGlosaElemento = ""
'
'    vgSql = ""
'    vgSql = "select gls_elemento "
'    vgSql = vgSql & "from MA_TPAR_TABCOD where "
'    vgSql = vgSql & "cod_tabla = '" & iTabla & "' and "
'    vgSql = vgSql & "cod_elemento = '" & iElemento & "' "
'    Set vlRsDescripcion = vgConexionBD.Execute(vgSql)
'    If Not vlRsDescripcion.EOF Then
'        fgBuscarGlosaElemento = vlRsDescripcion!gls_elemento
'    End If
'    vlRsDescripcion.Close
'
'Exit Function
'Err_BuscarGlosa:
'    Screen.MousePointer = 0
'    Select Case Err
'        Case Else
'        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
'    End Select
'End Function

Function fgBuscaFecServ() As String
    
    If vgTipoBase = "ORACLE" Then
       vgSql = ""
       vgSql = "SELECT SYSDATE AS FEC_ACTUAL FROM MA_TCOD_GENERAL"
       Set vgRs4 = vgConexionBD.Execute(vgSql)
       If Not vgRs4.EOF Then
          fgBuscaFecServ = Mid((vgRs4!FEC_ACTUAL), 1, 10)
       End If
    Else
      If vgTipoBase = "SQL" Then
         vgSql = ""
         vgSql = "SELECT GETDATE()AS FEC_ACTUAL FROM MA_TCOD_GENERAL"
         Set vgRs4 = vgConexionBD.Execute(vgSql)
         If Not vgRs4.EOF Then
            fgBuscaFecServ = Mid((vgRs4!FEC_ACTUAL), 1, 10)
         End If
      End If
    End If
    
End Function

'------------------------------------------------------------
'Permite Obtener la Glosa de un Código especifico de la Tabla
'de Parámetros Generales (TABCOD)
'iTabla = código de la Tabla, iElemento = código de Elemento
'------------------------------------------------------------
Function fgBuscarGlosaElemento(iTabla, iElemento) As String
Dim vlRsDescripcion As ADODB.Recordset
On Error GoTo Err_BuscarGlosa

    fgBuscarGlosaElemento = ""
    
    vgSql = ""
    vgSql = "SELECT gls_elemento "
    vgSql = vgSql & "FROM ma_tpar_tabcod WHERE "
    vgSql = vgSql & "cod_tabla = '" & iTabla & "' AND "
    vgSql = vgSql & "cod_elemento = '" & iElemento & "' "
    Set vlRsDescripcion = vgConexionBD.Execute(vgSql)
    If Not vlRsDescripcion.EOF Then
        fgBuscarGlosaElemento = vlRsDescripcion!gls_elemento
    End If
    vlRsDescripcion.Close
    
Exit Function
Err_BuscarGlosa:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgBuscarGlosaCauInv(iElemento) As String

On Error GoTo Err_fgBuscarGlosaCauInv

    fgBuscarGlosaCauInv = ""
    
    vgSql = ""
    vgSql = "SELECT gls_patologia "
    vgSql = vgSql & "FROM ma_tpar_patologia WHERE "
    vgSql = vgSql & "cod_patologia = '" & iElemento & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        fgBuscarGlosaCauInv = (vgRegistro!gls_patologia)
    End If
    vgRegistro.Close
    
Exit Function
Err_fgBuscarGlosaCauInv:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgBuscarNombreTipoIden(iCodTipoIden As String, Optional iNombreLargo As Boolean) As String
Dim vlRegNombre As ADODB.Recordset
On Error GoTo Err_BuscarNombreCorto
    
    fgBuscarNombreTipoIden = ""
    
    If (iNombreLargo = True) Then
        vlSql = "SELECT gls_tipoiden as gls_tipoidencor "
    Else
        vlSql = "SELECT gls_tipoidencor "
    End If
    vlSql = vlSql & "FROM ma_tpar_tipoiden "
    vlSql = vlSql & "WHERE "
    vlSql = vlSql & "cod_tipoiden = " & iCodTipoIden & " "
    Set vlRegNombre = vgConexionBD.Execute(vlSql)
    If Not vlRegNombre.EOF Then
         If Not IsNull(vlRegNombre!GLS_TIPOIDENCOR) Then fgBuscarNombreTipoIden = vlRegNombre!GLS_TIPOIDENCOR
    End If
    vlRegNombre.Close
  
Exit Function
Err_BuscarNombreCorto:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgBuscarNroLiquidacion(iPoliza As String) As String
Dim vlRsNroLiquidacion As ADODB.Recordset
On Error GoTo Err_BuscarNroLiquidacion

    fgBuscarNroLiquidacion = ""
    
    vgSql = "SELECT cod_liquidacion "
    vgSql = vgSql & "FROM pd_tmae_polprirec WHERE "
    vgSql = vgSql & "num_poliza = '" & iPoliza & "'"
    Set vlRsNroLiquidacion = vgConexionBD.Execute(vgSql)
    If Not vlRsNroLiquidacion.EOF Then
        If Not IsNull(vlRsNroLiquidacion!cod_liquidacion) Then
            fgBuscarNroLiquidacion = vlRsNroLiquidacion!cod_liquidacion
        End If
    End If
    vlRsNroLiquidacion.Close
    
Exit Function
Err_BuscarNroLiquidacion:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
Function fgComboNivelGlosa(vlCombo As ComboBox)
Dim vlGlosa As String
On Error GoTo Err_Combo

    vlCombo.Clear
    vgQuery = "SELECT cod_nivel as codigo "
    vgQuery = vgQuery & ",gls_nivel as glosa "
    vgQuery = vgQuery & "from "
    vgQuery = vgQuery & "MA_TPAR_NIVEL WHERE "
    vgQuery = vgQuery & "cod_sistema = '" & vgTipoSistema & "' "
    vgQuery = vgQuery & "ORDER BY cod_nivel "
    Set vgCmb = vgConexionBD.Execute(vgQuery)
    If Not (vgCmb.EOF) Then
        'vgCmb.MoveFirst
        While Not (vgCmb.EOF)
            vlGlosa = ""
            If Not IsNull(vgCmb!glosa) Then vlGlosa = Trim(vgCmb!glosa)
            vlCombo.AddItem (Trim(vgCmb!Codigo) & " - " & vlGlosa)
            vgCmb.MoveNext
        Wend
        If (vlCombo.ListCount <> 0) Then
            vlCombo.ListIndex = 0
        End If
    End If
    vgCmb.Close
    
Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgComboTipoIdentificacion(vlCombo As ComboBox)
Dim vlRegCombo As ADODB.Recordset
On Error GoTo Err_Combo

    vlCombo.Clear
    vgSql = ""
    vlCont = 0
    vgQuery = "SELECT cod_tipoiden as codigo, gls_tipoidencor as Nombre, "
    vgQuery = vgQuery & "num_lartipoiden as largo "
    vgQuery = vgQuery & "FROM MA_TPAR_TIPOIDEN "
    vgQuery = vgQuery & "ORDER BY codigo "
    Set vlRegCombo = vgConexionBD.Execute(vgQuery)
    If Not (vlRegCombo.EOF) Then
        While Not (vlRegCombo.EOF)
            vlCombo.AddItem Space(2 - Len(vlRegCombo!Codigo)) & vlRegCombo!Codigo & " - " & (Trim(vlRegCombo!nombre))
            vlCont = vlCombo.ListCount - 1
            vlCombo.ItemData(vlCont) = (vlRegCombo!largo)
            vlRegCombo.MoveNext
        Wend
    End If
    vlRegCombo.Close
        
    If vlCombo.ListCount <> 0 Then
        vlCombo.ListIndex = 0
    End If

Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'--Begin Ricardo Huerta Dextre : 10/01/2018
Function fgComboNacionalidad(vlCombo As ComboBox)
Dim vlRegCombo As ADODB.Recordset
On Error GoTo Err_Combo

    vlCombo.Clear
    vgSql = ""
    vlCont = 0
    
    vgQuery = " select '0' as cod_nacionalidad, '---Seleccione---' as gls_nacionalidad from dual "
    vgQuery = vgQuery & " Union"
    vgQuery = vgQuery & " select cod_nacionalidad, gls_nacionalidad from MA_TPAR_NACIONALIDAD  order by gls_nacionalidad "
       
    Set vlRegCombo = vgConexionBD.Execute(vgQuery)
    If Not (vlRegCombo.EOF) Then
        While Not (vlRegCombo.EOF)
            vlCombo.AddItem ((Trim(vlRegCombo!cod_nacionalidad) & " - " & Trim(vlRegCombo!gls_nacionalidad)))
            vlCont = vlCombo.ListCount - 1
            vlCombo.ItemData(vlCombo.NewIndex) = Val(vlRegCombo!cod_nacionalidad)
            vlRegCombo.MoveNext
        Wend
    End If
    vlRegCombo.Close
    
    Dim vgI As Integer
    vgI = 0
    vlCombo.ListIndex = 0
   
    If vlCombo.ListCount <> 0 Then
    
        Do While vgI < vlCombo.ListCount
            
            If InStr(1, Trim(UCase(vlCombo.Text)), "PERU") > 0 Then
                       Exit Do
            End If
          vgI = vgI + 1
          vlCombo.ListIndex = vgI
          
      Loop
   End If

Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fg_obtener_descripcion_nacionalidad(pcod_nacionalidad As String) As String
Dim vlReg As ADODB.Recordset
On Error GoTo Err_obtener_descripcion_nacionalidad

    vgQuery = "select cod_nacionalidad, gls_nacionalidad from MA_TPAR_NACIONALIDAD where cod_nacionalidad = " & pcod_nacionalidad
       
    Set vlReg = vgConexionBD.Execute(vgQuery)
    If Not (vlReg.EOF) Then
            fg_obtener_descripcion_nacionalidad = vlReg!gls_nacionalidad
    End If
    vlReg.Close
  
Exit Function
Err_obtener_descripcion_nacionalidad:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function
'--End Ricardo Huerta Dextre : 10/01/2018

'------------------------------------------------------------
'Permite Obtener la Posición, dentro de un Combo, del Código
'del Elemento indicado
'iElemento = código de Elemento
'------------------------------------------------------------
Function fgBuscarPosicionCodigoCombo(iElemento, icombo As ComboBox) As Long
Dim iContador As Long
On Error GoTo Err_BuscarPosicion

    fgBuscarPosicionCodigoCombo = -1
    
    iContador = 0
    icombo.ListIndex = 0
    Do While iContador < icombo.ListCount
        If (Trim(icombo) <> "") Then
            If (Trim(iElemento) = Trim(Mid(icombo.Text, 1, (InStr(1, icombo, "-") - 1)))) Then
                fgBuscarPosicionCodigoCombo = iContador
                Exit Do
            End If
        End If
        
        iContador = iContador + 1
        'If (vgI = iCombo.ListCount) Then
        '    MsgBox "Código no identificado o no existente.", vbExclamation, "Dato Inexistente"
        '    iCombo.ListIndex = -1
        '    Exit Do
        'End If
        icombo.ListIndex = iContador
    Loop

Exit Function
Err_BuscarPosicion:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgComboSucursalGlosa(vlCombo As ComboBox)
Dim vlGlosa As String
On Error GoTo Err_Combo

    vlCombo.Clear
    vgQuery = "SELECT cod_sucursal as codigo "
    vgQuery = vgQuery & ",gls_sucursal as glosa "
    vgQuery = vgQuery & "from "
    vgQuery = vgQuery & "pd_tpar_sucursal "
    vgQuery = vgQuery & "ORDER BY cod_sucursal "
    Set vgCmb = vgConexionBD.Execute(vgQuery)
    If Not (vgCmb.EOF) Then
        'vgCmb.MoveFirst
        While Not (vgCmb.EOF)
            vlGlosa = ""
            If Not IsNull(vgCmb!glosa) Then vlGlosa = Trim(vgCmb!glosa)
            vlCombo.AddItem (Trim(vgCmb!Codigo) & " - " & vlGlosa)
            vgCmb.MoveNext
        Wend
        If (vlCombo.ListCount <> 0) Then
            vlCombo.ListIndex = 0
        End If
    End If
    vgCmb.Close
    
Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Public Function fgCargarTablaMoneda(icodtabla As String, oEstructura() As TypeTablaMoneda, oNumTotal As Long)
'Función : Llenar la Estructura de los Tipos de Moneda que se encuentran registrados en la BD
'Parámetros de Entrada:
'Parámetros de Salida:
'- Llenar la Estructura de Tipos de Monedas
'------------------------------------------------------
'Fecha de Creación     : 05/07/2007 - ABV
'Fecha de Modificación :
'------------------------------------------------------
Dim iRegistro As ADODB.Recordset
Dim iSql As String
On Error GoTo Err_Tabla

    oNumTotal = 0
    
    'Selecciona el Número Máximo de los Códigos de Parámetros
    vgQuery = "SELECT count(cod_elemento) as numero "
    vgQuery = vgQuery & "from ma_tpar_tabcod "
    vgQuery = vgQuery & "WHERE cod_tabla = '" & icodtabla & "' "
    vgQuery = vgQuery & "AND (cod_sistema <> 'PP' OR cod_sistema is null) "
    Set iRegistro = vgConexionBD.Execute(vgQuery)
    If Not (iRegistro.EOF) Then
        If Not IsNull(iRegistro!Numero) Then
            oNumTotal = iRegistro!Numero
        End If
    End If
    iRegistro.Close
    
    'Llena la Estructura con los códigos de Parámetros
    If (oNumTotal <> 0) Then
        ReDim oEstructura(oNumTotal) As TypeTablaMoneda
        
        iSql = "SELECT cod_elemento as codigo,gls_elemento as descripcion,"
        iSql = iSql & "cod_scomp as cod_asociado "
        iSql = iSql & "FROM ma_tpar_tabcod "
        iSql = iSql & "WHERE cod_tabla = '" & icodtabla & "' "
        iSql = iSql & "AND (cod_sistema <> 'PP' OR cod_sistema is null) "
        iSql = iSql & "ORDER BY cod_elemento "
        Set iRegistro = vgConexionBD.Execute(iSql)
        If Not (iRegistro.EOF) Then
            vgX = 1
            While Not (iRegistro.EOF)
                oEstructura(vgX).Codigo = iRegistro!Codigo
                oEstructura(vgX).Descripcion = IIf(IsNull(iRegistro!Descripcion), "", Trim(iRegistro!Descripcion))
                oEstructura(vgX).Scomp = IIf(IsNull(iRegistro!Cod_Asociado), "", iRegistro!Cod_Asociado)
                
                iRegistro.MoveNext
                vgX = vgX + 1
            Wend
        End If
        iRegistro.Close
    End If
    
Exit Function
Err_Tabla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Public Function fgObtenerCodMonedaScomp(iEstructura() As TypeTablaMoneda, iNumTotal As Long, iTipoCodigo As String) As String
'Función: Permite obtener el Código de Scomp de una Moneda específica
'Parámetros de Entrada :
'- iEstructura => Estructura que contiene los Tipos de Moneda
'- iNumTotal   => Número Total de Filas de la Estructura
'- iTipoCodigo => Código de la Moneda a buscar
'Parámetros de Salida :
'- Devuelve el código de la Moneda definida para SCOMP
    
    fgObtenerCodMonedaScomp = ""
    
    If (iNumTotal <> 0) Then
        For vgX = 1 To iNumTotal
            If (iEstructura(vgX).Codigo = iTipoCodigo) Then
                fgObtenerCodMonedaScomp = iEstructura(vgX).Scomp
                Exit For
            End If
        Next vgX
    End If
    
End Function

Function fgComboMoneda(vlCombo As ComboBox)
On Error GoTo Err_Combo
    
    vlCombo.Clear
    vgQuery = "SELECT cod_elemento as codigo , gls_elemento as Nombre from "
    vgQuery = vgQuery & "MA_TPAR_TABCOD WHERE "
    vgQuery = vgQuery & "cod_tabla = 'TM' "
'I--- ABV 25/06/2007 ---
    vgQuery = vgQuery & "AND (cod_sistema <> 'PP' OR cod_sistema is null) "
'F--- ABV 25/06/2007 ---
    vgQuery = vgQuery & "ORDER BY cod_elemento " 'desc"
    Set vgCmb = vgConexionBD.Execute(vgQuery)
    If Not (vgCmb.EOF) Then
        
        While Not (vgCmb.EOF)
            vgPalabra = Trim(vgCmb!Codigo)
            vlCombo.AddItem (vgPalabra & _
            " - " & Trim(vgCmb!nombre))
            vgCmb.MoveNext
        Wend
        If (vlCombo.ListCount <> 0) Then
            vlCombo.ListIndex = 0
        End If
    End If
    vgCmb.Close
    
Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Public Function fgBuscarMonedaOfiTran(oMonedaOfi As String, oMonedaTran As String)
'Función: Permite buscar el Código de la Moneda a expresar los valores, y
'el Código de la Moneda en que se deben transformar los valores
'Parámetros de Entrada:
'Parámetros de Salida:
'- oMonedaOfi   => Código de la Moneda Oficial
'- oMonedaTran  => Código de la Moneda a Transformar
'----------------------------------------------------
'Fecha Creación     : 07/07/2007
'Fecha Modificación :
'----------------------------------------------------
Dim vlRegistroMon As ADODB.Recordset
On Error GoTo Err_BuscarMoneda

'    fgBuscarMonedaOfiTran = False
    oMonedaOfi = cgCodTipMonedaUF
    oMonedaTran = cgCodTipMonedaUF
    
    vgSql = "SELECT cod_monedaofi,cod_monedatrans "
    vgSql = vgSql & "FROM ma_tcod_moneda "
    Set vlRegistroMon = vgConexionBD.Execute(vgSql)
    If Not vlRegistroMon.EOF Then
        If Not IsNull(vlRegistroMon!cod_monedaofi) Then oMonedaOfi = vlRegistroMon!cod_monedaofi
        If Not IsNull(vlRegistroMon!cod_monedatrans) Then oMonedaTran = vlRegistroMon!cod_monedatrans
'        fgBuscarMonedaOfiTran = True
    End If
    vlRegistroMon.Close
    
Exit Function
Err_BuscarMoneda:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

'Public Function fgCalcularFechaPrimerPagoEst(iFecAcep As String, iFecDev As String, iTipoPension As String, iTipoRenta As String, iMesDif As Long) As String
''Permite estimar la Fecha de Primer Pago de la Póliza
''Parámetros de Entrada:
''- iFecAcep As String       => Fecha de Aceptación o Incorporación
''- iFecDev As String        => Fecha de Devengue de la Pensión
''- iTipoPension As String   => Tipo de Pensión
''- iTipoRenta As String     => Tipo de Renta
''- iMesDif                  => Número de Meses Diferidos
''Parámetros de Salida:
''------------------------------------------------------
''Fecha de Creación     : 07/07/2007
''Fecha de Modificación :
''------------------------------------------------------
'
'Dim vlRegFecha As ADODB.Recordset
'Dim vlDias As Long
'Dim vlFechaPago As String
'
'    fgCalcularFechaPrimerPagoEst = ""
'    vlDias = 0
'
'    vgQuery = "SELECT mto_elemento as dias FROM "
'    vgQuery = vgQuery & "MA_TPAR_TABCOD WHERE "
'    vgQuery = vgQuery & "cod_tabla = '" & vgCodTabla_TipPen & "' AND "
'    vgQuery = vgQuery & "cod_elemento = '" & iTipoPension & "' "
'    Set vlRegFecha = vgConexionBD.Execute(vgQuery)
'    If Not (vlRegFecha.EOF) Then
'        If Not IsNull(vlRegFecha!dias) Then
'            vlDias = vlRegFecha!dias
'        End If
'    End If
'    vlRegFecha.Close
'
'    If (iTipoRenta = cgTipoRentaInmediata) Then
'        vlFechaPago = DateSerial(Mid(iFecAcep, 1, 4), Mid(iFecAcep, 5, 2) + vlDias, Mid(iFecAcep, 7, 2))
'    Else
'        vlFechaPago = DateSerial(Mid(iFecDev, 1, 4), Mid(iFecDev, 5, 2) + iMesDif, 1)
'    End If
'
'    fgCalcularFechaPrimerPagoEst = vlFechaPago
'
'End Function

Public Function fgCalcularFechaPrimerPagoEst(iFecAcep As String, iFecDev As String, iFecRecPrima As String, iTipoPension As String, iTipoRenta As String, iMesDif As Long, iNumDias As Long) As String
'Permite estimar la Fecha de Primer Pago de la Póliza
'Parámetros de Entrada:
'- iFecAcep As String       => Fecha de Aceptación o Incorporación
'- iFecDev As String        => Fecha de Devengue de la Pensión
'- iTipoPension As String   => Tipo de Pensión
'- iTipoRenta As String     => Tipo de Renta
'- iMesDif                  => Número de Meses Diferidos
'- iNumDias                 => Número de Días de Plazo para el Primer Pago
'Parámetros de Salida:
'- Retorno Funcion          => Fecha Tope Primer Pago en Formato dd/mm/yyyy o formato del equipo
'------------------------------------------------------
'Fecha de Creación     : 07/07/2007
'Fecha de Modificación :
'------------------------------------------------------

Dim vlRegFecha As ADODB.Recordset
Dim vlDias As Long
Dim vlFechaPago As String

    fgCalcularFechaPrimerPagoEst = ""
   
    If (iTipoRenta = cgTipoRentaInmediata) Or (iTipoRenta = "6") Then
        If (iTipoPension = clCodTipPensionInvTot Or iTipoPension = clCodTipPensionInvPar Or iTipoPension = clCodTipPensionSob) Then
            If iFecRecPrima <> "" Then
                vlFechaPago = DateSerial(Mid(iFecRecPrima, 1, 4), Mid(iFecRecPrima, 5, 2), Mid(iFecRecPrima, 7, 2) + iNumDias)
            End If
        Else
            If iFecAcep <> "" Then
                vlFechaPago = DateSerial(Mid(iFecAcep, 1, 4), Mid(iFecAcep, 5, 2), Mid(iFecAcep, 7, 2) + iNumDias)
            End If
        End If
    Else
        If iFecDev <> "" Then
            vlFechaPago = DateSerial(Mid(iFecDev, 1, 4), Mid(iFecDev, 5, 2) + iMesDif, 1)
        End If
    End If
    
    fgCalcularFechaPrimerPagoEst = vlFechaPago
End Function

Public Function fgObtieneDiasPrimerPagoEst(iTipoPension As String) As Long
'Permite estimar los Dias de Plazo para realizar el Primer Pago
'Parámetros de Entrada:
'- iTipoPension As String   => Tipo de Pensión
'Retorno Función:
'- Número de Días
'------------------------------------------------------
'Fecha de Creación     : 11/07/2007
'Fecha de Modificación :
'------------------------------------------------------

Dim vlRegFecha As ADODB.Recordset
Dim vlDias As Long
Dim vlFechaPago As String

    fgObtieneDiasPrimerPagoEst = 0
    vlDias = 0
    
    vgQuery = "SELECT mto_elemento as dias FROM "
    vgQuery = vgQuery & "MA_TPAR_TABCOD WHERE "
    vgQuery = vgQuery & "cod_tabla = '" & vgCodTabla_TipPen & "' AND "
    vgQuery = vgQuery & "cod_elemento = '" & iTipoPension & "' "
    Set vlRegFecha = vgConexionBD.Execute(vgQuery)
    If Not (vlRegFecha.EOF) Then
        If Not IsNull(vlRegFecha!dias) Then
            vlDias = vlRegFecha!dias
            vlDias = vlDias - 1  'ABV 01/09/2007
        End If
    End If
    vlRegFecha.Close
    fgObtieneDiasPrimerPagoEst = vlDias
    
End Function

Public Function fgObtenerCodigo_TextoCompuesto(iTexto As String) As String
'Función: Permite obtener el Código de un Texto que tiene el Código y la
'Descripción separados por un Guión
'Parámetros de Entrada :
'- iTexto     => Texto que contiene el Código y Descripción
'Parámetros de Salida :
'- Devuelve el código del Texto
    
    If (InStr(1, iTexto, "-") <> 0) Then
        fgObtenerCodigo_TextoCompuesto = Trim(Mid(iTexto, 1, InStr(1, iTexto, "-") - 1))
    Else
        fgObtenerCodigo_TextoCompuesto = UCase(Trim(iTexto))
    End If

End Function

Public Function fgObtenerNombre_TextoCompuesto(iTexto As String) As String
'Función: Permite obtener el Nombre o Descripción de un Texto que tiene el Código y la
'Descripción separados por un Guión
'Parámetros de Entrada :
'- iTexto     => Texto que contiene el Código y Descripción
'Parámetros de Salida :
'- Devuelve la descripción del Texto
    
    If (InStr(1, iTexto, "-") <> 0) Then
        fgObtenerNombre_TextoCompuesto = Trim(Mid(iTexto, InStr(1, iTexto, "-") + 1, Len(iTexto)))
    Else
        fgObtenerNombre_TextoCompuesto = UCase(Trim(iTexto))
    End If

End Function

Public Function fgBuscarPorcBenSocial(iFecha As String, oValor As Double) As Boolean
Dim vlRegFecha As ADODB.Recordset
    
    fgBuscarPorcBenSocial = False
    
    vgQuery = "SELECT mto_valor "
    vgQuery = vgQuery & "FROM ma_tval_bensocial "
    vgQuery = vgQuery & "WHERE '" & iFecha & "' "
    vgQuery = vgQuery & "BETWEEN fec_inivig AND fec_tervig "
    Set vlRegFecha = vgConexionBD.Execute(vgQuery)
    If Not (vlRegFecha.EOF) Then
        If Not IsNull(vlRegFecha!mto_valor) Then
            oValor = vlRegFecha!mto_valor
            fgBuscarPorcBenSocial = True
        End If
    End If
    vlRegFecha.Close

End Function

Public Function fgBuscarSucCorredor(iTipoIden As String, iNumIden As String) As String
Dim vlRegBuscar As ADODB.Recordset
    
    fgBuscarSucCorredor = "00000"
    
    vgQuery = "SELECT cod_sucursal "
    vgQuery = vgQuery & "FROM pt_tmae_corredor "
    vgQuery = vgQuery & "WHERE "
    vgQuery = vgQuery & "cod_tipoidencor = '" & iTipoIden & "' AND "
    vgQuery = vgQuery & "num_idencor = '" & iNumIden & "' "
    Set vlRegBuscar = vgConexionBD.Execute(vgQuery)
    If Not (vlRegBuscar.EOF) Then
        If Not IsNull(vlRegBuscar!Cod_Sucursal) Then
            fgBuscarSucCorredor = vlRegBuscar!Cod_Sucursal
        End If
    End If
    vlRegBuscar.Close

End Function

Public Function fgCalcularFechaFinPerDiferido(iFecDev As String, iMesDif As Long) As String
'Permite determinar la Fecha de Termino del Periodo Diferido
'Parámetros de Entrada:
'- iMesDif      => Número de Meses Diferidos
'- iFechaDev    => Fecha de Devengue
'Parámetros de Salida:
'- Retorna      => Fecha Fin Periodo Diferido "yyyymmdd"
'------------------------------------------------------
'Fecha de Creación     : 07/07/2007
'Fecha de Modificación : 09/08/2007
'------------------------------------------------------
    
    fgCalcularFechaFinPerDiferido = ""
    
    If (iMesDif > 0) Then
        fgCalcularFechaFinPerDiferido = Format(DateSerial(Mid(iFecDev, 1, 4), Mid(iFecDev, 5, 2) + iMesDif, 1 - 1), "yyyymmdd")
    End If
    
End Function

Public Function fgCalcularFechaFinPerGarantizado(iFecDev As String, iMesDif As Long, iMesGar As Long) As String
'Permite determinar la Fecha de Termino del Periodo Garantizado
'Parámetros de Entrada:
'- iMesDif      => Número de Meses Diferidos
'- iMesGar      => Número de Meses Garantizados
'- iFechaDev    => Fecha de Devengue
'Parámetros de Salida:
'- Retorna      => Fecha Fin Periodo Garantizado "yyyymmdd"
'------------------------------------------------------
'Fecha de Creación     : 07/07/2007
'Fecha de Modificación : 09/08/2007
'------------------------------------------------------
Dim vlFecha As String

    fgCalcularFechaFinPerGarantizado = ""
    
    If (iMesGar > 0) Then
        If (iMesDif > 0) Then
            vlFecha = Format(DateSerial(Mid(iFecDev, 1, 4), Mid(iFecDev, 5, 2) + iMesDif, 1), "yyyymmdd")
        Else
            vlFecha = iFecDev
        End If
        fgCalcularFechaFinPerGarantizado = Format(DateSerial(Mid(vlFecha, 1, 4), Mid(vlFecha, 5, 2) + iMesGar, 1 - 1), "yyyymmdd")
    End If

End Function

Function fgCalcularFechaIniPagoPensiones(iFecDev As String, iMesDif As Long) As String
'Permite determinar la Fecha de Inicio de Pago de Pensiones
'Parámetros de Entrada:
'- iMesDif      => Número de Meses Diferidos
'- iFechaDev    => Fecha de Devengue
'Parámetros de Salida:
'- Retorna      => Fecha de Inicio de Pago de Pensiones "yyyymmdd"
'-----------------------------------------------------------------
'Fecha de Creación     : 09/08/2007
'Fecha de Modificación :
'------------------------------------------------------
Dim vlFecha As String

    fgCalcularFechaIniPagoPensiones = ""
    
    If (iMesDif > 0) Then
        fgCalcularFechaIniPagoPensiones = Format(DateSerial(Mid(iFecDev, 1, 4), Mid(iFecDev, 5, 2) + iMesDif, 1), "yyyymmdd")
    Else
        fgCalcularFechaIniPagoPensiones = iFecDev
    End If

End Function

Function fgBuscarGlosaCobConyuge(iElemento) As String
On Error GoTo Err_fgBuscarGlosaCobConyuge

    fgBuscarGlosaCobConyuge = ""
    
    vgSql = ""
    vgSql = "SELECT gls_cobercon "
    vgSql = vgSql & "FROM ma_tpar_cobercon WHERE "
    vgSql = vgSql & "cod_cobercon = '" & iElemento & "' "
    Set vgRegistro = vgConexionBD.Execute(vgSql)
    If Not vgRegistro.EOF Then
        fgBuscarGlosaCobConyuge = (vgRegistro!GLS_COBERCON)
    End If
    vgRegistro.Close
    
Exit Function
Err_fgBuscarGlosaCobConyuge:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgComboSiNo(vlCombo As ComboBox)
On Error GoTo Err_Combo

    vlCombo.Clear
    vlCombo.AddItem ("N - No")
    vlCombo.AddItem ("S - Si")
    vlCombo.ListIndex = 0

Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgComboCoberturaConyuge(vlCombo As ComboBox)
Dim vlRegCombo As ADODB.Recordset
On Error GoTo Err_Combo

    vlCombo.Clear
    
    vgQuery = "SELECT cod_cobercon as codigo "
    vgQuery = vgQuery & "FROM MA_TPAR_COBERCON "
    vgQuery = vgQuery & "ORDER BY codigo "
    Set vlRegCombo = vgConexionBD.Execute(vgQuery)
    If Not (vlRegCombo.EOF) Then
        While Not (vlRegCombo.EOF)
            vlCombo.AddItem (vlRegCombo!Codigo)
            vlRegCombo.MoveNext
        Wend
    End If
    vlRegCombo.Close
        
    If vlCombo.ListCount <> 0 Then
        vlCombo.ListIndex = 0
    End If

Exit Function
Err_Combo:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Public Function fgObtenerCod_Identificacion(iNombreCorto As String, oCodigo As Long) As Boolean
Dim vlRegBuscar As ADODB.Recordset
    
    fgObtenerCod_Identificacion = False
    oCodigo = 0
    
    vgQuery = "SELECT cod_tipoiden as codigo "
    vgQuery = vgQuery & "FROM ma_tpar_tipoiden "
    vgQuery = vgQuery & "WHERE "
    vgQuery = vgQuery & "gls_tipoidencor = '" & iNombreCorto & "'"
    Set vlRegBuscar = vgConexionBD.Execute(vgQuery)
    If Not (vlRegBuscar.EOF) Then
        If Not IsNull(vlRegBuscar!Codigo) Then
            oCodigo = vlRegBuscar!Codigo
            fgObtenerCod_Identificacion = True
        End If
    End If
    vlRegBuscar.Close

End Function

Public Function fgObtenerNombreSuc_Usuario(iCodigo As String) As String
Dim vlRegBuscar As ADODB.Recordset
    
    fgObtenerNombreSuc_Usuario = ""
    
    vgQuery = "SELECT gls_sucursal as nombre "
    vgQuery = vgQuery & "FROM pd_tpar_sucursal "
    vgQuery = vgQuery & "WHERE "
    vgQuery = vgQuery & "cod_sucursal = '" & iCodigo & "'"
    Set vlRegBuscar = vgConexionBD.Execute(vgQuery)
    If Not (vlRegBuscar.EOF) Then
        If Not IsNull(vlRegBuscar!nombre) Then
            fgObtenerNombreSuc_Usuario = Trim(vlRegBuscar!nombre)
        End If
    End If
    vlRegBuscar.Close

End Function

Function fgFormarNombreCompleto(iNombre As String, iNombreSeg As String, iPaterno As String, iMaterno As String) As String

fgFormarNombreCompleto = ""

If (iNombre = "") Then iNombre = "" Else iNombre = iNombre & " "
If (iNombreSeg = "") Then iNombreSeg = "" Else iNombreSeg = iNombreSeg & " "
If (iPaterno = "") Then iPaterno = "" Else iPaterno = iPaterno & " "
If (iMaterno = "") Then iMaterno = "" Else iMaterno = iMaterno & " "

fgFormarNombreCompleto = Trim(iNombre & iNombreSeg & iPaterno & iMaterno)

End Function

Sub p_centerForm(ByRef objFormMain As Form, ByRef objFormChild As Form)
   
   objFormChild.Left = (objFormMain.ScaleWidth - objFormChild.Width) / 2
   objFormChild.Top = (objFormMain.ScaleHeight - objFormChild.Height) / 2

End Sub

'marco agrego de la fuentes RV

Public Function fgObtenerPolizaCod_AFP(iNumPoliza As String, inumendoso As String) As String
Dim vlRegBuscar As ADODB.Recordset
    
    fgObtenerPolizaCod_AFP = ""
    
    vgQuery = "SELECT cod_afp as codigo "
    vgQuery = vgQuery & "FROM pp_tmae_poliza "
    vgQuery = vgQuery & "WHERE "
    vgQuery = vgQuery & "num_poliza = '" & iNumPoliza & "' AND "
    vgQuery = vgQuery & "num_endoso = " & inumendoso & " "
    Set vlRegBuscar = vgConexionBD.Execute(vgQuery)
    If Not (vlRegBuscar.EOF) Then
        If Not IsNull(vlRegBuscar!Codigo) Then
            fgObtenerPolizaCod_AFP = vlRegBuscar!Codigo
        End If
    End If
    vlRegBuscar.Close

End Function

Function fgValidaVigenciaPoliza(iNumPoliza As String, iFecha As String) As Boolean
Const clEstado = "9"
Dim inumend As Integer

    fgValidaVigenciaPoliza = False
    iFecha = Format(iFecha, "yyyymmdd")
    
    vgSql = "select max(num_endoso) as num_endoso from Pd_TMAE_oriPOLIZA "
    vgSql = vgSql & "where num_poliza = '" & iNumPoliza & "' "
    vgSql = vgSql & "order by num_endoso desc"
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        inumend = vgRs4!Num_Endoso
    Else
        vgRs4.Close
        Exit Function
    End If
    vgRs4.Close
    
    vgSql = "select cod_estado,FEC_VIGENCIA "
    vgSql = vgSql & "from PP_TMAE_POLIZA where "
    vgSql = vgSql & "num_poliza = '" & iNumPoliza & "' and "
    vgSql = vgSql & "num_endoso = " & inumend & " and "
    vgSql = vgSql & "cod_estado <> '" & clEstado & "' and "
    vgSql = vgSql & "fec_vigencia <= '" & iFecha & "'"
    Set vgRs4 = vgConexionBD.Execute(vgSql)
    If Not vgRs4.EOF Then
        fgValidaVigenciaPoliza = True
    End If
    vgRs4.Close
   
End Function
Function fgBuscarNombreProvinciaRegion(vlCodDir)
Dim vlRegistroNombre As ADODB.Recordset
On Error GoTo Err_Buscar

     vgSql = "SELECT r.Cod_Region,r.Gls_Region,p.Cod_Provincia,p.Gls_Provincia,c.Cod_Comuna,c.Gls_Comuna"
     vgSql = vgSql & " FROM MA_TPAR_COMUNA c, MA_TPAR_PROVINCIA p, MA_TPAR_REGION r"
     vgSql = vgSql & " Where c.Cod_Direccion = '" & vlCodDir & "' and  "
     vgSql = vgSql & " c.cod_region = p.cod_region and"
     vgSql = vgSql & " c.cod_provincia = p.cod_provincia and"
     vgSql = vgSql & " p.cod_region = r.cod_region"
     Set vlRegistroNombre = vgConexionBD.Execute(vgSql)
     If Not vlRegistroNombre.EOF Then
        vgNombreRegion = IIf(IsNull(vlRegistroNombre!gls_region), "", vlRegistroNombre!gls_region)
        vgNombreProvincia = IIf(IsNull(vlRegistroNombre!gls_provincia), "", vlRegistroNombre!gls_provincia)
        vgNombreComuna = IIf(IsNull(vlRegistroNombre!gls_comuna), "", vlRegistroNombre!gls_comuna)
     End If
     vlRegistroNombre.Close

Exit Function
Err_Buscar:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select

End Function

Function fgValidaFechaEfecto(iFecha As String, iNumPoliza As String, iNumOrden As Integer) As String
Dim vlOpcionPago As String
Dim vlPagoReg    As String
Dim vlAnno       As String, vlMes As String, vlDia As String
Dim iFechaInicio As String
Dim iFechaTermino As String
Dim iAnno As Integer
Dim iMes As Integer
Dim iDia As Integer
Dim vlRs3 As ADODB.Recordset
Dim vlRs2 As ADODB.Recordset

    fgValidaFechaEfecto = ""

    iFecha = Format(iFecha, "yyyymmdd")
    
'Calcula último día del mes
    iAnno = CInt(Mid(iFecha, 1, 4))
    iMes = CInt(Mid(iFecha, 5, 2))
    iDia = 1
    iFechaTermino = Format(DateSerial(iAnno, iMes + 1, iDia - 1), "yyyymmdd")
    
'Calcula Primer día del mes
    iAnno = CInt(Mid(iFecha, 1, 4))
    iMes = CInt(Mid(iFecha, 5, 2))
    iDia = 1
    iFechaInicio = Format(DateSerial(iAnno, iMes, iDia), "yyyymmdd")
            
    iFecha = Mid(iFecha, 1, 6)

    vlPagoReg = ""
    vlAnno = ""
    vlMes = ""
    vlDia = ""
    vgI = 0
    vgFechaEfecto = ""
    
    'Estados del Pago de Pensión
    'PP: Primer Pago
    'PR: Pago en Regimen
    vlOpcionPago = ""
    
    'Estados que puede tener el Periodo
    'A : Abierto
    'P : Provisorio
    'C : Cerrado

    'Determinar si el Caso es Primer Pago o Pago en Regimen
    'vgSql = "SELECT NUM_POLIZA,NUM_ENDOSO,NUM_ORDEN "
    'vgSql = vgSql & " FROM PP_TMAE_LIQPAGOPENDEF WHERE "
    'vgSql = vgSql & " NUM_POLIZA = '" & iNumPoliza & "' "
    ''vgSql = vgSql & " AND NUM_ORDEN = " & iNumOrden & ""
    '''vgSql = vgSql & " NUM_ENDOSO = " & inumend & " "
    vgSql = "SELECT num_poliza,num_endoso "
    vgSql = vgSql & " FROM pp_tmae_poliza A WHERE "
    vgSql = vgSql & " num_poliza = '" & iNumPoliza & "' "
    'vgSql = vgSql & " AND NUM_ORDEN = " & iNumOrden & ""
    vgSql = vgSql & " AND NUM_ENDOSO = "
    vgSql = vgSql & " (SELECT MAX(NUM_ENDOSO) FROM PP_TMAE_POLIZA " 'Último Endoso
    vgSql = vgSql & " WHERE NUM_POLIZA = A.NUM_POLIZA)"
    vgSql = vgSql & " AND FEC_INIPAGOPEN BETWEEN '" & iFechaInicio & "'"
    vgSql = vgSql & " AND '" & iFechaTermino & "'"
    Set vlRs3 = vgConexionBD.Execute(vgSql)
    If vlRs3.EOF Then
        'Pago Régimen
        vlOpcionPago = "PR"
    Else
        'Primer Pago
        vlOpcionPago = "PP"
    End If
    vlRs3.Close

    vlOpcionPago = "PR"
    
    'Determinar si el periodo a registrar es posterior al que se desea ingresar
    vgSql = ""
    vgSql = "SELECT NUM_PERPAGO,COD_ESTADOREG " ',COD_ESTADOPRI
    vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "num_perpago >= '" & iFecha & "' AND "
    If (vlOpcionPago = "PR") Then
        vgSql = vgSql & "cod_estadoreg <> 'C' "
    Else
        vgSql = vgSql & "cod_estadopri <> 'C' "
    End If
    vgSql = vgSql & "ORDER BY num_perpago ASC"
    Set vlRs2 = vgConexionBD.Execute(vgSql)
    If Not vlRs2.EOF Then
        'If vlOpcionPago = "PR" Then
            vlPagoReg = vlRs2!Num_PerPago
        '    'Pago Régimen
        '    If (vlRs2!cod_estadoreg) <> "C" Then
        '        fgValidaPagoPension = True
        '    Else
        '        vgI = 1
        '    End If
        'Else
        '    'Primer Pago
        '    vlPagoReg = vlRs2!Num_PerPago
        '    If (vlRs2!cod_estadopri) <> "C" Then
        '        fgValidaPagoPension = True
        '    Else
        '        vgI = 1
        '    End If
        'End If
    Else
        'Determinar si el periodo a registrar es posterior al que se desea ingresar
        vgSql = ""
        vgSql = "SELECT NUM_PERPAGO,COD_ESTADOREG " ',COD_ESTADOPRI
        vgSql = vgSql & "FROM PP_TMAE_PROPAGOPEN "
        vgSql = vgSql & "WHERE "
        vgSql = vgSql & "num_perpago >= '" & iFecha & "' AND "
        If (vlOpcionPago = "PR") Then
            vgSql = vgSql & "cod_estadoreg = 'C' "
        Else
            vgSql = vgSql & "cod_estadopri = 'C' "
        End If
        vgSql = vgSql & "ORDER BY num_perpago DESC"
        Set vlRs3 = vgConexionBD.Execute(vgSql)
        If Not vlRs3.EOF Then
            vlPagoReg = vlRs3!Num_PerPago
            vgI = 1
        Else
            vlPagoReg = iFecha
        End If
        vlRs3.Close
    End If
    vlRs2.Close
    
    If (vlPagoReg <> "") Then
        vlAnno = Mid(vlPagoReg, 1, 4)
        vlMes = Mid(vlPagoReg, 5, 2) + vgI
        vlDia = "01"
        vgFechaEfecto = DateSerial(vlAnno, vlMes, vlDia)
    End If
    
    fgValidaFechaEfecto = vgFechaEfecto
End Function
Public Function ObtenerFechaServer() As Date
    Dim fecha As Date
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    On Error GoTo mierror
    rs.CursorLocation = adUseClient
        rs.Open "select to_char(sysdate,'dd/mm/yyyy')as fecha from dual", vgConexionBD, adOpenStatic, adLockReadOnly
        If Not rs.EOF Then
            fecha = CDate(rs!fecha)
        End If
    ObtenerFechaServer = fecha
    Exit Function
mierror:
        MsgBox "No se puede ver Fecha del servidor", vbExclamation
End Function

Function FgGuardaLog(sQuery As String, sUser As String, nError As String) As Boolean
    Dim vlSql As String
    Dim oValor As Long
    Dim vlRegistro As ADODB.Recordset
    
    vlSql = "SELECT count(*) + 1 as Cor"
    vlSql = vlSql & " FROM PP_TMAE_LOGACTUAL a"
    Set vlRegistro = vgConexionBD.Execute(vlSql)
    If Not vlRegistro.EOF Then
        oValor = vlRegistro!Cor
    End If
    'vlRegistro.Close
    sQuery = Replace(sQuery, "'", "|")
    vlSql = ""
    vlSql = "INSERT INTO PP_TMAE_LOGACTUAL (LOG_NNUMREG, LOG_CCODERR, LOG_SQUERY ,LOG_DFECREG ,LOG_CUSUCRE) VALUES ("
    vlSql = vlSql & oValor & ",'" & Trim(nError) & "','" & Trim(sQuery) & "','" & Format(Date, "yyyymmdd") & "','" & sUser & "')"
    vgConexionBD.Execute (vlSql)

End Function


