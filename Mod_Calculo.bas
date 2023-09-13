Attribute VB_Name = "Mod_Calculo"
'Constantes para las Tablas de Mortalidad
Global Const vgTipoPeriodoMensual = "M"
Global Const vgTipoPeriodoAnual = "A"

Global Const vgTipoTablaRentista = "RV"
Global Const vgTipoTablaTotal = "MIT"
Global Const vgTipoTablaParcial = "MIP"
Global Const vgTipoTablaBeneficiario = "B"
Global Const vgEdadJubilacionHombre = 65
Global Const vgEdadJubilacionMujer = 65
Global Const vgTasaCapitalBono = 4

'Constantes para la agrupación por Tipo de Pensión
Const cgPensionInvVejez  As String = "02,04,05,06,07,14"
Const cgPensionSobOrigen As String = "08,13"
Const cgPensionSobTransf As String = "01,03,09,10,11,12,15"
Const cgParConyugeMadre  As String = "10,11,20,21"

'Costantes del Causante de la Póliza
Global Const cgCauCodPar As String * 2 = "99"

'Constantes de Estado de Pago Pensión
Global Const cgEstPension_NoPago As String = "10"
Global Const cgEstPension_SiPago As String = "99"

'Número de Endoso Inicial
Global Const cgNumeroEndosoInicial As String * 1 = "1"

'Estructura de Poliza
Public Type TyPoliza
    Num_Poliza As String
    Num_Endoso As Integer
    Num_Cot As String
    Num_Correlativo As Long
    Num_Operacion As String
    Num_Archivo As String
    Cod_AFP As String
    Cod_TipPension As String
    Cod_Cuspp As String
'    Rut_Afi As Long
'    Dgv_Afi As String
    Fec_Solicitud As String
    Fec_Vigencia As String
    Fec_Dev As String
    Fec_Acepta As String
    Fec_Emision As String
    Fec_Calculo As String
    Cod_MonedaFon As String
    Mto_MonedaFon As Double
    Mto_PrimaFon As Double
    Mto_ApoAdiFon As Double
    Mto_CtaIndFon As Double
    Mto_BonoFon As Double
    Mto_Prima As Double
    Mto_ApoAdi As Double
    Mto_CtaInd As Double
    Mto_Bono As Double
    Prc_TasaRPRT As Double
    Rut_Cor As Long
    Num_IdenCor As String
    Prc_CorCom As Double
    Prc_CorComReal As Double
    Mto_CorCom As Double
    Cod_BenSocial As String
    Ind_Cob As String
    Cod_Moneda As String
    Mto_ValMoneda As Double
    Mto_PriUniMod As Double
    Mto_CtaIndMod As Double
    Mto_BonoMod   As Double
    Mto_ApoAdiMod As Double
    Cod_TipRen As String
    Num_MesDif As Integer
    Cod_Modalidad As String
    Num_MesGar As Integer
    Cod_CoberCon As String
    Mto_FacPenElla   As Double
    Prc_FacPenElla   As Double
    Cod_DerCre As String
    Cod_DerGra As String
    Prc_RentaAFP As Double
    Prc_RentaTMP As Double
    Mto_CuoMor As Double
    Prc_TasaCe As Double
    Prc_TasaVta As Double
    Prc_TasaTir As Double
    Prc_TasaIntPerGar As Double
    Mto_CNU As Double
    Mto_PriUniSim As Double
    Mto_PriUniDif As Double
    Mto_Pension As Double
    Mto_PensionGar As Double
    Mto_CtaIndAFP As Double
    Mto_RentaTMPAFP As Double
    Mto_ResMat As Double
    Mto_ValPrePenTmp As Double
    Mto_PerCon As Double
    Prc_PerCon As Double
    Mto_SumPension As Double
    Mto_PenAnual As Double
    Mto_RMPension As Double
    Mto_RMGtoSep As Double
    Mto_RMGtoSepRV As Double
    Cod_TipoCot As String
    Cod_EstCot As String
    Num_Cargas As Integer
    Num_AnnoJub As Long
    Mto_AjusteIPC As Double
    Ind_CalSobDif As String      'I--- ABV 04/12/2009 ---
    Cod_TipReajuste As String    'I--- ABV 05/02/2011 ---
    Mto_ValReajusteTri As Double 'I--- ABV 05/02/2011 ---
    Mto_ValReajusteMen As Double 'I--- ABV 05/02/2011 ---
End Type

Public Type TyBeneficiarios
    Num_Poliza As String
    Num_Endoso As Integer
    Num_Orden As Integer
    Fec_Ingreso As String
    Rut_Ben As Long
    Dgv_Ben As String
    Gls_NomBen As String
    Gls_NomSegBen As String
    Gls_PatBen As String
    Gls_MatBen As String
    Gls_DirBen As String
    Cod_Direccion As Integer
    Gls_FonoBen As String
    Gls_CorreoBen As String
    Cod_GruFam  As String
    Cod_Par     As String
    Cod_Sexo    As String
    Cod_SitInv  As String
    Cod_DerCre  As String
    Cod_DerPen  As String
    Cod_CauInv  As String
    Fec_NacBen  As String
    Fec_NacHM   As String
    Fec_InvBen  As String
    Cod_MotReqPen As String
    Mto_Pension As Double
    Mto_PensionGar As Double
    Prc_Pension    As Double
    Prc_PensionLeg As Double
    Prc_PensionRep As Double
    Prc_PensionGar As Double
    Cod_InsSalud   As String
    Cod_ModSalud   As String
    Mto_PlanSalud  As Double
    Cod_EstPension As String
    Cod_CajaCompen As String
    Cod_ViaPago    As String
    Cod_Banco      As String
    Cod_TipCuenta  As String
    Num_Cuenta     As String
    Cod_Sucursal   As String
    Fec_FallBen    As String
    Fec_Matrimonio As String
    Cod_CauSusBen  As String
    Fec_SusBen     As String
    Fec_IniPagoPen As String
    Fec_TerPagoPenGar As String
    Cod_UsuarioCrea As String
    Fec_Crea As String
    Hor_Crea As String
    Cod_ModSalud2 As String
    Mto_PlanSalud2 As Double
    Num_Fun As String
End Type

Global stPolizaOri        As TyPoliza 'Registro de Poliza Original
Global stPolizaMod        As TyPoliza 'Registro de Poliza Modificada
Global stBeneficiariosOri() As TyBeneficiarios 'Registro de Beneficiarios Originales
Global stBeneficiariosMod() As TyBeneficiarios 'Registro de Beneficiarios Modificados
Global vgValorParametro As Double
Global vgValorPorcentaje As Double

'Definición de la Estructura que guardará a las Tablas de Mortalidad
Public Type TypeTabla
    Correlativo As Double
    TipoTabla   As String
    Nombre      As String
    Sexo        As String
    FechaIni    As Long
    FechaFin    As Long
    TipoPeriodo As String
    TipoGenerar As String
    IniTab      As Long
    Fintab      As Long
    Tasa        As Double
    Oficial     As String
    Estado      As String
    TipoMovimiento As String
    AñoBase     As Integer
    Descripcion As String
End Type
Global egTablaMortal()      As TypeTabla
Global vgNumeroTotalTablas      As Long

'Códigos de Tablas Mensuales
Global vgMortalVit_M As Long
Global vgMortalTot_M As Long
Global vgMortalPar_M As Long
Global vgMortalBen_M As Long
Global vgMortalVit_F As Long
Global vgMortalTot_F As Long
Global vgMortalPar_F As Long
Global vgMortalBen_F As Long

Global vgPalabra_MortalVit_M As String
Global vgPalabra_MortalTot_M As String
Global vgPalabra_MortalPar_M As String
Global vgPalabra_MortalBen_M As String
Global vgPalabra_MortalVit_F As String
Global vgPalabra_MortalTot_F As String
Global vgPalabra_MortalPar_F As String
Global vgPalabra_MortalBen_F As String

Global Fintab        As Long
Global vgFinTabVit_F As Long
Global vgFinTabTot_F As Long
Global vgFinTabPar_F As Long
Global vgFinTabBen_F As Long
Global vgFinTabVit_M As Long
Global vgFinTabTot_M As Long
Global vgFinTabPar_M As Long
Global vgFinTabBen_M As Long

Global L24      As Double
Global L21      As Double
Global L18      As Double
Global Lx()     As Double
Global Ly()     As Double

Global vgBuscarMortalVit_F As String
Global vgBuscarMortalTot_F As String
Global vgBuscarMortalPar_F As String
Global vgBuscarMortalBen_F As String
Global vgBuscarMortalVit_M As String
Global vgBuscarMortalTot_M As String
Global vgBuscarMortalPar_M As String
Global vgBuscarMortalBen_M As String

Global vgFechaIniMortalVit_F As String
Global vgFechaFinMortalVit_F As String
Global vgFechaIniMortalTot_F As String
Global vgFechaFinMortalTot_F As String
Global vgFechaIniMortalPar_F As String
Global vgFechaFinMortalPar_F As String
Global vgFechaIniMortalBen_F As String
Global vgFechaFinMortalBen_F As String
Global vgFechaIniMortalVit_M As String
Global vgFechaFinMortalVit_M As String
Global vgFechaIniMortalTot_M As String
Global vgFechaFinMortalTot_M As String
Global vgFechaIniMortalPar_M As String
Global vgFechaFinMortalPar_M As String
Global vgFechaIniMortalBen_M As String
Global vgFechaFinMortalBen_M As String

Global vgFechaAnterior     As String
Global vgError             As Long

'I--- ABV 15/03/2005 ---
Global vgIndicadorTipoMovimiento_F As String
Global vgIndicadorTipoMovimiento_M As String
Global vgDinamicaAñoBase_F         As Integer
Global vgDinamicaAñoBase_M         As Integer
'F--- ABV 15/03/2005 ---

Global vgFactorElla          As Double
Global vgPorcentajeElla      As Double
Global vgRentabilidadAFP     As Double
Global vgFactorAjusteIPC     As Double

Global vgUtilizarNormativa As String
Global vgReservaUtilizada  As String
Global vgBotonEscogido     As String
Global vgCodDireccion      As String

Function fgCarga_Param(iTabla As String, iElemento As String, iFecha As String) As Boolean
Dim vlRegistro As ADODB.Recordset

    fgCarga_Param = False
    vgValorParametro = 0

    vgSql = "SELECT mto_elemento FROM MA_TPAR_TABCODVIG WHERE "
    vgSql = vgSql & "COD_TABLA = '" & iTabla & "' and "
    vgSql = vgSql & "COD_ELEMENTO = '" & iElemento & "' "
    vgSql = vgSql & "AND (FEC_INIVIG <= '" & iFecha & "' "
    vgSql = vgSql & "AND FEC_TERVIG >= '" & iFecha & "') "
    Set vlRegistro = vgConexionBD.Execute(vgSql)
    'If Not (vlRegistro.EOF) Then
    If Not (vlRegistro.EOF) Then
        If Not IsNull(vlRegistro!mto_elemento) Then
            vgValorParametro = Trim(vlRegistro!mto_elemento)

            fgCarga_Param = True
        End If
    End If
    vlRegistro.Close

End Function

Public Function fgBuscarPorcCobConyuge(iCodigo As String, oValor As Double) As Boolean
Dim vlRegFecha As ADODB.Recordset
    
    fgBuscarPorcCobConyuge = False
    oValor = 0
    
    vgQuery = "SELECT mto_cobercon as mto_valor "
    vgQuery = vgQuery & "FROM ma_tpar_cobercon "
    vgQuery = vgQuery & "WHERE "
    vgQuery = vgQuery & "cod_cobercon = '" & iCodigo & "'"
    Set vlRegFecha = vgConexionBD.Execute(vgQuery)
    If Not (vlRegFecha.EOF) Then
        If Not IsNull(vlRegFecha!mto_valor) Then
            oValor = vlRegFecha!mto_valor
            fgBuscarPorcCobConyuge = True
        End If
    End If
    vlRegFecha.Close

End Function

Function fgCargaEstPoliza(iFormulario As Form, istPolizas As TyPoliza, iNumPoliza As String, _
iFecCalculo As String, iFecVigencia As String, iFecEmision As String, iBotonEscogido As String, _
iNumCotizacion As String, iNumCorrelativo As String)

Dim vlPos, vlNumero As Integer
Const clNumEndoso As String * 1 = "0"

Dim vlNumPoliza As String, vlNumEndoso As Long
Dim vlNumCot As String, vlNumCorrelativo As Long
Dim vlNumOperacion As String, vlNumArchivo As String
Dim vlCodAFP As String, vlCodTipPension As String
'Dim vlCodCuspp As String
'Dim vlRutAfi As Long, vlDgvAfi As String
Dim vlFecCalculo As String, vlFecSolicitud As String, vlFecVigencia As String
Dim vlFecDev As String, vlFecAcepta As String, Fec_Emision As String
Dim vlCodMonedaFon As String, vlMtoMonedaFon As Double
Dim vlMtoPrimaFon  As Double, vlMtoApoAdiFon As Double
Dim vlMtoCtaIndFon As Double, vlMtoBonoFon As Double
Dim vlMtoPrima     As Double, vlMtoApoAdi    As Double
Dim vlMtoCtaInd    As Double, vlMtoBono      As Double
Dim vlPrcTasaRPRT  As Double
Dim vlRutCor As Long, vlNumIdenCor As String
Dim vlPrcCorCom As Double, vlPrcCorComReal As Double, vlMtoCorCom As Double
Dim vlCodBenSocial As String
Dim vlIndCob As String
Dim vlCodMoneda As String, vlMtoValMoneda As Double
Dim vlMtoPriUniMod As Double, vlMtoCtaIndMod As Double, vlMtoBonoMod As Double
Dim vlCodTipRen As String, vlNumMesDif As Integer
Dim vlCodModalidad As String, vlNumMesGar As Integer
Dim vlCodCoberCon As String, vlMtoFacPenElla As Double, vlPrcFacPenElla As Double
Dim vlCodDerCre As String, vlCodDerGra As String
Dim vlPrcRentaAFP As Double, vlPrcRentaTMP As Double
Dim vlMtoCuoMor As Double
Dim vlPrcTasaCe As Double, vlPrcTasaVta As Double
Dim vlPrcTasaTir As Double, vlPrcTasaPerGar As Double
Dim vlMtoCNU As Double, vlMtoResMat As Double
Dim vlMtoPerCon As Double, vlPrcPerCon As Double
Dim vlMtoPenAnual As Double, vlMtoRMPension As Double
Dim vlMtoRMGtoSep As Double
Dim vlCodTipoCot As String, vlCodEstCot As String
Dim vlNumCargas As Integer, vlNumAnnoJub As Long
Dim vlPorcentaje As Double, vlSexo As String
Dim vlExiste As Boolean
Dim vlcuspp As String

    vlNumPoliza = iNumPoliza
    vlNumEndoso = clNumEndoso
    vlFecCalculo = iFecCalculo
    vlFecVigencia = iFecVigencia
    vlFecEmision = iFecEmision
    
    'Los Datos registrados desde la Póliza y que no se ingresan desde la pantalla
    If (iBotonEscogido = "C") Then
        vlSql = " SELECT d.num_cot,d.num_correlativo,"
        vlSql = vlSql & " d.num_operacion,d.num_archivo,c.fec_suscripcion as fec_solicitud,"
        vlSql = vlSql & " d.mto_cuomor,d.prc_tasatce as prc_tasace,d.prc_tasavta,d.prc_tasatir,"
        vlSql = vlSql & " d.prc_tasapergar,d.mto_cnu,d.mto_resmat,d.mto_percon,d.prc_percon,"
        vlSql = vlSql & " d.mto_penanual,d.mto_rmpension,d.mto_rmgtosep,c.cod_tipcot,"
        vlSql = vlSql & " d.cod_estcot,d.mto_valmoneda "
        vlSql = vlSql & " FROM pt_tmae_detcotizacion d, pt_tmae_cotizacion c "
        vlSql = vlSql & " WHERE "
        vlSql = vlSql & " d.num_cot = '" & iNumCotizacion & "' AND "
        vlSql = vlSql & " d.num_correlativo = '" & iNumCorrelativo & "' "
    Else
        vlSql = "SELECT num_cot,num_correlativo,"
        vlSql = vlSql & "num_operacion,num_archivo,fec_solicitud,"
        vlSql = vlSql & "mto_cuomor,prc_tasace,prc_tasavta,prc_tasatir,"
        vlSql = vlSql & "prc_tasapergar,mto_cnu,mto_resmat,mto_percon,prc_percon,"
        vlSql = vlSql & "mto_penanual,mto_rmpension,mto_rmgtosep,cod_tipcot,"
        vlSql = vlSql & "cod_estcot,mto_valmoneda "
        vlSql = vlSql & "FROM pd_tmae_oripoliza "
        vlSql = vlSql & "WHERE num_poliza = '" & iNumPoliza & "' "
    End If
    Set vgRs = vgConexionBD.Execute(vlSql)
    If Not (vgRs.EOF) Then
            
        vlNumCot = Trim(vgRs!Num_Cot)
        vlNumCorrelativo = vgRs!Num_Correlativo
        vlNumOperacion = vgRs!Num_Operacion
        vlNumArchivo = vgRs!Num_Archivo
        vlFecSolicitud = vgRs!Fec_Solicitud
        vlMtoCuoMor = vgRs!Mto_CuoMor
        vlPrcTasaCe = vgRs!Prc_TasaCe
        vlPrcTasaVta = vgRs!Prc_TasaVta
        vlPrcTasaTir = vgRs!Prc_TasaTir
        vlPrcTasaPerGar = vgRs!prc_tasapergar
        vlMtoCNU = vgRs!Mto_CNU
        vlMtoResMat = vgRs!Mto_ResMat
        vlMtoPerCon = vgRs!Mto_PerCon
        vlPrcPerCon = vgRs!Prc_PerCon
        vlMtoPenAnual = vgRs!Mto_PenAnual
        vlMtoRMPension = vgRs!Mto_RMPension
        vlMtoRMGtoSep = vgRs!Mto_RMGtoSep
        vlCodTipoCot = vgRs!cod_tipcot
        vlCodEstCot = vgRs!Cod_EstCot
        vlMtoValMoneda = vgRs!Mto_ValMoneda
        vlPrcTasaRPRT = vgRentabilidadAFP
        
    End If
    vgRs.Close
    
    'Obtener los Datos desde el Formulario
    vlCodAFP = fgObtenerCodigo_TextoCompuesto(iFormulario.Cmb_Afp)
    vlCodTipPension = fgObtenerCodigo_TextoCompuesto(iFormulario.Cmb_TipoPension)
    vlNumCargas = iFormulario.Txt_Asegurados
    vlFecDev = Format(iFormulario.Txt_FecDev, "yyyymmdd")
    vlFecAcepta = Format(iFormulario.Txt_FecIncorpora, "yyyymmdd")
    
    vlIndCob = fgObtenerCodigo_TextoCompuesto(iFormulario.Cmb_IndCob)
    
    vlcuspp = iFormulario.Txt_Cuspp
    
    'Cálculo de la Prima
    vlCodMonedaFon = cgCodTipMonedaUF
    vlMtoMonedaFon = cgMonedaValorNS
    vlMtoPrimaFon = iFormulario.Lbl_PriUnica
    vlMtoApoAdiFon = 0
    vlMtoCtaIndFon = iFormulario.Txt_CtaInd
    vlMtoBonoFon = iFormulario.Txt_BonoAct
    vlMtoPrima = Format(vlMtoPrimaFon / vlMtoMonedaFon, "#0.00")
    vlMtoApoAdi = Format(vlMtoApoAdiFon / vlMtoMonedaFon, "#0.00")
    vlMtoCtaInd = Format(vlMtoCtaIndFon / vlMtoMonedaFon, "#0.00")
    vlMtoBono = Format(vlMtoBonoFon / vlMtoMonedaFon, "#0.00")
    
    'Antecedentes de la Modalidad
    vlCodMoneda = fgObtenerCodigo_TextoCompuesto(iFormulario.Cmb_Moneda)
    'vlMtoValMoneda = 3.15
    'vlMtoValMoneda
    vlMtoPriUniMod = Format(vlMtoPrima / vlMtoValMoneda, "#0.00")
    vlMtoCtaIndMod = Format(vlMtoCtaInd / vlMtoValMoneda, "#0.00")
    vlMtoBonoMod = Format(vlMtoBono / vlMtoValMoneda, "#0.00")
    vlCodTipRen = fgObtenerCodigo_TextoCompuesto(iFormulario.Cmb_TipoRenta)
    vlNumMesDif = iFormulario.Txt_AnnosDif * 12
    vlCodModalidad = fgObtenerCodigo_TextoCompuesto(iFormulario.Cmb_Modalidad)
    vlNumMesGar = iFormulario.Txt_MesesGar
    vlPrcRentaAFP = vgRentabilidadAFP
    vlPrcRentaTMP = iFormulario.Txt_PrcRentaTmp
    vlCodCoberCon = Trim(iFormulario.Cmb_CobConyuge)
    vlCodDerCre = fgObtenerCodigo_TextoCompuesto(iFormulario.Cmb_DerCre)
    vlCodDerGra = fgObtenerCodigo_TextoCompuesto(iFormulario.Cmb_DerGra)
    
    'Cálcular el Factor de la Cobertura a la Cónyuge
    vlMtoFacPenElla = 1
    If (vlCodCoberCon <> "0") Then
        If (fgBuscarPorcCobConyuge(vlCodCoberCon, vlPrcFacPenElla) = True) Then
            'Buscar porcentaje del o la Cónyuge S/Hijos
            vlSexo = fgObtenerCodigo_TextoCompuesto(iFormulario.Cmb_Sexo)
            If (vlSexo = "M") Then vlSexo = "F" Else vlSexo = "M"
            If (fgObtenerPorcentaje("10", "N", vlSexo, vlFecCalculo) = True) Then
                vlPorcentaje = vgValorPorcentaje
                vlMtoFacPenElla = Format(vlPrcFacPenElla / vlPorcentaje, "#0.00000000")
            End If
        End If
    End If
    
    'Calcular los Datos del Intermediario
    vlRutCor = fgObtenerCodigo_TextoCompuesto(iFormulario.Lbl_TipoIdentCorr)
    vlNumIdenCor = iFormulario.Lbl_NumIdentCorr
    
    If (vlRutCor <> 0) Then
        vlPrcCorComReal = iFormulario.Txt_ComInt
        vlCodBenSocial = Mid(iFormulario.Lbl_BenSocial, 1, 1)
        If (vlCodBenSocial = "S") Then
            If (fgObtenerPorcentajeBenSocial(vlFecCalculo, vlPorcentaje) = False) Then
                vlPrcCorCom = vlPrcCorComReal
                vlMtoCorCom = Format((vlMtoPrima * (vlPrcCorComReal / 100)), "#0.00")
            Else
                vlPrcCorCom = Format((vlPrcCorComReal * (vlPorcentaje / 100)) + vlPrcCorComReal, "#0.00")
                vlMtoCorCom = Format((vlMtoPrima * (vlPrcCorCom / 100)), "#0.00")
            End If
        Else
            vlPrcCorCom = vlPrcCorComReal
            vlMtoCorCom = Format((vlMtoPrima * (vlPrcCorComReal / 100)), "#0.00")
        End If
    Else
        vlPrcCorCom = 0
        vlPrcCorComReal = 0
        vlMtoCorCom = 0
        vlCodBenSocial = "N"
    End If
    
    '-***************************
    'Calcular el Año de Jubilación
    vlSexo = fgObtenerCodigo_TextoCompuesto(iFormulario.Cmb_Sexo)
    If (vlSexo = "M") Then
        vlNumAnnoJub = Year(CDate(iFormulario.Txt_FecNac)) + vgEdadJubilacionHombre
    Else
        vlNumAnnoJub = Year(CDate(iFormulario.Txt_FecNac)) + vgEdadJubilacionMujer
    End If

    vgX = 1
    With istPolizas
        .Num_Poliza = vlNumPoliza
        .Num_Endoso = vlNumEndoso
        .Num_Cot = vlNumCot
        .Num_Correlativo = vlNumCorrelativo
        .Num_Operacion = vlNumOperacion
        .Num_Archivo = vlNumArchivo
        .Cod_AFP = vlCodAFP
        .Cod_TipPension = vlCodTipPension
        .Fec_Solicitud = vlFecSolicitud
        .Fec_Vigencia = vlFecVigencia
        .Fec_Dev = vlFecDev
        .Fec_Acepta = vlFecAcepta
        .Fec_Emision = vlFecEmision
        .Fec_Calculo = iFecCalculo
        .Cod_MonedaFon = vlCodMonedaFon
        .Mto_MonedaFon = vlMtoMonedaFon
        .Mto_PrimaFon = vlMtoPrimaFon
        '.Mto_ApoAdiFon = vlMtoApoAdiFon
        .Mto_CtaIndFon = vlMtoCtaIndFon
        .Mto_BonoFon = vlMtoBonoFon
        .Mto_Prima = vlMtoPrima
        .Mto_ApoAdi = vlMtoApoAdi
        .Mto_CtaInd = vlMtoCtaInd
        .Mto_Bono = vlMtoBono
        .Prc_TasaRPRT = vlPrcTasaRPRT
        .Rut_Cor = vlRutCor
        .Num_IdenCor = vlNumIdenCor
        .Prc_CorCom = vlPrcCorCom
        .Prc_CorComReal = vlPrcCorComReal
        .Mto_CorCom = vlMtoCorCom
        .Cod_BenSocial = vlCodBenSocial
        .Ind_Cob = vlIndCob
        .Cod_Moneda = vlCodMoneda
        .Mto_ValMoneda = vlMtoValMoneda
        .Mto_PriUniMod = vlMtoPriUniMod
        .Mto_CtaIndMod = vlMtoCtaIndMod
        .Mto_BonoMod = vlMtoBonoMod
        .Cod_TipRen = vlCodTipRen
        .Num_MesDif = vlNumMesDif
        .Cod_Modalidad = vlCodModalidad
        .Num_MesGar = vlNumMesGar
        .Cod_CoberCon = vlCodCoberCon
        .Mto_FacPenElla = vlMtoFacPenElla
        .Prc_FacPenElla = vlPrcFacPenElla
        .Cod_DerCre = vlCodDerCre
        .Cod_DerGra = vlCodDerGra
        .Prc_RentaAFP = vlPrcRentaAFP
        .Prc_RentaTMP = vlPrcRentaTMP
        .Mto_CuoMor = vlMtoCuoMor
        .Prc_TasaCe = vlPrcTasaCe
        .Prc_TasaVta = vlPrcTasaVta
        .Prc_TasaTir = vlPrcTasaTir
        .Prc_TasaIntPerGar = vlPrcTasaPerGar
        .Mto_CNU = vlMtoCNU
        .Mto_ResMat = vlMtoResMat
        .Mto_PerCon = vlMtoPerCon
        .Prc_PerCon = vlPrcPerCon
        .Mto_PenAnual = vlMtoPenAnual
        .Mto_RMPension = vlMtoRMPension
        .Mto_RMGtoSep = vlMtoRMGtoSep
        .Cod_TipoCot = vlCodTipoCot
        .Cod_EstCot = vlCodEstCot
        .Num_Cargas = vlNumCargas
        .Cod_Cuspp = vlcuspp
        .Num_AnnoJub = vlNumAnnoJub
         
    End With
    
End Function

Function fgCargaEstBenGrilla(iGrilla As MSFlexGrid, istBeneficiarios() As TyBeneficiarios, iFecDevengue)
On Error GoTo Err_fgCargaEstBenGrilla
Dim vlPos, vlNumero As Integer
Const clNumPoliza As String * 10 = "0000000000"
Const clNumEndoso As String * 1 = "0"
Const clCodEstPension As String * 2 = "10"
Const clCodMotReqPen  As String * 1 = "1"
Const clCodCauSus     As String * 1 = "0"
    
If iGrilla.Rows > 1 Then
    vlPos = 1
    iGrilla.Col = 0
    vgX = 0
    vgNumBen = (iGrilla.Rows - 1)
    ReDim istBeneficiarios(vgNumBen) As TyBeneficiarios
    While vlPos <= (iGrilla.Rows - 1)
        iGrilla.Row = vlPos
        iGrilla.Col = 0
    
            vgX = vgX + 1
            With istBeneficiarios(vgX)
                 'iGrilla.Col = 11
                 .Num_Poliza = clNumPoliza
                 'iGrilla.Col = 12
                 .Num_Endoso = clNumEndoso
                 iGrilla.Col = 0
                 .Num_Orden = (iGrilla.Text)
                 iGrilla.Col = 1
                 .Cod_Par = (iGrilla.Text)
                 iGrilla.Col = 2
                 .Cod_GruFam = (iGrilla.Text)
                 iGrilla.Col = 3
                 .Cod_Sexo = (iGrilla.Text)
                 iGrilla.Col = 4
                 .Cod_SitInv = (iGrilla.Text)
                 iGrilla.Col = 5
                 '.Fec_InvBen = (iGrilla.Text)
                 .Fec_InvBen = Format(iGrilla.Text, "yyyymmdd")
                 iGrilla.Col = 6
                 .Cod_CauInv = (iGrilla.Text)
                 iGrilla.Col = 7
                 .Cod_DerPen = (iGrilla.Text)
                 iGrilla.Col = 8
                 .Cod_DerCre = (iGrilla.Text)
                 iGrilla.Col = 9
                 '.Fec_NacBen = (iGrilla.Text)
                 .Fec_NacBen = Format(iGrilla.Text, "yyyymmdd")
                 iGrilla.Col = 10
                 '.Fec_NacHM = (iGrilla.Text)
                 .Fec_NacHM = Format(iGrilla.Text, "yyyymmdd")
                 iGrilla.Col = 11
                 .Rut_Ben = fgObtenerCodigo_TextoCompuesto(iGrilla.Text)
                 iGrilla.Col = 12
                 .Dgv_Ben = Trim(iGrilla.Text)
                 iGrilla.Col = 13
                 .Gls_NomBen = (iGrilla.Text)
                 iGrilla.Col = 14
                 .Gls_NomSegBen = (iGrilla.Text)
                 iGrilla.Col = 15
                 .Gls_PatBen = (iGrilla.Text)
                 iGrilla.Col = 16
                 .Gls_MatBen = (iGrilla.Text)
                 iGrilla.Col = 17
                 .Prc_Pension = (iGrilla.Text)
                 iGrilla.Col = 18
                 .Mto_Pension = (iGrilla.Text)
                 iGrilla.Col = 19
                 .Mto_PensionGar = (iGrilla.Text)
                 iGrilla.Col = 20
                 '.Fec_FallBen = (iGrilla.Text)
                 .Fec_FallBen = Format(iGrilla.Text, "yyyymmdd")
                 'iGrilla.Col = 0
                 '.Cod_MotReqPen = clCodMotReqPen
                 iGrilla.Col = 22
                 .Cod_EstPension = (iGrilla.Text)
                 iGrilla.Col = 23
                 .Prc_PensionGar = (iGrilla.Text)
                 iGrilla.Col = 24
                 .Prc_PensionLeg = (iGrilla.Text)
                 
                 'iGrilla.Col = 23
                 '.Cod_CauSusBen = clCodCauSus
                 'iGrilla.Col = 24
                 ''.Fec_SusBen = (iGrilla.Text)
                 '.Fec_SusBen = "" 'Format(iGrilla.Text, "yyyymmdd")
                 'iGrilla.Col = 25
                 '.Fec_IniPagoPen = (iGrilla.Text)
                 '.Fec_IniPagoPen = "" 'Format(iGrilla.Text, "yyyymmdd")
                 'iGrilla.Col = 26
                 '.Fec_TerPagoPenGar = (iGrilla.Text)
                 '.Fec_TerPagoPenGar = "" 'Format(iGrilla.Text, "yyyymmdd")
                 'iGrilla.Col = 27
                 '.Fec_Matrimonio = (iGrilla.Text)
                 '.Fec_Matrimonio = "" 'Format(iGrilla.Text, "yyyymmdd")
                
            End With
            
            vlPos = vlPos + 1
    Wend
End If

Exit Function
Err_fgCargaEstBenGrilla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgObtenerPorcentaje(iParentesco As String, iInvalidez As String, iSexo As String, iFecha As String) As Boolean
'Función : Permite validar la existencia del valor del Porcentaje de Pensión a buscar
'Parámetros de Entrada:
'       - iParentesco => Código del Parentesco del Beneficiario
'       - iInvalidez  => Código de la Situación de Invalidez del Beneficiario
'       - iSexo       => Código del Sexo del Beneficiario
'       - iFecha      => Fecha con la cual se compara la Vigencia del Porcentaje (Vigencia Póliza)
'Parámetros de Salida:
'       - Retorna un Falso o True de acuerdo a su existencia
'Variables de Salida:
'       - vgValorPorcentaje => Permite guardar el Porcentaje buscado
Dim Tb_Por As ADODB.Recordset
Dim Sql    As String

    fgObtenerPorcentaje = False
    
    Sql = "select prc_pension as valor_porcentaje "
    Sql = Sql & "from MA_TVAL_PORPAR where "
    Sql = Sql & "Cod_par = '" & iParentesco & "' AND "
    Sql = Sql & "Cod_sitinv = '" & iInvalidez & "' AND "
    Sql = Sql & "Cod_sexo = '" & iSexo & "' AND "
    Sql = Sql & "fec_inivigpor <= '" & iFecha & "' AND "
    Sql = Sql & "fec_tervigpor >= '" & iFecha & "' "
    
    Set Tb_Por = vgConexionBD.Execute(Sql)
    If Not Tb_Por.EOF Then
        
        If Not IsNull(Tb_Por!Valor_Porcentaje) Then
            vgValorPorcentaje = Tb_Por!Valor_Porcentaje
            
            fgObtenerPorcentaje = True
        End If
    End If
    Tb_Por.Close
    
End Function

Function fgCalcularPorcentajeBenef(iFechaIniVig As String, iNumBenef As Integer, ostBeneficiarios() As TyBeneficiarios, Optional iTipoPension As String, Optional iPensionRef As Double, Optional iCalcularPension As Boolean, Optional iDerCrecerCotizacion As String, Optional iCobCobertura As String, Optional iCalcularPorcentaje As Boolean, Optional iPerGar As Long) As Boolean
'Función: Permite actualizar los Porcentajes de Pensión de los Beneficiarios,
'         su Derecho a Acrecer y la Fecha de Nacimiento del Hijo Menor
'Parámetros de Entrada/Salida:
'iFechaIniVig     => Fecha de Inicio de Vigencia de la Póliza
'iNumBenef        => Número de Beneficiarios
'ostBeneficiarios => Estructura desde la cual se obtienen los datos de los
'                    Beneficiarios y al mismo tiempo se calcula el Porcentaje
'                    de Pensión al cual tienen Dº
'iCalcularPension => Permite indicar si se debe realizar el cálculo del Monto de la Pensión que le corresponde a cada Beneficiario
'iTipoPension     => Tipo de Pensión de la Póliza
'iPensionRef      => Monto de la Pensión de Referencia utilizada para el Calculo de la Pensión si el campo anterior esta en Verdadero
'iDerCrecerCotizacion => Indicador de Derecho a Crecer definido en la Cotización (S o N)
'iCobCobertura    => Indicador de Cobertura de la Cotización (S o N)

Dim vlValor  As Double
Dim Numor()  As Integer
Dim Ncorbe() As Integer
Dim Codrel() As Integer
Dim Cod_Grfam() As Integer
Dim Sexobe() As String, Inv()   As String, Coinbe()   As String
Dim derpen() As Integer
Dim Nanbe()  As Integer, Nmnbe() As Integer, Ndnbe() As Integer
Dim Porben() As Double
Dim Codcbe() As String
Dim Iaap     As Integer, Immp   As Integer, Iddp    As Integer
Dim vlNum_Ben As Integer
Dim Hijos()  As Integer, Hijos_Inv() As Integer
Dim Hijos_SinDerechoPension()  As Integer 'hqr 08/07/2006
Dim Hijo_Menor() As Date
Dim Hijo_Menor_Ant() As Date
Dim Fec_NacHM() As Date

'I--- ABV 18/07/2006 ---
Dim Hijos_SinConyugeMadre()  As Integer
Dim Fec_Fall() As String
Dim vlFechaFallCau As String
Dim cont_esp_Totales As Integer
Dim cont_esp_Tot_GF() As Integer
Dim cont_mhn_Totales As Integer
Dim cont_mhn_Tot_GF() As Integer
Dim vlValorHijo  As Double
'F--- ABV 18/07/2006 ---

Dim cont_mhn() As Integer
Dim cont_causante As Integer
Dim cont_esposa As Integer
Dim cont_mhn_tot As Integer
Dim cont_hijo As Integer
Dim cont_padres As Integer

Dim L24 As Long, i As Long, edad_mes_ben As Long
Dim fecha_sin As Long, vlContBen As Long
Dim sexo_cau As String
Dim g As Long, Q As Long, X As Long, j As Long, u As String, k As Long
Dim v_hijo As Double

Dim vlFechaFallecimiento As String
Dim vlFechaMatrimonio    As String

Dim vlPorcBenef As Double
Dim vlPenBenef As Double
Dim vlPenGarBenef As Double

Dim vlSumarTotalPorcentajePension As Double, vlSumaDef As Double, vlDif As Double
Dim vlPorcentajeRecal As Double

Dim vlRemuneracion As Double, vlRemuneracionProm As Double
Dim vlRemuneracionBase As Double
Dim vlPrcCobertura As Double

        Dim vlSumaPjePenPadres As Double
'I--- ABV 21/06/2007 ---
'On Error GoTo Err_fgCalcularPorcentaje
'F--- ABV 21/06/2007 ---

'    Call flAsignaPorcentajesLegales
'    Call flValidaBeneficiarios
'    Call flDerechoAcrecer
'    Call flMadreHijoMenor
'    Call flVariosConyuges
'    Call flHijosSolos
    
    fgCalcularPorcentajeBenef = False
    L24 = 0
    'Debiera tomar la Fecha de Devengue
    
    If (fgCarga_Param("LI", "L24", iFechaIniVig) = True) Then
        L24 = vgValorParametro
    Else
        vgError = 1000
        MsgBox "No existe Edad de tope para los 24 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    
    'Mensualizar la Edad de 24 Años
    L24 = L24 * 12
'I--- ABV 21/07/2006 ---
'Edad de 18 años
Dim L18 As Long
If (fgCarga_Param("LI", "L18", iFechaIniVig) = True) Then
    L18 = vgValorParametro
Else
    vgError = 1000
    MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
    Exit Function
End If
'Mensualizar la Edad de 24 Años
L18 = L18 * 12
'F--- ABV 21/07/2006 ---

    'If Not IsDate(txt_fecha_devengo) Then
    '   X = MsgBox("Debe ingresar la Fecha de Inicio de la Pensión", 16)
    '   txt_fecha_devengo.SetFocus
    '   Exit Function
    'End If
    'If CDate(txt_fecha_devengo) > CDate(lbl_cotizacion) Then
    '    X = MsgBox("Error", "La fecha de devengamiento no puede ser mayor que la fecha de cotización", 16)
    '    Exit Function
    'End If
    'Iaap = CInt(Year(CDate(txt_fecha_devengo))) 'a Fecha de siniestro
    'Immp = CInt(Month(CDate(txt_fecha_devengo))) 'm Fecha de siniestro
    'Iddp = CInt(Day(CDate(txt_fecha_devengo))) 'd Fecha de siniestro
    
    Iaap = CInt(Mid(iFechaIniVig, 1, 4)) 'a Fecha de siniestro
    Immp = CInt(Mid(iFechaIniVig, 5, 2)) 'm Fecha de siniestro
    Iddp = CInt(Mid(iFechaIniVig, 7, 2)) 'd Fecha de siniestro
    
    fecha_sin = Iaap * 12 + Immp
    vlNum_Ben = iNumBenef 'grd_beneficiarios.Rows - 1 '.Rows - 1
    
    ReDim Numor(vlNum_Ben) As Integer
    ReDim Ncorbe(vlNum_Ben) As Integer
    ReDim Codrel(vlNum_Ben) As Integer
    ReDim Cod_Grfam(vlNum_Ben) As Integer
    ReDim Sexobe(vlNum_Ben) As String
    ReDim Inv(vlNum_Ben) As String
    ReDim Coinbe(vlNum_Ben) As String
    ReDim derpen(vlNum_Ben) As Integer
    ReDim Nanbe(vlNum_Ben) As Integer
    ReDim Nmnbe(vlNum_Ben) As Integer
    ReDim Ndnbe(vlNum_Ben) As Integer
    ReDim Hijos(vlNum_Ben) As Integer
    ReDim Hijos_Inv(vlNum_Ben) As Integer
    ReDim Hijos_SinDerechoPension(vlNum_Ben) As Integer
    ReDim Hijo_Menor(vlNum_Ben) As Date
    ReDim Hijo_Menor_Ant(vlNum_Ben) As Date
    ReDim Porben(vlNum_Ben) As Double
    ReDim Codcbe(vlNum_Ben) As String
    ReDim Fec_NacHM(vlNum_Ben) As Date
    ReDim cont_mhn(vlNum_Ben) As Integer
    
'I--- ABV 18/07/2006 ---
    ReDim Hijos_SinConyugeMadre(vlNum_Ben) As Integer
    ReDim Fec_Fall(vlNum_Ben) As String
    ReDim cont_mhn_Tot_GF(vlNum_Ben) As Integer
    ReDim cont_esp_Tot_GF(vlNum_Ben) As Integer
    vlFechaFallCau = ""
    vlValorHijo = 0
'F--- ABV 18/07/2006 ---
    
If (iCalcularPorcentaje = True) Then
    vlContBen = 1 '0
    i = 1
    Do While i <= vlNum_Ben
    
        vlContBen = vlContBen + 1
        'Nº Orden
        
        'Msf_GrillaBenef.Row = i
        'Msf_GrillaBenef.Col = 0
        'Numor(i) = Msf_GrillaBenef.Text ''tb_ben!cod_numordben 'N° de orden  NUMOR(I)
        
        'If Trim(grd_beneficiarios.TextMatrix(i, 0)) = "" Then
        If Trim(ostBeneficiarios(i).Num_Orden) = "" Then
            vgError = 1000
            MsgBox "No existe Número de Orden de Beneficiario.", vbCritical, "Error de Datos"
            Exit Do
        End If
        
        'Número de Orden
        Numor(i) = ostBeneficiarios(i).Num_Orden
        
        'Parentesco
        Ncorbe(i) = ostBeneficiarios(i).Cod_Par ''tb_ben!cod_par
        Codrel(i) = ostBeneficiarios(i).Cod_Par ''tb_ben!cod_par
        
        'Grupo Familiar
        Cod_Grfam(i) = ostBeneficiarios(i).Cod_GruFam
        
        'Sexo
        Sexobe(i) = ostBeneficiarios(i).Cod_Sexo
        If (Ncorbe(i) = "99") Then
            sexo_cau = Sexobe(i)
        End If
        
        'Situación de Invalidez
        Inv(i) = ostBeneficiarios(i).Cod_SitInv
        
        'Derecho Pensión
        derpen(i) = ostBeneficiarios(i).Cod_EstPension
        
        'Fecha de Nacimiento
        Nanbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 1, 4)) 'a Fecha de nacimiento
        Nmnbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 5, 2)) 'm Fecha de nacimiento
        Ndnbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 7, 2)) 'd Fecha de nacimiento
                    
        'Fecha nacimiento hijo menor =IJAM(I),IJMN(I),IJDN(I)
        
        'Codificación de Situación de Invalidez
        If Inv(i) = "P" Then Coinbe(i) = "P"
        If Inv(i) = "T" Then Coinbe(i) = "T"
        If Inv(i) = "N" Then Coinbe(i) = "N"
        
        '*********
        edad_mes_ben = fecha_sin - (Nanbe(i) * 12 + Nmnbe(i))
        
        vlFechaFallecimiento = ostBeneficiarios(i).Fec_FallBen
'I--- ABV 18/07/2006 ---
        Fec_Fall(i) = vlFechaFallecimiento
        If Codrel(i) = 99 Then
            vlFechaFallCau = vlFechaFallecimiento
        End If
'F--- ABV 18/07/2006 ---

        vlFechaMatrimonio = ostBeneficiarios(i).Fec_Matrimonio
        
        'I--- ABV 16/04/2005 ---
        If (Codrel(i) >= 30 And Codrel(i) < 40) Then
            If (vlFechaFallecimiento <> "") Or (vlFechaMatrimonio <> "") Then
                derpen(i) = 10
                If ostBeneficiarios(i).Cod_EstPension <> "10" Then
                    ostBeneficiarios(i).Cod_EstPension = "10"
                    Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
                End If
            Else
                If edad_mes_ben > L24 And Coinbe(i) = "N" Then
                    derpen(i) = 10
                    'hqr 08/07/2006 cambiar Estado de pension a No Vigente.  Se debe cambiar estado de madre si es hijo unico
                    If ostBeneficiarios(i).Cod_EstPension <> "10" Then
                        ostBeneficiarios(i).Cod_EstPension = "10"
                        Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
                    End If
                Else
'I--- ABV 21/07/2006 ---
'Validación Especial para la Cía Chilena. Si ellos indican que no se encuentra
'vigente cuando tiene menos de 24 años, se debe modificar
                    derpen(i) = 99
'                    If ostBeneficiarios(i).Cod_DerPen = "10" Then
'                        If edad_mes_ben > L18 Then
'                            Derpen(i) = 10
'                        End If
'                    End If
'F--- ABV 21/07/2006 ---
                End If
            End If
        Else
            'El resto de los Beneficiarios que no sean Hijos, solo se dejan como
            'Sin Derecho a Pensión cuando están fallecidos
            If (vlFechaFallecimiento <> "") Then
                derpen(i) = 10
'I--- ABV 18/07/2006 ---
'                If (Codrel(i) = "11") Or (Codrel(i) = "21") Then
'                    If ostBeneficiarios(i).Cod_EstPension <> "10" Then
'                        ostBeneficiarios(i).Cod_EstPension = "10"
'                        Hijos_SinConyugeMadre(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinConyugeMadre(ostBeneficiarios(i).Cod_GruFam) + 1
'                    End If
'                End If
'F--- ABV 18/07/2006 ---
            Else
                derpen(i) = 99
            End If
        End If
        ostBeneficiarios(i).Cod_DerPen = derpen(i)
        'F--- ABV 16/04/2005 ---
        
        i = i + 1
    Loop
    
    cont_causante = 0
    cont_esposa = 0
    'cont_mhn = 0
    cont_mhn_tot = 0
    cont_hijo = 0
    cont_padres = 0

'I--- ABV 18/07/2006 ---
    cont_esp_Totales = 0
    cont_mhn_Totales = 0
'F--- ABV 18/07/2006 ---
    
    'Primer Ciclo
    For g = 1 To vlNum_Ben
        If derpen(g) <> 10 Then
            '99: con derecho a pension,20: con Derecho Pendiente
            '10: sin derecho a pension
            If Ncorbe(g) = 99 Then
                cont_causante = cont_causante + 1
            Else
                Select Case Ncorbe(g)
                    Case 10, 11
                        cont_esposa = cont_esposa + 1
                    Case 20, 21
                        Q = Cod_Grfam(g)
                        cont_mhn(Q) = cont_mhn(Q) + 1
                        cont_mhn_tot = cont_mhn_tot + 1
                    Case 30
                        'edad = fgEdadBen(vg_fecsin, vgBen(g).fec_nacben)
                        Q = Cod_Grfam(g)
                        Hijos(Q) = Hijos(Q) + 1
                        If Coinbe(g) <> "N" Then Hijos_Inv(Q) = Hijos_Inv(Q) + 1
                        Hijo_Menor(Q) = DateSerial(Nanbe(g), Nmnbe(g), Ndnbe(g))
                        If Hijos(Q) > 1 Then
                            If Hijo_Menor(Q) > Hijo_Menor_Ant(Q) Then
                                Hijo_Menor_Ant(Q) = Hijo_Menor(Q)
                            End If
                        Else
                            Hijo_Menor_Ant(Q) = Hijo_Menor(Q)
                        End If
                        'Case 35
                        edad_mes_ben = fecha_sin - (Nanbe(g) * 12 + Nmnbe(g))
                        If Coinbe(g) = "N" And edad_mes_ben <= L24 Then
                            cont_hijo = cont_hijo + 1
                        Else
                            If Coinbe(g) = "T" Or Coinbe(g) = "P" Then
                                cont_hijo = cont_hijo + 1
                            End If
                        End If
                    Case 41, 42
                        cont_padres = cont_padres + 1
                    Case Else
                        vgError = 1000
                        X = MsgBox("Error en codificación de codigo de relación", vbCritical)
                        Exit Function
                End Select
            End If
        Else
            'Verificar si la cónyuge o la Madre Falleció antes que el Causante para contarla
            Select Case Ncorbe(g)
                Case 11:
                    If (vlFechaFallCau > Fec_Fall(g)) Then
                        cont_esposa = cont_esposa + 1
                        Q = Cod_Grfam(g)
                        cont_esp_Tot_GF(Q) = cont_esp_Tot_GF(Q) + 1
                        cont_esp_Totales = cont_esp_Totales + 1
                    End If
                Case 21:
'                    If (vlFechaFallCau > Fec_Fall(g)) Then
'                        Q = Cod_Grfam(g)
'                        cont_mhn_Tot_GF(Q) = cont_mhn_Tot_GF(Q) + 1
'                        cont_mhn_Totales = cont_mhn_Totales + 1
''                    Else
''                        Q = Cod_Grfam(g)
''                        cont_mhn_Tot_GF(Q) = cont_mhn_Tot_GF(Q) - 1
''                        cont_mhn_Totales = cont_mhn_Totales - 1
'                    End If
            End Select
        End If
    Next g
                
''I--- ABV 18/07/2006 ---
''Corregir el Parentesco de los Hijos cuando la Cónyuge o la Madre haya Muerto antes que el Causante
'    For j = 1 To vlNum_Ben
'        '99: con derecho a pension,20: con Derecho Pendiente
'        '10: sin derecho a pension
'        If derpen(j) <> 10 Then
''            edad_mes_ben = fecha_sin - (Nanbe(j) * 12 + Nmnbe(j))
'            Select Case Ncorbe(j)
'
'                Case 30
'                    Q = Cod_Grfam(j)
'                    'Cuando No existe Conyuge y tampoco no Existe MHN para el Hijo a Analizar
'                    'se deben modificar los Códigos de los Hijos a 35
''I--- ABV 20/07/2006 ---
''                    If cont_esp_Tot_GF(Q) <= 0 And cont_mhn_Tot_GF(Q) <= 0 Then
'                    If cont_esposa <= 0 And cont_mhn(Q) <= 0 Then
''F--- ABV 20/07/2006 ---
'                        Ncorbe(j) = 35
'                        Hijos(Q) = Hijos(Q) - 1
'                        cont_hijo = cont_hijo + 1
'                        If ostBeneficiarios(j).Cod_Par <> "35" Then
'                            ostBeneficiarios(j).Cod_Par = "35"
'                            'Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) = Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) + 1
'                        End If
'                    End If
'                Case 35
'                    Q = Cod_Grfam(j)
'                    'Cuando existe Conyuge o MHN, considerándola como existente
'                    'cuando ésta muere antes del Causante de la Póliza,
'                    'se deben modificar los Códigos de los Hijos a 30
''I--- ABV 20/07/2006 ---
''                    If cont_esp_Tot_GF(Q) > 0 Or cont_mhn_Tot_GF(Q) > 0 Then
'                    If cont_esposa > 0 Or cont_mhn(Q) > 0 Then
''F--- ABV 20/07/2006 ---
'                        Ncorbe(j) = 30
'                        Hijos(Q) = Hijos(Q) + 1
'                        cont_hijo = cont_hijo - 1
'                        If ostBeneficiarios(j).Cod_Par <> "30" Then
'                            ostBeneficiarios(j).Cod_Par = "30"
'                            'Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) = Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) + 1
'                        End If
'                    End If
'            End Select
'        End If
'    Next j
''F--- ABV 18/07/2006 ---
                
    j = 1
    For j = 1 To vlNum_Ben
        '99: con derecho a pension,20: con Derecho Pendiente
        '10: sin derecho a pension
        If derpen(j) <> 10 Then
            edad_mes_ben = fecha_sin - (Nanbe(j) * 12 + Nmnbe(j))
            Select Case Ncorbe(j)
                Case 99
                    If cont_causante > 1 Then
                        vgError = 1000
                        X = MsgBox("Error en codificación de codigo de relación, No puede ingresar otro causante", vbCritical)
                        Exit Function
                    End If
                    'I--- ABV 25/02/2005 ---
                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                        vlValor = vgValorPorcentaje
                    Else
                        vgError = 1000
                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                        Exit Function
                    End If
                    'F--- ABV 25/02/2005 ---
                    If (vlValor < 0) Then
                        vgError = 1000
                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                        'error
                        Exit Function
                    Else
                        Porben(j) = vlValor
                        'Porben(j) = 100
                        Codcbe(j) = "N"
                    End If
                Case 10, 11
                    If sexo_cau = "M" Then
                        If Sexobe(j) <> "F" Then
                            vgError = 1000
                            X = MsgBox("Error de código de sexo, el Sexo de la Cónyuge debe ser Femenino.", vbCritical)
                            Exit Function
                        End If
                    Else
                        If Sexobe(j) <> "M" Then
                            vgError = 1000
                            X = MsgBox("Error de codigo de sexo, el Sexo del Cónyuge debe ser Masculino.", vbCritical)
                            Exit Function
                        End If
                    End If
                    'hqr 08/07/2006 validacion y cambio parentesco
                    u = Cod_Grfam(j)
                    If Hijos(u) = 0 And Ncorbe(j) = 11 Then
                        If Hijos_SinDerechoPension(u) = 0 Then
                            vgError = 1000
                            X = MsgBox("Error de código de relación, 'Cónyuge Con Hijos con Dº Pensión', no tiene Hijos.", vbCritical)
                            Exit Function
                        Else
                            Ncorbe(j) = 10
                            ostBeneficiarios(j).Cod_Par = 10
                        End If
                    End If
                    
                    'HQR 08/07/2006 se deja al final porque se cambia el tipo de parentesco
                    'I--- ABV 25/02/2005 ---
                     'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                     If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                         vlValor = vgValorPorcentaje
                     Else
                         vgError = 1000
                         MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                         & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                         Exit Function
                     End If
                     'F--- ABV 25/02/2005 ---
                     If (vlValor < 0) Then
                         vgError = 1000
                         MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                         'error
                         Exit Function
                     End If

                    'fin hqr 08/07/2006
                    If sexo_cau = "M" Or sexo_cau = "F" Then

                        'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                        'If (vlValor < 0) Then
                        '    Exit Sub
                        'Else
                            Porben(j) = CDbl(Format(vlValor / cont_esposa, "#0.00"))
                        'End If
                                         
                        If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
                            If Hijos(u) > 0 Then
                                Codcbe(j) = "S"
                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
                            End If
                        Else
                            Codcbe(j) = "N"
                        End If
                            
                        If Hijos(u) > 0 And Ncorbe(j) = 10 Then
                            vgError = 1000
                            X = MsgBox("Error de código de relación, 'Cónyuge Sin Hijos con Dº Pensión', tiene Hijos.", vbCritical)
                            Exit Function
                        End If
'                    Else
'                        u = Cod_Grfam(j)
'                        If Hijos(u) > 0 Then
'                            'I--- ABV 25/02/2005 ---
'                            'If Coinbe(j) = "T" Then
'                            '    Porben(j) = vlValor     '50
'                            'Else
'                            '    'HQR 16-06-2004
'                            '    'Porben(j) = 36
'                            '    If Coinbe(j) = "P" Then
'                            '        Porben(j) = 36
'                            '    Else
'                            '        Porben(j) = 0
'                            '        cont_esposa = cont_esposa - 1
'                            '    End If
'                            '    'FIN HQR 16-06-2004
'                            'End If
'
'                            Porben(j) = vlValor
'                            If (Coinbe(j) = "N") Then
'                                cont_esposa = cont_esposa - 1
'                            End If
'                            'F--- ABV 25/02/2005 ---
'
'                            If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
'                                Codcbe(j) = "S"
'                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
'                            Else
'                                Codcbe(j) = "N"
'                            End If
'                        Else
'                            'I--- ABV 25/02/2005 ---
'                            'If Coinbe(j) = "T" Then
'                            '    Porben(j) = vlValor     '60
'                            'Else
'                            '    'HQR 16-06-2004
'                            '    'Porben(j) = 43
'                            '    If Coinbe(j) = "P" Then
'                            '        Porben(j) = 43
'                            '    Else
'                            '        Porben(j) = 0
'                            '        cont_esposa = cont_esposa - 1
'                            '    End If
'                            '    'FIN HQR 16-06-2004
'                            'End If
'
'                            Porben(j) = vlValor
'                            If (Coinbe(j) = "N") Then
'                                cont_esposa = cont_esposa - 1
'                            End If
'                            'F--- ABV 25/02/2005 ---
'
'                            If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
'                                Codcbe(j) = "S"
'                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
'                            Else
'                                Codcbe(j) = "N"
'                            End If
'                        End If
                    End If
                Case 20, 21

                    If sexo_cau = "M" Then
                        If Sexobe(j) <> "F" Then
                            vgError = 1000
                            X = MsgBox("Error de código de sexo, el Sexo de la Cónyuge debe ser Femenino.", vbCritical)
                            Exit Function
                        End If
                    Else
                        If Sexobe(j) <> "M" Then
                            vgError = 1000
                            X = MsgBox("Error de codigo de sexo, el Sexo del Cónyuge debe ser Masculino.", vbCritical)
                            Exit Function
                        End If
                    End If
                    
                    u = Cod_Grfam(j)
                    If Hijos(u) = 0 And Ncorbe(j) = 21 Then
                        If Hijos_SinDerechoPension(u) = 0 Then
                            vgError = 1000
                            X = MsgBox("Error en código de relación 'Madre Con Hijos con Dº Pensión, no tiene Hijos.", vbCritical)
                            Exit Function
                        Else
                            Ncorbe(j) = 20
                            ostBeneficiarios(j).Cod_Par = 20
                        End If
                    End If
                    
                    'I--- ABV 25/02/2005 ---
                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                        vlValor = vgValorPorcentaje
                    Else
                        vgError = 1000
                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                        Exit Function
                    End If
                    'F--- ABV 25/02/2005 ---
                    If (vlValor < 0) Then
                        vgError = 1000
                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                        Exit Function
                    Else
                        Porben(j) = vlValor / cont_mhn_tot
                    End If

                    If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
                        If Hijos(u) > 0 Then
                            Codcbe(j) = "S"
                        Else
                            Codcbe(j) = "N"
                        End If
                        If Hijos_Inv(u) > 0 Then
                            Codcbe(j) = "N"
                        End If
                    Else
                        Codcbe(j) = "N"
                    End If

                    If Hijos(u) > 0 And Ncorbe(j) = 20 Then
                        vgError = 1000
                        X = MsgBox("Error en código de relación 'Madre Sin Hijos con Dº Pensión', tiene Hijos.", vbCritical)
                        Exit Function
                    End If
                    
                Case 30
                    Codcbe(j) = "N"
                    Q = Cod_Grfam(j)
                    If cont_esposa > 0 Or cont_mhn(Q) > 0 Then
                        If Coinbe(j) = "N" And edad_mes_ben > L24 Then
                            Porben(j) = 0
                        Else
                            If (Coinbe(j) = "P" Or Coinbe(j) = "T") And edad_mes_ben > L24 Then
                                'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                                    vlValor = vgValorPorcentaje
                                Else
                                    vgError = 1000
                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                    & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                    Exit Function
                                End If
                                
                                If (vlValor < 0) Then
                                    vgError = 1000
                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                    'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
                                    Exit Function
                                Else
                                    Porben(j) = vlValor
                                End If
                                'F--- ABV 25/02/2005 ---
                            
                            Else
                                
                                'I--- ABV 25/02/2005 ---
                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                    vlValor = vgValorPorcentaje
                                Else
                                    vgError = 1000
                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                    & Ncorbe(j) & " - N - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                    Exit Function
                                End If
                                'F--- ABV 25/02/2005 ---
                                
                                If (vlValor < 0) Then
                                    vgError = 1000
                                    'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                    Exit Function
                                Else
                                    'Porben(j) = 15
                                    Porben(j) = vlValor
                                End If
                            End If
                        End If
                    Else
                        Q = Cod_Grfam(j)
                        Codcbe(j) = "N"

                        If cont_esposa = 0 And cont_mhn(Q) = 0 Then
                            If (Coinbe(j) = "P" Or Coinbe(j) = "T") And edad_mes_ben > L24 Then
                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                    vlValor = vgValorPorcentaje
                                Else
                                    vgError = 1000
                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                    & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                    Exit Function
                                End If
                                
                                If vlValor < 0 Then
                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                    Exit Function
                                Else
                                    Porben(j) = vlValor
                                End If

                            Else
                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                    vlValor = vgValorPorcentaje
                                Else
                                    vgError = 1000
                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                    & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                    Exit Function
                                End If
                                
                                If vlValor < 0 Then
                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                    Exit Function
                                Else
                                    vlValorHijo = vlValor
                                    Porben(j) = vlValor
                                End If
                            End If
                            
                            'Buscar el Porcentaje de una Cónyuge
                            If (fgObtenerPorcentaje("10", "N", "F", iFechaIniVig) = True) Then
                                'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                vlValor = vgValorPorcentaje
                            Else
                                vgError = 1000
                                MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                & "10" & " - " & "N" & " - " & "F" & " - " & iFechaIniVig & "."
                                Exit Function
                            End If
                            
                            If vlValor < 0 Then
                                MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                Exit Function
                            Else
                                v_hijo = vlValor
                                If Coinbe(j) = "N" And edad_mes_ben <= L24 Then
                                    If cont_hijo = 1 Then
                                        Porben(j) = v_hijo
                                    Else
                                        Porben(j) = v_hijo / cont_hijo + vlValorHijo
                                    End If
                                Else
                                    If Coinbe(j) = "T" Or Coinbe(j) = "N" Then
                                        If cont_hijo = 1 Then
                                            Porben(j) = v_hijo
                                        Else
                                            Porben(j) = v_hijo / cont_hijo + vlValorHijo
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case 41, 42
                    'I--- ABV 25/02/2005 ---
                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                        'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                        vlValor = vgValorPorcentaje
                    Else
                        vgError = 1000
                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                        Exit Function
                    End If
                    'F--- ABV 25/02/2005 ---
                    
                    If (vlValor < 0) Then
                        vgError = 1000
                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                        Exit Function
                    Else
                        Codcbe(j) = "N"
                        Porben(j) = vlValor
                    End If

            End Select
        End If
    Next j
    
    For k = 1 To vlNum_Ben '60
        If derpen(k) <> 10 Then
            Select Case Ncorbe(k)
                Case 11
                    Q = Cod_Grfam(k)
                    'If codcbe(k) = "S" Then DAJ 10/08/2007
                        Fec_NacHM(k) = Hijo_Menor_Ant(Q)  '******Agruegué 05/11/2000 HILDA
                    'End If DAJ 10/08/2007
                Case 21
                    Q = Cod_Grfam(k)
                    'If codcbe(k) = "S" Then DAJ 10/08/2007
                        Fec_NacHM(k) = Hijo_Menor_Ant(Q)  '*******Agruegué 05/11/2000 HILDA
                    'End If DAJ 10/08/2007
            End Select
        End If
    Next k
    
    vlSumarTotalPorcentajePension = 0
    vlSumaPjePenPadres = 0
    j = 0
    For j = 1 To vlNum_Ben
        'Guardar el Valor del Porcentaje Calculado
        'If IsNumeric(Porben(j)) Then
        '    grd_beneficiarios.Text = Format(Porben(j), "##0.00")
        'Else
        '    grd_beneficiarios.Text = Format("0", "0.00")
        'End If
        If IsNumeric(Porben(j)) Then
            'I--- ABV 22/06/2005 ---
            'ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
            If (ostBeneficiarios(j).Cod_Par = "99" Or ostBeneficiarios(j).Cod_Par = "0") Then
                If (derpen(j) <> 10) Then
                    ostBeneficiarios(j).Prc_Pension = 100
                    ostBeneficiarios(j).Prc_PensionLeg = 100 '*-+
                Else
                    ostBeneficiarios(j).Prc_Pension = 0
                    ostBeneficiarios(j).Prc_PensionLeg = 0 '*-+
                End If
            Else
                ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
                ostBeneficiarios(j).Prc_PensionLeg = Format(Porben(j), "#0.00") '*-+
            End If
            'F--- ABV 22/06/2005 ---
        Else
            ostBeneficiarios(j).Prc_Pension = 0
            ostBeneficiarios(j).Prc_PensionLeg = 0 '*-+
        End If
        
        'Guardar el Derecho a Acrecer de los Beneficiarios
        'Inicio
        'If (Codcbe(j) <> Empty And Not IsNull(Codcbe(j))) Then
        '    grd_beneficiarios.Text = Codcbe(j)
        'Else
        '    'Por Defecto Negar el Derecho a Acrecer de los Beneficiarios
        '    grd_beneficiarios.Text = "N"
        'End If
        If (Codcbe(j) <> Empty And Not IsNull(Codcbe(j))) Then
            ostBeneficiarios(j).Cod_DerCre = Codcbe(j)
        Else
            ostBeneficiarios(j).Cod_DerCre = "N"
        End If
        'Fin
        
        'Guardar la Fecha de Nacimiento del Hijo Menor de la Cónyuge
        'Inicio
        If Format(CDate(Fec_NacHM(j)), "yyyymmdd") > "18991230" Then
            'ostBeneficiarios(j).Fec_NacHM = CDate(Fec_NacHM(j))
            ostBeneficiarios(j).Fec_NacHM = Format(CDate(Fec_NacHM(j)), "yyyymmdd")
        Else
            ostBeneficiarios(j).Fec_NacHM = ""
        End If
        'Fin

'I--- ABV 10/08/2007 ---
        'Inicializar el Monto de Pensión a Cero
        '--------------------------------------
        'Actualizar el Monto de la Pensión
        ostBeneficiarios(j).Mto_Pension = Format(0, "#0.00")
        
        'Actualizar el Monto de la Pensión Garantizada
        ostBeneficiarios(j).Mto_PensionGar = Format(0, "#0.00")
'F--- ABV 10/08/2007 ---
        
        If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
            vlSumarTotalPorcentajePension = vlSumarTotalPorcentajePension + ostBeneficiarios(j).Prc_Pension
        End If
'F--- ABV 21/06/2006 ---
    Next j
    
    'Validar si los Porcentajes de Pensión Suman más del 100%
    'Validacion de existencia de padres y eliminar o disminuir su % segun corresponda
    If cont_padres > 0 And vlSumarTotalPorcentajePension > 100 Then
        If vlSumarTotalPorcentajePension - vlSumaPjePenPadres <= 100 Then
            For j = 1 To vlNum_Ben
                If (ostBeneficiarios(j).Cod_DerPen <> "10") And ((ostBeneficiarios(j).Cod_Par = "41" Or ostBeneficiarios(j).Cod_Par = "42") And ostBeneficiarios(j).Cod_Par <> "0") Then
                    vlPorcentajeRecal = (100 - (vlSumarTotalPorcentajePension - vlSumaPjePenPadres)) / cont_padres
                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                End If
            Next j
        Else
            For j = 1 To vlNum_Ben
                If (ostBeneficiarios(j).Cod_DerPen <> "10") And ((ostBeneficiarios(j).Cod_Par = "41" Or ostBeneficiarios(j).Cod_Par = "42") And ostBeneficiarios(j).Cod_Par <> "0") Then
                    vlPorcentajeRecal = 0
                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                    ostBeneficiarios(j).Cod_DerPen = "10"
                End If
            Next j

            vlSumaDef = 0
            For j = 1 To vlNum_Ben
                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                    vlPorcentajeRecal = (ostBeneficiarios(j).Prc_PensionLeg / (vlSumarTotalPorcentajePension - vlSumaPjePenPadres)) * 100
                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                    vlSumaDef = vlSumaDef + ostBeneficiarios(j).Prc_Pension
                End If
            Next j
            vlDif = Format(100 - vlSumaDef, "#0.00")
            If (vlDif < 0) Or (vlDif > 0) Then
                For j = 1 To vlNum_Ben
                    If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                        vlPorcentajeRecal = ostBeneficiarios(j).Prc_Pension + vlDif
                        ostBeneficiarios(j).Prc_Pension = vlPorcentajeRecal
                        Exit For
                    End If
                Next
            End If
        End If
    Else
        If (vlSumarTotalPorcentajePension > 100) Then
            vlSumaDef = 0
            For j = 1 To vlNum_Ben
                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                    vlPorcentajeRecal = (ostBeneficiarios(j).Prc_Pension / vlSumarTotalPorcentajePension) * 100
                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                    vlSumaDef = vlSumaDef + ostBeneficiarios(j).Prc_Pension
                End If
            Next j
            vlDif = Format(100 - vlSumaDef, "#0.00")
            If (vlDif < 0) Or (vlDif > 0) Then
                'Asignar la diferencia de Porcentaje al Primer Beneficiario con Derecho distinto del Causante
                For j = 1 To vlNum_Ben
                    If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                        vlPorcentajeRecal = ostBeneficiarios(j).Prc_Pension + vlDif
                        ostBeneficiarios(j).Prc_Pension = vlPorcentajeRecal
                        Exit For
                    End If
                Next j
            End If
        End If
    End If

'*-+
'I--- ABV 04/08/2007 ---
'Recalcular Porcentajes si se trata de un Caso de Invalidez - Con Cobertura
    If (iTipoPension = "06" Or iTipoPension = "07") And (iCobCobertura = "S") Then
        If (fgObtenerPorcCobertura(iTipoPension, iFechaIniVig, vlRemuneracion, vlPrcCobertura) = True) Then
            'Determinar la Remuneración Promedio para el Causante
            vlRemuneracionBase = vlRemuneracion * (vlPrcCobertura / 100)
            
            'Determinar para cada Beneficiario el Nuevo Porcentaje
            For j = 1 To (vlContBen - 1)
                If (ostBeneficiarios(j).Cod_DerPen <> "10") Then
                    If (ostBeneficiarios(j).Cod_Par = "99") Then
                        vlRemuneracionProm = vlRemuneracion * (vlPrcCobertura / 100)
                    Else
                        vlRemuneracionProm = vlRemuneracion * (ostBeneficiarios(j).Prc_Pension / 100)
                    End If
                    vlPorcentajeRecal = (vlRemuneracionProm / vlRemuneracionBase) * 100
                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                End If
            Next j
        Else
            vgError = 1000
            MsgBox "No existen valores de Porcentaje de Cobertura para el Recalculo de Pensiones.", vbCritical, "Inexistencia de Datos"
            Exit Function
        End If
    End If
'F--- ABV 04/08/2007 ---
'*-+
End If

'Recalcular los Montos de Pensión
If (iCalcularPension = True) Then
    For j = 1 To (vlNum_Ben)
            
        vlPorcBenef = 0
        vlPenBenef = 0
        vlPenGarBenef = 0
                
        'Verificar el Estado de Derecho a Pensión (Cod_EstPension)
'            If (ostBeneficiarios(j).Cod_EstPension <> "10") Then
        If (ostBeneficiarios(j).Cod_DerPen <> "10") Then
            vlPorcBenef = ostBeneficiarios(j).Prc_Pension
            
            'Determinar el monto de la Pensión del Beneficiario
            'Penben(i) = IIf(Msf_BMGrilla.TextMatrix(vlI, 17) = "", 0, Msf_BMGrilla.TextMatrix(vlI, 17))
            vlPenBenef = iPensionRef * (vlPorcBenef / 100)
        
            'Definir la Pensión del Garantizado
            vlPenGarBenef = ostBeneficiarios(j).Mto_PensionGar
        
            'I--- ABV 20/04/2005 ---
            vgPalabra = iTipoPension
            vgI = InStr(1, cgPensionInvVejez, vgPalabra)
            If (vgI <> 0) Then
                'Valida que sea un caso de Invalidez o Vejez
                vlPenGarBenef = vlPenBenef
            End If
            
            vgI = InStr(1, cgPensionSobOrigen, vgPalabra)
            If (vgI <> 0) Then
                'Valida que sea un caso de Sobrevivencia por Origen
                'Se supone que la Cónyuge o Madre es la única que Garantiza su pensión
                'vlNumero = InStr(Cmb_BMPar.Text, "-")
                vgPalabraAux = ostBeneficiarios(j).Cod_Par
                
                vgX = InStr(1, cgParConyugeMadre, vgPalabraAux)
                If (vgX <> 0) Then
                    vlPenGarBenef = vlPenBenef
                'Else
                '    Txt_BMMtoPensionGar.Enabled = True
                End If
            End If
            
            vgI = InStr(1, cgPensionSobTransf, vgPalabra)
            If (vgI <> 0) Then
                ''Valida que sea un caso de Invalidez o Vejez
                'Txt_BMMtoPensionGar.Enabled = True
            End If

            'F--- ABV 19/04/2005 ---
        End If
        
        'Actualizar el Monto de la Pensión
        ostBeneficiarios(j).Mto_Pension = Format(vlPenBenef, "#0.00")
        
        'Actualizar el Monto de la Pensión Garantizada
        If (iPerGar > 0) Then
            ostBeneficiarios(j).Mto_PensionGar = Format(vlPenGarBenef, "#0.00")
            ostBeneficiarios(j).Prc_PensionGar = vlPorcBenef
        Else
            ostBeneficiarios(j).Mto_PensionGar = 0
            ostBeneficiarios(j).Prc_PensionGar = 0
        End If
        
    Next j
End If

    fgCalcularPorcentajeBenef = True
    
Exit Function
Err_fgCalcularPorcentaje:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        vgError = 1000
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgCalcularPorcentajeBenef_Old(iFechaIniVig As String, iNumBenef As Integer, ostBeneficiarios() As TyBeneficiarios, Optional iTipoPension As String, Optional iPensionRef As Double, Optional iCalcularPension As Boolean, Optional iDerCrecerCotizacion As String, Optional iCobCobertura As String, Optional iCalcularPorcentaje As Boolean, Optional iPerGar As Long) As Boolean
'Función: Permite actualizar los Porcentajes de Pensión de los Beneficiarios,
'         su Derecho a Acrecer y la Fecha de Nacimiento del Hijo Menor
'Parámetros de Entrada/Salida:
'iFechaIniVig     => Fecha de Inicio de Vigencia de la Póliza
'iNumBenef        => Número de Beneficiarios
'ostBeneficiarios => Estructura desde la cual se obtienen los datos de los
'                    Beneficiarios y al mismo tiempo se calcula el Porcentaje
'                    de Pensión al cual tienen Dº
'iCalcularPension => Permite indicar si se debe realizar el cálculo del Monto de la Pensión que le corresponde a cada Beneficiario
'iTipoPension     => Tipo de Pensión de la Póliza
'iPensionRef      => Monto de la Pensión de Referencia utilizada para el Calculo de la Pensión si el campo anterior esta en Verdadero
'iDerCrecerCotizacion => Indicador de Derecho a Crecer definido en la Cotización (S o N)
'iCobCobertura    => Indicador de Cobertura de la Cotización (S o N)

Dim vlValor  As Double
Dim Numor()  As Integer
Dim Ncorbe() As Integer
Dim Codrel() As Integer
Dim Cod_Grfam() As Integer
Dim Sexobe() As String, Inv()   As String, Coinbe()   As String
Dim derpen() As Integer
Dim Nanbe()  As Integer, Nmnbe() As Integer, Ndnbe() As Integer
Dim Porben() As Double
Dim Codcbe() As String
Dim Iaap     As Integer, Immp   As Integer, Iddp    As Integer
Dim vlNum_Ben As Integer
Dim Hijos()  As Integer, Hijos_Inv() As Integer
Dim Hijos_SinDerechoPension()  As Integer 'hqr 08/07/2006
Dim Hijo_Menor() As Date
Dim Hijo_Menor_Ant() As Date
Dim Fec_NacHM() As Date

'I--- ABV 18/07/2006 ---
Dim Hijos_SinConyugeMadre()  As Integer
Dim Fec_Fall() As String
Dim vlFechaFallCau As String
Dim cont_esp_Totales As Integer
Dim cont_esp_Tot_GF() As Integer
Dim cont_mhn_Totales As Integer
Dim cont_mhn_Tot_GF() As Integer
Dim vlValorHijo  As Double
'F--- ABV 18/07/2006 ---

Dim cont_mhn() As Integer
Dim cont_causante As Integer
Dim cont_esposa As Integer
Dim cont_mhn_tot As Integer
Dim cont_hijo As Integer
Dim cont_padres As Integer

Dim L24 As Long, i As Long, edad_mes_ben As Long
Dim fecha_sin As Long, vlContBen As Long
Dim sexo_cau As String
Dim g As Long, Q As Long, X As Long, j As Long, u As String, k As Long
Dim v_hijo As Double

Dim vlFechaFallecimiento As String
Dim vlFechaMatrimonio    As String

Dim vlPorcBenef As Double
Dim vlPenBenef As Double
Dim vlPenGarBenef As Double

Dim vlSumarTotalPorcentajePension As Double, vlSumaDef As Double, vlDif As Double
Dim vlPorcentajeRecal As Double
Dim vlRemuneracion As Double, vlRemuneracionProm As Double
Dim vlRemuneracionBase As Double
Dim vlPrcCobertura As Double

'I--- ABV 21/06/2007 ---
'On Error GoTo Err_fgCalcularPorcentaje
'F--- ABV 21/06/2007 ---

'    Call flAsignaPorcentajesLegales
'    Call flValidaBeneficiarios
'    Call flDerechoAcrecer
'    Call flMadreHijoMenor
'    Call flVariosConyuges
'    Call flHijosSolos
    
    fgCalcularPorcentajeBenef_Old = False
    L24 = 0
    'Debiera tomar la Fecha de Devengue
    
    If (fgCarga_Param("LI", "L24", iFechaIniVig) = True) Then
        L24 = vgValorParametro
    Else
        vgError = 1000
        MsgBox "No existe Edad de tope para los 24 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    
    'Mensualizar la Edad de 24 Años
    L24 = L24 * 12
'I--- ABV 21/07/2006 ---
'Edad de 18 años
Dim L18 As Long
If (fgCarga_Param("LI", "L18", iFechaIniVig) = True) Then
    L18 = vgValorParametro
Else
    vgError = 1000
    MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
    Exit Function
End If
'Mensualizar la Edad de 24 Años
L18 = L18 * 12
'F--- ABV 21/07/2006 ---

    'If Not IsDate(txt_fecha_devengo) Then
    '   X = MsgBox("Debe ingresar la Fecha de Inicio de la Pensión", 16)
    '   txt_fecha_devengo.SetFocus
    '   Exit Function
    'End If
    'If CDate(txt_fecha_devengo) > CDate(lbl_cotizacion) Then
    '    X = MsgBox("Error", "La fecha de devengamiento no puede ser mayor que la fecha de cotización", 16)
    '    Exit Function
    'End If
    'Iaap = CInt(Year(CDate(txt_fecha_devengo))) 'a Fecha de siniestro
    'Immp = CInt(Month(CDate(txt_fecha_devengo))) 'm Fecha de siniestro
    'Iddp = CInt(Day(CDate(txt_fecha_devengo))) 'd Fecha de siniestro
    
    Iaap = CInt(Mid(iFechaIniVig, 1, 4)) 'a Fecha de siniestro
    Immp = CInt(Mid(iFechaIniVig, 5, 2)) 'm Fecha de siniestro
    Iddp = CInt(Mid(iFechaIniVig, 7, 2)) 'd Fecha de siniestro
    
    fecha_sin = Iaap * 12 + Immp
    'sexo_cau = Trim(Mid(cbo_sexo, 1, (InStr(1, cbo_sexo, "-") - 1)))
    'vlNum_Ben = txt_n_orden - 1 '.Rows - 1
    vlNum_Ben = iNumBenef 'grd_beneficiarios.Rows - 1 '.Rows - 1
    
    ReDim Numor(vlNum_Ben) As Integer
    ReDim Ncorbe(vlNum_Ben) As Integer
    ReDim Codrel(vlNum_Ben) As Integer
    ReDim Cod_Grfam(vlNum_Ben) As Integer
    ReDim Sexobe(vlNum_Ben) As String
    ReDim Inv(vlNum_Ben) As String
    ReDim Coinbe(vlNum_Ben) As String
    ReDim derpen(vlNum_Ben) As Integer
    ReDim Nanbe(vlNum_Ben) As Integer
    ReDim Nmnbe(vlNum_Ben) As Integer
    ReDim Ndnbe(vlNum_Ben) As Integer
    ReDim Hijos(vlNum_Ben) As Integer
    ReDim Hijos_Inv(vlNum_Ben) As Integer
    ReDim Hijos_SinDerechoPension(vlNum_Ben) As Integer
    ReDim Hijo_Menor(vlNum_Ben) As Date
    ReDim Hijo_Menor_Ant(vlNum_Ben) As Date
    ReDim Porben(vlNum_Ben) As Double
    ReDim Codcbe(vlNum_Ben) As String
    ReDim Fec_NacHM(vlNum_Ben) As Date
    ReDim cont_mhn(vlNum_Ben) As Integer
    
'I--- ABV 18/07/2006 ---
    ReDim Hijos_SinConyugeMadre(vlNum_Ben) As Integer
    ReDim Fec_Fall(vlNum_Ben) As String
    ReDim cont_mhn_Tot_GF(vlNum_Ben) As Integer
    ReDim cont_esp_Tot_GF(vlNum_Ben) As Integer
    vlFechaFallCau = ""
    vlValorHijo = 0
'F--- ABV 18/07/2006 ---
    
If (iCalcularPorcentaje = True) Then
    'tb_cau!cod_sexo
    vlContBen = 1 '0
    i = 1
    Do While i <= vlNum_Ben
    
        vlContBen = vlContBen + 1
        'Nº Orden
        
        'Msf_GrillaBenef.Row = i
        'Msf_GrillaBenef.Col = 0
        'Numor(i) = Msf_GrillaBenef.Text ''tb_ben!cod_numordben 'N° de orden  NUMOR(I)
        
        'If Trim(grd_beneficiarios.TextMatrix(i, 0)) = "" Then
        If Trim(ostBeneficiarios(i).Num_Orden) = "" Then
            vgError = 1000
            MsgBox "No existe Número de Orden de Beneficiario.", vbCritical, "Error de Datos"
            Exit Do
        End If
        
        'Número de Orden
        Numor(i) = ostBeneficiarios(i).Num_Orden
        
        'Parentesco
        Ncorbe(i) = ostBeneficiarios(i).Cod_Par ''tb_ben!cod_par
        Codrel(i) = ostBeneficiarios(i).Cod_Par ''tb_ben!cod_par
        
        'Grupo Familiar
        Cod_Grfam(i) = ostBeneficiarios(i).Cod_GruFam
        
        'Sexo
        Sexobe(i) = ostBeneficiarios(i).Cod_Sexo
        If (Ncorbe(i) = "99") Then
            sexo_cau = Sexobe(i)
        End If
        
        'Situación de Invalidez
        Inv(i) = ostBeneficiarios(i).Cod_SitInv
        
        'Derecho Pensión
        derpen(i) = ostBeneficiarios(i).Cod_EstPension
        
        'Fecha de Nacimiento
        Nanbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 1, 4)) 'a Fecha de nacimiento
        Nmnbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 5, 2)) 'm Fecha de nacimiento
        Ndnbe(i) = CInt(Mid(ostBeneficiarios(i).Fec_NacBen, 7, 2)) 'd Fecha de nacimiento
                    
        'Fecha nacimiento hijo menor =IJAM(I),IJMN(I),IJDN(I)
        
        'Codificación de Situación de Invalidez
        If Inv(i) = "P" Then Coinbe(i) = "P"
        If Inv(i) = "T" Then Coinbe(i) = "T"
        If Inv(i) = "N" Then Coinbe(i) = "N"
        
        '*********
        edad_mes_ben = fecha_sin - (Nanbe(i) * 12 + Nmnbe(i))
        
        vlFechaFallecimiento = ostBeneficiarios(i).Fec_FallBen
'I--- ABV 18/07/2006 ---
        Fec_Fall(i) = vlFechaFallecimiento
        If Codrel(i) = 99 Then
            vlFechaFallCau = vlFechaFallecimiento
        End If
'F--- ABV 18/07/2006 ---

        vlFechaMatrimonio = ostBeneficiarios(i).Fec_Matrimonio
        
        'I--- ABV 16/04/2005 ---
        If (Codrel(i) >= 30 And Codrel(i) < 40) Then
            If (vlFechaFallecimiento <> "") Or (vlFechaMatrimonio <> "") Then
                derpen(i) = 10
                If ostBeneficiarios(i).Cod_EstPension <> "10" Then
                    ostBeneficiarios(i).Cod_EstPension = "10"
                    Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
                End If
            Else
                If edad_mes_ben > L24 And Coinbe(i) = "N" Then
                    derpen(i) = 10
                    'hqr 08/07/2006 cambiar Estado de pension a No Vigente.  Se debe cambiar estado de madre si es hijo unico
                    If ostBeneficiarios(i).Cod_EstPension <> "10" Then
                        ostBeneficiarios(i).Cod_EstPension = "10"
                        Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinDerechoPension(ostBeneficiarios(i).Cod_GruFam) + 1
                    End If
                Else
'I--- ABV 21/07/2006 ---
'Validación Especial para la Cía Chilena. Si ellos indican que no se encuentra
'vigente cuando tiene menos de 24 años, se debe modificar
                    derpen(i) = 99
'                    If ostBeneficiarios(i).Cod_DerPen = "10" Then
'                        If edad_mes_ben > L18 Then
'                            Derpen(i) = 10
'                        End If
'                    End If
'F--- ABV 21/07/2006 ---
                End If
            End If
        Else
            'El resto de los Beneficiarios que no sean Hijos, solo se dejan como
            'Sin Derecho a Pensión cuando están fallecidos
            If (vlFechaFallecimiento <> "") Then
                derpen(i) = 10
'I--- ABV 18/07/2006 ---
'                If (Codrel(i) = "11") Or (Codrel(i) = "21") Then
'                    If ostBeneficiarios(i).Cod_EstPension <> "10" Then
'                        ostBeneficiarios(i).Cod_EstPension = "10"
'                        Hijos_SinConyugeMadre(ostBeneficiarios(i).Cod_GruFam) = Hijos_SinConyugeMadre(ostBeneficiarios(i).Cod_GruFam) + 1
'                    End If
'                End If
'F--- ABV 18/07/2006 ---
            Else
                derpen(i) = 99
            End If
        End If
        ostBeneficiarios(i).Cod_DerPen = derpen(i)
        'F--- ABV 16/04/2005 ---
        
        i = i + 1
    Loop
    
    cont_causante = 0
    cont_esposa = 0
    'cont_mhn = 0
    cont_mhn_tot = 0
    cont_hijo = 0
    cont_padres = 0

'I--- ABV 18/07/2006 ---
    cont_esp_Totales = 0
    cont_mhn_Totales = 0
'F--- ABV 18/07/2006 ---
    
    'Primer Ciclo
    For g = 1 To vlNum_Ben
        If derpen(g) <> 10 Then
            '99: con derecho a pension,20: con Derecho Pendiente
            '10: sin derecho a pension
            If Ncorbe(g) = 99 Then
                cont_causante = cont_causante + 1
            End If
            If Ncorbe(g) <> 99 Then
                Select Case Ncorbe(g)
                    Case 10, 11
                        cont_esposa = cont_esposa + 1
''I--- ABV 18/07/2006 ---
'                        Q = Cod_Grfam(g)
'                        cont_esp_Tot_GF(Q) = cont_esp_Tot_GF(Q) + 1
'                        cont_esp_Totales = cont_esp_Totales + 1
''F--- ABV 18/07/2006 ---
                    Case 20, 21
                        Q = Cod_Grfam(g)
                        cont_mhn(Q) = cont_mhn(Q) + 1
                        cont_mhn_tot = cont_mhn_tot + 1
''I--- ABV 18/07/2006 ---
'                        cont_mhn_Tot_GF(Q) = cont_mhn_Tot_GF(Q) + 1
'                        cont_mhn_Totales = cont_mhn_Totales + 1
''F--- ABV 18/07/2006 ---
                    Case 30
                        'edad = fgEdadBen(vg_fecsin, vgBen(g).fec_nacben)
                        Q = Cod_Grfam(g)
                        'q = ncorbe(g)
                        Hijos(Q) = Hijos(Q) + 1
                        If Coinbe(g) <> "N" Then Hijos_Inv(Q) = Hijos_Inv(Q) + 1
                        Hijo_Menor(Q) = DateSerial(Nanbe(g), Nmnbe(g), Ndnbe(g))
                        'hijo_menor_ant = hijo_menor(q)
                        If Hijos(Q) > 1 Then
                            If Hijo_Menor(Q) > Hijo_Menor_Ant(Q) Then
                                Hijo_Menor_Ant(Q) = Hijo_Menor(Q)
                            End If
                        Else
                            Hijo_Menor_Ant(Q) = Hijo_Menor(Q)
                        End If
                    Case 35
                        edad_mes_ben = fecha_sin - (Nanbe(g) * 12 + Nmnbe(g))
                        If Coinbe(g) = "P" And edad_mes_ben <= L24 Then
                            cont_hijo = cont_hijo + 1
                        Else
                            If Coinbe(g) = "T" Or Coinbe(g) = "N" Then
                                cont_hijo = cont_hijo + 1
                            End If
                        End If
    
                    Case 41, 42
                        cont_padres = cont_padres + 1
                    Case Else
                        vgError = 1000
                        X = MsgBox("Error en codificación de codigo de relación", vbCritical)
                        Exit Function
                End Select
            End If
'I--- ABV 18/07/2006 ---
        Else
            'Verificar si la cónyuge o la Madre Falleció antes que el Causante para contarla
            Select Case Ncorbe(g)
                Case 11:
                    If (vlFechaFallCau > Fec_Fall(g)) Then
                        cont_esposa = cont_esposa + 1
                        Q = Cod_Grfam(g)
                        cont_esp_Tot_GF(Q) = cont_esp_Tot_GF(Q) + 1
                        cont_esp_Totales = cont_esp_Totales + 1
'                    Else
'                        Q = Cod_Grfam(g)
'                        cont_esp_Tot_GF(Q) = cont_esp_Tot_GF(Q) - 1
'                        cont_esp_Totales = cont_esp_Totales - 1
                    End If
                Case 21:
'                    If (vlFechaFallCau > Fec_Fall(g)) Then
'                        Q = Cod_Grfam(g)
'                        cont_mhn_Tot_GF(Q) = cont_mhn_Tot_GF(Q) + 1
'                        cont_mhn_Totales = cont_mhn_Totales + 1
''                    Else
''                        Q = Cod_Grfam(g)
''                        cont_mhn_Tot_GF(Q) = cont_mhn_Tot_GF(Q) - 1
''                        cont_mhn_Totales = cont_mhn_Totales - 1
'                    End If
            End Select
'F--- ABV 18/07/2006 ---
        End If
    Next g
                
'I--- ABV 18/07/2006 ---
'Corregir el Parentesco de los Hijos cuando la Cónyuge o la Madre haya Muerto antes que el Causante
    For j = 1 To vlNum_Ben
        '99: con derecho a pension,20: con Derecho Pendiente
        '10: sin derecho a pension
        If derpen(j) <> 10 Then
'            edad_mes_ben = fecha_sin - (Nanbe(j) * 12 + Nmnbe(j))
            Select Case Ncorbe(j)
                
                Case 30
                    Q = Cod_Grfam(j)
                    'Cuando No existe Conyuge y tampoco no Existe MHN para el Hijo a Analizar
                    'se deben modificar los Códigos de los Hijos a 35
'I--- ABV 20/07/2006 ---
'                    If cont_esp_Tot_GF(Q) <= 0 And cont_mhn_Tot_GF(Q) <= 0 Then
                    If cont_esposa <= 0 And cont_mhn(Q) <= 0 Then
'F--- ABV 20/07/2006 ---
                        Ncorbe(j) = 35
                        Hijos(Q) = Hijos(Q) - 1
                        cont_hijo = cont_hijo + 1
                        If ostBeneficiarios(j).Cod_Par <> "35" Then
                            ostBeneficiarios(j).Cod_Par = "35"
                            'Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) = Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) + 1
                        End If
                    End If
                Case 35
                    Q = Cod_Grfam(j)
                    'Cuando existe Conyuge o MHN, considerándola como existente
                    'cuando ésta muere antes del Causante de la Póliza,
                    'se deben modificar los Códigos de los Hijos a 30
'I--- ABV 20/07/2006 ---
'                    If cont_esp_Tot_GF(Q) > 0 Or cont_mhn_Tot_GF(Q) > 0 Then
                    If cont_esposa > 0 Or cont_mhn(Q) > 0 Then
'F--- ABV 20/07/2006 ---
                        Ncorbe(j) = 30
                        Hijos(Q) = Hijos(Q) + 1
                        cont_hijo = cont_hijo - 1
                        If ostBeneficiarios(j).Cod_Par <> "30" Then
                            ostBeneficiarios(j).Cod_Par = "30"
                            'Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) = Hijos_SinConyugeMadre(ostBeneficiarios(j).Cod_GruFam) + 1
                        End If
                    End If
            End Select
        End If
    Next j
'F--- ABV 18/07/2006 ---
                
    j = 1
    For j = 1 To vlNum_Ben
        '99: con derecho a pension,20: con Derecho Pendiente
        '10: sin derecho a pension
        If derpen(j) <> 10 Then
            edad_mes_ben = fecha_sin - (Nanbe(j) * 12 + Nmnbe(j))
            Select Case Ncorbe(j)
                Case 99
                    If cont_causante > 1 Then
                        vgError = 1000
                        X = MsgBox("Error en codificación de codigo de relación, No puede ingresar otro causante", vbCritical)
                        Exit Function
                    End If
                    'I--- ABV 25/02/2005 ---
                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                        vlValor = vgValorPorcentaje
                    Else
                        vgError = 1000
                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                        Exit Function
                    End If
                    'F--- ABV 25/02/2005 ---
                    If (vlValor < 0) Then
                        vgError = 1000
                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                        'error
                        Exit Function
                    Else
                        Porben(j) = vlValor
                        'Porben(j) = 100
                        Codcbe(j) = "N"
                    End If
                Case 10, 11
                    If sexo_cau = "M" Then
                        If Sexobe(j) <> "F" Then
                            vgError = 1000
                            X = MsgBox("Error de código de sexo, el Sexo de la Cónyuge debe ser Femenino.", vbCritical)
                            Exit Function
                        End If
                    Else
                        If Sexobe(j) <> "M" Then
                            vgError = 1000
                            X = MsgBox("Error de codigo de sexo, el Sexo del Cónyuge debe ser Masculino.", vbCritical)
                            Exit Function
                        End If
                    End If
                    'hqr 08/07/2006 validacion y cambio parentesco
                    u = Cod_Grfam(j)
                    If Hijos(u) = 0 And Ncorbe(j) = 11 Then
                        If Hijos_SinDerechoPension(u) = 0 Then
                            vgError = 1000
                            X = MsgBox("Error de código de relación, 'Cónyuge Con Hijos con Dº Pensión', no tiene Hijos.", vbCritical)
                            Exit Function
                        Else
                            Ncorbe(j) = 10
                            ostBeneficiarios(j).Cod_Par = 10
                        End If
                    End If
                    
                    'HQR 08/07/2006 se deja al final porque se cambia el tipo de parentesco
                    'I--- ABV 25/02/2005 ---
                     'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                     If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                         vlValor = vgValorPorcentaje
                     Else
                         vgError = 1000
                         MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                         & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                         Exit Function
                     End If
                     'F--- ABV 25/02/2005 ---
                     If (vlValor < 0) Then
                         vgError = 1000
                         MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                         'error
                         Exit Function
                     End If

                    'fin hqr 08/07/2006
                    If sexo_cau = "M" Then
                        'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                        'If (vlValor < 0) Then
                        '    Exit Sub
                        'Else
                            Porben(j) = CDbl(Format(vlValor / cont_esposa, "#0.00"))
                        'End If
                                         
                        If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
                            If Hijos(u) > 0 Then
                                Codcbe(j) = "S"
                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
                            End If
                        Else
                            Codcbe(j) = "N"
                        End If
                            
                        If Hijos(u) > 0 And Ncorbe(j) = 10 Then
                            vgError = 1000
                            X = MsgBox("Error de código de relación, 'Cónyuge Sin Hijos con Dº Pensión', tiene Hijos.", vbCritical)
                            Exit Function
                        End If
                    Else
                        u = Cod_Grfam(j)
                        If Hijos(u) > 0 Then
                            'I--- ABV 25/02/2005 ---
                            'If Coinbe(j) = "T" Then
                            '    Porben(j) = vlValor     '50
                            'Else
                            '    'HQR 16-06-2004
                            '    'Porben(j) = 36
                            '    If Coinbe(j) = "P" Then
                            '        Porben(j) = 36
                            '    Else
                            '        Porben(j) = 0
                            '        cont_esposa = cont_esposa - 1
                            '    End If
                            '    'FIN HQR 16-06-2004
                            'End If
                            
                            Porben(j) = vlValor
                            If (Coinbe(j) = "N") Then
                                cont_esposa = cont_esposa - 1
                            End If
                            'F--- ABV 25/02/2005 ---
                            
                            If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
                                Codcbe(j) = "S"
                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
                            Else
                                Codcbe(j) = "N"
                            End If
                        Else
                            'I--- ABV 25/02/2005 ---
                            'If Coinbe(j) = "T" Then
                            '    Porben(j) = vlValor     '60
                            'Else
                            '    'HQR 16-06-2004
                            '    'Porben(j) = 43
                            '    If Coinbe(j) = "P" Then
                            '        Porben(j) = 43
                            '    Else
                            '        Porben(j) = 0
                            '        cont_esposa = cont_esposa - 1
                            '    End If
                            '    'FIN HQR 16-06-2004
                            'End If
                            
                            Porben(j) = vlValor
                            If (Coinbe(j) = "N") Then
                                cont_esposa = cont_esposa - 1
                            End If
                            'F--- ABV 25/02/2005 ---
                            
                            If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
                                Codcbe(j) = "S"
                                If Hijos_Inv(u) > 0 Then Codcbe(j) = "N"
                            Else
                                Codcbe(j) = "N"
                            End If
                        End If
                    End If
                Case 20, 21
                    If sexo_cau = "M" Then
                        If Sexobe(j) <> "F" Then
                            vgError = 1000
                            X = MsgBox("Error de código de sexo, el Sexo de la Cónyuge debe ser Femenino.", vbCritical)
                            Exit Function
                        End If
                    Else
                        vgError = 1000
                        X = MsgBox("Error de codigo de sexo, el Sexo del Cónyuge debe ser Masculino.", vbCritical)
                        Exit Function
                    End If
                    
                    u = Cod_Grfam(j)
                    If Hijos(u) = 0 And Ncorbe(j) = 21 Then
                        If Hijos_SinDerechoPension(u) = 0 Then
                            vgError = 1000
                            X = MsgBox("Error en código de relación 'Madre Con Hijos con Dº Pensión, no tiene Hijos.", vbCritical)
                            Exit Function
                        Else
                            Ncorbe(j) = 20
                            ostBeneficiarios(j).Cod_Par = 20
                        End If
                    End If
                    
                    'I--- ABV 25/02/2005 ---
                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                        vlValor = vgValorPorcentaje
                    Else
                        vgError = 1000
                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                        Exit Function
                    End If
                    'F--- ABV 25/02/2005 ---
                    If (vlValor < 0) Then
                        vgError = 1000
                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                        Exit Function
                    Else
                        Porben(j) = vlValor / cont_mhn_tot
                    End If

                    If (iDerCrecerCotizacion = "S") Then 'I--- ABV 14/07/2007
                        If Hijos(u) > 0 Then
                            Codcbe(j) = "S"
                        Else
                            Codcbe(j) = "N"
                        End If
                        If Hijos_Inv(u) > 0 Then
                            Codcbe(j) = "N"
                        End If
                    Else
                        Codcbe(j) = "N"
                    End If

                    If Hijos(u) > 0 And Ncorbe(j) = 20 Then
                        vgError = 1000
                        X = MsgBox("Error en código de relación 'Madre Sin Hijos con Dº Pensión', tiene Hijos.", vbCritical)
                        Exit Function
                    End If
                    
                Case 30
                    Codcbe(j) = "N"
                    Q = Cod_Grfam(j)
                    
                    If cont_esposa > 0 Or cont_mhn(Q) > 0 Then
                        If Coinbe(j) = "N" And edad_mes_ben > L24 Then
                            Porben(j) = 0
                        Else
                            'I--- ABV 25/02/2005 ---
                            'If Coinbe(j) = "P" And edad_mes_ben > L24 Then
                            If (Coinbe(j) = "P" Or Coinbe(j) = "T") And edad_mes_ben > L24 Then
                            'F--- ABV 25/02/2005 ---
                                
                                'Porben(j) = 11
                                'I--- ABV 25/02/2005 ---
                                'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                                    vlValor = vgValorPorcentaje
                                Else
                                    vgError = 1000
                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                    & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                    Exit Function
                                End If
                                
                                If (vlValor < 0) Then
                                    vgError = 1000
                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                    'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
                                    Exit Function
                                Else
                                    Porben(j) = vlValor
                                End If
                                'F--- ABV 25/02/2005 ---
                            
                            Else
                                
                                'I--- ABV 25/02/2005 ---
                                If (fgObtenerPorcentaje(CStr(Ncorbe(j)), "N", Sexobe(j), iFechaIniVig) = True) Then
                                    'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                    vlValor = vgValorPorcentaje
                                Else
                                    vgError = 1000
                                    MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                    & Ncorbe(j) & " - N - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                    Exit Function
                                End If
                                'F--- ABV 25/02/2005 ---
                                
                                If (vlValor < 0) Then
                                    vgError = 1000
                                    'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
                                    MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                    Exit Function
                                Else
                                    'Porben(j) = 15
                                    Porben(j) = vlValor
                                End If
                            End If
                        End If
                        'If cont_esposa = 0 Or cont_mhn > 0 Then
                    Else
                        vgError = 1000
                        X = MsgBox("Error: Los códigos de beneficiarios de hijos estan mal ingresados", vbCritical)
                        Exit Function
                    End If
                Case 35
                    Q = Cod_Grfam(j)
                    Codcbe(j) = "N"
                    
'I--- ABV 18/07/2006 ---
'                    vlValorHijo = 0
'                    If cont_esp_Tot_GF(Q) > 0 Or cont_mhn_Tot_GF(Q) > 0 Then
'                        If Hijos_SinConyugeMadre(Q) = 0 Then
'                            vgError = 1000
'                            X = MsgBox("Error: Los códigos de beneficiarios de hijos estan mal ingresados", vbCritical)
'                            Exit Function
'                        End If
'                    End If
'F--- ABV 18/07/2006 ---
                    
'I--- ABV 20/07/2006 ---
'                    If cont_esposa = 0 And cont_mhn(Q) = 0 Then
                    If cont_esposa = 0 And cont_mhn(Q) = 0 Then
'F--- ABV 20/07/2006 ---

                        'I--- ABV 25/02/2005 ---
                        'If Coinbe(j) = "P" And edad_mes_ben > L24 Then
                        If (Coinbe(j) = "P" Or Coinbe(j) = "T") And edad_mes_ben > L24 Then
                        'F--- ABV 25/02/2005 ---
                            'Porben(j) = 11
                            'I--- ABV 25/02/2005 ---
                            'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                            If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                                vlValor = vgValorPorcentaje
                            Else
                                vgError = 1000
                                MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                Exit Function
                            End If
                            
                            If (vlValor < 0) Then
                                vgError = 1000
                                'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
                                MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                Exit Function
                            Else
                                Porben(j) = vlValor
                            End If
                            'F--- ABV 25/02/2005 ---
                        Else
                            'Porben(j) = 15
                            'A = L24
                        
                            'I--- ABV 25/02/2005 ---
                            If (fgObtenerPorcentaje(CStr(Ncorbe(j)), "N", Sexobe(j), iFechaIniVig) = True) Then
                                'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                                vlValor = vgValorPorcentaje
'I--- ABV 18/07/2006 ---
                                vlValorHijo = vlValor
'F--- ABV 18/07/2006 ---
                            Else
                                vgError = 1000
                                MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                                & Ncorbe(j) & " - N - " & Sexobe(j) & " - " & iFechaIniVig & "."
                                Exit Function
                            End If
                            'F--- ABV 25/02/2005 ---
                            
                            If (vlValor < 0) Then
                                vgError = 1000
                                MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                                'X = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
                                Exit Function
                            Else
                                'Porben(j) = 15
                                Porben(j) = vlValor
                            End If
                        End If
                        
                        'Sql = "select prc_par from porpar where cod_par = 11"
                        'Set tb_por = vlConBD.Execute(Sql)
                        'If Not tb_por.EOF Then
                        '    v_hijo = tb_por!prc_par
                        '    If coinbe(J) = "P" And edad_mes_ben <= l24 Then
                        '        porben(J) = v_hijo / cont_hijo + 15
                        '    Else
                        '        If coinbe(J) = "T" Or coinbe(J) = "N" Then
                        '            porben(J) = v_hijo / cont_hijo + 15
                        '        End If
                        '    End If
                        'Else
                        '    x = MsgBox("Error: Código de cónyuge no esta incorporado", vbCritical)
                        'End If
                        
                        'Obtener el Porcentaje de la Cónyuge
                        'I--- ABV 25/02/2005 ---
                        If (fgObtenerPorcentaje("10", "N", "F", iFechaIniVig) = True) Then
                            'vlValor = fgValorPorcentaje(1, j, 11)
                            vlValor = vgValorPorcentaje
                        Else
                            vgError = 1000
                            MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                            & "10 - N - F  - " & iFechaIniVig & "."
                            Exit Function
                        End If
                        'F--- ABV 25/02/2005 ---
                        
                        If (vlValor < 0) Then
                            vgError = 1000
                            'X = MsgBox("Error, el porcentaje para la 'Cónyuge Con Hijos' no se encuentra.", vbCritical)
                            MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & "11" & ".", vbCritical, "Error de Datos"
                            Exit Function
                        Else
                            v_hijo = vlValor
                            If Coinbe(j) = "P" And edad_mes_ben <= L24 Then
                                Porben(j) = v_hijo / cont_hijo + vlValorHijo
                            Else
                                If Coinbe(j) = "T" Or Coinbe(j) = "N" Then
                                    Porben(j) = v_hijo / cont_hijo + vlValorHijo
                                End If
                            End If
                        End If
                        
                    Else
                        vgError = 1000
                        X = MsgBox("Error: Los códigos de beneficiarios de hijos estan mal ingresados", vbCritical)
                        Exit Function
                    End If
                    
                Case 41, 42
                    'I--- ABV 25/02/2005 ---
                    If (fgObtenerPorcentaje(CStr(Ncorbe(j)), Coinbe(j), Sexobe(j), iFechaIniVig) = True) Then
                        'vlValor = fgValorPorcentaje(1, j, Ncorbe(j))
                        vlValor = vgValorPorcentaje
                    Else
                        vgError = 1000
                        MsgBox "No existe valor de Porcentaje para el parámetro de :" _
                        & Ncorbe(j) & " - " & Coinbe(j) & " - " & Sexobe(j) & " - " & iFechaIniVig & "."
                        Exit Function
                    End If
                    'F--- ABV 25/02/2005 ---
                    
                    If (vlValor < 0) Then
                        vgError = 1000
                        MsgBox "Valor del Porcentaje registrado es <= a Cero para el parentesco " & Ncorbe(j) & ".", vbCritical, "Error de Datos"
                        Exit Function
                    Else
                        Codcbe(j) = "N"
                        Porben(j) = vlValor
                    End If
                'Case 42
                '    Codcbe(j) = "N"
                '    Porben(j) = 50
            End Select
        End If
    Next j
    
    For k = 1 To vlNum_Ben '60
        If derpen(k) <> 10 Then
            Select Case Ncorbe(k)
                Case 11
                    'q = codrel(k)
                    Q = Cod_Grfam(k)
                    'If codcbe(k) = "S" Then DAJ 10/08/2007
                        Fec_NacHM(k) = Hijo_Menor_Ant(Q)  '******Agruegué 05/11/2000 HILDA
                    'End If DAJ 10/08/2007
                Case 21
                    'q = codrel(k)
                    Q = Cod_Grfam(k)
                    'If codcbe(k) = "S" Then DAJ 10/08/2007
                        Fec_NacHM(k) = Hijo_Menor_Ant(Q)  '*******Agruegué 05/11/2000 HILDA
                    'End If DAJ 10/08/2007
            End Select
        End If
    Next k
    
    vlSumarTotalPorcentajePension = 0
    
    For j = 1 To (vlContBen - 1)
        'Guardar el Valor del Porcentaje Calculado
        'If IsNumeric(Porben(j)) Then
        '    grd_beneficiarios.Text = Format(Porben(j), "##0.00")
        'Else
        '    grd_beneficiarios.Text = Format("0", "0.00")
        'End If
        If IsNumeric(Porben(j)) Then
            'I--- ABV 22/06/2005 ---
            'ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
            If (ostBeneficiarios(j).Cod_Par = "99" Or ostBeneficiarios(j).Cod_Par = "0") Then
                If (derpen(j) <> 10) Then
                    ostBeneficiarios(j).Prc_Pension = 100
                    ostBeneficiarios(j).Prc_PensionLeg = 100 '*-+
                Else
                    ostBeneficiarios(j).Prc_Pension = 0
                    ostBeneficiarios(j).Prc_PensionLeg = 0 '*-+
                End If
            Else
                ostBeneficiarios(j).Prc_Pension = Format(Porben(j), "#0.00")
                ostBeneficiarios(j).Prc_PensionLeg = Format(Porben(j), "#0.00") '*-+
            End If
            'F--- ABV 22/06/2005 ---
        Else
            ostBeneficiarios(j).Prc_Pension = 0
            ostBeneficiarios(j).Prc_PensionLeg = 0 '*-+
        End If
        
        'Guardar el Derecho a Acrecer de los Beneficiarios
        'Inicio
        'If (Codcbe(j) <> Empty And Not IsNull(Codcbe(j))) Then
        '    grd_beneficiarios.Text = Codcbe(j)
        'Else
        '    'Por Defecto Negar el Derecho a Acrecer de los Beneficiarios
        '    grd_beneficiarios.Text = "N"
        'End If
        If (Codcbe(j) <> Empty And Not IsNull(Codcbe(j))) Then
            ostBeneficiarios(j).Cod_DerCre = Codcbe(j)
        Else
            ostBeneficiarios(j).Cod_DerCre = "N"
        End If
        'Fin
        
        'Guardar la Fecha de Nacimiento del Hijo Menor de la Cónyuge
        'Inicio
        ''If Format((Fec_NacHM(J)), "yyyy/mm/dd") > "1899/12/30" Then
        'If Format(CDate(Fec_NacHM(j)), "yyyymmdd") > "18991230" Then
        '    grd_beneficiarios.Text = CDate(Fec_NacHM(j))
        'Else
        '    'Guardar la Fecha de Nacimiento del Hijo Menor de la Cónyuge
        '    grd_beneficiarios.Text = ""
        'End If
        If Format(CDate(Fec_NacHM(j)), "yyyymmdd") > "18991230" Then
            'ostBeneficiarios(j).Fec_NacHM = CDate(Fec_NacHM(j))
            ostBeneficiarios(j).Fec_NacHM = Format(CDate(Fec_NacHM(j)), "yyyymmdd")
        Else
            ostBeneficiarios(j).Fec_NacHM = ""
        End If
        'Fin

'I--- ABV 10/08/2007 ---
        'Inicializar el Monto de Pensión a Cero
        '--------------------------------------
        'Actualizar el Monto de la Pensión
        ostBeneficiarios(j).Mto_Pension = Format(0, "#0.00")
        
        'Actualizar el Monto de la Pensión Garantizada
        ostBeneficiarios(j).Mto_PensionGar = Format(0, "#0.00")
'F--- ABV 10/08/2007 ---
        
'I--- ABV 21/06/2006 ---
'        If (iCalcularPension = True) Then
'
'            vlPorcBenef = 0
'            vlPenBenef = 0
'            vlPenGarBenef = 0
'
'            'Verificar el Estado de Derecho a Pensión (Cod_EstPension)
'            If (ostBeneficiarios(j).Cod_EstPension <> "10") Then
'                vlPorcBenef = ostBeneficiarios(j).Prc_Pension
'
'                'Determinar el monto de la Pensión del Beneficiario
'                'Penben(i) = IIf(Msf_BMGrilla.TextMatrix(vlI, 17) = "", 0, Msf_BMGrilla.TextMatrix(vlI, 17))
'                vlPenBenef = iPensionRef * (vlPorcBenef / 100)
'
'                'Definir la Pensión del Garantizado
'                vlPenGarBenef = ostBeneficiarios(j).Mto_PensionGar
'
'                'I--- ABV 20/04/2005 ---
'                vgPalabra = iTipoPension
'                vgI = InStr(1, cgPensionInvVejez, vgPalabra)
'                If (vgI <> 0) Then
'                    'Valida que sea un caso de Invalidez o Vejez
'                    vlPenGarBenef = vlPenBenef
'                End If
'
'                vgI = InStr(1, cgPensionSobOrigen, vgPalabra)
'                If (vgI <> 0) Then
'                    'Valida que sea un caso de Sobrevivencia por Origen
'                    'Se supone que la Cónyuge o Madre es la única que Garantiza su pensión
'                    'vlNumero = InStr(Cmb_BMPar.Text, "-")
'                    vgPalabraAux = ostBeneficiarios(j).Cod_Par
'
'                    vgX = InStr(1, cgParConyugeMadre, vgPalabraAux)
'                    If (vgX <> 0) Then
'                        vlPenGarBenef = vlPenBenef
'                    'Else
'                    '    Txt_BMMtoPensionGar.Enabled = True
'                    End If
'                End If
'
'                vgI = InStr(1, cgPensionSobTransf, vgPalabra)
'                If (vgI <> 0) Then
'                    ''Valida que sea un caso de Invalidez o Vejez
'                    'Txt_BMMtoPensionGar.Enabled = True
'                End If
'
'                'F--- ABV 19/04/2005 ---
'            End If
'
'            'Actualizar el Monto de la Pensión
'            ostBeneficiarios(j).Mto_Pension = Format(vlPenBenef, "#0.00")
'
'            'Actualizar el Monto de la Pensión Garantizada
'            ostBeneficiarios(j).Mto_PensionGar = Format(vlPenGarBenef, "#0.00")
'
'        End If
        
        If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
            vlSumarTotalPorcentajePension = vlSumarTotalPorcentajePension + ostBeneficiarios(j).Prc_Pension
        End If
'F--- ABV 21/06/2006 ---
    Next j
    
'Validar si los Porcentajes de Pensión Suman más del 100%
    If (vlSumarTotalPorcentajePension > 100) Then
        vlSumaDef = 0
        For j = 1 To (vlContBen - 1)
            If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                vlPorcentajeRecal = (ostBeneficiarios(j).Prc_Pension / vlSumarTotalPorcentajePension) * 100
                ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                vlSumaDef = vlSumaDef + ostBeneficiarios(j).Prc_Pension
            End If
        Next j
        vlDif = Format(100 - vlSumaDef, "#0.00")
        If (vlDif < 0) Or (vlDif > 0) Then
            'Asignar la diferencia de Porcentaje al Primer Beneficiario con Derecho distinto del Causante
            For j = 1 To (vlContBen - 1)
                If (ostBeneficiarios(j).Cod_DerPen <> "10") And (ostBeneficiarios(j).Cod_Par <> "99" And ostBeneficiarios(j).Cod_Par <> "0") Then
                    vlPorcentajeRecal = ostBeneficiarios(j).Prc_Pension + vlDif
                    ostBeneficiarios(j).Prc_Pension = vlPorcentajeRecal
                    Exit For
                End If
            Next j
        End If
    End If

'*-+
'I--- ABV 04/08/2007 ---
'Recalcular Porcentajes si se trata de un Caso de Invalidez - Con Cobertura
    If (iTipoPension = "06" Or iTipoPension = "07") And (iCobCobertura = "S") Then
        If (fgObtenerPorcCobertura(iTipoPension, iFechaIniVig, vlRemuneracion, vlPrcCobertura) = True) Then
            'Determinar la Remuneración Promedio para el Causante
            vlRemuneracionBase = vlRemuneracion * (vlPrcCobertura / 100)
            
            'Determinar para cada Beneficiario el Nuevo Porcentaje
            For j = 1 To (vlContBen - 1)
                If (ostBeneficiarios(j).Cod_DerPen <> "10") Then
                    If (ostBeneficiarios(j).Cod_Par = "99") Then
                        vlRemuneracionProm = vlRemuneracion * (vlPrcCobertura / 100)
                    Else
                        vlRemuneracionProm = vlRemuneracion * (ostBeneficiarios(j).Prc_Pension / 100)
                    End If
                    vlPorcentajeRecal = (vlRemuneracionProm / vlRemuneracionBase) * 100
                    ostBeneficiarios(j).Prc_Pension = Format(vlPorcentajeRecal, "#0.00")
                End If
            Next j
        Else
            vgError = 1000
            MsgBox "No existen valores de Porcentaje de Cobertura para el Recalculo de Pensiones.", vbCritical, "Inexistencia de Datos"
            Exit Function
        End If
    End If
'F--- ABV 04/08/2007 ---
'*-+
End If

'Recalcular los Montos de Pensión
If (iCalcularPension = True) Then
    For j = 1 To (vlNum_Ben)
            
        vlPorcBenef = 0
        vlPenBenef = 0
        vlPenGarBenef = 0
                
        'Verificar el Estado de Derecho a Pensión (Cod_EstPension)
'            If (ostBeneficiarios(j).Cod_EstPension <> "10") Then
        If (ostBeneficiarios(j).Cod_DerPen <> "10") Then
            vlPorcBenef = ostBeneficiarios(j).Prc_Pension
            
            'Determinar el monto de la Pensión del Beneficiario
            'Penben(i) = IIf(Msf_BMGrilla.TextMatrix(vlI, 17) = "", 0, Msf_BMGrilla.TextMatrix(vlI, 17))
            vlPenBenef = iPensionRef * (vlPorcBenef / 100)
        
            'Definir la Pensión del Garantizado
            vlPenGarBenef = ostBeneficiarios(j).Mto_PensionGar
        
            'I--- ABV 20/04/2005 ---
            vgPalabra = iTipoPension
            vgI = InStr(1, cgPensionInvVejez, vgPalabra)
            If (vgI <> 0) Then
                'Valida que sea un caso de Invalidez o Vejez
                vlPenGarBenef = vlPenBenef
            End If
            
            vgI = InStr(1, cgPensionSobOrigen, vgPalabra)
            If (vgI <> 0) Then
                'Valida que sea un caso de Sobrevivencia por Origen
                'Se supone que la Cónyuge o Madre es la única que Garantiza su pensión
                'vlNumero = InStr(Cmb_BMPar.Text, "-")
                vgPalabraAux = ostBeneficiarios(j).Cod_Par
                
                vgX = InStr(1, cgParConyugeMadre, vgPalabraAux)
                If (vgX <> 0) Then
                    vlPenGarBenef = vlPenBenef
                'Else
                '    Txt_BMMtoPensionGar.Enabled = True
                End If
            End If
            
            vgI = InStr(1, cgPensionSobTransf, vgPalabra)
            If (vgI <> 0) Then
                ''Valida que sea un caso de Invalidez o Vejez
                'Txt_BMMtoPensionGar.Enabled = True
            End If

            'F--- ABV 19/04/2005 ---
        End If
        
        'Actualizar el Monto de la Pensión
        ostBeneficiarios(j).Mto_Pension = Format(vlPenBenef, "#0.00")
        
        'Actualizar el Monto de la Pensión Garantizada
        If (iPerGar > 0) Then
            ostBeneficiarios(j).Mto_PensionGar = Format(vlPenGarBenef, "#0.00")
            ostBeneficiarios(j).Prc_PensionGar = vlPorcBenef
        Else
            ostBeneficiarios(j).Mto_PensionGar = 0
            ostBeneficiarios(j).Prc_PensionGar = 0
        End If
        
    Next j
End If

    'vll_numorden = grd_beneficiarios.Rows
    'txt_n_orden.Caption = vll_numorden
    
    fgCalcularPorcentajeBenef_Old = True
    
Exit Function
Err_fgCalcularPorcentaje:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        vgError = 1000
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgActualizaGrillaBeneficiarios(iGrilla As MSFlexGrid, istBeneficiarios() As TyBeneficiarios, iNumBen As Integer, iNumMesGar As Long, iDerCrePol As String, iFecDev As String, iCodTipPen As String)
On Error GoTo Err_flCargaGrillaBeneficiarios
Dim vlCodPar As String, vlDerCre As String
Dim vlCargaFecNac   As String
Dim vlCargaFecNacHM As String
Dim vlCargaFecFall  As String
Dim vlCargaFecMat   As String
Dim vlCargaFecInv   As String
Dim vlCargaFecSus   As String
Dim vlCargaFecIniPagoPen    As String
Dim vlCargaFecTerpagoPenGar As String
Dim vlPrcPension As String
Dim vlPrcPensionGar As String
Dim vlFecNacBen As String, vlFecFallBen As String
Dim vlEstPen As String, vlSitInv As String

    vgX = 0
    
    'I--- ABV 25/04/2005 ---
    'iNumBen = iGrilla.Rows
    'F--- ABV 25/04/2005 ---
    
    'Call flInicializaGrillaBenef(iGrilla)
    While vgX < iNumBen
          vgX = vgX + 1
          With istBeneficiarios(vgX)
                
'                'Formatear la Fecha de Nacimiento
'                If (Trim(.Fec_NacBen) <> "") Then
'                    vlCargaFecNac = DateSerial(CInt(Mid(.Fec_NacBen, 1, 4)), CInt(Mid(.Fec_NacBen, 5, 2)), CInt(Mid(.Fec_NacBen, 7, 2)))
'                Else
'                    vlCargaFecNac = ""
'                End If
'                'Formatear la Fecha de Nacimiento del Hijo Menor
'                If (Trim(.Fec_NacHM) <> "") Then
'                    vlCargaFecNacHM = DateSerial(CInt(Mid(.Fec_NacHM, 1, 4)), CInt(Mid(.Fec_NacHM, 5, 2)), CInt(Mid(.Fec_NacHM, 7, 2)))
'                Else
'                    vlCargaFecNacHM = ""
'                End If
'                'Formatear la Fecha de Invalidez
'                If (Trim(.Fec_InvBen) <> "") Then
'                    vlCargaFecInv = DateSerial(CInt(Mid(.Fec_InvBen, 1, 4)), CInt(Mid(.Fec_InvBen, 5, 2)), CInt(Mid(.Fec_InvBen, 7, 2)))
'                Else
'                    vlCargaFecInv = ""
'                End If
'
'                'Formatear la Fecha de Fallecimiento
'                If (Trim(.Fec_FallBen) <> "") Then
'                    vlCargaFecFall = DateSerial(CInt(Mid(.Fec_FallBen, 1, 4)), CInt(Mid(.Fec_FallBen, 5, 2)), CInt(Mid(.Fec_FallBen, 7, 2)))
'                Else
'                    vlCargaFecFall = ""
'                End If
'
'                'Formatear la Fecha de Suspención del Beneficiario
'                If (Trim(.Fec_SusBen) <> "") Then
'                    vlCargaFecSus = DateSerial(CInt(Mid(.Fec_SusBen, 1, 4)), CInt(Mid(.Fec_SusBen, 5, 2)), CInt(Mid(.Fec_SusBen, 7, 2)))
'                Else
'                    vlCargaFecSus = ""
'                End If
'
'                'Formatear la Fecha de Inicio de Pago de Pensiones
'                If (Trim(.Fec_IniPagoPen) <> "") Then
'                    vlCargaFecIniPagoPen = DateSerial(CInt(Mid(.Fec_IniPagoPen, 1, 4)), CInt(Mid(.Fec_IniPagoPen, 5, 2)), CInt(Mid(.Fec_IniPagoPen, 7, 2)))
'                Else
'                    vlCargaFecIniPagoPen = ""
'                End If
'
'                'Formatear la Fecha de Termino de Pago del Periodo Garantizado
'                If (Trim(.Fec_TerPagoPenGar) <> "") Then
'                    vlCargaFecTerpagoPenGar = DateSerial(CInt(Mid(.Fec_TerPagoPenGar, 1, 4)), CInt(Mid(.Fec_TerPagoPenGar, 5, 2)), CInt(Mid(.Fec_TerPagoPenGar, 7, 2)))
'                Else
'                    vlCargaFecTerpagoPenGar = ""
'                End If
'
'                'Formatear la Fecha de Termino de Pago del Periodo Garantizado
'                If (Trim(.Fec_Matrimonio) <> "") Then
'                    vlCargaFecMat = DateSerial(CInt(Mid(.Fec_Matrimonio, 1, 4)), CInt(Mid(.Fec_Matrimonio, 5, 2)), CInt(Mid(.Fec_Matrimonio, 7, 2)))
'                Else
'                    vlCargaFecMat = ""
'                End If
'
'               'vgCodPar = " " & Trim(.Cod_Par) & " - " & Trim(fgBuscarGlosaElemento(vgCodTabla_Par, Trim(.Cod_Par)))
'
               If (iGrilla.TextMatrix(vgX, 0) = .Num_Orden) Then
               
               
'I--- ABV 10/08/2007 ---
                    vlPrcPension = (Trim(.Prc_Pension))
                    If (iNumMesGar > 0) Then
                        vlPrcPensionGar = vlPrcPension
                    Else
                        vlPrcPensionGar = "0"
                    End If
                    
                    'Corregir el Derecho a Crecer de la Cónyuge o Madre con Hijos => 11 ó 21
                    vlCodPar = iGrilla.TextMatrix(vgX, 1)
                    If vlCodPar = "11" Or vlCodPar = "21" Then
                        vlDerCre = iDerCrePol
                    Else
                        vlDerCre = .Cod_DerCre
                    End If
                    ''Corregir Porcentaje para la Cobertura de la Cónyuge (S/Hijos) definida en la Cotizazión
                    'If (vlCodCoberCon <> "0") And (vlCodCoberCon <> "") Then
                    '    If vlCodPar = "10" Then vlPrcPension = vlPrcFacPenElla
                    'End If
                    
                    'Corregir el Estado de Pago para la Pensión
                    vlFecNacBen = Format(iGrilla.TextMatrix(vgX, 9), "yyyymmdd")
                    If Trim(iGrilla.TextMatrix(vgX, 20)) <> "" Then
                        vlFecFallBen = Format(iGrilla.TextMatrix(vgX, 20), "yyyymmdd")
                    End If
                    vlSitInv = (iGrilla.TextMatrix(vgX, 4))
                    vlEstPen = fgCalcularEstadoPagoPension(iFecDev, iCodTipPen, vlCodPar, vlFecNacBen, vlFecFallBen, "", vlSitInv)
'F--- ABV 10/08/2007 ---
                    
                    iGrilla.TextMatrix(vgX, 7) = (Trim(.Cod_DerPen))
                    iGrilla.TextMatrix(vgX, 8) = vlDerCre '(Trim(.Cod_DerCre))
                    'iGrilla.TextMatrix(vgX, 10) = (vlCargaFecNacHM)
                    iGrilla.TextMatrix(vgX, 17) = (Trim(.Prc_Pension))
                    iGrilla.TextMatrix(vgX, 18) = (Trim(.Mto_Pension))
                    iGrilla.TextMatrix(vgX, 19) = (Trim(.Mto_PensionGar))
                    iGrilla.TextMatrix(vgX, 22) = vlEstPen '(Trim(.Cod_EstPension))
                    iGrilla.TextMatrix(vgX, 23) = vlPrcPensionGar '(Trim(.Prc_PensionGar))
                    iGrilla.TextMatrix(vgX, 24) = (Trim(.Prc_PensionLeg))
                    
'                    iGrilla.AddItem (.Num_Orden) & vbTab _
'                    & (" " & Format((Trim(.Rut_Ben)), "##,###,##0") & " - " & (Trim(.Dgv_Ben))) & vbTab _
'                    & (Trim(.Gls_NomBen)) & vbTab & (Trim(.Gls_PatBen)) & vbTab & (Trim(.Gls_MatBen)) & vbTab _
'                    & (Trim(.Cod_Par)) & vbTab _
'                    & (Trim(.Cod_GruFam)) & vbTab _
'                    & (Trim(.Cod_Sexo)) & vbTab _
'                    & (Trim(.Cod_SitInv)) & vbTab _
'                    & (Trim(.Cod_EstPension)) & vbTab _
'                    & (Trim(.Cod_DerCre)) & vbTab _
'                    & (Trim(.Num_Poliza)) & vbTab & (Trim(.num_endoso)) & vbTab _
'                    & (Trim(.Cod_CauInv)) & vbTab _
'                    & (vlCargaFecNac) & vbTab _
'                    & (vlCargaFecNacHM) & vbTab _
'                    & (vlCargaFecInv) & vbTab _
'                    & (Trim(.Mto_Pension)) & vbTab _
'                    & (Trim(.Prc_Pension)) & vbTab _
'                    & (vlCargaFecFall) & vbTab _
'                    & (Trim(.Cod_DerPen)) & vbTab _
'                    & (Trim(.Cod_MotReqPen)) & vbTab _
'                    & (Trim(.Mto_PensionGar)) & vbTab _
'                    & (Trim(.Cod_CauSusBen)) & vbTab _
'                    & (vlCargaFecSus) & vbTab _
'                    & (vlCargaFecIniPagoPen) & vbTab _
'                    & (vlCargaFecTerpagoPenGar) & vbTab _
'                    & (vlCargaFecMat)
               End If
          End With
    Wend

Exit Function
Err_flCargaGrillaBeneficiarios:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgObtenerPorcCobertura(iTipoPension As String, iFecha As String, oRemuneracion As Double, oPrcCobertura As Double) As Boolean
Dim vlRegCober As ADODB.Recordset

    fgObtenerPorcCobertura = False
    oRemuneracion = 0
    oPrcCobertura = 0

    vgSql = "SELECT mto_rem as remuneracion, prc_cober as porcentaje "
    vgSql = vgSql & "FROM ma_tcod_cober WHERE "
    vgSql = vgSql & "cod_tippension = '" & iTipoPension & "' "
    vgSql = vgSql & "AND (fec_inivig <= '" & iFecha & "' "
    vgSql = vgSql & "AND fec_tervig >= '" & iFecha & "') "
    Set vlRegCober = vgConexionBD.Execute(vgSql)
    If Not (vlRegCober.EOF) Then
        If Not IsNull(vlRegCober!remuneracion) Then oRemuneracion = (vlRegCober!remuneracion)
        If Not IsNull(vlRegCober!porcentaje) Then oPrcCobertura = (vlRegCober!porcentaje)
        
        fgObtenerPorcCobertura = True
    End If
    vlRegCober.Close

End Function

Function fgCalcularEstadoPagoPension(iFechaIniVig As String, iTipoPension As String, iParentesco As String, iFechaNacimiento As String, iFechaFallecimiento As String, iFechaMatrimonio As String, iSitInv As String) As String
'Función: Permite calcular el Estado de Pago de la Pensión para un Beneficiario
'         Es decir, si le corresponde pago o no
'Parámetros de Entrada:
'iFechaIniVig     => Fecha de Inicio de Vigencia de la Póliza
'iTipoPension     => Tipo de Pensión de la Póliza
'iParentesco      => Monto de la Pensión de Referencia utilizada para el Calculo de la Pensión si el campo anterior esta en Verdadero
'iFechaFallecimiento => Fecha de Fallecimiento del Beneficiario
'iFechaNacimiento => Fecha de Nacimiento del Beneficiario
'iSitInv          => Situación de Invalidez del Beneficiario
'Parámetros de Salida:
'Estado de Pensión => Estado de Pago de la Pensión
'-------------------------------------------------------------
'Fecha de Creación     : 07/08/2007 - ABV
'Fecha de Modificación :
'-------------------------------------------------------------

Dim Iaap   As Integer, Immp   As Integer, Iddp   As Integer
Dim fecha_sin As Long
Dim L24 As Long, L18 As Long
Dim edad_mes_ben As Long
Dim Nanbe  As Integer, Nmnbe  As Integer, Ndnbe  As Integer
Dim Codrel As Integer, derpen As Integer, Estpen As Integer
Dim Inv    As String, Coinbe As String
Dim vlFechaFallecimiento As String
Dim vlFechaMatrimonio    As String

    fgCalcularEstadoPagoPension = cgEstPension_NoPago
    
    L24 = 0
    'Debiera tomar la Fecha de Devengue
    If (fgCarga_Param("LI", "L24", iFechaIniVig) = True) Then
        L24 = vgValorParametro
    Else
        vgError = 1000
        MsgBox "No existe Edad de tope para los 24 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    'Mensualizar la Edad de 24 Años
    L24 = L24 * 12
    
    L18 = 0
    'Edad de 18 años
    If (fgCarga_Param("LI", "L18", iFechaIniVig) = True) Then
        L18 = vgValorParametro
    Else
        vgError = 1000
        MsgBox "No existe Edad de tope para los 18 años.", vbCritical, "Proceso Cancelado"
        Exit Function
    End If
    'Mensualizar la Edad de 18 Años
    L18 = L18 * 12

    Iaap = CInt(Mid(iFechaIniVig, 1, 4)) 'a Fecha de siniestro
    Immp = CInt(Mid(iFechaIniVig, 5, 2)) 'm Fecha de siniestro
    Iddp = CInt(Mid(iFechaIniVig, 7, 2)) 'd Fecha de siniestro
    
    fecha_sin = Iaap * 12 + Immp
    
    'Parentesco
    Codrel = iParentesco
    
    'Situación de Invalidez
    Inv = iSitInv
    
    'Fecha de Nacimiento
    Nanbe = CInt(Mid(iFechaNacimiento, 1, 4)) 'a Fecha de nacimiento
    Nmnbe = CInt(Mid(iFechaNacimiento, 5, 2)) 'm Fecha de nacimiento
    Ndnbe = CInt(Mid(iFechaNacimiento, 7, 2)) 'd Fecha de nacimiento
        
    'Codificación de Situación de Invalidez
    If Inv = "P" Then Coinbe = "P"
    If Inv = "T" Then Coinbe = "T"
    If Inv = "N" Then Coinbe = "N"
    
    'Edad en Meses
    edad_mes_ben = fecha_sin - (Nanbe * 12 + Nmnbe)
    
    'Fecha de Fallecimiento
    vlFechaFallecimiento = iFechaFallecimiento
    
    'Fecha de Matrimonio
    vlFechaMatrimonio = iFechaMatrimonio
    
    'Derecho Pensión
    derpen = cgEstPension_NoPago
    Estpen = cgEstPension_NoPago

    'Determinar su Estado de Pago Pensión
    Select Case Codrel
    Case 99, 0:
        'Si es el Causante
        vgPalabra = iTipoPension
        'Valida que sea un caso de Invalidez o Vejez
        vgI = InStr(1, cgPensionInvVejez, vgPalabra)
        If (vgI <> 0) Then
            derpen = 99
            Estpen = 99
        Else
            derpen = 10
            Estpen = 10
        End If
    Case 30 To 40:
        'Si el Beneficiario es un Hijo
        If (vlFechaFallecimiento <> "") Or (vlFechaMatrimonio <> "") Then
            derpen = 10
            Estpen = 10
        Else
            If edad_mes_ben > L24 And Coinbe = "N" Then
                derpen = 10
                Estpen = 10
            Else
                vgPalabra = iTipoPension
                'Valida que sea un caso de Invalidez o Vejez
                vgI = InStr(1, cgPensionInvVejez, vgPalabra)
                If (vgI <> 0) Then
                    derpen = 99
                    Estpen = 10
                Else
                    derpen = 99
                    Estpen = 99
                End If
            End If
        End If
    Case Else
        'El resto de los Beneficiarios que no sean Hijos, solo se dejan como
        'Sin Derecho a Pensión cuando están fallecidos o el Tipo de Pensión indica Causante Vivo
        If (vlFechaFallecimiento <> "") Then
            derpen = 10
            Estpen = 10
        Else
            vgPalabra = iTipoPension
            'Valida que sea un caso de Invalidez o Vejez
            vgI = InStr(1, cgPensionInvVejez, vgPalabra)
            If (vgI <> 0) Then
                derpen = 99
                Estpen = 10
            Else
                derpen = 99
                Estpen = 99
            End If
        End If
    End Select
    
    fgCalcularEstadoPagoPension = Estpen
End Function

Function fgObtenerRentabilidadAFP(iCodAFP As String, iFecCalculo As String, oRentabilidad As Double) As Boolean
'Función: Permite obtener la Rentabilidad de la AFP consultada
'Parámetros de Entrada:
'- iCodAFP       => Código de la AFP
'- iFecCalculo   => Fecha a la cual se solicita la Rentabilidad ("yyyymmdd")
'Parámetros de Salida:
'- Retorna un verdadero o falso si existe o no el valor consultado en Parámetros
'- oRentabilidad => contiene la Rentabilidad de la AFP
'----------------------------------------------------------
'Fecha de Creación     : 13/08/2007
'Fecha de Modificación :
'----------------------------------------------------------
Dim vlRegRenta As ADODB.Recordset
    
    fgObtenerRentabilidadAFP = False
    oRentabilidad = 0
    
    vlSql = "SELECT mto_elemento as valor "
    vlSql = vlSql & "FROM ma_tpar_tabcodvig "
    vlSql = vlSql & "WHERE cod_tabla = '" & vgCodTabla_AFP & "' AND "
    vlSql = vlSql & "cod_elemento = '" & iCodAFP & "' AND "
    vlSql = vlSql & "fec_inivig <= '" & iFecCalculo & "' AND "
    vlSql = vlSql & "fec_tervig >= '" & iFecCalculo & "'"
    Set vlRegRenta = vgConexionBD.Execute(vlSql)
    If Not vlRegRenta.EOF Then
        If Not IsNull(vlRegRenta!valor) Then
            oRentabilidad = vlRegRenta!valor
            fgObtenerRentabilidadAFP = True
        End If
    End If
    vlRegRenta.Close
    
End Function

Function fgObtenerFactorAjusteIPC(iFechaDevengue As String, iFechaCotizacion As String) As Boolean
'Permite obtener el valor de Ajuste del IPC entre la Fecha de Devengue y Cálculo
'Parámetros de Entrada:
'- iFechaDevengue   => Fecha de Devengue de la Pensión (yyyymmdd)
'- iFechaCotizacion => Fecha de Cálculo de la Pensión (yyyymmdd)
'Parámetros de Salida:
'- Retorna un verdadero si existen los valores, y en caso contrario, un valor Falso
'- vgFactorAjusteIPC  => Asigna el valor del Factor de Ajuste calculado
'-----------------------------------------------------
'Fecha de Creación : 14/08/2007  ABV
'Fecha de Modificación :
'-----------------------------------------------------
Dim vlIpcDevengue As Double, vlIpcCotizacion As Double
Dim vlNum         As Integer
Dim vlSwFecDev    As Boolean, vlSwFecCot    As Boolean
Dim vlRegValor As ADODB.Recordset
Dim vlFechaCotizacionAnterior As String

    fgObtenerFactorAjusteIPC = False
    
    vgFactorAjusteIPC = 0
    
    'Calcula el Ipc a la Fecha de Devengue
    vlNum = CInt(Mid(iFechaDevengue, 5, 2))
    
    vlSql = "SELECT mto_ipc as valor "
    vlSql = vlSql & "FROM ma_tval_ipc "
    vlSql = vlSql & "WHERE "
    vlSql = vlSql & "fec_ipc = '" & (Mid(iFechaDevengue, 1, 6)) & "01' "
    Set vlRegValor = vgConexionBD.Execute(vlSql)
    If Not vlRegValor.EOF Then
        If Not IsNull(vlRegValor!valor) Then
            vlIpcDevengue = vlRegValor!valor
            vlSwFecDev = True
        End If
    End If
    vlRegValor.Close

    'Calcula el IPC a la Fecha de Cotización
    vlNum = CInt(Mid(iFechaCotizacion, 5, 2))
'I--- ABV 03/09/2007 ---
    'Se debe obtener el IPC del Mes Anterior al de la Operación
    vlFechaCotizacionAnterior = Format(DateSerial(Mid(iFechaCotizacion, 1, 4), CInt(Mid(iFechaCotizacion, 5, 2)) - 1, 1), "yyyymmdd")
'F--- ABV 03/09/2007 ---
    
    vlSql = "SELECT mto_ipc as valor "
    vlSql = vlSql & "FROM ma_tval_ipc "
    vlSql = vlSql & "WHERE "
'I--- ABV 03/09/2007 ---
'    vlSql = vlSql & "fec_ipc = '" & (Mid(iFechaCotizacion, 1, 6)) & "01' "
    vlSql = vlSql & "fec_ipc = '" & (Mid(vlFechaCotizacionAnterior, 1, 6)) & "01' "
'F--- ABV 03/09/2007 ---
    Set vlRegValor = vgConexionBD.Execute(vlSql)
    If Not vlRegValor.EOF Then
        If Not IsNull(vlRegValor!valor) Then
            vlIpcCotizacion = vlRegValor!valor
            vlSwFecCot = True
        End If
    End If
    vlRegValor.Close
    
    If (vlSwFecCot = False) Or (vlSwFecDev = False) Then
        Exit Function
    End If
    
    If (vlIpcDevengue <> 0) Then
        vgFactorAjusteIPC = Format(vlIpcCotizacion / vlIpcDevengue, "#0.000000")
    Else
        vgFactorAjusteIPC = 0
    End If
    
    fgObtenerFactorAjusteIPC = True

End Function

Function fgObtenerPorcentajeBenSocial(iFecha As String, oValor As Double) As Boolean
'Función : Permite obtener el Porcentaje por Beneficio Social del Intermediario
'Parámetros de Entrada:
'- iFecha      => Fecha a la cual se solicita la información (yyyymmdd)
'Parámetros de Salida:
'- Retorna un Falso o True de acuerdo a su existencia
'- oValor      => Permite guardar el Porcentaje buscado
Dim Tb_Por As ADODB.Recordset
Dim Sql    As String

    fgObtenerPorcentajeBenSocial = False
    oValor = 0
    
    Sql = "select mto_valor as valor_porcentaje "
    Sql = Sql & "from ma_tval_bensocial where "
    Sql = Sql & "fec_inivig <= '" & iFecha & "' AND "
    Sql = Sql & "fec_tervig >= '" & iFecha & "' "
    Set Tb_Por = vgConexionBD.Execute(Sql)
    If Not Tb_Por.EOF Then

        If Not IsNull(Tb_Por!Valor_Porcentaje) Then
            oValor = Tb_Por!Valor_Porcentaje
            
            fgObtenerPorcentajeBenSocial = True
        End If
    End If
    Tb_Por.Close
    
End Function

Function fgMaximo(param1, param2) As Variant
    If param1 > param2 Then
        fgMaximo = param1
    Else
        fgMaximo = param2
    End If
End Function


Function fgMinimo(param1, param2) As Variant
    If param1 > param2 Then
        fgMinimo = param2
    Else
        fgMinimo = param1
    End If
End Function

Function fgLimpiarVariablesGlobales()
    vgFechaIniMortalVit_F = ""
    vgFechaIniMortalTot_F = ""
    vgFechaIniMortalPar_F = ""
    vgFechaIniMortalBen_F = ""
    vgFechaIniMortalVit_M = ""
    vgFechaIniMortalTot_M = ""
    vgFechaIniMortalPar_M = ""
    vgFechaIniMortalBen_M = ""
    
    vgFechaFinMortalVit_F = ""
    vgFechaFinMortalTot_F = ""
    vgFechaFinMortalPar_F = ""
    vgFechaFinMortalBen_F = ""
    vgFechaFinMortalVit_M = ""
    vgFechaFinMortalTot_M = ""
    vgFechaFinMortalPar_M = ""
    vgFechaFinMortalBen_M = ""
    
    vgIndicadorTipoMovimiento_F = ""
    vgIndicadorTipoMovimiento_M = ""
    
    vgFechaAnterior = ""
    vgError = 0
End Function

Function fgCargarTablaMortalidad(ioPeriodo As String)
Dim vgCmb As ADODB.Recordset
Dim iRegistro As ADODB.Recordset
Dim iSql As String
On Error GoTo Err_Tabla
    
    'If (iFecha <> "") And (iPeriodo <> "") And (iSexo <> "") And (iTipoTabla <> "") Then
    If (ioPeriodo <> "") Then
        
        vgNumeroTotalTablas = 0
        
        vgQuery = "SELECT count(num_correlativo) as numero "
        vgQuery = vgQuery & "from MA_TVAL_MORTAL WHERE "
        vgQuery = vgQuery & "cod_tipoper = '" & ioPeriodo & "' "
        vgQuery = vgQuery & "and "
        vgQuery = vgQuery & "cod_tiptabmor in ("
        vgQuery = vgQuery & "'" & vgTipoTablaParcial & "',"
        vgQuery = vgQuery & "'" & vgTipoTablaTotal & "',"
        vgQuery = vgQuery & "'" & vgTipoTablaBeneficiario & "',"
        vgQuery = vgQuery & "'" & vgTipoTablaRentista & "') "
        ''vgQuery = vgQuery & "cod_sexo = '" & iSexo & "' and "
        ''vgQuery = vgQuery & "#" & Format(iFecha, "yyyy/mm/dd") & "# beetwen "
        ''vgQuery = vgQuery & "fec_ini and fec_ter and "
        'vgQuery = vgQuery & "and cod_estado = 'A' "
        Set vgCmb = vgConexionBD.Execute(vgQuery)
        If Not (vgCmb.EOF) Then
            If Not IsNull(vgCmb!Numero) Then
                vgNumeroTotalTablas = vgCmb!Numero
            End If
        End If
        vgCmb.Close
        
        If (vgNumeroTotalTablas <> 0) Then
            ReDim egTablaMortal(vgNumeroTotalTablas) As TypeTabla
            
            iSql = "SELECT num_correlativo,cod_tiptabmor,cod_sexo,fec_ini,fec_ter,gls_nombre,"
            iSql = iSql & "cod_tipogen,cod_tipoper,num_initabmor,num_tertabmor,prc_tasaint,cod_estado "
            iSql = iSql & ",cod_oficial "
            'I--- ABV 15/03/2005 ---
            iSql = iSql & ",cod_tipotabla,num_annobase,gls_descripcion "
            'F--- ABV 15/03/2005 ---
            'iSql = "SELECT * "
            iSql = iSql & "from MA_TVAL_MORTAL WHERE "
            iSql = iSql & "cod_tipoper = '" & ioPeriodo & "' "
            iSql = iSql & "and "
            iSql = iSql & "cod_tiptabmor in ("
            iSql = iSql & "'" & vgTipoTablaParcial & "',"
            iSql = iSql & "'" & vgTipoTablaTotal & "',"
            iSql = iSql & "'" & vgTipoTablaBeneficiario & "',"
            iSql = iSql & "'" & vgTipoTablaRentista & "') "
            ''vgQuery = vgQuery & "cod_sexo = '" & iSexo & "' and "
            ''vgQuery = vgQuery & "#" & Format(iFecha, "yyyy/mm/dd") & "# beetwen "
            ''vgQuery = vgQuery & "fec_ini and fec_ter and "
            'vgQuery = vgQuery & "and cod_estado = 'A' "
            'iSql = iSql & "ORDER BY gls_nombre,fec_ini "
            'Debug.Print iSql
            Set iRegistro = vgConexionBD.Execute(iSql)
            If Not (iRegistro.EOF) Then
                'iRegistro.MoveFirst
                'ReDim TablaMortal(vgCmb!Numero)
                vgX = 1
                While Not (iRegistro.EOF)
                    egTablaMortal(vgX).Correlativo = iRegistro!Num_Correlativo
                    egTablaMortal(vgX).TipoTabla = Trim(iRegistro!cod_tiptabmor)
                    egTablaMortal(vgX).Sexo = Trim(iRegistro!Cod_Sexo)
                    egTablaMortal(vgX).FechaIni = Trim(iRegistro!fec_ini)
                    egTablaMortal(vgX).FechaFin = Trim(iRegistro!fec_ter)
                    egTablaMortal(vgX).Nombre = Trim(iRegistro!gls_nombre)
                    egTablaMortal(vgX).TipoGenerar = Trim(iRegistro!cod_tipogen)
                    egTablaMortal(vgX).TipoPeriodo = Trim(iRegistro!cod_tipoper)
                    egTablaMortal(vgX).IniTab = iRegistro!num_initabmor
                    egTablaMortal(vgX).Fintab = iRegistro!num_tertabmor
                    egTablaMortal(vgX).Tasa = iRegistro!prc_tasaint
                    egTablaMortal(vgX).Oficial = Trim(iRegistro!cod_oficial)
                    egTablaMortal(vgX).Estado = Trim(iRegistro!Cod_Estado)
                    egTablaMortal(vgX).TipoMovimiento = IIf(IsNull(iRegistro!cod_tipotabla), "E", Trim(iRegistro!cod_tipotabla))
                    egTablaMortal(vgX).AñoBase = IIf(IsNull(iRegistro!num_annobase), "0", iRegistro!num_annobase)
                    egTablaMortal(vgX).Descripcion = Trim(iRegistro!gls_descripcion)
                    vgX = vgX + 1
                    iRegistro.MoveNext
                Wend
            End If
            iRegistro.Close
        End If
    End If
    
Exit Function
Err_Tabla:
    Screen.MousePointer = 0
    Select Case Err
        Case Else
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgComboMortalNombre(iFecha, iTipoTabla, iPeriodo, iSexo) As String
Dim iFechaCot As Long

    'vlCombo.Clear
    fgComboMortalNombre = ""
    
    If (iFecha <> "") And (iPeriodo <> "") And (iSexo <> "") And (iTipoTabla <> "") Then
        
        vgI = vgNumeroTotalTablas
        vgX = 1
        vgJ = 1
        'iFechaCot = Format(iFecha, "yyyymmdd") ABV 21/08/2007
        iFechaCot = iFecha
        
        Do While vgX <= vgI
        
            If (egTablaMortal(vgX).FechaIni <= iFechaCot) And _
               (egTablaMortal(vgX).FechaFin >= iFechaCot) And _
               (egTablaMortal(vgX).Sexo = iSexo) And _
               (egTablaMortal(vgX).TipoTabla = iTipoTabla) And _
               (egTablaMortal(vgX).TipoPeriodo = iPeriodo) _
               And (egTablaMortal(vgX).Estado = "A") Then
                
                If (egTablaMortal(vgX).Oficial = "S") Then
                    fgComboMortalNombre = egTablaMortal(vgX).Nombre
                    Exit Do
                '    vlCombo.AddItem egTablaMortal(vgX).Nombre, 0
                '    vlCombo.ItemData(0) = egTablaMortal(vgX).Correlativo
                'Else
                '    vlCombo.AddItem egTablaMortal(vgX).Nombre
                '    vgJ = vlCombo.ListCount - 1
                '    vlCombo.ItemData(vgJ) = egTablaMortal(vgX).Correlativo
                End If
                'vgJ = vgJ + 1
                'CInt(egTablaMortal(vgX).Correlativo)
                'vlCombo.List = egTablaMortal(vgX).Correlativo
            End If
            
            vgX = vgX + 1
        Loop
        
        'If (vlCombo.ListCount <> 0) Then
        '    vlCombo.ListIndex = 0
        'End If
    End If
    
End Function

Function fgBuscarMortalCodigo(iNombre) As String

    fgBuscarMortalCodigo = ""
    
    If (iNombre <> "") Then
        
        vgI = vgNumeroTotalTablas
        vgX = 1
        vgJ = 1
        
        Do While vgX <= vgI
            If (Trim(egTablaMortal(vgX).Nombre) = Trim(iNombre)) Then
                fgBuscarMortalCodigo = egTablaMortal(vgX).Correlativo
                Exit Do
            End If
            vgX = vgX + 1
        Loop
    
    End If
    
End Function

Function fgFinTab_Mortal(iCorrelativo As Long) As Long
    
    fgFinTab_Mortal = -1
    For vgI = 1 To vgNumeroTotalTablas
        If (egTablaMortal(vgI).Correlativo = iCorrelativo) Then
            fgFinTab_Mortal = egTablaMortal(vgI).Fintab
            Exit For
        End If
    Next vgI
    
End Function

Function fgBuscarMortalidadNormativa(iNavig, iNmvig, iNdvig, iNap, iNmp, iNdp, iSexoCau, iFechaNacCau) As Boolean
Dim vlResPregunta As Boolean

    fgBuscarMortalidadNormativa = False
    
    '1. Leer Tabla de Mortalidad de Rtas. Vitalicias Mujer
    vgBuscarMortalVit_F = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalVit_F = "") And (vgFechaFinMortalVit_F = "") Then
            vgBuscarMortalVit_F = "S"
        Else
            If (vgIndicadorTipoMovimiento_F = "E") Or (vgIndicadorTipoMovimiento_F <> "E" And iSexoCau = "M") Then
                vgBuscarMortalVit_F = "N"
            Else
                vgBuscarMortalVit_F = "S"
            End If
        End If
    Else
        If (vgFechaIniMortalVit_F <> "") And (vgFechaFinMortalVit_F <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalVit_F)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalVit_F)) Then
                If (vgIndicadorTipoMovimiento_F = "E") Or (vgIndicadorTipoMovimiento_F <> "E" And iSexoCau = "M") Then
                    vgBuscarMortalVit_F = "N"
                Else
                    vgBuscarMortalVit_F = "S"
                End If
            Else
                vgBuscarMortalVit_F = "S"
            End If
        Else
            vgBuscarMortalVit_F = "S"
        End If
    End If
    
    If (vgBuscarMortalVit_F = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 1
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalVit_F_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalVit_F) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                    
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalVit_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalVit_F = egTablaMortal(vgI).FechaFin
                        vgIndicadorTipoMovimiento_F = egTablaMortal(vgI).TipoMovimiento
                        vgDinamicaAñoBase_F = egTablaMortal(vgI).AñoBase
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Lx(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            If (vgIndicadorTipoMovimiento_F <> "D") Then
                                Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                                Sql = Sql & " where num_correlativo = " & vgs_Nro
                                Sql = Sql & " order by num_edad "
                                Set Tb2 = vgConexionBD.Execute(Sql)
                                If Not (Tb2.EOF) Then
                                    vgSw = True
                                    'tb2.MoveFirst
                                    'k = 1
                                    Do While Not Tb2.EOF
                                        k = Tb2!Edad
                                        'If h = 1 Then   'Causante
                                            Lx(i, j, k) = Tb2!mto_lx
                                        'Else    'Beneficiario
                                        '    ly(i, j, k) = tb2!mto_lx
                                        'End If
                                        'k = k + 1
                                        Tb2.MoveNext
                                    Loop
                                Else
                                    vgError = 1061
                                    Exit Function
                                End If
                                Tb2.Close
                                
                                Exit For
                            Else
                                'Obtener la Tabla Temporal Dinánica
                                'If (fgCrearMortalidadDinamica(iNavig, iNmvig, iNdvig, _
                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_F, iFechaNacCau) = False) Then
                                    
                                If (fgCrearMortalidadDinamica_DaniMensual(iNavig, iNmvig, iNdvig, _
                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_F, iFechaNacCau) = False) Then
                                    vgError = 1061
                                    Exit Function
                                Else
                                    vgSw = True
                                End If
                            End If
                        Else
                            vgError = 1061
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1061
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1061
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

    '2. Leer Tabla de Mortalidad de Inv. Totales Mujer
    vgBuscarMortalTot_F = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalTot_F = "") And (vgFechaFinMortalTot_F = "") Then
            vgBuscarMortalTot_F = "S"
        End If
    Else
        If (vgFechaIniMortalTot_F <> "") And (vgFechaFinMortalTot_F <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalTot_F)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalTot_F)) Then
                vgBuscarMortalTot_F = "N"
            Else
                vgBuscarMortalTot_F = "S"
            End If
        Else
            vgBuscarMortalTot_F = "S"
        End If
    End If
    
    If (vgBuscarMortalTot_F = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 1
                    If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalTot_F_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalTot_F) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalTot_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalTot_F = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1062
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1062
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1062
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1062
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '3. Leer Tabla de Mortalidad de Inv. Parciales Mujer
    vgBuscarMortalPar_F = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalPar_F = "") And (vgFechaFinMortalPar_F = "") Then
            vgBuscarMortalPar_F = "S"
        End If
    Else
        If (vgFechaIniMortalPar_F <> "") And (vgFechaFinMortalPar_F <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalPar_F)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalPar_F)) Then
                vgBuscarMortalPar_F = "N"
            Else
                vgBuscarMortalPar_F = "S"
            End If
        Else
            vgBuscarMortalPar_F = "S"
        End If
    End If
    
    If (vgBuscarMortalPar_F = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 3
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalPar_F_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalPar_F) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalPar_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalPar_F = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1063
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1063
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1063
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1063
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '4. Leer Tabla de Mortalidad de Beneficiarios Mujer
    vgBuscarMortalBen_F = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalBen_F = "") And (vgFechaFinMortalBen_F = "") Then
            vgBuscarMortalBen_F = "S"
        End If
    Else
        If (vgFechaIniMortalBen_F <> "") And (vgFechaFinMortalBen_F <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalBen_F)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalBen_F)) Then
                vgBuscarMortalBen_F = "N"
            Else
                vgBuscarMortalBen_F = "S"
            End If
        Else
            vgBuscarMortalBen_F = "S"
        End If
    End If
    
    If (vgBuscarMortalBen_F = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 2
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalBen_F_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalBen_F) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalBen_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalBen_F = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Ly(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    'If h = 1 Then   'Causante
                                    '    lx(i, j, k) = tb2!mto_lx
                                    'Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    'End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1064
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1064
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1064
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1064
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

'--------------------------------------------------------------------
    '5. Leer Tabla de Mortalidad de Rtas. Vitalicias Hombre
    vgBuscarMortalVit_M = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalVit_M = "") And (vgFechaFinMortalVit_M = "") Then
            vgBuscarMortalVit_M = "S"
        Else
            If (vgIndicadorTipoMovimiento_M = "E") Or (vgIndicadorTipoMovimiento_M <> "E" And iSexoCau = "F") Then
                vgBuscarMortalVit_M = "N"
            Else
                vgBuscarMortalVit_M = "S"
            End If
        End If
    Else
        If (vgFechaIniMortalVit_M <> "") And (vgFechaFinMortalVit_M <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalVit_M)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalVit_M)) Then
                If (vgIndicadorTipoMovimiento_M = "E") Or (vgIndicadorTipoMovimiento_M <> "E" And iSexoCau = "F") Then
                    vgBuscarMortalVit_M = "N"
                Else
                    vgBuscarMortalVit_M = "S"
                End If
            Else
                vgBuscarMortalVit_M = "S"
            End If
        Else
            vgBuscarMortalVit_M = "S"
        End If
    End If
    
    If (vgBuscarMortalVit_M = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 1
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalVit_M_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalVit_M) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalVit_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalVit_M = egTablaMortal(vgI).FechaFin
                        vgIndicadorTipoMovimiento_M = egTablaMortal(vgI).TipoMovimiento
                        vgDinamicaAñoBase_M = egTablaMortal(vgI).AñoBase
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Lx(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            If (vgIndicadorTipoMovimiento_M <> "D") Then
                                Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                                Sql = Sql & " where num_correlativo = " & vgs_Nro
                                Sql = Sql & " order by num_edad "
                                Set Tb2 = vgConexionBD.Execute(Sql)
                                If Not (Tb2.EOF) Then
                                    vgSw = True
                                    'tb2.MoveFirst
                                    'k = 1
                                    Do While Not Tb2.EOF
                                        k = Tb2!Edad
                                        'If h = 1 Then   'Causante
                                            Lx(i, j, k) = Tb2!mto_lx
                                        'Else    'Beneficiario
                                        '    ly(i, j, k) = tb2!mto_lx
                                        'End If
                                        'k = k + 1
                                        Tb2.MoveNext
                                    Loop
                                Else
                                    vgError = 1065
                                    Exit Function
                                End If
                                Tb2.Close
                                
                                Exit For
                            Else
                                'Obtener la Tabla Temporal Dinánica
                                'If (fgCrearMortalidadDinamica(iNavig, iNmvig, iNdvig, _
                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_M, iFechaNacCau) = False) Then
                                    
                                If (fgCrearMortalidadDinamica_DaniMensual(iNavig, iNmvig, iNdvig, _
                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_M, iFechaNacCau) = False) Then
                                    vgError = 1061
                                    Exit Function
                                Else
                                    vgSw = True
                                End If
                            End If
                        Else
                            vgError = 1065
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1065
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1065
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

    '6. Leer Tabla de Mortalidad de Inv. Totales Hombre
    vgBuscarMortalTot_M = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalTot_M = "") And (vgFechaFinMortalTot_M = "") Then
            vgBuscarMortalTot_M = "S"
        End If
    Else
        If (vgFechaIniMortalTot_M <> "") And (vgFechaFinMortalTot_M <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalTot_M)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalTot_M)) Then
                vgBuscarMortalTot_M = "N"
            Else
                vgBuscarMortalTot_M = "S"
            End If
        Else
            vgBuscarMortalTot_M = "S"
        End If
    End If
    
    If (vgBuscarMortalTot_M = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 1
                    If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalTot_M_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalTot_M) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalTot_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalTot_M = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1066
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1066
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1066
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1066
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '7. Leer Tabla de Mortalidad de Inv. Parciales Hombre
    vgBuscarMortalPar_M = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalPar_M = "") And (vgFechaFinMortalPar_M = "") Then
            vgBuscarMortalPar_M = "S"
        End If
    Else
        If (vgFechaIniMortalPar_M <> "") And (vgFechaFinMortalPar_M <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalPar_M)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalPar_M)) Then
                vgBuscarMortalPar_M = "N"
            Else
                vgBuscarMortalPar_M = "S"
            End If
        Else
            vgBuscarMortalPar_M = "S"
        End If
    End If
    
    If (vgBuscarMortalPar_M = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 3
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalPar_M_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalPar_M) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalPar_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalPar_M = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1067
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1067
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1067
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1067
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '8. Leer Tabla de Mortalidad de Beneficiarios Hombre
    vgBuscarMortalBen_M = ""
    If (vgUtilizarNormativa = "S") Then
        If (vgFechaIniMortalBen_M = "") And (vgFechaFinMortalBen_M = "") Then
            vgBuscarMortalBen_M = "S"
        End If
    Else
        If (vgFechaIniMortalBen_M <> "") And (vgFechaFinMortalBen_M <> "") Then
            If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalBen_M)) And _
            ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalBen_M)) Then
                vgBuscarMortalBen_M = "N"
            Else
                vgBuscarMortalBen_M = "S"
            End If
        Else
            vgBuscarMortalBen_M = "S"
        End If
    End If
    
    If (vgBuscarMortalBen_M = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 2
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    vlResPregunta = False
                    If (vgUtilizarNormativa = "S") Then
                        If (egTablaMortal(vgI).Correlativo = vgMortalBen_M_Normativa) Then
                            vlResPregunta = True
                        End If
                    Else
                        If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                        (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                        (egTablaMortal(vgI).Nombre = vgPalabra_MortalBen_M) And _
                        (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                        (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                            vlResPregunta = True
                        End If
                    End If
                    
                    If (vlResPregunta = True) Then
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalBen_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalBen_M = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Ly(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from MA_TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    'If h = 1 Then   'Causante
                                    '    lx(i, j, k) = tb2!mto_lx
                                    'Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    'End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1068
                                Exit Function
                            End If
                            Tb2.Close
                            
                            Exit For
                        Else
                            vgError = 1068
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1068
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1068
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

    fgBuscarMortalidadNormativa = True
End Function

Function fgCrearMortalidadDinamica_DaniMensual(iNavig, iNmvig, iNdvig, _
    iab, imb, idb, iSexo, iInval, iCorrelativo, iFinTab, iAñoBase, iFNac) As Boolean
'Función: permite generar la Tabla de Mortalidad Dinámica desde la descripción de una
'         Tabla Anual
'Parámetros de Entrada:
'   - iNavig = Año de Vigencia de la Tabla de Mortalidad
'   - iNmvig = Mes de Vigencia de la Tabla de Mortalidad
'   - iNdvig = Día de Vigencia de la Tabla de Mortalidad
'   - iab    = Año de Proceso
'   - imb    = Mes de Proceso
'   - idb    = Día de Proceso
'   - iSexo  = Sexo a generar
'   - iInval = Invalidez a generar
'   - iCorrelativo = Número de la Tabla de Mortalidad a Leer
'   - iFinTab = Valor que indica el termino de la Tabla de Mortalidad
'   - iAñoBase = Indica el Año Base desde el cual se genera la Tabla Dinámica
'Valor de Salida:
'   Retorna un valor True o False si pudo realizar la Actualización de la Matriz

'Dim tledad(0 To 90), qx(0 To 90), facmejor(0 To 90)
'Dim Lxm(1 To 2, 1 To 3, 1 To 1332) As Double
Dim Lxm(1 To 1332) As Double

Dim QxModif(0 To 120) As Double
Dim FacMejor(0 To 120) As Double, TlEdad(0 To 120) As Double
Dim qx(0 To 120) As Double
Dim AñoBase As Long
Dim AñoProceso As Integer, MesProceso As Integer, DiaProceso As Integer
Dim AñoNac     As Integer, MesNac     As Integer, DiaNac     As Integer
Dim QxMensModif As Double, Parte1 As Double, Parte2 As Double
Dim i As Long, j As Long, k As Long
Dim Tb2 As ADODB.Recordset
Dim FProceso As String
Dim Edad As Long, edaca As Long
Dim difdia As Integer
Dim factor_exponente(0 To 2000) As Double

'On Error GoTo Err_Mortal

    fgCrearMortalidadDinamica_DaniMensual = False
'    imb = 12
'    idb = 31
'    iFNac = "19401231"
    
    AñoBase = iAñoBase

    FProceso = Format(iab, "0000") & Format(imb, "00") & Format(idb, "00")
    AñoProceso = iab
    MesProceso = imb
    DiaProceso = idb
    Fechap = AñoProceso * 12 + MesProceso

    AñoNac = Mid(iFNac, 1, 4)
    MesNac = Mid(iFNac, 5, 2)
    DiaNac = Mid(iFNac, 7, 2)
    Fechan = AñoNac * 12 + MesNac

    Edad = Fechap - Fechan

    difdia = idb - DiaNac
    If difdia > 15 Then Edad = Edad + 1
    If Edad <= 240 Then Edad = 240
    If Edad > (110 * 12) Then
        vgError = 1023
        Exit Function
    End If
    edaca = Fix(Edad / 12)
'    For i = 0 To 1332
'        factor_exponente(i) = 0
'    Next i
    'Lectura de tabla de mortalidad
    vgSql = "SELECT num_edad AS edad, mto_qx AS qx, prc_factor AS factor "
    vgSql = vgSql & "FROM ma_tval_mordet "
    vgSql = vgSql & "WHERE num_correlativo = " & vgs_Nro & " "
    vgSql = vgSql & "ORDER BY num_edad "
    Set Tb2 = vgConexionBD.Execute(vgSql)
    If Not (Tb2.EOF) Then
        Do While Not Tb2.EOF
            k = Tb2!Edad
            TlEdad(k) = Tb2!Edad
            qx(k) = Tb2!qx
            FacMejor(k) = Tb2!factor
            'FacMejor(k) = 0
            Tb2.MoveNext
        Loop
    Else
        vgError = 1061
        Tb2.Close
        Exit Function
    End If
    Tb2.Close
    '***************************************************************************
    'Modificar la aplicacion el exponenete del factor de mejoramiento
    'Ej. pol 1446 interrrentas, año 2004 son 3 meses pasar al factor 2
    incremento = 0
    indice = 0
    inicio = 0
    For i = edaca To 110
        If i = edaca Then
            inicio = imb + 1
            If inicio > 12 Then inicio = 1
        End If
        For j = inicio To 12
            indice = indice + 1
            k = Edad + indice - 1
            'If j = 1 And i <> edaca Then
            If j = 1 Then
                incremento = incremento + Fix(j / 12) + 1
            Else
                incremento = incremento
            End If
            factor_exponente(k) = incremento + (AñoProceso - AñoBase)
            
        Next j
        inicio = 1
    Next i
    QxMensModif = 0
    For i = edaca To 110
        For j = 1 To 12
            k = (i * 12) + j - 1
            If k < Edad Then
            Else
                If k = Edad Then
                    Lxm(k) = 100000
                Else
                    Lxm(k) = Lxm(k - 1) - (Lxm(k - 1) * QxMensModif)
                End If
                QxModif(i) = qx(i) * (1 - FacMejor(i)) ^ factor_exponente(k)
                Parte1 = ((1 / 12) * QxModif(i))
                Parte2 = (k / 12 - Fix(k / 12))
                If ((1 - Parte2 * QxModif(i)) = 0) Then
                    QxMensModif = 0
                Else
                    QxMensModif = Parte1 / (1 - Parte2 * QxModif(i))
                End If
                If k > (110 * 12) Then Exit For
                Lx(iSexo, iInval, k) = Lxm(k)
'                'Borra - Daniela
'                vgQuery = "INSERT into DANI (poliza,agno ,edad,qx,qxanualmodif,qxanual,MEJORA,lx,sexo) values( "
'                vgQuery = vgQuery & "'1', "
'                vgQuery = vgQuery & Str(Format(AñoProceso, "#0")) & ", "
'                vgQuery = vgQuery & Str(Format(k, "#000")) & ", "
'                vgQuery = vgQuery & Str(Format(QxMensModif, "#0.0000000000000")) & ", "
'                vgQuery = vgQuery & Str(Format(QxModif(i), "#0.0000000000000")) & ", "
'                vgQuery = vgQuery & Str(Format(Qx(i), "#0.0000000000000")) & ", "
'                vgQuery = vgQuery & Str(Format(FacMejor(i), "#0.0000000000000")) & ", "
'                vgQuery = vgQuery & Str(Format(Lxm(k), "#000.000000000")) & ", "
'                vgQuery = vgQuery & Str(Format(iSexo, "#0")) & ") "
'                vgConexionBD.Execute (vgQuery)

'                'vgQuery = "update DANI set  "
'                'vgQuery = vgQuery & "agno = " & Str(Format(iNavig, "#0")) & ", "
'                'vgQuery = vgQuery & "edad = " & Str(Format(k, "#000")) & ", "
'                'vgQuery = vgQuery & "qx = " & Str(Format(QxMensModif, "#0.0000000")) & ", "
'                'vgQuery = vgQuery & "lx = " & Str(Format(Lxm(iSexo, iInval, k), "#000.000")) & " "
'                'vgConexionBD.Execute (vgQuery)
'                'End If
             End If
        Next j
    Next i

    fgCrearMortalidadDinamica_DaniMensual = True

Exit Function   'Buscar otra Póliza a calcular
Err_Mortal:
    'Screen.MousePointer = 0
    Select Case Err
        Case Else
        'ProgressBar.Value = 0
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgCrearMortalidadDinamica_DaniAnual(iNavig, iNmvig, iNdvig, _
    iab, imb, idb, iSexo, iInval, iCorrelativo, iFinTab, iAñoBase, iFNac) As Boolean
'Función: permite generar la Tabla de Mortalidad Dinámica desde la descripción de una
'         Tabla Anual
'Parámetros de Entrada:
'   - iNavig = Año de Vigencia de la Tabla de Mortalidad
'   - iNmvig = Mes de Vigencia de la Tabla de Mortalidad
'   - iNdvig = Día de Vigencia de la Tabla de Mortalidad
'   - iab    = Año de Proceso
'   - imb    = Mes de Proceso
'   - idb    = Día de Proceso
'   - iSexo  = Sexo a generar
'   - iInval = Invalidez a generar
'   - iCorrelativo = Número de la Tabla de Mortalidad a Leer
'   - iFinTab = Valor que indica el termino de la Tabla de Mortalidad
'   - iAñoBase = Indica el Año Base desde el cual se genera la Tabla Dinámica
'Valor de Salida:
'   Retorna un valor True o False si pudo realizar la Actualización de la Matriz

'Dim tledad(0 To 90), qx(0 To 90), facmejor(0 To 90)
'Dim Lxm(1 To 2, 1 To 3, 1 To 1332) As Double
Dim Lxm(1 To 1332) As Double

Dim QxModif(0 To 120) As Double
Dim FacMejor(0 To 120) As Double, TlEdad(0 To 120) As Double
Dim qx(0 To 120) As Double
Dim AñoBase As Long
Dim AñoProceso As Integer, MesProceso As Integer, DiaProceso As Integer
Dim AñoNac     As Integer, MesNac     As Integer, DiaNac     As Integer
Dim QxMensModif As Double, Parte1 As Double, Parte2 As Double
Dim i As Long, j As Long, k As Long
Dim Tb2 As ADODB.Recordset
Dim FProceso As String
Dim Edad As Long, edaca As Long
Dim difdia As Integer

'On Error GoTo Err_Mortal

    fgCrearMortalidadDinamica_DaniAnual = False
    
    AñoBase = iAñoBase
    
    FProceso = Format(iab, "0000") & Format(imb, "00") & Format(idb, "00")
    AñoProceso = iab
    MesProceso = imb
    DiaProceso = idb
    Fechap = AñoProceso * 12 + MesProceso
    
    AñoNac = Mid(iFNac, 1, 4)
    MesNac = Mid(iFNac, 5, 2)
    DiaNac = Mid(iFNac, 7, 2)
    Fechan = AñoNac * 12 + MesNac
    
    Edad = Fechap - Fechan
    
    difdia = idb - DiaNac
    If difdia > 15 Then Edad = Edad + 1
    If Edad <= 240 Then Edad = 240
    If Edad > (110 * 12) Then
        vgError = 1023
        Exit Function
    End If
    edaca = Fix(Edad / 12)
    
    'Lectura de tabla de mortalidad
    vgSql = "SELECT num_edad AS edad, mto_qx AS qx, prc_factor AS factor "
    vgSql = vgSql & "FROM ma_tval_mordet "
    vgSql = vgSql & "WHERE num_correlativo = " & vgs_Nro & " "
    vgSql = vgSql & "ORDER BY num_edad "
    Set Tb2 = vgConexionBD.Execute(vgSql)
    If Not (Tb2.EOF) Then
        Do While Not Tb2.EOF
            k = Tb2!Edad
            TlEdad(k) = Tb2!Edad
            qx(k) = Tb2!qx
            FacMejor(k) = Tb2!factor
            Tb2.MoveNext
        Loop
    Else
        vgError = 1061
        Tb2.Close
        Exit Function
    End If
    Tb2.Close
    
    j = -1
    For i = edaca To 110 '- edaca)
        j = j + 1
        QxModif(i) = qx(i) * (1 - FacMejor(i)) ^ (j + (AñoProceso - AñoBase))
    Next i
    
    QxMensModif = 0
    For i = edaca To 110 '- edaca)
        If i = edaca Then
            Lxm(i) = 100000
        Else
            Lxm(i) = Lxm(i - 1) - (Lxm(i - 1) * QxMensModif)
        End If
        QxMensModif = QxModif(i)
                
        Lx(iSexo, iInval, i) = Lxm(i)
                
        If i >= 110 Then Exit For
                
'        'Borra - Daniela
'        vgQuery = "INSERT into DANI (poliza,agno ,edad,qx ,lx,sexo) values( "
'        vgQuery = vgQuery & "'1', "
'        vgQuery = vgQuery & Str(Format(AñoProceso, "#0")) & ", "
'        vgQuery = vgQuery & Str(Format(k, "#000")) & ", "
'        vgQuery = vgQuery & Str(Format(QxMensModif, "#0.0000000000000")) & ", "
'        vgQuery = vgQuery & Str(Format(Lxm(k), "#000.000000000")) & ", "
'        vgQuery = vgQuery & Str(Format(iSexo, "#0")) & ") "
'        vgConexionBD.Execute (vgQuery)
'
'        'vgQuery = "update DANI set  "
'        'vgQuery = vgQuery & "agno = " & Str(Format(iNavig, "#0")) & ", "
'        'vgQuery = vgQuery & "edad = " & Str(Format(k, "#000")) & ", "
'        'vgQuery = vgQuery & "qx = " & Str(Format(QxMensModif, "#0.0000000")) & ", "
'        'vgQuery = vgQuery & "lx = " & Str(Format(Lxm(iSexo, iInval, k), "#000.000")) & " "
'        'vgConexionBD.Execute (vgQuery)
    Next i
    
    fgCrearMortalidadDinamica_DaniAnual = True
    
Exit Function   'Buscar otra Póliza a calcular
Err_Mortal:
    'Screen.MousePointer = 0
    Select Case Err
        Case Else
        'ProgressBar.Value = 0
        MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
    End Select
End Function

Function fgBuscarMortalidad_Old(iNavig, iNmvig, iNdvig, iNap, iNmp, iNdp, iSexoCau, iFechaNacCau) As Boolean
'Correcciones realizadas para Rimac-Perú con Fecha 18/12/2005
    fgBuscarMortalidad = False
    
    'If (iNavig = 2003) Then
    '    iNdvig = iNdvig
    'End If
    
    '1. Leer Tabla de Mortalidad de Rtas. Vitalicias Mujer
    vgBuscarMortalVit_F = ""
    If (vgFechaIniMortalVit_F <> "") And (vgFechaFinMortalVit_F <> "") Then
        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalVit_F)) And _
        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalVit_F)) Then
'I--- ABV 18/12/2005 ---
'            If (vgIndicadorTipoMovimiento_F = "E") Or (vgIndicadorTipoMovimiento_F <> "E" And iSexoCau = "M") Then
                vgBuscarMortalVit_F = "N"
'            Else
'                vgBuscarMortalVit_F = "S"
'            End If
'F--- ABV 18/12/2005 ---
        Else
            vgBuscarMortalVit_F = "S"
        End If
    Else
        vgBuscarMortalVit_F = "S"
    End If
    
    If (vgBuscarMortalVit_F = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 1
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalVit_F) And _
                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                    
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalVit_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalVit_F = egTablaMortal(vgI).FechaFin
                        vgIndicadorTipoMovimiento_F = egTablaMortal(vgI).TipoMovimiento
                        vgDinamicaAñoBase_F = egTablaMortal(vgI).AñoBase
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Lx(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            If (vgIndicadorTipoMovimiento_F <> "D") Then
                                Sql = "Select num_edad AS edad,mto_lx from TVAL_MORDET "
                                Sql = Sql & " where num_correlativo = " & vgs_Nro
                                Sql = Sql & " order by num_edad "
                                Set Tb2 = vgConexionBD.Execute(Sql)
                                If Not (Tb2.EOF) Then
                                    vgSw = True
                                    'tb2.MoveFirst
                                    'k = 1
                                    Do While Not Tb2.EOF
                                        k = Tb2!Edad
                                        'If h = 1 Then   'Causante
                                            Lx(i, j, k) = Tb2!mto_lx
                                        'Else    'Beneficiario
                                        '    ly(i, j, k) = tb2!mto_lx
                                        'End If
                                        'k = k + 1
                                        Tb2.MoveNext
                                    Loop
                                Else
                                    vgError = 1061
                                    Exit Function
                                End If
                                Tb2.Close
                            Else
                                'Obtener la Tabla Temporal Dinánica
                                If (fgCrearMortalidadDinamica_DaniMensual(iNavig, iNmvig, iNdvig, _
                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_F, iFechaNacCau) = False) Then
                                    vgError = 1061
                                    Exit Function
                                Else
                                    vgSw = True
                                End If
                            End If
                        Else
                            vgError = 1061
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1061
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1061
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

    '2. Leer Tabla de Mortalidad de Inv. Totales Mujer
    vgBuscarMortalTot_F = ""
    If (vgFechaIniMortalTot_F <> "") And (vgFechaFinMortalTot_F <> "") Then
        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalTot_F)) And _
        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalTot_F)) Then
            vgBuscarMortalTot_F = "N"
        Else
            vgBuscarMortalTot_F = "S"
        End If
    Else
        vgBuscarMortalTot_F = "S"
    End If
    
    If (vgBuscarMortalTot_F = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 1
                    If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalTot_F) And _
                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                    
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalTot_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalTot_F = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1062
                                Exit Function
                            End If
                            Tb2.Close
                        Else
                            vgError = 1062
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1062
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1062
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '3. Leer Tabla de Mortalidad de Inv. Parciales Mujer
    vgBuscarMortalPar_F = ""
    If (vgFechaIniMortalPar_F <> "") And (vgFechaFinMortalPar_F <> "") Then
        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalPar_F)) And _
        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalPar_F)) Then
            vgBuscarMortalPar_F = "N"
        Else
            vgBuscarMortalPar_F = "S"
        End If
    Else
        vgBuscarMortalPar_F = "S"
    End If
    
    If (vgBuscarMortalPar_F = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 3
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalPar_F) And _
                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                    
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalPar_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalPar_F = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1063
                                Exit Function
                            End If
                            Tb2.Close
                        Else
                            vgError = 1063
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1063
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1063
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '4. Leer Tabla de Mortalidad de Beneficiarios Mujer
    vgBuscarMortalBen_F = ""
    If (vgFechaIniMortalBen_F <> "") And (vgFechaFinMortalBen_F <> "") Then
        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalBen_F)) And _
        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalBen_F)) Then
            vgBuscarMortalBen_F = "N"
        Else
            vgBuscarMortalBen_F = "S"
        End If
    Else
        vgBuscarMortalBen_F = "S"
    End If
    
    If (vgBuscarMortalBen_F = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 2
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 2
                vgs_Sexo = "F"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalBen_F) And _
                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                    
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalBen_F = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalBen_F = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Ly(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    'If h = 1 Then   'Causante
                                    '    lx(i, j, k) = tb2!mto_lx
                                    'Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    'End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1064
                                Exit Function
                            End If
                            Tb2.Close
                        Else
                            vgError = 1064
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1064
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1064
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

'--------------------------------------------------------------------
    '5. Leer Tabla de Mortalidad de Rtas. Vitalicias Hombre
    vgBuscarMortalVit_M = ""
    If (vgFechaIniMortalVit_M <> "") And (vgFechaFinMortalVit_M <> "") Then
        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalVit_M)) And _
        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalVit_M)) Then
'I--- ABV 18/12/2005 ---
'            If (vgIndicadorTipoMovimiento_M = "E") Or (vgIndicadorTipoMovimiento_M <> "E" And iSexoCau = "F") Then
                vgBuscarMortalVit_M = "N"
'            Else
'                vgBuscarMortalVit_M = "S"
'            End If
'F--- ABV 18/12/2005 ---
        Else
            vgBuscarMortalVit_M = "S"
        End If
    Else
        vgBuscarMortalVit_M = "S"
    End If
    
    If (vgBuscarMortalVit_M = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 1
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalVit_M) And _
                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                    
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalVit_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalVit_M = egTablaMortal(vgI).FechaFin
                        vgIndicadorTipoMovimiento_M = egTablaMortal(vgI).TipoMovimiento
                        vgDinamicaAñoBase_M = egTablaMortal(vgI).AñoBase
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Lx(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            If (vgIndicadorTipoMovimiento_M <> "D") Then
                                Sql = "Select num_edad AS edad,mto_lx from TVAL_MORDET "
                                Sql = Sql & " where num_correlativo = " & vgs_Nro
                                Sql = Sql & " order by num_edad "
                                Set Tb2 = vgConexionBD.Execute(Sql)
                                If Not (Tb2.EOF) Then
                                    vgSw = True
                                    'tb2.MoveFirst
                                    'k = 1
                                    Do While Not Tb2.EOF
                                        k = Tb2!Edad
                                        'If h = 1 Then   'Causante
                                            Lx(i, j, k) = Tb2!mto_lx
                                        'Else    'Beneficiario
                                        '    ly(i, j, k) = tb2!mto_lx
                                        'End If
                                        'k = k + 1
                                        Tb2.MoveNext
                                    Loop
                                Else
                                    vgError = 1065
                                    Exit Function
                                End If
                                Tb2.Close
                            Else
                                'Obtener la Tabla Temporal Dinánica
                                If (fgCrearMortalidadDinamica_DaniMensual(iNavig, iNmvig, iNdvig, _
                                iNap, iNmp, iNdp, i, j, vgs_Nro, Fintab, vgDinamicaAñoBase_M, iFechaNacCau) = False) Then
                                    vgError = 1061
                                    Exit Function
                                Else
                                    vgSw = True
                                End If
                            End If
                        Else
                            vgError = 1065
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1065
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1065
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

    '6. Leer Tabla de Mortalidad de Inv. Totales Hombre
    vgBuscarMortalTot_M = ""
    If (vgFechaIniMortalTot_M <> "") And (vgFechaFinMortalTot_M <> "") Then
        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalTot_M)) And _
        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalTot_M)) Then
            vgBuscarMortalTot_M = "N"
        Else
            vgBuscarMortalTot_M = "S"
        End If
    Else
        vgBuscarMortalTot_M = "S"
    End If
    
    If (vgBuscarMortalTot_M = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 1
                    If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalTot_M) And _
                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                    
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalTot_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalTot_M = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1066
                                Exit Function
                            End If
                            Tb2.Close
                        Else
                            vgError = 1066
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1066
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1066
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '7. Leer Tabla de Mortalidad de Inv. Parciales Hombre
    vgBuscarMortalPar_M = ""
    If (vgFechaIniMortalPar_M <> "") And (vgFechaFinMortalPar_M <> "") Then
        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalPar_M)) And _
        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalPar_M)) Then
            vgBuscarMortalPar_M = "N"
        Else
            vgBuscarMortalPar_M = "S"
        End If
    Else
        vgBuscarMortalPar_M = "S"
    End If
    
    If (vgBuscarMortalPar_M = "S") Then
        For h = 1 To 2  '1=Causante '2=Beneficiario
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 3
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    'If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalPar_M) And _
                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                    
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalPar_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalPar_M = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        If (h = 1) Then
                            For vgX = 1 To Fintab
                                Lx(i, j, vgX) = 0
                            Next vgX
                        Else
                            For vgX = 1 To Fintab
                                Ly(i, j, vgX) = 0
                            Next vgX
                        End If
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    If h = 1 Then   'Causante
                                        Lx(i, j, k) = Tb2!mto_lx
                                    Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1067
                                Exit Function
                            End If
                            Tb2.Close
                        Else
                            vgError = 1067
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1067
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1067
                    Exit Function
                End If
                'Next j
            'Next i
        Next h
    End If

    '8. Leer Tabla de Mortalidad de Beneficiarios Hombre
    vgBuscarMortalBen_M = ""
    If (vgFechaIniMortalBen_M <> "") And (vgFechaFinMortalBen_M <> "") Then
        If ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) >= (vgFechaIniMortalBen_M)) And _
        ((Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00")) <= (vgFechaFinMortalBen_M)) Then
            vgBuscarMortalBen_M = "N"
        Else
            vgBuscarMortalBen_M = "S"
        End If
    Else
        vgBuscarMortalBen_M = "S"
    End If
    
    If (vgBuscarMortalBen_M = "S") Then
        'For h = 1 To 2  '1=Causante '2=Beneficiario
        h = 2
            vgs_Sexo = ""
            vgs_Tipo = ""
            vgs_Nro = ""
            'For i = 1 To 2
            '    If i = 1 Then vgs_Sexo = "M"    ' hombre  (M)asculino
            '    If i = 2 Then vgs_Sexo = "F"    ' mujer   (F)emenino
                i = 1
                vgs_Sexo = "M"
                
                'For j = 1 To 3
                j = 2
                    'If j = 1 Then vgs_Tipo = "MIT"    '1  = Invalido Total
                    'If j = 2 Then vgs_Tipo = "RV"    '2  = No Invalido
                    If h = 2 And j = 2 Then vgs_Tipo = "B" ' Tabla de beneficiarios no invalidos
                    'If j = 3 Then vgs_Tipo = "MIP"    '3  = invalido Parcial
                
                vgSw = False
                
                'Buscar Número Correlativo desde la Estructura
                For vgI = 1 To vgNumeroTotalTablas
                                      
                    'Sql = " SELECT * "
                    'Sql = Sql & " from PR_TVAL_MORTAL "
                    'Sql = Sql & " where cod_sexo = '" & vgs_Sexo & "' and "
                    'Sql = Sql & " cod_tiptabmor ='" & vgs_Tipo & "' and "
                    'Sql = Sql & " fec_ini <= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "# AND "
                    'Sql = Sql & " fec_ter >= #" & DateSerial(Format(Navig, "0000"), Format(Nmvig, "00"), Format(Ndvig, "00")) & "#"
                    'Set tb = vgConexionBD.Execute(Sql)
                    'If Not (tb.EOF) Then
                    
                    If (egTablaMortal(vgI).Sexo = vgs_Sexo) And _
                    (egTablaMortal(vgI).TipoTabla = vgs_Tipo) And _
                    (egTablaMortal(vgI).Nombre = vgPalabra_MortalBen_M) And _
                    (egTablaMortal(vgI).FechaIni <= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) And _
                    (egTablaMortal(vgI).FechaFin >= (Format(iNavig, "0000") & Format(iNmvig, "00") & Format(iNdvig, "00"))) Then
                    
                        vgs_error = 0
                        vgs_Nro = egTablaMortal(vgI).Correlativo
                        vgFechaIniMortalBen_M = egTablaMortal(vgI).FechaIni
                        vgFechaFinMortalBen_M = egTablaMortal(vgI).FechaFin
                        
                        'Limpiar columna de Datos
                        For vgX = 1 To Fintab
                            Ly(i, j, vgX) = 0
                        Next vgX
                        
                        'vgs_Nro = tb!num_correlativo
                        'vgFechaInicioMortal = Format(tb!fec_ini, "yyyymmdd")
                        'vgFechaFinMortal = Format(tb!fec_ter, "yyyymmdd")
                        ''vgFechaInicioMortal = DateSerial(Year(tb!fec_ini), Month(tb!fec_ini), Day(tb!fec_ini))
                        ''vgFechaFinMortal = DateSerial(Year(tb!fec_ter), Month(tb!fec_ter), Day(tb!fec_ter))
                        If vgs_Nro <> 0 Then
                            Sql = "Select num_edad AS edad,mto_lx from TVAL_MORDET "
                            Sql = Sql & " where num_correlativo = " & vgs_Nro
                            Sql = Sql & " order by num_edad "
                            Set Tb2 = vgConexionBD.Execute(Sql)
                            If Not (Tb2.EOF) Then
                                vgSw = True
                                'tb2.MoveFirst
                                'k = 1
                                Do While Not Tb2.EOF
                                    k = Tb2!Edad
                                    'If h = 1 Then   'Causante
                                    '    lx(i, j, k) = tb2!mto_lx
                                    'Else    'Beneficiario
                                        Ly(i, j, k) = Tb2!mto_lx
                                    'End If
                                    'k = k + 1
                                    Tb2.MoveNext
                                Loop
                            Else
                                vgError = 1068
                                Exit Function
                            End If
                            Tb2.Close
                        Else
                            vgError = 1068
                            Exit Function
                        End If
                    'Else
                    '    vgError = 1068
                    '    Exit Function
                    End If
                    'tb.Close
                Next vgI
                If (vgSw = False) Then
                    vgError = 1068
                    Exit Function
                End If
                'Next j
            'Next i
        'Next h
    End If

    fgBuscarMortalidad = True
End Function

Function fgCalcularRentaVitalicia(istPolizas As TyPoliza, istBeneficiarios() As TyBeneficiarios, Coti As String, codigo_afp As String, iRentaAFP As Double, iNumCargas As Integer) As Boolean
Dim Prodin() As Double
Dim Flupen() As Double, Flucm() As Double, Exced() As Double
Dim impres(9, 110) As Double
Dim Ncorbe(20) As Integer
Dim Penben(20) As Double, Porcbe(20) As Double, porcbe_ori(20) As Double
Dim Coinbe(20) As String, Codcbe(20) As String, Sexobe(20) As String
Dim Nanbe(20) As Integer, Nmnbe(20) As Integer, Ndnbe(20) As Integer
Dim Ijam(20) As Integer, Ijmn(20) As Integer, Ijdn(20) As Integer
Dim Npolbe(20) As String, derpen(20) As Integer
Dim i As Integer
Dim Totpor As Double
Dim cob(5) As String, alt1(3) As String, tip(2) As String

Dim Npolca As String, Mone As String
Dim Cober As String, Alt As String, Indi As String, cplan As String
Dim Nben As Long
Dim Nap As Integer, Nmp As Integer, Ndp As Integer
Dim Fechan As Long, Fechap As Long
Dim Mesdif As Long, Mesgar As Long
Dim Bono As Double, Bono_Pesos1 As Double, GtoFun As Double
Dim CtaInd As Double, SalCta As Double, Salcta_Sol As Double
Dim Ffam As Double, porfam As Double
Dim Prc_Tasa_Afp As Double, Prc_Pension_Afp As Double
Dim vgs_Coti As String

Dim edbedi As Long, mdif As Long
Dim large As Integer
Dim edaca As Long, edalca As Long, edacai As Long, edacas As Long, edabe As Long, edalbe As Long
Dim Fasolp As Long, Fmsolp As Long, Fdsolp As Long, pergar As Long, numrec As Long, numrep As Long
Dim nrel As Long, nmdif As Long, numbep As Long, ni As Long, ns As Long, nibe As Long, nsbe As Long, limite As Long
Dim limite1 As Long, limite2 As Long, limite3 As Long, limite4 As Long, imas1 As Long, kdif As Long, nt As Long
Dim nmax As Integer, j As Integer
Dim rmpol As Double, px As Double, py As Double, qx As Double, relres As Double
Dim comisi As Double, tasac As Double, timp As Double, tm As Double, tmm As Double
Dim gasemi As Double
Dim tm3 As Double, sumapx As Double, sumaqx As Double, actual As Double, actua1 As Double, PenBase As Double, tce As Double
Dim vpte As Double, difres As Double, difre1 As Double, tir As Double, tinc As Double
Dim Tasa As Double, tastce As Double, tirvta As Double, tvmax As Double
Dim vppen As Double, vpcm As Double, penanu As Double, reserva As Double, gastos As Double, rdeuda As Double
Dim resfin As Double, rend As Double, varrm As Double, resant As Double, flupag As Double, gto As Double
Dim sumaex As Double, sumaex1 As Double, tirmax As Double
Dim Sql As String, Numero As String
Dim Linea1 As String
Dim Inserta As String
Dim Var As String, Nombre As String
Dim cuenta As Integer
Dim nom_moneda As String
Dim nom_alt As String
Dim nom_plan As String
Dim nom_modalidad As String
Dim vlMargenDespuesImpuesto As Double
Dim facfam As Double
Dim fprob As Double
Dim vlI As Long
Dim tirmax_ori As Double
Dim tasac_mod As Double
Dim vp_tasac As Double
Dim vlContarMaximo As Long, vlMtoCtaIndAfp As Double

Dim vlCorrCot As Integer, vgd_tasa_vta As Double, FecDev As String, FecCot As String
Dim h As Integer, Cor As Integer
Dim intX As Integer, k As Integer, ltot As Long, comint As Double
Dim mto As Double
Dim add_porc_ben As Double, gto_sepelio As Double, mesga2 As Double
Dim vlFechaNacCausante As String
Dim vlSexoCausante As String, vlMoneda As String
Dim sumaporcsob As Double, vgPensionCot As Double, Mto_ValPrePenTmp As Double
Dim fapag As Long, fechas As Long, mesdif1 As Long, pergar1 As Long, mescon As Long
Dim fmpag As Integer
Dim icont10 As Integer, icont20 As Integer, icont11 As Integer, icont21 As Integer
Dim icont30 As Integer, icont35 As Integer, icont30Inv As Integer
Dim icont40 As Integer, icont77 As Integer
Dim vlSumPension As Double, MtoMoneda As Double
Dim DerCrecer As String
Dim DerGratificacion As String
Dim lrefun As Long
Dim facgratif() As Double
Dim fecha1 As Date

'YO
Dim LL As Integer, ij As Integer
Dim X As Long
Dim perdif As Long
Dim nmdiga As Long
Dim edhm As Long
Dim swg As String
Dim flumax  As Double
Dim pension  As Double
Dim renta  As Double
Dim ax As Double
Dim vlNumero As Integer
Dim Tasa_afp As Double
Dim Navig As Integer, Nmvig As Integer, Ndvig As Integer
Dim vlMtoPenSim As Double, vlMtoPriUniDif As Double
Dim vlRtaTmpAFP As Double, vlPriUniSim As Double
Dim vlPenGar As Double
Dim TipoCot As String

'---------------------------------------------------------------------------
'Ultima Modificación realizada el 18/12/2005
'Agregar Tablas de Mortalidad con lectura desde la BD y no desde un archivo
'Además, manejar dichas tablas por Fecha de Vigencia, para que opere la que
'corresponda a la Fecha de Cotización
'La Tabla de Mortalidad en esta función es MENSUAL
'---------------------------------------------------------------------------

    ReDim facgratif(Fintab)
    TipoCot = "C"
    
    'Lee y Calcula por Modalidad, en este caso solo se trata de una
    'se pasan los parametros a variables
    For vlNumero = 1 To 1
        cuenta = 1
        
        'vgd_tasa_vta = 0
        'Npolca = (dr("Num_Cot").ToString)
        'Fintab = (dr("FinTab").ToString)
'*-* I Agregado por ABV
        Navig = Mid(istPolizas.Fec_Vigencia, 1, 4)
        Nmvig = Mid(istPolizas.Fec_Vigencia, 5, 2)
        Ndvig = Mid(istPolizas.Fec_Vigencia, 7, 2)
'*-* F
        
        Nben = iNumCargas
        'If istPolizas.Cod_TipPension = "08" Then Nben = Nben - 1
        
        Cober = istPolizas.Cod_TipPension '(dr("Plan").ToString)
        Indi = istPolizas.Cod_TipRen 'CInt((dr("Indicador").ToString))   ' I o D
        Alt = istPolizas.Cod_Modalidad '(dr("Alternativa").ToString)
        'I - KVR 06/08/2007 -
        pergar = istPolizas.Num_MesGar 'CLng(dr("MesGar").ToString)
        'Mesgar = CLng(dr("MesGar").ToString)
        'F - KVR 06/08/2007 -
        Mone = istPolizas.Cod_Moneda '(dr("Moneda").ToString) 'vgMonedaOficial ABV 17-07-2007
        FecCot = istPolizas.Fec_Calculo '(dr("FecCot").ToString)
        Nap = CInt(Mid(FecCot, 1, 4))
        Nmp = CInt(Mid(FecCot, 5, 2))
        Ndp = CInt(Mid(FecCot, 7, 2))
        'I - KVR 06/08/2007 -
        'idp = CInt(Mid((dr("FecCot").ToString), 7, 2))
        'Bono = CDbl((dr("Mto_BonoAct").ToString))
        'Bono_Pesos1 = CDbl((dr("Mto_BonoActPesos").ToString))
        CtaInd = istPolizas.Mto_CtaIndMod 'CDbl((dr("CtaInd").ToString))   'EN SOLES
        Prima_unica = istPolizas.Mto_PriUniMod  'CDbl((dr("Mto_PriUni").ToString))
        SalCta = istPolizas.Mto_PriUniMod  'CDbl((dr("Mto_PriUni").ToString)) ' SIEMPRE VIENE EN LA MONEDA DE LA MODALIDAD
        'F - KVR 06/08/2007 -
        Ffam = istPolizas.Mto_FacPenElla  'CDbl((dr("FacPenElla").ToString))
        'I - KVR 06/08/2007 - se agrega porcentaje pensando en ella
        porfam = istPolizas.Prc_FacPenElla 'CDbl((dr("PrcFacPenElla").ToString))
        'F - KVR 06/08/2007 -
        Mesdif = istPolizas.Num_MesDif 'CLng((dr("MesDif").ToString))
        FecDev = istPolizas.Fec_Dev '(dr("FecDev").ToString)
        Fasolp = CLng(Mid(FecDev, 1, 4))   'a_sol_pen
        Fmsolp = CLng(Mid(FecDev, 5, 2))  'm_sol_pen
        Fdsolp = CLng(Mid(FecDev, 7, 2))    'd_sol_pen
        GtoFun = istPolizas.Mto_CuoMor 'CDbl((dr("Gassep").ToString))  'siempre es en soles
        MtoMoneda = istPolizas.Mto_ValMoneda 'CDbl(dr("MtoMoneda").ToString)
        If Mone <> vgMonedaCodOfi Then
            GtoFun = Format(CDbl(GtoFun / MtoMoneda), "#0.00000")
        End If
        vlCorrCot = istPolizas.Num_Correlativo 'CInt((dr("Num_Correlativo").ToString))
        Tasa = istPolizas.Prc_TasaVta 'CDbl((dr("Prc_TasaVta").ToString))
        Tasa = Format(Tasa, "#0.00")
        'I - KVR 06/08/2007 - comente estos campos ya que no aparecen en funcion de Daniela
        Prc_Tasa_Afp = istPolizas.Prc_RentaAFP / 100 'CDbl((dr("RtaAfp").ToString)) / 100
        Tasa_afp = istPolizas.Prc_RentaAFP / 100
        Prc_Pension_Afp = istPolizas.Prc_RentaTMP / 100 'CDbl((dr("RtaTmp").ToString)) / 100
        comint = istPolizas.Prc_CorCom 'CDbl((dr("Prc_ComCor").ToString))
        'F - KVR 06/08/2007 -
        'I - KVR 11/08/2007 -
        DerCrecer = istPolizas.Cod_DerCre '(dr("DerCre").ToString)  ' S/N Variable si tiene o no Derecho a Crecer la modalidad
        DerGratificacion = istPolizas.Cod_DerGra '(dr("DerGra").ToString) ' S/N
        'F - KVR 11/08/2007 -

        If TipoCot = "S" Then
            'PenBase = CDbl((dr("Pension").ToString))
        End If

        fecha1 = DateSerial(Nap, Nmp, 1)
        For i = 1 To Fintab
            facgratif(i) = 1
            If (Month(fecha1) = 7 Or Month(fecha1) = 12) And DerGratificacion = "S" Then facgratif(i) = 2
            fecha1 = DateSerial(Nap, Nmp + i, 1)
        Next i

        'La conversión de estos códigos debe ser corregida a la Oficial
        If Cober = "08" Then Cober = "S"
        If Cober = "06" Then Cober = "I"
        If Cober = "07" Then Cober = "P"
        If Cober = "04" Or Cober = "05" Then Cober = "V"
        'SalCta = Salcta_Sol
        Totpor = 0
        'I - KVR 06/08/20007 -
        If Indi = "1" Then Indi = "I"
        If Indi = "2" Then Indi = "D"

        If Alt = "1" Then Alt = "S"
        If Alt = "3" Then Alt = "G"
        If Alt = "4" Then Alt = "F"
        'F - KVR 06/08/2007 -
        'Obtiene los Datos de los Beneficiarios
'*-*            If vlCorrCot = 1 Or TipoCot = "M" Then
        If vlNumero = 1 Then
            i = 1

            For i = 1 To iNumCargas
                Ncorbe(i) = istBeneficiarios(i).Cod_Par '(dRow("Parentesco").ToString)
                Porcbe(i) = istBeneficiarios(i).Prc_Pension '(dRow("Porcentaje").ToString)
                'I - KVR 06/08/2007 -
                porcbe_ori(i) = istBeneficiarios(i).Prc_PensionLeg '(dRow("PorcentajeLeg").ToString)
                If (Ncorbe(i) = 99) Or (Ncorbe(i) = 0) Then
                    vlFechaNacCausante = istBeneficiarios(i).Fec_NacBen '(dRow("Fec.Nac.").ToString)
                    vlSexoCausante = istBeneficiarios(i).Cod_Sexo '(dRow("Sexo").ToString)
                End If
                'F - KVR 06/08/2007 -
                Dim fecha As String
                fecha = istBeneficiarios(i).Fec_NacBen '(dRow("Fec.Nac.").ToString)
                Nanbe(i) = Mid(fecha, 1, 4)  'aa_nac
                Nmnbe(i) = Mid(fecha, 5, 2) 'mm_nac
                Ndnbe(i) = Mid(fecha, 7, 2) 'mm_nac
                Sexobe(i) = istBeneficiarios(i).Cod_Sexo '(dRow("Sexo").ToString)
                Coinbe(i) = istBeneficiarios(i).Cod_SitInv '(dRow("Sit.Inv.").ToString)
                Codcbe(i) = istBeneficiarios(i).Cod_DerCre '(dRow("Dº Crecer").ToString)
                'If Len((dRow("Fec.Nac.HM").ToString)) > 0 Then
                If Len(istBeneficiarios(i).Fec_NacHM) > 0 Then
                    fecha = ""
                    fecha = istBeneficiarios(i).Fec_NacHM '(dRow("Fec.Nac.HM").ToString)
                    Ijam(i) = Mid(fecha, 1, 4)  'aa_hijom
                    Ijmn(i) = Mid(fecha, 5, 2)    'mm_hijom
                    Ijdn(i) = Mid(fecha, 7, 2)    'mm_hijom
                Else
                    Ijam(i) = "0000" ' Year(tb_difben!fec_nachm)   'aa_hijom
                    Ijmn(i) = "00" 'Month(tb_difben!fec_nachm)     'mm_hijom
                    Ijdn(i) = "00" 'Month(tb_difben!fec_nachm)     'mm_hijom
                End If
                Npolbe(i) = istPolizas.Num_Cot
                Porcbe(i) = Porcbe(i) / 100
                porcbe_ori(i) = porcbe_ori(i) / 100
                If Cober = "S" And (Ncorbe(i) <> 0 Or Ncorbe(i) <> 99) Then sumaporcsob = sumaporcsob + Porcbe(i)

                'Penben(i) = Porcbe(i)
                'derpen(i) = (dRow("Dº Pension").ToString) 'Dº Pensión
                'If derpen(i) <> 10 Then
                '    If Cober <> "S" Then
                '        If Ncorbe(i) <> 99 Then
                '            Totpor = Totpor + Porcbe(i)
                '        End If
                '    Else
                '        Totpor = Totpor + Porcbe(i)
                '    End If
                'End If
'*-*                i = i + 1
            Next i
            
'*-* I Dentro del VB ya se encuentran registradas en el L24, L21, L18
'            ' Nben = i - 1
'            'validar los topes de edad de pago de pensiones
'            Dim LimEdad As New Limite_Edad
'            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope24, FecDev)
'            If LimEdad.Mensaje = Nothing Then
'                L24 = (LimEdad.LimEdad)
'            Else
'                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
'                Exit Function
'            End If
'
'            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope21, FecDev)
'            If LimEdad.Mensaje = Nothing Then
'                L21 = (LimEdad.LimEdad)
'            Else
'                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
'                Exit Function
'            End If
'
'            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope18, FecDev)
'            If LimEdad.Mensaje = Nothing Then
'                L18 = (LimEdad.LimEdad)
'            Else
'                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
'                Exit Function
'            End If
'
'            If fgValorFactorAjusteIPC(FecDev, FecCot) = False Then
'                vgFactorAjusteIPC = 0
'            End If
'
'            'I - KVR 11/08/2007 - SOLO UNA VEZ
'            L24 = L24 * 12
'            L21 = L21 * 12
'            L18 = L18 * 12
'            'F - KVR 11/08/2007 -
'*-* F
        End If

        If Cober = "S" And sumaporcsob > 1 And porfam > 0 Then
            X = MsgBox("La suma de los porcentajes de pensión corregidos por factor Pensando en Ella es mayor al 100%.", vbCritical, "Proceso de cálculo Abortado")
            'Renta_Vitalicia = False
            Exit Function
        End If

        'ReDim Cp(Fintab)
        'ReDim Prodin(Fintab)
        ReDim Flupen(Fintab)
        'ReDim Flucm(Fintab)
        'ReDim Exced(Fintab)

'*-* I Modificación de Carga de Tablas de Mortalidad
        '-------------------------------------------------
        'Leer Tabla de Mortalidad
        '-------------------------------------------------
        If (fgBuscarMortalidadNormativa(Navig, Nmvig, Ndvig, Nap, Nmp, Ndp, vlSexoCausante, vlFechaNacCausante) = False) Then
            'Se produjeron errores en la obtención de una de las Tablas de Mortalidad
            Exit Function
        End If
        
'        'llenar las matrices Lx y Ly
'        ReDim Lx(2, 3, Fintab)
'        ReDim Ly(2, 3, Fintab)
'
'        For Each dataRow In dtMatriz.Rows
'            i = (dataRow("i").ToString)
'            j = (dataRow("j").ToString)
'            h = (dataRow("h").ToString)
'            k = (dataRow("k").ToString)
'            mto = CDbl((dataRow("mto_lx").ToString))
'
'            'If (h = 1) Then
'            '    For intX = 1 To FinTab
'            '        Lx(i, j, intX) = 0
'            '    Next intX
'            'Else
'            '    For intX = 1 To FinTab
'            '        Ly(i, j, intX) = 0
'            '    Next intX
'            'End If
'
'            If h = 1 Then   'Causante
'                Lx(i, j, k) = mto
'            Else    'Beneficiario
'                Ly(i, j, k) = mto
'            End If
'
'        Next
'*-* F

        cuenta = 0
        numrec = -1
        lrefun = 288

        'Inicializacion de variables
        Fechap = Nap * 12 + Nmp
        perdif = 0
        fapag = 0
        fmpag = 0
        fechas = 0
        
        'Recalculo de periodo garantizado y diferido despues de la fecha de devengamiento.
        mesdif1 = Mesdif 'debe venir en meses
        pergar1 = pergar
        mescon = Fechap - ((Fasolp * 12) + Fmsolp)
        If (mescon < (mesdif1 + pergar1)) Then
            If (mescon > mesdif1) Then
                mescosto = mescon - mesdif1
            Else
                mescosto = 0
            End If
        Else
            mescosto = mescon - mesdif1
        End If
        'Periodo Diferido
        If (mescon > mesdif1) Then
            Mesdif = 0
        Else
            If (mesdif1 > mescon) Then
                Mesdif = (mesdif1 - mescon)
            Else
                Mesdif = (mescon - mesdif1)
            End If
        End If
        'Periodo Garantizado
        If (mescon > (pergar1 + mesdif1)) Then
            pergar = 0
        Else
            If (mescon < mesdif1) Then
                pergar = pergar
            Else
                pergar = (pergar1 + mesdif1) - mescon
            End If
        End If
        perdif = Mesdif

        If Indi = "D" Then
            icont10 = 0: icont20 = 0: icont11 = 0: icont21 = 0
            icont30 = 0: icont35 = 0
            icont40 = 0: icont77 = 0: icont30Inv = 0
            For j = 1 To Nben
                nibe = 0
                If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
                If Coinbe(j) = "N" Then nibe = 2
                If Coinbe(j) = "P" Then nibe = 2
                If nibe = 0 Then
                    X = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
                    'Renta_Vitalicia = False
                    Exit Function
                End If
                nsbe = 0
                If Sexobe(j) = "M" Then nsbe = 1
                If Sexobe(j) = "F" Then nsbe = 2
                If nsbe = 0 Then
                    X = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
                    'Renta_Vitalicia = False
                    Exit Function
                End If
                Fechan = Nanbe(j) * 12 + Nmnbe(j)
                edabe = Fechap - Fechan
                If edabe < 1 Then edabe = 1
                If edabe > Fintab Then
                    'vgError = 1023
                    Exit Function
                End If
                If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Then icont10 = icont10 + 1
                If Ncorbe(j) = 20 Or Ncorbe(j) = 21 Then icont20 = icont20 + 1
                If Ncorbe(j) = 30 Then icont30 = icont30 + 1
                If Ncorbe(j) = 30 And Coinbe(j) <> "N" Then icont30Inv = icont30Inv + 1
                If Ncorbe(j) = 35 Then icont35 = icont35 + 1
                If Ncorbe(j) > 40 And Ncorbe(j) < 50 Then icont40 = icont40 + 1
                If Ncorbe(j) = 77 Then icont77 = icont77 + 1
            Next j
            If (icont10 > 0 Or icont20 > 0) And icont30 > 0 And icont30Inv = 0 Then
                For j = 1 To Nben
                    If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or Ncorbe(j) = 20 Or Ncorbe(j) = 21 Then
                        edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j))) + perdif
                        If edhm >= L18 Then
                            Porcbe(j) = 0.42
                        End If
                    End If
                Next j
            End If
        End If

        tmtce = (1 + Tasa / 100) ^ (1 / 12)

        ''If Indi = 2 Then
        ''    ' Mesdif = Mesdif * 12
        ''    PerDif = Mesdif
        ''End If
        ''rmpol = 0
        ''If Alt = 3 Or (Alt = 4 And pergar > 0) Then Mesgar = pergar

        'If Cober = 8 Or Cober = 9 Or Cober = 10 Or Cober = 11 Or Cober = 12 Then ABV 17-07-2007
        If Cober = "S" Then

            For ij = 1 To Fintab
                Flupen(ij) = 0
            Next ij
            Mesgar = pergar
            'I - KVR 17/08/2007 -
            If Alt = "S" Then Mesgar = 0
            'If Alt = 1 Then Mesgar = 0
            'F - KVR 17/08/2007 -
            nmdiga = perdif + Mesgar
            For j = 1 To Nben
                pension = Porcbe(j)
                swg = "N"
                'nrel = 0
                nibe = 0
                If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
                If Coinbe(j) = "N" Then nibe = 2
                If Coinbe(j) = "P" Then nibe = 2
                If nibe = 0 Then
                    X = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
                    'Renta_Vitalicia = False
                    Exit Function
                End If
                nsbe = 0
                If Sexobe(j) = "M" Then nsbe = 1
                If Sexobe(j) = "F" Then nsbe = 2
                If nsbe = 0 Then
                    X = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
                    'Renta_Vitalicia = False
                    Exit Function
                End If

                'Calculo de la edad de los beneficiarios
                Fechan = Nanbe(j) * 12 + Nmnbe(j)
                edabe = Fechap - Fechan
                If edabe > Fintab Then
                    X = MsgBox("Error edad del beneficiario es mayor a la tabla de mortalidad.", vbCritical, "Proceso de cálculo Abortado")
                    ' Renta_Vitalicia = False
                    Exit Function

                End If
                If edabe < 1 Then edabe = 1
                If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
                    Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
                    Ncorbe(j) = 41 Or Ncorbe(j) = 46 Or _
                    Ncorbe(j) = 42 Or Ncorbe(j) = 45 Or _
                    (Ncorbe(j) >= 30 And Ncorbe(j) < 40) And (Coinbe(j) <> "N" And edabe > L18) Then

                    'PRIMA SOBREVIVENCIA VITALICIA
                    pension = Porcbe(j)
                    limite1 = Fintab - edabe - 1
                    nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
                    For i = 0 To limite1
                        imas1 = i + 1
                        edalbe = edabe + i
                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
                        If i < nmdiga Then py = 1
                        If i < perdif Then py = 0
                        'Flupen(imas1) = Flupen(imas1) + py * Pension
                        Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
                    Next i
                    'DERECHO A ACRECER
                    'If Codcbe(j) <> "N" Then
                    If DerCrecer <> "N" Then
                        edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
                        If edhm > L18 Then
                            nmdif = 0
                        Else
                            nmdif = L18 - edhm
                        End If
                        Ecadif = edabe + nmdif
                        limite1 = Fintab - Ecadif - 1
                        pension = Penben(j) * 0.2
                        nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
                        For i = 0 To limite1
                            imas1 = nmdif + i + 1
                            edalbe = Ecadif + i
                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
                            If (i + nmdif) < nmdiga And nrel = 2 Then py = 1
                            If (i + nmdif) < perdif Then py = 0
                            'Flupen(imas1) = Flupen(imas1) + py * Pension
                            Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
                        Next i

                    End If

                Else
                    If Ncorbe(j) >= 30 And Ncorbe(j) < 40 And ((Coinbe(j) = "N" And edabe <= L18)) Then
                        'PRIMA DE PENSIONES TEMPORALES
                        If (edabe > L18 And Coinbe(j) = "N") Then
                            X = MsgBox("Error edad de hijo mayor a la edad legal", vbCritical)
                        Else
                            If edabe < L18 Then
                                mdif = L18 - edabe
                                nmdif = mdif - 1
                                nmax = fgMaximo(nmax, CInt(nmdif)) '*-*nmax = amax.amax0(nmax, CInt(nmdif))
                                For i = 0 To nmdif
                                    imas1 = i + 1
                                    edalbe = edabe + i
                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
                                    If swg = "S" And i < nmdiga Then py = 1
                                    'En el Siscot2 estaba esta Línea ?????
                                    If i < nmdiga Then py = 1
                                    'Fin ???
                                    If i < perdif Then py = 0
                                    'Flupen(imas1) = Flupen(imas1) + py * Pension
                                    Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
                                Next i
                            End If
                        End If
                        'PRIMA DE HIJOS INVALIDOS
                        If Coinbe(j) <> "N" Then
                            kdif = mdif
                            edbedi = edabe + kdif
                            limite3 = Fintab - edbedi - 1
                            pension = Porcbe(j)
                            nmax = fgMaximo(nmax, CInt(limite3)) '*-*nmax = amax.amax0(nmax, CInt(limite3))
                            For i = 0 To limite3
                                edalbe = edbedi + i
                                nmdifi = i + kdif
                                imas1 = nmdifi + 1
                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
                                'En el Siscot2 estaba esta Línea ?????
                                If i < nmdiga Then py = 1
                                'Fin ???
                                If nmdifi < perdif Then py = 0
                                'Flupen(imas1) = Flupen(imas1) + py * Pension
                                Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
                            Next i
                        End If
                    Else
                        'X = MsgBox("Error en códificación de parentesco.", vbCritical, "Proceso de Cálculo Abortado")
                        ' Renta_Vitalicia = False
                        'Exit Function

                    End If
                End If
            Next j

            Dim ax_sob As Double

            '********************************************************************************************************************
            'Si es simulación PARA SOBREVIVENCIA


            If TipoCot = "S" Then


                    'Pension Leida  *********************************************************** Colocar Nombre que tu tienes en esta variable
                    PensionIngresada = PenBase

                    'Prima Unica Leida  ******************************************************* Colocar Nombre que tu tienes en esta variable
                    PUIngresada = SalCta

                    'Tasa de venta leida ******************************************************* Colocar Nombre que tu tienes en esta variable
                    TasaIngresada = Tasa

                    tmtce = (1 + TasaIngresada / 100) ^ (1 / 12)

                    ax = 0
                    For LL = 1 To nmax
                        ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
                    Next LL

                    If ax <= 0 Then
                        renta = 0
                        ax = 0
                    Else
                        ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
                        ax = CDbl(Format(ax, "#,#0.000000"))
                    End If

                    If Num_Cot = "PENSION" Then 'Calculo de Pension *************************************************
                        renta = SalCta / ax
                        renta = CDbl(Format(renta, "#,#0.00"))

                        vlMtoPenSim = renta

                    End If

                    If Num_Cot = "PRIMA" Then 'Calculo de Prima Unica *************************************************

                        If ax <= 0 Then
                            renta = 0
                            ax = 0
                            Prima_unica = 0
                        Else

                            Prima_unica = ax * PensionIngresada
                        End If
                    End If
            Else


                    'nO ES SIMULACION

                    ax = 0
                    For LL = 1 To nmax
                        ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
                    Next LL
        
                    If ax <= 0 Then
                        renta = 0
                        ax = 0
                    Else
                        ax_sob = ax
                        ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
                        renta = SalCta / ax
                        '******
                        renta = CDbl(Format(renta, "#,#0.00"))
                        ax = CDbl(Format(ax, "#,#0.000000"))
                        '*******
                    End If
                    If Indi = "D" Then
                        renta = 0
                    End If
                    If Indi = "D" Then
                        vgPensionCot = renta
                    Else
                        vgPensionCot = (renta / vgFactorAjusteIPC)
                    End If
                    If (Mone = vgMonedaCodOfi) Then
                        renta = (renta / vgFactorAjusteIPC)
                    End If
        
                    vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
                    vlMtoPenSim = Format(renta, "##0.00")
                    vlPenGar = Format(0, "##0.00") 'sumaqx
                    '----------------------------------------------------------------------

                    Dim SumaPensbenef As Double, MesHijoDif_l18 As Integer
                    Dim new_periodo As Double, sumanew_periodo As Double



                    If Indi = "I" Then
                        If mescosto > 0 Then
                            sumanew_periodo = 0
                            For j = 1 To Nben
                                new_periodo = 0
                                If Ncorbe(j) = 30 Then
                                    Fechan = Nanbe(j) * 12 + Nmnbe(j)
                                    edabe = (Fasolp * 12 + Fmsolp) - Fechan
                                    If edabe > Fintab Then
                                        edabe = Fintab
                                    Else
                                        If edabe < 1 Then edabe = 1
                                    End If
                                    MesHijoDif_l18 = L18 - edabe
                                    If MesHijoDif_l18 < 0 Then MesHijoDif_l18 = 0
                                    new_periodo = (MesHijoDif_l18 * Porcbe(j)) / sumaporcsob
                                Else
                                    new_periodo = mescon * Porcbe(j) / sumaporcsob
                                End If
                                sumanew_periodo = sumanew_periodo + new_periodo

                            Next j
                            renta = (SalCta / ((ax_sob / sumaporcsob) + sumanew_periodo)) / sumaporcsob
                            renta = CDbl(Format(renta, "#,#0.00"))
                            vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
                            vlMtoPenSim = Format(renta, "##0.00")
                            vlPenGar = Format(0, "##0.00") 'sumaqx
                        Else
                            renta = (SalCta / ((ax_sob / sumaporcsob) + mescosto)) / sumaporcsob
                            renta = CDbl(Format(renta, "#,#0.00"))
                            vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
                            vlMtoPenSim = Format(renta, "##0.00")
                            vlPenGar = Format(0, "##0.00") 'sumaqx
                        End If
                    Else
                        'calcula diferida
                        add_porc_ben = 0
                        Vpptem = 0
                        Tasa_afp = 0
                        'Prima_unica = 0
                        Rete_sim = 0
                        Prun_sim = 0
                        Sald_sim = 0
                        mesga2 = 0
                        vlMoneda = ""
                        vgPensionCot = 0
        
                        vlMoneda = Mone
                        If Mesdif > 0 Then
                            gto_sepelio = vlPenGar
                            If Prc_Tasa_Afp = 0 Then
                                Vpptem = 0
                            Else
                                Vpptem = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ (mesdif1 / 12))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
                                Vpptem = CDbl(Format(CDbl(Vpptem), "##0.000000"))
                            End If
                        Else
                            Vpptem = mesdif1
                        End If

                        'If Indi = "D" Then
                        '    Dim edadHijo(20) As Integer, sumavpptem1 As Double
                        '    Dim EdadDifer(20) As Integer, vpptem1(20) As Double
                        '    For i = 1 To Nben
                        '        If Ncorbe(i) = 30 And Coinbe(i) <> "N" Then
                        '            edadHijo(i) = Fechap - (Nanbe(i) * 12 + Nmnbe(i))
                        '            If edadHijo(i) < 1 Then edadHijo(i) = 1
                        '            If edadHijo(i) > Fintab Then
                        '                Exit Function
                        '            End If
                        '            EdadDifer(i) = CInt((L18 - edadHijo(i)) / 12)

                        '            If EdadDifer(i) < mesdif1 Then
                        '                vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ EdadDifer(i))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
                        '                vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
                        '            Else
                        '                vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ (mesdif1 / 12))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
                        '                vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
                        '            End If
                        '        Else
                        '            vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ (mesdif1 / 12))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
                        '            vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
                        '        End If
                        '        sumavpptem1 = sumavpptem1 + vpptem1(i)
                        '    Next i
                        '    'vpptem = sumavpptem1 / Nben
                        'End If

                        'tasa_afp=rentabilidad de la afp
                        'Distribucion de pensiones de benefciiarios
                        If mescosto > 0 Then
                            sumanew_periodo = 0
                            For j = 1 To Nben
                                new_periodo = 0
                                If Ncorbe(j) = 30 Then
                                    Fechan = Nanbe(j) * 12 + Nmnbe(j)
                                    edabe = (Fasolp * 12 + Fmsolp) - Fechan
                                    If edabe > Fintab Then
                                        edabe = Fintab
                                    Else
                                        If edabe < 1 Then edabe = 1
                                    End If
                                    MesHijoDif_l18 = L18 - edabe
                                    If MesHijoDif_l18 < 0 Then MesHijoDif_l18 = 0
                                    new_periodo = (MesHijoDif_l18 * Porcbe(j)) / sumaporcsob
                                Else
                                    new_periodo = mescon * Porcbe(j) / sumaporcsob
                                End If
                                sumanew_periodo = sumanew_periodo + new_periodo

                            Next j
                        Else
                            sumanew_periodo = mescosto
                        End If

                        'renta = (SalCta / ((ax_sob / sumaporcsob) + sumanew_periodo)) / sumaporcsob
                        'renta = CDbl(Format(renta, "#,#0.00"))
                        'vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
                        'vlMtoPenSim = Format(renta, "##0.00")
                        'vlPenGar = Format(0, "##0.00") 'sumaqx

                        If vlPriUniSim > 0 Then
                            If (vlMoneda = vgMonedaCodOfi) Then
                                If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
                                    Rete_sim = 0
                                Else
                                    'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * (Dif!mto_priunisim * vgFactorAjusteIPC))))), "##0.00"))
                                    'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + Prc_Pension_Afp * ((vlPriUniSim / sumaporcsob) * vgFactorAjusteIPC))))), "##0.00"))
                                    'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + Prc_Pension_Afp * ((vlPriUniSim / sumaporcsob) * vgFactorAjusteIPC))))), "##0.00"))
                                    'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + Prc_Pension_Afp * ((ax_sob / sumaporcsob) + mescosto)) * vgFactorAjusteIPC))), "##0.00"))
                                    Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * ((ax_sob / sumaporcsob) + sumanew_periodo)) * vgFactorAjusteIPC))), "##0.00"))
                                End If
                            Else
                                If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
                                    Rete_sim = 0
                                Else
                                    'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
                                    'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + Prc_Pension_Afp * (vlPriUniSim / sumaporcsob))))), "##0.00"))
                                    'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + Prc_Pension_Afp * ((ax_sob / sumaporcsob) + mescosto))))), "##0.00"))
                                    Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * ((ax_sob / sumaporcsob) + sumanew_periodo))))), "##0.00"))
                                End If
                            End If
                        End If

                        Dim PensionBenef(20) As Double, sumapenben As Double
                        sumapenben = 0
                        For i = 1 To Nben
                            PensionBenef(i) = CDbl(Format(CDbl((Porcbe(i) / sumaporcsob) * Rete_sim), "##0.00"))
                            sumapenben = sumapenben + PensionBenef(i)
                        Next i
        
                        '' para sobrevivencia corresponde a la pension real / suma%benef
                        Rete_sim = CDbl(Format(CDbl(Rete_sim / sumaporcsob), "##0.00"))
        
                        'Saldo necesario AFP
                        'Sald_sim = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
                        Sald_sim = CDbl(Format(CDbl(Vpptem * sumapenben), "#,#0.00"))

                        If vlPriUniSim > 0 Then
                            'prun_sim = CDbl(Format(CDbl(Salcta_Sol - Sald_sim), "#,#0.00"))
                            Prun_sim = CDbl(Format(CDbl(Prima_unica - Sald_sim), "#,#0.00"))
        
                            'vlPensim = CDbl(Format(CDbl((prun_sim - gto_sepelio) / vlPriUniSim), "#,#0.00"))
                            'pension modificarla a pension/ suma % benef
                                    'vlPensim = CDbl(Format(CDbl((prun_sim - gto_sepelio) / ((ax_sob / sumaporcsob) + mescosto)) / sumaporcsob, "#,#0.00"))
                                    vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / ((ax_sob / sumaporcsob) + sumanew_periodo)) / sumaporcsob, "#,#0.00"))
                                    If (vlMoneda = "NS") Then
                                        'vgPensionCot = CDbl(Format(CDbl(((prun_sim - gto_sepelio) / (((ax_sob / sumaporcsob) + mescosto) * vgFactorAjusteIPC)) / sumaporcsob), "##0.00"))
                                        vgPensionCot = CDbl(Format(CDbl(((Prun_sim - gto_sepelio) / (((ax_sob / sumaporcsob) + sumanew_periodo) * vgFactorAjusteIPC)) / sumaporcsob), "##0.00"))
                                    Else
                                        'vgPensionCot = CDbl(Format(CDbl(((prun_sim - gto_sepelio) / ((ax_sob / sumaporcsob) + mescosto)) / sumaporcsob), "##0.00"))
                                        vgPensionCot = CDbl(Format(CDbl(((Prun_sim - gto_sepelio) / ((ax_sob / sumaporcsob) + sumanew_periodo)) / sumaporcsob), "##0.00"))
                                    End If
                            vlPensim = vgPensionCot
                        Else
                            vgPensionCot = 0
                        End If
                        Mto_ValPrePenTmp = Vpptem
                        vlMtoPenSim = vlPensim 'hqr para reporte
                        vlMtoPriUniDif = Prun_sim
                        vlMtoCtaIndAfp = Sald_sim
                        vlRtaTmpAFP = sumapenben        'rete_sim
                    End If
            
                    Dim vlSumaPension As Double
                    'Registrar los valores de la Pensión para cada
                    'Beneficiario para el caso de Sobrevivencia
                    vlSumaPension = 0
                    For vlI = 1 To Nben
                        If (Ncorbe(vlI) <> 0) Then
                            If (Ncorbe(vlI) = 10) And (Ffam > 1) Then
                                vlSumaPension = vlSumaPension + CDbl(Format((vgPensionCot * (porfam / 100)), "#0.00"))
                            Else
                                vlSumaPension = vlSumaPension + CDbl(Format((vgPensionCot * Porcbe(vlI)), "#0.00"))
                            End If
                        End If
                    Next vlI
                    vlSumPension = Format(vlSumaPension, "#0.00")

            End If  ' FIN DE PREGUNTA DE SI ES SIMULACION

            '********************************************************************************************************************

        Else

            'FLUJOS DE RENTAS VITALICIAS DE VEJEZ E INVALIDEZ
            qx = 0
            For ij = 1 To Fintab
                Flupen(ij) = 0
            Next ij
            facfam = Ffam

            'I - KVR 17/08/2007 -
            If Alt = "S" Then Mesgar = 0
            If Alt = "G" Or (Alt = "F" And Mesgar > 0) Then Mesgar = pergar 'Mesgar

            'If Alt = "F" Or Alt = "S" Then Mesgar = 0
            'If Alt = "G" Then Mesgar = pergar
            'If Alt = "S" Or Alt = "F" Then Mesgar = 0
            'F - 17/08/2007 -
            
            'Definicion del periodo garantizado en 0 para las
            'alternativas Simple o pensiones con distinto porcentaje legal
            For j = 1 To Nben
                'If derpen(j) <> 10 Then
                pension = Porcbe(j)
                'CALCULO DE LA PRIMA DEL AFILIADO
                If (Ncorbe(j) = 0 Or Ncorbe(j) = 99) And j = 1 Then
                    ni = 0
                    If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then ni = 1
                    If Coinbe(j) = "N" Then ni = 2
                    If Coinbe(j) = "P" Then ni = 3
                    If ni = 0 Then
                        X = MsgBox("Error de códificación de tipo de inavlidez", vbCritical, "Proceso de Cálculo Abortado")
                        'Renta_Vitalicia = False
                        Exit Function
                    End If
                    ns = 0
                    If Sexobe(j) = "M" Then ns = 1
                    If Sexobe(j) = "F" Then ns = 2
                    If ns = 0 Then
                        X = MsgBox("Error de códificación de sexo.", vbCritical, "Proceso de Cálculo Abortado")
                        'Renta_Vitalicia = False
                        Exit Function
                    End If
                    'Calculo de edad del causante
                    Fechan = Nanbe(j) * 12 + Nmnbe(j)
                    edaca = Fechap - Fechan
                    If edaca < 780 And ns = 1 And ni = 2 Then cplan = "A"
                    If edaca < 720 And ns = 2 And ni = 2 Then cplan = "A"
                    If edaca <= 216 Then edaca = 216
                    If edaca > Fintab Then
                        X = MsgBox("Error en edad del beneficiario mayor a tabla de mortalidad", vbCritical, "Proceso de cálculo Abortado")
                        'Renta_Vitalicia = False
                        Exit Function
                    End If
                    sumaqx = 0
                    limite1 = CInt(Fintab - edaca - 1)
                    nmax = CInt(limite1)
                    For i = 0 To limite1
                        imas1 = i + 1
                        edacai = edaca + i
                        If edacai > Fintab Then
                            X = MsgBox("Edad fuera de Rangos establecidos.", vbCritical, "Proceso de Cálculo Abortado")
                            'Renta_Vitalicia = False
                            Exit Function
                        End If
                        px = Lx(ns, ni, edacai) / Lx(ns, ni, edaca)
                        edacas = edacai + 1
                        qx = ((Lx(ns, ni, edacai) - Lx(ns, ni, edacas))) / Lx(ns, ni, edaca)
                        'Flupen(imas1) = Flupen(imas1) + px * Pension
                        Flupen(imas1) = Flupen(imas1) + px * pension * facgratif(imas1)
                        sumaqx = sumaqx + GtoFun * qx / tmtce ^ (i + 0.5)
                    Next i
                End If

                If Ncorbe(j) <> 0 And Ncorbe(j) <> 99 Then
                    'Prima de los Beneficiarios
                    nibe = 0
                    If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
                    If Coinbe(j) = "N" Then nibe = 2
                    If Coinbe(j) = "P" Then nibe = 3
                    If nibe = 0 Then
                        X = MsgBox("Error de códificación de tipo de invalidez.", vbCritical, "Proceso de cálculo Abortado")
                        'Renta_Vitalicia = False
                        Exit Function
                    End If
                    nsbe = 0
                    If Sexobe(j) = "M" Then nsbe = 1
                    If Sexobe(j) = "F" Then nsbe = 2
                    If nsbe = 0 Then
                        X = MsgBox("Error de códificación de sexo.", vbCritical, "Proceso de cálculo Abortado")
                        'Renta_Vitalicia = False
                        Exit Function
                    End If
                    'Calculo de la edad del beneficiario
                    edabe = Fechap - (Nanbe(j) * 12 + Nmnbe(j))
                    If edabe < 1 Then edabe = 1
                    If edabe > Fintab Then
                        X = MsgBox("Error Edad del beneficario es mayor a la tabla.", vbCritical, "Proceso de cálculo Abortado")
                        'Renta_Vitalicia = False
                        Exit Function
                    End If
                    If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
                       Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
                       Ncorbe(j) = 41 Or Ncorbe(j) = 46 Or _
                       Ncorbe(j) = 42 Or Ncorbe(j) = 45 Or _
                       ((Ncorbe(j) >= 30 And Ncorbe(j) < 40) And _
                       (Coinbe(j) <> "N" And edabe > L18)) Then
                        'FLUJOS DE VIDAS CONJUNTAS VITALICIAS
                        'Probabilidad del beneficiario solo
                        limite1 = Fintab - edabe - 1
                        nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
                        For i = 0 To limite1
                            imas1 = i + 1
                            edalbe = edabe + i
                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
                            'Flupen(imas1) = Flupen(imas1) + py * Pension * facfam
                            Flupen(imas1) = Flupen(imas1) + py * pension * facfam * facgratif(imas1)
                        Next i
                        'Probabilidad conjunta de causante y beneficiario
                        limite2 = Fintab - edaca - 1
                        limite = fgMinimo(limite1, CInt(limite2)) '*-* limite = amax.amin0(limite1, CInt(limite2))
                        nmax = fgMaximo(nmax, CInt(limite)) '*-* nmax = amax.amax0(nmax, CInt(limite))
                        For i = 0 To limite
                            imas1 = i + 1
                            edalca = edaca + i
                            edalbe = edabe + i
                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
                            px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
                            'Flupen(imas1) = Flupen(imas1) - (py * px * Pension * facfam)
                            Flupen(imas1) = Flupen(imas1) - (py * px * pension * facfam) * facgratif(imas1)
                        Next i

                        'DERECHO A ACRECER
                        'If Codcbe(j) <> "N" Then
                        If DerCrecer <> "N" Then
                            edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
                            If edhm > L18 Then
                                nmdif = 0
                            Else
                                nmdif = L18 - edhm
                            End If
                            Ecadif = edabe + nmdif
                            limite1 = Fintab - Ecadif - 1
                            pension = Porcbe(j) * 0.2
                            nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
                            For i = 0 To limite1
                                imas1 = nmdif + i + 1
                                edalbe = Ecadif + i
                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
                                'Flupen(imas1) = Flupen(imas1) + py * Pension * facfam
                                Flupen(imas1) = Flupen(imas1) + py * pension * facfam * facgratif(imas1)
                            Next i
                            
                            Ecadif = edaca + nmdif
                            limite4 = Fintab - Ecadif - 1
                            limite = fgMinimo(limite1, CInt(limite4)) '*-* limite = amax.amin0(limite1, CInt(limite4))

                            For i = 0 To limite
                                imas1 = nmdif + i + 1
                                edalbe = Ecadif + i
                                edalca = (edaca + nmdif) + i
                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
                                px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
                                'Flupen(imas1) = Flupen(imas1) - (py * px * Pension * facfam)
                                Flupen(imas1) = Flupen(imas1) - (py * px * pension * facfam) * facgratif(imas1)
                            Next i

                        End If

                    Else
                        If Ncorbe(j) >= 30 And Ncorbe(j) < 40 Then
                            'Prima Rentas Temporales
                            If edabe <= L18 Then
                                mdif = L18 - edabe
                                nmdif = mdif
                                'Probabilidad conjunta del causante y beneficiario
                                limite2 = Fintab - edaca
                                limite = fgMinimo(nmdif, CInt(limite2)) - 1 '*-* limite = amax.amin0(nmdif, CInt(limite2)) - 1
                                nmax = fgMaximo(nmax, CInt(limite)) '*-* nmax = amax.amax0(nmax, CInt(limite))
                                For i = 0 To limite
                                    imas1 = i + 1
                                    edalca = edaca + i
                                    edalbe = edabe + i
                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
                                    px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
                                    'Flupen(imas1) = Flupen(imas1) + ((py * Pension) - (py * px * Pension)) * facfam
                                    Flupen(imas1) = Flupen(imas1) + ((py * pension) - (py * px * pension)) * facfam * facgratif(imas1)
                                Next i
                            Else
                                'I---- ABV 06/01/2004 ---
                                'Verificar la acción a realizar
                                'para que continue el proceso, ya que los
                                'Hijos mayores a la Edad Legal no se deben
                                'calcular.
                                'Preguntar a Daniela la implicancia de esto
                                'x = MsgBox("La Edad del Hijo es mayor a la Edad Legal.", vbCritical, "Proceso de Cálculo Abortado")
                                'Renta_vitalicia = False
                                'Exit Function
                                'F---- ABV 06/01/2004 ---
                            End If
                            'Prima del Hijo Invalido
                            If Coinbe(j) <> "N" Then
                                'Probabilidad conjunta del causante y beneficiario
                                edbedi = edabe + nmdif
                                limite3 = Fintab - edbedi - 1
                                limite4 = Fintab - (edaca + nmdif) - 1
                                nmax = fgMaximo(nmax, CInt(limite3)) '*-* nmax = amax.amax0(nmax, CInt(limite3))
                                For i = 0 To limite3
                                    nmdifi = nmdif + i
                                    imas1 = nmdifi + 1
                                    edalca = edaca + nmdif + i
                                    edalbe = edbedi + i
                                    edalca = fgMinimo(edalca, CInt(Fintab)) '*-* edalca = amax.amin0(edalca, CInt(Fintab))
                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
                                    px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
                                    'Flupen(imas1) = Flupen(imas1) + (py - (py * px)) * Pension * facfam
                                    Flupen(imas1) = Flupen(imas1) + (py - (py * px)) * pension * facfam * facgratif(imas1)
                                Next i
                            End If
                        Else
                            X = MsgBox("Error en codificación del parentesco.", vbCritical, "Proceso de cálculo Abortado")
                            'Renta_Vitalicia = False
                            Exit Function
                        End If
                    End If
                End If
                'End If
            Next j


                '********************************************************************************************************************
                'Si es simulación PARA VEJEZ E INVALIDEZ

            If TipoCot = "S" Then


                    'Pension Leida  *********************************************************** Colocar Nombre que tu tienes en esta variable
                    PensionIngresada = PenBase

                    'Prima Unica Leida  ******************************************************* Colocar Nombre que tu tienes en esta variable
                    PUIngresada = SalCta

                    'Tasa de venta leida ******************************************************* Colocar Nombre que tu tienes en esta variable
                    TasaIngresada = Tasa

                    tmtce = (1 + TasaIngresada / 100) ^ (1 / 12)


                    ax = 0
                    flumax = 0
                    If Alt = "S" Then  'KVR 17/08/2007 SE ELIMINO DE ESTA LINEA Alt = "F"
                        Mesgar = 0
                        nmdiga = perdif + Mesgar
                        ax = 0
                        For LL = 1 To nmax
                            flumax = Flupen(LL)
                            If LL <= perdif Then flumax = 0
                            ax = ax + flumax / tmtce ^ (LL - 1)
                        Next LL
                    Else
                        Mesgar = pergar
                        nmdiga = perdif + Mesgar
                        ax = 0
                        For LL = 1 To nmdiga
                            flumax = fgMaximo(1, Flupen(LL)) '*-* flumax = amax.amax1(1, Flupen(LL))
                            If LL <= perdif Then flumax = 0
                            ax = ax + flumax / tmtce ^ (LL - 1)
                        Next LL
                        For LL = nmdiga + 1 To nmax
                            ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
                        Next LL
                    End If


                    If Num_Cot = "PENSION" Then 'Calculo de Pension *************************************************

                        If Indi = "I" Then
                            If ax <= 0 Then
                                renta = 0
                                ax = 0
                            Else
                                sumaqx = CDbl(Format(sumaqx, "#,#0.000000"))
                                ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
                                renta = CDbl(Format(((SalCta - sumaqx) / ax), "#,#0.00"))
                            End If
                        Else
                            ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
                            renta = CDbl(Format(((SalCta - sumaqx) / ax), "#,#0.00"))   '0
                        End If

                        vlMtoPenSim = renta
                        vlMtoPriUniDif = 0
                        vlMtoCtaIndAfp = 0
                        vlRtaTmpAFP = 0
                        vlPriUniSim = ax
    
                    End If

                    If Num_Cot = "PRIMA" Then 'Calculo de Prima Unica *************************************************

                        If ax <= 0 Then
                            renta = 0
                            ax = 0
                            Prima_unica = 0
                        Else
                            sumaqx = CDbl(Format(sumaqx, "#,#0.000000"))
                            ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
                            'renta = CDbl(Format(((SalCta - sumaqx) / ax), "#,#0.00"))

                            Prima_unica = (ax * PensionIngresada) + sumaqx
                        End If
                        vlMtoPenSim = renta
                        vlMtoPriUniDif = Prima_unica
                        vlMtoCtaIndAfp = 0
                        vlRtaTmpAFP = 0
                        vlPriUniSim = ax

                    End If
            Else
                    'NO ES SIMULACION

                    'Calculo de tarifa y Pension
                    ax = 0
                    flumax = 0
                    If Alt = "S" Then  'KVR 17/08/2007 SE ELIMINO DE ESTA LINEA Alt = "F"
                        Mesgar = 0
                        nmdiga = perdif + Mesgar
                        ax = 0
                        For LL = 1 To nmax
                            flumax = Flupen(LL)
                            If LL <= perdif Then flumax = 0
                            ax = ax + flumax / tmtce ^ (LL - 1)
                        Next LL
                    Else
                        Mesgar = pergar
                        nmdiga = perdif + Mesgar
                        ax = 0
                        For LL = 1 To nmdiga
                            flumax = fgMaximo(1, Flupen(LL)) 'amax.amax1(1, Flupen(LL))
                            If LL <= perdif Then flumax = 0
                            ax = ax + flumax / tmtce ^ (LL - 1)
                        Next LL
                        For LL = nmdiga + 1 To nmax
                            ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
                        Next LL
                    End If

                    If Indi = "I" Then
                        If ax <= 0 Then
                            renta = 0
                            ax = 0
                        Else
                            sumaqx = CDbl(Format(sumaqx, "#,#0.000000"))
                            ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
                            renta = CDbl(Format(((SalCta - sumaqx) / ax), "#,#0.00"))
                        End If
                    Else
                        ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
                        renta = 0
                    End If
                    If (Mone = vgMonedaCodOfi) Then
                        renta = (renta * vgFactorAjusteIPC)
                    End If
                    'HQR 24/05/2004
                    vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
                    vlMtoPenSim = Format(renta, "##0.00")
                    vlPenGar = Format(sumaqx, "##0.00")
                    'FIN HQR 24/05/2004

                    If Indi = "D" Then
                        'aca se iba a la función que calcula diferida
                        add_porc_ben = 0
                        Vpptem = 0
                        Tasa_afp = 0
                        ' Prima_unica = 0
                        Rete_sim = 0
                        Prun_sim = 0
                        Sald_sim = 0
                        mesga2 = 0
                        vlMoneda = ""
                        vgPensionCot = 0
        
                        vlMoneda = Mone
                        If Mesdif > 0 Then
                            gto_sepelio = vlPenGar
                            If Prc_Tasa_Afp = 0 Then
                                Vpptem = 0
                            Else
                                Vpptem = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ (mesdif1 / 12))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
                                Vpptem = CDbl(Format(CDbl(Vpptem), "##0.000000"))
                            End If
                        Else
                            Vpptem = mesdif1
                        End If

                        'Dim edadHijo(20) As Integer, sumavpptem1 As Double
                        'Dim EdadDifer(20) As Integer, vpptem1(20) As Double
                        'sumavpptem1 = 0
                        'For i = 1 To Nben
                        '    If Ncorbe(i) = 30 And Coinbe(i) = "N" Then
                        '        edadHijo(i) = Fechap - (Nanbe(i) * 12 + Nmnbe(i))
                        '        If edadHijo(i) < 1 Then edadHijo(i) = 1
                        '        If edadHijo(i) > Fintab Then
                        '            Exit Function
                        '        End If
                        '        EdadDifer(i) = CInt((L18 - edadHijo(i)) / 12)
                        '        If EdadDifer(i) > (mesdif1 / 12) Then
                        '            edaddifer(i) = (mesdif1 / 12)
                        '        End If

                        '        If EdadDifer(i) < mesdif1 Then
                        '            vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ EdadDifer(i))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
                        '            vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
                        '        Else
                        '            vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ (mesdif1 / 12))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
                        '            vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
                        '        End If
                        '    Else
                        '        vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ (mesdif1 / 12))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
                        '        vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
                        '    End If

                        'Next i


                        'tasa_afp=rentabilidad de la afp
        
                        If vlPriUniSim > 0 Then
                            If (vlMoneda = vgMonedaCodOfi) Then
                                If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
                                    Rete_sim = 0
                                Else
                                    'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * (Dif!mto_priunisim * vgFactorAjusteIPC))))), "##0.00"))
                                    Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * (vlPriUniSim * vgFactorAjusteIPC))))), "##0.00"))
                                End If
                            Else
                                If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
                                    Rete_sim = 0
                                Else
                                    'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
                                    Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * vlPriUniSim)))), "##0.00"))
                                End If
                            End If
                        End If


                        'Saldo necesario AFP
                        'Dim sumapensionbenef As Double
                        ''Dim vpptem1(20) As Double
                        'sumapensionbenef = 0
                        'For i = 1 To Nben
                        '    sumapensionbenef = sumapensionbenef + (rete_sim * vpptem1(i) * Porcbe(i))
                        'Next i
                        Sald_sim = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
                        'Sald_sim = CDbl(Format(CDbl(sumapensionbenef), "#,#0.00"))

                        If vlPriUniSim > 0 Then
                            Prun_sim = CDbl(Format(CDbl(Prima_unica - Sald_sim), "#,#0.00"))
                            'prun_sim = CDbl(Format(CDbl(Salcta_Sol - Sald_sim), "#,#0.00"))
                            vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "#,#0.00"))
        
                            If (vlMoneda = vgMonedaCodOfi) Then
                                vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / (vlPriUniSim * vgFactorAjusteIPC)), "##0.00"))
                            Else
                                vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "##0.00"))
                            End If
                            vlPensim = vgPensionCot
                        Else
                            vgPensionCot = 0
                        End If
                        Mto_ValPrePenTmp = Vpptem
                        vlMtoPenSim = vlPensim 'hqr para reporte
                        vlMtoPriUniDif = Prun_sim
                        vlMtoCtaIndAfp = Sald_sim
                        vlRtaTmpAFP = Rete_sim
                End If


            End If  'FIN DE PREGUNTA DE SIMULACION ="S"

                '   ********************************************************************************************************************************

        End If

        'llenar el datatable
'*-*        dtDataRow = Renta_Vitalicia.NewRow
        istPolizas.Mto_ValPrePenTmp = Mto_ValPrePenTmp 'dtDataRow("MTO_VALPREPENTMP") = mto_valprepentmp
        istPolizas.Mto_Pension = vlMtoPenSim 'dtDataRow("MTO_PENSION") = vlMtoPenSim
        istPolizas.Mto_PriUniDif = vlMtoPriUniDif 'dtDataRow("MTO_PRIUNIDIF") = vlMtoPriUniDif
        istPolizas.Mto_CtaIndAFP = vlMtoCtaIndAfp 'dtDataRow("MTO_CTAINDAFP") = vlMtoCtaIndAfp
        istPolizas.Mto_RentaTMPAFP = vlRtaTmpAFP 'dtDataRow("MTO_RENTATMPAFP") = vlRtaTmpAFP
        istPolizas.Mto_PriUniSim = vlPriUniSim 'dtDataRow("MTO_PRIUNISIM") = vlPriUniSim
        'I - KVR 13/08/2007 -
        'dtDataRow("MTO_PENSIONGAR") = vlMtoPenSim
        istPolizas.Mto_RMGtoSepRV = vlPenGar 'dtDataRow("MTO_PENSIONGAR") = vlPenGar
        istPolizas.Mto_SumPension = Format(vlSumPension, "#0.00") 'dtDataRow("MTO_SUMPENSION") = Format(vlSumPension, "#0.00")
        'F - KVR 13/08/2007 -

        'I - KVR 18/08/2007 -
        If Mone = vgMonedaCodOfi And Indi = "D" Then
            istPolizas.Mto_AjusteIPC = Format(vgFactorAjusteIPC, "#0.00000000") 'dtDataRow("MTO_AJUSTEIPC") = Format(vgFactorAjusteIPC, "#0.00000000")
        Else
            istPolizas.Mto_AjusteIPC = 1 'dtDataRow("MTO_AJUSTEIPC") = 1
        End If
        'F - KVR 18/08/2007 -

        If (Mesgar > 0) Then
            istPolizas.Mto_PensionGar = vlMtoPenSim
        Else
            istPolizas.Mto_PensionGar = 0
        End If
'*-*        dtDataRow("NUM_CORRELATIVO") = vlCorrCot
'*-*        dtDataRow("PRIMA_UNICA") = Prima_unica

        Mto_ValPrePenTmp = 0
        vlMtoPenSim = 0
        vlMtoPriUniDif = 0
        vlMtoCtaIndAfp = 0
        vlRtaTmpAFP = 0
        vlPriUniSim = 0
        vlPenGar = 0
        vlCorrCot = 0
'*-*        Renta_Vitalicia.Rows.Add (dtDataRow)

    Next vlNumero

    fgCalcularRentaVitalicia = True

End Function

Function fgCalcularRentaVitalicia_Old3(istPolizas As TyPoliza, istBeneficiarios() As TyBeneficiarios, Coti As String, codigo_afp As String, iRentaAFP As Double, iNumCargas As Integer) As Boolean
'Dim Prodin() As Double
'Dim Flupen() As Double, Flucm() As Double, Exced() As Double
'Dim impres(9, 110) As Double
'Dim Ncorbe(20) As Integer
'Dim Penben(20) As Double, Porcbe(20) As Double, porcbe_ori(20) As Double
'Dim Coinbe(20) As String, Codcbe(20) As String, Sexobe(20) As String
'Dim Nanbe(20) As Integer, Nmnbe(20) As Integer, Ndnbe(20) As Integer
'Dim Ijam(20) As Integer, Ijmn(20) As Integer, Ijdn(20) As Integer
'Dim Npolbe(20) As String, derpen(20) As Integer
'Dim i As Integer
'Dim Totpor As Double
'Dim cob(5) As String, alt1(3) As String, tip(2) As String
'
'Dim Npolca As String, Mone As String
'Dim Cober As String, Alt As String, Indi As String, cplan As String
'Dim Nben As Long
'Dim Nap As Integer, Nmp As Integer, Ndp As Integer
'Dim Fechan As Long, Fechap As Long
'Dim Mesdif As Long, Mesgar As Long
'Dim Bono As Double, Bono_Pesos1 As Double, GtoFun As Double
'Dim CtaInd As Double, SalCta As Double, Salcta_Sol As Double
'Dim Ffam As Double, porfam As Double
'Dim Prc_Tasa_Afp As Double, Prc_Pension_Afp As Double
'Dim vgs_Coti As String
'
'Dim edbedi As Long, mdif As Long
'Dim large As Integer
'Dim edaca As Long, edalca As Long, edacai As Long, edacas As Long, edabe As Long, edalbe As Long
'Dim Fasolp As Long, Fmsolp As Long, Fdsolp As Long, pergar As Long, numrec As Long, numrep As Long
'Dim nrel As Long, nmdif As Long, numbep As Long, ni As Long, ns As Long, nibe As Long, nsbe As Long, limite As Long
'Dim limite1 As Long, limite2 As Long, limite3 As Long, limite4 As Long, imas1 As Long, kdif As Long, nt As Long
'Dim nmax As Integer, j As Integer
'Dim rmpol As Double, px As Double, py As Double, qx As Double, relres As Double
'Dim comisi As Double, tasac As Double, timp As Double, tm As Double, tmm As Double
'Dim gasemi As Double
'Dim tm3 As Double, sumapx As Double, sumaqx As Double, actual As Double, actua1 As Double, PenBase As Double, tce As Double
'Dim vpte As Double, difres As Double, difre1 As Double, tir As Double, tinc As Double
'Dim Tasa As Double, tastce As Double, tirvta As Double, tvmax As Double
'Dim vppen As Double, vpcm As Double, penanu As Double, reserva As Double, gastos As Double, rdeuda As Double
'Dim resfin As Double, rend As Double, varrm As Double, resant As Double, flupag As Double, gto As Double
'Dim sumaex As Double, sumaex1 As Double, tirmax As Double
'Dim Sql As String, Numero As String
'Dim Linea1 As String
'Dim Inserta As String
'Dim Var As String, Nombre As String
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
'Dim vlContarMaximo As Long, vlMtoCtaIndAfp As Double
'
'Dim vlCorrCot As Integer, vgd_tasa_vta As Double, FecDev As String, FecCot As String
'Dim h As Integer, Cor As Integer
'Dim intX As Integer, k As Integer, ltot As Long, comint As Double
'Dim mto As Double
'Dim add_porc_ben As Double, gto_sepelio As Double, mesga2 As Double
'Dim vlFechaNacCausante As String
'Dim vlSexoCausante As String, vlMoneda As String
'Dim sumaporcsob As Double, vgPensionCot As Double, Mto_ValPrePenTmp As Double
'Dim fapag As Long, fechas As Long, mesdif1 As Long, pergar1 As Long, mescon As Long
'Dim fmpag As Integer
'Dim icont10 As Integer, icont20 As Integer, icont11 As Integer, icont21 As Integer
'Dim icont30 As Integer, icont35 As Integer, icont30Inv As Integer
'Dim icont40 As Integer, icont77 As Integer
'Dim vlSumPension As Double, MtoMoneda As Double
'Dim DerCrecer As String
'Dim DerGratificacion As String
'Dim lrefun As Long
'Dim facgratif() As Double
'Dim fecha1 As Date
'
''YO
'Dim LL As Integer, ij As Integer
'Dim x As Long
'Dim perdif As Long
'Dim nmdiga As Long
'Dim edhm As Long
'Dim swg As String
'Dim flumax  As Double
'Dim pension  As Double
'Dim renta  As Double
'Dim ax As Double
'Dim vlNumero As Integer
'Dim Tasa_afp As Double
'Dim Navig As Integer, Nmvig As Integer, Ndvig As Integer
'Dim vlMtoPenSim As Double, vlMtoPriUniDif As Double
'Dim vlRtaTmpAFP As Double, vlPriUniSim As Double
'Dim vlPenGar As Double
'
''---------------------------------------------------------------------------
''Ultima Modificación realizada el 18/12/2005
''Agregar Tablas de Mortalidad con lectura desde la BD y no desde un archivo
''Además, manejar dichas tablas por Fecha de Vigencia, para que opere la que
''corresponda a la Fecha de Cotización
''La Tabla de Mortalidad en esta función es MENSUAL
''---------------------------------------------------------------------------
'
'    ReDim facgratif(Fintab)
'
'    'Lee y Calcula por Modalidad, en este caso solo se trata de una
'    'se pasan los parametros a variables
'    For vlNumero = 1 To 1
'        cuenta = 1
'
'        'vgd_tasa_vta = 0
'        'Npolca = (dr("Num_Cot").ToString)
'        'Fintab = (dr("FinTab").ToString)
''*-* I Agregado por ABV
'        Navig = Mid(istPolizas.Fec_Vigencia, 1, 4)
'        Nmvig = Mid(istPolizas.Fec_Vigencia, 5, 2)
'        Ndvig = Mid(istPolizas.Fec_Vigencia, 7, 2)
''*-* F
'
'        Nben = iNumCargas
'        If istPolizas.Cod_TipPension = "08" Then Nben = Nben - 1
'
'        Cober = istPolizas.Cod_TipPension '(dr("Plan").ToString)
'        Indi = istPolizas.Cod_TipRen 'CInt((dr("Indicador").ToString))   ' I o D
'        Alt = istPolizas.Cod_Modalidad '(dr("Alternativa").ToString)
'        'I - KVR 06/08/2007 -
'        pergar = istPolizas.Num_MesGar 'CLng(dr("MesGar").ToString)
'        'Mesgar = CLng(dr("MesGar").ToString)
'        'F - KVR 06/08/2007 -
'        Mone = istPolizas.Cod_Moneda '(dr("Moneda").ToString) 'vgMonedaOficial ABV 17-07-2007
'        FecCot = istPolizas.Fec_Calculo '(dr("FecCot").ToString)
'        Nap = CInt(Mid(FecCot, 1, 4))
'        Nmp = CInt(Mid(FecCot, 5, 2))
'        Ndp = CInt(Mid(FecCot, 7, 2))
'        'I - KVR 06/08/2007 -
'        'idp = CInt(Mid((dr("FecCot").ToString), 7, 2))
'        'Bono = CDbl((dr("Mto_BonoAct").ToString))
'        'Bono_Pesos1 = CDbl((dr("Mto_BonoActPesos").ToString))
'        CtaInd = istPolizas.Mto_CtaInd 'CDbl((dr("CtaInd").ToString))   'EN SOLES
'        Prima_unica = istPolizas.Mto_PriUniMod  'CDbl((dr("Mto_PriUni").ToString))
'        SalCta = istPolizas.Mto_PriUniMod  'CDbl((dr("Mto_PriUni").ToString)) ' SIEMPRE VIENE EN LA MONEDA DE LA MODALIDAD
'        'F - KVR 06/08/2007 -
'        Ffam = istPolizas.Mto_FacPenElla  'CDbl((dr("FacPenElla").ToString))
'        'I - KVR 06/08/2007 - se agrega porcentaje pensando en ella
'        porfam = istPolizas.Prc_FacPenElla 'CDbl((dr("PrcFacPenElla").ToString))
'        'F - KVR 06/08/2007 -
'        Mesdif = istPolizas.Num_MesDif 'CLng((dr("MesDif").ToString))
'        FecDev = istPolizas.Fec_Dev '(dr("FecDev").ToString)
'        Fasolp = CLng(Mid(FecDev, 1, 4))   'a_sol_pen
'        Fmsolp = CLng(Mid(FecDev, 5, 2))  'm_sol_pen
'        Fdsolp = CLng(Mid(FecDev, 7, 2))    'd_sol_pen
'        GtoFun = istPolizas.Mto_CuoMor 'CDbl((dr("Gassep").ToString))  'siempre es en soles
'        MtoMoneda = istPolizas.Mto_ValMoneda 'CDbl(dr("MtoMoneda").ToString)
'        If Mone <> vgMonedaCodOfi Then
'            GtoFun = Format(CDbl(GtoFun / MtoMoneda), "#0.00000")
'        End If
'        vlCorrCot = istPolizas.Num_Correlativo 'CInt((dr("Num_Correlativo").ToString))
'        Tasa = istPolizas.Prc_TasaVta 'CDbl((dr("Prc_TasaVta").ToString))
'        Tasa = Format(Tasa, "#0.00")
'        'I - KVR 06/08/2007 - comente estos campos ya que no aparecen en funcion de Daniela
'        Prc_Tasa_Afp = istPolizas.Prc_RentaAFP / 100 'CDbl((dr("RtaAfp").ToString)) / 100
'        Tasa_afp = istPolizas.Prc_RentaAFP / 100
'        Prc_Pension_Afp = istPolizas.Prc_RentaTMP / 100 'CDbl((dr("RtaTmp").ToString)) / 100
'        comint = istPolizas.Prc_CorCom 'CDbl((dr("Prc_ComCor").ToString))
'        'F - KVR 06/08/2007 -
'        'I - KVR 11/08/2007 -
'        DerCrecer = istPolizas.Cod_DerCre '(dr("DerCre").ToString)  ' S/N Variable si tiene o no Derecho a Crecer la modalidad
'        DerGratificacion = istPolizas.Cod_DerGra '(dr("DerGra").ToString) ' S/N
'        'F - KVR 11/08/2007 -
'
'        fecha1 = DateSerial(Nap, Nmp, 1)
'        For i = 1 To Fintab
'            facgratif(i) = 1
'            If (Month(fecha1) = 7 Or Month(fecha1) = 12) And DerGratificacion = "S" Then facgratif(i) = 2
'            fecha1 = DateSerial(Nap, Nmp + i, 1)
'        Next i
'
'        'La conversión de estos códigos debe ser corregida a la Oficial
'        If Cober = "08" Then Cober = "S"
'        If Cober = "06" Then Cober = "I"
'        If Cober = "07" Then Cober = "P"
'        If Cober = "04" Or Cober = "05" Then Cober = "V"
'        'SalCta = Salcta_Sol
'        Totpor = 0
'        'I - KVR 06/08/20007 -
'        If Indi = "1" Then Indi = "I"
'        If Indi = "2" Then Indi = "D"
'
'        If Alt = "1" Then Alt = "S"
'        If Alt = "3" Then Alt = "G"
'        If Alt = "4" Then Alt = "F"
'        'F - KVR 06/08/2007 -
'        'Obtiene los Datos de los Beneficiarios
''*-*            If vlCorrCot = 1 Or TipoCot = "M" Then
'        If vlNumero = 1 Then
'            i = 1
'
'            For i = 1 To iNumCargas
'                Ncorbe(i) = istBeneficiarios(i).Cod_Par '(dRow("Parentesco").ToString)
'                Porcbe(i) = istBeneficiarios(i).Prc_Pension '(dRow("Porcentaje").ToString)
'                'I - KVR 06/08/2007 -
'                porcbe_ori(i) = istBeneficiarios(i).Prc_PensionLeg '(dRow("PorcentajeLeg").ToString)
'                If (Ncorbe(i) = 99) Or (Ncorbe(i) = 0) Then
'                    vlFechaNacCausante = istBeneficiarios(i).Fec_NacBen '(dRow("Fec.Nac.").ToString)
'                    vlSexoCausante = istBeneficiarios(i).Cod_Sexo '(dRow("Sexo").ToString)
'                End If
'                'F - KVR 06/08/2007 -
'                Dim fecha As String
'                fecha = istBeneficiarios(i).Fec_NacBen '(dRow("Fec.Nac.").ToString)
'                Nanbe(i) = Mid(fecha, 1, 4)  'aa_nac
'                Nmnbe(i) = Mid(fecha, 5, 2) 'mm_nac
'                Ndnbe(i) = Mid(fecha, 7, 2) 'mm_nac
'                Sexobe(i) = istBeneficiarios(i).Cod_Sexo '(dRow("Sexo").ToString)
'                Coinbe(i) = istBeneficiarios(i).Cod_SitInv '(dRow("Sit.Inv.").ToString)
'                Codcbe(i) = istBeneficiarios(i).Cod_DerCre '(dRow("Dº Crecer").ToString)
'                'If Len((dRow("Fec.Nac.HM").ToString)) > 0 Then
'                If Len(istBeneficiarios(i).Fec_NacHM) > 0 Then
'                    fecha = ""
'                    fecha = istBeneficiarios(i).Fec_NacHM '(dRow("Fec.Nac.HM").ToString)
'                    Ijam(i) = Mid(fecha, 1, 4)  'aa_hijom
'                    Ijmn(i) = Mid(fecha, 5, 2)    'mm_hijom
'                    Ijdn(i) = Mid(fecha, 7, 2)    'mm_hijom
'                Else
'                    Ijam(i) = "0000" ' Year(tb_difben!fec_nachm)   'aa_hijom
'                    Ijmn(i) = "00" 'Month(tb_difben!fec_nachm)     'mm_hijom
'                    Ijdn(i) = "00" 'Month(tb_difben!fec_nachm)     'mm_hijom
'                End If
'                Npolbe(i) = istPolizas.Num_Cot
'                Porcbe(i) = Porcbe(i) / 100
'                porcbe_ori(i) = porcbe_ori(i) / 100
'                If Cober = "S" And (Ncorbe(i) <> 0 Or Ncorbe(i) <> 99) Then sumaporcsob = sumaporcsob + Porcbe(i)
'
'                'Penben(i) = Porcbe(i)
'                'derpen(i) = (dRow("Dº Pension").ToString) 'Dº Pensión
'                'If derpen(i) <> 10 Then
'                '    If Cober <> "S" Then
'                '        If Ncorbe(i) <> 99 Then
'                '            Totpor = Totpor + Porcbe(i)
'                '        End If
'                '    Else
'                '        Totpor = Totpor + Porcbe(i)
'                '    End If
'                'End If
''*-*                i = i + 1
'            Next i
'
''*-* I Dentro del VB ya se encuentran registradas en el L24, L21, L18
''            ' Nben = i - 1
''            'validar los topes de edad de pago de pensiones
''            Dim LimEdad As New Limite_Edad
''            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope24, FecDev)
''            If LimEdad.Mensaje = Nothing Then
''                L24 = (LimEdad.LimEdad)
''            Else
''                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
''                Exit Function
''            End If
''
''            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope21, FecDev)
''            If LimEdad.Mensaje = Nothing Then
''                L21 = (LimEdad.LimEdad)
''            Else
''                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
''                Exit Function
''            End If
''
''            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope18, FecDev)
''            If LimEdad.Mensaje = Nothing Then
''                L18 = (LimEdad.LimEdad)
''            Else
''                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
''                Exit Function
''            End If
''
''            If fgValorFactorAjusteIPC(FecDev, FecCot) = False Then
''                vgFactorAjusteIPC = 0
''            End If
''
''            'I - KVR 11/08/2007 - SOLO UNA VEZ
''            L24 = L24 * 12
''            L21 = L21 * 12
''            L18 = L18 * 12
''            'F - KVR 11/08/2007 -
''*-* F
'        End If
'
'        If Cober = "S" And sumaporcsob > 1 And porfam > 0 Then
'            x = MsgBox("La suma de los porcentajes de pensión corregidos por factor Pensando en Ella es mayor al 100%.", vbCritical, "Proceso de cálculo Abortado")
'            'Renta_Vitalicia = False
'            Exit Function
'        End If
'
'        'ReDim Cp(Fintab)
'        'ReDim Prodin(Fintab)
'        ReDim Flupen(Fintab)
'        'ReDim Flucm(Fintab)
'        'ReDim Exced(Fintab)
'
''*-* I Modificación de Carga de Tablas de Mortalidad
'        '-------------------------------------------------
'        'Leer Tabla de Mortalidad
'        '-------------------------------------------------
'        If (fgBuscarMortalidadNormativa(Navig, Nmvig, Ndvig, Nap, Nmp, Ndp, vlSexoCausante, vlFechaNacCausante) = False) Then
'            'Se produjeron errores en la obtención de una de las Tablas de Mortalidad
'            Exit Function
'        End If
'
''        'llenar las matrices Lx y Ly
''        ReDim Lx(2, 3, Fintab)
''        ReDim Ly(2, 3, Fintab)
''
''        For Each dataRow In dtMatriz.Rows
''            i = (dataRow("i").ToString)
''            j = (dataRow("j").ToString)
''            h = (dataRow("h").ToString)
''            k = (dataRow("k").ToString)
''            mto = CDbl((dataRow("mto_lx").ToString))
''
''            'If (h = 1) Then
''            '    For intX = 1 To FinTab
''            '        Lx(i, j, intX) = 0
''            '    Next intX
''            'Else
''            '    For intX = 1 To FinTab
''            '        Ly(i, j, intX) = 0
''            '    Next intX
''            'End If
''
''            If h = 1 Then   'Causante
''                Lx(i, j, k) = mto
''            Else    'Beneficiario
''                Ly(i, j, k) = mto
''            End If
''
''        Next
''*-* F
'
'        cuenta = 0
'        numrec = -1
'        lrefun = 288
'
'        'Inicializacion de variables
'        Fechap = Nap * 12 + Nmp
'        perdif = 0
'        fapag = 0
'        fmpag = 0
'        fechas = 0
'
'        'Recalculo de periodo garantizado y diferido despues de la fecha de devengamiento.
'        mesdif1 = Mesdif 'debe venir en meses
'        pergar1 = pergar
'        mescon = Fechap - ((Fasolp * 12) + Fmsolp)
'        If (mescon < (mesdif1 + pergar1)) Then
'            If (mescon > mesdif1) Then
'                mescosto = mescon - mesdif1
'            Else
'                mescosto = 0
'            End If
'        Else
'            mescosto = mescon - mesdif1
'        End If
'        If (mescon > mesdif1) Then
'            Mesdif = 0
'        Else
'            If (mesdif1 > mescon) Then
'                Mesdif = (mesdif1 - mescon)
'            Else
'                Mesdif = (mescon - mesdif1)
'            End If
'        End If
'        If (mescon > (pergar1 + mesdif1)) Then
'            pergar = 0
'        Else
'            If (mescon < mesdif1) Then
'                pergar = pergar
'            Else
'                pergar = (pergar1 + mesdif1) - mescon
'            End If
'        End If
'        perdif = Mesdif
'
'        If Indi = "D" Then
'            icont10 = 0: icont20 = 0: icont11 = 0: icont21 = 0
'            icont30 = 0: icont35 = 0
'            icont40 = 0: icont77 = 0: icont30Inv = 0
'            For j = 1 To Nben
'                nibe = 0
'                If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
'                If Coinbe(j) = "N" Then nibe = 2
'                If Coinbe(j) = "P" Then nibe = 2
'                If nibe = 0 Then
'                    x = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'                nsbe = 0
'                If Sexobe(j) = "M" Then nsbe = 1
'                If Sexobe(j) = "F" Then nsbe = 2
'                If nsbe = 0 Then
'                    x = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'                Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                edabe = Fechap - Fechan
'                If edabe < 1 Then edabe = 1
'                If edabe > Fintab Then
'                    'vgError = 1023
'                    Exit Function
'                End If
'                If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Then icont10 = icont10 + 1
'                If Ncorbe(j) = 20 Or Ncorbe(j) = 21 Then icont20 = icont20 + 1
'                If Ncorbe(j) = 30 Then icont30 = icont30 + 1
'                If Ncorbe(j) = 30 And Coinbe(j) <> "N" Then icont30Inv = icont30Inv + 1
'                If Ncorbe(j) = 35 Then icont35 = icont35 + 1
'                If Ncorbe(j) > 40 And Ncorbe(j) < 50 Then icont40 = icont40 + 1
'                If Ncorbe(j) = 77 Then icont77 = icont77 + 1
'            Next j
'            If (icont10 > 0 Or icont20 > 0) And icont30 > 0 And icont30Inv = 0 Then
'                For j = 1 To Nben
'                    If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or Ncorbe(j) = 20 Or Ncorbe(j) = 21 Then
'                        edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j))) + perdif
'                        If edhm >= L18 Then
'                            Porcbe(j) = 0.42
'                        End If
'                    End If
'                Next j
'            End If
'        End If
'
'        tmtce = (1 + Tasa / 100) ^ (1 / 12)
'
'        ''If Indi = 2 Then
'        ''    ' Mesdif = Mesdif * 12
'        ''    PerDif = Mesdif
'        ''End If
'        ''rmpol = 0
'        ''If Alt = 3 Or (Alt = 4 And pergar > 0) Then Mesgar = pergar
'
'        'If Cober = 8 Or Cober = 9 Or Cober = 10 Or Cober = 11 Or Cober = 12 Then ABV 17-07-2007
'        If Cober = "S" Then
'
'            For ij = 1 To Fintab
'                Flupen(ij) = 0
'            Next ij
'            Mesgar = pergar
'            'I - KVR 17/08/2007 -
'            If Alt = "S" Then Mesgar = 0
'            'If Alt = 1 Then Mesgar = 0
'            'F - KVR 17/08/2007 -
'            nmdiga = perdif + Mesgar
'            For j = 1 To Nben
'                pension = Porcbe(j)
'                swg = "N"
'                'nrel = 0
'                nibe = 0
'                If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
'                If Coinbe(j) = "N" Then nibe = 2
'                If Coinbe(j) = "P" Then nibe = 2
'                If nibe = 0 Then
'                    x = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'                nsbe = 0
'                If Sexobe(j) = "M" Then nsbe = 1
'                If Sexobe(j) = "F" Then nsbe = 2
'                If nsbe = 0 Then
'                    x = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'
'                'Calculo de la edad de los beneficiarios
'                Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                edabe = Fechap - Fechan
'                If edabe > Fintab Then
'                    x = MsgBox("Error edad del beneficiario es mayor a la tabla de mortalidad.", vbCritical, "Proceso de cálculo Abortado")
'                    ' Renta_Vitalicia = False
'                    Exit Function
'
'                End If
'                If edabe < 1 Then edabe = 1
'                If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
'                    Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
'                    Ncorbe(j) = 41 Or Ncorbe(j) = 46 Or _
'                    Ncorbe(j) = 42 Or Ncorbe(j) = 45 Or _
'                    (Ncorbe(j) >= 30 And Ncorbe(j) < 40) And (Coinbe(j) <> "N" And edabe > L18) Then
'
'                    'PRIMA SOBREVIVENCIA VITALICIA
'                    pension = Porcbe(j)
'                    limite1 = Fintab - edabe - 1
'                    nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                    For i = 0 To limite1
'                        imas1 = i + 1
'                        edalbe = edabe + i
'                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                        If i < nmdiga Then py = 1
'                        If i < perdif Then py = 0
'                        'Flupen(imas1) = Flupen(imas1) + py * Pension
'                        Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
'                    Next i
'                    'DERECHO A ACRECER
'                    'If Codcbe(j) <> "N" Then
'                    If DerCrecer <> "N" Then
'                        edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
'                        If edhm > L18 Then
'                            nmdif = 0
'                        Else
'                            nmdif = L18 - edhm
'                        End If
'                        Ecadif = edabe + nmdif
'                        limite1 = Fintab - Ecadif - 1
'                        pension = Penben(j) * 0.2
'                        nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                        For i = 0 To limite1
'                            imas1 = nmdif + i + 1
'                            edalbe = Ecadif + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            If (i + nmdif) < nmdiga And nrel = 2 Then py = 1
'                            If (i + nmdif) < perdif Then py = 0
'                            'Flupen(imas1) = Flupen(imas1) + py * Pension
'                            Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
'                        Next i
'
'                    End If
'
'                Else
'                    If Ncorbe(j) >= 30 And Ncorbe(j) < 40 And ((Coinbe(j) = "N" And edabe <= L18)) Then
'                        'PRIMA DE PENSIONES TEMPORALES
'                        If (edabe > L18 And Coinbe(j) = "N") Then
'                            x = MsgBox("Error edad de hijo mayor a la edad legal", vbCritical)
'                        Else
'                            If edabe < L18 Then
'                                mdif = L18 - edabe
'                                nmdif = mdif - 1
'                                nmax = fgMaximo(nmax, CInt(nmdif)) '*-*nmax = amax.amax0(nmax, CInt(nmdif))
'                                For i = 0 To nmdif
'                                    imas1 = i + 1
'                                    edalbe = edabe + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    If swg = "S" And i < nmdiga Then py = 1
'                                    'En el Siscot2 estaba esta Línea ?????
'                                    If i < nmdiga Then py = 1
'                                    'Fin ???
'                                    If i < perdif Then py = 0
'                                    'Flupen(imas1) = Flupen(imas1) + py * Pension
'                                    Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
'                                Next i
'                            End If
'                        End If
'                        'PRIMA DE HIJOS INVALIDOS
'                        If Coinbe(j) <> "N" Then
'                            kdif = mdif
'                            edbedi = edabe + kdif
'                            limite3 = Fintab - edbedi - 1
'                            pension = Porcbe(j)
'                            nmax = fgMaximo(nmax, CInt(limite3)) '*-*nmax = amax.amax0(nmax, CInt(limite3))
'                            For i = 0 To limite3
'                                edalbe = edbedi + i
'                                nmdifi = i + kdif
'                                imas1 = nmdifi + 1
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                'En el Siscot2 estaba esta Línea ?????
'                                If i < nmdiga Then py = 1
'                                'Fin ???
'                                If nmdifi < perdif Then py = 0
'                                'Flupen(imas1) = Flupen(imas1) + py * Pension
'                                Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
'                            Next i
'                        End If
'                    Else
'                        'X = MsgBox("Error en códificación de parentesco.", vbCritical, "Proceso de Cálculo Abortado")
'                        ' Renta_Vitalicia = False
'                        'Exit Function
'
'                    End If
'                End If
'            Next j
'
'            Dim ax_sob As Double
'
'            ax = 0
'            For LL = 1 To nmax
'                ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
'            Next LL
'
'            If ax <= 0 Then
'                renta = 0
'                ax = 0
'            Else
'                ax_sob = ax
'                ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
'                renta = SalCta / ax
'                '******
'                renta = CDbl(Format(renta, "#,#0.00"))
'                ax = CDbl(Format(ax, "#,#0.000000"))
'                '*******
'            End If
'            If Indi = "D" Then
'                renta = 0
'            End If
'            If Indi = "D" Then
'                vgPensionCot = renta
'            Else
'                vgPensionCot = (renta / vgFactorAjusteIPC)
'            End If
'            If (Mone = vgMonedaCodOfi) Then
'                renta = (renta / vgFactorAjusteIPC)
'            End If
'
'            vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
'            vlMtoPenSim = Format(renta, "##0.00")
'            vlPenGar = Format(0, "##0.00") 'sumaqx
'            '----------------------------------------------------------------------
'            If Indi = "I" Then
'                'calcula inmediata
'                renta = (SalCta / ((ax_sob / sumaporcsob) + mescosto)) / sumaporcsob
'                renta = CDbl(Format(renta, "#,#0.00"))
'                vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
'                vlMtoPenSim = Format(renta, "##0.00")
'                vlPenGar = Format(0, "##0.00") 'sumaqx
'
'            Else
'                'calcula diferida
'                add_porc_ben = 0
'                Vpptem = 0
'                Tasa_afp = 0
'                'Prima_unica = 0
'                Rete_sim = 0
'                Prun_sim = 0
'                Sald_sim = 0
'                mesga2 = 0
'                vlMoneda = ""
'                vgPensionCot = 0
'
'                vlMoneda = Mone
'                If Mesdif > 0 Then
'                    gto_sepelio = vlPenGar
'                    If Prc_Tasa_Afp = 0 Then
'                        Vpptem = 0
'                    Else
'                        Vpptem = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ Mesdif)) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                        Vpptem = CDbl(Format(CDbl(Vpptem), "##0.000000"))
'                    End If
'                Else
'                    Vpptem = mesdif1
'                End If
'
'                If Indi = "D" Then
'                    Dim edadHijo(20) As Integer, sumavpptem1 As Double
'                    Dim EdadDifer(20) As Integer, vpptem1(20) As Double
'                    For i = 1 To Nben
'                        If Ncorbe(i) = 30 And Coinbe(i) <> "N" Then
'                            edadHijo(i) = Fechap - (Nanbe(i) * 12 + Nmnbe(i))
'                            If edadHijo(i) < 1 Then edadHijo(i) = 1
'                            If edadHijo(i) > Fintab Then
'                                Exit Function
'                            End If
'                            EdadDifer(i) = CInt((L18 - edadHijo(i)) / 12)
'
'                            If EdadDifer(i) < mesdif1 Then
'                                vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ EdadDifer(i))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                                vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
'                            Else
'                                vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ (mesdif1 / 12))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                                vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
'                            End If
'                        Else
'                            vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ (mesdif1 / 12))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                            vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
'                        End If
'                        sumavpptem1 = sumavpptem1 + vpptem1(i)
'                    Next i
'                    'vpptem = sumavpptem1 / Nben
'                End If
'
'                'tasa_afp=rentabilidad de la afp
'
'                If vlPriUniSim > 0 Then
'                    If (vlMoneda = vgMonedaCodOfi) Then
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * (Dif!mto_priunisim * vgFactorAjusteIPC))))), "##0.00"))
'                            'Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * (vlPriUniSim * vgFactorAjusteIPC))))), "##0.00"))
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + Prc_Pension_Afp * ((vlPriUniSim / sumaporcsob) * vgFactorAjusteIPC))))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * ((ax_sob / sumaporcsob) + mescosto)) * vgFactorAjusteIPC))), "##0.00"))
'                        End If
'                    Else
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
'                            'Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * vlPriUniSim)))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * ((ax_sob / sumaporcsob) + mescosto))))), "##0.00"))
'                        End If
'                    End If
'                End If
'
'                Dim PensionBenef(20) As Double, sumapenben As Double
'                sumapenben = 0
'                For i = 1 To Nben
'                    PensionBenef(i) = CDbl(Format(CDbl((Porcbe(i) / sumaporcsob) * Rete_sim), "##0.00"))
'                    sumapenben = sumapenben + PensionBenef(i)
'                Next i
'
'                '' para sobrevivencia corresponde a la pension real / suma%benef
'                Rete_sim = CDbl(Format(CDbl(Rete_sim / sumaporcsob), "##0.00"))
'
'                'Saldo necesario AFP
'                'Sald_sim = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
'                Sald_sim = CDbl(Format(CDbl(Vpptem * sumapenben), "#,#0.00"))
'
'                If vlPriUniSim > 0 Then
'                    'prun_sim = CDbl(Format(CDbl(Salcta_Sol - Sald_sim), "#,#0.00"))
'                    Prun_sim = CDbl(Format(CDbl(Prima_unica - Sald_sim), "#,#0.00"))
'
'                    'vlPensim = CDbl(Format(CDbl((prun_sim - gto_sepelio) / vlPriUniSim), "#,#0.00"))
'                    'pension modificarla a pension/ suma % benef
'                    vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / ((ax_sob / sumaporcsob) + mescosto)) / sumaporcsob, "#,#0.00"))
'                    If (vlMoneda = vgMonedaCodOfi) Then
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / (vlPriUniSim * vgFactorAjusteIPC)), "##0.00"))
'                    Else
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "##0.00"))
'                    End If
'                    vlPensim = vgPensionCot
'                Else
'                    vgPensionCot = 0
'                End If
'                Mto_ValPrePenTmp = Vpptem
'                vlMtoPenSim = vlPensim 'hqr para reporte
'                vlMtoPriUniDif = Prun_sim
'                vlMtoCtaIndAfp = Sald_sim
'                vlRtaTmpAFP = sumapenben        'rete_sim
'            End If
'
'            Dim vlSumaPension As Double
'            'Registrar los valores de la Pensión para cada
'            'Beneficiario para el caso de Sobrevivencia
'            vlSumaPension = 0
'            For vlI = 1 To Nben
'                If (Ncorbe(vlI) <> 0) Then
'                    If (Ncorbe(vlI) = 10) And (Ffam > 1) Then
'                        vlSumaPension = vlSumaPension + CDbl(Format((vgPensionCot * (porfam / 100)), "#0.00"))
'                    Else
'                        vlSumaPension = vlSumaPension + CDbl(Format((vgPensionCot * Porcbe(vlI)), "#0.00"))
'                    End If
'                End If
'            Next vlI
'            vlSumPension = Format(vlSumaPension, "#0.00")
'
'        Else
'
'            'FLUJOS DE RENTAS VITALICIAS DE VEJEZ E INVALIDEZ
'            qx = 0
'            For ij = 1 To Fintab
'                Flupen(ij) = 0
'            Next ij
'            facfam = Ffam
'
'            'I - KVR 17/08/2007 -
'            If Alt = "S" Then Mesgar = 0
'            If Alt = "G" Or (Alt = "F" And Mesgar > 0) Then Mesgar = pergar 'Mesgar
'
'            'If Alt = "F" Or Alt = "S" Then Mesgar = 0
'            'If Alt = "G" Then Mesgar = pergar
'            'If Alt = "S" Or Alt = "F" Then Mesgar = 0
'            'F - 17/08/2007 -
'
'            'Definicion del periodo garantizado en 0 para las
'            'alternativas Simple o pensiones con distinto porcentaje legal
'            For j = 1 To Nben
'                'If derpen(j) <> 10 Then
'                pension = Porcbe(j)
'                'CALCULO DE LA PRIMA DEL AFILIADO
'                If (Ncorbe(j) = 0 Or Ncorbe(j) = 99) And j = 1 Then
'                    ni = 0
'                    If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then ni = 1
'                    If Coinbe(j) = "N" Then ni = 2
'                    If Coinbe(j) = "P" Then ni = 3
'                    If ni = 0 Then
'                        x = MsgBox("Error de códificación de tipo de inavlidez", vbCritical, "Proceso de Cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    ns = 0
'                    If Sexobe(j) = "M" Then ns = 1
'                    If Sexobe(j) = "F" Then ns = 2
'                    If ns = 0 Then
'                        x = MsgBox("Error de códificación de sexo.", vbCritical, "Proceso de Cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    'Calculo de edad del causante
'                    Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                    edaca = Fechap - Fechan
'                    If edaca < 780 And ns = 1 And ni = 2 Then cplan = "A"
'                    If edaca < 720 And ns = 2 And ni = 2 Then cplan = "A"
'                    If edaca <= 216 Then edaca = 216
'                    If edaca > Fintab Then
'                        x = MsgBox("Error en edad del beneficiario mayor a tabla de mortalidad", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    sumaqx = 0
'                    limite1 = CInt(Fintab - edaca - 1)
'                    nmax = CInt(limite1)
'                    For i = 0 To limite1
'                        imas1 = i + 1
'                        edacai = edaca + i
'                        If edacai > Fintab Then
'                            x = MsgBox("Edad fuera de Rangos establecidos.", vbCritical, "Proceso de Cálculo Abortado")
'                            'Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                        px = Lx(ns, ni, edacai) / Lx(ns, ni, edaca)
'                        edacas = edacai + 1
'                        qx = ((Lx(ns, ni, edacai) - Lx(ns, ni, edacas))) / Lx(ns, ni, edaca)
'                        'Flupen(imas1) = Flupen(imas1) + px * Pension
'                        Flupen(imas1) = Flupen(imas1) + px * pension * facgratif(imas1)
'                        sumaqx = sumaqx + GtoFun * qx / tmtce ^ (i + 0.5)
'                    Next i
'                End If
'
'                If Ncorbe(j) <> 0 And Ncorbe(j) <> 99 Then
'                    'Prima de los Beneficiarios
'                    nibe = 0
'                    If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
'                    If Coinbe(j) = "N" Then nibe = 2
'                    If Coinbe(j) = "P" Then nibe = 3
'                    If nibe = 0 Then
'                        x = MsgBox("Error de códificación de tipo de invalidez.", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    nsbe = 0
'                    If Sexobe(j) = "M" Then nsbe = 1
'                    If Sexobe(j) = "F" Then nsbe = 2
'                    If nsbe = 0 Then
'                        x = MsgBox("Error de códificación de sexo.", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    'Calculo de la edad del beneficiario
'                    edabe = Fechap - (Nanbe(j) * 12 + Nmnbe(j))
'                    If edabe < 1 Then edabe = 1
'                    If edabe > Fintab Then
'                        x = MsgBox("Error Edad del beneficario es mayor a la tabla.", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
'                       Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
'                       Ncorbe(j) = 41 Or Ncorbe(j) = 46 Or _
'                       Ncorbe(j) = 42 Or Ncorbe(j) = 45 Or _
'                       ((Ncorbe(j) >= 30 And Ncorbe(j) < 40) And _
'                       (Coinbe(j) <> "N" And edabe > L18)) Then
'                        'FLUJOS DE VIDAS CONJUNTAS VITALICIAS
'                        'Probabilidad del beneficiario solo
'                        limite1 = Fintab - edabe - 1
'                        nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                        For i = 0 To limite1
'                            imas1 = i + 1
'                            edalbe = edabe + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            'Flupen(imas1) = Flupen(imas1) + py * Pension * facfam
'                            Flupen(imas1) = Flupen(imas1) + py * pension * facfam * facgratif(imas1)
'                        Next i
'                        'Probabilidad conjunta de causante y beneficiario
'                        limite2 = Fintab - edaca - 1
'                        limite = fgMinimo(limite1, CInt(limite2)) '*-* limite = amax.amin0(limite1, CInt(limite2))
'                        nmax = fgMaximo(nmax, CInt(limite)) '*-* nmax = amax.amax0(nmax, CInt(limite))
'                        For i = 0 To limite
'                            imas1 = i + 1
'                            edalca = edaca + i
'                            edalbe = edabe + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                            'Flupen(imas1) = Flupen(imas1) - (py * px * Pension * facfam)
'                            Flupen(imas1) = Flupen(imas1) - (py * px * pension * facfam) * facgratif(imas1)
'                        Next i
'
'                        'DERECHO A ACRECER
'                        'If Codcbe(j) <> "N" Then
'                        If DerCrecer <> "N" Then
'                            edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
'                            If edhm > L18 Then
'                                nmdif = 0
'                            Else
'                                nmdif = L18 - edhm
'                            End If
'                            Ecadif = edabe + nmdif
'                            limite1 = Fintab - Ecadif - 1
'                            pension = Porcbe(j) * 0.2
'                            nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                            For i = 0 To limite1
'                                imas1 = nmdif + i + 1
'                                edalbe = Ecadif + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                'Flupen(imas1) = Flupen(imas1) + py * Pension * facfam
'                                Flupen(imas1) = Flupen(imas1) + py * pension * facfam * facgratif(imas1)
'                            Next i
'
'                            Ecadif = edaca + nmdif
'                            limite4 = Fintab - Ecadif - 1
'                            limite = fgMinimo(limite1, CInt(limite4)) '*-* limite = amax.amin0(limite1, CInt(limite4))
'
'                            For i = 0 To limite
'                                imas1 = nmdif + i + 1
'                                edalbe = Ecadif + i
'                                edalca = (edaca + nmdif) + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                'Flupen(imas1) = Flupen(imas1) - (py * px * Pension * facfam)
'                                Flupen(imas1) = Flupen(imas1) - (py * px * pension * facfam) * facgratif(imas1)
'                            Next i
'
'                        End If
'
'                    Else
'                        If Ncorbe(j) >= 30 And Ncorbe(j) < 40 Then
'                            'Prima Rentas Temporales
'                            If edabe <= L18 Then
'                                mdif = L18 - edabe
'                                nmdif = mdif
'                                'Probabilidad conjunta del causante y beneficiario
'                                limite2 = Fintab - edaca
'                                limite = fgMinimo(nmdif, CInt(limite2)) - 1 '*-* limite = amax.amin0(nmdif, CInt(limite2)) - 1
'                                nmax = fgMaximo(nmax, CInt(limite)) '*-* nmax = amax.amax0(nmax, CInt(limite))
'                                For i = 0 To limite
'                                    imas1 = i + 1
'                                    edalca = edaca + i
'                                    edalbe = edabe + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                    'Flupen(imas1) = Flupen(imas1) + ((py * Pension) - (py * px * Pension)) * facfam
'                                    Flupen(imas1) = Flupen(imas1) + ((py * pension) - (py * px * pension)) * facfam * facgratif(imas1)
'                                Next i
'                            Else
'                                'I---- ABV 06/01/2004 ---
'                                'Verificar la acción a realizar
'                                'para que continue el proceso, ya que los
'                                'Hijos mayores a la Edad Legal no se deben
'                                'calcular.
'                                'Preguntar a Daniela la implicancia de esto
'                                'x = MsgBox("La Edad del Hijo es mayor a la Edad Legal.", vbCritical, "Proceso de Cálculo Abortado")
'                                'Renta_vitalicia = False
'                                'Exit Function
'                                'F---- ABV 06/01/2004 ---
'                            End If
'                            'Prima del Hijo Invalido
'                            If Coinbe(j) <> "N" Then
'                                'Probabilidad conjunta del causante y beneficiario
'                                edbedi = edabe + nmdif
'                                limite3 = Fintab - edbedi - 1
'                                limite4 = Fintab - (edaca + nmdif) - 1
'                                nmax = fgMaximo(nmax, CInt(limite3)) '*-* nmax = amax.amax0(nmax, CInt(limite3))
'                                For i = 0 To limite3
'                                    nmdifi = nmdif + i
'                                    imas1 = nmdifi + 1
'                                    edalca = edaca + nmdif + i
'                                    edalbe = edbedi + i
'                                    edalca = fgMinimo(edalca, CInt(Fintab)) '*-* edalca = amax.amin0(edalca, CInt(Fintab))
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                    'Flupen(imas1) = Flupen(imas1) + (py - (py * px)) * Pension * facfam
'                                    Flupen(imas1) = Flupen(imas1) + (py - (py * px)) * pension * facfam * facgratif(imas1)
'                                Next i
'                            End If
'                        Else
'                            x = MsgBox("Error en codificación del parentesco.", vbCritical, "Proceso de cálculo Abortado")
'                            'Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                    End If
'                End If
'                'End If
'            Next j
'
'            'Calculo de tarifa y Pension
'            ax = 0
'            flumax = 0
'                If Alt = "S" Then  'KVR 17/08/2007 SE ELIMINO DE ESTA LINEA Alt = "F"
'                Mesgar = 0
'                nmdiga = perdif + Mesgar
'                ax = 0
'                For LL = 1 To nmax
'                    flumax = Flupen(LL)
'                    If LL <= perdif Then flumax = 0
'                    ax = ax + flumax / tmtce ^ (LL - 1)
'                Next LL
'            Else
'                Mesgar = pergar
'                nmdiga = perdif + Mesgar
'                ax = 0
'                For LL = 1 To nmdiga
'                    flumax = fgMaximo(1, Flupen(LL)) '*-* flumax = amax.amax1(1, Flupen(LL))
'                    If LL <= perdif Then flumax = 0
'                    ax = ax + flumax / tmtce ^ (LL - 1)
'                Next LL
'                For LL = nmdiga + 1 To nmax
'                    ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
'                Next LL
'            End If
'
'            If Indi = "I" Then
'                If ax <= 0 Then
'                    renta = 0
'                    ax = 0
'                Else
'                    sumaqx = CDbl(Format(sumaqx, "#,#0.000000"))
'                    ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
'                    renta = CDbl(Format(((SalCta - sumaqx) / ax), "#,#0.00"))
'                End If
'            Else
'                ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
'                renta = 0
'            End If
'            If (Mone = vgMonedaCodOfi) Then
'                renta = (renta * vgFactorAjusteIPC)
'            End If
'            'HQR 24/05/2004
'            vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
'            vlMtoPenSim = Format(renta, "##0.00")
'            vlPenGar = Format(sumaqx, "##0.00")
'            'FIN HQR 24/05/2004
'
'            If Indi = "D" Then
'                'aca se iba a la función que calcula diferida
'                add_porc_ben = 0
'                Vpptem = 0
'                Tasa_afp = 0
'                ' Prima_unica = 0
'                Rete_sim = 0
'                Prun_sim = 0
'                Sald_sim = 0
'                mesga2 = 0
'                vlMoneda = ""
'                vgPensionCot = 0
'
'                vlMoneda = Mone
'                If Mesdif > 0 Then
'                    gto_sepelio = vlPenGar
'                    If Prc_Tasa_Afp = 0 Then
'                        Vpptem = 0
'                    Else
'                        Vpptem = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ Mesdif)) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                        Vpptem = CDbl(Format(CDbl(Vpptem), "##0.000000"))
'                    End If
'                Else
'                    Vpptem = mesdif1
'                End If
'
'                        'Dim edadHijo(20) As Integer, sumavpptem1 As Double
'                        'Dim EdadDifer(20) As Integer, vpptem1(20) As Double
'                        'sumavpptem1 = 0
'                        'For i = 1 To Nben
'                        '    If Ncorbe(i) = 30 And Coinbe(i) = "N" Then
'                        '        edadHijo(i) = Fechap - (Nanbe(i) * 12 + Nmnbe(i))
'                        '        If edadHijo(i) < 1 Then edadHijo(i) = 1
'                        '        If edadHijo(i) > Fintab Then
'                        '            Exit Function
'                        '        End If
'                        '        EdadDifer(i) = CInt((L18 - edadHijo(i)) / 12)
'                        '        If EdadDifer(i) > (mesdif1 / 12) Then
'                        '            edaddifer(i) = (mesdif1 / 12)
'                        '        End If
'
'                        '        If EdadDifer(i) < mesdif1 Then
'                        '            vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ EdadDifer(i))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                        '            vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
'                        '        Else
'                        '            vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ (mesdif1 / 12))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                        '            vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
'                        '        End If
'                        '    Else
'                        '        vpptem1(i) = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ (mesdif1 / 12))) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                        '        vpptem1(i) = CDbl(Format(CDbl(vpptem1(i)), "##0.000000"))
'                        '    End If
'
'                        'Next i
'
'
'                'tasa_afp=rentabilidad de la afp
'
'                If vlPriUniSim > 0 Then
'                    If (vlMoneda = vgMonedaCodOfi) Then
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * (Dif!mto_priunisim * vgFactorAjusteIPC))))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * (vlPriUniSim * vgFactorAjusteIPC))))), "##0.00"))
'                        End If
'                    Else
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * vlPriUniSim)))), "##0.00"))
'                        End If
'                    End If
'                End If
'
'                'Saldo necesario AFP
'                'Dim sumapensionbenef As Double
'                ''Dim vpptem1(20) As Double
'                'sumapensionbenef = 0
'                'For i = 1 To Nben
'                '    sumapensionbenef = sumapensionbenef + (rete_sim * vpptem1(i) * Porcbe(i))
'                'Next i
'                Sald_sim = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
'                'Sald_sim = CDbl(Format(CDbl(sumapensionbenef), "#,#0.00"))
'
'                If vlPriUniSim > 0 Then
'                    Prun_sim = CDbl(Format(CDbl(Prima_unica - Sald_sim), "#,#0.00"))
'                    'prun_sim = CDbl(Format(CDbl(Salcta_Sol - Sald_sim), "#,#0.00"))
'                    vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "#,#0.00"))
'
'                    If (vlMoneda = vgMonedaCodOfi) Then
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / (vlPriUniSim * vgFactorAjusteIPC)), "##0.00"))
'                    Else
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "##0.00"))
'                    End If
'                    vlPensim = vgPensionCot
'                Else
'                    vgPensionCot = 0
'                End If
'                Mto_ValPrePenTmp = Vpptem
'                vlMtoPenSim = vlPensim 'hqr para reporte
'                vlMtoPriUniDif = Prun_sim
'                vlMtoCtaIndAfp = Sald_sim
'                vlRtaTmpAFP = Rete_sim
'            End If
'        End If
'
'        'llenar el datatable
''*-*        dtDataRow = Renta_Vitalicia.NewRow
'        istPolizas.Mto_ValPrePenTmp = Mto_ValPrePenTmp 'dtDataRow("MTO_VALPREPENTMP") = mto_valprepentmp
'        istPolizas.Mto_Pension = vlMtoPenSim 'dtDataRow("MTO_PENSION") = vlMtoPenSim
'        istPolizas.Mto_PriUniDif = vlMtoPriUniDif 'dtDataRow("MTO_PRIUNIDIF") = vlMtoPriUniDif
'        istPolizas.Mto_CtaIndAFP = vlMtoCtaIndAfp 'dtDataRow("MTO_CTAINDAFP") = vlMtoCtaIndAfp
'        istPolizas.Mto_RentaTMPAFP = vlRtaTmpAFP 'dtDataRow("MTO_RENTATMPAFP") = vlRtaTmpAFP
'        istPolizas.Mto_PriUniSim = vlPriUniSim 'dtDataRow("MTO_PRIUNISIM") = vlPriUniSim
'        'I - KVR 13/08/2007 -
'        'dtDataRow("MTO_PENSIONGAR") = vlMtoPenSim
'        istPolizas.Mto_RMGtoSepRV = vlPenGar 'dtDataRow("MTO_PENSIONGAR") = vlPenGar
'        istPolizas.Mto_SumPension = Format(vlSumPension, "#0.00") 'dtDataRow("MTO_SUMPENSION") = Format(vlSumPension, "#0.00")
'        'F - KVR 13/08/2007 -
'
'        'I - KVR 18/08/2007 -
'        If Mone = vgMonedaCodOfi And Indi = "D" Then
'            istPolizas.Mto_AjusteIPC = Format(vgFactorAjusteIPC, "#0.00000000") 'dtDataRow("MTO_AJUSTEIPC") = Format(vgFactorAjusteIPC, "#0.00000000")
'        Else
'            istPolizas.Mto_AjusteIPC = 1 'dtDataRow("MTO_AJUSTEIPC") = 1
'        End If
'        'F - KVR 18/08/2007 -
'
'        If (Mesgar > 0) Then
'            istPolizas.Mto_PensionGar = vlMtoPenSim
'        Else
'            istPolizas.Mto_PensionGar = 0
'        End If
''*-*        dtDataRow("NUM_CORRELATIVO") = vlCorrCot
'
'        Mto_ValPrePenTmp = 0
'        vlMtoPenSim = 0
'        vlMtoPriUniDif = 0
'        vlMtoCtaIndAfp = 0
'        vlRtaTmpAFP = 0
'        vlPriUniSim = 0
'        vlPenGar = 0
'        vlCorrCot = 0
''*-*        Renta_Vitalicia.Rows.Add (dtDataRow)
'
'    Next vlNumero
'
'    fgCalcularRentaVitalicia = True

End Function

Function fgCalcularRentaVitalicia_Old2(istPolizas As TyPoliza, istBeneficiarios() As TyBeneficiarios, Coti As String, codigo_afp As String, iRentaAFP As Double, iNumCargas As Integer) As Boolean
'Dim Prodin() As Double
'Dim Flupen() As Double, Flucm() As Double, Exced() As Double
'Dim impres(9, 110) As Double
'Dim Ncorbe(20) As Integer
'Dim Penben(20) As Double, Porcbe(20) As Double, porcbe_ori(20) As Double
'Dim Coinbe(20) As String, Codcbe(20) As String, Sexobe(20) As String
'Dim Nanbe(20) As Integer, Nmnbe(20) As Integer, Ndnbe(20) As Integer
'Dim Ijam(20) As Integer, Ijmn(20) As Integer, Ijdn(20) As Integer
'Dim Npolbe(20) As String, derpen(20) As Integer
'Dim i As Integer
'Dim Totpor As Double
'Dim cob(5) As String, alt1(3) As String, tip(2) As String
'
'Dim Npolca As String, Mone As String
'Dim Cober As String, Alt As String, Indi As String, cplan As String
'Dim Nben As Long
'Dim Nap As Integer, Nmp As Integer, Ndp As Integer
'Dim Fechan As Long, Fechap As Long
'Dim Mesdif As Long, Mesgar As Long
'Dim Bono As Double, Bono_Pesos1 As Double, GtoFun As Double
'Dim CtaInd As Double, SalCta As Double, Salcta_Sol As Double
'Dim Ffam As Double, porfam As Double
'Dim Prc_Tasa_Afp As Double, Prc_Pension_Afp As Double
'Dim vgs_Coti As String
'
'Dim edbedi As Long, mdif As Long
'Dim large As Integer
'Dim edaca As Long, edalca As Long, edacai As Long, edacas As Long, edabe As Long, edalbe As Long
'Dim Fasolp As Long, Fmsolp As Long, Fdsolp As Long, pergar As Long, numrec As Long, numrep As Long
'Dim nrel As Long, nmdif As Long, numbep As Long, ni As Long, ns As Long, nibe As Long, nsbe As Long, limite As Long
'Dim limite1 As Long, limite2 As Long, limite3 As Long, limite4 As Long, imas1 As Long, kdif As Long, nt As Long
'Dim nmax As Integer, j As Integer
'Dim rmpol As Double, px As Double, py As Double, qx As Double, relres As Double
'Dim comisi As Double, tasac As Double, timp As Double, tm As Double, tmm As Double
'Dim gasemi As Double
'Dim tm3 As Double, sumapx As Double, sumaqx As Double, actual As Double, actua1 As Double, PenBase As Double, tce As Double
'Dim vpte As Double, difres As Double, difre1 As Double, tir As Double, tinc As Double
'Dim Tasa As Double, tastce As Double, tirvta As Double, tvmax As Double
'Dim vppen As Double, vpcm As Double, penanu As Double, reserva As Double, gastos As Double, rdeuda As Double
'Dim resfin As Double, rend As Double, varrm As Double, resant As Double, flupag As Double, gto As Double
'Dim sumaex As Double, sumaex1 As Double, tirmax As Double
'Dim Sql As String, Numero As String
'Dim Linea1 As String
'Dim Inserta As String
'Dim Var As String, Nombre As String
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
'Dim vlContarMaximo As Long, vlMtoCtaIndAfp As Double
'
'Dim vlCorrCot As Integer, vgd_tasa_vta As Double, FecDev As String, FecCot As String
'Dim h As Integer, Cor As Integer
'Dim intX As Integer, k As Integer, ltot As Long, comint As Double
'Dim mto As Double
'Dim add_porc_ben As Double, gto_sepelio As Double, mesga2 As Double
'Dim vlFechaNacCausante As String
'Dim vlSexoCausante As String, vlMoneda As String
'Dim sumaporcsob As Double, vgPensionCot As Double, Mto_ValPrePenTmp As Double
'Dim fapag As Long, fechas As Long, mesdif1 As Long, pergar1 As Long, mescon As Long
'Dim fmpag As Integer
'Dim icont10 As Integer, icont20 As Integer, icont11 As Integer, icont21 As Integer
'Dim icont30 As Integer, icont35 As Integer, icont30Inv As Integer
'Dim icont40 As Integer, icont77 As Integer
'Dim vlSumPension As Double, MtoMoneda As Double
'Dim DerCrecer As String
'Dim DerGratificacion As String
'Dim lrefun As Long
'Dim facgratif() As Double
'Dim fecha1 As Date
'
''YO
'Dim LL As Integer, ij As Integer
'Dim x As Long
'Dim perdif As Long
'Dim nmdiga As Long
'Dim edhm As Long
'Dim swg As String
'Dim flumax  As Double
'Dim pension  As Double
'Dim renta  As Double
'Dim ax As Double
'Dim vlNumero As Integer
'Dim Tasa_afp As Double
'Dim Navig As Integer, Nmvig As Integer, Ndvig As Integer
'
''---------------------------------------------------------------------------
''Ultima Modificación realizada el 18/12/2005
''Agregar Tablas de Mortalidad con lectura desde la BD y no desde un archivo
''Además, manejar dichas tablas por Fecha de Vigencia, para que opere la que
''corresponda a la Fecha de Cotización
''La Tabla de Mortalidad en esta función es MENSUAL
''---------------------------------------------------------------------------
'
'    ReDim facgratif(Fintab)
'
'    'Lee y Calcula por Modalidad, en este caso solo se trata de una
'    'se pasan los parametros a variables
'    For vlNumero = 1 To 1
'        cuenta = 1
'
'        'vgd_tasa_vta = 0
'        'Npolca = (dr("Num_Cot").ToString)
'        'Fintab = (dr("FinTab").ToString)
''*-* I Agregado por ABV
'        Navig = Mid(istPolizas.Fec_Vigencia, 1, 4)
'        Nmvig = Mid(istPolizas.Fec_Vigencia, 5, 2)
'        Ndvig = Mid(istPolizas.Fec_Vigencia, 7, 2)
''*-* F
'
'        Nben = iNumCargas
'        If istPolizas.Cod_TipPension = "08" Then Nben = Nben - 1
'
'        Cober = istPolizas.Cod_TipPension '(dr("Plan").ToString)
'        Indi = istPolizas.Cod_TipRen 'CInt((dr("Indicador").ToString))   ' I o D
'        Alt = istPolizas.Cod_Modalidad '(dr("Alternativa").ToString)
'        'I - KVR 06/08/2007 -
'        pergar = istPolizas.Num_MesGar 'CLng(dr("MesGar").ToString)
'        'Mesgar = CLng(dr("MesGar").ToString)
'        'F - KVR 06/08/2007 -
'        Mone = istPolizas.Cod_Moneda '(dr("Moneda").ToString) 'vgMonedaOficial ABV 17-07-2007
'        FecCot = istPolizas.Fec_Calculo '(dr("FecCot").ToString)
'        Nap = CInt(Mid(FecCot, 1, 4))
'        Nmp = CInt(Mid(FecCot, 5, 2))
'        Ndp = CInt(Mid(FecCot, 7, 2))
'        'I - KVR 06/08/2007 -
'        'idp = CInt(Mid((dr("FecCot").ToString), 7, 2))
'        'Bono = CDbl((dr("Mto_BonoAct").ToString))
'        'Bono_Pesos1 = CDbl((dr("Mto_BonoActPesos").ToString))
'        CtaInd = istPolizas.Mto_CtaInd 'CDbl((dr("CtaInd").ToString))   'EN SOLES
'        Prima_unica = istPolizas.Mto_PriUniMod  'CDbl((dr("Mto_PriUni").ToString))
'        SalCta = istPolizas.Mto_PriUniMod  'CDbl((dr("Mto_PriUni").ToString)) ' SIEMPRE VIENE EN LA MONEDA DE LA MODALIDAD
'        'F - KVR 06/08/2007 -
'        Ffam = istPolizas.Mto_FacPenElla  'CDbl((dr("FacPenElla").ToString))
'        'I - KVR 06/08/2007 - se agrega porcentaje pensando en ella
'        porfam = istPolizas.Prc_FacPenElla 'CDbl((dr("PrcFacPenElla").ToString))
'        'F - KVR 06/08/2007 -
'        Mesdif = istPolizas.Num_MesDif 'CLng((dr("MesDif").ToString))
'        FecDev = istPolizas.Fec_Dev '(dr("FecDev").ToString)
'        Fasolp = CLng(Mid(FecDev, 1, 4))   'a_sol_pen
'        Fmsolp = CLng(Mid(FecDev, 5, 2))  'm_sol_pen
'        Fdsolp = CLng(Mid(FecDev, 7, 2))    'd_sol_pen
'        GtoFun = istPolizas.Mto_CuoMor 'CDbl((dr("Gassep").ToString))  'siempre es en soles
'        MtoMoneda = istPolizas.Mto_ValMoneda 'CDbl(dr("MtoMoneda").ToString)
'        If Mone <> vgMonedaCodOfi Then
'            GtoFun = Format(CDbl(GtoFun / MtoMoneda), "#0.00000")
'        End If
'        vlCorrCot = istPolizas.Num_Correlativo 'CInt((dr("Num_Correlativo").ToString))
'        Tasa = istPolizas.Prc_TasaVta 'CDbl((dr("Prc_TasaVta").ToString))
'        Tasa = Format(Tasa, "#0.00")
'        'I - KVR 06/08/2007 - comente estos campos ya que no aparecen en funcion de Daniela
'        Prc_Tasa_Afp = istPolizas.Prc_RentaAFP / 100 'CDbl((dr("RtaAfp").ToString)) / 100
'        Tasa_afp = istPolizas.Prc_RentaAFP / 100
'        Prc_Pension_Afp = istPolizas.Prc_RentaTMP / 100 'CDbl((dr("RtaTmp").ToString)) / 100
'        comint = istPolizas.Prc_CorCom 'CDbl((dr("Prc_ComCor").ToString))
'        'F - KVR 06/08/2007 -
'        'I - KVR 11/08/2007 -
'        DerCrecer = istPolizas.Cod_DerCre '(dr("DerCre").ToString)  ' S/N Variable si tiene o no Derecho a Crecer la modalidad
'        DerGratificacion = istPolizas.Cod_DerGra '(dr("DerGra").ToString) ' S/N
'        'F - KVR 11/08/2007 -
'
'        fecha1 = DateSerial(Nap, Nmp, 1)
'        For i = 1 To Fintab
'            facgratif(i) = 1
'            If (Month(fecha1) = 7 Or Month(fecha1) = 12) And DerGratificacion = "S" Then facgratif(i) = 2
'            fecha1 = DateSerial(Nap, Nmp + i, 1)
'        Next i
'
'        'La conversión de estos códigos debe ser corregida a la Oficial
'        If Cober = "08" Then Cober = "S"
'        If Cober = "06" Then Cober = "I"
'        If Cober = "07" Then Cober = "P"
'        If Cober = "04" Or Cober = "05" Then Cober = "V"
'        'SalCta = Salcta_Sol
'        Totpor = 0
'        'I - KVR 06/08/20007 -
'        If Indi = "1" Then Indi = "I"
'        If Indi = "2" Then Indi = "D"
'
'        If Alt = "1" Then Alt = "S"
'        If Alt = "3" Then Alt = "G"
'        If Alt = "4" Then Alt = "F"
'        'F - KVR 06/08/2007 -
'        'Obtiene los Datos de los Beneficiarios
''*-*            If vlCorrCot = 1 Or TipoCot = "M" Then
'        If vlNumero = 1 Then
'            i = 1
'
'            For i = 1 To iNumCargas
'                Ncorbe(i) = istBeneficiarios(i).Cod_Par '(dRow("Parentesco").ToString)
'                Porcbe(i) = istBeneficiarios(i).Prc_Pension '(dRow("Porcentaje").ToString)
'                'I - KVR 06/08/2007 -
'                porcbe_ori(i) = istBeneficiarios(i).Prc_PensionLeg '(dRow("PorcentajeLeg").ToString)
'                If (Ncorbe(i) = 99) Or (Ncorbe(i) = 0) Then
'                    vlFechaNacCausante = istBeneficiarios(i).Fec_NacBen '(dRow("Fec.Nac.").ToString)
'                    vlSexoCausante = istBeneficiarios(i).Cod_Sexo '(dRow("Sexo").ToString)
'                End If
'                'F - KVR 06/08/2007 -
'                Dim fecha As String
'                fecha = istBeneficiarios(i).Fec_NacBen '(dRow("Fec.Nac.").ToString)
'                Nanbe(i) = Mid(fecha, 1, 4)  'aa_nac
'                Nmnbe(i) = Mid(fecha, 5, 2) 'mm_nac
'                Ndnbe(i) = Mid(fecha, 7, 2) 'mm_nac
'                Sexobe(i) = istBeneficiarios(i).Cod_Sexo '(dRow("Sexo").ToString)
'                Coinbe(i) = istBeneficiarios(i).Cod_SitInv '(dRow("Sit.Inv.").ToString)
'                Codcbe(i) = istBeneficiarios(i).Cod_DerCre '(dRow("Dº Crecer").ToString)
'                'If Len((dRow("Fec.Nac.HM").ToString)) > 0 Then
'                If Len(istBeneficiarios(i).Fec_NacHM) > 0 Then
'                    fecha = ""
'                    fecha = istBeneficiarios(i).Fec_NacHM '(dRow("Fec.Nac.HM").ToString)
'                    Ijam(i) = Mid(fecha, 1, 4)  'aa_hijom
'                    Ijmn(i) = Mid(fecha, 5, 2)    'mm_hijom
'                    Ijdn(i) = Mid(fecha, 7, 2)    'mm_hijom
'                Else
'                    Ijam(i) = "0000" ' Year(tb_difben!fec_nachm)   'aa_hijom
'                    Ijmn(i) = "00" 'Month(tb_difben!fec_nachm)     'mm_hijom
'                    Ijdn(i) = "00" 'Month(tb_difben!fec_nachm)     'mm_hijom
'                End If
'                Npolbe(i) = istPolizas.Num_Cot
'                Porcbe(i) = Porcbe(i) / 100
'                porcbe_ori(i) = porcbe_ori(i) / 100
'                If Cober = "S" And (Ncorbe(i) <> 0 Or Ncorbe(i) <> 99) Then sumaporcsob = sumaporcsob + Porcbe(i)
'
'                'Penben(i) = Porcbe(i)
'                'derpen(i) = (dRow("Dº Pension").ToString) 'Dº Pensión
'                'If derpen(i) <> 10 Then
'                '    If Cober <> "S" Then
'                '        If Ncorbe(i) <> 99 Then
'                '            Totpor = Totpor + Porcbe(i)
'                '        End If
'                '    Else
'                '        Totpor = Totpor + Porcbe(i)
'                '    End If
'                'End If
''*-*                i = i + 1
'            Next i
'
''*-* I Dentro del VB ya se encuentran registradas en el L24, L21, L18
''            ' Nben = i - 1
''            'validar los topes de edad de pago de pensiones
''            Dim LimEdad As New Limite_Edad
''            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope24, FecDev)
''            If LimEdad.Mensaje = Nothing Then
''                L24 = (LimEdad.LimEdad)
''            Else
''                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
''                Exit Function
''            End If
''
''            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope21, FecDev)
''            If LimEdad.Mensaje = Nothing Then
''                L21 = (LimEdad.LimEdad)
''            Else
''                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
''                Exit Function
''            End If
''
''            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope18, FecDev)
''            If LimEdad.Mensaje = Nothing Then
''                L18 = (LimEdad.LimEdad)
''            Else
''                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
''                Exit Function
''            End If
''
''            If fgValorFactorAjusteIPC(FecDev, FecCot) = False Then
''                vgFactorAjusteIPC = 0
''            End If
''
''            'I - KVR 11/08/2007 - SOLO UNA VEZ
''            L24 = L24 * 12
''            L21 = L21 * 12
''            L18 = L18 * 12
''            'F - KVR 11/08/2007 -
''*-* F
'        End If
'
'        If Cober = "S" And sumaporcsob > 1 And porfam > 0 Then
'            x = MsgBox("La suma de los porcentajes de pensión corregidos por factor Pensando en Ella es mayor al 100%.", vbCritical, "Proceso de cálculo Abortado")
'            'Renta_Vitalicia = False
'            Exit Function
'        End If
'
'        'ReDim Cp(Fintab)
'        'ReDim Prodin(Fintab)
'        ReDim Flupen(Fintab)
'        'ReDim Flucm(Fintab)
'        'ReDim Exced(Fintab)
'
''*-* I Modificación de Carga de Tablas de Mortalidad
'        '-------------------------------------------------
'        'Leer Tabla de Mortalidad
'        '-------------------------------------------------
'        If (fgBuscarMortalidadNormativa(Navig, Nmvig, Ndvig, Nap, Nmp, Ndp, vlSexoCausante, vlFechaNacCausante) = False) Then
'            'Se produjeron errores en la obtención de una de las Tablas de Mortalidad
'            Exit Function
'        End If
'
''        'llenar las matrices Lx y Ly
''        ReDim Lx(2, 3, Fintab)
''        ReDim Ly(2, 3, Fintab)
''
''        For Each dataRow In dtMatriz.Rows
''            i = (dataRow("i").ToString)
''            j = (dataRow("j").ToString)
''            h = (dataRow("h").ToString)
''            k = (dataRow("k").ToString)
''            mto = CDbl((dataRow("mto_lx").ToString))
''
''            'If (h = 1) Then
''            '    For intX = 1 To FinTab
''            '        Lx(i, j, intX) = 0
''            '    Next intX
''            'Else
''            '    For intX = 1 To FinTab
''            '        Ly(i, j, intX) = 0
''            '    Next intX
''            'End If
''
''            If h = 1 Then   'Causante
''                Lx(i, j, k) = mto
''            Else    'Beneficiario
''                Ly(i, j, k) = mto
''            End If
''
''        Next
''*-* F
'
'        cuenta = 0
'        numrec = -1
'        lrefun = 288
'
'        'Inicializacion de variables
'        Fechap = Nap * 12 + Nmp
'        perdif = 0
'        fapag = 0
'        fmpag = 0
'        fechas = 0
'
'        'Recalculo de periodo garantizado y diferido despues de la fecha de devengamiento.
'        mesdif1 = Mesdif 'debe venir en meses
'        pergar1 = pergar
'        mescon = Fechap - ((Fasolp * 12) + Fmsolp)
'        If (mescon < (mesdif1 + pergar1)) Then
'            If (mescon > mesdif1) Then
'                mescosto = mescon - mesdif1
'            Else
'                mescosto = 0
'            End If
'        Else
'            mescosto = mescon - mesdif1
'        End If
'        If (mescon > mesdif1) Then
'            Mesdif = 0
'        Else
'            If (mesdif1 > mescon) Then
'                Mesdif = (mesdif1 - mescon)
'            Else
'                Mesdif = (mescon - mesdif1)
'            End If
'        End If
'        If (mescon > (pergar1 + mesdif1)) Then
'            pergar = 0
'        Else
'            If (mescon < mesdif1) Then
'                pergar = pergar
'            Else
'                pergar = (pergar1 + mesdif1) - mescon
'            End If
'        End If
'        perdif = Mesdif
'
'        If Indi = "D" Then
'            icont10 = 0: icont20 = 0: icont11 = 0: icont21 = 0
'            icont30 = 0: icont35 = 0
'            icont40 = 0: icont77 = 0: icont30Inv = 0
'            For j = 1 To Nben
'                nibe = 0
'                If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
'                If Coinbe(j) = "N" Then nibe = 2
'                If Coinbe(j) = "P" Then nibe = 2
'                If nibe = 0 Then
'                    x = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'                nsbe = 0
'                If Sexobe(j) = "M" Then nsbe = 1
'                If Sexobe(j) = "F" Then nsbe = 2
'                If nsbe = 0 Then
'                    x = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'                Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                edabe = Fechap - Fechan
'                If edabe < 1 Then edabe = 1
'                If edabe > Fintab Then
'                    'vgError = 1023
'                    Exit Function
'                End If
'                If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Then icont10 = icont10 + 1
'                If Ncorbe(j) = 20 Or Ncorbe(j) = 21 Then icont20 = icont20 + 1
'                If Ncorbe(j) = 30 Then icont30 = icont30 + 1
'                If Ncorbe(j) = 30 And Coinbe(j) <> "N" Then icont30Inv = icont30Inv + 1
'                If Ncorbe(j) = 35 Then icont35 = icont35 + 1
'                If Ncorbe(j) > 40 And Ncorbe(j) < 50 Then icont40 = icont40 + 1
'                If Ncorbe(j) = 77 Then icont77 = icont77 + 1
'            Next j
'            If (icont10 > 0 Or icont20 > 0) And icont30 > 0 And icont30Inv = 0 Then
'                For j = 1 To Nben
'                    If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or Ncorbe(j) = 20 Or Ncorbe(j) = 21 Then
'                        edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j))) + perdif
'                        If edhm >= L18 Then
'                            Porcbe(j) = 0.42
'                        End If
'                    End If
'                Next j
'            End If
'        End If
'
'        tmtce = (1 + Tasa / 100) ^ (1 / 12)
'
'        ''If Indi = 2 Then
'        ''    ' Mesdif = Mesdif * 12
'        ''    PerDif = Mesdif
'        ''End If
'        ''rmpol = 0
'        ''If Alt = 3 Or (Alt = 4 And pergar > 0) Then Mesgar = pergar
'
'        'If Cober = 8 Or Cober = 9 Or Cober = 10 Or Cober = 11 Or Cober = 12 Then ABV 17-07-2007
'        If Cober = "S" Then
'
'            For ij = 1 To Fintab
'                Flupen(ij) = 0
'            Next ij
'            Mesgar = pergar
'            'I - KVR 17/08/2007 -
'            If Alt = "S" Then Mesgar = 0
'            'If Alt = 1 Then Mesgar = 0
'            'F - KVR 17/08/2007 -
'            nmdiga = perdif + Mesgar
'            For j = 1 To Nben
'                pension = Porcbe(j)
'                swg = "N"
'                'nrel = 0
'                nibe = 0
'                If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
'                If Coinbe(j) = "N" Then nibe = 2
'                If Coinbe(j) = "P" Then nibe = 2
'                If nibe = 0 Then
'                    x = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'                nsbe = 0
'                If Sexobe(j) = "M" Then nsbe = 1
'                If Sexobe(j) = "F" Then nsbe = 2
'                If nsbe = 0 Then
'                    x = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'
'                'Calculo de la edad de los beneficiarios
'                Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                edabe = Fechap - Fechan
'                If edabe > Fintab Then
'                    x = MsgBox("Error edad del beneficiario es mayor a la tabla de mortalidad.", vbCritical, "Proceso de cálculo Abortado")
'                    ' Renta_Vitalicia = False
'                    Exit Function
'
'                End If
'                If edabe < 1 Then edabe = 1
'                If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
'                    Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
'                    Ncorbe(j) = 41 Or Ncorbe(j) = 46 Or _
'                    Ncorbe(j) = 42 Or Ncorbe(j) = 45 Or _
'                    (Ncorbe(j) >= 30 And Ncorbe(j) < 40) And (Coinbe(j) <> "N" And edabe > L18) Then
'
'                    'PRIMA SOBREVIVENCIA VITALICIA
'                    pension = Porcbe(j)
'                    limite1 = Fintab - edabe - 1
'                    nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                    For i = 0 To limite1
'                        imas1 = i + 1
'                        edalbe = edabe + i
'                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                        If i < nmdiga Then py = 1
'                        If i < perdif Then py = 0
'                        'Flupen(imas1) = Flupen(imas1) + py * Pension
'                        Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
'                    Next i
'                    'DERECHO A ACRECER
'                    'If Codcbe(j) <> "N" Then
'                    If DerCrecer <> "N" Then
'                        edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
'                        If edhm > L18 Then
'                            nmdif = 0
'                        Else
'                            nmdif = L18 - edhm
'                        End If
'                        Ecadif = edabe + nmdif
'                        limite1 = Fintab - Ecadif - 1
'                        pension = Penben(j) * 0.2
'                        nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                        For i = 0 To limite1
'                            imas1 = nmdif + i + 1
'                            edalbe = Ecadif + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            If (i + nmdif) < nmdiga And nrel = 2 Then py = 1
'                            If (i + nmdif) < perdif Then py = 0
'                            'Flupen(imas1) = Flupen(imas1) + py * Pension
'                            Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
'                        Next i
'
'                    End If
'
'                Else
'                    If Ncorbe(j) >= 30 And Ncorbe(j) < 40 And ((Coinbe(j) = "N" And edabe <= L18)) Then
'                        'PRIMA DE PENSIONES TEMPORALES
'                        If (edabe > L18 And Coinbe(j) = "N") Then
'                            x = MsgBox("Error edad de hijo mayor a la edad legal", vbCritical)
'                        Else
'                            If edabe < L18 Then
'                                mdif = L18 - edabe
'                                nmdif = mdif - 1
'                                nmax = fgMaximo(nmax, CInt(nmdif)) '*-*nmax = amax.amax0(nmax, CInt(nmdif))
'                                For i = 0 To nmdif
'                                    imas1 = i + 1
'                                    edalbe = edabe + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    If swg = "S" And i < nmdiga Then py = 1
'                                    'En el Siscot2 estaba esta Línea ?????
'                                    If i < nmdiga Then py = 1
'                                    'Fin ???
'                                    If i < perdif Then py = 0
'                                    'Flupen(imas1) = Flupen(imas1) + py * Pension
'                                    Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
'                                Next i
'                            End If
'                        End If
'                        'PRIMA DE HIJOS INVALIDOS
'                        If Coinbe(j) <> "N" Then
'                            kdif = mdif
'                            edbedi = edabe + kdif
'                            limite3 = Fintab - edbedi - 1
'                            pension = Porcbe(j)
'                            nmax = fgMaximo(nmax, CInt(limite3)) '*-*nmax = amax.amax0(nmax, CInt(limite3))
'                            For i = 0 To limite3
'                                edalbe = edbedi + i
'                                nmdifi = i + kdif
'                                imas1 = nmdifi + 1
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                'En el Siscot2 estaba esta Línea ?????
'                                If i < nmdiga Then py = 1
'                                'Fin ???
'                                If nmdifi < perdif Then py = 0
'                                'Flupen(imas1) = Flupen(imas1) + py * Pension
'                                Flupen(imas1) = Flupen(imas1) + py * pension * facgratif(imas1)
'                            Next i
'                        End If
'                    Else
'                        x = MsgBox("Error en códificación de parentesco.", vbCritical, "Proceso de Cálculo Abortado")
'                        ' Renta_Vitalicia = False
'                        Exit Function
'
'                    End If
'                End If
'            Next j
'
'
'            ax = 0
'            For LL = 1 To nmax
'                ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
'            Next LL
'
'            If ax <= 0 Then
'                renta = 0
'                ax = 0
'            Else
'                ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
'                renta = SalCta / ax
'                '******
'                renta = CDbl(Format(renta, "#,#0.00"))
'                ax = CDbl(Format(ax, "#,#0.000000"))
'                '*******
'            End If
'            If Indi = "D" Then
'                renta = 0
'            End If
'            If Indi = "D" Then
'                vgPensionCot = renta
'            Else
'                vgPensionCot = (renta / vgFactorAjusteIPC)
'            End If
'            If (Mone = "NS") Then
'                renta = (renta / vgFactorAjusteIPC)
'            End If
'
'            vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
'            vlMtoPenSim = Format(renta, "##0.00")
'            vlPenGar = Format(0, "##0.00") 'sumaqx
'            '----------------------------------------------------------------------
'            If Indi = "D" Then 'calcula diferida
'                add_porc_ben = 0
'                Vpptem = 0
'                Tasa_afp = 0
'                'Prima_unica = 0
'                Rete_sim = 0
'                Prun_sim = 0
'                Sald_sim = 0
'                mesga2 = 0
'                vlMoneda = ""
'                vgPensionCot = 0
'
'                vlMoneda = Mone
'                If Mesdif > 0 Then
'                    gto_sepelio = vlPenGar
'                    If Prc_Tasa_Afp = 0 Then
'                        Vpptem = 0
'                    Else
'                        Vpptem = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ Mesdif)) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                        Vpptem = CDbl(Format(CDbl(Vpptem), "##0.000000"))
'                    End If
'                End If
'
'                'tasa_afp=rentabilidad de la afp
'
'                If vlPriUniSim > 0 Then
'                    If (vlMoneda = "NS") Then
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * (Dif!mto_priunisim * vgFactorAjusteIPC))))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * (vlPriUniSim * vgFactorAjusteIPC))))), "##0.00"))
'                        End If
'                    Else
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * vlPriUniSim)))), "##0.00"))
'                        End If
'                    End If
'                End If
'
'
'                'Saldo necesario AFP
'                Sald_sim = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
'
'                If vlPriUniSim > 0 Then
'                    Prun_sim = CDbl(Format(CDbl(Prima_unica - Sald_sim), "#,#0.00"))
'                    'prun_sim = CDbl(Format(CDbl(Salcta_Sol - Sald_sim), "#,#0.00"))
'                    vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "#,#0.00"))
'
'                    If (vlMoneda = "NS") Then
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / (vlPriUniSim * vgFactorAjusteIPC)), "##0.00"))
'                    Else
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "##0.00"))
'                    End If
'                    vlPensim = vgPensionCot
'                Else
'                    vgPensionCot = 0
'                End If
'                Mto_ValPrePenTmp = Vpptem
'                vlMtoPenSim = vlPensim 'hqr para reporte
'                vlMtoPriUniDif = Prun_sim
'                vlMtoCtaIndAfp = Sald_sim
'                vlRtaTmpAFP = Rete_sim
'            End If
'            Dim vlSumaPension As Double
'            'Registrar los valores de la Pensión para cada
'            'Beneficiario para el caso de Sobrevivencia
'            vlSumaPension = 0
'            For vlI = 1 To Nben
'                If (Ncorbe(vlI) <> 0) Then
'                    If (Ncorbe(vlI) = 10) And (Ffam > 1) Then
'                        vlSumaPension = vlSumaPension + CDbl(Format((vgPensionCot * (porfam / 100)), "#0.00"))
'                    Else
'                        vlSumaPension = vlSumaPension + CDbl(Format((vgPensionCot * Porcbe(vlI)), "#0.00"))
'                    End If
'                End If
'            Next vlI
'            vlSumPension = Format(vlSumaPension, "#0.00")
'
'        Else
'
'            'FLUJOS DE RENTAS VITALICIAS DE VEJEZ E INVALIDEZ
'            qx = 0
'            For ij = 1 To Fintab
'                Flupen(ij) = 0
'            Next ij
'            facfam = Ffam
'
'            'I - KVR 17/08/2007 -
'            If Alt = "S" Then Mesgar = 0
'            If Alt = "G" Or (Alt = "F" And Mesgar > 0) Then Mesgar = pergar 'Mesgar
'
'            'If Alt = "F" Or Alt = "S" Then Mesgar = 0
'            'If Alt = "G" Then Mesgar = pergar
'            'If Alt = "S" Or Alt = "F" Then Mesgar = 0
'            'F - 17/08/2007 -
'
'            'Definicion del periodo garantizado en 0 para las
'            'alternativas Simple o pensiones con distinto porcentaje legal
'            For j = 1 To Nben
'                'If derpen(j) <> 10 Then
'                pension = Porcbe(j)
'                'CALCULO DE LA PRIMA DEL AFILIADO
'                If (Ncorbe(j) = 0 Or Ncorbe(j) = 99) And j = 1 Then
'                    ni = 0
'                    If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then ni = 1
'                    If Coinbe(j) = "N" Then ni = 2
'                    If Coinbe(j) = "P" Then ni = 3
'                    If ni = 0 Then
'                        x = MsgBox("Error de códificación de tipo de inavlidez", vbCritical, "Proceso de Cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    ns = 0
'                    If Sexobe(j) = "M" Then ns = 1
'                    If Sexobe(j) = "F" Then ns = 2
'                    If ns = 0 Then
'                        x = MsgBox("Error de códificación de sexo.", vbCritical, "Proceso de Cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    'Calculo de edad del causante
'                    Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                    edaca = Fechap - Fechan
'                    If edaca < 780 And ns = 1 And ni = 2 Then cplan = "A"
'                    If edaca < 720 And ns = 2 And ni = 2 Then cplan = "A"
'                    If edaca <= 216 Then edaca = 216
'                    If edaca > Fintab Then
'                        x = MsgBox("Error en edad del beneficiario mayor a tabla de mortalidad", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    sumaqx = 0
'                    limite1 = CInt(Fintab - edaca - 1)
'                    nmax = CInt(limite1)
'                    For i = 0 To limite1
'                        imas1 = i + 1
'                        edacai = edaca + i
'                        If edacai > Fintab Then
'                            x = MsgBox("Edad fuera de Rangos establecidos.", vbCritical, "Proceso de Cálculo Abortado")
'                            'Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                        px = Lx(ns, ni, edacai) / Lx(ns, ni, edaca)
'                        edacas = edacai + 1
'                        qx = ((Lx(ns, ni, edacai) - Lx(ns, ni, edacas))) / Lx(ns, ni, edaca)
'                        'Flupen(imas1) = Flupen(imas1) + px * Pension
'                        Flupen(imas1) = Flupen(imas1) + px * pension * facgratif(imas1)
'                        sumaqx = sumaqx + GtoFun * qx / tmtce ^ (i + 0.5)
'                    Next i
'                End If
'
'                If Ncorbe(j) <> 0 And Ncorbe(j) <> 99 Then
'                    'Prima de los Beneficiarios
'                    nibe = 0
'                    If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
'                    If Coinbe(j) = "N" Then nibe = 2
'                    If Coinbe(j) = "P" Then nibe = 3
'                    If nibe = 0 Then
'                        x = MsgBox("Error de códificación de tipo de invalidez.", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    nsbe = 0
'                    If Sexobe(j) = "M" Then nsbe = 1
'                    If Sexobe(j) = "F" Then nsbe = 2
'                    If nsbe = 0 Then
'                        x = MsgBox("Error de códificación de sexo.", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    'Calculo de la edad del beneficiario
'                    edabe = Fechap - (Nanbe(j) * 12 + Nmnbe(j))
'                    If edabe < 1 Then edabe = 1
'                    If edabe > Fintab Then
'                        x = MsgBox("Error Edad del beneficario es mayor a la tabla.", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
'                       Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
'                       Ncorbe(j) = 41 Or Ncorbe(j) = 46 Or _
'                       Ncorbe(j) = 42 Or Ncorbe(j) = 45 Or _
'                       ((Ncorbe(j) >= 30 And Ncorbe(j) < 40) And _
'                       (Coinbe(j) <> "N" And edabe > L18)) Then
'                        'FLUJOS DE VIDAS CONJUNTAS VITALICIAS
'                        'Probabilidad del beneficiario solo
'                        limite1 = Fintab - edabe - 1
'                        nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                        For i = 0 To limite1
'                            imas1 = i + 1
'                            edalbe = edabe + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            'Flupen(imas1) = Flupen(imas1) + py * Pension * facfam
'                            Flupen(imas1) = Flupen(imas1) + py * pension * facfam * facgratif(imas1)
'                        Next i
'                        'Probabilidad conjunta de causante y beneficiario
'                        limite2 = Fintab - edaca - 1
'                        limite = fgMinimo(limite1, CInt(limite2)) '*-* limite = amax.amin0(limite1, CInt(limite2))
'                        nmax = fgMaximo(nmax, CInt(limite)) '*-* nmax = amax.amax0(nmax, CInt(limite))
'                        For i = 0 To limite
'                            imas1 = i + 1
'                            edalca = edaca + i
'                            edalbe = edabe + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                            'Flupen(imas1) = Flupen(imas1) - (py * px * Pension * facfam)
'                            Flupen(imas1) = Flupen(imas1) - (py * px * pension * facfam) * facgratif(imas1)
'                        Next i
'
'                        'DERECHO A ACRECER
'                        'If Codcbe(j) <> "N" Then
'                        If DerCrecer <> "N" Then
'                            edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
'                            If edhm > L18 Then
'                                nmdif = 0
'                            Else
'                                nmdif = L18 - edhm
'                            End If
'                            Ecadif = edabe + nmdif
'                            limite1 = Fintab - Ecadif - 1
'                            pension = Porcbe(j) * 0.2
'                            nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                            For i = 0 To limite1
'                                imas1 = nmdif + i + 1
'                                edalbe = Ecadif + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                'Flupen(imas1) = Flupen(imas1) + py * Pension * facfam
'                                Flupen(imas1) = Flupen(imas1) + py * pension * facfam * facgratif(imas1)
'                            Next i
'
'                            Ecadif = edaca + nmdif
'                            limite4 = Fintab - Ecadif - 1
'                            limite = fgMinimo(limite1, CInt(limite4)) '*-* limite = amax.amin0(limite1, CInt(limite4))
'
'                            For i = 0 To limite
'                                imas1 = nmdif + i + 1
'                                edalbe = Ecadif + i
'                                edalca = (edaca + nmdif) + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                'Flupen(imas1) = Flupen(imas1) - (py * px * Pension * facfam)
'                                Flupen(imas1) = Flupen(imas1) - (py * px * pension * facfam) * facgratif(imas1)
'                            Next i
'
'                        End If
'
'                    Else
'                        If Ncorbe(j) >= 30 And Ncorbe(j) < 40 Then
'                            'Prima Rentas Temporales
'                            If edabe <= L18 Then
'                                mdif = L18 - edabe
'                                nmdif = mdif
'                                'Probabilidad conjunta del causante y beneficiario
'                                limite2 = Fintab - edaca
'                                limite = fgMinimo(nmdif, CInt(limite2)) - 1 '*-* limite = amax.amin0(nmdif, CInt(limite2)) - 1
'                                nmax = fgMaximo(nmax, CInt(limite)) '*-* nmax = amax.amax0(nmax, CInt(limite))
'                                For i = 0 To limite
'                                    imas1 = i + 1
'                                    edalca = edaca + i
'                                    edalbe = edabe + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                    'Flupen(imas1) = Flupen(imas1) + ((py * Pension) - (py * px * Pension)) * facfam
'                                    Flupen(imas1) = Flupen(imas1) + ((py * pension) - (py * px * pension)) * facfam * facgratif(imas1)
'                                Next i
'                            Else
'                                'I---- ABV 06/01/2004 ---
'                                'Verificar la acción a realizar
'                                'para que continue el proceso, ya que los
'                                'Hijos mayores a la Edad Legal no se deben
'                                'calcular.
'                                'Preguntar a Daniela la implicancia de esto
'                                'x = MsgBox("La Edad del Hijo es mayor a la Edad Legal.", vbCritical, "Proceso de Cálculo Abortado")
'                                'Renta_vitalicia = False
'                                'Exit Function
'                                'F---- ABV 06/01/2004 ---
'                            End If
'                            'Prima del Hijo Invalido
'                            If Coinbe(j) <> "N" Then
'                                'Probabilidad conjunta del causante y beneficiario
'                                edbedi = edabe + nmdif
'                                limite3 = Fintab - edbedi - 1
'                                limite4 = Fintab - (edaca + nmdif) - 1
'                                nmax = fgMaximo(nmax, CInt(limite3)) '*-* nmax = amax.amax0(nmax, CInt(limite3))
'                                For i = 0 To limite3
'                                    nmdifi = nmdif + i
'                                    imas1 = nmdifi + 1
'                                    edalca = edaca + nmdif + i
'                                    edalbe = edbedi + i
'                                    edalca = fgMinimo(edalca, CInt(Fintab)) '*-* edalca = amax.amin0(edalca, CInt(Fintab))
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                    'Flupen(imas1) = Flupen(imas1) + (py - (py * px)) * Pension * facfam
'                                    Flupen(imas1) = Flupen(imas1) + (py - (py * px)) * pension * facfam * facgratif(imas1)
'                                Next i
'                            End If
'                        Else
'                            x = MsgBox("Error en codificación del parentesco.", vbCritical, "Proceso de cálculo Abortado")
'                            'Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                    End If
'                End If
'                'End If
'            Next j
'
'            'Calculo de tarifa y Pension
'            ax = 0
'            flumax = 0
'                If Alt = "S" Then  'KVR 17/08/2007 SE ELIMINO DE ESTA LINEA Alt = "F"
'                Mesgar = 0
'                nmdiga = perdif + Mesgar
'                ax = 0
'                For LL = 1 To nmax
'                    flumax = Flupen(LL)
'                    If LL <= perdif Then flumax = 0
'                    ax = ax + flumax / tmtce ^ (LL - 1)
'                Next LL
'            Else
'                Mesgar = pergar
'                nmdiga = perdif + Mesgar
'                ax = 0
'                For LL = 1 To nmdiga
'                    flumax = fgMaximo(1, Flupen(LL)) '*-* flumax = amax.amax1(1, Flupen(LL))
'                    If LL <= perdif Then flumax = 0
'                    ax = ax + flumax / tmtce ^ (LL - 1)
'                Next LL
'                For LL = nmdiga + 1 To nmax
'                    ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
'                Next LL
'            End If
'
'            If Indi = "I" Then
'                If ax <= 0 Then
'                    renta = 0
'                    ax = 0
'                Else
'                    sumaqx = CDbl(Format(sumaqx, "#,#0.000000"))
'                    ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
'                    renta = CDbl(Format(((SalCta - sumaqx) / ax), "#,#0.00"))
'                End If
'            Else
'                ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
'                renta = 0
'            End If
'            If (Mone = "NS") Then
'                renta = (renta * vgFactorAjusteIPC)
'            End If
'            'HQR 24/05/2004
'            vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
'            vlMtoPenSim = Format(renta, "##0.00")
'            vlPenGar = Format(sumaqx, "##0.00")
'            'FIN HQR 24/05/2004
'
'            If Indi = "D" Then
'                'aca se iba a la función que calcula diferida
'                add_porc_ben = 0
'                Vpptem = 0
'                Tasa_afp = 0
'                ' Prima_unica = 0
'                Rete_sim = 0
'                Prun_sim = 0
'                Sald_sim = 0
'                mesga2 = 0
'                vlMoneda = ""
'                vgPensionCot = 0
'
'                vlMoneda = Mone
'                If Mesdif > 0 Then
'                    gto_sepelio = vlPenGar
'                    If Prc_Tasa_Afp = 0 Then
'                        Vpptem = 0
'                    Else
'                        Vpptem = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ Mesdif)) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                        Vpptem = CDbl(Format(CDbl(Vpptem), "##0.000000"))
'                    End If
'                End If
'
'                'tasa_afp=rentabilidad de la afp
'
'                If vlPriUniSim > 0 Then
'                    If (vlMoneda = "NS") Then
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * (Dif!mto_priunisim * vgFactorAjusteIPC))))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * (vlPriUniSim * vgFactorAjusteIPC))))), "##0.00"))
'                        End If
'                    Else
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * vlPriUniSim)))), "##0.00"))
'                        End If
'                    End If
'                End If
'
'                'Saldo necesario AFP
'                Sald_sim = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
'
'                If vlPriUniSim > 0 Then
'                    Prun_sim = CDbl(Format(CDbl(Prima_unica - Sald_sim), "#,#0.00"))
'                    'prun_sim = CDbl(Format(CDbl(Salcta_Sol - Sald_sim), "#,#0.00"))
'                    vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "#,#0.00"))
'
'                    If (vlMoneda = "NS") Then
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / (vlPriUniSim * vgFactorAjusteIPC)), "##0.00"))
'                    Else
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "##0.00"))
'                    End If
'                    vlPensim = vgPensionCot
'                Else
'                    vgPensionCot = 0
'                End If
'                Mto_ValPrePenTmp = Vpptem
'                vlMtoPenSim = vlPensim 'hqr para reporte
'                vlMtoPriUniDif = Prun_sim
'                vlMtoCtaIndAfp = Sald_sim
'                vlRtaTmpAFP = Rete_sim
'            End If
'        End If
'
'        'llenar el datatable
'        istPolizas.Mto_ValPrePenTmp = Mto_ValPrePenTmp 'dtDataRow("MTO_VALPREPENTMP") = mto_valprepentmp
'        istPolizas.Mto_Pension = vlMtoPenSim 'dtDataRow("MTO_PENSION") = vlMtoPenSim
'        istPolizas.Mto_PriUniDif = vlMtoPriUniDif 'dtDataRow("MTO_PRIUNIDIF") = vlMtoPriUniDif
'        istPolizas.Mto_CtaIndAFP = vlMtoCtaIndAfp 'dtDataRow("MTO_CTAINDAFP") = vlMtoCtaIndAfp
'        istPolizas.Mto_RentaTMPAFP = vlRtaTmpAFP 'dtDataRow("MTO_RENTATMPAFP") = vlRtaTmpAFP
'        istPolizas.Mto_PriUniSim = vlPriUniSim 'dtDataRow("MTO_PRIUNISIM") = vlPriUniSim
'        'I - KVR 13/08/2007 -
'        'dtDataRow("MTO_PENSIONGAR") = vlMtoPenSim
'        istPolizas.Mto_RMGtoSepRV = vlPenGar 'dtDataRow("MTO_PENSIONGAR") = vlPenGar
'        istPolizas.Mto_SumPension = Format(vlSumPension, "#0.00") 'dtDataRow("MTO_SUMPENSION") = Format(vlSumPension, "#0.00")
'        'F - KVR 13/08/2007 -
'
''*-*        'I - KVR 18/08/2007 -
''*-*        If Mone = vgMonedaOficial And Indi = "D" Then
''*-*            dtDataRow("MTO_AJUSTEIPC") = Format(vgFactorAjusteIPC, "#0.00000000")
''*-*        Else
''*-*            dtDataRow("MTO_AJUSTEIPC") = 1
''*-*        End If
''*-*        'F - KVR 18/08/2007 -
'
'        If (Mesgar > 0) Then
'            istPolizas.Mto_PensionGar = vlMtoPenSim
'        Else
'            istPolizas.Mto_PensionGar = 0
'        End If
''*-*        dtDataRow("NUM_CORRELATIVO") = vlCorrCot
'
'        Mto_ValPrePenTmp = 0
'        vlMtoPenSim = 0
'        vlMtoPriUniDif = 0
'        vlMtoCtaIndAfp = 0
'        vlRtaTmpAFP = 0
'        vlPriUniSim = 0
'        vlPenGar = 0
'        vlCorrCot = 0
''*-*        Renta_Vitalicia.Rows.Add (dtDataRow)
'
'    Next vlNumero
'
'    fgCalcularRentaVitalicia_Old2 = True
'
End Function

Function fgCalcularRentaVitalicia_Old(istPolizas As TyPoliza, istBeneficiarios() As TyBeneficiarios, Coti As String, codigo_afp As String, iRentaAFP As Double, iNumCargas As Integer) As Boolean
'Dim Prodin() As Double
'Dim Flupen() As Double, Flucm() As Double, Exced() As Double
'Dim impres(9, 110) As Double
'Dim Ncorbe(20) As Integer
'Dim Penben(20) As Double, Porcbe(20) As Double, porcbe_ori(20) As Double
'Dim Coinbe(20) As String, Codcbe(20) As String, Sexobe(20) As String
'Dim Nanbe(20) As Integer, Nmnbe(20) As Integer, Ndnbe(20) As Integer
'Dim Ijam(20) As Integer, Ijmn(20) As Integer, Ijdn(20) As Integer
'Dim Npolbe(20) As String, derpen(20) As Integer
'Dim i As Integer
'Dim Totpor As Double
'Dim cob(5) As String, alt1(3) As String, tip(2) As String
'
'Dim Npolca As String, Mone As String
'Dim Cober As String, Alt As String, Indi As String, cplan As String
'Dim Nben As Long
'Dim Nap As Integer, Nmp As Integer
'Dim Fechan As Long, Fechap As Long
'Dim Mesdif As Long, Mesgar As Long
'Dim Bono As Double, Bono_Pesos1 As Double, GtoFun As Double
'Dim CtaInd As Double, SalCta As Double, Salcta_Sol As Double
'Dim Ffam As Double, porfam As Double
'Dim Prc_Tasa_Afp As Double, Prc_Pension_Afp As Double
'Dim vgs_Coti As String
'
'Dim edbedi As Long, mdif As Long
'Dim large As Integer
'Dim edaca As Long, edalca As Long, edacai As Long, edacas As Long, edabe As Long, edalbe As Long
'Dim Fasolp As Long, Fmsolp As Long, Fdsolp As Long, pergar As Long, numrec As Long, numrep As Long
'Dim nrel As Long, nmdif As Long, numbep As Long, ni As Long, ns As Long, nibe As Long, nsbe As Long, limite As Long
'Dim limite1 As Long, limite2 As Long, limite3 As Long, limite4 As Long, imas1 As Long, kdif As Long, nt As Long
'Dim nmax As Integer, j As Integer
'Dim rmpol As Double, px As Double, py As Double, qx As Double, relres As Double
'Dim comisi As Double, tasac As Double, timp As Double, tm As Double, tmm As Double
'Dim gasemi As Double
'Dim tm3 As Double, sumapx As Double, sumaqx As Double, actual As Double, actua1 As Double, PenBase As Double, tce As Double
'Dim vpte As Double, difres As Double, difre1 As Double, tir As Double, tinc As Double
'Dim Tasa As Double, tastce As Double, tirvta As Double, tvmax As Double
'Dim vppen As Double, vpcm As Double, penanu As Double, reserva As Double, gastos As Double, rdeuda As Double
'Dim resfin As Double, rend As Double, varrm As Double, resant As Double, flupag As Double, gto As Double
'Dim sumaex As Double, sumaex1 As Double, tirmax As Double
'Dim Sql As String, Numero As String
'Dim Linea1 As String
'Dim Inserta As String
'Dim Var As String, Nombre As String
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
'Dim vlContarMaximo As Long, vlMtoCtaIndAfp As Double
'
'Dim vlCorrCot As Integer, vgd_tasa_vta As Double, FecDev As String, FecCot As String
'Dim h As Integer, Cor As Integer
'Dim intX As Integer, k As Integer, ltot As Long, comint As Double
'Dim mto As Double
'Dim add_porc_ben As Double, gto_sepelio As Double, mesga2 As Double
'Dim vlFechaNacCausante As String
'Dim vlSexoCausante As String, vlMoneda As String
'Dim sumaporcsob As Double, vgPensionCot As Double, mto_valprepentmp As Double
'Dim fapag As Long, fechas As Long, mesdif1 As Long, pergar1 As Long, mescon As Long
'Dim fmpag As Integer
'Dim icont10 As Integer, icont20 As Integer, icont11 As Integer, icont21 As Integer
'Dim icont30 As Integer, icont35 As Integer, icont30Inv As Integer
'Dim icont40 As Integer, icont77 As Integer
'Dim vlSumPension As Double, MtoMoneda As Double
'Dim DerCrecer As String, DerGratificacion As String
'Dim lrefun As Long
'
''YO
'Dim LL As Integer, ij As Integer
'Dim X As Long
'Dim perdif As Long
'Dim nmdiga As Long
'Dim edhm As Long
'Dim swg As String
'Dim flumax  As Double
'Dim pension  As Double
'Dim renta  As Double
'Dim ax As Double
'Dim vlNumero As Integer
'Dim Tasa_afp As Double
'Dim Navig As Integer, Nmvig As Integer, Ndvig As Integer
'
''---------------------------------------------------------------------------
''Ultima Modificación realizada el 18/12/2005
''Agregar Tablas de Mortalidad con lectura desde la BD y no desde un archivo
''Además, manejar dichas tablas por Fecha de Vigencia, para que opere la que
''corresponda a la Fecha de Cotización
''La Tabla de Mortalidad en esta función es MENSUAL
''---------------------------------------------------------------------------
'
'
'    'Lee y Calcula por Modalidad, en este caso solo se trata de una
'    'se pasan los parametros a variables
'    For vlNumero = 1 To 1
'        cuenta = 1
'
'        'vgd_tasa_vta = 0
'        'Npolca = (dr("Num_Cot").ToString)
'        'Fintab = (dr("FinTab").ToString)
''*-* I Agregado por ABV
'        Navig = Mid(istPolizas.Fec_Vigencia, 1, 4)
'        Nmvig = Mid(istPolizas.Fec_Vigencia, 5, 2)
'        Ndvig = Mid(istPolizas.Fec_Vigencia, 7, 2)
''*-* F
'
'        Nben = iNumCargas
'        If istPolizas.Cod_TipPension = "08" Then Nben = Nben - 1
'
'        Cober = istPolizas.Cod_TipPension '(dr("Plan").ToString)
'        Indi = istPolizas.Cod_TipRen 'CInt((dr("Indicador").ToString))   ' I o D
'        Alt = istPolizas.Cod_Modalidad '(dr("Alternativa").ToString)
'        'I - KVR 06/08/2007 -
'        pergar = istPolizas.Num_MesGar 'CLng(dr("MesGar").ToString)
'        'Mesgar = CLng(dr("MesGar").ToString)
'        'F - KVR 06/08/2007 -
'        Mone = istPolizas.Cod_Moneda '(dr("Moneda").ToString) 'vgMonedaOficial ABV 17-07-2007
'        FecCot = istPolizas.Fec_Calculo '(dr("FecCot").ToString)
'        Nap = CInt(Mid(FecCot, 1, 4))
'        Nmp = CInt(Mid(FecCot, 5, 2))
'        'I - KVR 06/08/2007 -
'        'idp = CInt(Mid((dr("FecCot").ToString), 7, 2))
'        'Bono = CDbl((dr("Mto_BonoAct").ToString))
'        'Bono_Pesos1 = CDbl((dr("Mto_BonoActPesos").ToString))
'        CtaInd = istPolizas.Mto_CtaInd 'CDbl((dr("CtaInd").ToString))   'EN SOLES
'        Prima_unica = istPolizas.Mto_PriUniMod  'CDbl((dr("Mto_PriUni").ToString))
'        SalCta = istPolizas.Mto_PriUniMod  'CDbl((dr("Mto_PriUni").ToString)) ' SIEMPRE VIENE EN LA MONEDA DE LA MODALIDAD
'        'F - KVR 06/08/2007 -
'        Ffam = istPolizas.Mto_FacPenElla  'CDbl((dr("FacPenElla").ToString))
'        'I - KVR 06/08/2007 - se agrega porcentaje pensando en ella
'        porfam = istPolizas.Prc_FacPenElla 'CDbl((dr("PrcFacPenElla").ToString))
'        'F - KVR 06/08/2007 -
'        Mesdif = istPolizas.Num_MesDif 'CLng((dr("MesDif").ToString))
'        FecDev = istPolizas.Fec_Dev '(dr("FecDev").ToString)
'        Fasolp = CLng(Mid(FecDev, 1, 4))   'a_sol_pen
'        Fmsolp = CLng(Mid(FecDev, 5, 2))  'm_sol_pen
'        Fdsolp = CLng(Mid(FecDev, 7, 2))    'd_sol_pen
'        GtoFun = istPolizas.Mto_CuoMor 'CDbl((dr("Gassep").ToString))  'siempre es en soles
'        MtoMoneda = istPolizas.Mto_ValMoneda 'CDbl(dr("MtoMoneda").ToString)
'        If Mone <> vgMonedaCodOfi Then
'            GtoFun = Format(CDbl(GtoFun / MtoMoneda), "#0.00000")
'        End If
'        vlCorrCot = istPolizas.Num_Correlativo 'CInt((dr("Num_Correlativo").ToString))
'        Tasa = istPolizas.Prc_TasaVta 'CDbl((dr("Prc_TasaVta").ToString))
'        Tasa = Format(Tasa, "#0.00")
'        'I - KVR 06/08/2007 - comente estos campos ya que no aparecen en funcion de Daniela
'        Prc_Tasa_Afp = istPolizas.Prc_RentaAFP / 100 'CDbl((dr("RtaAfp").ToString)) / 100
'        Tasa_afp = istPolizas.Prc_RentaAFP / 100
'        Prc_Pension_Afp = istPolizas.Prc_RentaTMP / 100 'CDbl((dr("RtaTmp").ToString)) / 100
'        comint = istPolizas.Prc_CorCom 'CDbl((dr("Prc_ComCor").ToString))
'        'F - KVR 06/08/2007 -
'        'I - KVR 11/08/2007 -
'        DerCrecer = istPolizas.Cod_DerCre '(dr("DerCre").ToString)  ' S/N Variable si tiene o no Derecho a Crecer la modalidad
'        DerGratificacion = istPolizas.Cod_DerGra '(dr("DerGra").ToString)  ' S/N Variable si tiene o no Derecho a Gratificación
'        'F - KVR 11/08/2007 -
'
'        'La conversión de estos códigos debe ser corregida a la Oficial
'        If Cober = "08" Then Cober = "S"
'        If Cober = "06" Then Cober = "I"
'        If Cober = "07" Then Cober = "P"
'        If Cober = "04" Or Cober = "05" Then Cober = "V"
'        'SalCta = Salcta_Sol
'        Totpor = 0
'        'I - KVR 06/08/20007 -
'        If Indi = "1" Then Indi = "I"
'        If Indi = "2" Then Indi = "D"
'
'        If Alt = "1" Then Alt = "S"
'        If Alt = "3" Then Alt = "G"
'        If Alt = "4" Then Alt = "F"
'        'F - KVR 06/08/2007 -
'        'Obtiene los Datos de los Beneficiarios
''*-*        If vlCorrCot = 1 Then
'        If vlNumero = 1 Then
'            i = 1
'
'            For i = 1 To iNumCargas
'                Ncorbe(i) = istBeneficiarios(i).Cod_Par '(dRow("Parentesco").ToString)
'                Porcbe(i) = istBeneficiarios(i).Prc_Pension '(dRow("Porcentaje").ToString)
'                'I - KVR 06/08/2007 -
'                porcbe_ori(i) = istBeneficiarios(i).Prc_PensionLeg '(dRow("PorcentajeLeg").ToString)
'                If (Ncorbe(i) = 99) Or (Ncorbe(i) = 0) Then
'                    vlFechaNacCausante = istBeneficiarios(i).Fec_NacBen '(dRow("Fec.Nac.").ToString)
'                    vlSexoCausante = istBeneficiarios(i).Cod_Sexo '(dRow("Sexo").ToString)
'                End If
'                'F - KVR 06/08/2007 -
'                Dim fecha As String
'                fecha = istBeneficiarios(i).Fec_NacBen '(dRow("Fec.Nac.").ToString)
'                Nanbe(i) = Mid(fecha, 1, 4)  'aa_nac
'                Nmnbe(i) = Mid(fecha, 5, 2) 'mm_nac
'                Ndnbe(i) = Mid(fecha, 7, 2) 'mm_nac
'                Sexobe(i) = istBeneficiarios(i).Cod_Sexo '(dRow("Sexo").ToString)
'                Coinbe(i) = istBeneficiarios(i).Cod_SitInv '(dRow("Sit.Inv.").ToString)
'                Codcbe(i) = istBeneficiarios(i).Cod_DerCre '(dRow("Dº Crecer").ToString)
'                'If Len((dRow("Fec.Nac.HM").ToString)) > 0 Then
'                If Len(istBeneficiarios(i).Fec_NacHM) > 0 Then
'                    fecha = ""
'                    fecha = istBeneficiarios(i).Fec_NacHM '(dRow("Fec.Nac.HM").ToString)
'                    Ijam(i) = Mid(fecha, 1, 4)  'aa_hijom
'                    Ijmn(i) = Mid(fecha, 5, 2)    'mm_hijom
'                    Ijdn(i) = Mid(fecha, 7, 2)    'mm_hijom
'                Else
'                    Ijam(i) = "0000" ' Year(tb_difben!fec_nachm)   'aa_hijom
'                    Ijmn(i) = "00" 'Month(tb_difben!fec_nachm)     'mm_hijom
'                    Ijdn(i) = "00" 'Month(tb_difben!fec_nachm)     'mm_hijom
'                End If
'                Npolbe(i) = istPolizas.Num_Cot
'                Porcbe(i) = Porcbe(i) / 100
'                porcbe_ori(i) = porcbe_ori(i) / 100
'                If Cober = "S" And (Ncorbe(i) <> 0 Or Ncorbe(i) <> 99) Then sumaporcsob = sumaporcsob + Porcbe(i)
'
'                'Penben(i) = Porcbe(i)
'                'derpen(i) = (dRow("Dº Pension").ToString) 'Dº Pensión
'                'If derpen(i) <> 10 Then
'                '    If Cober <> "S" Then
'                '        If Ncorbe(i) <> 99 Then
'                '            Totpor = Totpor + Porcbe(i)
'                '        End If
'                '    Else
'                '        Totpor = Totpor + Porcbe(i)
'                '    End If
'                'End If
''*-*                i = i + 1
'            Next i
'
''*-* I Dentro del VB ya se encuentran registradas en el L24, L21, L18
''            ' Nben = i - 1
''            'validar los topes de edad de pago de pensiones
''            Dim LimEdad As New Limite_Edad
''            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope24, FecDev)
''            If LimEdad.Mensaje = Nothing Then
''                L24 = (LimEdad.LimEdad)
''            Else
''                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
''                Exit Function
''            End If
''
''            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope21, FecDev)
''            If LimEdad.Mensaje = Nothing Then
''                L21 = (LimEdad.LimEdad)
''            Else
''                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
''                Exit Function
''            End If
''
''            LimEdad.Limite_Edad(vgCodTabla_LimEdad, vgEdadtope18, FecDev)
''            If LimEdad.Mensaje = Nothing Then
''                L18 = (LimEdad.LimEdad)
''            Else
''                MsgBox(LimEdad.Mensaje, MsgBoxStyle.Critical, "Falta Información")
''                Exit Function
''            End If
''
''            If fgValorFactorAjusteIPC(FecDev, FecCot) = False Then
''                vgFactorAjusteIPC = 0
''            End If
''
''            'I - KVR 11/08/2007 - SOLO UNA VEZ
''            L24 = L24 * 12
''            L21 = L21 * 12
''            L18 = L18 * 12
''            'F - KVR 11/08/2007 -
''*-* F
'        End If
'
'        If Cober = "S" And sumaporcsob > 1 And porfam > 0 Then
'            X = MsgBox("La suma de los porcentajes de pensión corregidos por factor Pensando en Ella es mayor al 100%.", vbCritical, "Proceso de cálculo Abortado")
'            'Renta_Vitalicia = False
'            Exit Function
'        End If
'
'        'ReDim Cp(Fintab)
'        'ReDim Prodin(Fintab)
'        ReDim Flupen(Fintab)
'        'ReDim Flucm(Fintab)
'        'ReDim Exced(Fintab)
'
''*-* I Modificación de Carga de Tablas de Mortalidad
'        '-------------------------------------------------
'        'Leer Tabla de Mortalidad
'        '-------------------------------------------------
'        If (fgBuscarMortalidadNormativa(Navig, Nmvig, Ndvig, Nap, Nmp, Ndp, vlSexoCausante, vlFechaNacCausante) = False) Then
'            'Se produjeron errores en la obtención de una de las Tablas de Mortalidad
'            Exit Function
'        End If
'
''        'llenar las matrices Lx y Ly
''        ReDim Lx(2, 3, Fintab)
''        ReDim Ly(2, 3, Fintab)
''
''        For Each dataRow In dtMatriz.Rows
''            i = (dataRow("i").ToString)
''            j = (dataRow("j").ToString)
''            h = (dataRow("h").ToString)
''            k = (dataRow("k").ToString)
''            mto = CDbl((dataRow("mto_lx").ToString))
''
''            'If (h = 1) Then
''            '    For intX = 1 To FinTab
''            '        Lx(i, j, intX) = 0
''            '    Next intX
''            'Else
''            '    For intX = 1 To FinTab
''            '        Ly(i, j, intX) = 0
''            '    Next intX
''            'End If
''
''            If h = 1 Then   'Causante
''                Lx(i, j, k) = mto
''            Else    'Beneficiario
''                Ly(i, j, k) = mto
''            End If
''
''        Next
''*-* F
'
'        cuenta = 0
'        numrec = -1
'        lrefun = 288
'
'        'Inicializacion de variables
'        Fechap = Nap * 12 + Nmp
'        perdif = 0
'        fapag = 0
'        fmpag = 0
'        fechas = 0
'
'        'Recalculo de periodo garantizado y diferido despues de la fecha de devengamiento.
'        mesdif1 = Mesdif 'debe venir en meses
'        pergar1 = pergar
'        mescon = Fechap - ((Fasolp * 12) + Fmsolp)
'        If (mescon < (mesdif1 + pergar1)) Then
'            If (mescon > mesdif1) Then
'                mescosto = mescon - mesdif1
'            Else
'                mescosto = 0
'            End If
'        Else
'            mescosto = mescon - mesdif1
'        End If
'        If (mescon > mesdif1) Then
'            Mesdif = 0
'        Else
'            If (mesdif1 > mescon) Then
'                Mesdif = (mesdif1 - mescon)
'            Else
'                Mesdif = (mescon - mesdif1)
'            End If
'        End If
'        If (mescon > (pergar1 + mesdif1)) Then
'            pergar = 0
'        Else
'            If (mescon < mesdif1) Then
'                pergar = pergar
'            Else
'                pergar = (pergar1 + mesdif1) - mescon
'            End If
'        End If
'        perdif = Mesdif
'
'        If Indi = "D" Then
'            icont10 = 0: icont20 = 0: icont11 = 0: icont21 = 0
'            icont30 = 0: icont35 = 0
'            icont40 = 0: icont77 = 0: icont30Inv = 0
'            For j = 1 To Nben
'                nibe = 0
'                If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
'                If Coinbe(j) = "N" Then nibe = 2
'                If Coinbe(j) = "P" Then nibe = 2
'                If nibe = 0 Then
'                    X = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'                nsbe = 0
'                If Sexobe(j) = "M" Then nsbe = 1
'                If Sexobe(j) = "F" Then nsbe = 2
'                If nsbe = 0 Then
'                    X = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'                Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                edabe = Fechap - Fechan
'                If edabe < 1 Then edabe = 1
'                If edabe > Fintab Then
'                    'vgError = 1023
'                    Exit Function
'                End If
'                If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Then icont10 = icont10 + 1
'                If Ncorbe(j) = 20 Or Ncorbe(j) = 21 Then icont20 = icont20 + 1
'                If Ncorbe(j) = 30 Then icont30 = icont30 + 1
'                If Ncorbe(j) = 30 And Coinbe(j) <> "N" Then icont30Inv = icont30Inv + 1
'                If Ncorbe(j) = 35 Then icont35 = icont35 + 1
'                If Ncorbe(j) > 40 And Ncorbe(j) < 50 Then icont40 = icont40 + 1
'                If Ncorbe(j) = 77 Then icont77 = icont77 + 1
'            Next j
'            If (icont10 > 0 Or icont20 > 0) And icont30 > 0 And icont30Inv = 0 Then
'                For j = 1 To Nben
'                    If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or Ncorbe(j) = 20 Or Ncorbe(j) = 21 Then
'                        edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j))) + perdif
'                        If edhm >= L18 Then
'                            Porcbe(j) = 0.42
'                        End If
'                    End If
'                Next j
'            End If
'        End If
'
'        tmtce = (1 + Tasa / 100) ^ (1 / 12)
'
'        ''If Indi = 2 Then
'        ''    ' Mesdif = Mesdif * 12
'        ''    PerDif = Mesdif
'        ''End If
'        ''rmpol = 0
'        ''If Alt = 3 Or (Alt = 4 And pergar > 0) Then Mesgar = pergar
'
'        'If Cober = 8 Or Cober = 9 Or Cober = 10 Or Cober = 11 Or Cober = 12 Then ABV 17-07-2007
'        If Cober = "S" Then
'
'            For ij = 1 To Fintab
'                Flupen(ij) = 0
'            Next ij
'            Mesgar = pergar
'            If Alt = 1 Then Mesgar = 0
'            nmdiga = perdif + Mesgar
'            For j = 1 To Nben
'                pension = Porcbe(j)
'                swg = "N"
'                '                        nrel = 0
'                nibe = 0
'                If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
'                If Coinbe(j) = "N" Then nibe = 2
'                If Coinbe(j) = "P" Then nibe = 2
'                If nibe = 0 Then
'                    X = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'                nsbe = 0
'                If Sexobe(j) = "M" Then nsbe = 1
'                If Sexobe(j) = "F" Then nsbe = 2
'                If nsbe = 0 Then
'                    X = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                    'Renta_Vitalicia = False
'                    Exit Function
'                End If
'
'                'Calculo de la edad de los beneficiarios
'                Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                edabe = Fechap - Fechan
'                If edabe > Fintab Then
'                    X = MsgBox("Error edad del beneficiario es mayor a la tabla de mortalidad.", vbCritical, "Proceso de cálculo Abortado")
'                    ' Renta_Vitalicia = False
'                    Exit Function
'
'                End If
'                If edabe < 1 Then edabe = 1
'                If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
'                    Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
'                    Ncorbe(j) = 41 Or Ncorbe(j) = 46 Or _
'                    Ncorbe(j) = 42 Or Ncorbe(j) = 45 Or _
'                    (Ncorbe(j) >= 30 And Ncorbe(j) < 40) And (Coinbe(j) <> "N" And edabe > L18) Then
'
'                    'PRIMA SOBREVIVENCIA VITALICIA
'                    pension = Porcbe(j)
'                    limite1 = Fintab - edabe - 1
'                    nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                    For i = 0 To limite1
'                        imas1 = i + 1
'                        edalbe = edabe + i
'                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                        If i < nmdiga Then py = 1
'                        If i < perdif Then py = 0
'                        Flupen(imas1) = Flupen(imas1) + py * pension
'                    Next i
'                    'DERECHO A ACRECER
'                    'If Codcbe(j) <> "N" Then
'                    If DerCrecer <> "N" Then
'                        edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
'                        If edhm > L18 Then
'                            nmdif = 0
'                        Else
'                            nmdif = L18 - edhm
'                        End If
'                        Ecadif = edabe + nmdif
'                        limite1 = Fintab - Ecadif - 1
'                        pension = Penben(j) * 0.2
'                        nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                        For i = 0 To limite1
'                            imas1 = nmdif + i + 1
'                            edalbe = Ecadif + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            If (i + nmdif) < nmdiga And nrel = 2 Then py = 1
'                            If (i + nmdif) < perdif Then py = 0
'                            Flupen(imas1) = Flupen(imas1) + py * pension
'                        Next i
'
'                    End If
'
'                Else
'                    If Ncorbe(j) >= 30 And Ncorbe(j) < 40 And ((Coinbe(j) = "N" And edabe <= L18)) Then
'                        'PRIMA DE PENSIONES TEMPORALES
'                        If (edabe > L18 And Coinbe(j) = "N") Then
'                            X = MsgBox("Error edad de hijo mayor a la edad legal", vbCritical)
'                        Else
'                            If edabe < L18 Then
'                                mdif = L18 - edabe
'                                nmdif = mdif - 1
'                                nmax = fgMaximo(nmax, CInt(nmdif)) '*-*nmax = amax.amax0(nmax, CInt(nmdif))
'                                For i = 0 To nmdif
'                                    imas1 = i + 1
'                                    edalbe = edabe + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    If swg = "S" And i < nmdiga Then py = 1
'                                    'En el Siscot2 estaba esta Línea ?????
'                                    If i < nmdiga Then py = 1
'                                    'Fin ???
'                                    If i < perdif Then py = 0
'                                    Flupen(imas1) = Flupen(imas1) + py * pension
'                                Next i
'                            End If
'                        End If
'                        'PRIMA DE HIJOS INVALIDOS
'                        If Coinbe(j) <> "N" Then
'                            kdif = mdif
'                            edbedi = edabe + kdif
'                            limite3 = Fintab - edbedi - 1
'                            pension = Porcbe(j)
'                            nmax = fgMaximo(nmax, CInt(limite3)) '*-*nmax = amax.amax0(nmax, CInt(limite3))
'                            For i = 0 To limite3
'                                edalbe = edbedi + i
'                                nmdifi = i + kdif
'                                imas1 = nmdifi + 1
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                'En el Siscot2 estaba esta Línea ?????
'                                If i < nmdiga Then py = 1
'                                'Fin ???
'                                If nmdifi < perdif Then py = 0
'                                Flupen(imas1) = Flupen(imas1) + py * pension
'                            Next i
'                        End If
'                    Else
'                        X = MsgBox("Error en códificación de parentesco.", vbCritical, "Proceso de Cálculo Abortado")
'                        ' Renta_Vitalicia = False
'                        Exit Function
'
'                    End If
'                End If
'            Next j
'
'
'            ax = 0
'            For LL = 1 To nmax
'                ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
'            Next LL
'
'            If ax <= 0 Then
'                renta = 0
'                ax = 0
'            Else
'                ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
'                renta = SalCta / ax
'                '******
'                renta = CDbl(Format(renta, "#,#0.00"))
'                ax = CDbl(Format(ax, "#,#0.000000"))
'                '*******
'            End If
'            If Indi = "D" Then
'                renta = 0
'            End If
'            If Indi = "D" Then
'                vgPensionCot = renta
'            Else
'                vgPensionCot = (renta / vgFactorAjusteIPC)
'            End If
'            If (Mone = "NS") Then
'                renta = (renta / vgFactorAjusteIPC)
'            End If
'
'            vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
'            vlMtoPenSim = Format(renta, "##0.00")
'            vlPenGar = Format(0, "##0.00") 'sumaqx
'            '----------------------------------------------------------------------
'            If Indi = "D" Then 'calcula diferida
'                add_porc_ben = 0
'                Vpptem = 0
'                Tasa_afp = 0
'                'Prima_unica = 0
'                Rete_sim = 0
'                Prun_sim = 0
'                Sald_sim = 0
'                mesga2 = 0
'                vlMoneda = ""
'                vgPensionCot = 0
'
'                vlMoneda = Mone
'                If Mesdif > 0 Then
'                    gto_sepelio = vlPenGar
'                    If Prc_Tasa_Afp = 0 Then
'                        Vpptem = 0
'                    Else
'                        Vpptem = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ Mesdif)) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                        Vpptem = CDbl(Format(CDbl(Vpptem), "##0.000000"))
'                    End If
'                End If
'
'                'tasa_afp=rentabilidad de la afp
'
'                If vlPriUniSim > 0 Then
'                    If (vlMoneda = "NS") Then
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * (Dif!mto_priunisim * vgFactorAjusteIPC))))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * (vlPriUniSim * vgFactorAjusteIPC))))), "##0.00"))
'                        End If
'                    Else
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * vlPriUniSim)))), "##0.00"))
'                        End If
'                    End If
'                End If
'
'
'                'Saldo necesario AFP
'                Sald_sim = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
'
'                If vlPriUniSim > 0 Then
'                    Prun_sim = CDbl(Format(CDbl(Prima_unica - Sald_sim), "#,#0.00"))
'                    'prun_sim = CDbl(Format(CDbl(Salcta_Sol - Sald_sim), "#,#0.00"))
'                    vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "#,#0.00"))
'
'                    If (vlMoneda = "NS") Then
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / (vlPriUniSim * vgFactorAjusteIPC)), "##0.00"))
'                    Else
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "##0.00"))
'                    End If
'                    vlPensim = vgPensionCot
'                Else
'                    vgPensionCot = 0
'                End If
'                mto_valprepentmp = Vpptem
'                vlMtoPenSim = vlPensim 'hqr para reporte
'                vlMtoPriUniDif = Prun_sim
'                vlMtoCtaIndAfp = Sald_sim
'                vlRtaTmpAFP = Rete_sim
'            End If
'            Dim vlSumaPension As Double
'            'Registrar los valores de la Pensión para cada
'            'Beneficiario para el caso de Sobrevivencia
'            vlSumaPension = 0
'            For vlI = 1 To Nben
'                If (Ncorbe(vlI) <> 0) Then
'                    If (Ncorbe(vlI) = 10) And (Ffam > 1) Then
'                        vlSumaPension = vlSumaPension + CDbl(Format((vgPensionCot * (porfam / 100)), "#0.00"))
'                    Else
'                        vlSumaPension = vlSumaPension + CDbl(Format((vgPensionCot * Porcbe(vlI)), "#0.00"))
'                    End If
'                End If
'            Next vlI
'            vlSumPension = Format(vlSumaPension, "#0.00")
'
'        Else
'
'            'FLUJOS DE RENTAS VITALICIAS DE VEJEZ E INVALIDEZ
'            qx = 0
'            For ij = 1 To Fintab
'                Flupen(ij) = 0
'            Next ij
'            facfam = Ffam
'
'            ''If Alt = 1 Then Mesgar = 0
'            ''If Alt = 3 Or (Alt = 4 And Mesgar > 0) Then Mesgar = Mesgar
'
'            If Alt = "F" Or Alt = "S" Then Mesgar = 0
'            If Alt = "G" Then Mesgar = pergar
'            If Alt = "S" Or Alt = "F" Then Mesgar = 0
'
'            'Definicion del periodo garantizado en 0 para las
'            'alternativas Simple o pensiones con distinto porcentaje legal
'            For j = 1 To Nben
'                'If derpen(j) <> 10 Then
'                pension = Porcbe(j)
'                'CALCULO DE LA PRIMA DEL AFILIADO
'                If (Ncorbe(j) = 0 Or Ncorbe(j) = 99) And j = 1 Then
'                    ni = 0
'                    If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then ni = 1
'                    If Coinbe(j) = "N" Then ni = 2
'                    If Coinbe(j) = "P" Then ni = 3
'                    If ni = 0 Then
'                        X = MsgBox("Error de códificación de tipo de inavlidez", vbCritical, "Proceso de Cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    ns = 0
'                    If Sexobe(j) = "M" Then ns = 1
'                    If Sexobe(j) = "F" Then ns = 2
'                    If ns = 0 Then
'                        X = MsgBox("Error de códificación de sexo.", vbCritical, "Proceso de Cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    'Calculo de edad del causante
'                    Fechan = Nanbe(j) * 12 + Nmnbe(j)
'                    edaca = Fechap - Fechan
'                    If edaca < 780 And ns = 1 And ni = 2 Then cplan = "A"
'                    If edaca < 720 And ns = 2 And ni = 2 Then cplan = "A"
'                    If edaca <= 216 Then edaca = 216
'                    If edaca > Fintab Then
'                        X = MsgBox("Error en edad del beneficiario mayor a tabla de mortalidad", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    sumaqx = 0
'                    limite1 = CInt(Fintab - edaca - 1)
'                    nmax = CInt(limite1)
'                    For i = 0 To limite1
'                        imas1 = i + 1
'                        edacai = edaca + i
'                        If edacai > Fintab Then
'                            X = MsgBox("Edad fuera de Rangos establecidos.", vbCritical, "Proceso de Cálculo Abortado")
'                            'Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                        px = Lx(ns, ni, edacai) / Lx(ns, ni, edaca)
'                        edacas = edacai + 1
'                        qx = ((Lx(ns, ni, edacai) - Lx(ns, ni, edacas))) / Lx(ns, ni, edaca)
'                        Flupen(imas1) = Flupen(imas1) + px * pension
'                        sumaqx = sumaqx + GtoFun * qx / tmtce ^ (i + 0.5)
'                    Next i
'                End If
'
'                If Ncorbe(j) <> 0 And Ncorbe(j) <> 99 Then
'                    'Prima de los Beneficiarios
'                    nibe = 0
'                    If Coinbe(j) = "S" Or Coinbe(j) = "T" Or Coinbe(j) = "I" Then nibe = 1
'                    If Coinbe(j) = "N" Then nibe = 2
'                    If Coinbe(j) = "P" Then nibe = 3
'                    If nibe = 0 Then
'                        X = MsgBox("Error de códificación de tipo de invalidez.", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    nsbe = 0
'                    If Sexobe(j) = "M" Then nsbe = 1
'                    If Sexobe(j) = "F" Then nsbe = 2
'                    If nsbe = 0 Then
'                        X = MsgBox("Error de códificación de sexo.", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    'Calculo de la edad del beneficiario
'                    edabe = Fechap - (Nanbe(j) * 12 + Nmnbe(j))
'                    If edabe < 1 Then edabe = 1
'                    If edabe > Fintab Then
'                        X = MsgBox("Error Edad del beneficario es mayor a la tabla.", vbCritical, "Proceso de cálculo Abortado")
'                        'Renta_Vitalicia = False
'                        Exit Function
'                    End If
'                    If Ncorbe(j) = 10 Or Ncorbe(j) = 11 Or _
'                       Ncorbe(j) = 20 Or Ncorbe(j) = 21 Or _
'                       Ncorbe(j) = 41 Or Ncorbe(j) = 46 Or _
'                       Ncorbe(j) = 42 Or Ncorbe(j) = 45 Or _
'                       ((Ncorbe(j) >= 30 And Ncorbe(j) < 40) And _
'                       (Coinbe(j) <> "N" And edabe > L18)) Then
'                        'FLUJOS DE VIDAS CONJUNTAS VITALICIAS
'                        'Probabilidad del beneficiario solo
'                        limite1 = Fintab - edabe - 1
'                        nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                        For i = 0 To limite1
'                            imas1 = i + 1
'                            edalbe = edabe + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            Flupen(imas1) = Flupen(imas1) + py * pension * facfam
'                        Next i
'                        'Probabilidad conjunta de causante y beneficiario
'                        limite2 = Fintab - edaca - 1
'                        limite = fgMinimo(limite1, CInt(limite2)) '*-* limite = amax.amin0(limite1, CInt(limite2))
'                        nmax = fgMaximo(nmax, CInt(limite)) '*-* nmax = amax.amax0(nmax, CInt(limite))
'                        For i = 0 To limite
'                            imas1 = i + 1
'                            edalca = edaca + i
'                            edalbe = edabe + i
'                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                            px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                            Flupen(imas1) = Flupen(imas1) - (py * px * pension * facfam)
'                        Next i
'
'                        'DERECHO A ACRECER
'                        'If Codcbe(j) <> "N" Then
'                        If DerCrecer <> "N" Then
'                            edhm = (Fechap - (Ijam(j) * 12 + Ijmn(j)))
'                            If edhm > L18 Then
'                                nmdif = 0
'                            Else
'                                nmdif = L18 - edhm
'                            End If
'                            Ecadif = edabe + nmdif
'                            limite1 = Fintab - Ecadif - 1
'                            pension = Porcbe(j) * 0.2
'                            nmax = fgMaximo(nmax, CInt(limite1)) '*-* nmax = amax.amax0(nmax, CInt(limite1))
'                            For i = 0 To limite1
'                                imas1 = nmdif + i + 1
'                                edalbe = Ecadif + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                Flupen(imas1) = Flupen(imas1) + py * pension * facfam
'                            Next i
'                            For i = 0 To limite1
'                                imas1 = nmdif + i + 1
'                                edalbe = Ecadif + i
'                                edalca = (edaca + nmdif) + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                Flupen(imas1) = Flupen(imas1) - (py * px * pension * facfam)
'                            Next i
'
'                        End If
'
'                    Else
'                        If Ncorbe(j) >= 30 And Ncorbe(j) < 40 Then
'                            'Prima Rentas Temporales
'                            If edabe <= L18 Then
'                                mdif = L18 - edabe
'                                nmdif = mdif
'                                'Probabilidad conjunta del causante y beneficiario
'                                limite2 = Fintab - edaca
'                                limite = fgMinimo(nmdif, CInt(limite2)) - 1 '*-* limite = amax.amin0(nmdif, CInt(limite2)) - 1
'                                nmax = fgMaximo(nmax, CInt(limite)) '*-* nmax = amax.amax0(nmax, CInt(limite))
'                                For i = 0 To limite
'                                    imas1 = i + 1
'                                    edalca = edaca + i
'                                    edalbe = edabe + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                    Flupen(imas1) = Flupen(imas1) + ((py * pension) - (py * px * pension)) * facfam
'                                Next i
'                            Else
'                                'I---- ABV 06/01/2004 ---
'                                'Verificar la acción a realizar
'                                'para que continue el proceso, ya que los
'                                'Hijos mayores a la Edad Legal no se deben
'                                'calcular.
'                                'Preguntar a Daniela la implicancia de esto
'                                'x = MsgBox("La Edad del Hijo es mayor a la Edad Legal.", vbCritical, "Proceso de Cálculo Abortado")
'                                'Renta_vitalicia = False
'                                'Exit Function
'                                'F---- ABV 06/01/2004 ---
'                            End If
'                            'Prima del Hijo Invalido
'                            If Coinbe(j) <> "N" Then
'                                'Probabilidad conjunta del causante y beneficiario
'                                edbedi = edabe + nmdif
'                                limite3 = Fintab - edbedi - 1
'                                limite4 = Fintab - (edaca + nmdif) - 1
'                                nmax = fgMaximo(nmax, CInt(limite3)) '*-* nmax = amax.amax0(nmax, CInt(limite3))
'                                For i = 0 To limite3
'                                    nmdifi = nmdif + i
'                                    imas1 = nmdifi + 1
'                                    edalca = edaca + nmdif + i
'                                    edalbe = edbedi + i
'                                    edalca = fgMinimo(edalca, CInt(Fintab)) '*-* edalca = amax.amin0(edalca, CInt(Fintab))
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                    Flupen(imas1) = Flupen(imas1) + (py - (py * px)) * pension * facfam
'                                Next i
'                            End If
'                        Else
'                            X = MsgBox("Error en codificación del parentesco.", vbCritical, "Proceso de cálculo Abortado")
'                            'Renta_Vitalicia = False
'                            Exit Function
'                        End If
'                    End If
'                End If
'                'End If
'            Next j
'
'            'Calculo de tarifa y Pension
'            ax = 0
'            flumax = 0
'            If Alt = "F" Or Alt = "S" Then
'                Mesgar = 0
'                nmdiga = perdif + Mesgar
'                ax = 0
'                For LL = 1 To nmax
'                    flumax = Flupen(LL)
'                    If LL <= perdif Then flumax = 0
'                    ax = ax + flumax / tmtce ^ (LL - 1)
'                Next LL
'            Else
'                Mesgar = pergar
'                nmdiga = perdif + Mesgar
'                ax = 0
'                For LL = 1 To nmdiga
'                    flumax = fgMaximo(1, Flupen(LL)) '*-* flumax = amax.amax1(1, Flupen(LL))
'                    If LL <= perdif Then flumax = 0
'                    ax = ax + flumax / tmtce ^ (LL - 1)
'                Next LL
'                For LL = nmdiga + 1 To nmax
'                    ax = ax + Flupen(LL) / tmtce ^ (LL - 1)
'                Next LL
'            End If
'
'            If Indi = "I" Then
'                If ax <= 0 Then
'                    renta = 0
'                    ax = 0
'                Else
'                    sumaqx = CDbl(Format(sumaqx, "#,#0.000000"))
'                    ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
'                    renta = CDbl(Format(((SalCta - sumaqx) / ax), "#,#0.00"))
'                End If
'            Else
'                ax = CDbl(Format(ax, "#,#0.000000")) + mescosto
'                renta = 0
'            End If
'            If (Mone = "NS") Then
'                renta = (renta * vgFactorAjusteIPC)
'            End If
'            'HQR 24/05/2004
'            vlPriUniSim = Format(ax, "##0.00") 'HQR REVISAR
'            vlMtoPenSim = Format(renta, "##0.00")
'            vlPenGar = Format(sumaqx, "##0.00")
'            'FIN HQR 24/05/2004
'
'            If Indi = "D" Then
'                'aca se iba a la función que calcula diferida
'                add_porc_ben = 0
'                Vpptem = 0
'                Tasa_afp = 0
'                ' Prima_unica = 0
'                Rete_sim = 0
'                Prun_sim = 0
'                Sald_sim = 0
'                mesga2 = 0
'                vlMoneda = ""
'                vgPensionCot = 0
'
'                vlMoneda = Mone
'                If Mesdif > 0 Then
'                    gto_sepelio = vlPenGar
'                    If Prc_Tasa_Afp = 0 Then
'                        Vpptem = 0
'                    Else
'                        Vpptem = ((1 - 1 / ((1 + Prc_Tasa_Afp) ^ Mesdif)) / Prc_Tasa_Afp) * (1 + Prc_Tasa_Afp) * 12
'                        Vpptem = CDbl(Format(CDbl(Vpptem), "##0.000000"))
'                    End If
'                End If
'
'                'tasa_afp=rentabilidad de la afp
'
'                If vlPriUniSim > 0 Then
'                    If (vlMoneda = "NS") Then
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * (Dif!mto_priunisim * vgFactorAjusteIPC))))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * (vlPriUniSim * vgFactorAjusteIPC))))), "##0.00"))
'                        End If
'                    Else
'                        If Vpptem = 0 Or Prc_Pension_Afp = 0 Then
'                            Rete_sim = 0
'                        Else
'                            'rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
'                            Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + Prc_Pension_Afp * vlPriUniSim)))), "##0.00"))
'                        End If
'                    End If
'                End If
'
'                'Saldo necesario AFP
'                Sald_sim = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
'
'                If vlPriUniSim > 0 Then
'                    Prun_sim = CDbl(Format(CDbl(Prima_unica - Sald_sim), "#,#0.00"))
'                    'prun_sim = CDbl(Format(CDbl(Salcta_Sol - Sald_sim), "#,#0.00"))
'                    vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "#,#0.00"))
'
'                    If (vlMoneda = "NS") Then
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / (vlPriUniSim * vgFactorAjusteIPC)), "##0.00"))
'                    Else
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / vlPriUniSim), "##0.00"))
'                    End If
'                    vlPensim = vgPensionCot
'                Else
'                    vgPensionCot = 0
'                End If
'                mto_valprepentmp = Vpptem
'                vlMtoPenSim = vlPensim 'hqr para reporte
'                vlMtoPriUniDif = Prun_sim
'                vlMtoCtaIndAfp = Sald_sim
'                vlRtaTmpAFP = Rete_sim
'            End If
'        End If
'
'        'llenar el datatable
'        istPolizas.mto_valprepentmp = mto_valprepentmp 'dtDataRow("MTO_VALPREPENTMP") = mto_valprepentmp
'        istPolizas.Mto_Pension = vlMtoPenSim 'dtDataRow("MTO_PENSION") = vlMtoPenSim
'        istPolizas.Mto_PriUniDif = vlMtoPriUniDif 'dtDataRow("MTO_PRIUNIDIF") = vlMtoPriUniDif
'        istPolizas.Mto_CtaIndAFP = vlMtoCtaIndAfp 'dtDataRow("MTO_CTAINDAFP") = vlMtoCtaIndAfp
'        istPolizas.Mto_RentaTMPAFP = vlRtaTmpAFP 'dtDataRow("MTO_RENTATMPAFP") = vlRtaTmpAFP
'        istPolizas.Mto_PriUniSim = vlPriUniSim 'dtDataRow("MTO_PRIUNISIM") = vlPriUniSim
'        istPolizas.Mto_RMGtoSepRV = vlPenGar 'dtDataRow("MTO_PENSIONGAR") = vlPenGar
'        If (Mesgar > 0) Then
'            istPolizas.Mto_PensionGar = vlMtoPenSim
'        Else
'            istPolizas.Mto_PensionGar = 0
'        End If
'        'dtDataRow("NUM_CORRELATIVO") = vlCorrCot
'        mto_valprepentmp = 0
'        vlMtoPenSim = 0
'        vlMtoPriUniDif = 0
'        vlMtoCtaIndAfp = 0
'        vlRtaTmpAFP = 0
'        vlPriUniSim = 0
'        vlPenGar = 0
'        vlCorrCot = 0
'        'Renta_Vitalicia.Rows.Add (dtDataRow)
'
'    Next vlNumero
'
'    fgCalcularRentaVitalicia = True

End Function

Function Calcula_Diferida(Cotizacion As String, Cod_AFP As String, Comision As Double, indicador As Long, iRentabilidad As Double)
'Dim Add_porc_be As Double, Vpptem As Double, Tasa_afp As Double, Prima_unica As Double
'Dim Rete_sim As Double, Prun_sim As Double, Sald_sim As Double, mesga2 As Double
'Dim vlNumCoti As String
'Dim vlCorrCot As Integer
'
'Dim Dif As ADODB.Recordset
''I--- ABV 04/02/2006 ---
'Dim vlMoneda As String
''F--- ABV 04/02/2006 ---
'
'    add_porc_ben = 0
'    Vpptem = 0
'    Tasa_afp = 0
'    Prima_unica = 0
'    Rete_sim = 0
'    Prun_sim = 0
'    Sald_sim = 0
'    mesga2 = 0
'    vlMoneda = ""
'
'    vgPensionCot = 0
'
'    Query = "SELECT "
'    Query = Query & " num_cot,num_pro,num_mesdif,mto_priuni,mto_pengar, COD_MONEDA, "
'    Query = Query & " num_correlcot,mto_priunisim,prc_rentatmp "
'    Query = Query & " FROM tmae_propuesta "
'    Query = Query & " WHERE "
'    Query = Query & " num_cot = '" & Cotizacion & "' and "
'    Query = Query & " num_pro = " & indicador & ""
'    Set Dif = vgConectarBD.Execute(Query)
'    If Not Dif.EOF Then
'
'        vlMoneda = Trim(Dif!Cod_Moneda)
'
'        'Dif.MoveFirst
'        'Do While Not Dif.EOF
'            If Dif!Num_MesDif > 0 Then
'                'Select Case Dif!Moneda
'                '    Case "US"
'                 '       Prima_unica = Dif!prima_us
'                 '   Case "NS"
'                        'I---- ABV 10/02/2004 ---
'                        'Prima_unica = Dif!mto_priunius
'                        Prima_unica = Dif!mto_priuni
'                        'F---- ABV 10/02/2004 ---
'                        gto_sepelio = Dif!mto_pengar
'                'End Select
'
'                'I---- ABV 10/02/2004 ---
'                'Paso1 = "SELECT * FROM tpar_tabcod WHERE cod_tabla = 'AF' and "
'                'Paso1 = Paso1 & "trim(cod_elemento) = '" & Cod_afp & "'"
'                'Set Q = vgConectarBD.Execute(Paso1)
'                'If Not (Q.EOF) Then
'                '    Tasa_afp = Q!mto_elemento / 100
'                    Tasa_afp = iRentabilidad / 100
'                'End If
'                'Q.Close
'                'F---- ABV 10/02/2004 ---
'
'                'I --- Daniela  13/10/2004
''                'Determina la Suma Total de Porcentajes del Grupo Familiar
''                Paso1 = "SELECT sum(prc_legal) as porcen FROM tmae_benpro WHERE "
''                Paso1 = Paso1 & "num_cot = '" & Cotizacion & "'"
''                Set Q = vgConectarBD.Execute(Paso1)
''                If Not Q.EOF Then
''                    If Not IsNull(Q!porcen) Then
''                        Add_porc_be = Q!porcen
''                    Else
''                        Add_porc_be = 0
''                    End If
''                End If
''                Q.Close
'                'F --- Daniela  13/10/2004
'
''                'I  -- Daniela 13/10/2004
''                If cober = "S" Then
''                    Add_porc_be = 100
''                End If
''                'F  -- Daniela 13/10/2004
'
'                Vpptem = ((1 - 1 / ((1 + Tasa_afp) ^ Dif!Num_MesDif)) / Tasa_afp) * (1 + Tasa_afp) * 12
'                'I---- ABV 13/11/2003 ---
'                'Vpptem = Format(CLng(Vpptem), "##0.0000")
'                Vpptem = CDbl(Format(CDbl(Vpptem), "##0.000000"))
'                'F---- ABV 13/11/2003 ---
''''                Dif.Edit
''''                'I---- ABV 13/11/2003 ---
''''                'Dif!fac_afp = Format(CLng(Vpptem), "##0.000")
''''                Dif!fac_afp = Vpptem
''''                'I---- ABV 13/11/2003 ---
''''                Dif.Update
'
'                vlCorrCot = Dif!num_correlcot
'
'                vlNumCoti = Mid(Cotizacion, 1, 13) & Format(vlCorrCot, "00") & Mid(Cotizacion, 16, 15)
'
'                'I---- ABV 09/11/2004 ---
'                'vlSql = ""
'                'vlSql = "Update tmae_propuesta set "
'                'vlSql = vlSql & "mto_valprepentmp = " & Str(Vpptem) & " "
'                'vlSql = vlSql & "WHERE "
'                'vlSql = vlSql & "num_cot = '" & Cotizacion & "' and "
'                'vlSql = vlSql & "num_pro = " & indicador & " "
'                'vgConectarBD.Execute (vlSql)
'                '
'                'vlSql = "Update tmae_cotizacion set "
'                'vlSql = vlSql & "mto_valprepentmp = " & Str(Vpptem) & " "
'                'vlSql = vlSql & "WHERE "
'                'vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                'vgConectarBD.Execute (vlSql)
'                'F---- ABV 09/11/2004 ---
'
'                Add_porc_be = Add_porc_be / 100
'
'                'I---- ABV 10/02/2004 ---
'                'Query = "SELECT * FROM tmae_propuesta WHERE "
'                'Query = Query & " num_cot = '" & Cotizacion & "' and "
'                'Query = Query & " num_pro = " & indicador & ""
'                'Set Dif = vgConectarBD.Execute(Query)
'                'F---- ABV 10/02/2004 ---
'
'                If Dif!Mto_PriUniSim > 0 Then
'
''I--- PAC 04/02/2006 ---
'''I--- ABV 04/01/2006 ---
'''                    Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
''                    Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + (Dif!prc_rentatmp / 100) * Dif!mto_priunisim)))), "##0.00"))
'                    If (vlMoneda = "NS") Then
'                        Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + (Dif!Prc_RentaTMP / 100) * (Dif!Mto_PriUniSim * vgFactorAjusteIPC))))), "##0.00"))
'                    Else
'                        Rete_sim = CDbl(Format(CDbl(((Prima_unica - gto_sepelio) * (1 / (Vpptem + (Dif!Prc_RentaTMP / 100) * Dif!Mto_PriUniSim)))), "##0.00"))
'                    End If
'''F--- ABV 04/01/2006 ---
''I--- PAC 04/02/2006 ---
'                End If
'
'''I--- ABV 04/02/2006 ---
''                If (vlMoneda = "NS") Then
''                    Rete_sim = (Rete_sim / vgFactorAjusteIPC)
''                End If
'''F--- ABV 04/02/2006 ---
'
'                '**********
'                Sald_sim = CDbl(Format(CDbl(Vpptem * Rete_sim), "#,#0.00"))
'                '**********
'                If (Dif!Mto_PriUniSim > 0) Then
'                    '******
'                    Prun_sim = CDbl(Format(CDbl(Prima_unica - Sald_sim), "#,#0.00"))
'                    '******
'                    'Prun_sim = CDbl(Prima_unica - Sald_sim)
'                End If
'
''''                Dif.Edit
''''                Dif!pensim = 0
''''                Dif.Update
'
''I--- PAC 04/02/2006 ---
'                If (Dif!Mto_PriUniSim > 0) Then
'                    If (vlMoneda = "NS") Then
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / (Dif!Mto_PriUniSim * vgFactorAjusteIPC)), "##0.00"))
'                    Else
'                        vgPensionCot = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / Dif!Mto_PriUniSim), "##0.00"))
'                    End If
'                Else
'                    vgPensionCot = 0
'                End If
''F--- PAC 04/02/2006 ---
'
'                'I---- ABV 09/11/2004 ---
'                'vlSql = "Update tmae_propuesta set "
'                'vlSql = vlSql & "mto_pensim = 0 "
'                'vlSql = vlSql & "WHERE "
'                'vlSql = vlSql & "num_cot = '" & Cotizacion & "' and "
'                'vlSql = vlSql & "num_pro = " & indicador & " "
'                'vgConectarBD.Execute (vlSql)
'                '
'                'vlSql = "Update tmae_cotizacion set "
'                'vlSql = vlSql & "mto_pensim = 0 "
'                'vlSql = vlSql & "WHERE "
'                'vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                'vgConectarBD.Execute (vlSql)
'                'F---- ABV 09/11/2004 ---
'
'                vlPensim = 0
'
'                If (Dif!Mto_PriUniSim > 0) Then
''''                    Dif.Edit
''''                    '***********
''''                    Dif!pensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / Dif!pufsim), "#,#0.00"))
'''''                    Dif!pensim = CDbl(Prun_sim / Dif!pufsim)
''''                    '*****************
''''                    Dif.Update
'
''I--- PAC 04/02/2006 ---
''                    vlPensim = CDbl(Format(CDbl((Prun_sim - gto_sepelio) / Dif!mto_priunisim), "#,#0.00"))
'                    vlPensim = vgPensionCot
''F--- PAC 04/02/2006 ---
'
'                    'I---- ABV 09/11/2004 ---
'                    'vlSql = "Update tmae_propuesta set "
'                    'vlSql = vlSql & "mto_pensim = " & Str(vlPensim) & " "
'                    'vlSql = vlSql & "WHERE "
'                    'vlSql = vlSql & "num_cot = '" & Cotizacion & "' and "
'                    'vlSql = vlSql & "num_pro = " & indicador & " "
'                    'vgConectarBD.Execute (vlSql)
'                    '
'                    'vlSql = "Update tmae_cotizacion set "
'                    'vlSql = vlSql & "mto_pensim = " & Str(vlPensim) & " "
'                    'vlSql = vlSql & "WHERE "
'                    'vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                    'vgConectarBD.Execute (vlSql)
'                    'F---- ABV 09/11/2004 ---
'                End If
'
'''''                Dif.Edit
'''''                Dif!pu_dif_sim = Prun_sim
'''''                Dif!cta_rt_sim = Sald_sim
'''''                Dif!rt_afp_sim = Rete_sim
'''''                Dif!factor_afp = 0
'''''                Dif!tasa_rt = 0
'''''                Dif.Update
'
'                vlSql = "Update tmae_propuesta set "
'                'I---- ABV 09/11/2004 ---
'                vlSql = vlSql & "mto_valprepentmp = " & Str(Vpptem) & ", "
'                vlSql = vlSql & "mto_pensim = " & Str(vlPensim) & ", "
'                'I---- ABV 09/11/2004 ---
'                vlSql = vlSql & "mto_priunidif = " & Str(Prun_sim) & ", "
'                vlSql = vlSql & "mto_ctaindafp = " & Str(Sald_sim) & ", "
'                vlSql = vlSql & "mto_rentatmpafp = " & Str(Rete_sim) & " "
'                vlSql = vlSql & "WHERE "
'                vlSql = vlSql & "num_cot = '" & Cotizacion & "' and "
'                vlSql = vlSql & "num_pro = " & indicador & " "
'                vgConectarBD.Execute (vlSql)
'
'                If (vlCorrCot <> 0) Then
'                    vlSql = "Update tmae_cotizacion set "
'                    'I---- ABV 09/11/2004 ---
'                    vlSql = vlSql & "mto_valprepentmp = " & Str(Vpptem) & ", "
'                    vlSql = vlSql & "mto_pensim = " & Str(vlPensim) & ", "
'                    'I---- ABV 09/11/2004 ---
'                    vlSql = vlSql & "mto_priunidif = " & Str(Prun_sim) & ","
'                    vlSql = vlSql & "mto_ctaindafp = " & Str(Sald_sim) & ","
'                    vlSql = vlSql & "mto_rentatmpafp = " & Str(Rete_sim) & " "
'                    vlSql = vlSql & "WHERE "
'                    vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                    vgConectarBD.Execute (vlSql)
'                End If
'            End If
'        '    Dif.MoveNext
'        'Loop
'    End If
'    Dif.Close
End Function

Function Tarifa_Todo(Coti As String, iOperacion As String, Cual_Form As String) As Boolean
'ReDim Cp(1 To 112) As Double, impres(9, 110) As Double
'ReDim Penben(1 To 20) As Double, prodin(112) As Double, Flupen(1 To 112) As Double, flucm(1 To 112) As Double, exced(1 To 112) As Double, porcbe(1 To 20) As Double
'ReDim coinbe(1 To 20), codcbe(1 To 20), sexobe(1 To 20), cob(1 To 5), alt1(1 To 3), tip(1 To 2) As String
'Dim ncorbe(1 To 20) As Integer, nanbe(1 To 20) As Integer, nmnbe(1 To 20) As Integer, i As Integer
'Dim ijam(1 To 20) As Integer, ijmn(1 To 20) As Integer
'Dim npolbe(1 To 20) As String, porcbe_ori(1 To 20) As Double
'Dim isuc(1 To 20) As Long
'Dim mesdif As Long, Fechan As Long, Fechap As Long, edbedi As Long, mdif As Long
'Dim mesgar As Long
'Dim large As Integer
'Dim edaca As Long, edalca As Long, edacai As Long, edacas As Long, edabe As Long, edalbe As Long
'Dim Fasolp As Long, Fmsolp As Long, Fdsolp As Long, pergar As Long, codsuc As String, L18 As Long, numrec As Long, numrep As Long
'Dim nmax As Integer, nap As Integer, nmp As Integer, j As Integer
'Dim nrel As Long, nmdif As Long, nben As Long, numbep As Long, ni As Long, ns As Long, nibe As Long, nsbe As Long, limite As Long
'Dim limite1 As Long, limite2 As Long, limite3 As Long, limite4 As Long, imas1 As Long, kdif As Long, nt  As Long
'Dim rmpol As Double, gtofun As Double, px As Double, py As Double, Qx As Double, relres As Double
'Dim salcta As Double, ffam As Double, comisi As Double, facdec As Double, tasac As Double, timp As Double, tm As Double, tmm As Double
'Dim tm3 As Double, sumapx As Double, sumaqx As Double, actual As Double, actua1 As Double, penbase As Double, tce As Double
'Dim vpte As Double, difres As Double, difre1 As Double, tir As Double, tinc As Double
'Dim cober, alt, indi, cplan As String
'Dim Tasa As Double
'Dim tastce As Double
'Dim tirvta As Double
'Dim tvmax As Double
'Dim vppen As Double, vpcm As Double, penanu As Double, reserva As Double, gastos As Double, rdeuda As Double
'Dim resfin As Double, rend As Double, varrm As Double, resant As Double, flupag As Double, gto As Double
'Dim sumaex As Double, sumaex1 As Double, tirmax As Double, penmin As Double, penmax As Double
'Dim gto_supervivencia As Double
'Dim Sql, Numero As String
'Dim Linea1 As String
'Dim Inserta As String
'Dim Var, Nombre, vgs_coti As String
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
'Dim salcta_eva As Double
'Dim Vpptem As Integer
'
'Dim Orden(1 To 20) As Integer
'
'Dim vlFechaNacCausante As String
'Dim vlSexoCausante     As String
'
''---------------------------------------------------------------------------
''Ultima Modificación realizada el 18/12/2005
''Agregar Tablas de Mortalidad con lectura desde la BD y no desde un archivo
''Además, manejar dichas tablas por Fecha de Vigencia, para que opere la que
''corresponda a la Fecha de Cotización
''La Tabla de Mortalidad en esta función es ANUAL
''---------------------------------------------------------------------------
'
'    Screen.MousePointer = 11
'    Tarifa_Todo = False
'    cuenta = 0
'    L18 = 18
'    numrec = 0
'    numrep = 0
'    numrec = -1
'    '-------------------------------------------------
'    'Leer Tabla de Mortalidad
'    '-------------------------------------------------
'    Navig = CInt(Mid(vgFechaCotizacion, 1, 4))
'    Nmvig = CInt(Mid(vgFechaCotizacion, 5, 2))
'    Ndvig = CInt(Mid(vgFechaCotizacion, 7, 2))
'    nap = CInt(Mid(vgFechaCotizacion, 1, 4))
'    nmp = CInt(Mid(vgFechaCotizacion, 5, 2))
'    Ndp = CInt(Mid(vgFechaCotizacion, 7, 2))
'    If (fgBuscarMortalidad(Navig, Nmvig, Ndvig, nap, nmp, Ndp, vlSexoCausante, vlFechaNacCausante) = False) Then
'        'Se produjeron errores en la obtención de una de las Tablas de Mortalidad
'        Exit Function
'    End If
'    'Permite determinar el Número de Propuestas existentes
'    Sql = "SELECT count(num_cot) as numero from tmae_propuesta "
'    Sql = Sql & "where num_cot = '" & Coti & "'"
'    Set vgRs = vgConectarBD.Execute(Sql)
'    If Not (vgRs.EOF) Then
'        qwer = vgRs!Numero
'    Else
'        MsgBox "Proceso de Cálculo Abortado por inexistencia de Datos.", vbCritical, "Proceso de cálculo Abortado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'    vgRs.Close
'
'    If qwer = 0 Then
'        vlAumento = 0
'    Else
'        vlAumento = 100 / qwer   'PORQUE NO TOMA EL VALOR REAL Y DICE DIVISION POR CERO ?????
'    End If
'
'    Frm_Progress.Show
'    Frm_Progress.Caption = "Progreso del cálulo"
'    Frm_Progress.Refresh
'    Frm_Progress.ProgressBar1.Value = 0
'    Frm_Progress.lbl_progress = "Realizando Cálculos de Tarifa..."
'    Frm_Progress.Refresh
'    Frm_Progress.ProgressBar1.Visible = True
'    Frm_Progress.Refresh
'
'    Sql = "select num_cot from tmae_propuesta "
'    Sql = Sql & " where num_cot = '" & Coti & "'"
'    Set Tb_Difpol = vgConectarBD.Execute(Sql)
'    If (Tb_Difpol.EOF) Then
'        MsgBox "No existen Propuestas de cotizaciones en la Base de Datos.", vbCritical, "Proceso de cálculo Abortado"
'        Tarifa_Todo = False
'        Exit Function
'    End If
'
'    'Borrar el contenido de la Tabla que registra los valores de la Evaluación
'    vlSql = "delete from tmae_evapro where num_cot = '" & Coti & "'"
'    vgConectarBD.Execute (vlSql)
'
'    'Borrar el contenido de la Tabla de Evaluaciones de Cotizaciones
'    vlSql = "delete from tmae_evacot where "
'    vlSql = vlSql & "mid(num_cot,1,13) = '" & Mid(Coti, 1, 13) & "' and "
'    vlSql = vlSql & "mid(num_cot,16,15) = '" & Mid(Coti, 16, 15) & "' "
'    vgConectarBD.Execute (vlSql)
'
'    For indicador = 1 To qwer
'
'        If Frm_Progress.ProgressBar1.Value + vlAumento < 100 Then
'            Frm_Progress.ProgressBar1.Value = Frm_Progress.ProgressBar1.Value + vlAumento
'            Frm_Progress.Refresh
'        End If
'
'        Sql = "select "
'        Sql = Sql & " num_cot,num_ben,cod_plan,cod_modalidad,cod_alternativa,"
'        Sql = Sql & " num_mesgar,cod_Moneda,fec_ingcot,fec_ingcot,mto_bonact,"
'        Sql = Sql & " mto_bonactus,mto_ctaind,mto_priuni,mto_priunius,mto_facpenella,"
'        Sql = Sql & " num_mesdif,fec_dev,fec_dev,mto_gassep,prc_rentaafpori,"
'        Sql = Sql & " prc_rentatmp,num_correlcot "
'        Sql = Sql & " from tmae_propuesta "
'        Sql = Sql & " where "
'        Sql = Sql & " num_cot = '" & Coti & "' and "
'        Sql = Sql & " num_pro = " & indicador
'        Set CtaDifpol = vgConectarBD.Execute(Sql)
'        If Not CtaDifpol.EOF Then
'            'If ctadifpol.RecordCount > 0 Then
'                cuenta = 1
'                vgd_tasa_vta = 0
'                npolca = CtaDifpol!Num_Cot
'                nben = CtaDifpol!num_ben
'                '************27/03/2001
'                If vgTipoPension = "S" Then
'                    nben = nben - 1
'                End If
'                '*********************
'                cober = CtaDifpol!cod_plan
'                indi = CtaDifpol!Cod_Modalidad
'                alt = CtaDifpol!cod_alternativa
'                mesgar = CtaDifpol!Num_MesGar
'                mone = UCase(CtaDifpol!Cod_Moneda)
'                nap = Year(CtaDifpol!fec_ingcot)
'                nmp = Month(CtaDifpol!fec_ingcot)
'                bono_sol1 = CtaDifpol!mto_bonact
'                bono_us1 = CtaDifpol!mto_bonactus
'                ctaind = CtaDifpol!Mto_CtaInd
'                salcta_sol = CtaDifpol!mto_priuni
'                salcta_us = CtaDifpol!mto_priunius
'                codsuc = vgSucursal
'                ffam = CtaDifpol!Mto_FacPenElla
'                mesdif = CtaDifpol!Num_MesDif
'                Fasolp = Year(CtaDifpol!Fec_Dev)
'                Fmsolp = Month(CtaDifpol!Fec_Dev)
'                Fdsolp = Day(CtaDifpol!Fec_Dev)
'                gtofun = IIf(IsNull(CtaDifpol!mto_gassep), 0, CtaDifpol!mto_gassep)
'                prc_Tasa_afp = CtaDifpol!prc_rentaafpori / 100
'                prc_Pension_afp = CtaDifpol!Prc_RentaTMP / 100
'                vlCorrCot = CtaDifpol!num_correlcot
'                If indi = "RVFA" Then indi = "I"
'                If indi = "RTRV" Then indi = "D"
'                If mone = "NS" Then
'                    salcta = salcta_sol
'                Else
'                    salcta = salcta_us
'                End If
'                If indicador = 1 Then vgs_coti = CtaDifpol!Num_Cot
'                totpor = 0
'                Sql = "select "
'                Sql = Sql & "num_orden,cod_par,prc_legal,fec_nacben,cod_sexo, "
'                Sql = Sql & "cod_sitinv,cod_dercre,fec_nachm,num_cot "
'                Sql = Sql & ",prc_legalaux "
'                Sql = Sql & "from tmae_benpro "
'                Sql = Sql & "where "
'                Sql = Sql & "num_cot = '" & vgs_coti & "' "
'                If (vgTipoPension = "S") Then
'                    Sql = Sql & " and num_orden <> 1 "  'el orden empieza de 1 no de 0
'                End If
'                Sql = Sql & "order by cod_par"
'                Set Tb_Difben = vgConectarBD.Execute(Sql)
'                If (Tb_Difben.EOF) Then
'                    MsgBox "Falta de antecedentes de Beneficiarios en Propuestas de Cotizaciones para realización de cálculo.", vbCritical, "Proceso de Cálculo Abortado"
'                    Tarifa_Todo = False
'                    Exit Function
'                End If
'                sumaporcsob = 0
'                For i = 1 To nben
'                    Orden(i) = Tb_Difben!Num_Orden
'                    ncorbe(i) = Tb_Difben!Cod_Par
'                    If (ncorbe(i) = 99) Or (ncorbe(i) = 0) Then
'                        vlFechaNacCausante = Tb_Difben!Fec_NacBen
'                        vlSexoCausante = Trim(Tb_Difben!Cod_Sexo)
'                    End If
'                    porcbe(i) = Tb_Difben!prc_legal
'                    porcbe_ori(i) = Tb_Difben!prc_legalaux 'porcbe(i)
'                    nanbe(i) = Year(Tb_Difben!Fec_NacBen)    'aa_nac
'                    nmnbe(i) = Month(Tb_Difben!Fec_NacBen)   'mm_nac
'                    sexobe(i) = Tb_Difben!Cod_Sexo
'                    coinbe(i) = Tb_Difben!Cod_SitInv
'
'                    codcbe(i) = Tb_Difben!Cod_DerCre
'
'                    If Not IsNull(Tb_Difben!Fec_NacHM) Then
'                        ijam(i) = Year(Tb_Difben!Fec_NacHM)   'aa_hijom
'                        ijmn(i) = Month(Tb_Difben!Fec_NacHM)     'mm_hijom
'                    Else
'                        ijam(i) = "0000" ' Year(tb_difben!fec_nachm)   'aa_hijom
'                        ijmn(i) = "00" 'Month(tb_difben!fec_nachm)     'mm_hijom
'                    End If
'                    npolbe(i) = Tb_Difben!Num_Cot
'                    isuc(i) = 0                             'tb_difben!Sucursal
'                    porcbe(i) = porcbe(i) / 100
'                    porcbe_ori(i) = porcbe_ori(i) / 100 'porcbe(i)
'                    If cober = "S" And (ncorbe(i) <> 0 Or ncorbe(i) <> 99) Then sumaporcsob = sumaporcsob + porcbe(i)
'                    totpo = totpo + porcbe(i)
'                    Tb_Difben.MoveNext
'                Next i
'                Tb_Difben.Close
'                If cober = "S" And sumaporcsob > 1 And ffam > 0 Then
'                    x = MsgBox("La suma de los porcentajes de pensión corregidos por factor Pensando en Ella es mayor al 100%.", vbCritical, "Proceso de cálculo Abortado")
'                    Tarifa_Todo = False
'                    Exit Function
'                End If
'                    Sql = "select * from tval_gasto where "
'                    Sql = Sql & "cod_moneda = '" & mone & "' "
'                    Set Tb_EvaparUS = vgConectarBD.Execute(Sql)
'                    If (Tb_EvaparUS.EOF) Then
'                        MsgBox "Inexistencia de datos de Parámetros de Evaluación en Dólares.", vbCritical, "Proceso de Cálculo Abortado"
'                        Tarifa_Todo = False
'                        Exit Function
'                    End If
'
'                    comisi = Format(Tb_EvaparUS!prc_gasint, "#0.00")               'comisi
'                    tasac = Format(Tb_EvaparUS!prc_ctocap, "#0.00")                'tasac
'                    gastos = Format(Tb_EvaparUS!mto_gasadm, " #0.00")               'gastos
'                    gto_supervivencia = Format(Tb_EvaparUS!mto_gasconsup, "#0.00") 'gastos_sup
'                    rdeuda = Format(Tb_EvaparUS!prc_endeuda, "#0.00")             'rdeuda
'                    timp = Format(Tb_EvaparUS!prc_impuesto, "#0.00")               'timp
'                    penmin = Format(Tb_EvaparUS!prc_tasaminvta, "#0.00")           'tasa_min
'                    penmax = Format(Tb_EvaparUS!prc_tasamaxvta, "#0.00")           'tasa_max
'                    gasemi = Format(Tb_EvaparUS!mto_gasemi, "#0.00")               'gastos_emi
'                    facdec = Format(Tb_EvaparUS!prc_gasconsvs, "#0.00")            'gastos_eva
'                    If Not IsNull(Tb_EvaparUS!prc_permax) Then
'                        PerMax = Tb_EvaparUS!prc_permax            'Porc. Pérdida Máxima
'                    Else
'                        PerMax = 0
'                    End If
'
'
'
'                    Tb_EvaparUS.Close
'
'                    Sql = "select "
'                    Sql = Sql & "num_anno,prc_tasamer,prc_cpk "
'                    Sql = Sql & " from tval_calce where "
'                    Sql = Sql & "cod_moneda = '" & mone & "' "
'                    Sql = Sql & "order by num_anno "
'                    Set Tb_EvapasUS = vgConectarBD.Execute(Sql)
'                    If Not (Tb_EvapasUS.EOF) Then
'                        vlI = 1
'                        tm = Tb_EvapasUS!prc_tasamer
'
'                        While Not (Tb_EvapasUS.EOF)
'                            vlI = Tb_EvapasUS!num_anno
'                            Cp(vlI) = Tb_EvapasUS!prc_cpk
'                            Tb_EvapasUS.MoveNext
'                        Wend
'                    Else
'                        MsgBox "Inexistencia de Datos de Parámetros de Tabla de Calce en Dólares.", vbCritical, "Proceso de cálculo Abortado"
'                        Tarifa_Todo = False
'                        Exit Function
'                    End If
'                    Tb_EvapasUS.Close
'                    Sql = "select "
'                    Sql = Sql & "num_anno,prc_tasatip "
'                    Sql = Sql & "from tval_rentabilidad where "
'                    Sql = Sql & "cod_moneda = '" & mone & "' "
'                    Sql = Sql & "order by num_anno "
'                    Set Tb_EvaproUS = vgConectarBD.Execute(Sql)
'                    If (Tb_EvaproUS.EOF) Then
'                        MsgBox "Inexistencia de datos de Parámetros de Productos de Inversiones en Dólares.", vbCritical, "Proceso de Cálculo Abortado"
'                        Tarifa_Todo = False
'                        Exit Function
'                    End If
'
'                    While Not (Tb_EvaproUS.EOF)
'                        vlI = Tb_EvaproUS!num_anno
'                        prodin(vlI) = Format(Tb_EvaproUS!prc_tasatip, "#0.00") / 100
'                        Tb_EvaproUS.MoveNext
'                    Wend
'                    Tb_EvaproUS.Close
'                comisi = comisi / 100
'                tasac = tasac / 100
'                timp = timp / 100
'                tmm = (1 + tm / 100)
'                tm3 = (1.03)
'                facdec = facdec / 100
'                'Inicializacion de variables
'                For i = 1 To Fintab
'                    exced(i) = 0
'                    Flupen(i) = 0
'                    flucm(i) = 0
'                Next i
'                nmax = 0
'                cplan = cober
'                Fechap = nap * 12 + nmp
'                If indi = "D" Then
'                    mesdif = mesdif
'                Else
'                    mesdif = 0
'                End If
'                mesgar = mesgar / 12
'                ltot = mesgar + mesdif
'                If indi = "D" Then
'                    icont10 = 0:  icont20 = 0: icont11 = 0: icont21 = 0
'                    icont30 = 0: icont35 = 0: icont30Inv = 0
'                    icont40 = 0: icont77 = 0
'                    For j = 1 To nben
'                        nibe = 0
'                        If coinbe(j) = "S" Or coinbe(j) = "T" Or coinbe(j) = "I" Then nibe = 1
'                        If coinbe(j) = "N" Then nibe = 2
'                        If coinbe(j) = "P" Then nibe = 2
'                        If nibe = 0 Then
'                            x = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                        nsbe = 0
'                        If sexobe(j) = "M" Then nsbe = 1
'                        If sexobe(j) = "F" Then nsbe = 2
'                        If nsbe = 0 Then
'                            x = MsgBox("Error en ingreso de codigo de invalidez del beneficiario.", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                        Fechan = nanbe(j) * 12 + nmnbe(j)
'                        edabe = Fechap - Fechan
'                        If edabe < 1 Then edabe = 1
'                        If edabe > (Fintab * 12) Then
'                            vgError = 1023
'                            Exit Function
'                        End If
'                        If ncorbe(j) = 10 Or ncorbe(j) = 11 Then icont10 = icont10 + 1
'                        If ncorbe(j) = 20 Or ncorbe(j) = 21 Then icont20 = icont20 + 1
'                        If ncorbe(j) = 30 Then icont30 = icont30 + 1
'                        If ncorbe(j) = 30 And coinbe(j) <> "N" Then icont30Inv = icont30Inv + 1
'                        If ncorbe(j) = 35 Then icont35 = icont35 + 1
'                        If ncorbe(j) > 40 And ncorbe(j) < 50 Then icont40 = icont40 + 1
'                        If ncorbe(j) = 77 Then icont77 = icont77 + 1
'                    Next j
'                    If (icont10 > 0 Or icont20 > 0) And icont30 > 0 And icont30Inv = 0 Then
'                        For j = 1 To nben
'                            If ncorbe(j) = 10 Or ncorbe(j) = 11 Or ncorbe(j) = 20 Or ncorbe(j) = 21 Then
'                                edhm = (Fechap - (ijam(j) * 12 + ijmn(j))) + (mesdif * 12)
'                                If edhm >= (L18 * 12) Then
'                                    porcbe(j) = 0.42
'                                End If
'                            End If
'                        Next j
'                    End If
'                End If
'
'                If cober <> "S" Then
'                    'Calculo de flujos de vejez e invalidez
'                    facfam = ffam
'                    For j = 1 To nben
'                        Penben(j) = porcbe(j)
'                        numbep = numbep + 1
'                        If ncorbe(j) = 0 And j = 1 Then
'                            ni = 0
'                            If coinbe(j) = "S" Or coinbe(j) = "T" Or coinbe(j) = "I" Then ni = 1
'                            If coinbe(j) = "N" Then ni = 2
'                            If coinbe(j) = "P" Then ni = 3
'                            If ni = 0 Then
'                                x = MsgBox("Error en código de invalidez", vbCritical, "Proceso de cálculo Abortado")
'                                Tarifa_Todo = False
'                                Exit Function
'                            End If
'                            ns = 0
'                            If sexobe(j) = "M" Then ns = 1
'                            If sexobe(j) = "F" Then ns = 2
'                            If ns = 0 Then
'                                x = MsgBox("Error en código de sexo", vbCritical, "Proceso de cálculo Abortado")
'                                Tarifa_Todo = False
'                                Exit Function
'                            End If
'                            Fechan = nanbe(j) * 12 + nmnbe(j)
'                            edaca = Fechap - Fechan
'                            If edaca < 780 And ns = 1 And ni = 2 Then cplan = "A"
'                            If edaca < 720 And ns = 2 And ni = 2 Then cplan = "A"
'                            edaca = CInt(edaca / 12)
'                            If edaca <= 0 Or edaca > Fintab Then
'                                x = MsgBox("Error en edad de causante ", vbCritical, "Proceso de cálculo Abortado")
'                                Tarifa_Todo = False
'                                Exit Function
'                            End If
'                            limite1 = Fintab - edaca - 1
'                            nmax = limite1
'                            For i = 0 To limite1
'                                imas1 = i + 1
'                                edacai = edaca + i
'                                px = Lx(ns, ni, edacai) / Lx(ns, ni, edaca)
'                                edacas = edacai + 1
'                                Qx = ((Lx(ns, ni, edacai) - Lx(ns, ni, edacas))) / Lx(ns, ni, edaca)
'                                Flupen(imas1) = Flupen(imas1) + px * Penben(j)
'                                flucm(imas1) = flucm(imas1) + gtofun * Qx
'                            Next i
'                        End If
'                        If ncorbe(j) <> 0 Then
'                            nibe = 0
'                            If coinbe(j) = "S" Or coinbe(j) = "T" Or coinbe(j) = "I" Then nibe = 1
'                            If coinbe(j) = "N" Then nibe = 2
'                            If coinbe(j) = "P" Then nibe = 2
'                            If nibe = 0 Then
'                                x = MsgBox("Error en codificacion de invalidez beneficiario", vbCritical, "Proceso de cálculo Abortado")
'                                Tarifa_Todo = False
'                                Exit Function
'                            End If
'                            nsbe = 0
'                            If sexobe(j) = "M" Then nsbe = 1
'                            If sexobe(j) = "F" Then nsbe = 2
'                            If nsbe = 0 Then
'                                x = MsgBox("Error en codificacion de sexo beneficiario", vbCritical, "Proceso de cálculo Abortado")
'                                Tarifa_Todo = False
'                                Exit Function
'                            End If
'                            edabe = Fechap - (nanbe(j) * 12 + nmnbe(j))
'                            edabe = CInt(edabe / 12)
'                            If edabe < 1 Then edabe = 1
'                            If edabe > Fintab Then
'                                x = MsgBox("Error en Edad del beneficiario es mayor al limite de la tabla de mortalidad", vbCritical, "Proceso de cálculo Abortado")
'                                Tarifa_Todo = False
'                                Exit Function
'                            End If
'
'                            'Calculo de rentas vitalicias
'                            If ncorbe(j) = 10 Or ncorbe(j) = 11 Or _
'                               ncorbe(j) = 20 Or ncorbe(j) = 21 Or _
'                               ncorbe(j) = 41 Or ncorbe(j) = 46 Or _
'                               ncorbe(j) = 42 Or ncorbe(j) = 45 _
'                               Or (ncorbe(j) >= 30 And ncorbe(j) < 40 And (coinbe(j) = "P" Or coinbe(j) = "T") And edabe > L18) Then
'                                limite1 = Fintab - edabe - 1
'                                nmax = amax0(nmax, CInt(limite1))
'                                For i = 0 To limite1
'                                    imas1 = i + 1
'                                    edalbe = edabe + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    Flupen(imas1) = Flupen(imas1) + py * Penben(j) * facfam
'                                Next i
'                                limite2 = Fintab - edaca - 1
'                                limite = amin0(limite1, CInt(limite2))
'                                nmax = amax0(nmax, CInt(limite))
'                                For i = 0 To limite
'                                    imas1 = i + 1
'                                    edalca = edaca + i
'                                    edalbe = edabe + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                    Flupen(imas1) = Flupen(imas1) - (py * px * Penben(j) * facfam)
'                                Next i
'
'
'    '                           DERECHO A ACRECER
'                                If codcbe(j) <> "N" Then
'                                    edhm = (Fechap - (ijam(j) * 12 + ijmn(j)))
'                                    If edhm > L18 Then
'                                        nmdif = 0
'                                    Else
'                                        nmdif = L18 - edhm
'                                    End If
'                                    Ecadif = edabe + nmdif
'                                    limite1 = Fintab - Ecadif - 1
'                                    pension = Penben(j) * 0.2
'                                    nmax = amax0(nmax, CInt(limite1))
'                                    For i = 0 To limite1
'                                        imas1 = i + 1
'                                        edalbe = Ecadif + i
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        Flupen(imas1) = Flupen(imas1) + py * pension
'                                    Next i
'                                    limite2 = Fintab - edaca - 1
'                                    limite = amin0(limite1, CInt(limite2))
'                                    nmax = amax0(nmax, CInt(limite))
'                                    For i = 0 To limite1
'                                        imas1 = i + 1
'                                        edalbe = Ecadif + i
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                        Flupen(imas1) = Flupen(imas1) - (py * py * pension)
'                                    Next i
'                                End If
'                            Else
'                                'Calculo de rentas temporales
'                                If ncorbe(j) >= 30 And ncorbe(j) < 40 Then
'                                    If edabe > L18 Then
'                                    Else
'                                        mdif = L18 - edabe
'                                        nmdif = mdif
'                                        limite2 = Fintab - edaca
'                                        limite = amin0(nmdif, CInt(limite2)) - 1
'                                        nmax = amax0(nmax, CInt(limite))
'                                        For i = 0 To limite
'                                            imas1 = i + 1
'                                            edalca = edaca + i
'                                            edalbe = edabe + i
'                                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                            px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                            Flupen(imas1) = Flupen(imas1) + (py * Penben(j) - py * px * Penben(j)) * facfam
'                                        Next i
'                                    End If
'                                    If coinbe(j) <> "N" Then
'                                        edbedi = edabe + nmdif
'                                        limite3 = Fintab - edbedi - 1
'                                        limite4 = Fintab - (edaca + nmdif) - 1
'                                        nmax = amax0(nmax, CInt(limite3))
'                                        For i = 0 To limite3
'                                            imas1 = nmdif + i + 1
'                                            edalca = edaca + nmdif + i
'                                            edalbe = edbedi + i
'                                            edalca = amin0(edalca, CInt(Fintab))
'                                            py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                            px = Lx(ns, ni, edalca) / Lx(ns, ni, edaca)
'                                            Flupen(imas1) = Flupen(imas1) + (py - py * px) * porcbe(j) * facfam
'                                        Next i
'                                    End If
'                                End If
'                            End If
'                        End If
'                    Next j
'                Else
'                    'Calculo de flujos de Sobrevivencia
'                    For j = 1 To nben
'                        Penben(j) = porcbe(j)
'                        numbep = numbep + 1
'                        nibe = 0
'                        If coinbe(j) = "S" Or coinbe(j) = "T" Or coinbe(j) = "I" Then nibe = 1
'                        If coinbe(j) = "N" Then nibe = 2
'                        If coinbe(j) = "P" Then nibe = 2
'                        If nibe = 0 Then
'                            x = MsgBox("Error en codificacion de invalidez beneficiario", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'
'                        nsbe = 0
'                        If (sexobe(j) = "M") Then nsbe = 1
'                        If (sexobe(j) = "F") Then nsbe = 2
'                        If nsbe = 0 Then
'                            x = MsgBox("Error en codificacion de sexo beneficiario", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                        Fechan = nanbe(j) * 12 + nmnbe(j)
'                        edabe = Fechap - Fechan
'                        edabe = CInt(edabe / 12)
'                        If edabe > Fintab Then
'                            x = MsgBox("Error en Edad del beneficiario es mayor al limite de la tabla de mortalidad", vbCritical, "Proceso de cálculo Abortado")
'                            Tarifa_Todo = False
'                            Exit Function
'                        End If
'                        If edabe < 1 Then edabe = 1
'                        If ncorbe(j) = 10 Or ncorbe(j) = 11 Or _
'                           ncorbe(j) = 20 Or ncorbe(j) = 21 Or _
'                           ncorbe(j) = 41 Or ncorbe(j) = 46 Or _
'                           ncorbe(j) = 42 Or ncorbe(j) = 45 Or _
'                           (ncorbe(j) >= 30 And ncorbe(j) < 40 And _
'                           (coinbe(j) = "P" Or coinbe(j) = "T") And edabe > L18) Then
'                            limite1 = Fintab - edabe - 1
'                            nmax = amax0(nmax, CInt(limite1))
'                            For i = 0 To limite1
'                                imas1 = i + 1
'                                edalbe = edabe + i
'                                py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                If i <= ltot Then py = 1
'                                Flupen(imas1) = Flupen(imas1) + py * Penben(j)
'                            Next i
'
'                            If codcbe(j) <> "N" Then
'                                edhm = (Fechap - (ijam(j) * 12 + ijmn(j)))
'                                If edhm > L18 Then
'                                    nmdif = 0
'                                Else
'                                    nmdif = L18 - edhm
'                                End If
'                                Ecadif = edabe + nmdif
'                                limite1 = Fintab - Ecadif - 1
'                                pension = Penben(j) * 0.2
'                                nmax = amax0(nmax, CInt(limite1))
'                                For i = 0 To limite1
'                                    imas1 = i + 1
'                                    edalbe = Ecadif + i
'                                    py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                    Flupen(imas1) = Flupen(imas1) + py * pension
'                                Next i
'                            End If
'                        Else
'                            If ncorbe(j) >= 30 And ncorbe(j) < 40 Then
'                                If edabe > L18 Then
'                                Else
'                                    mdif = L18 - edabe
'                                    nmdif = amax0(mdif, ltot)
'                                    nmax = amax0(nmax, CInt(nmdif))
'                                    For i = 0 To nmdif
'                                        imas1 = i + 1
'                                        edalbe = edabe + i
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        If i <= ltot Then py = 1
'                                        Flupen(imas1) = Flupen(imas1) + py * Penben(j)
'                                    Next i
'                                End If
'                                If coinbe(j) <> "N" Then
'                                    kdif = mdif
'                                    edbedi = edabe + kdif
'                                    limite3 = Fintab - edbedi - 1
'                                    nmax = amax0(nmax, CInt(limite3))
'                                    For i = 0 To limite3
'                                        edalbe = edbedi + i
'                                        imas1 = kdif + i + 1
'                                        py = Ly(nsbe, nibe, edalbe) / Ly(nsbe, nibe, edabe)
'                                        If i <= ltot Then py = 1
'                                        Flupen(imas1) = Flupen(imas1) + py * Penben(j)
'                                    Next i
'                                End If
'                            Else
'                                x = MsgBox("Error en Código de relación del beneficiario", vbCritical, "Proceso de cálculo Abortado")
'                                Tarifa_Todo = False
'                                Exit Function
'                            End If
'                        End If
'                    Next j
'                End If
'
'                '**************************
'                '**************************
'                '
'                'Evaluación de Cotizaciones
'                '
'                '**************************
'                '**************************
'
'                If cober = "I" Or cober = "V" Or cober = "A" Or cober = "P" Then
'                    For i = 1 To nmax
'                        If i <= (ltot + 1) Then
'                            Flupen(i) = amax1(Flupen(i), 1)
'                            If i <= (mesdif + 1) Then
'                                Flupen(i) = 0
'                                flucm(i) = 0
'                            End If
'                        End If
'                    Next i
'                End If
'                If cober = "S" And mesdif > 0 Then
'                    For i = 1 To nmax
'                        If i <= (mesdif + 1) Then
'                            Flupen(i) = 0
'                            flucm(i) = 0
'                        End If
'                    Next i
'                End If
'                rmpol = 0
'                nmax = nmax + 1
'                sumapx = 0
'                sumaqx = 0
'                For i = 1 To nmax
'                    actual = (0.8 * Cp(i) / tmm ^ (i - 1)) + ((1 - 0.8 * Cp(i)) / tm3 ^ (i - 1))
'                    sumapx = sumapx + Flupen(i) * actual
'                    actua1 = (0.8 * Cp(i) / tmm ^ (i - 0.5)) + ((1 - 0.8 * Cp(i)) / tm3 ^ (i - 0.5))
'                    sumaqx = sumaqx + flucm(i) * actua1
'                Next i
'                If sumapx <= 0 Then
'                    penbase = 0
'                Else
'                    penbase = (salcta - sumaqx) / sumapx
'                End If
'                If sumapx <= 0 Then
'                    rmpol = 0
'                Else
'                    rmpol = sumapx * penbase + sumaqx
'                End If
'                tce = 0
'                vpte = 0
'                difres = 0
'                difre1 = 0
'                tir = 3
'                tinc = 0.1
'                TINC1 = 0.01
'225:
'                Tasa = (1 + tir / 100)
'                i = 1
'                For i = 1 To nmax
'                    vpte = vpte + (Flupen(i) * penbase / Tasa ^ (i - 1)) + (flucm(i) / Tasa ^ (i - 0.5))
'                Next i
'                difres = vpte - rmpol
'                If CDbl(Format(difres, "#0.00000")) >= 0 Then
'                    tir = tir + tinc
'                    If tir > 100 Then
'                        x = MsgBox("TASA TIR MAYOR A 100%", vbCritical)
'                        Tarifa_Todo = False
'                        Exit Function
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
'                salcta_eva = salcta
'                vppen = 0
'                vpcm = 0
'                For i = 1 To nmax
'                    vppen = vppen + Flupen(i) / (1 + tvmax) ^ (i - 1)
'                    vpcm = vpcm + flucm(i) / (1 + tvmax) ^ (i - 0.5)
'                Next i
'                penanu = (salcta_eva - vpcm) / vppen
'                If indi = "D" Then
'                    Vpptem = ((1 - 1 / ((1 + prc_Tasa_afp) ^ mesdif)) / prc_Tasa_afp) * (1 + prc_Tasa_afp)
'                    Add_porc_be = totpor
'                    If vppen > 0 Then
'                        Rete_sim = CDbl(Format(CDbl((salcta_eva * (1 / (Vpptem + (prc_Pension_afp) * vppen)))), "##0.00"))
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
'                exced(1) = (salcta_eva * (1 - comisi) - gastos - gto_supervivencia - reserva - gasemi - (CDbl(salcta_eva) * facdec)) * (1 - timp) - 1 / rdeuda * reserva
'                perdi = (reserva - salcta_eva + (salcta_eva * comisi)) / salcta_eva * 100
'                vld_comision = (salcta_eva * comisi)
'                vld_gtosbs = (salcta_eva * facdec)
'                vld_impuesto = (salcta_eva * (1 - comisi) - gastos - gto_supervivencia - reserva - gasemi - (salcta_eva * facdec)) * (1 - timp)
'                vld_puesta = 1 / rdeuda * reserva
'                flupag = penanu * Flupen(1) + flucm(1)
'                relres = 1
'                resfin = resfin * relres
'                rend = ((reserva + resfin) / 2) * (1 + 1 / rdeuda) * prodin(1)
'                varrm = reserva
'                resant = reserva
'                vlContarMaximo = nmax
'                For i = 2 To nmax
'                    flupag = penanu * Flupen(i) + flucm(i)
'                    relres = 1
'                    resfin = (resant - flupag) * tastce
'                    resfin = resfin * relres
'                    resfin = amax1(resfin, 0)
'                    varrm = resfin - resant
'                    gto = (gastos * Flupen(i)) + (gto_supervivencia * Flupen(i))
'                    exced(i) = (-flupag - gto - varrm + rend) * (1 - timp) - 1 / rdeuda * varrm
'                    If resfin <= 0 Then GoTo 131
'                    rend = ((resant + resfin) / 2) * (1 + 1 / rdeuda) * prodin(i)
'                    resant = resfin
'                    vld_comision = 0
'                    vld_gtosbs = gto
'                    vld_impuesto = (-flupag - gto - varrm + rend) * (1 - timp)
'                    vld_puesta = 1 / rdeuda * varrm
'                    'Se debe cortar la evaluacion del Flujo cuando la
'                    'Rentabilidad se haga Negativa
'                    If rend <= 0 Then
'                        vlContarMaximo = i
'                        Exit For
'                    End If
'                    'Se debe cortar la impresión del Flujo cuando el
'                    'Ajuste de Reservas se haga positivo
'                    If (varrm >= 0 And i > (mesdif + 1)) Then
'                        vlContarMaximo = i - 1
'                        Exit For
'                    End If
'                Next i
'131:
'                sumaex = 0
'                For i = 1 To vlContarMaximo
'                    sumaex = sumaex + exced(i) / (1 + tasac) ^ i
'                Next i
'
'                If sumaex >= 0 Then
'                    tirvta = tirvta + tinc
'                    If tirvta > 100 Then
'                        x = MsgBox("TASA TIR MAYOR A 100%", vbCritical)
'                        Tarifa_Todo = False
'                        Exit Function
'                    End If
'                    sumaex1 = sumaex
'                    sumaex = 0
'                    GoTo 222
'                End If
'                tirmax = tirvta + tinc * (sumaex / (sumaex1 - sumaex))
'                tirmax_ori = tirmax
'                tce = Format(tce, "###0.00")
'                TTirMax = (1 + tirmax / 100)
'                vppen = 0
'                vpcm = 0
'                For i = 1 To nmax
'                    vppen = vppen + Flupen(i) / TTirMax ^ (i - 1)
'                    vpcm = vpcm + flucm(i) / TTirMax ^ (i - 0.5)
'                Next i
'                penanu = (salcta_eva - vpcm) / vppen
'151:
'                If (perdi > PerMax) Then
'                    tirmax = tirmax - TINC1
'                    TTirMax = (1 + tirmax / 100)
'                    vppen = 0
'                    vpcm = 0
'                    RM = 0
'                    For i = 1 To nmax
'                        vppen = vppen + Flupen(i) / TTirMax ^ (i - 1)
'                        vpcm = vpcm + flucm(i) / TTirMax ^ (i - 0.5)
'                    Next i
'                    NewPen = (salcta_eva - vpcm) / vppen
'                    RM = (NewPen * sumapx) + sumaqx
'                    perdi = ((RM - salcta_eva + (salcta_eva * comisi)) / salcta_eva) * 100
'                    GoTo 151
'                End If
'                perdis = perdi
'
'                tirmax = amax1(penmin, tirmax)
'                tirmax = amin1(tirmax, penmax)
'                tirmax = Format(tirmax, "###0.00")
'                tassim = tirmax
'                'Grabar nmax en tabla difpol en campo tasa_simple
'                vgd_tasa_vta = tirmax
'                vgd_tce = tce
'
'                'Impresion informe tarifa.lis
'                If alt = "S" Then ia = 1
'                If alt = "G" Then ia = 2
'                If cplan = "A" Then IC = 1
'                If cplan = "V" Then IC = 2
'                If cplan = "I" Or cplan = "P" Then IC = 3
'                If cplan = "S" Then IC = 4
'                vppen = 0
'                vpcm = 0
'                TTirMax = (1 + tirmax / 100)
'                vppen = 0
'                vpcm = 0
'                For i = 1 To nmax
'                    vppen = vppen + Flupen(i) / TTirMax ^ (i - 1)
'                    vpcm = vpcm + flucm(i) / TTirMax ^ (i - 0.5)
'                Next i
'                vld_ReservaPensiones = 0
'                vld_ReservaSepelio = 0
'                vld_PensionAnual = 0
'                'Primer periodo
'                salcta_eva = salcta
'                penanu = (salcta_eva - vpcm) / vppen
'                If indi = "D" Then
'                    Vpptem = ((1 - 1 / ((1 + prc_Tasa_afp) ^ mesdif)) / prc_Tasa_afp) * (1 + prc_Tasa_afp)
'                    Add_porc_be = totpor
'                    If vppen > 0 Then
'                        Rete_sim = CDbl(Format(CDbl((salcta_eva * (1 / (Vpptem + (prc_Pension_afp) * vppen)))), "##0.00"))
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
'                flupag = penanu * Flupen(1) + flucm(1)
'                vld_ReservaPensiones = penanu * sumapx
'                vld_ReservaSepelio = sumaqx
'                vld_PensionAnual = penanu
'
'                'Guardar Monto de la Reserva
'                act = "UPDATE tmae_propuesta SET "
'                act = act & " prc_tasavta = " & Str(tirmax) & ", "
'                act = act & " prc_tasatce = " & Str(vgd_tce) & ", "
'                act = act & " mto_penanual = " & Str(Format(vld_PensionAnual, "#0.00")) & ", "
'                act = act & " mto_rmpension = " & Str(Format(vld_ReservaPensiones, "#0.00")) & ", "
'                act = act & " mto_rmgtosep = " & Str(Format(vld_ReservaSepelio, "#0.00")) & ", "
'                act = act & " mto_resmat = " & Str(Format(reserva, "#0.00")) & " "
'                act = act & " WHERE num_cot = '" & Coti & "' and "
'                act = act & " num_pro = " & indicador & ""
'                vgConectarBD.Execute (act)
'
'                'vlCorrCot  variable que indica que cotizacion corresponde a la propuesta
'                vlNumCoti = Mid(Coti, 1, 13) & Format(vlCorrCot, "00") & Mid(Coti, 16, 15)
'
'                If (vlCorrCot <> 0) Then
'                    act = "UPDATE tmae_cotizacion SET "
'                    act = act & " prc_tasavta = " & Str(tirmax) & ", "
'                    act = act & " prc_tasatce = " & Str(vgd_tce) & ", "
'                    act = act & " mto_penanual = " & Str(Format(vld_PensionAnual, "#0.00")) & ", "
'                    act = act & " mto_rmpension = " & Str(Format(vld_ReservaPensiones, "#0.00")) & ", "
'                    act = act & " mto_rmgtosep = " & Str(Format(vld_ReservaSepelio, "#0.00")) & ", "
'                    act = act & " mto_resmat = " & Str(Format(reserva, "#0.00")) & " "
'                    act = act & " WHERE num_cot = '" & vlNumCoti & "'"
'                    vgConectarBD.Execute (act)
'                End If
'
'                For j = 1 To nmax
'                    impres(1, j) = 0
'                    impres(2, j) = 0
'                    impres(3, j) = 0
'                    impres(4, j) = 0
'                    impres(5, j) = 0
'                    impres(6, j) = 0
'                    impres(7, j) = 0
'                    impres(8, j) = 0
'                    impres(9, j) = 0
'                Next j
'
'                gto_inicial = gasemi + (salcta_eva * facdec)
'                exced(1) = (salcta_eva * (1 - comisi) - gastos - gto_supervivencia - reserva - gasemi - (salcta_eva * facdec)) * (1 - timp) - 1 / rdeuda * reserva
'                Comision = salcta_eva * comisi
'                margen = salcta_eva - Comision - gastos - gto_supervivencia - gto_inicial - reserva
'                vlMargenDespuesImpuesto = margen * (1 - timp)
'                rend = 0
'
'                impres(1, 1) = salcta_eva
'                impres(2, 1) = Comision
'                impres(3, 1) = 0
'                impres(4, 1) = gastos + gto_supervivencia + gto_inicial
'                impres(5, 1) = reserva
'                impres(6, 1) = rend
'                impres(7, 1) = margen
'                impres(8, 1) = exced(1)
'                impres(9, 1) = vlMargenDespuesImpuesto
'                relres = 1
'                rend = ((reserva + resfin) / 2) * (1 + 1 / rdeuda) * prodin(1)
'                varrm = reserva
'                resant = reserva
'                vlContarMaximo = nmax
'
'                For i = 2 To nmax
'                    flupag = penanu * Flupen(i) + flucm(i)
'                    gto = (gastos * Flupen(i)) + (gto_supervivencia * Flupen(i))
'                    relres = 1
'                    resfin = (resant - flupag) * tastce
'                    resfin = resfin * relres
'                    varrm = resfin - resant
'                    rend = ((resant + resfin) / 2) * (1 + 1 / rdeuda) * prodin(i)
'                    exced(i) = (-flupag - gto - varrm + rend) * (1 - timp) - 1 / rdeuda * varrm
'                    margen = (-flupag - gto - varrm + rend)
'                    impres(1, i) = 0
'                    impres(2, i) = 0
'                    impres(3, i) = flupag
'                    impres(4, i) = gto
'                    impres(5, i) = varrm
'                    impres(6, i) = rend
'                    impres(7, i) = margen
'                    impres(8, i) = exced(i)
'                    vlMargenDespuesImpuesto = margen * (1 - timp)
'                    impres(9, i) = vlMargenDespuesImpuesto
'                    If (resfin <= 0) Then
'                        vlContarMaximo = i
'                        Exit For
'                    End If
'                    'Se debe cortar la impresión del Flujo cuando la
'                    'Rentabilidad se haga Negativa
'                    If rend <= 0 Then
'                        vlContarMaximo = i
'                        Exit For
'                    End If
'                    resant = resfin
'                    'Se debe cortar la impresión del Flujo cuando el
'                    'Ajuste de Reservas se haga positivo
'                    If (varrm >= 0 And i > (mesdif + 1)) Then
'                        vlContarMaximo = i - 1
'                        Exit For
'                    End If
'                Next i
'                If tirmax_ori > penmax Or tirmax_ori < penmin Then
'                    tasac_mod = 0
'                    vp_tasac = 0
'                    difres = 0
'                    difre1 = 0
'                    tasac_mod = 3
'                    tinc = 0.01
'300:
'                    Tasa = tasac_mod
'                    i = 1
'                    For i = 1 To vlContarMaximo
'                        vp_tasac = vp_tasac + exced(i) / (1 + Tasa / 100) ^ (i - 1)
'                    Next i
'                    difres = vp_tasac
'                    If difres >= 0 Then
'                        tasac_mod = tasac_mod + tinc
'                        If tasac_mod > 100 Then
'                            GoTo 301
'                            tasac_fin = tasac_mod
'                        End If
'                        difre1 = difres
'                        vp_tasac = 0
'                        GoTo 300
'                    End If
'                    tasac_fin = Tasa + tinc * (difres / (difre1 - difres))
'301:
'                    tasac = tasac_fin / 100
'                    tasa_tir = Format(CDbl(tasac * 100), "#0.00")
'                Else
'                    tasa_tir = Format(CDbl(tasac * 100), "#0.00")
'                End If
'
'                vlSql = "update tmae_propuesta set "
'                vlSql = vlSql & "prc_tasatir = " & Str(tasa_tir) & " "
'                vlSql = vlSql & " WHERE num_cot = '" & Coti & "' and "
'                vlSql = vlSql & " num_pro = " & indicador & ""
'                vgConectarBD.Execute (vlSql)
'
'                If (vlCorrCot <> 0) Then
'                    vlSql = "update tmae_cotizacion set "
'                    vlSql = vlSql & "prc_tasatir = " & Str(tasa_tir) & " where "
'                    vlSql = vlSql & "num_cot = '" & vlNumCoti & "'"
'                    vgConectarBD.Execute (vlSql)
'                End If
'                vlNumCoti = Mid(Coti, 1, 13) & Format(vlCorrCot, "00") & Mid(Coti, 16, 15)
'
'                    '--------------------------------------------------------
'                    'ABV : Este número máximo se puede reemplazar por la sgte.
'                    'Fórmula = fintab - edaca -1,
'                    'para lo cual se modificaría el número máximo de impresión
'                    '--------------------------------------------------------
'
'                For i = 1 To vlContarMaximo
'                        'Realizar modificación de la forma de impresión del
'                        'Detalle de la Evaluación
'                        '1. Comenzar la numeración de los años desde Cero
'                        '2. Mover las pensiones una posición, es decir, comienzan en el año 1
'                        '3. Cuando la Rentabilidad sea negativa, asignarle un Cero
'
'                        'Se debe cortar la impresión del Flujo cuando la
'                        'Rentabilidad se haga Negativa
'                    'Evaluaciones de la Propuesta
'                    vlSql = "insert into tmae_evapro ("
'                    vlSql = vlSql & "num_cot,num_pro,num_anno,mto_prima,mto_comision,"
'                    vlSql = vlSql & "mto_pension,mto_gasto,mto_renta,"
'                    vlSql = vlSql & "mto_ajuste,mto_margen,mto_margenimp,mto_excedente"
'                    vlSql = vlSql & ") values ("
'                    vlSql = vlSql & "'" & Coti & "',"
'                    vlSql = vlSql & "" & indicador & ","
'                    vlSql = vlSql & "" & i - 1 & ","
'                    vlSql = vlSql & "" & Str(Format(impres(1, i), "#,#0.00")) & "," 'Prima
'                    vlSql = vlSql & "" & Str(Format(impres(2, i), "#,#0.00")) & "," 'Comisión
'                    vlSql = vlSql & "" & Str(Format(impres(3, i), "#,#0.00")) & "," 'Pensiones
'                    vlSql = vlSql & "" & Str(Format(impres(4, i), "#,#0.00")) & "," 'Gastos
'                    If (impres(6, i) < 0) Then 'Reserva
'                        vlSql = vlSql & "" & Str(Format(0, "#,#0.00")) & ","
'                    Else
'                        vlSql = vlSql & "" & Str(Format(impres(6, i), "#,#0.00")) & ","
'                    End If
'                    vlSql = vlSql & "" & Str(Format(impres(5, i), "#,#0.00")) & "," 'Ajuste
'                    vlSql = vlSql & "" & Str(Format(impres(7, i), "#,#0.00")) & "," 'Margen
'                    vlSql = vlSql & "" & Str(Format(impres(9, i), "#,#0.00")) & "," 'Margen Des. Impuesto
'                    vlSql = vlSql & "" & Str(Format(impres(8, i), "#,#0.00")) & " " 'Excedente
'                    vlSql = vlSql & ")"
'                    vgConectarBD.Execute (vlSql)
'
'                    If (vlCorrCot <> 0) Then
'                        'Evaluaciones de la Cotización
'                        vlSql = "insert into tmae_evacot ("
'                        vlSql = vlSql & "num_cot,num_pro,num_anno,mto_prima,mto_comision,"
'                        vlSql = vlSql & "mto_pension,mto_gasto,mto_ajuste,"
'                        vlSql = vlSql & "mto_renta,mto_margen,mto_margenimp,mto_excedente"
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
'                        vlSql = vlSql & "" & Str(Format(impres(8, i), "#,#0.00")) & ""
'                        vlSql = vlSql & ")"
'                        vgConectarBD.Execute (vlSql)
'                    End If
'                Next i
'
'        End If
'    Next indicador
'    Tarifa_Todo = True
'
'    If (Frm_Progress.ProgressBar1.Value < 100) Then
'       Frm_Progress.ProgressBar1.Value = 100
'    End If
'
End Function
