Attribute VB_Name = "Mod_Pensiones"
'Estructura del Pago de Pensiones
Type TyDetPension
    Num_PerPago As String
    Num_Poliza As String
    Num_Orden As String
    Cod_ConHabDes As String
    Fec_IniPago As String
    Fec_TerPago As String
    Mto_ConHabDes As Double
    Edad As Integer
    EdadAños As Integer
    Num_IdenReceptor As String
    Cod_TipoIdenReceptor As Long
    Cod_Modulo As String
    Cod_TipReceptor As String
End Type

Type TyLiquidacion
    Num_PerPago As String
    Num_Poliza As String
    Num_Orden As String
    Fec_Pago As String
    Gls_Direccion As String
    Cod_Direccion As String
    Cod_TipPension As String
    Cod_ViaPago As String
    Cod_Banco As String
    Cod_TipCuenta As String
    Num_Cuenta As String
    Cod_Sucursal As String
    Cod_InsSalud As String
    Num_IdenReceptor As String
    Cod_TipoIdenReceptor As Long
    Gls_NomReceptor As String
    Gls_NomSegReceptor As String
    Gls_PatReceptor As String
    Gls_MatReceptor As String
    Cod_TipReceptor As String
    Mto_Haber As Double
    Mto_Descuento As Double
    Mto_LiqPagar As Double
    Cod_TipoPago As String
    Mto_BaseImp As Double
    Mto_BaseTri As Double
    Cod_Modulo As String
    Mto_Pension As Double
    'Cod_ModSalud As String
    'Mto_PlanSalud As Double
    Cod_TipoIdenPensionado As Long
    Num_IdenPensionado As String
    Cod_Parentesco As String
    Fec_IniPago As String
    Fec_TerPago As String
    Mto_Salud As Double
    Fac_Ajuste As Double
    Gls_TipoIdentCor As String
    Cod_Moneda As String
    'Para Derecho a Crecer
    Mto_PensionTotal As Double
    Prc_Pension As Double
    Mto_MaxSalud As Double
    Cod_DerGra As String
    Gls_Moneda As String
    Mto_pensiongarTotal As Double
    prc_factorAjus As Double
End Type

Type TyDatosGenerales
    Prc_SaludMin As Double
    Mto_MaxSalud As Double
    Mto_TopeBaseImponible As Double
    Val_UF As Double
    Val_UFUltDiaMes As Double
    MesesEdad18 As Long
    MesesEdad24 As Long
    MesesEdad60 As Long
    MesesEdad65 As Long
    Cod_ConceptoPension As String
    Cod_ConceptoDesctoSalud As String
    Cod_ConceptoGratificacion As String
End Type

Type TyTutor
    Num_IdenReceptor As String
    Cod_TipoIdenReceptor As Long
    Gls_NomReceptor As String
    Gls_NomSegReceptor As String
    Gls_PatReceptor As String
    Gls_MatReceptor As String
    Cod_TipReceptor As String
    Cod_GruFami As String
    Gls_Direccion As String
    Cod_Direccion As String
    Cod_TipPension As String
    Cod_ViaPago As String
    Cod_Banco As String
    Cod_TipCuenta As String
    Num_Cuenta As String
    Cod_Sucursal As String
    Num_Orden As Long
End Type

Type TyCrecer
    Num_PerPago As String 'Periodo en el que debe crecer
    Num_OrdenCon As Long 'Numero de Orden de la Conyuge
    Prc_Pension As Double 'Porcentaje de Pension del Beneficiario que ya no recibe pension
End Type

Global stDetPension() As TyDetPension
Global stDatGenerales As TyDatosGenerales
Global stLiquidacion() As TyLiquidacion
Global stTutor As TyTutor
Global Const clCodIsapreExento As String * 2 = "00"
Global vgNumIdenCliente As String
Global Const clNumeroMaximoDiasPrimerPago = 30
Global stCrecer() As TyCrecer 'Para guardar los porcentajes en que debe crecer la Conyuge
Global Const clModSaludDefecto As String = "PORCE"
Global Const clAjusteDesdeFechaDevengamiento As Boolean = True 'hqr 03/03/2011 Indica si se debe ajustar desde la fecha de devengamiento, TRUE: Se ajusta desde el devengamiento, FALSE: Se ajusta desde el inicio de vigencia de la Póliza
Function fgCargarVariablesGlobales()

    'Carga Constantes para Conceptos Fijos
    stDatGenerales.Cod_ConceptoDesctoSalud = "24"
    stDatGenerales.Cod_ConceptoPension = "01"
        
    stDatGenerales.MesesEdad18 = 18 * 12
    stDatGenerales.MesesEdad24 = 24 * 12
    stDatGenerales.MesesEdad65 = 65 * 12
    stDatGenerales.MesesEdad60 = 60 * 12
    stDatGenerales.Cod_ConceptoGratificacion = "80"
End Function


Function fgCalcularPrimerPago(iNumPoliza As String, iFecPago As String, iFecIniPag As String, iFecTerPag As String, iPension As Double, iPensionGar As Double, iMoneda As String, oExistenPagos As Boolean, iTipoAjuste, iMontoAjusteTri, iMontoAjusteMen, iFecVigencia, vlTipPer) As Boolean
    'Realiza el Cálculo del Primer Pago
    Dim vlNumBeneficiarios As Long
    Dim vlNumConceptos As Long
    Dim vlFecIniPag As String 'Para el mes actual
    Dim vlFecTerPag As String 'Para el mes actual
    Dim vlTermino As Boolean
    Dim vlPension As Double 'Monto de la Pensión por Beneficiaros
    Dim bResp As Integer 'Retorno de las Funciones
    Dim vlEdad As Long, vlEdadAños As Long, vlMaximoSalud As Double
    Dim vlFactorAjuste As Double, vlFactorAjuste2 As Double
    Dim vlMes As Long
    Dim vlAño As Long
    Dim vlPrimerMesAjuste As Boolean
    Dim vlPensionAjustada As Double
    Dim vlPensionGarAjustada As Double
    Dim vlPensionTotalAjustada As Double
    Dim vlPensionTotalGarAjustada As Double
    Dim vlPensionTotal As Double
    Dim vlFechaAjuste As String
    Dim vlNumCrecer As Long
    
    Dim vlPensionGar As Double  'RRR 28/10/2014
    Dim vlPensionGarTotal As Double 'RRR 28/10/2014
    Dim VlPeriodoGarVig As Boolean
    
    'hqr 05/09/2007 Agregados para Calculo de Primera Pensión Diferida
    Dim vlPrimerMesAjusteDif As Boolean
    Dim vlFactorAjusteDif As Double
    Dim vlFactorAjusteDif2 As Double
    Dim vlFechaIniDif As Date
    Dim vlFechaFinDif As Date
    Dim vlMesDif As Long, vlAñoDif As Long
    Dim vlFecDesdeAjuste As String
    Dim vlPensionInicial As Double
    Dim vlPrimerPagoBen As Boolean
    'fin hqr 05/09/2007
    
    Dim vlDiaDif As Long 'RRR 27/04/2016
    Dim vlPensionGarInicial As Double 'RRR 28/10/2014
    
    
    Dim vlFactorAjusteTasaFija As Double 'hqr 11/01/2011
    Dim vlFactorAjusteporIPC As Double 'hqr 11/01/2011
    Dim vlFactorAjusteDifTasaFijaTri As Double 'hqr 11/01/2011
    Dim vlFactorAjusteDifTasaFijaMen As Double 'hqr 14/02/2011
    Dim vlNumPension As Double 'hqr 14/02/2011
    Dim vlPasaTrimestre As Boolean 'hqr 14/02/2011
    Dim vlPasaTrimestreBen As Boolean 'hqr 15/02/2011 Para no pisar el valor de variables anteriores por cada beneficiario
    Dim vlNumPensionBen As Double 'hqr 15/02/2011 Para no pisar el valor de variables anteriores por cada beneficiario
    Dim vlFecDesdeAjustePension As String 'hqr 03/03/2011 Fecha desde la cual se realizan los ajustes
     
    Dim vlGlosaMoneda As String
    Dim ValQui As Boolean
    Dim vlDia As Long
    Dim vlFecDevVac As String
    Dim vlDiaDV As String
    Dim cuentaTRimestres As Integer
    'Dim vlTipRen As String
    Dim vlFactorAjusteDif_tmp As Double
    Dim fAjuste(1 To 1332) ', Porcbe(1 To 20) As Double
    Dim i As Integer
    Dim men As Integer

    On Error GoTo Errores
    Screen.MousePointer = vbHourglass
    fgCalcularPrimerPago = False
        
    vlPrimerMesAjuste = True
    vlFecDesdeAjuste = ""
    oExistenPagos = False
    vlFecDevVac = iFecIniPag
    cuentaTRimestres = 0
    'Obtiene Porcentaje Minimo de Salud
    If Not fgObtieneParametroVigencia("PS", "PSM", iFecPago, stDatGenerales.Prc_SaludMin) Then
        MsgBox "Debe ingresar Porcentaje Mínimo de Salud", vbCritical ', Me.Caption
        Exit Function
    End If
    
'    'Obtiene Tope Base Imponible (Para calculo de Aporte por CCAF)
'    If Not fgObtieneParametroVigencia("TBI", "MBM", vgFecPago, stDatGenerales.Mto_TopeBaseImponible) Then
'        MsgBox "Debe ingresar Monto Tope Base Imponible en UF", vbCritical, Me.Caption
'        Exit Function
'    End If
    
'    'Obtiene Valor UF a Fecha de Pago (Para no estarla obteniendo por Cada Caso
'    If Not fgObtieneConversion(vgFecPago, "UF", stDatGenerales.Val_UF) Then
'        MsgBox "Debe ingresar el Tipo de Cambio de la Moneda '" & iModalidad & "' a la Fecha de Pago", vbCritical, "Falta Tipo de Cambio"
'        Exit Function
'    End If

    fgCargarVariablesGlobales 'Carga conceptos globales
    
    '**************************************************************
    'Calcula Monto de la Pensión Actualizada hasta periodo anterior
    'a fecha de primer pago para polizas diferidas
    vlPrimerMesAjusteDif = True
    vlFactorAjusteDif = 1
    vlFactorAjusteDif_tmp = 1
    vlFactorAjusteDifTasaFijaTri = (1 + (iMontoAjusteTri / 100)) 'hqr 11/01/2011
    vlFactorAjusteDifTasaFijaMen = (1 + (iMontoAjusteMen / 100)) 'hqr 14/02/2011
    vlPensionInicial = iPension
    vlPensionGarInicial = iPensionGar
    vlGlosaMoneda = ""
    'If iMoneda = vgMonedaCodOfi Then 'Nuevos Soles
    
    
    
    'SI LA MONEDA ES INDEXADA
    If iTipoAjuste = cgAJUSTESOLES Then  'hqr 11/01/2011 --SOLES INDEXADOS
            'hqr  03/03/2011
        If clAjusteDesdeFechaDevengamiento Then 'hqr  03/03/2011
            vlFecDesdeAjustePension = iFecIniPag 'Fecha de Devengamiento
        Else
            vlFecDesdeAjustePension = iFecVigencia 'Fecha de Inicio de Vigencia de la Póliza
        End If
        'fin hqr  03/03/2011
        vgSql = "SELECT a.num_mesdif,a.fec_finperdif"
        vgSql = vgSql & " FROM pd_tmae_oripoliza a"
        vgSql = vgSql & " WHERE a.num_poliza = '" & Trim(iNumPoliza) & "'"
        Set vlRegistro1 = vgConexionBD.Execute(vgSql)
        If Not (vlRegistro1.EOF) Then
            vlFactorAjusteDif = 1 'Sin Ajuste
            vlFactorAjusteDifTasaFija = 0 'hqr 11/01/2011
            'vlFechaIniDif = DateSerial(Mid(vlFecDesdeAjustePension, 1, 4), Mid(vlFecDesdeAjustePension, 5, 2) + 1, Mid(vlFecDesdeAjustePension, 7, 2)) 'hqr 03/03/2011
            'vlFechaFinDif = DateSerial(Mid(vlRegistro1!fec_finperdif, 1, 4), Mid(vlRegistro1!fec_finperdif, 5, 2), Mid(vlRegistro1!fec_finperdif, 7, 2))
            vlFechaIniDif = DateSerial(Mid(vlFecDesdeAjustePension, 1, 4), Mid(vlFecDesdeAjustePension, 5, 2) + 1, 1)
            vlFechaFinDif = DateSerial(Mid(iFecPago, 1, 4), Mid(iFecPago, 5, 2), Mid(iFecPago, 7, 2))
                
            vlMontoPensionAct = iPension
            vlNumPension = 2 'hqr 14/02/2011
            men = 2
            vlPasaTrimestre = False 'hqr 14/02/2011
            fAjuste(1) = vlFactorAjusteDif
            Do While vlFechaFinDif >= vlFechaIniDif
                    ValQui = False
                    vlDiaDif = Day(vlFechaFinDif)
                    vlMesDif = Month(vlFechaIniDif)
                    vlAñoDif = Year(vlFechaIniDif)
                    
                    'CALCULA PRIMERO EL AJSUTE
                    'Obtener Factor anterior en la tabla de IPC
                    vlFechaAjuste = Format(DateAdd("m", -1, DateSerial(Year(vlFechaIniDif), Month(vlFechaIniDif), 1)), "yyyymmdd") 'Format(DateSerial(Year(vlFechaIniDif), Month(vlFechaIniDif), 1), "yyyymmdd")
                    If Not fgObtieneFactorAjusteIPCSolInd(vlFechaAjuste, ValQui, vlFactorAjusteDif) Then
                        MsgBox "No se encuentra Factor de Ajuste del Periodo: " & DateSerial(Mid(vlFechaAjuste, 1, 4), Mid(vlFechaAjuste, 5, 2), Mid(vlFechaAjuste, 7, 2)), vbCritical, "Error de Datos"
                        GoTo Deshacer
                    End If
                        
                    If men > 3 Then
                        men = 4
                    End If
'                    'Obtiene Factor de Ajuste Actual de 3 meses atras
                    vlFechaAjuste = Format(DateAdd("m", -(men), DateSerial(Year(vlFechaIniDif), Month(vlFechaIniDif), 1)), "yyyymmdd")
                    If Not fgObtieneFactorAjusteIPCSolInd(vlFechaAjuste, ValQui, vlFactorAjusteDif2) Then
                        MsgBox "No se encuentra Factor de Ajuste del Periodo: " & DateSerial(Mid(vlFechaAjuste, 1, 4), Mid(vlFechaAjuste, 5, 2), Mid(vlFechaAjuste, 7, 2)), vbCritical, "Error de Datos"
                        GoTo Deshacer
                    End If
                        
                    vlFactorAjusteDif = Format(Math.Round(vlFactorAjusteDif, 2) / Math.Round(vlFactorAjusteDif2, 2), "##0.000000")
                    
                    If ((vlMesDif Mod 3) - 1) = 0 Then 'TRIMESTRE
                        'vlFactorAjusteDif = Format(vlFactorAjusteDif / vlFactorAjusteDif2, "##0.0000")
                        vlFactorAjusteDif_tmp = vlFactorAjusteDif * vlFactorAjusteDif_tmp
                        If men = 4 Then
                            vlFactorAjusteDif = vlFactorAjusteDif_tmp
                        End If
                    Else 'No está en mes del trimestre
                        vlFactorAjusteDif = vlFactorAjusteDif_tmp
                    End If
                    fAjuste(vlNumPension) = vlFactorAjusteDif
                    vlFechaIniDif = DateAdd("m", 1, vlFechaIniDif)
                    vlNumPension = vlNumPension + 1
                    men = men + 1
            Loop
        End If
    ElseIf iTipoAjuste = cgAJUSTETASAFIJA Then '-- DOLARES Y SOLES AJUSTADOS
            'hqr  03/03/2011
        If clAjusteDesdeFechaDevengamiento Then 'hqr  03/03/2011
            vlFecDesdeAjustePension = iFecIniPag 'Fecha de Devengamiento
        Else
            vlFecDesdeAjustePension = iFecVigencia 'Fecha de Inicio de Vigencia de la Póliza
        End If
        'fin hqr  03/03/2011
        vgSql = "SELECT a.num_mesdif,a.fec_finperdif, a.cod_tipren"
        vgSql = vgSql & " FROM pd_tmae_oripoliza a"
        vgSql = vgSql & " WHERE a.num_poliza = '" & Trim(iNumPoliza) & "'"
        Set vlRegistro1 = vgConexionBD.Execute(vgSql)
        If Not (vlRegistro1.EOF) Then
            'If vlRegistro1!Num_MesDif > 0 And vlTipPer <> "6" Then 'Si se trata de periodo diferido
                vlFactorAjuste = 1
                vlFactorAjusteDif = 1 'Sin Ajuste
                vlFactorAjusteDifTasaFija = 0 'hqr 11/01/2011
                vlFactorAjusteDif2 = 1
                vlFactorAjusteDif_tmp = 1
                vlNumPension = 1
                vlFechaIniDif = DateSerial(Mid(vlFecDesdeAjustePension, 1, 4), Mid(vlFecDesdeAjustePension, 5, 2) + 1, Mid(vlFecDesdeAjustePension, 7, 2)) 'hqr 03/03/2011
                'vlFechaFinDif = DateSerial(Mid(vlRegistro1!fec_finperdif, 1, 4), Mid(vlRegistro1!fec_finperdif, 5, 2), Mid(vlRegistro1!fec_finperdif, 7, 2))
                vlFechaFinDif = DateSerial(Mid(iFecPago, 1, 4), Mid(iFecPago, 5, 2), Mid(iFecPago, 7, 2))
                vlMontoPensionAct = iPension
                'vlPasaTrimestre = False 'hqr 14/02/2011
                fAjuste(1) = vlFactorAjusteDif
                Do While vlFechaFinDif >= vlFechaIniDif
                    vlNumPension = vlNumPension + 1
                    vlMesDif = Month(vlFechaIniDif)
                    vlAñoDif = Year(vlFechaIniDif)
                    vlFactorAjusteDif = (1 + (iMontoAjusteMen / 100))
                    vlFactorAjusteDif2 = vlFactorAjusteDif2 * vlFactorAjusteDif
                    If (((vlMesDif Mod 3) - 1) = 0) Then

                        vlFactorAjuste = vlFactorAjusteDif2
                        vlFactorAjusteDif_tmp = vlFactorAjuste
                    
                    Else 'No está en mes del trimestre
                        vlFactorAjuste = vlFactorAjusteDif_tmp
                    End If
                    
                    fAjuste(vlNumPension) = vlFactorAjuste
                    vlFechaIniDif = DateAdd("m", 1, vlFechaIniDif)
                    
                Loop
                
            If vlRegistro1!Num_MesDif > 0 And vlTipPer <> "6" Then 'Si se trata de periodo diferido
           
            End If
        End If
    End If
    '**************************************************************
    
    
    vlNumPension = 0
    vlNumBeneficiarios = 0
    vlNumConceptos = 0
    vgSql = "SELECT a.num_orden, a.cod_par, a.cod_grufam, a.prc_pension, "
    vgSql = vgSql & "a.cod_tipoidenben, a.num_idenben, a.fec_nacben, b.cod_dercre, b.cod_dergra,"
    vgSql = vgSql & "a.gls_nomben, a.gls_nomsegben, a.gls_patben, a.gls_matben,"
    vgSql = vgSql & "b.cod_banco, b.cod_direccion, b.cod_isapre, "
    vgSql = vgSql & "b.cod_sucursal, b.cod_tipcuenta, b.cod_tippension, "
    vgSql = vgSql & "b.cod_viapago, b.gls_direccion, b.num_cuenta, "
    vgSql = vgSql & "b.fec_finpergar, b.fec_finperdif, c.gls_tipoidencor, a.cod_sitinv, b.num_mesdif, d.gls_elemento as glosamoneda, a.prc_pensiongar, b.cod_tipren, b.num_mesgar, b.cod_tippension "
    vgSql = vgSql & "FROM pd_tmae_oripolben a, pd_tmae_oripoliza b, ma_tpar_tipoiden c, ma_tpar_tabcod d "
    vgSql = vgSql & "WHERE "
    vgSql = vgSql & "a.num_poliza = b.num_poliza AND "
    vgSql = vgSql & "a.cod_tipoidenben = c.cod_tipoiden AND "
    vgSql = vgSql & "a.num_poliza = '" & Trim(iNumPoliza) & "' "
    vgSql = vgSql & "AND a.cod_estpension = 99 " 'campo aun no existe
    vgSql = vgSql & "AND d.cod_tabla = 'TM' " 'Tabla de Monedas
    vgSql = vgSql & "AND d.cod_elemento = b.cod_moneda "
    vgSql = vgSql & "ORDER BY a.num_poliza,a.cod_grufam, a.cod_par"
    vlNumCrecer = 0
    Set vlRegistro1 = vgConexionBD.Execute(vgSql)
    If Not (vlRegistro1.EOF) Then
        stTutor.Cod_GruFami = "-1"
        vlPrimerPagoBen = True 'Para que pregunte una sola vez
        Do While Not (vlRegistro1.EOF)
            'vlTipRen = vlRegistro1!Cod_TipRen
            'stDetPension
            vlNumPension = 0
            If vlRegistro1!Num_MesDif > 0 Then
                vlNumPension = vlRegistro1!Num_MesDif
            End If
            
            If vlRegistro1!Prc_Pension <= 0 Then
                GoTo Siguiente
            End If
            vlPasaTrimestreBen = vlPasaTrimestre 'hqr 15/02/2011
            vlNumPensionBen = IIf(vlNumPension = 0, 1, vlNumPension) 'hqr 15/02/2011
            'Monto Pensión
            vlPensionAjustada = Format((vlRegistro1!Prc_Pension / 100) * vlPensionInicial, "0.00")
            If vlRegistro1!Num_MesGar > 0 Then
                If vlNumPensionBen = 1 Then
                    vlPensionGarAjustada = Format((vlRegistro1!Prc_Pension / 100) * vlPensionInicial, "0.00")
                Else
                    vlPensionGarAjustada = Format((vlRegistro1!Prc_PensionGar / 100) * vlPensionGarInicial, "0.00")
                End If
            End If
            
            vlPensionTotalAjustada = Format(vlPensionInicial, "0.00")
            If vlRegistro1!Cod_TipPension <> "08" Then
                vlPensionTotalGarAjustada = Format(vlPensionGarInicial, "0.00")
            Else
                vlPensionTotalGarAjustada = Format(vlPensionGarAjustada, "0.00")
                vlMtoPenGarUf = vlPensionTotalGarAjustada
            End If
            
            
            If vlTipPer = "6" Then
                vlFecIniPag = Mid(iFecIniPag, 1, 6) & "01" 'Primer dia del mes
                vlFecTerPag = Format(DateSerial(Mid(vlFecIniPag, 1, 4), Mid(vlFecIniPag, 5, 2) + 1, 0), "yyyymmdd")
                If iFecTerPag < vlFecTerPag Then
                    vlFecTerPag = iFecTerPag 'Por si es solo un mes
                End If
            Else
                If Not IsNull(vlRegistro1!fec_finperdif) Then
                    'If vlRegistro1!fec_finperdif >= vlFecTerPag Then
                        'GoTo Siguiente 'No corresponde pago, ya que esté en el periodo diferifo
                        vlFecIniPag = Format(DateSerial(Mid(vlRegistro1!fec_finperdif, 1, 4), Mid(vlRegistro1!fec_finperdif, 5, 2) + 1, 1), "yyyymmdd")
                        If vlFecIniPag = iFecPago Then 'Es la misma fecha tope
                            vlFecTerPag = Format(DateSerial(Mid(vlFecIniPag, 1, 4), Mid(vlFecIniPag, 5, 2), 0), "yyyymmdd")
                        Else
                            vlFecTerPag = Format(DateSerial(Mid(vlFecIniPag, 1, 4), Mid(vlFecIniPag, 5, 2) + 1, 0), "yyyymmdd")
                        End If
                        If iFecTerPag < vlFecTerPag Then
                            vlFecTerPag = iFecTerPag 'Por si es solo un mes
                        End If
                    'End If
                Else
                    vlFecIniPag = Mid(iFecIniPag, 1, 6) & "01" 'Primer dia del mes
                    vlFecTerPag = Format(DateSerial(Mid(vlFecIniPag, 1, 4), Mid(vlFecIniPag, 5, 2) + 1, 0), "yyyymmdd")
                    If iFecTerPag < vlFecTerPag Then
                        vlFecTerPag = iFecTerPag 'Por si es solo un mes
                    End If
                End If
            End If
    
            If vlRegistro1!Cod_Par < 30 Then 'Padres quedan registrados para Tutores
                stTutor.Cod_GruFami = vlRegistro1!Cod_GruFam
                stTutor.Cod_TipReceptor = "M"
                stTutor.Cod_TipoIdenReceptor = vlRegistro1!Cod_TipoIdenBen
                stTutor.Num_IdenReceptor = vlRegistro1!Num_IdenBen
                stTutor.Gls_MatReceptor = IIf(IsNull(vlRegistro1!Gls_MatBen), "", vlRegistro1!Gls_MatBen)
                stTutor.Gls_PatReceptor = vlRegistro1!Gls_PatBen
                stTutor.Gls_NomReceptor = IIf(IsNull(vlRegistro1!Gls_NomBen), "", vlRegistro1!Gls_NomBen)
                stTutor.Gls_NomSegReceptor = IIf(IsNull(vlRegistro1!Gls_NomSegBen), "", vlRegistro1!Gls_NomSegBen)
                stTutor.Gls_Direccion = IIf(IsNull(vlRegistro1!Gls_Direccion), "", vlRegistro1!Gls_Direccion)
                stTutor.Cod_Direccion = vlRegistro1!Cod_Direccion
                stTutor.Num_Orden = vlRegistro1!Num_Orden
                'stTutor.Cod_ViaPago = vlRegistro1!Cod_ViaPago
                'stTutor.Cod_Banco = IIf(IsNull(vlRegistro1!Cod_Banco), "NULL", vlRegistro1!Cod_Banco)
                'stTutor.Cod_TipCuenta = IIf(IsNull(vlRegistro1!Cod_TipCuenta), "NULL", vlRegistro1!Cod_TipCuenta)
                'stTutor.Num_Cuenta = IIf(IsNull(vlRegistro1!Num_Cuenta), "NULL", vlRegistro1!Num_Cuenta)
                'stTutor.Cod_Sucursal = IIf(IsNull(vlRegistro1!Cod_Sucursal), "NULL", vlRegistro1!Cod_Sucursal)
            End If
                        
            vlTermino = False
            vlPrimerMesAjuste = True 'para que se reinicie con cada beneficiario
            Do While vlFecTerPag <= iFecTerPag And vlFecIniPag < iFecPago And vlFecIniPag < vlFecTerPag 'Or vlTipRen = "6"
                vlNumBeneficiarios = vlNumBeneficiarios + 1
                vlNumPension = vlNumPension + 1
                ReDim Preserve stLiquidacion(vlNumBeneficiarios)
                stLiquidacion(vlNumBeneficiarios - 1).Cod_Moneda = iMoneda
                stLiquidacion(vlNumBeneficiarios - 1).Gls_Moneda = vlRegistro1!glosamoneda
                stLiquidacion(vlNumBeneficiarios - 1).Gls_TipoIdentCor = vlRegistro1!GLS_TIPOIDENCOR
                stLiquidacion(vlNumBeneficiarios - 1).Num_Orden = vlRegistro1!Num_Orden
                stLiquidacion(vlNumBeneficiarios - 1).Num_Poliza = iNumPoliza
                stLiquidacion(vlNumBeneficiarios - 1).Cod_Banco = vlRegistro1!Cod_Banco
                stLiquidacion(vlNumBeneficiarios - 1).Cod_Direccion = vlRegistro1!Cod_Direccion
                stLiquidacion(vlNumBeneficiarios - 1).Cod_InsSalud = IIf(IsNull(vlRegistro1!cod_isapre), "NULL", vlRegistro1!cod_isapre)
                stLiquidacion(vlNumBeneficiarios - 1).Cod_Sucursal = vlRegistro1!Cod_Sucursal
                stLiquidacion(vlNumBeneficiarios - 1).Cod_TipCuenta = vlRegistro1!Cod_TipCuenta
                stLiquidacion(vlNumBeneficiarios - 1).Cod_TipoPago = "P" 'Primer Pago
                stLiquidacion(vlNumBeneficiarios - 1).Cod_TipPension = vlRegistro1!Cod_TipPension
                stLiquidacion(vlNumBeneficiarios - 1).Cod_ViaPago = vlRegistro1!Cod_ViaPago
                stLiquidacion(vlNumBeneficiarios - 1).Fec_Pago = iFecPago
                stLiquidacion(vlNumBeneficiarios - 1).Gls_Direccion = IIf(IsNull(vlRegistro1!Gls_Direccion), "", vlRegistro1!Gls_Direccion)
                stLiquidacion(vlNumBeneficiarios - 1).Num_Cuenta = IIf(IsNull(vlRegistro1!Num_Cuenta), "NULL", vlRegistro1!Num_Cuenta)
                stLiquidacion(vlNumBeneficiarios - 1).Cod_TipoIdenPensionado = vlRegistro1!Cod_TipoIdenBen
                stLiquidacion(vlNumBeneficiarios - 1).Num_IdenPensionado = vlRegistro1!Num_IdenBen
                stLiquidacion(vlNumBeneficiarios - 1).Cod_Parentesco = vlRegistro1!Cod_Par
            
                stLiquidacion(vlNumBeneficiarios - 1).Num_PerPago = Mid(vlFecIniPag, 1, 6)
                bResp = fgCalculaEdad(vlRegistro1!Fec_NacBen, vlFecTerPag)
                If bResp = "-1" Then 'Error
                    GoTo Deshacer
                Else
                    If bResp = "-2" Then 'Aun no nacía
                        vlNumBeneficiarios = vlNumBeneficiarios - 1
                        ReDim Preserve stLiquidacion(vlNumBeneficiarios)
                        GoTo Siguiente
                    End If
                End If
                vlEdad = bResp
                vlEdadAños = fgConvierteEdadAños(vlEdad)
                'stDetPension.Edad = bResp
                'stDetPension.EdadAños = fgConvierteEdadAños(stDetPension.Edad)
                If vlRegistro1!Cod_Par >= 30 And vlRegistro1!Cod_Par <= 35 Then 'Hijos
                    If vlEdad >= stDatGenerales.MesesEdad18 And vlRegistro1!Cod_SitInv = "N" Then 'Hijos Sanos
                        'OBS: Se asume que el mes de los 18 años se paga completo
                        If vlRegistro1!Cod_DerCre = "S" Then
                            'Para registrar los Derechos a Crecer
                            vlNumCrecer = vlNumCrecer + 1
                            ReDim Preserve stCrecer(vlNumCrecer)
                            stCrecer(vlNumCrecer).Num_OrdenCon = stTutor.Num_Orden
                            stCrecer(vlNumCrecer).Prc_Pension = vlRegistro1!Prc_Pension
                            stCrecer(vlNumCrecer).Num_PerPago = stLiquidacion(vlNumBeneficiarios - 1).Num_PerPago
                        End If
                        vlNumBeneficiarios = vlNumBeneficiarios - 1
                        ReDim Preserve stLiquidacion(vlNumBeneficiarios)
                        GoTo Siguiente 'Va al Siguiente Beneficiario, ya que éste no tiene Derecho
                    End If
                End If
                'Inicializa Monto Haber y Descuento
                stLiquidacion(vlNumBeneficiarios - 1).Mto_Haber = 0
                stLiquidacion(vlNumBeneficiarios - 1).Mto_Descuento = 0
                
                'Asignación de Tutores
                If vlRegistro1!Cod_Par >= 30 And vlRegistro1!Cod_Par <= 35 And vlEdad < stDatGenerales.MesesEdad18 And stTutor.Cod_GruFami = vlRegistro1!Cod_GruFam Then 'El Tutor debe ser la Madre
                    stLiquidacion(vlNumBeneficiarios - 1).Cod_TipReceptor = stTutor.Cod_TipReceptor 'MADRE
                    stLiquidacion(vlNumBeneficiarios - 1).Cod_TipoIdenReceptor = stTutor.Cod_TipoIdenReceptor
                    stLiquidacion(vlNumBeneficiarios - 1).Num_IdenReceptor = stTutor.Num_IdenReceptor
                    stLiquidacion(vlNumBeneficiarios - 1).Gls_NomReceptor = stTutor.Gls_NomReceptor
                    stLiquidacion(vlNumBeneficiarios - 1).Gls_NomSegReceptor = stTutor.Gls_NomSegReceptor
                    stLiquidacion(vlNumBeneficiarios - 1).Gls_PatReceptor = stTutor.Gls_PatReceptor
                    stLiquidacion(vlNumBeneficiarios - 1).Gls_MatReceptor = stTutor.Gls_MatReceptor
                    stLiquidacion(vlNumBeneficiarios - 1).Gls_Direccion = stTutor.Gls_Direccion
                    stLiquidacion(vlNumBeneficiarios - 1).Cod_Direccion = stTutor.Cod_Direccion
                Else 'Else se le Pagará a El Mismo
                    stLiquidacion(vlNumBeneficiarios - 1).Cod_TipReceptor = "P" 'Causante
                    stLiquidacion(vlNumBeneficiarios - 1).Cod_TipoIdenReceptor = vlRegistro1!Cod_TipoIdenBen
                    stLiquidacion(vlNumBeneficiarios - 1).Num_IdenReceptor = vlRegistro1!Num_IdenBen
                    stLiquidacion(vlNumBeneficiarios - 1).Gls_NomReceptor = vlRegistro1!Gls_NomBen
                    stLiquidacion(vlNumBeneficiarios - 1).Gls_NomSegReceptor = IIf(IsNull(vlRegistro1!Gls_NomSegBen), "", vlRegistro1!Gls_NomSegBen)
                    stLiquidacion(vlNumBeneficiarios - 1).Gls_PatReceptor = vlRegistro1!Gls_PatBen
                    stLiquidacion(vlNumBeneficiarios - 1).Gls_MatReceptor = IIf(IsNull(vlRegistro1!Gls_MatBen), "", vlRegistro1!Gls_MatBen)
                    stLiquidacion(vlNumBeneficiarios - 1).Gls_Direccion = IIf(IsNull(vlRegistro1!Gls_Direccion), "", vlRegistro1!Gls_Direccion)
                    stLiquidacion(vlNumBeneficiarios - 1).Cod_Direccion = vlRegistro1!Cod_Direccion
                End If
                
                'Pago por Conceptos
                vlNumConceptos = vlNumConceptos + 1
                ReDim Preserve stDetPension(vlNumConceptos)
                
                'Ajusta Pensión
                stLiquidacion(vlNumBeneficiarios - 1).Fac_Ajuste = IIf(fAjuste(vlNumPension) = 0, 1, fAjuste(vlNumPension))
                
                'Monto Pensión
                
                vlFactorAjuste = fAjuste(vlNumPension)
                If vlFactorAjuste = 0 Then
                    vlFactorAjuste = 1
                End If
                vlPensionTotalAjustada = Format(iPension * vlFactorAjuste, "##0.00")
                vlPensionTotalGarAjustada = Format(iPensionGar * vlFactorAjuste, "##0.00")
                vlPensionAjustada = Format((vlRegistro1!Prc_Pension / 100) * vlPensionTotalAjustada, "##0.00")
                vlPensionGarAjustada = Format((vlRegistro1!Prc_PensionGar / 100) * vlPensionTotalGarAjustada, "##0.00")
                vlPasaTrimestreBen = True
                
                vlNumPensionBen = vlNumPensionBen + 1 'hqr 17/02/2011
                
                vlPensionGar = 0
                vlPensionGarTotal = 0
                
                vlPension = 0
                vlPensionTotal = 0
                
                VlPeriodoGarVig = False
                
                'Las pensiones de supervivencia no consideran monto garantizado - MateriaGris JaimeRios 09/02/2018
                If vlRegistro1!Cod_TipPension <> "08" Then
                    If Not IsNull(vlRegistro1!fec_finpergar) Then
                        If vlRegistro1!fec_finpergar <= vlFecIniPag Then
                            vlPension = vlPensionAjustada
                            vlPensionTotal = vlPensionTotalAjustada
                        Else
                            VlPeriodoGarVig = True
                            vlPension = vlPensionAjustada
                            vlPensionTotal = vlPensionTotalAjustada
                            
                            vlPensionGar = vlPensionGarAjustada
                            vlPensionGarTotal = vlPensionTotalGarAjustada
                        End If
                    Else
                        vlPension = vlPensionAjustada
                        vlPensionTotal = vlPensionTotalAjustada
                    End If
                Else
                    vlPension = vlPensionAjustada
                    vlPensionTotal = vlPensionTotalAjustada
                    vlPensionGar = vlPensionGarAjustada
                    vlPensionGarTotal = vlPensionTotalGarAjustada
                End If
                                
                'vlPension = Format(vlPension * stLiquidacion(vlNumBeneficiarios - 1).Fac_Ajuste, "##0.00")
                stDetPension(vlNumConceptos - 1).Num_IdenReceptor = stLiquidacion(vlNumBeneficiarios - 1).Num_IdenReceptor
                stDetPension(vlNumConceptos - 1).Cod_TipoIdenReceptor = stLiquidacion(vlNumBeneficiarios - 1).Cod_TipoIdenReceptor
                stDetPension(vlNumConceptos - 1).Cod_TipReceptor = stLiquidacion(vlNumBeneficiarios - 1).Cod_TipReceptor
                stDetPension(vlNumConceptos - 1).Cod_ConHabDes = stDatGenerales.Cod_ConceptoPension
                stDetPension(vlNumConceptos - 1).Fec_IniPago = vlFecIniPag
                stDetPension(vlNumConceptos - 1).Fec_TerPago = vlFecTerPag
                stDetPension(vlNumConceptos - 1).Mto_ConHabDes = IIf(VlPeriodoGarVig = True, vlPensionGar, vlPension)
                stDetPension(vlNumConceptos - 1).Num_Orden = stLiquidacion(vlNumBeneficiarios - 1).Num_Orden
                stDetPension(vlNumConceptos - 1).Num_PerPago = stLiquidacion(vlNumBeneficiarios - 1).Num_PerPago
                stDetPension(vlNumConceptos - 1).Num_Poliza = stLiquidacion(vlNumBeneficiarios - 1).Num_Poliza
                stLiquidacion(vlNumBeneficiarios - 1).Mto_Pension = stDetPension(vlNumConceptos - 1).Mto_ConHabDes
                'Gratificacion
                
                'xxx
                
                vlMes = Mid(vlFecIniPag, 5, 2)
                If vlRegistro1!Cod_DerGra = "S" Then
                    If (vlMes = 7) Or (vlMes = 12) Then 'Gratificación se paga en Julio y Diciembre
                        vlNumConceptos = vlNumConceptos + 1
                        ReDim Preserve stDetPension(vlNumConceptos)
                        stDetPension(vlNumConceptos - 1).Num_IdenReceptor = stLiquidacion(vlNumBeneficiarios - 1).Num_IdenReceptor
                        stDetPension(vlNumConceptos - 1).Cod_TipoIdenReceptor = stLiquidacion(vlNumBeneficiarios - 1).Cod_TipoIdenReceptor
                        stDetPension(vlNumConceptos - 1).Cod_TipReceptor = stLiquidacion(vlNumBeneficiarios - 1).Cod_TipReceptor
                        stDetPension(vlNumConceptos - 1).Cod_ConHabDes = stDatGenerales.Cod_ConceptoGratificacion
                        stDetPension(vlNumConceptos - 1).Fec_IniPago = vlFecIniPag
                        stDetPension(vlNumConceptos - 1).Fec_TerPago = vlFecTerPag
                        stDetPension(vlNumConceptos - 1).Mto_ConHabDes = stLiquidacion(vlNumBeneficiarios - 1).Mto_Pension
                        stDetPension(vlNumConceptos - 1).Num_Orden = stLiquidacion(vlNumBeneficiarios - 1).Num_Orden
                        stDetPension(vlNumConceptos - 1).Num_PerPago = stLiquidacion(vlNumBeneficiarios - 1).Num_PerPago
                        stDetPension(vlNumConceptos - 1).Num_Poliza = stLiquidacion(vlNumBeneficiarios - 1).Num_Poliza
                        stLiquidacion(vlNumBeneficiarios - 1).Mto_Pension = stLiquidacion(vlNumBeneficiarios - 1).Mto_Pension + stDetPension(vlNumConceptos - 1).Mto_ConHabDes
                    End If
                End If
                
                'Monto Salud
                stLiquidacion(vlNumBeneficiarios - 1).Mto_Salud = 0
                
                'Para el Derecho a Crecer
                stLiquidacion(vlNumBeneficiarios - 1).Mto_PensionTotal = vlPensionTotal
                stLiquidacion(vlNumBeneficiarios - 1).Mto_pensiongarTotal = vlPensionGarTotal
                stLiquidacion(vlNumBeneficiarios - 1).Prc_Pension = vlRegistro1!Prc_Pension
                stLiquidacion(vlNumBeneficiarios - 1).Cod_DerGra = vlRegistro1!Cod_DerGra
                stLiquidacion(vlNumBeneficiarios - 1).prc_factorAjus = vlFactorAjuste
                'Hasta acá para el Derecho a Crecer
                stLiquidacion(vlNumBeneficiarios - 1).Mto_MaxSalud = 0
                If Not IsNull(vlRegistro1!cod_isapre) Then
                    If vlRegistro1!cod_isapre <> clCodIsapreExento Then
                        If Not fgObtieneMontoMaximoSalud(iMoneda, vlFecIniPag, vlMaximoSalud) Then
                            GoTo Deshacer
                        End If
                        vlNumConceptos = vlNumConceptos + 1
                        ReDim Preserve stDetPension(vlNumConceptos)
                        stDetPension(vlNumConceptos - 1).Num_IdenReceptor = stLiquidacion(vlNumBeneficiarios - 1).Num_IdenReceptor
                        stDetPension(vlNumConceptos - 1).Cod_TipoIdenReceptor = stLiquidacion(vlNumBeneficiarios - 1).Cod_TipoIdenReceptor
                        stDetPension(vlNumConceptos - 1).Cod_TipReceptor = stLiquidacion(vlNumBeneficiarios - 1).Cod_TipReceptor
                        stDetPension(vlNumConceptos - 1).Cod_ConHabDes = stDatGenerales.Cod_ConceptoDesctoSalud
                        stDetPension(vlNumConceptos - 1).Fec_IniPago = vlFecIniPag
                        stDetPension(vlNumConceptos - 1).Fec_TerPago = vlFecTerPag
                        stDetPension(vlNumConceptos - 1).Mto_ConHabDes = Format((stLiquidacion(vlNumBeneficiarios - 1).Mto_Pension * stDatGenerales.Prc_SaludMin / 100), "##0.00")
                        If stDetPension(vlNumConceptos - 1).Mto_ConHabDes > vlMaximoSalud Then
                            stDetPension(vlNumConceptos - 1).Mto_ConHabDes = vlMaximoSalud
                        End If
                        stDetPension(vlNumConceptos - 1).Num_Orden = stLiquidacion(vlNumBeneficiarios - 1).Num_Orden
                        stDetPension(vlNumConceptos - 1).Num_PerPago = stLiquidacion(vlNumBeneficiarios - 1).Num_PerPago
                        stDetPension(vlNumConceptos - 1).Num_Poliza = stLiquidacion(vlNumBeneficiarios - 1).Num_Poliza
                        stLiquidacion(vlNumBeneficiarios - 1).Mto_Salud = stDetPension(vlNumConceptos - 1).Mto_ConHabDes
                        stLiquidacion(vlNumBeneficiarios - 1).Mto_MaxSalud = vlMaximoSalud
                    End If
                End If
                                
                'Datos del Pago
                stLiquidacion(vlNumBeneficiarios - 1).Fec_IniPago = vlFecIniPag
                stLiquidacion(vlNumBeneficiarios - 1).Fec_TerPago = vlFecTerPag
                stLiquidacion(vlNumBeneficiarios - 1).Mto_BaseImp = stLiquidacion(vlNumBeneficiarios - 1).Mto_Pension
                stLiquidacion(vlNumBeneficiarios - 1).Mto_BaseTri = stLiquidacion(vlNumBeneficiarios - 1).Mto_BaseImp - stLiquidacion(vlNumBeneficiarios - 1).Mto_Salud
                stLiquidacion(vlNumBeneficiarios - 1).Mto_Descuento = stLiquidacion(vlNumBeneficiarios - 1).Mto_Salud
                stLiquidacion(vlNumBeneficiarios - 1).Mto_Haber = stLiquidacion(vlNumBeneficiarios - 1).Mto_Pension
                stLiquidacion(vlNumBeneficiarios - 1).Mto_LiqPagar = stLiquidacion(vlNumBeneficiarios - 1).Mto_BaseTri
                vlFecIniPag = Format(DateSerial(Mid(vlFecIniPag, 1, 4), Mid(vlFecIniPag, 5, 2) + 1, Mid(vlFecIniPag, 7, 2)), "yyyymmdd")
                'If vlTipRen <> "6" Then
                    vlFecTerPag = Format(DateSerial(Mid(vlFecIniPag, 1, 4), Mid(vlFecIniPag, 5, 2) + 1, 0), "yyyymmdd")
                'End If
                
                If iFecTerPag < vlFecTerPag And vlTermino = False Then
                    vlFecTerPag = iFecTerPag 'Para el último mes
                    vlTermino = True
                End If
                
            Loop
Siguiente:
            vlRegistro1.MoveNext
        Loop
    End If
    vlRegistro1.Close
    
    'Recalcular Pension de la Conyuge si tiene Derecho a Crecer
    If vlNumCrecer > 0 Then
        Dim k As Long
        Dim vlEncontro As Boolean
        Dim vlGratificacion As Double
        For i = 1 To vlNumCrecer
            k = 1
            For j = 1 To vlNumBeneficiarios - 1
                vlGratificacion = 0
                If stLiquidacion(j).Num_PerPago >= stCrecer(i).Num_PerPago And stLiquidacion(j).Num_Orden = stCrecer(i).Num_OrdenCon Then 'Si es el periodo y es la conyuge
                    stLiquidacion(j).Prc_Pension = (stLiquidacion(j).Prc_Pension + stCrecer(i).Prc_Pension)
                    stLiquidacion(j).Mto_Pension = Format(stLiquidacion(j).Mto_PensionTotal * (stLiquidacion(j).Prc_Pension / 100), "#0.00")
                    If stLiquidacion(j).Cod_DerGra = "S" Then
                        If (CLng(Mid(stLiquidacion(j).Num_PerPago, 5, 2)) Mod 6) = 0 Then
                            'Pagar Gratificacion
                            vlGratificacion = stLiquidacion(j).Mto_Pension
                            stLiquidacion(j).Mto_Pension = stLiquidacion(j).Mto_Pension + vlGratificacion
                        End If
                    End If
                    If stLiquidacion(j).Mto_Salud > 0 Then
                        'Se debe calcular Monto Salud
                        stLiquidacion(j).Mto_Salud = Format((stLiquidacion(j).Mto_Pension * stDatGenerales.Prc_SaludMin / 100), "##0.00")
                        If stLiquidacion(j).Mto_MaxSalud < stLiquidacion(j).Mto_Salud Then
                            stLiquidacion(j).Mto_Salud = stLiquidacion(j).Mto_MaxSalud
                        End If
                    End If
                    stLiquidacion(j).Mto_BaseImp = stLiquidacion(j).Mto_Pension
                    stLiquidacion(j).Mto_BaseTri = stLiquidacion(j).Mto_BaseImp - stLiquidacion(j).Mto_Salud
                    stLiquidacion(j).Mto_Descuento = stLiquidacion(j).Mto_Salud
                    stLiquidacion(j).Mto_Haber = stLiquidacion(j).Mto_Pension
                    stLiquidacion(j).Mto_LiqPagar = stLiquidacion(j).Mto_BaseTri
                    'Actualizar Detalle
                    vlEncontro = False
                    For l = k To vlNumConceptos
                        If stDetPension(l).Num_PerPago = stLiquidacion(j).Num_PerPago And stDetPension(l).Num_Orden = stLiquidacion(j).Num_Orden Then
                            vlEncontro = True
                            If stDetPension(l).Cod_ConHabDes = stDatGenerales.Cod_ConceptoPension Then
                                stDetPension(l).Mto_ConHabDes = stLiquidacion(j).Mto_Pension
                            End If
                            If stDetPension(l).Cod_ConHabDes = stDatGenerales.Cod_ConceptoDesctoSalud Then
                                stDetPension(l).Mto_ConHabDes = stLiquidacion(j).Mto_Salud
                            End If
                            If stDetPension(l).Cod_ConHabDes = stDatGenerales.Cod_ConceptoGratificacion Then
                                stDetPension(l).Mto_ConHabDes = vlGratificacion
                            End If
                            k = k + 1
                        Else
                            If vlEncontro Then 'Si estaba en el beneficiario y periodo y ahora ya no está, no se sigue con el loop
                                Exit For
                            Else
                                k = k + 1
                            End If
                        End If
                    Next l
                End If
            Next j
        Next i
    End If
    If vlNumBeneficiarios > 0 Then
        oExistenPagos = True
    End If
    fgCalcularPrimerPago = True
    Screen.MousePointer = vbDefault

Errores:
    If Err.Number <> 0 Then
Deshacer:
        If Err.Number <> 0 Then
            MsgBox "Error Grave [ " & Err & Space(4) & Err.Description & " ]", vbCritical, "¡ERROR!..."
        End If
    End If
    Screen.MousePointer = vbDefault
End Function

Function fgCalculaEdad(iFecNac, iFecIniPag) As Integer
'Calcula Edad del Pensionado a la Fecha de Pago
On Error GoTo Errores
Dim vlEdad As Integer

fgCalculaEdad = "-1"

vlEdad = DateDiff("m", DateSerial(Mid(iFecNac, 1, 4), Mid(iFecNac, 5, 2), Mid(iFecNac, 7, 2)), DateSerial(Mid(iFecIniPag, 1, 4), Mid(iFecIniPag, 5, 2), Mid(iFecIniPag, 7, 2)) - 1) 'Edad del Beneficiario
If vlEdad >= 0 Then
    fgCalculaEdad = vlEdad
Else
    fgCalculaEdad = "-2"
End If

Errores:
    If Err.Number <> 0 Then
        MsgBox "Error al Calcular Edad del Pensionado" & Chr(13) & Err.Description, vbCritical, "Error"
    End If
End Function

Function fgConvierteEdadAños(iEdad)
'Convierte la Edad Calculada en Meses a Años

fgConvierteEdadAños = (iEdad \ 12)

End Function



Function fgObtieneParametroVigencia(iTabla, iElemento, iVigencia, oValor) As Boolean
'Obtiene Parámetros Generales de Tabla de Vigencias MA_TPAR_TABCODVIG
On Error GoTo Errores
Dim vlSql As String
Dim vlTb As ADODB.Recordset

fgObtieneParametroVigencia = False
vlSql = "SELECT MTO_ELEMENTO"
vlSql = vlSql & " FROM MA_TPAR_TABCODVIG"
vlSql = vlSql & " WHERE COD_TABLA = '" & iTabla & "'"
vlSql = vlSql & " AND COD_ELEMENTO = '" & iElemento & "'"
vlSql = vlSql & " AND FEC_INIVIG <= '" & iVigencia & "'"
vlSql = vlSql & " AND FEC_TERVIG >= '" & iVigencia & "'"
Set vlTb = vgConexionBD.Execute(vlSql)
If Not vlTb.EOF Then
    oValor = vlTb!mto_elemento
Else
    vlTb.Close
    Exit Function
End If
vlTb.Close
fgObtieneParametroVigencia = True

Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al Obtener Parámetros Generales de Vigencia" & Chr(13) & Err.Description, vbCritical, "Error al Obtener Datos Generales"
End If

End Function

Function fgObtieneMontoMaximoSalud(iMoneda, iFecha, oValor) As Boolean
'Obtiene Monto Máximo de Salud

On Error GoTo Errores
Dim vlSql As String
Dim vlTb As ADODB.Recordset

fgObtieneMontoMaximoSalud = False
vlSql = "SELECT a.mto_elemento"
vlSql = vlSql & " FROM ma_tval_saludmax a"
vlSql = vlSql & " WHERE cod_moneda = '" & iMoneda & "'"
vlSql = vlSql & " AND fec_inivig <= '" & iFecha & "'"
vlSql = vlSql & " AND fec_tervig >= '" & iFecha & "'"
Set vlTb = vgConexionBD.Execute(vlSql)
If Not vlTb.EOF Then
    oValor = vlTb!mto_elemento
Else
    vlTb.Close
    MsgBox "No está definido Monto Máximo de Salud para la Moneda de la Pensión", vbCritical
    Exit Function
End If
vlTb.Close
fgObtieneMontoMaximoSalud = True

Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al Obtener Monto Máximo de Salud" & Chr(13) & Err.Description, vbCritical, "Error al Obtener Datos Generales"
End If

End Function


Function fgObtieneFactorAjuste(iFecha, oValor) As Boolean
'Obtiene Factor IPC
On Error GoTo Errores
Dim vlSql As String
Dim vlTb As ADODB.Recordset

fgObtieneFactorAjuste = False
vlSql = "SELECT a.mto_ipc"
vlSql = vlSql & " FROM ma_tval_ipc a"
vlSql = vlSql & " WHERE a.fec_ipc = '" & iFecha & "'"
Set vlTb = vgConexionBD.Execute(vlSql)
If Not vlTb.EOF Then
    oValor = vlTb!mto_ipc
Else
    vlTb.Close
    MsgBox "No existe Factor de Variación para la Fecha : " & DateSerial(Mid(iFecha, 1, 4), Mid(iFecha, 5, 2), Mid(iFecha, 7, 2)), vbCritical, "Faltan Datos"
    Exit Function
End If
vlTb.Close
fgObtieneFactorAjuste = True

Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al Obtener Factor de Variación Pensión" & Chr(13) & Err.Description, vbCritical, "Error al Obtener Datos Generales"
End If

End Function


Function fgObtieneFactorAjusteIPCSolInd(iFecha, bValQ, oValor) As Boolean
'Obtiene Factor IPC
On Error GoTo Errores
Dim vlSql As String
Dim vlTb As ADODB.Recordset

fgObtieneFactorAjusteIPCSolInd = False

'If bValQ = False Then
'    vlSql = "SELECT MTO_FACTOR FROM MA_TVAL_VALVAC WHERE '" & iFecha & "' BETWEEN FEC_INICUOMOR AND FEC_TERCUOMOR"
'    Set vlTb = vgConexionBD.Execute(vlSql)
'    If Not vlTb.EOF Then
'            oValor = vlTb!MTO_FACTOR
'    Else
'        vlTb.Close
'        MsgBox "No existe Factor de Variación para la Fecha : " & DateSerial(Mid(iFecha, 1, 4), Mid(iFecha, 5, 2), Mid(iFecha, 7, 2)), vbCritical, "Faltan Datos"
'        Exit Function
'    End If
'Else
    vlSql = "SELECT MTO_IPCMEN FROM MA_TVAL_IPCVAC WHERE FEC_VIGIPC='" & iFecha & "'"
    Set vlTb = vgConexionBD.Execute(vlSql)
    If Not vlTb.EOF Then
            oValor = vlTb!MTO_IPCMEN
    Else
        vlTb.Close
        MsgBox "No existe Factor de Variación para la Fecha : " & DateSerial(Mid(iFecha, 1, 4), Mid(iFecha, 5, 2), Mid(iFecha, 7, 2)), vbCritical, "Faltan Datos"
        Exit Function
    End If
'End If

'vlTb.Close
fgObtieneFactorAjusteIPCSolInd = True

Errores:
If Err.Number <> 0 Then
    MsgBox "Se ha producido un Error al Obtener Factor de Variación Pensión" & Chr(13) & Err.Description, vbCritical, "Error al Obtener Datos Generales"
End If

End Function

Function fgConvierteMontoenPalabras(iMonto As Double)
'Convierte Monto en palabras
Dim vlMontoPalabras As String


End Function

Function fgConvierteDigito(iDigito As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Select Case iDigito
    Case 1
        vlMontoPalabras = "uno"
    Case 2
        vlMontoPalabras = "dos"
    Case 3
        vlMontoPalabras = "tres"
    Case 4
        vlMontoPalabras = "cuatro"
    Case 5
        vlMontoPalabras = "cinco"
    Case 6
        vlMontoPalabras = "seis"
    Case 7
        vlMontoPalabras = "siete"
    Case 8
        vlMontoPalabras = "ocho"
    Case 9
        vlMontoPalabras = "nueve"
    Case 10
        vlMontoPalabras = "diez"
    Case 11
        vlMontoPalabras = "once"
    Case 12
        vlMontoPalabras = "doce"
    Case 13
        vlMontoPalabras = "trece"
    Case 14
        vlMontoPalabras = "catorce"
    Case 15
        vlMontoPalabras = "quince"
    Case 16
        vlMontoPalabras = "dieciseis"
    Case 17
        vlMontoPalabras = "diecisiete"
    Case 18
        vlMontoPalabras = "dieciocho"
    Case 19
        vlMontoPalabras = "diecinueve"
    Case 20
        vlMontoPalabras = "veinte"
End Select
fgConvierteDigito = vlMontoPalabras
End Function


Function fgConvierteDecenas(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlDigito As Double
vlDigito = iMonto Mod 10
Select Case iMonto
    Case Is > 90
        vlMontoPalabras = "noventa y " + fgConvierteDigito(vlDigito)
    Case 90
        vlMontoPalabras = "noventa"
    Case Is > 80
        vlMontoPalabras = "ochenta y " + fgConvierteDigito(vlDigito)
    Case 80
        vlMontoPalabras = "ochenta"
    Case Is > 70
        vlMontoPalabras = "setenta y " + fgConvierteDigito(vlDigito)
    Case 70
        vlMontoPalabras = "setenta"
    Case Is > 60
        vlMontoPalabras = "sesenta y " + fgConvierteDigito(vlDigito)
    Case 60
        vlMontoPalabras = "sesenta"
    Case Is > 50
        vlMontoPalabras = "cincuenta y " + fgConvierteDigito(vlDigito)
    Case 50
        vlMontoPalabras = "cincuenta"
    Case Is > 40
        vlMontoPalabras = "cuarenta y " + fgConvierteDigito(vlDigito)
    Case 40
        vlMontoPalabras = "cuarenta"
    Case Is > 30
        vlMontoPalabras = "treinta y " + fgConvierteDigito(vlDigito)
    Case 30
        vlMontoPalabras = "treinta"
    Case Is > 20
        vlMontoPalabras = "veinti" + fgConvierteDigito(vlDigito)
    Case Is <= 20
        vlMontoPalabras = fgConvierteDigito(iMonto)
End Select
fgConvierteDecenas = vlMontoPalabras
End Function

Function fgConvierteCentenas(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlDecena As Double
vlDecena = iMonto Mod 100
Select Case iMonto
    Case Is > 900
        vlMontoPalabras = "novecientos " + fgConvierteDecenas(vlDecena)
    Case 900
        vlMontoPalabras = "novecientos"
    Case Is > 800
        vlMontoPalabras = "ochocientos " + fgConvierteDecenas(vlDecena)
    Case 800
        vlMontoPalabras = "ochocientos"
    Case Is > 700
        vlMontoPalabras = "setecientos " + fgConvierteDecenas(vlDecena)
    Case 700
        vlMontoPalabras = "setecientos"
    Case Is > 600
        vlMontoPalabras = "seiscientos " + fgConvierteDecenas(vlDecena)
    Case 600
        vlMontoPalabras = "seiscientos"
    Case Is > 500
        vlMontoPalabras = "quinientos " + fgConvierteDecenas(vlDecena)
    Case 500
        vlMontoPalabras = "quinientos"
    Case Is > 400
        vlMontoPalabras = "cuatrocientos " + fgConvierteDecenas(vlDecena)
    Case 400
        vlMontoPalabras = "cuatrocientos"
    Case Is > 300
        vlMontoPalabras = "trescientos " + fgConvierteDecenas(vlDecena)
    Case 300
        vlMontoPalabras = "trescientos"
    Case Is > 200
        vlMontoPalabras = "doscientos " + fgConvierteDecenas(vlDecena)
    Case 200
        vlMontoPalabras = "doscientos"
    Case Is > 100
        vlMontoPalabras = "ciento " + fgConvierteDecenas(vlDecena)
    Case 100
        vlMontoPalabras = "cien"
    Case Is < 100
        vlMontoPalabras = fgConvierteDecenas(iMonto)
End Select
fgConvierteCentenas = vlMontoPalabras
End Function


Function fgConvierteMiles(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlFraccion As Double
Dim vlCantidad As Double
vlFraccion = iMonto Mod 1000
vlCantidad = Int(iMonto / 1000) 'Parte Entera
Select Case iMonto
    Case 1000
        vlMontoPalabras = "mil"
    Case Is > 999
        vlMontoPalabras = Trim(fgConvierteCentenas(vlCantidad) + " mil " + fgConvierteCentenas(vlFraccion))
    Case Else
        vlMontoPalabras = fgConvierteCentenas(iMonto)
End Select
fgConvierteMiles = vlMontoPalabras
End Function

Function fgConvierteMillones(iMonto As Double) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlFraccion As Double
Dim vlCantidad As Double
vlFraccion = iMonto Mod 1000000
vlCantidad = Int(iMonto / 1000000) 'Parte Entera
Select Case iMonto
    Case 1000000
        vlMontoPalabras = "un millón"
    Case Is > 999999
        vlMontoPalabras = Trim(fgConvierteMiles(vlCantidad) + " millones " + fgConvierteMiles(vlFraccion))
    Case Else
        vlMontoPalabras = fgConvierteMiles(iMonto)
End Select
fgConvierteMillones = vlMontoPalabras
End Function

Function fgConvierteNumeroLetras(iMonto As Double, Optional iMoneda As String) As String
'Convierte Monto en palabras
Dim vlMontoPalabras As String
Dim vlDecimales As Double
Dim vlEntero As Double
vlEntero = Fix(iMonto)
vlDecimales = Format((iMonto - vlEntero) * 100, "#0.00")
vlMontoPalabras = fgConvierteMillones(vlEntero)
If iMonto > 2 Then
    If Mid(vlMontoPalabras, 1, 3) = "uno" Then
        vlMontoPalabras = Mid(vlMontoPalabras, 1, 2) + Mid(vlMontoPalabras, 4)
    End If
End If
If vlDecimales > 0 Then
    vlMontoPalabras = vlMontoPalabras + " con " & vlDecimales & "/100"
End If
If Not IsMissing(iMoneda) Then
    vlMontoPalabras = vlMontoPalabras + " " & iMoneda
End If
fgConvierteNumeroLetras = UCase(vlMontoPalabras)

End Function


Function ValidaMenorEdad(fechaNac As Date) As Boolean

Dim fechaLimite As Date
fechaLimite = DateAdd("YYYY", -18, ObtenerFechaServer)
If fechaNac < fechaLimite Then
    ValidaMenorEdad = False
    Exit Function
End If
ValidaMenorEdad = True

End Function
