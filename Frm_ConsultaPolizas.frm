VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ConsultaPolizas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plantilla de Polizas por meses"
   ClientHeight    =   1905
   ClientLeft      =   6855
   ClientTop       =   3465
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Polizas"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtPoliza 
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "A Excel"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame frm_Modalidad 
      Caption         =   "Fecha Calculo"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   6855
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   5040
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Año/Mes desde"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Año/mes Hasta"
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Frm_ConsultaPolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()
Dim xlapp As Excel.Application
Dim sArchivo As String
Dim ix As Long
Dim Sql As String
Dim vlRegistro As ADODB.Recordset

    Screen.MousePointer = 11
    
    Set xlapp = CreateObject("excel.application")

    sArchivo = App.Path & "\Plantilla_polizas.xls"

    xlapp.Visible = True 'para ver vista previa
    xlapp.WindowState = 2 ' minimiza excel
    xlapp.Workbooks.Open (sArchivo)

    ix = 2
    
'    sql = ""
'    sql = sql & "select a.num_cot AS COTIZACION, a.num_correlativo, num_orden, i.gls_elemento parentesco, to_date(fec_nacben,'YYYYMMDD') as FecNac,"
'    sql = sql & " to_date(g.fec_dev,'YYYYMMDD') as fec_dev,"
'    sql = sql & " g.cod_cuspp, a.num_operacion as solicitud , d.cod_scomp as MonedaAju, h.gls_elemento as Tipopension, b.gls_elemento as tipo_renta,"
'    sql = sql & " c.gls_elemento as Modalidad, a.num_mesdif / 12  as mesdif, a.prc_rentatmp, a.num_mesgar / 12 mesgar,"
'    sql = sql & " a.mto_priunimod,   a.PRC_TASATIR as tir , a.PRC_PERCON as perdida, a.PRC_CORCOMREAL as comision,"
'    sql = sql & " a.PRC_TASAVTA as tasa, SUBSTR(a.FEC_CALCULO,1,6) as mes, to_date(g.fec_cierre, 'YYYYMMDD') as fechacierre, 'N' as gana"
'    sql = sql & " from pt_tmae_detcotizacion a"
'    sql = sql & " join ma_tpar_tabcod b on b.cod_elemento=a.cod_tipren and b.cod_tabla='TR'"
'    sql = sql & " join ma_tpar_tabcod c on c.cod_elemento=a.cod_modalidad and c.cod_tabla='AL'"
'    sql = sql & " JOIN MA_TPAR_MONEDATIPOREAJU d on a.cod_moneda=d.cod_moneda and a.cod_tipreajuste=d.cod_tipreajuste"
'    sql = sql & " join PT_TMAE_COTIZACION e on a.num_cot=e.num_cot"
'    sql = sql & " join PT_TMAE_COTBEN f on a.num_cot=f.num_cot"
'    sql = sql & " join pt_tmae_cotizacion g on a.num_cot=g.num_cot"
'    sql = sql & " join ma_tpar_tabcod h on h.cod_elemento=g.cod_tippension and h.cod_tabla='TP'"
'    sql = sql & " join ma_tpar_tabcod i on i.cod_elemento=f.cod_par and i.cod_tabla='PA'"
'    sql = sql & " full join pd_tmae_poliza j on a.num_cot=j.num_cot"
'    sql = sql & " where SUBSTR(a.FEC_CALCULO,1,6) in ('201303') AND a.COD_ESTCOT IN ('E','P','S','C')"
'    sql = sql & " and a.num_cot is not null"
'    sql = sql & " Union All"
'    sql = sql & " select a.num_cot AS COTIZACION, a.num_correlativo, num_orden, i.gls_elemento parentesco, to_date(fec_nacben,'YYYYMMDD') as FecNac,"
'    sql = sql & " to_date(g.fec_dev,'YYYYMMDD') as fec_dev,"
'    sql = sql & " g.cod_cuspp, a.num_operacion as solicitud , d.cod_scomp as MonedaAju, h.gls_elemento as Tipopension, b.gls_elemento as tipo_renta,"
'    sql = sql & " c.gls_elemento as Modalidad, a.num_mesdif / 12  as mesdif, a.prc_rentatmp, a.num_mesgar / 12 mesgar,"
'    sql = sql & " a.mto_priunimod,   a.PRC_TASATIR as tir , a.PRC_PERCON as perdida, a.PRC_CORCOMREAL as comision,"
'    sql = sql & " a.PRC_TASAVTA as tasa, SUBSTR(a.FEC_CALCULO,1,6) as mes, to_date(g.fec_cierre, 'YYYYMMDD') as fechacierre, 'G' as gana"
'    sql = sql & " from pt_tmae_detcotizacion a"
'    sql = sql & " join ma_tpar_tabcod b on b.cod_elemento=a.cod_tipren and b.cod_tabla='TR'"
'    sql = sql & " join ma_tpar_tabcod c on c.cod_elemento=a.cod_modalidad and c.cod_tabla='AL'"
'    sql = sql & " JOIN MA_TPAR_MONEDATIPOREAJU d on a.cod_moneda=d.cod_moneda and a.cod_tipreajuste=d.cod_tipreajuste"
'    sql = sql & " join PT_TMAE_COTIZACION e on a.num_cot=e.num_cot"
'    sql = sql & " join PT_TMAE_COTBEN f on a.num_cot=f.num_cot"
'    sql = sql & " join pt_tmae_cotizacion g on a.num_cot=g.num_cot"
'    sql = sql & " join ma_tpar_tabcod h on h.cod_elemento=g.cod_tippension and h.cod_tabla='TP'"
'    sql = sql & " join ma_tpar_tabcod i on i.cod_elemento=f.cod_par and i.cod_tabla='PA'"
'    sql = sql & " join pd_tmae_poliza j on a.num_cot=j.num_cot"
'    sql = sql & " where SUBSTR(a.FEC_CALCULO,1,6) in ('201303') AND a.COD_ESTCOT IN ('E','P','S','C')"
'    sql = sql & " order by 1, 2, 3"
    
    
    
'    Sql = Sql & "select a.num_poliza, b.num_orden, f.gls_elemento as ROL, cod_cuspp, to_date(c.fec_calculo,'YYYYMMDD') as fec_calculo, to_date(fec_cierre,'YYYYMMDD') as fec_cierre , d.num_cot, cod_afp, e.gls_elemento as AFP,"
'    Sql = Sql & " b.gls_nomben || ' ' || gls_nomsegben || ' ' || gls_patben || ' ' || gls_matben as nombres, cod_sexo,"
'    Sql = Sql & " to_date(fec_nacben,'YYYYMMDD') as fec_nacben, to_date(fec_fallben,'YYYYMMDD') as fec_fallben, g.gls_elemento as TIPOPEN, h.gls_elemento as modalidad, to_date(fec_dev,'YYYYMMDD') as fec_dev,"
'    Sql = Sql & " num_mesdif, num_mesgar, b.cod_dercre, cod_dergra, num_idenreceptor, i.cod_par, k.mto_priinf, a.cod_moneda,"
'    Sql = Sql & " j.mto_valmoneda  as TCCotiza, j.mto_pension, to_date(fec_solicitud,'YYYYMMDD') SOLICITUD, prc_tasavta, to_date(c.fec_acepta,'YYYYMMDD') as fec_acepta, a.mto_valmoneda as TCTrasAFP,"
'    Sql = Sql & " mto_pritotal as PRIRECAL, mto_pricia, b.mto_pension as PenEmitida, to_date(c.fec_emision,'YYYYMMDD') as fec_emision, to_date(b.fec_inipagopen,'YYYYMMDD') as fec_inipagopen, to_date(fec_traspaso,'YYYYMMDD') as fec_traspaso,"
'    Sql = Sql & " num_idencor, l.gls_nomcor || ' ' || gls_patcor || ' ' || gls_matcor as nomasesor, c.prc_corcomreal, to_date(a.fec_devsol,'YYYYMMDD') as fec_devsol "
'    Sql = Sql & " from pp_tmae_poliza a"
'    Sql = Sql & " join pp_tmae_ben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
'    Sql = Sql & " join pd_tmae_poliza c on a.num_poliza=c.num_poliza"
'    Sql = Sql & " join pt_tmae_cotizacion d on c.num_cot=d.num_cot"
'    Sql = Sql & " join pt_tmae_detcotizacion j on d.num_cot= j.num_cot"
'    Sql = Sql & " join ma_tpar_tabcod e on a.cod_afp=e.cod_elemento and e.cod_tabla='AF'"
'    Sql = Sql & " join ma_tpar_tabcod f on b.cod_par=f.cod_elemento and f.cod_tabla='PA'"
'    Sql = Sql & " join ma_tpar_tabcod g on a.cod_tippension=g.cod_elemento and g.cod_tabla='TP'"
'    Sql = Sql & " join ma_tpar_tabcod h on a.cod_modalidad=h.cod_elemento and h.cod_tabla='AL'"
'    Sql = Sql & " left join (select a.num_poliza, num_idenreceptor, max(fec_pago), cod_par"
'    Sql = Sql & " from pp_tmae_liqpagopendef a join pp_tmae_ben b on b.num_idenben=a.num_idenreceptor"
'    Sql = Sql & " group by a.num_poliza, num_idenreceptor,cod_par ) i on b.num_idenben=i.num_idenreceptor"
'    Sql = Sql & " join pd_tmae_polprirec k on a.num_poliza=k.num_poliza"
'    Sql = Sql & " join pt_tmae_corredor l on c.num_idencor=l.num_idencor"
'    Sql = Sql & " where a.num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=a.num_poliza)"
'    Sql = Sql & " and a.num_poliza >= " & txtPoliza & ""
'    Sql = Sql & " order by 1,2"

'    Sql = Sql & "select distinct a.num_poliza, b.num_orden, f.gls_elemento as ROL, cod_cuspp, to_date(c.fec_calculo,'YYYYMMDD') as fec_calculo, to_date(fec_cierre,'YYYYMMDD') as fec_cierre ,"
'    Sql = Sql & " d.num_cot, cod_afp, e.gls_elemento as AFP, b.gls_nomben || ' ' || gls_nomsegben || ' ' || gls_patben || ' ' || gls_matben as nombres, cod_sexo,"
'    Sql = Sql & " to_date(fec_nacben,'YYYYMMDD') as fec_nacben, to_date(fec_fallben,'YYYYMMDD') as fec_fallben, g.gls_elemento as TIPOPEN, h.gls_elemento as modalidad,"
'    Sql = Sql & " to_date(fec_dev,'YYYYMMDD') as fec_dev, num_mesdif, num_mesgar, b.cod_dercre, cod_dergra, num_idenreceptor, i.cod_par, k.mto_priinf,"
'    Sql = Sql & " a.cod_moneda, j.mto_valmoneda  as TCCotiza, j.mto_pension, to_date(fec_solicitud,'YYYYMMDD') SOLICITUD, prc_tasavta, to_date(c.fec_acepta,'YYYYMMDD') as fec_acepta,"
'    Sql = Sql & " a.mto_valmoneda as TCTrasAFP, mto_pritotal as PRIRECAL, mto_pricia, b.mto_pension as PenEmitida, to_date(c.fec_emision,'YYYYMMDD') as fec_emision,"
'    Sql = Sql & " to_date(b.fec_inipagopen,'YYYYMMDD') as fec_inipagopen, to_date(fec_traspaso,'YYYYMMDD') as fec_traspaso, num_idencor,"
'    Sql = Sql & " l.gls_nomcor || ' ' || gls_patcor || ' ' || gls_matcor as nomasesor, c.prc_corcomreal, to_date(a.fec_devsol,'YYYYMMDD') as fec_devsol, ind_cob,"
'    Sql = Sql & " (select mto_ipc from ma_tval_ipc where fec_ipc = substr(c.fec_calculo,1,6) || '01') as IPC,"
'    Sql = Sql & " (select prc_mes4 from ma_tval_tasatm where num_anno = substr(c.fec_calculo,1,4) and cod_moneda=a.cod_moneda and cod_tipreajuste=a.cod_tipreajuste) as TasaMer,"
'    Sql = Sql & " b.Prc_Pension , b.Mto_Pension, b.Prc_PensionGar, b.Mto_PensionGar, a.Prc_TasaCe, b.Cod_SitInv, a.Cod_CoberCon,"
'    Sql = Sql & " case when a.cod_moneda='NS' then (cb1.MTO_CNTBAS+cb1.MTO_CNGBAS) else (cb2.MTO_CNTBAS+cb2.MTO_CNGBAS) end AS RMB,"
'    Sql = Sql & " case when a.cod_moneda='NS' then (cb1.MTO_CNTFIN+cb1.MTO_CNGFIN) else (cb2.MTO_CNTFIN+cb2.MTO_CNGFIN) end AS RMF,  c.prc_tasatir, c.prc_percon , l.gls_nomcor || ' ' ||  l.gls_nomsegcor || ' ' ||  l.gls_patcor || ' ' ||  l.gls_matcor as supervisor"
'    Sql = Sql & " from pp_tmae_poliza a"
'    Sql = Sql & " join pp_tmae_ben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
'    Sql = Sql & " join pd_tmae_poliza c on a.num_poliza=c.num_poliza"
'    Sql = Sql & " join pt_tmae_cotizacion d on c.num_cot=d.num_cot"
'    Sql = Sql & " join pt_tmae_detcotizacion j on d.num_cot= j.num_cot"
'    Sql = Sql & " join ma_tpar_tabcod e on a.cod_afp=e.cod_elemento and e.cod_tabla='AF'"
'    Sql = Sql & " join ma_tpar_tabcod f on b.cod_par=f.cod_elemento and f.cod_tabla='PA'"
'    Sql = Sql & " join ma_tpar_tabcod g on a.cod_tippension=g.cod_elemento and g.cod_tabla='TP'"
'    Sql = Sql & " join ma_tpar_tabcod h on a.cod_modalidad=h.cod_elemento and h.cod_tabla='AL'"
'    Sql = Sql & " left join (select a.num_poliza, num_idenreceptor, max(fec_pago), cod_par from pp_tmae_liqpagopendef a join pp_tmae_ben b on b.num_idenben=a.num_idenreceptor group by a.num_poliza, num_idenreceptor, cod_par ) i on b.num_idenben=i.num_idenreceptor"
'    Sql = Sql & " join pd_tmae_polprirec k on a.num_poliza=k.num_poliza"
'    Sql = Sql & " join pt_tmae_corredor l on c.num_idencor=l.num_idencor"
'    Sql = Sql & " left join pr_tmae_calben1 cb1 on b.num_poliza=cb1.num_poliza and b.num_orden=cb1.num_orden"
'    Sql = Sql & " left join pr_tmae_calben2 cb2 on b.num_poliza=cb2.num_poliza and b.num_orden=cb2.num_orden"
'    Sql = Sql & " where a.num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=a.num_poliza)"
'    Sql = Sql & " and a.num_poliza >=  " & txtPoliza & ""
'    Sql = Sql & " order by 1,2"
    
    
'    Sql = Sql & " select distinct a.num_poliza, b.num_orden, f.gls_elemento as ROL, cod_cuspp, to_date(c.fec_calculo,'YYYYMMDD') as fec_calculo, to_date(fec_cierre,'YYYYMMDD') as fec_cierre , d.num_cot, cod_afp,"
'    Sql = Sql & " e.gls_elemento as AFP, b.num_idenben dniben, b.gls_nomben || ' ' || gls_nomsegben || ' ' || gls_patben || ' ' || gls_matben as nombres, cod_sexo, to_date(fec_nacben,'YYYYMMDD') as fec_nacben, to_date(fec_fallben,'YYYYMMDD') as fec_fallben,"
'    Sql = Sql & " g.gls_elemento as TIPOPEN, h.gls_elemento as modalidad, to_date(fec_dev,'YYYYMMDD') as fec_dev, num_mesdif, num_mesgar, b.cod_dercre, cod_dergra, num_idenreceptor, i.cod_par, k.mto_priinf, a.cod_moneda, j.mto_valmoneda  as TCCotiza,"
'    Sql = Sql & " j.mto_pension, to_date(fec_solicitud,'YYYYMMDD') SOLICITUD, prc_tasavta, to_date(c.fec_acepta,'YYYYMMDD') as fec_acepta, a.mto_valmoneda as TCTrasAFP, mto_pritotal as PRIRECAL, mto_pricia, b.mto_pension as PenEmitida,"
'    Sql = Sql & " to_date(c.fec_emision,'YYYYMMDD') as fec_emision, to_date(b.fec_inipagopen,'YYYYMMDD') as fec_inipagopen, to_date(fec_traspaso,'YYYYMMDD') as fec_traspaso, l.num_idencor,"
'    Sql = Sql & " l.gls_nomcor || ' ' || l.gls_nomsegcor || ' ' || l.gls_patcor || ' ' || l.gls_matcor as nomasesor, c.prc_corcomreal, to_date(a.fec_devsol,'YYYYMMDD') as fec_devsol, ind_cob,"
'    Sql = Sql & " (select mto_ipc from ma_tval_ipc where fec_ipc = substr(c.fec_calculo,1,6) || '01') as IPC, (select prc_mes4 from ma_tval_tasatm where num_anno = substr(c.fec_calculo,1,4) and cod_moneda=a.cod_moneda and cod_tipreajuste=a.cod_tipreajuste) as TasaMer,"
'    Sql = Sql & " b.Prc_Pension , b.Mto_Pension, b.Prc_PensionGar, b.Mto_PensionGar, a.Prc_TasaCe, b.Cod_SitInv, a.Cod_CoberCon, case when a.cod_moneda='NS' then (cb1.MTO_CNTBAS+cb1.MTO_CNGBAS) else (cb2.MTO_CNTBAS+cb2.MTO_CNGBAS) end AS RMB, case when a.cod_moneda='NS' then (cb1.MTO_CNTFIN+cb1.MTO_CNGFIN) else (cb2.MTO_CNTFIN+cb2.MTO_CNGFIN) end AS RMF,"
'    Sql = Sql & " c.prc_tasatir, c.prc_percon , ls.gls_nomcor || ' ' ||  ls.gls_nomsegcor || ' ' ||  ls.gls_patcor || ' ' ||  ls.gls_matcor as supervisor"
'    Sql = Sql & " from pp_tmae_poliza a join pp_tmae_ben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
'    Sql = Sql & " join pd_tmae_poliza c on a.num_poliza=c.num_poliza join pt_tmae_cotizacion d on c.num_cot=d.num_cot"
'    Sql = Sql & " join pt_tmae_detcotizacion j on d.num_cot= j.num_cot"
'    Sql = Sql & " join ma_tpar_tabcod e on a.cod_afp=e.cod_elemento and e.cod_tabla='AF'"
'    Sql = Sql & " join ma_tpar_tabcod f on b.cod_par=f.cod_elemento and f.cod_tabla='PA'"
'    Sql = Sql & " join ma_tpar_tabcod g on a.cod_tippension=g.cod_elemento and g.cod_tabla='TP'"
'    Sql = Sql & " join ma_tpar_tabcod h on a.cod_modalidad=h.cod_elemento and h.cod_tabla='AL'"
'    Sql = Sql & " left join (select a.num_poliza, num_idenreceptor, max(fec_pago), cod_par from pp_tmae_liqpagopendef a join pp_tmae_ben b on b.num_idenben=a.num_idenreceptor group by a.num_poliza, num_idenreceptor, cod_par ) i on b.num_idenben=i.num_idenreceptor"
'    Sql = Sql & " join pd_tmae_polprirec k on a.num_poliza=k.num_poliza"
'    Sql = Sql & " join pt_tmae_corredor l on c.num_idencor=l.num_idencor"
'    Sql = Sql & " join pt_tmae_corredor ls on l.num_idenjefe=ls.num_idencor"
'    Sql = Sql & " left join pr_tmae_calben1 cb1 on b.num_poliza=cb1.num_poliza and b.num_orden=cb1.num_orden"
'    Sql = Sql & " left join pr_tmae_calben2 cb2 on b.num_poliza=cb2.num_poliza and b.num_orden=cb2.num_orden"
'    Sql = Sql & " where a.num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=a.num_poliza) and a.num_poliza >=  " & txtPoliza & " order by 1,2"
'
    Sql = ""
    Sql = Sql & " select distinct a.num_poliza,"
    Sql = Sql & " b.num_orden,"
    Sql = Sql & " f.gls_elemento as ROL,"
    Sql = Sql & " c.cod_cuspp,"
    Sql = Sql & " to_date(c.fec_calculo,'YYYYMMDD') as fec_calculo,"
    Sql = Sql & " to_date(fec_cierre,'YYYYMMDD') as fec_cierre ,"
    Sql = Sql & " d.num_cot,"
    Sql = Sql & " c.cod_afp,"
    Sql = Sql & " e.gls_elemento as AFP,"
    Sql = Sql & " b.num_idenben dniben, b.gls_nomben || ' ' || b.gls_nomsegben || ' ' || b.gls_patben || ' ' || b.gls_matben as nombres,"
    Sql = Sql & " b.cod_sexo,"
    Sql = Sql & " to_date(b.fec_nacben,'YYYYMMDD') as fec_nacben,"
    Sql = Sql & " to_date(b.fec_fallben,'YYYYMMDD') as fec_fallben,"
    Sql = Sql & " g.gls_elemento as TIPOPEN,"
    Sql = Sql & " h.gls_elemento as modalidad,"
    Sql = Sql & " to_date(c.fec_dev,'YYYYMMDD') as fec_dev,"
    Sql = Sql & " c.num_mesdif,"
    Sql = Sql & " c.num_mesgar,"
    Sql = Sql & " b.cod_dercre,"
    Sql = Sql & " c.cod_dergra,"
    Sql = Sql & " i.num_idenreceptor,"
    Sql = Sql & " i.cod_par,"
    Sql = Sql & " k.mto_priinf,"
    Sql = Sql & " a.cod_moneda,"
    Sql = Sql & " j.mto_valmoneda  as TCCotiza,"
    Sql = Sql & " j.mto_pension,"
    Sql = Sql & " to_date(c.fec_solicitud,'YYYYMMDD') SOLICITUD,"
    Sql = Sql & " c.prc_tasavta,"
    Sql = Sql & " to_date(c.fec_acepta,'YYYYMMDD') as fec_acepta,"
    Sql = Sql & " a.mto_valmoneda as TCTrasAFP,"
    Sql = Sql & " k.mto_pritotal as PRIRECAL,"
    Sql = Sql & " k.mto_pricia,"
    Sql = Sql & " b.mto_pension as PenEmitida,"
    Sql = Sql & " to_date(c.fec_emision,'YYYYMMDD') as fec_emision,"
    Sql = Sql & " to_date(b.fec_inipagopen,'YYYYMMDD') as fec_inipagopen,"
    Sql = Sql & " to_date(k.fec_traspaso,'YYYYMMDD') as fec_traspaso,"
    Sql = Sql & " l.num_idencor,"
    Sql = Sql & " l.gls_nomcor || ' ' || l.gls_nomsegcor || ' ' || l.gls_patcor || ' ' || l.gls_matcor as nomasesor,"
    Sql = Sql & " c.prc_corcomreal,"
    Sql = Sql & " to_date(a.fec_devsol,'YYYYMMDD') as fec_devsol,"
    Sql = Sql & " c.ind_cob,"
    Sql = Sql & " (select mto_ipc from ma_tval_ipc where fec_ipc = substr(c.fec_calculo,1,6) || '01') as IPC,"
    Sql = Sql & " (select prc_mes4 from ma_tval_tasatm where num_anno = substr(c.fec_calculo,1,4) and cod_moneda=a.cod_moneda and cod_tipreajuste=a.cod_tipreajuste) as TasaMer,"
    Sql = Sql & " b.Prc_Pension,"
    Sql = Sql & " b.Mto_Pension,"
    Sql = Sql & " b.Prc_PensionGar,"
    Sql = Sql & " b.Mto_PensionGar,"
    Sql = Sql & " a.Prc_TasaCe,"
    Sql = Sql & " b.Cod_SitInv,"
    Sql = Sql & " a.Cod_CoberCon,"
    Sql = Sql & " case when a.cod_moneda='NS' then (cb1.MTO_CNTBAS+cb1.MTO_CNGBAS) else (cb2.MTO_CNTBAS+cb2.MTO_CNGBAS) end AS RMB,"
    Sql = Sql & " case when a.cod_moneda='NS' then (cb1.MTO_CNTFIN+cb1.MTO_CNGFIN) else (cb2.MTO_CNTFIN+cb2.MTO_CNGFIN) end AS RMF,"
    Sql = Sql & " c.prc_tasatir,"
    Sql = Sql & " c.prc_percon ,"
    Sql = Sql & " ls.gls_nomcor || ' ' ||  ls.gls_nomsegcor || ' ' ||  ls.gls_patcor || ' ' ||  ls.gls_matcor as supervisor"
    Sql = Sql & " from pp_tmae_poliza a join pp_tmae_ben b on a.num_poliza=b.num_poliza and a.num_endoso=b.num_endoso"
    Sql = Sql & " join pd_tmae_poliza c on a.num_poliza=c.num_poliza join pt_tmae_cotizacion d on c.num_cot=d.num_cot"
    Sql = Sql & " join pt_tmae_detcotizacion j on d.num_cot= j.num_cot"
    Sql = Sql & " join ma_tpar_tabcod e on a.cod_afp=e.cod_elemento and e.cod_tabla='AF'"
    Sql = Sql & " join ma_tpar_tabcod f on b.cod_par=f.cod_elemento and f.cod_tabla='PA'"
    Sql = Sql & " join ma_tpar_tabcod g on a.cod_tippension=g.cod_elemento and g.cod_tabla='TP'"
    Sql = Sql & " join ma_tpar_tabcod h on a.cod_modalidad=h.cod_elemento and h.cod_tabla='AL'"
    Sql = Sql & " left join (select a.num_poliza, num_idenreceptor, max(fec_pago), cod_par from pp_tmae_liqpagopendef a join pp_tmae_ben b on b.num_idenben=a.num_idenreceptor group by a.num_poliza, num_idenreceptor, cod_par ) i on b.num_idenben=i.num_idenreceptor"
    Sql = Sql & " join pd_tmae_polprirec k on a.num_poliza=k.num_poliza"
    Sql = Sql & " join pt_tmae_corredor l on c.num_idencor=l.num_idencor"
    Sql = Sql & " join pt_tmae_corredor ls on l.num_idenjefe=ls.num_idencor"
    Sql = Sql & " left join pr_tmae_calben1 cb1 on b.num_poliza=cb1.num_poliza and b.num_orden=cb1.num_orden"
    Sql = Sql & " left join pr_tmae_calben2 cb2 on b.num_poliza=cb2.num_poliza and b.num_orden=cb2.num_orden"
    Sql = Sql & " where a.num_endoso=(select max(num_endoso) from pp_tmae_poliza where num_poliza=a.num_poliza) and a.num_poliza >=  " & txtPoliza & " order by 1,2"

    Set vlRegistro = vgConexionBD.Execute(Sql)
   
    Dim registros As Variant
    
    registros = vlRegistro.GetRows()
    
    'MsgBox "Cantidad de registros: " & UBound(registros, 2) + 1
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = UBound(registros, 2) + 1
    
    vlRegistro.MoveFirst
    
    If Not vlRegistro.EOF Then
        While Not vlRegistro.EOF
            xlapp.Range("A" & ix) = IIf(IsNull(vlRegistro(0).Value) = True, "", vlRegistro(0).Value)
            xlapp.Range("B" & ix) = IIf(IsNull(vlRegistro(1).Value) = True, "", vlRegistro(1).Value)
            xlapp.Range("C" & ix) = IIf(IsNull(vlRegistro(2).Value) = True, "", vlRegistro(2).Value)
            xlapp.Range("D" & ix) = IIf(IsNull(vlRegistro(3).Value) = True, "", vlRegistro(3).Value)
            xlapp.Range("E" & ix) = IIf(IsNull(vlRegistro(4).Value) = True, "", vlRegistro(4).Value)
            xlapp.Range("F" & ix) = IIf(IsNull(vlRegistro(5).Value) = True, "", vlRegistro(5).Value)
            xlapp.Range("G" & ix) = IIf(IsNull(vlRegistro(6).Value) = True, "", vlRegistro(6).Value)
            xlapp.Range("H" & ix) = IIf(IsNull(vlRegistro(7).Value) = True, "", vlRegistro(7).Value)
            xlapp.Range("I" & ix) = IIf(IsNull(vlRegistro(8).Value) = True, "", vlRegistro(8).Value)
            xlapp.Range("J" & ix) = IIf(IsNull(vlRegistro(9).Value) = True, "", vlRegistro(9).Value)
            xlapp.Range("K" & ix) = IIf(IsNull(vlRegistro(10).Value) = True, "", vlRegistro(10).Value)
            xlapp.Range("L" & ix) = IIf(IsNull(vlRegistro(11).Value) = True, "", vlRegistro(11).Value)
            xlapp.Range("M" & ix) = IIf(IsNull(vlRegistro(12).Value) = True, "", vlRegistro(12).Value)
            xlapp.Range("N" & ix) = IIf(IsNull(vlRegistro(13).Value) = True, "", vlRegistro(13).Value)
            xlapp.Range("O" & ix) = IIf(IsNull(vlRegistro(14).Value) = True, "", vlRegistro(14).Value)
            xlapp.Range("P" & ix) = IIf(IsNull(vlRegistro(15).Value) = True, "", vlRegistro(15).Value)
            xlapp.Range("Q" & ix) = IIf(IsNull(vlRegistro(16).Value) = True, "", vlRegistro(16).Value)
            xlapp.Range("R" & ix) = IIf(IsNull(vlRegistro(17).Value) = True, "", vlRegistro(17).Value)
            xlapp.Range("S" & ix) = IIf(IsNull(vlRegistro(18).Value) = True, "", vlRegistro(18).Value)
            xlapp.Range("T" & ix) = IIf(IsNull(vlRegistro(19).Value) = True, "", vlRegistro(19).Value)
            xlapp.Range("U" & ix) = IIf(IsNull(vlRegistro(20).Value) = True, "", vlRegistro(20).Value)
            xlapp.Range("V" & ix) = IIf(IsNull(vlRegistro(21).Value) = True, "", vlRegistro(21).Value)
            xlapp.Range("W" & ix) = IIf(IsNull(vlRegistro(22).Value) = True, "", vlRegistro(22).Value)
            xlapp.Range("X" & ix) = IIf(IsNull(vlRegistro(23).Value) = True, "", vlRegistro(23).Value)
            xlapp.Range("Y" & ix) = IIf(IsNull(vlRegistro(24).Value) = True, "", vlRegistro(24).Value)
            xlapp.Range("Z" & ix) = IIf(IsNull(vlRegistro(25).Value) = True, "", vlRegistro(25).Value)
            xlapp.Range("AA" & ix) = IIf(IsNull(vlRegistro(26).Value) = True, "", vlRegistro(26).Value)
            xlapp.Range("AB" & ix) = IIf(IsNull(vlRegistro(27).Value) = True, "", vlRegistro(27).Value)
            xlapp.Range("AC" & ix) = IIf(IsNull(vlRegistro(28).Value) = True, "", vlRegistro(28).Value)
            xlapp.Range("AD" & ix) = IIf(IsNull(vlRegistro(29).Value) = True, "", vlRegistro(29).Value)
            xlapp.Range("AE" & ix) = IIf(IsNull(vlRegistro(30).Value) = True, "", vlRegistro(30).Value)
            xlapp.Range("AF" & ix) = IIf(IsNull(vlRegistro(31).Value) = True, "", vlRegistro(31).Value)
            xlapp.Range("AG" & ix) = IIf(IsNull(vlRegistro(32).Value) = True, "", vlRegistro(32).Value)
            xlapp.Range("AH" & ix) = IIf(IsNull(vlRegistro(33).Value) = True, "", vlRegistro(33).Value)
            xlapp.Range("AI" & ix) = IIf(IsNull(vlRegistro(34).Value) = True, "", vlRegistro(34).Value)
            xlapp.Range("AJ" & ix) = IIf(IsNull(vlRegistro(35).Value) = True, "", vlRegistro(35).Value)
            xlapp.Range("AK" & ix) = IIf(IsNull(vlRegistro(36).Value) = True, "", vlRegistro(36).Value)
            xlapp.Range("AL" & ix) = IIf(IsNull(vlRegistro(37).Value) = True, "", vlRegistro(37).Value)
            xlapp.Range("AM" & ix) = IIf(IsNull(vlRegistro(38).Value) = True, "", vlRegistro(38).Value)
            xlapp.Range("AN" & ix) = IIf(IsNull(vlRegistro(39).Value) = True, "", vlRegistro(39).Value)
            xlapp.Range("AO" & ix) = IIf(IsNull(vlRegistro(40).Value) = True, "", vlRegistro(40).Value)
            xlapp.Range("AP" & ix) = IIf(IsNull(vlRegistro(41).Value) = True, "", vlRegistro(41).Value)
            xlapp.Range("AQ" & ix) = IIf(IsNull(vlRegistro(42).Value) = True, "", vlRegistro(42).Value)
            xlapp.Range("AR" & ix) = IIf(IsNull(vlRegistro(43).Value) = True, "", vlRegistro(43).Value)
            xlapp.Range("AS" & ix) = IIf(IsNull(vlRegistro(44).Value) = True, "", vlRegistro(44).Value)
            xlapp.Range("AT" & ix) = IIf(IsNull(vlRegistro(45).Value) = True, "", vlRegistro(45).Value)
            xlapp.Range("AU" & ix) = IIf(IsNull(vlRegistro(46).Value) = True, "", vlRegistro(46).Value)
            xlapp.Range("AV" & ix) = IIf(IsNull(vlRegistro(47).Value) = True, "", vlRegistro(47).Value)
            xlapp.Range("AW" & ix) = IIf(IsNull(vlRegistro(48).Value) = True, "", vlRegistro(48).Value)
            xlapp.Range("AX" & ix) = IIf(IsNull(vlRegistro(49).Value) = True, "", vlRegistro(49).Value)
            xlapp.Range("AY" & ix) = IIf(IsNull(vlRegistro(50).Value) = True, "", vlRegistro(50).Value)
            xlapp.Range("AZ" & ix) = IIf(IsNull(vlRegistro(51).Value) = True, "", vlRegistro(51).Value)
            xlapp.Range("BA" & ix) = IIf(IsNull(vlRegistro(52).Value) = True, "", vlRegistro(52).Value)
            xlapp.Range("BB" & ix) = IIf(IsNull(vlRegistro(53).Value) = True, "", vlRegistro(53).Value)
            xlapp.Range("BC" & ix) = IIf(IsNull(vlRegistro(54).Value) = True, "", vlRegistro(54).Value)
            xlapp.Range("BC" & ix) = IIf(IsNull(vlRegistro(55).Value) = True, "", vlRegistro(55).Value)
            ix = ix + 1
            
            ProgressBar1.Value = ix - 2
            
            vlRegistro.MoveNext
        Wend
    End If
    Screen.MousePointer = 0
    xlapp.WindowState = xlMaximized
    'xlapp.Workbooks.Close (sArchivo)
    ProgressBar1.Value = 0
    MsgBox ("Datos Exportados!")
End Sub

