VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_plantilla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Producción"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   1725
   ClientWidth     =   15105
   Icon            =   "frm_plantilla.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   15105
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer crw_object 
      Height          =   9045
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   16830
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frm_plantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objApp As New CRAXDRT.Application
Private objReport As CRAXDRT.Report
Private objSubReport As CRAXDRT.Report
Private objTotReport As CRAXDRT.Report

Private Sub Form_Load()

  'Call p_centerForm(mdi_principal, Me)
  Set objReport = objApp.OpenReport(strRpt & "" & sName_Reporte)
  Select Case sName_Reporte
    Case "PD_Rpt_SBSTasaCot.rpt"
      'objReport.ParameterFields.GetItemByName("pm_sTitulo").AddCurrentValue (svrpt_sTitulo)
      'objReport.ParameterFields.GetItemByName("pm_sPeriodo").AddCurrentValue (svrpt_sSubTitulo)
      objReport.Database.SetDataSource objRsRpt
      crw_object.ReportSource = objReport
    Case "PD_Rpt_LiquidacionRV.rpt"
      objReport.ParameterFields.GetItemByName("NombreCompania").AddCurrentValue (vgNombreCompania)
      objReport.ParameterFields.GetItemByName("rutCliente").AddCurrentValue (vgNumIdenCliente)
      objReport.Database.SetDataSource objRsRpt
      crw_object.ReportSource = objReport
  End Select
  crw_object.ViewReport
  Screen.MousePointer = 0
  Set objReport = Nothing
  DoEvents

End Sub

