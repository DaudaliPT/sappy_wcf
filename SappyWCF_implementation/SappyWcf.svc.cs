using System;
using System.Collections.Generic;
using System.Data;
using System.ServiceModel;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using System.ServiceModel.Web;
using System.Net;
using System.Web.Script.Serialization;
using CrystalDecisions.Shared;
using System.Drawing.Printing;
using WcfSappy.STAThread;
using System.Threading;



public class SappyWcf : I_SappyWcf
{
    public Stream GetPdf(string empresa, string docCode)
    {
        string sInfo = "GetPdf";
        try
        {
            Logger.LogInvoke(sInfo, "");

            using (HelperCrystalReports crw = new HelperCrystalReports())
            {
                var fname = crw.GetSAPReportTemplate(empresa, docCode);
                crw.OpenReport(fname, empresa);
                crw.setParametersDynamically(empresa, docCode);

                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GeneratedPdf");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                string fnamePdf = Path.Combine(path, Guid.NewGuid().ToString() + ".pdf");
                string fnameRpt = Path.Combine(path, Guid.NewGuid().ToString() + ".rpt");

                Logger.Log.Debug("Create rpt " + fnameRpt);

                crw.rptDoc.ExportToDisk(ExportFormatType.CrystalReport, fnameRpt);

                Logger.Log.Debug("Create pdf " + fnamePdf);
                crw.rptDoc.ExportToDisk(ExportFormatType.PortableDocFormat, fnamePdf);

                // Return the pdf
                WebOperationContext.Current.OutgoingResponse.ContentType = "application/pdf";
                Logger.LogResult(sInfo, new Object());

                return new FileStream_ThatDeletesFileAfterReading(fnamePdf, FileMode.Open, FileAccess.Read);
            }
        }
        catch (System.Exception ex)
        {
            Logger.Log.Error(ex.Message, ex);
            throw new WebFaultException<string>(ex.Message, HttpStatusCode.InternalServerError);
        }
    }


    public bool Print(string empresa, string docCode)
    {
        string sInfo = "Print";
        try
        {
            Logger.LogInvoke(sInfo, "");

            using (HelperCrystalReports crw = new HelperCrystalReports())
            {

                var fname = crw.GetSAPReportTemplate(empresa, docCode);
                crw.OpenReport(fname, empresa);
                crw.setParametersDynamically(empresa, docCode);

                CrystalDecisions.ReportAppServer.Controllers.PrintReportOptions popt = new CrystalDecisions.ReportAppServer.Controllers.PrintReportOptions();
                popt.PrinterName = SappyWCF_implementation.Properties.Settings.Default.SAPPY001_PrinterName;
                crw.rptDoc.ReportClientDocument.PrintOutputController.PrintReport(popt);


                return true;
            }
        }
        catch (System.Exception ex)
        {
            Logger.Log.Error(ex.Message, ex);

            throw new WebFaultException<string>(ex.Message, HttpStatusCode.InternalServerError);
        }
    }

    public string GetPdfParameters(string empresa, string docCode)
    {
        string sInfo = "GetPdfParameters";
        try
        {
            Logger.LogInvoke(sInfo, "");
            using (HelperCrystalReports crw = new HelperCrystalReports())
            {
                var fname = crw.GetSAPReportTemplate(empresa, docCode);
                crw.OpenReport(fname, empresa);


                WebOperationContext.Current.OutgoingResponse.Format = WebMessageFormat.Json;
                string jsonString = new JavaScriptSerializer().Serialize(crw.rptDoc.ParameterFields);
                return jsonString;
            }
        }
        catch (System.Exception ex)
        {
            Logger.Log.Error(ex.Message, ex);

            throw new WebFaultException<string>(ex.ToString(), HttpStatusCode.NotFound);
        }
    }


    // [STAOperationBehavior]
    public string AddDoc(string empresa, string objCode, string draftId)
    {
        Result result = new Result();
        string sInfo = "AddDoc";

        if (!SBOHandler.DIAPIConnections.ContainsKey(empresa))
            result.error = "Empresa não configurada ou inexistente.";


        var sboCon = SBOHandler.DIAPIConnections[empresa];
         
        if (Monitor.TryEnter(sboCon, new TimeSpan(0, 0, 10)))
        {
            try
            {
                Logger.LogInvoke(sInfo, "");
                int Id = Convert.ToInt32(draftId);

                result.result = sboCon.Confirmar_SAPPY_DOC(objCode, Id);
            }
            catch (System.Exception ex)
            {
                Logger.Log.Error(ex.Message, ex);
                result.error = ex.Message;
            }
            finally
            {
                Monitor.Exit(sboCon);
            }
        }
        else {         
            result.error =  "Busy, please try again later...";
        }

        Logger.LogResult(sInfo, result);
        return Logger.FormatToJson( result);

    }
}