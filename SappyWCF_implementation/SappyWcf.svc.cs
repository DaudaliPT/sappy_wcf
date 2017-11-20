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
using System.Text;



public class SappyWcf : I_SappyWcf
{
    public Stream GetPdf(string empresa, string docCode)
    {
        string sInfo = "GetPdf";
        Logger.LogInvoke(sInfo, empresa, docCode);

        try
        {
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

    public Stream ReportExport(string empresa, string docCode, string format)
    {
        string sInfo = "ReportExport";
        Logger.LogInvoke(sInfo, empresa, docCode);

        try
        {
            using (HelperCrystalReports crw = new HelperCrystalReports())
            {
                var rptFile = crw.GetSAPReportTemplate(empresa, docCode);
                crw.OpenReport(rptFile, empresa);
                crw.setParametersDynamically(empresa, docCode);

                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Exported");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);


                ExportFormatType formatType;
                string outFile;
                if (format == "xls")
                {
                    WebOperationContext.Current.OutgoingResponse.ContentType = "application/vnd.ms-excel";
                    outFile = Path.Combine(path, Guid.NewGuid().ToString() + ".xls");
                    formatType = ExportFormatType.Excel;
                }
                else if (format == "doc")
                {
                    WebOperationContext.Current.OutgoingResponse.ContentType = "application/msword";
                    outFile = Path.Combine(path, Guid.NewGuid().ToString() + ".doc");
                    formatType = ExportFormatType.WordForWindows;
                }
                else if (format == "csv")
                {
                    WebOperationContext.Current.OutgoingResponse.ContentType = "text/csv";
                    outFile = Path.Combine(path, Guid.NewGuid().ToString() + ".csv");
                    formatType = ExportFormatType.CharacterSeparatedValues;
                } 
                else if (format == "rpt")
                {
                    WebOperationContext.Current.OutgoingResponse.ContentType = "application/octet-stream";
                    outFile = Path.Combine(path, Guid.NewGuid().ToString() + ".rpt");
                    formatType = ExportFormatType.CrystalReport;
                }
                else if (format == "rtf")
                {
                    WebOperationContext.Current.OutgoingResponse.ContentType = "application/rtf";
                    outFile = Path.Combine(path, Guid.NewGuid().ToString() + ".rtf");
                    formatType = ExportFormatType.EditableRTF;
                }
                else  
                {
                    WebOperationContext.Current.OutgoingResponse.ContentType = "application/pdf";
                    outFile = Path.Combine(path, Guid.NewGuid().ToString() + ".pdf");
                    formatType = ExportFormatType.PortableDocFormat;
                }

                Logger.Log.Debug("Create " + outFile);
                
                crw.rptDoc.ExportToDisk(formatType, outFile);

                // Return the doc
                Logger.LogResult(sInfo, new Object());

                return new FileStream_ThatDeletesFileAfterReading(outFile, FileMode.Open, FileAccess.Read);
            }
        }
        catch (System.Exception ex)
        {
            Logger.Log.Error(ex.Message, ex);
            throw new WebFaultException<string>(ex.Message, HttpStatusCode.InternalServerError);
        }
    }


    public bool ReportPrint(string empresa, string docCode)
    {
        string sInfo = "Print";
        try
        {
            Logger.LogInvoke(sInfo, empresa, docCode);

            using (HelperCrystalReports crw = new HelperCrystalReports())
            {
                string toPrinter = WebOperationContext.Current.IncomingRequest.UriTemplateMatch.QueryParameters["toPrinter"];
                string toPrinterTaloes = WebOperationContext.Current.IncomingRequest.UriTemplateMatch.QueryParameters["toPrinterTaloes"];

                var fname = crw.GetSAPReportTemplate(empresa, docCode);
                crw.OpenReport(fname, empresa);
                crw.setParametersDynamically(empresa, docCode);

                CrystalDecisions.ReportAppServer.Controllers.PrintReportOptions popt = new CrystalDecisions.ReportAppServer.Controllers.PrintReportOptions();

                if (crw.rptDoc.SummaryInfo.KeywordsInReport != null &&
                    crw.rptDoc.SummaryInfo.KeywordsInReport.Contains("USE_POS_PRINTER") &&
                    toPrinterTaloes != "")
                    popt.PrinterName = toPrinterTaloes;
                else if (toPrinter != "")
                    popt.PrinterName = toPrinter;

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

    public bool Print(string empresa, string docCode)
    {
        string sInfo = "Print"; 
        try
        {
            Logger.LogInvoke(sInfo, empresa, docCode);

            using (HelperCrystalReports crw = new HelperCrystalReports())
            {
                string toPrinter = WebOperationContext.Current.IncomingRequest.UriTemplateMatch.QueryParameters["toPrinter"];
                string toPrinterTaloes = WebOperationContext.Current.IncomingRequest.UriTemplateMatch.QueryParameters["toPrinterTaloes"];

                var fname = crw.GetSAPReportTemplate(empresa, docCode);
                crw.OpenReport(fname, empresa);
                crw.setParametersDynamically(empresa, docCode);

                CrystalDecisions.ReportAppServer.Controllers.PrintReportOptions popt = new CrystalDecisions.ReportAppServer.Controllers.PrintReportOptions();

                if (crw.rptDoc.SummaryInfo.KeywordsInReport!=null && 
                    crw.rptDoc.SummaryInfo.KeywordsInReport.Contains("USE_POS_PRINTER") && 
                    toPrinterTaloes != "") 
                    popt.PrinterName = toPrinterTaloes;
                else if (toPrinter!="") 
                    popt.PrinterName = toPrinter;

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

    public string GetReportParameters(string empresa, string docCode)
    {
        string sInfo = "GetReportParameters";
        Logger.LogInvoke(sInfo, empresa, docCode);
        try
        {
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

    public string AddDoc(string empresa, string objCode, string draftId, string expectedTotal)
    {
        Result result = new Result();
        string sInfo = "AddDoc";
        Logger.LogInvoke(sInfo, empresa, objCode, draftId, expectedTotal);

        if (!SBOHandler.DIAPIConnections.ContainsKey(empresa))
            result.error = "Empresa não configurada ou inexistente.";
        else
        {
            var sboCon = SBOHandler.DIAPIConnections[empresa];

            if (Monitor.TryEnter(sboCon, new TimeSpan(0, 0, 10)))
            {
                try
                {
                    int Id = Convert.ToInt32(draftId);
                    double ExpectedTotal = Convert.ToDouble(expectedTotal, CultureInfo.InvariantCulture);

                    result.result = sboCon.SAPDOC_FROM_SAPPY_DRAFT(DocActions.ADD, objCode, Id, ExpectedTotal);
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
            else
            {
                result.error = "Busy, please try again later...";
            }
        }
        Logger.LogResult(sInfo, result);
        return Logger.FormatToJson(result);

    }

    public string AddDocPOS(string empresa, string objCode, string draftId, string expectedTotal)
    {
        Result result = new Result();
        string sInfo = "AddDocPOS";
        Logger.LogInvoke(sInfo, empresa, objCode, draftId, expectedTotal);

        if (!SBOHandler.DIAPIConnections.ContainsKey(empresa))
            result.error = "Empresa não configurada ou inexistente.";
        else
        {
            var sboCon = SBOHandler.DIAPIConnections[empresa];

            if (Monitor.TryEnter(sboCon, new TimeSpan(0, 0, 10)))
            {
                try
                {
                    int Id = Convert.ToInt32(draftId);
                    double ExpectedTotal = Convert.ToDouble(expectedTotal, CultureInfo.InvariantCulture);

                    result.result = sboCon.SAPDOC_FROM_SAPPY_DRAFT_POS(DocActions.ADD, objCode, Id, ExpectedTotal);
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
            else
            {
                result.error = "Busy, please try again later...";
            }
        }
        Logger.LogResult(sInfo, result);
        return Logger.FormatToJson(result);

    }

    public string SimulateDoc(string empresa, string objCode, string draftId)
    {
        Result result = new Result();
        string sInfo = "SimulateDoc";
        Logger.LogInvoke(sInfo, empresa, objCode, draftId);

        if (!SBOHandler.DIAPIConnections.ContainsKey(empresa))
            result.error = "Empresa não configurada ou inexistente.";
        else 
        {
            var sboCon = SBOHandler.DIAPIConnections[empresa];
            if (Monitor.TryEnter(sboCon, new TimeSpan(0, 0, 10)))
            {
                try
                {
                    int Id = Convert.ToInt32(draftId);
                    result.result = sboCon.SAPDOC_FROM_SAPPY_DRAFT(DocActions.SIMULATE, objCode, Id, 0);
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
            else
            {
                result.error = "Busy, please try again later...";
            }
        }
        Logger.LogResult(sInfo, result);
        return Logger.FormatToJson(result);

    }



    public string PatchDoc(string empresa, string objCode, string docEntry)
    {
        Result result = new Result();
        string sInfo = "PatchDoc";
        Logger.LogInvoke(sInfo, empresa, objCode, docEntry);

        if (!SBOHandler.DIAPIConnections.ContainsKey(empresa))
            result.error = "Empresa não configurada ou inexistente.";
        else
        {

            var sboCon = SBOHandler.DIAPIConnections[empresa];

            if (Monitor.TryEnter(sboCon, new TimeSpan(0, 0, 10)))
            {
                try
                {
                    int DocEntry = Convert.ToInt32(docEntry);

                    result.result = sboCon.SAPDOC_PATCH_WITH_SAPPY_CHANGES(objCode, DocEntry);
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
            else
            {
                result.error = "Busy, please try again later...";
            }
        }

        Logger.LogResult(sInfo, result);
        return Logger.FormatToJson(result);

    }

    public string CancelDoc(string empresa, string objCode, string docEntry)
    {
        Result result = new Result();
        string sInfo = "CancelDoc";
        Logger.LogInvoke(sInfo, empresa, objCode, docEntry);

        if (!SBOHandler.DIAPIConnections.ContainsKey(empresa))
            result.error = "Empresa não configurada ou inexistente.";
        else
        {

            var sboCon = SBOHandler.DIAPIConnections[empresa];

            if (Monitor.TryEnter(sboCon, new TimeSpan(0, 0, 10)))
            {
                try
                {
                    int DocEntry = Convert.ToInt32(docEntry);

                    result.result = sboCon.SAPDOC_CANCELDOC(objCode, DocEntry);
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
            else
            {
                result.error = "Busy, please try again later...";
            }
        }

        Logger.LogResult(sInfo, result);
        return Logger.FormatToJson(result);

    }

    public string CloseDoc(string empresa, string objCode, string docEntry)
    {
        Result result = new Result();
        string sInfo = "CloseDoc";
        Logger.LogInvoke(sInfo, empresa, objCode, docEntry);

        if (!SBOHandler.DIAPIConnections.ContainsKey(empresa))
            result.error = "Empresa não configurada ou inexistente.";
        else
        {

            var sboCon = SBOHandler.DIAPIConnections[empresa];

            if (Monitor.TryEnter(sboCon, new TimeSpan(0, 0, 10)))
            {
                try
                {
                    int DocEntry = Convert.ToInt32(docEntry);

                    result.result = sboCon.SAPDOC_CLOSEDOC(objCode, DocEntry);
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
            else
            {
                result.error = "Busy, please try again later...";
            }
        }

        Logger.LogResult(sInfo, result);
        return Logger.FormatToJson(result);

    }

    public string GetPrinters()
    {
        string sInfo = "GetPrinters";
        try
        {
            Logger.LogInvoke(sInfo, "");

            WebOperationContext.Current.OutgoingResponse.Format = WebMessageFormat.Json;
            string jsonString = new JavaScriptSerializer().Serialize(System.Drawing.Printing.PrinterSettings.InstalledPrinters);
            return jsonString;
        }
        catch (System.Exception ex)
        {
            Logger.Log.Error(ex.Message, ex);

            throw new WebFaultException<string>(ex.ToString(), HttpStatusCode.NotFound);
        }
    }

    public string PostAdiantamento(PostAdiantamentoInput body, string empresa)
    {
        Result result = new Result();
        string sInfo = "PostAdiantamento";
        Logger.LogInvoke(sInfo, empresa, body);

        if (!SBOHandler.DIAPIConnections.ContainsKey(empresa))
            result.error = "Empresa não configurada ou inexistente.";
        else
        {
            var sboCon = SBOHandler.DIAPIConnections[empresa];

            if (Monitor.TryEnter(sboCon, new TimeSpan(0, 0, 10)))
            {
                try
                {
                    result.result = sboCon.ADD_ADIANTAMENTO_PARA_DESPESAS(body);
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
            else
            {
                result.error = "Busy, please try again later...";
            }
        }
        Logger.LogResult(sInfo, result);
        return Logger.FormatToJson(result);
    }

    public string PostFecharAdiantamento(PostFecharAdiantamentoInput body, string empresa)
    {
        Result result = new Result();
        string sInfo = "PostFecharAdiantamento";
        Logger.LogInvoke(sInfo, empresa);

        if (!SBOHandler.DIAPIConnections.ContainsKey(empresa))
            result.error = "Empresa não configurada ou inexistente.";

        else
        {
            var sboCon = SBOHandler.DIAPIConnections[empresa];

            if (Monitor.TryEnter(sboCon, new TimeSpan(0, 0, 10)))
            {
                try
                {
                    result.result = sboCon.FECHAR_ADIANTAMENTO_PARA_DESPESAS(body);
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
            else
            {
                result.error = "Busy, please try again later...";
            }
        }

        Logger.LogResult(sInfo, result);
        return Logger.FormatToJson(result);
    }


    public string PostDespesa(PostDespesaInput body, string empresa)
    {
        Result result = new Result();
        string sInfo = "PostDespesa";

        if (!SBOHandler.DIAPIConnections.ContainsKey(empresa))
            result.error = "Empresa não configurada ou inexistente.";
        else
        {
            var sboCon = SBOHandler.DIAPIConnections[empresa];

            if (Monitor.TryEnter(sboCon, new TimeSpan(0, 0, 10)))
            {
                try
                {
                    Logger.LogInvoke(sInfo, "");

                    result.result = sboCon.ADD_DESPESA(body);
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
            else
            {
                result.error = "Busy, please try again later...";
            }
        }
        Logger.LogResult(sInfo, result);
        return Logger.FormatToJson(result);
    }
}