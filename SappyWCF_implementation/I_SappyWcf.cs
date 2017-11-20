using System;
using System.Collections.Generic;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.IO;
using System.Runtime.Serialization;



[DataContract]
public class Result
{
    public Object result = null;
    public Object error = null;
}


[ServiceContract]
public interface I_SappyWcf
{
    [OperationContract]
    [WebInvoke(Method = "GET", UriTemplate = "{empresa}/GetPdf({docCode})", ResponseFormat = WebMessageFormat.Json)]
    Stream GetPdf(string empresa, string docCode);

    [OperationContract]
    [WebInvoke(Method = "GET", UriTemplate = "{empresa}/ReportExport/{docCode}/{format}", ResponseFormat = WebMessageFormat.Json)]
    Stream ReportExport(string empresa, string docCode, string format);

    [OperationContract]
    [WebInvoke(Method = "POST", UriTemplate = "{empresa}/ReportPrint/{docCode}", ResponseFormat = WebMessageFormat.Json)]
    bool ReportPrint(string empresa, string docCode);

    [OperationContract]
    [WebInvoke(Method = "POST", UriTemplate = "{empresa}/print({docCode})", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
    bool Print(string empresa, string docCode);

    [OperationContract]
    [WebInvoke(Method = "GET", UriTemplate = "{empresa}/GetReportParameters/{docCode}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
    string GetReportParameters(string empresa, string docCode);

    [OperationContract]
    [WebInvoke(Method = "POST", UriTemplate = "{empresa}/AddDoc/{objCode}/{draftId}?expectedTotal={expectedTotal}", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.WrappedRequest)]
    string AddDoc(string empresa, string objCode, string draftId, string expectedTotal);

    [OperationContract]
    [WebInvoke(Method = "POST", UriTemplate = "{empresa}/AddDocPOS/{objCode}/{draftId}?expectedTotal={expectedTotal}", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.WrappedRequest)]
    string AddDocPOS(string empresa, string objCode, string draftId, string expectedTotal);

    [OperationContract]
    [WebInvoke(Method = "POST", UriTemplate = "{empresa}/SimulateDoc/{objCode}/{draftId}", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.WrappedRequest)]
    string SimulateDoc(string empresa, string objCode, string draftId);

    [OperationContract]
    [WebInvoke(Method = "PATCH", UriTemplate = "{empresa}/PatchDoc/{objCode}/{docEntry}", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.WrappedRequest)]
    string PatchDoc(string empresa, string objCode, string docEntry);

    [OperationContract]
    [WebInvoke(Method = "POST", UriTemplate = "{empresa}/CancelDoc/{objCode}/{docEntry}", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.WrappedRequest)]
    string CancelDoc(string empresa, string objCode, string docEntry);

    [OperationContract]
    [WebInvoke(Method = "POST", UriTemplate = "{empresa}/CloseDoc/{objCode}/{docEntry}", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.WrappedRequest)]
    string CloseDoc(string empresa, string objCode, string docEntry);

    [OperationContract]
    [WebInvoke(Method = "GET", UriTemplate = "printers", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
    string GetPrinters();

    [OperationContract]
    [WebInvoke(Method = "POST", UriTemplate = "{empresa}/adiantamento", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
    string PostAdiantamento(PostAdiantamentoInput body, string empresa);

    [OperationContract]
    [WebInvoke(Method = "POST", UriTemplate = "{empresa}/fecharadiantamento", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
    string PostFecharAdiantamento(PostFecharAdiantamentoInput body, string empresa);

    [OperationContract]
    [WebInvoke(Method = "POST", UriTemplate = "{empresa}/despesa", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
    string PostDespesa(PostDespesaInput body, string empresa);
}
