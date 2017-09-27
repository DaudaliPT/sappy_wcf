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
    [WebInvoke(Method = "GET",
        ResponseFormat = WebMessageFormat.Json,
        UriTemplate = "{empresa}/GetPdf({docCode})")]
    Stream GetPdf(string empresa, string docCode);

    [OperationContract]
    [WebInvoke(Method = "POST",
        ResponseFormat = WebMessageFormat.Json,
        UriTemplate = "{empresa}/print({docCode})")]
    bool Print(string empresa, string docCode);

    [OperationContract]
    [WebInvoke(Method = "GET",
        ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.Bare,
        UriTemplate = "{empresa}/GetPdfParameters({docCode})")]
    string GetPdfParameters(string empresa, string docCode);

    [OperationContract]
    [WebInvoke(Method = "POST",
        RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.WrappedRequest,
        UriTemplate = "{empresa}/AddDoc/{objCode}/{draftId}?expectedTotal={expectedTotal}")]
    string AddDoc(string empresa, string objCode, string draftId, string expectedTotal);
    
    [OperationContract]
    [WebInvoke(Method = "POST",
        RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.WrappedRequest,
        UriTemplate = "{empresa}/SimulateDoc/{objCode}/{draftId}")]
    string SimulateDoc(string empresa, string objCode, string draftId);

    [OperationContract]
    [WebInvoke(Method = "PATCH",
        RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.WrappedRequest,
        UriTemplate = "{empresa}/PatchDoc/{objCode}/{docEntry}")]
    string PatchDoc(string empresa, string objCode, string docEntry);

    [OperationContract]
    [WebInvoke(Method = "GET",
        ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.Bare,
        UriTemplate = "printers")]
    string GetPrinters();

}
