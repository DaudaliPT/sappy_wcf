using System;
using System.Collections.Generic;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.ReportSource;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.ServiceModel.Web;

class HelperCrystalReports : IDisposable
{
    public ReportDocument rptDoc;

    public bool OpenReport(string fName, string dbName)
    {

        Logger.LogInvoke("OpenReport", true);
        rptDoc = new CrystalDecisions.CrystalReports.Engine.ReportDocument();

        try
        {
            Logger.Log.Debug("Load file " + fName);
            rptDoc.Load(fName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault);
            Logger.Log.Debug("Loaded file " + fName);
        }
        catch (Exception e)
        {
            throw new Exception("Erro em myReport.Load: " + e.Message);
        }


        Logger.Log.Debug("Start set location...");
        // Check if 64-bit app 
        var driverName = "";
        if (IntPtr.Size == 8) driverName = "B1CRHPROXY"; else driverName = "B1CRHPROXY32";

        string strConnection = "DRIVER={" + driverName + "}";
        strConnection += ";UID=" + SappyWCF_implementation.Properties.Settings.Default.DBUSER;
        strConnection += ";PWD=" + SappyWCF_implementation.Properties.Settings.Default.DBUSERPASS;
        strConnection += ";SERVERNODE=" + SappyWCF_implementation.Properties.Settings.Default.DBSERVER;
        strConnection += ";DATABASE=" + dbName + ";";

        for (int i = 0; i < rptDoc.DataSourceConnections.Count; i++)
        {
            Logger.Log.Debug("Set location " + i + "...");

            NameValuePairs2 logonProps2 = rptDoc.DataSourceConnections[i].LogonProperties;
            logonProps2.Set("Provider", driverName);
            logonProps2.Set("Server Type", driverName);
            logonProps2.Set("Connection String", strConnection);

            rptDoc.DataSourceConnections[i].SetLogonProperties(logonProps2);
            rptDoc.DataSourceConnections[i].SetConnection(
                SappyWCF_implementation.Properties.Settings.Default.DBSERVER,
                dbName,
                SappyWCF_implementation.Properties.Settings.Default.DBUSER,
                SappyWCF_implementation.Properties.Settings.Default.DBUSERPASS);
        }

        rptDoc.Refresh();
        Logger.Log.Debug("End set location...");


        Logger.LogResult("OpenReport", true);
        return true;
    }

    public string GetSAPReportTemplate(string empresa, string docCode)
    {

        Logger.LogInvoke("GetSAPReportTemplate", empresa + "," + docCode);
        // read the report template
        string sql2 = "select \"Template\", coalesce(\"RptHash\",'') as \"RptHash\", \"Category\" from " + empresa + ".rdoc where \"DocCode\" = '" + docCode + "'";

        using (HelperOdbc dataLayer = new HelperOdbc())
        using (System.Data.Odbc.OdbcDataReader dr = dataLayer.ExecuteReader(sql2))
        {
            if (!dr.Read())
            {
                throw new Exception("Relatório " + docCode + " não existe em " + empresa);
            }

            if (dr.GetString(dr.GetOrdinal("Category")) != "C")
            {
                throw new Exception("O layout " + docCode + " não usa Crystal Reports na " + empresa);
            }

            var rptHash = dr.GetString(dr.GetOrdinal("RptHash"));

            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ReportCache");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            string fName = Path.Combine(path, docCode + "_" + rptHash + ".rpt");
            if (!File.Exists(fName))
            {
                using (FileStream stream = new FileStream(fName, FileMode.OpenOrCreate, FileAccess.Write))
                {
                    var writer = new BinaryWriter(stream);
                    int bufferSize = 100;                   // Size of the BLOB buffer.  
                    byte[] outByte = new byte[bufferSize];  // The BLOB byte[] buffer to be filled by GetBytes.  
                    long retval;                            // The bytes returned from GetBytes.  
                    long startIndex = 0;                    // The starting position in the BLOB output.  

                    int template = dr.GetOrdinal("Template");
                    if (!dr.IsDBNull(template))
                    {
                        // Read bytes into outByte[] and retain the number of bytes returned.  
                        retval = dr.GetBytes(template, startIndex, outByte, 0, bufferSize);

                        // Continue while there are bytes beyond the size of the buffer.  
                        while (retval == bufferSize)
                        {
                            writer.Write(outByte);
                            writer.Flush();

                            // Reposition start index to end of last buffer and fill buffer.  
                            startIndex += bufferSize;
                            retval = dr.GetBytes(template, startIndex, outByte, 0, bufferSize);
                        }

                        // Write the remaining buffer.  
                        writer.Write(outByte, 0, (int)retval - 1);
                        writer.Flush();

                        // Close the output file.  
                        writer.Close();
                    }
                }
            }

            Logger.LogResult("GetSAPReportTemplate", fName);
            return fName;
        }
    }


    public void setParametersDynamically(string empresa, string docCode)
    {

        string parValuesJson = WebOperationContext.Current.IncomingRequest.UriTemplateMatch.QueryParameters["parValues"];

        //Car car = new Car() { Id = 1, Name = "Polo", Company = "VW" };
        //string serialized = JsonConvert.SerializeObject(car);
        //dynamic deserialized = JsonConvert.DeserializeObject(serialized);


        dynamic parValues = Newtonsoft.Json.JsonConvert.DeserializeObject(parValuesJson);


        foreach (var par in parValues)
        {
            var thisVal = par.Value;
            ParameterField rptPar = this.rptDoc.ParameterFields.Find(par.Name, "");
            if (rptPar == null) continue;

            if (rptPar.PromptingType == DiscreteOrRangeKind.DiscreteValue)
            {
                if (rptPar.EnableAllowMultipleValue == false)
                {
                    var discrete = new ParameterDiscreteValue();

                    if (rptPar.ParameterValueType == ParameterValueKind.DateParameter) discrete.Value = (DateTime)thisVal.Value;
                    else if (rptPar.ParameterValueType == ParameterValueKind.DateTimeParameter) discrete.Value = (DateTime)thisVal.Value;
                    else if (rptPar.ParameterValueType == ParameterValueKind.StringParameter) discrete.Value = (string)thisVal.Value;
                    else discrete.Value = thisVal.Value;

                    rptPar.CurrentValues.Add(discrete);//Add the value
                }
                else
                {
                    foreach (var item in thisVal.Value)
                    {
                        var discrete = new ParameterDiscreteValue();

                        if (rptPar.ParameterValueType == ParameterValueKind.DateParameter) discrete.Value = (DateTime)item;
                        else if (rptPar.ParameterValueType == ParameterValueKind.DateTimeParameter) discrete.Value = (DateTime)item;
                        else if (rptPar.ParameterValueType == ParameterValueKind.StringParameter) discrete.Value = (string)item;
                        else discrete.Value = item;

                        rptPar.CurrentValues.Add(discrete);//Add the value
                    }
                }

            }
            else if (rptPar.PromptingType == DiscreteOrRangeKind.RangeValue)
            {
                var range = new ParameterRangeValue();
                range.StartValue = thisVal.Value.ToString();
                range.EndValue = thisVal.EndValue.ToString();
                rptPar.CurrentValues.Add(range);//Add the value 
            }
            else
            {
                Debug.Print("");
            }
        }
    }

    public void Dispose()
    {
        if (rptDoc != null)
        {
            rptDoc.Close();
            rptDoc.Dispose();
        }
    }
}


