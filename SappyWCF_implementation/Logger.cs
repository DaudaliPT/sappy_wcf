using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using log4net.Config;
using System.Web.Script.Serialization;


public static class Logger
{
    private class InitLogger
    {
        public InitLogger()
        {
            XmlConfigurator.Configure();
        } 
    }

    private static InitLogger initLogger = new InitLogger();
    /// <summary>
    /// The log4net instance.
    /// </summary>        
    public static log4net.ILog Log = log4net.LogManager.GetLogger("LOGFILE");// MethodBase.GetCurrentMethod().DeclaringType);

    internal static void LogInvoke(string procName, string request1, object request)
    {
        Log.Info("Invoked " + procName + " on DB server: " + SappyWCF_implementation.Properties.Settings.Default.DBSERVER);
        Log.Debug("Invoked " + procName + "((request: " + request1 + ", " + FormatToJson(request) + ") on DB server: " + SappyWCF_implementation.Properties.Settings.Default.DBSERVER);
    }

    public static void LogInvoke(string procName, object request)
    {
        Log.Info("Invoked " + procName + " on Db Server: " + SappyWCF_implementation.Properties.Settings.Default.DBSERVER);
        Log.Debug("Invoked " + procName + "(request: " + FormatToJson(request) + ") on DB server: " + SappyWCF_implementation.Properties.Settings.Default.DBSERVER);
    }

    public static void LogResult(string procName, object result)
    {
        Log.Info("Completed invocation of " + procName);
        Log.Debug("ResultOf " + procName + ": " + FormatToJson(result));
    }

    /// <summary>
    /// Adds indentation and line breaks to output of JavaScriptSerializer
    /// </summary>
    internal static string FormatToJson(object toConvert)
    {
        string jsonString = new JavaScriptSerializer().Serialize(toConvert);
        var stringBuilder = new StringBuilder();

        bool escaping = false;
        bool inQuotes = false;
        int indentation = 0;

        foreach (char character in jsonString)
        {
            if (escaping)
            {
                escaping = false;
                stringBuilder.Append(character);
            }
            else
            {
                if (character == '\\')
                {
                    escaping = true;
                    stringBuilder.Append(character);
                }
                else if (character == '\"')
                {
                    inQuotes = !inQuotes;
                    stringBuilder.Append(character);
                }
                else if (!inQuotes)
                {
                    if (character == ',')
                    {
                        stringBuilder.Append(character);
                        stringBuilder.Append("\r\n");
                        stringBuilder.Append('\t', indentation);
                    }
                    else if (character == '[' || character == '{')
                    {
                        stringBuilder.Append(character);
                        stringBuilder.Append("\r\n");
                        stringBuilder.Append('\t', ++indentation);
                    }
                    else if (character == ']' || character == '}')
                    {
                        stringBuilder.Append("\r\n");
                        stringBuilder.Append('\t', --indentation);
                        stringBuilder.Append(character);
                    }
                    else if (character == ':')
                    {
                        stringBuilder.Append(character);
                        stringBuilder.Append('\t');
                    }
                    else
                    {
                        stringBuilder.Append(character);
                    }
                }
                else
                {
                    stringBuilder.Append(character);
                }
            }
        }

        return stringBuilder.ToString();
    }
}
