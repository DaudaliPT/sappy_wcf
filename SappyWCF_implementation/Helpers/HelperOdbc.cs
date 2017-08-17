using System;
using System.Text;
using System.Data.Odbc;
using System.Data;

class HelperOdbc : IDisposable
{
    public void Dispose()
    {
        if (db != null) db.Dispose();
    }

    internal OdbcConnection db { get; set; }

    public HelperOdbc()
    {
        try
        {
            OdbcConnectionStringBuilder cnnBuilder = new OdbcConnectionStringBuilder();

            // Check if 64-bit app 
            if (IntPtr.Size == 8) cnnBuilder.Driver = "{HDBODBC}"; else cnnBuilder.Driver = "{HDBODBC32}";
            cnnBuilder.Add("SERVERNODE", SappyWCF_implementation.Properties.Settings.Default.DBSERVER);
            cnnBuilder.Add("SERVERDB", "NDB");
            cnnBuilder.Add("UID", SappyWCF_implementation.Properties.Settings.Default.DBUSER);
            cnnBuilder.Add("PWD", SappyWCF_implementation.Properties.Settings.Default.DBUSERPASS);

            var ConnectionString = cnnBuilder.ToString();
            db = new OdbcConnection(ConnectionString);
            db.Open();
        }
        catch (Exception)
        {
            Logger.Log.Error("HelperOdbc(" + SappyWCF_implementation.Properties.Settings.Default.DBSERVER + "): Não foi possivel estabelecer conexão");
            throw;
        }
    }

    public HelperOdbc(string server, string user, string password)
    {
        try
        { 
            OdbcConnectionStringBuilder cnnBuilder = new OdbcConnectionStringBuilder();

            // Check if 64-bit app 
            if (IntPtr.Size == 8) cnnBuilder.Driver = "{HDBODBC}"; else cnnBuilder.Driver = "{HDBODBC32}";
            cnnBuilder.Add("SERVERNODE", server);
            cnnBuilder.Add("SERVERDB", "NDB");
            cnnBuilder.Add("UID", user);
            cnnBuilder.Add("PWD", password);

            var ConnectionString = cnnBuilder.ToString();
            db = new OdbcConnection(ConnectionString);
            db.Open();
        }
        catch (Exception)
        {
            Logger.Log.Error("HelperOdbc(" + server + "): Não foi possivel estabelecer conexão");
            throw;
        }
    }

    internal DataTable Execute(string sqlQuery, bool CleanNulls = true)
    {
        DataTable dt = new DataTable(); 
        using (OdbcCommand cmd = db.CreateCommand())
        {
            try
            {
                cmd.CommandText = sqlQuery;

                OdbcDataAdapter adpt = new OdbcDataAdapter(cmd);
                adpt.Fill(dt);
                adpt.Dispose();
            }
            catch (Exception)
            {
                Logger.Log.Debug(sqlQuery);
                throw;
            }
        }

        if (CleanNulls) ConvertNullsToDefaults(ref dt);

        return dt;
    }

    internal int ExecuteNonQuery(string sqlQuery)
    {
        DataTable dt = new DataTable();
        using (OdbcCommand cmd = db.CreateCommand())
        {
            try
            {
                cmd.CommandText = sqlQuery;
                return cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {
                Logger.Log.Debug(sqlQuery);
                throw;
            }
        }
    }

    internal OdbcDataReader ExecuteReader(string sqlQuery)
    {
        DataTable dt = new DataTable();
        using (OdbcCommand cmd = db.CreateCommand())
        {
            try
            {
                cmd.CommandText = sqlQuery;
                return cmd.ExecuteReader();
            }
            catch (Exception)
            {
                Logger.Log.Debug(sqlQuery);
                throw;
            } 
        }
    }
     
    internal object ExecuteScalar(string sqlQuery)
    {
        DataTable dt = new DataTable();
        using (OdbcCommand cmd = db.CreateCommand())
        {
            try
            {
                cmd.CommandText = sqlQuery;
                return cmd.ExecuteScalar();
            }
            catch (Exception)
            {
                Logger.Log.Debug(sqlQuery);
                throw;
            }
        }
    }

    internal static void ConvertNullsToDefaults(ref DataTable dt)
    {
        for (int i = 0; i <= dt.Rows.Count - 1; i++)
        {
            DataRow row = dt.Rows[i];

            for (int c = 0; c <= dt.Columns.Count - 1; c++)
            {
                if (row.IsNull(c))
                {
                    Type mt = dt.Columns[c].DataType;
                    if (object.ReferenceEquals(mt, typeof(long))) { row[c] = 0; }
                    else if (object.ReferenceEquals(mt, typeof(int))) { row[c] = 0; }
                    else if (object.ReferenceEquals(mt, typeof(short))) { row[c] = 0; }
                    else if (object.ReferenceEquals(mt, typeof(double))) { row[c] = 0; }
                    else if (object.ReferenceEquals(mt, typeof(string))) { row[c] = string.Empty; }
                    else if (object.ReferenceEquals(mt, typeof(Guid))) { row[c] = Guid.Empty; }
                    else if (object.ReferenceEquals(mt, typeof(bool))) { row[c] = false; }
                    else if (object.ReferenceEquals(mt, typeof(byte))) { row[c] = 0; }
                    else if (object.ReferenceEquals(mt, typeof(decimal))) { row[c] = 0; }
                    else if (object.ReferenceEquals(mt, typeof(System.DateTime))) { row[c] = new System.DateTime(1900, 1, 1, 0, 0, 0); }
                    else { row[c] = null; }
                }
            }
        }
    } 
}