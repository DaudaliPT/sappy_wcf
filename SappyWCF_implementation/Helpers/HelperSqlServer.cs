using System;
using System.Text;
using System.Data.SqlClient;
using System.Data;

class HelperSqlServer : IDisposable
{
    public void Dispose()
    {
        if (db != null) db.Dispose();
    }

    internal SqlConnection db { get; set; } 

    public HelperSqlServer()
    {
        try
        {
            SqlConnectionStringBuilder cnnBuilder = new SqlConnectionStringBuilder();

            cnnBuilder.DataSource = SappyWCF_implementation.Properties.Settings.Default.DBSERVER;
            cnnBuilder.UserID = SappyWCF_implementation.Properties.Settings.Default.DBUSER;
            cnnBuilder.Password = SappyWCF_implementation.Properties.Settings.Default.DBUSERPASS;




            var ConnectionString = cnnBuilder.ToString();
            db = new SqlConnection(ConnectionString);
            db.Open();
        }
        catch (Exception)
        {
            Logger.Log.Debug("HelperSqlServer(" + SappyWCF_implementation.Properties.Settings.Default.DBSERVER + "): Não foi possivel establecer conexão.");
            throw;
        }
    }

    public HelperSqlServer(string server, string user, string password, string database)
    { 
        try
        {
            SqlConnectionStringBuilder cnnBuilder = new SqlConnectionStringBuilder();
            
            cnnBuilder.DataSource = server;
            cnnBuilder.InitialCatalog = database;
            cnnBuilder.UserID = user;
            cnnBuilder.Password = password;



            var ConnectionString = cnnBuilder.ToString();
            db = new SqlConnection(ConnectionString);
            db.Open();
        }
        catch (Exception)
        {
            Logger.Log.Debug("HelperSqlServer(" + server + ", " + database + "): Não foi possivel establecer conexão.");
            throw;
        }
    }

    internal DataTable Execute(string sqlQuery, bool CleanNulls = true)
    {
        DataTable dt = new DataTable();
        using (SqlCommand cmd = db.CreateCommand())
        {
            try
            {
                cmd.CommandText = sqlQuery;

                SqlDataAdapter adpt = new SqlDataAdapter(cmd);
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
        using (SqlCommand cmd = db.CreateCommand())
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

    internal SqlDataReader ExecuteReader(string sqlQuery)
    {
        DataTable dt = new DataTable();
        using (SqlCommand cmd = db.CreateCommand())
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
        using (SqlCommand cmd = db.CreateCommand())
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