using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Risa.ReplicateToHANA.Replication;

namespace Risa.ReplicateToHANA.RunContext
{
    class Context
    {
        private Context()
        {

        }

        public Context(string pSqlDatabaseName, string pHanaDatabaseName, string pHanaTablePrefix, int pDatabaseID, DatabaseStatusEnum pReplicationStatus)
        {
            SqlDatabaseName = pSqlDatabaseName;
            HanaDatabaseName = pHanaDatabaseName;
            HanaTablePrefix = pHanaTablePrefix;
            DatabaseID = pDatabaseID;
            replicationStatus = pReplicationStatus;

            this.DbServiceCNN = new RunContext.DBHelperServiceDB(); // Open connetion to ServiceDB 
            this.DbServiceCNNread = new RunContext.DBHelperServiceDB(); // Open connetion to ServiceDB 
        }

        public DBHelperServiceDB DbServiceCNN;
        public DBHelperServiceDB DbServiceCNNread;
        public DBHelperHANA DestHanaCNN;
        public DBHelperSQL SourceSqlCnn;

        public string SqlDatabaseName;
        public string HanaDatabaseName;
        public string HanaTablePrefix;
        public int DatabaseID;

        private DatabaseStatusEnum replicationStatus;
        public DatabaseStatusEnum ReplicationStatus
        {
            get { return replicationStatus; }

        }

        public string HanaFullName(string hanaTable)
        {
            return "\"" + this.HanaDatabaseName + "\"." + hanaTable;
        }

        public void SetTableStatus(string tableName, TableStatusEnum statusEnum, string StatusMsg = null, long changeTrakingVersion = -1000)
        {
            string sql = "UPDATE DatabaseTables ";
            sql += "\n SET StatusDate = GETDATE()";
            sql += "\n   , Status = " + (int)statusEnum;
            sql += "\n   , StatusMsg = " + (StatusMsg == null ? "null" : "'" + (StatusMsg.Length > 100 ? StatusMsg.Substring(0, 99) : StatusMsg) + "'");
            if (changeTrakingVersion != -1000)
            {
                sql += "\n   , ChangeTrackingLastVersion = " + changeTrakingVersion;
            }
            sql += "\n WHERE DatabaseID=" + this.DatabaseID + " and TableName = '" + tableName + "'";
            this.DbServiceCNN.Execute(sql);
        }

        public void SetDatabaseStatus(DatabaseStatusEnum value)
        {
            string sql = "UPDATE Databases SET ReplicationStatus = " + ((int)value).ToString() + " WHERE ID=" + this.DatabaseID;
            this.DbServiceCNN.Execute(sql);
            replicationStatus = value;
        }

        public Context CreateClone()
        {
            Context context = new Context(this.SqlDatabaseName, this.HanaDatabaseName, this.HanaTablePrefix, this.DatabaseID, this.replicationStatus);

            // Connect to source SQL
            context.SourceSqlCnn = new RunContext.DBHelperSQL(this.SourceSqlCnn.Server, this.SourceSqlCnn.User, this.SourceSqlCnn.Password, this.SourceSqlCnn.Database);

            // Connect to destination HANA
            context.DestHanaCNN = new RunContext.DBHelperHANA(this.DestHanaCNN.Server, this.DestHanaCNN.User, this.DestHanaCNN.Password);
            return context;
        }

        //public void StartTransactionOnALL()
        //{
        //    if (DbServiceCNN.currentTransaction != null || DestHanaCNN.currentTransaction != null || SourceSqlCnn.currentTransaction != null)
        //    {
        //        throw new Exception("Transaction is already in course");
        //    }

        //    SourceSqlCnn.currentTransaction = SourceSqlCnn.db.BeginTransaction(System.Data.IsolationLevel.Snapshot);
        //    DbServiceCNN.currentTransaction = DbServiceCNN.db.BeginTransaction();
        //    DestHanaCNN.currentTransaction = DestHanaCNN.db.BeginTransaction();
        //}
        //public void RoolbackTransactionOnALL()
        //{
        //    try
        //    {
        //        try { if (SourceSqlCnn.currentTransaction == null) Manager.Log.Error("RoolbackTransactionOnALL: No Transaction in course at SourceSqlCnn"); }
        //        finally { SourceSqlCnn.currentTransaction.Rollback(); }

        //        try { if (DbServiceCNN.currentTransaction == null) Manager.Log.Error("RoolbackTransactionOnALL: No Transaction in course at DbServiceCNN"); }
        //        finally { DbServiceCNN.currentTransaction.Rollback(); }

        //        try { if (DestHanaCNN.currentTransaction == null) Manager.Log.Error("RoolbackTransactionOnALL: No Transaction in course at DestHanaCNN"); }
        //        finally { DestHanaCNN.currentTransaction.Rollback(); }
        //    }
        //    finally
        //    {
        //        SourceSqlCnn.currentTransaction = null;
        //        DbServiceCNN.currentTransaction = null;
        //        DestHanaCNN.currentTransaction = null;
        //    }

        //}

        //public void CommitTransactionOnAll()
        //{
        //    try
        //    {
        //        try { if (SourceSqlCnn.currentTransaction == null) Manager.Log.Error("CommitTransactionOnALL: No Transaction in course at SourceSqlCnn"); }
        //        finally { SourceSqlCnn.currentTransaction.Commit(); }

        //        try { if (DbServiceCNN.currentTransaction == null) Manager.Log.Error("CommitTransactionOnALL: No Transaction in course at DbServiceCNN"); }
        //        finally { DbServiceCNN.currentTransaction.Commit(); }

        //        try { if (DestHanaCNN.currentTransaction == null) Manager.Log.Error("CommitTransactionOnALL: No Transaction in course at DestHanaCNN"); }
        //        finally { DestHanaCNN.currentTransaction.Commit(); }
        //    }
        //    finally
        //    {
        //        SourceSqlCnn.currentTransaction = null;
        //        DbServiceCNN.currentTransaction = null;
        //        DestHanaCNN.currentTransaction = null;
        //    }
        //}


    }
}
