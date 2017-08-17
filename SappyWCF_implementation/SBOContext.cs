using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using System.Data;

class SBOContext : IDisposable
{
    private SAPbobsCOM.Company company = null;

    internal SBOContext(string companydb)
    {
        //if (Thread.CurrentThread.GetApartmentState().ToString() != "STA")
        //{
        //    throw new Exception("Esta função usa componentes COM (SAP DI API), a tread em que é chamada tem que estar em no modo STA.");
        //}

        this.company = new SAPbobsCOM.Company();
        this.company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
        this.company.Server = SappyWCF_implementation.Properties.Settings.Default.DBSERVER;
        this.company.LicenseServer = SappyWCF_implementation.Properties.Settings.Default.LICENCESERVER; 
        this.company.CompanyDB = companydb;
        this.company.UseTrusted = false;
        this.company.UserName = "dora";
        this.company.Password = "dadobia";
        this.company.language = SAPbobsCOM.BoSuppLangs.ln_English;
        this.company.DbUserName = SappyWCF_implementation.Properties.Settings.Default.DBUSER;
        this.company.DbPassword = SappyWCF_implementation.Properties.Settings.Default.DBUSERPASS;

        Logger.Log.Debug("Connecting SAPB1 DIAPI to " + this.company.Server + ", database " + this.company.CompanyDB + "...");
        if (this.company.Connect() != 0)
        {
            int errCode = 0;
            string errMsg = string.Empty;
            this.company.GetLastError(out errCode, out errMsg);

            var ex = new Exception("Não foi possível ligar ao SAP B1 pela DI API.\n" + errCode + " - " + errMsg);
            Logger.Log.Error("Connecting to " + this.company.Server + ", database " + this.company.CompanyDB + "...", ex);

            this.company = null;
            throw ex;
        }
        Logger.Log.Info("Connected SAPB1 DIAPI to " + this.company.Server + ", database " + this.company.CompanyDB + "...");
    }

    public void Dispose()
    {
        if (this.company != null)
        {
            Logger.Log.Info("Disconnecting SAPB1 DIAPI from " + this.company.Server + ", database " + this.company.CompanyDB + "...");
            try { if (this.company.Connected == true) this.company.Disconnect(); }
            catch (Exception) { }
            finally { Marshal.ReleaseComObject(this.company); }
        }
    }

    internal int Confirmar_SAPPY_DOC(string objCode, int draftId)
    {

        using (HelperOdbc dataLayer = new HelperOdbc())
        using (DataTable headerDt = dataLayer.Execute("SELECT * FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC WHERE ID =" + draftId))
        using (DataTable detailsDt = dataLayer.Execute("SELECT * FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES WHERE ID =" + draftId + " ORDER BY LINENUM"))
        {
            DataRow header = headerDt.Rows[0];

            int objType = Convert.ToInt32(header["OBJTYPE"]);
            if (objType.ToString() != objCode) throw new Exception("objCode " + objCode + " is diferent of " + objType.ToString());


            DateTime DOCDATE = (DateTime)header["DOCDATE"];
            DateTime DOCDUEDATE = (DateTime)header["DOCDUEDATE"];

            SAPbobsCOM.Documents newDoc = (SAPbobsCOM.Documents)this.company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)objType);
            newDoc.DocDate = DateTime.Now;                      // DataLancamento;
            if (DOCDATE.Year > 1900) newDoc.TaxDate = DOCDATE;       // DataDocumento;
            if (DOCDUEDATE.Year > 1900) newDoc.DocDueDate = DOCDUEDATE; // DataEntrega;
            newDoc.CardCode = (string)header["CARDCODE"];
            newDoc.Comments = (string)header["COMMENTS"];
            newDoc.NumAtCard = (string)header["NUMATCARD"];
            newDoc.UserFields.Fields.Item("U_apyUSER").Value = (string)header["CREATED_BY_NAME"];

            foreach (DataRow line in detailsDt.Rows)
            {
                if (newDoc.Lines.ItemCode != "") newDoc.Lines.Add();

                newDoc.Lines.ItemCode = (string)line["ITEMCODE"];
                newDoc.Lines.Factor1 = (double)(decimal)line["QTCX"];   // Num caixas/pack
                newDoc.Lines.Factor2 = (double)(decimal)line["QTPK"];   // Qdd por Caixa/pack
                newDoc.Lines.UnitPrice = (double)(decimal)line["PRICE"];
                newDoc.Lines.UserFields.Fields.Item("U_apyUDISC").Value = (string)line["USER_DISC"];
                

                var BONUS_NAP = (short)line["BONUS_NAP"];
                if (BONUS_NAP == 1)
                {
                    // deixa que o SAP calcule o LineTotal na lona atual e na de Bonus
                    if (newDoc.Lines.ItemCode != "") newDoc.Lines.Add();

                    newDoc.Lines.ItemCode = "BONUS";
                    newDoc.Lines.Factor1 = -1 * (double)(decimal)line["QTCX"];   // Num caixas/pack
                    newDoc.Lines.Factor2 = (double)(decimal)line["QTPK"];   // Qdd por Caixa/pack
                    newDoc.Lines.UnitPrice = (double)(decimal)line["PRICE"];
                }
                else
                {
                    newDoc.Lines.LineTotal = (double)(decimal)line["LINETOTAL"];
                }

            }

            try
            {

                int docentry = 0;
                this.company.StartTransaction();

                if (newDoc.Add() != 0)
                {
                    throw new Exception("Não foi possível gravar em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                }
                else
                {
                    int.TryParse(this.company.GetNewObjectKey(), out docentry);
                }

                this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                //APGAR OS REGISTOS
                dataLayer.Execute("DELETE FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC WHERE ID =" + draftId);
                dataLayer.Execute("DELETE FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES WHERE ID =" + draftId);
                
                if (newDoc.GetByKey(docentry))
                {
                    return newDoc.DocNum;    
                }

                return docentry;
            }
            finally
            {
                if (this.company.InTransaction) this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }
        }
    }
}