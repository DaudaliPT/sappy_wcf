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
        this.company.UseTrusted = false;
        this.company.language = SAPbobsCOM.BoSuppLangs.ln_English;
        this.company.LicenseServer = SappyWCF_implementation.Properties.Settings.Default.LICENCESERVER;
        this.company.Server = SappyWCF_implementation.Properties.Settings.Default.DBSERVER;
        this.company.DbUserName = SappyWCF_implementation.Properties.Settings.Default.DBUSER;
        this.company.DbPassword = SappyWCF_implementation.Properties.Settings.Default.DBUSERPASS;
        this.company.CompanyDB = companydb;
        this.company.UserName = SappyWCF_implementation.Properties.Settings.Default.SAPB1USER;
        this.company.Password = SappyWCF_implementation.Properties.Settings.Default.SAPB1USERPASS;

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

    internal AddDocResult Confirmar_SAPPY_DOC(string objCode, int draftId, double expectedTotal)
    {
        var sqlHeader = "SELECT T0.* ";
        sqlHeader += "\n , OCPR.\"CntctCode\"";
        sqlHeader += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC T0";
        sqlHeader += "\n LEFT JOIN \"" + this.company.CompanyDB + "\".OCPR OCPR";
        sqlHeader += "\n        ON T0.CARDCODE = OCPR.\"CardCode\"";
        sqlHeader += "\n       AND T0.CONTACT  = OCPR.\"Name\"";
        sqlHeader += "\n WHERE T0.ID =" + draftId;

        var sqlDetail = "SELECT T1.*";
        sqlDetail += "\n , OITM.\"InvntryUom\"";
        sqlDetail += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES T1";
        sqlDetail += "\n INNER JOIN \"" + this.company.CompanyDB + "\".OITM OITM on T1.ITEMCODE = OITM.\"ItemCode\"";
        sqlDetail += "\n WHERE T1.ID =" + draftId;
        sqlDetail += "\n ORDER BY T1.LINENUM";

        using (HelperOdbc dataLayer = new HelperOdbc())
        using (DataTable headerDt = dataLayer.Execute(sqlHeader))
        using (DataTable detailsDt = dataLayer.Execute(sqlDetail))
        {
            DataRow header = headerDt.Rows[0];

            int objType = Convert.ToInt32(header["OBJTYPE"]);
            if (objType.ToString() != objCode) throw new Exception("objCode " + objCode + " is diferent of " + objType.ToString());


            SAPbobsCOM.Documents newDoc = (SAPbobsCOM.Documents)this.company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)objType);

            //Serie
            newDoc.Series = (int)header["DOCSERIES"];

            // DataDocumento;
            DateTime DOCDATE = (DateTime)header["DOCDATE"];
            if (DOCDATE.Year > 1900) newDoc.TaxDate = DOCDATE;

            // Data vencimento/entrega;
            DateTime DOCDUEDATE = (DateTime)header["DOCDUEDATE"];
            if (DOCDUEDATE.Year > 1900) newDoc.DocDueDate = DOCDUEDATE;

            // Informações do Parceiro
            newDoc.CardCode = (string)header["CARDCODE"];
            if ((string)header["SHIPADDR"] != "") newDoc.ShipToCode = (string)header["SHIPADDR"];
            if ((string)header["BILLADDR"] != "") newDoc.PayToCode = (string)header["BILLADDR"];
            if ((string)header["NUMATCARD"] != "") newDoc.NumAtCard = (string)header["NUMATCARD"];
            if ((string)header["COMMENTS"] != "") newDoc.Comments = (string)header["COMMENTS"];
            if ((int)header["CntctCode"] != 0) newDoc.ContactPersonCode = (int)header["CntctCode"];

            newDoc.UserFields.Fields.Item("U_apyUSER").Value = (string)header["CREATED_BY_NAME"];
            newDoc.UserFields.Fields.Item("U_apyINCONF").Value = (short)header["HASINCONF"] == 1 ? "Y" : "N";

            foreach (DataRow line in detailsDt.Rows)
            {
                if (newDoc.Lines.ItemCode != "") newDoc.Lines.Add();

                newDoc.Lines.ItemCode = (string)line["ITEMCODE"];
                newDoc.Lines.ItemDescription = (string)line["ITEMNAME"];

                //Quantidades
                newDoc.Lines.MeasureUnit = (string)line["InvntryUom"];
                newDoc.Lines.Factor1 = (double)(decimal)line["QTCX"];   // Num caixas/pack
                newDoc.Lines.Factor2 = (double)(decimal)line["QTPK"];   // Qdd por Caixa/pack

                //Preço
                newDoc.Lines.UnitPrice = (double)(decimal)line["PRICE"];
                newDoc.Lines.DiscountPercent = (double)(decimal)line["DISCOUNT"];
                newDoc.Lines.UserFields.Fields.Item("U_apyUDISC").Value = (string)line["USER_DISC"];

                //Iva
                newDoc.Lines.TaxCode = (string)line["VATGROUP"];
                //     newDoc.Lines.VatGroup = (string)line["VATGROUP"];
                newDoc.Lines.UserFields.Fields.Item("U_apyINCONF").Value = (short)line["HASINCONF"] == 1 ? "Y" : "N";


                var BONUS_NAP = (short)line["BONUS_NAP"];
                if (BONUS_NAP == 1)
                {
                    // deixa que o SAP calcule o LineTotal na linha atual e na de Bonus
                    if (newDoc.Lines.ItemCode != "") newDoc.Lines.Add();

                    newDoc.Lines.ItemCode = "BONUS";
                    newDoc.Lines.ItemDescription = (string)line["ITEMNAME"];
                    newDoc.Lines.Factor1 = -1 * (double)(decimal)line["QTCX"];      // Num caixas/pack
                    newDoc.Lines.Factor2 = (double)(decimal)line["QTPK"];           // Qdd por Caixa/pack
                    newDoc.Lines.UnitPrice = (double)(decimal)line["PRICE"];
                    newDoc.Lines.TaxCode = (string)line["VATGROUP"];
                    // newDoc.Lines.VatGroup = (string)line["VATGROUP"];
                }
                else
                {
                    // Estamos a deixar o SAP calcular o total.
                    // fazemos isso principalmente porque isso fará o SAP calcular a percentagem de desconto e a base de imposto será seguramente a mesma, 
                    // sem diferenças causadas por arredondamentos.
                    // Isso pode causar % de descontos muitos pequenas e até negativas (mas sem significado nos valores)
                    // newDoc.Lines.LineTotal = (double)(decimal)line["LINETOTAL"];
                }
            }




            // ARREDONDAMENTO DE TOTAL
            if ((double)(decimal)header["ROUNDVAL"] != 0)
            {
                newDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES;
                newDoc.RoundingDiffAmount = (double)(decimal)header["ROUNDVAL"];
            }


            try
            {
                this.company.StartTransaction();

                if (newDoc.Add() != 0)
                {
                    var ex = new Exception("Não foi possível gravar em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                    //log the xml to allow easier debug
                    var xml = newDoc.GetAsXML();
                    Logger.Log.Debug(xml, ex);

                    throw ex;
                }

                int docentry = 0;
                int.TryParse(this.company.GetNewObjectKey(), out docentry);


                if (newDoc.GetByKey(docentry) == false)
                {
                    throw new Exception("Não foi obter o documento criado em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                }

                AddDocResult result = new AddDocResult();
                if (newDoc.DocTotal != expectedTotal)
                {
                    result.DocTotal = newDoc.DocTotal;
                    result.message = "(TOTALDIF) O total não é o esperado.";
                    return result;
                }

                this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                //APGAR OS REGISTOS
                dataLayer.Execute("DELETE FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC WHERE ID =" + draftId);
                dataLayer.Execute("DELETE FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES WHERE ID =" + draftId);

                result.DocEntry = docentry;
                result.DocNum = newDoc.DocNum;
                result.DocTotal = newDoc.DocTotal;
                return result;
            }
            finally
            {
                if (this.company.InTransaction) this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }
        }
    }
}