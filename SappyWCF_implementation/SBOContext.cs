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
        this.company.language = SAPbobsCOM.BoSuppLangs.ln_Portuguese;
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

    internal AddDocResult SAPDOC_FROM_SAPPY_DRAFT(DocActions action, string objCode, int draftId, double expectedTotal)
    {

        int priceDecimals = 6;
        {
            var s = this.company.GetCompanyService();
            var ai = s.GetAdminInfo();
            priceDecimals = ai.PriceAccuracy;
        }

        var sqlHeader = "SELECT T0.* ";
        sqlHeader += "\n , OCPR.\"CntctCode\"";
        sqlHeader += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC T0";
        sqlHeader += "\n LEFT JOIN \"" + this.company.CompanyDB + "\".OCPR OCPR";
        sqlHeader += "\n        ON T0.CARDCODE = OCPR.\"CardCode\"";
        sqlHeader += "\n       AND T0.CONTACT  = OCPR.\"Name\"";
        sqlHeader += "\n WHERE T0.ID =" + draftId;

        var sqlDetail = "SELECT T1.*";
        sqlDetail += "\n , OITM.\"InvntryUom\"";
        sqlDetail += "\n , OITW.\"OnHand\"";
        sqlDetail += "\n , OITW.\"AvgPrice\"";
        sqlDetail += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES T1";
        sqlDetail += "\n INNER JOIN \"" + this.company.CompanyDB + "\".OITM OITM on T1.ITEMCODE = OITM.\"ItemCode\"";
        sqlDetail += "\n LEFT JOIN \"" + this.company.CompanyDB + "\".OITW OITW on T1.ITEMCODE = OITW.\"ItemCode\" AND T1.WHSCODE = OITW.\"WhsCode\"";
        sqlDetail += "\n WHERE T1.ID =" + draftId;
        sqlDetail += "\n ORDER BY T1.LINENUM";

        var sqlDetailNET = "SELECT T1.ITEMCODE, T1.WHSCODE, SUM(T1.QTSTK) AS QTSTK";
        sqlDetailNET += "\n , SUM(T1.LINETOTAL + COALESCE(CASE WHEN T1.BONUS_NAP=1 THEN T1.LINETOTALBONUS ELSE 0 END,0) ) AS TRANSCOST";
        sqlDetailNET += "\n , SUM(T1.NETTOTAL+ COALESCE(CASE WHEN T1.BONUS_NAP=1 THEN T1.NETTOTALBONUS ELSE 0 END,0) ) AS TRANSCOSTNET";
        sqlDetailNET += "\n , OITW.\"OnHand\"";
        sqlDetailNET += "\n , OITW.\"AvgPrice\"";
        sqlDetailNET += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES T1";
        sqlDetailNET += "\n INNER JOIN \"" + this.company.CompanyDB + "\".OITM OITM on T1.ITEMCODE = OITM.\"ItemCode\"";
        sqlDetailNET += "\n LEFT JOIN \"" + this.company.CompanyDB + "\".OITW OITW on T1.ITEMCODE = OITW.\"ItemCode\" AND T1.WHSCODE = OITW.\"WhsCode\"";
        sqlDetailNET += "\n WHERE T1.ID =" + draftId;
        sqlDetailNET += "\n GROUP BY T1.ITEMCODE, T1.WHSCODE";
        sqlDetailNET += "\n , OITW.\"OnHand\"";
        sqlDetailNET += "\n , OITW.\"AvgPrice\"";
        sqlDetailNET += "\n ORDER BY MIN(T1.LINENUM)";

        using (HelperOdbc dataLayer = new HelperOdbc())
        using (DataTable headerDt = dataLayer.Execute(sqlHeader))
        using (DataTable detailsDt = dataLayer.Execute(sqlDetail))
        using (DataTable detailsDtNET = dataLayer.Execute(sqlDetailNET))
        {
            DataRow header = headerDt.Rows[0];

            int objType = Convert.ToInt32(header["OBJTYPE"]);
            bool sujRevalorizacaoNET = (action == DocActions.ADD && objType == 18);

            if (objType.ToString() != objCode) throw new Exception("objCode " + objCode + " is diferent of " + objType.ToString());

            SAPbobsCOM.Documents newDoc = (SAPbobsCOM.Documents)this.company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)objType);
            newDoc.Series = (int)header["DOCSERIES"];
            DateTime TAXDATE = (DateTime)header["TAXDATE"];
            DateTime DOCDUEDATE = (DateTime)header["DOCDUEDATE"];
            if (TAXDATE.Year > 1900) newDoc.TaxDate = TAXDATE;
            if (DOCDUEDATE.Year > 1900) newDoc.DocDueDate = DOCDUEDATE;

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
                var QTCX = (double)(decimal)line["QTCX"];   // Num caixas/pack
                var QTPK = (double)(decimal)line["QTPK"];
                var QTSTK = (double)(decimal)line["QTSTK"];
                var QTBONUS = (double)(decimal)line["QTBONUS"];
                var BONUS_NAP = (short)line["BONUS_NAP"];

                if (QTSTK != 0)
                {
                    if (newDoc.Lines.ItemCode != "") newDoc.Lines.Add();
                    newDoc.Lines.ItemCode = (string)line["ITEMCODE"];
                    newDoc.Lines.ItemDescription = (string)line["ITEMNAME"];
                    newDoc.Lines.MeasureUnit = (string)line["InvntryUom"];
                    newDoc.Lines.Factor1 = QTCX;   // Num caixas/pack
                    newDoc.Lines.Factor2 = QTPK;   // Qdd por Caixa/pack 
                    //newDoc.Lines.InventoryQuantity = (double)(decimal)line["QTSTK"]; //Definir sobrepoe os fatores 1 e 2
                    newDoc.Lines.UnitPrice = (double)(decimal)line["PRICE"];
                    newDoc.Lines.WarehouseCode = (string)line["WHSCODE"];
                    //newDoc.Lines.VatGroup = (string)line["VATGROUP"];
                    newDoc.Lines.TaxCode = (string)line["VATGROUP"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyINCONF").Value = (short)line["HASINCONF"] == 1 ? "Y" : "N";

                    // Estes campos atualmente estão ao nivel de cabeçalho, mas são guardados no documento nas linhas,
                    // porque preve-se que no futuro esta tenha que ser uma informação linha a linha.
                    newDoc.Lines.UserFields.Fields.Item("U_apyDFIN").Value = (string)header["DESCFIN"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyDDEB").Value = (string)header["DESCDEB"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyDFINAC").Value = (short)header["DESCFINAC"] == 1 ? "Y" : "N";
                    newDoc.Lines.UserFields.Fields.Item("U_apyDDEBAC").Value = (short)header["DESCDEBAC"] == 1 ? "Y" : "N";
                    newDoc.Lines.UserFields.Fields.Item("U_apyDDEBPER").Value = (string)header["DESCDEBPER"];

                    newDoc.Lines.UserFields.Fields.Item("U_apyPRCNET").Value = (double)(decimal)line["NETPRICE"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyNETTOT").Value = (double)(decimal)line["NETTOTAL"];

                    newDoc.Lines.UserFields.Fields.Item("U_apyUDISC").Value = (string)line["USER_DISC"];
                    newDoc.Lines.LineTotal = (double)(decimal)line["LINETOTAL"];
                }
                if (QTBONUS != 0)
                {
                    //    double NRCXBONUS = QTBONUS;
                    //    double QTPKBONUS = 1;

                    //    if (QTBONUS % QTPK == 0)
                    //    {
                    //        // Usar grupagem quando é possível
                    //        NRCXBONUS = QTBONUS / QTPK;
                    //        QTPKBONUS = QTPK;
                    //    }


                    if (newDoc.Lines.ItemCode != "") newDoc.Lines.Add();
                    newDoc.Lines.ItemCode = (string)line["ITEMCODE"];
                    newDoc.Lines.ItemDescription = (string)line["ITEMNAME"];
                    newDoc.Lines.MeasureUnit = (string)line["InvntryUom"];
                    newDoc.Lines.Quantity = QTBONUS;
                    newDoc.Lines.UnitPrice = (double)(decimal)line["PRICE"];
                    //newDoc.Lines.VatGroup = (string)line["VATGROUP"];
                    newDoc.Lines.TaxCode = (string)line["VATGROUP"];
                    if (QTSTK != 0) newDoc.Lines.UserFields.Fields.Item("U_apyREFLIN").Value = newDoc.Lines.Count - 2;

                    newDoc.Lines.UserFields.Fields.Item("U_apyINCONF").Value = (short)line["HASINCONF"] == 1 ? "Y" : "N";

                    // Estes campos atualmente estão ao nivel de cabeçalho, mas são guardados no documento nas linhas,
                    // porque preve-se que no futuro esta tenha que ser uma informação linha a linha.
                    newDoc.Lines.UserFields.Fields.Item("U_apyDFIN").Value = (string)header["DESCFIN"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyDDEB").Value = (string)header["DESCDEB"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyDFINAC").Value = (short)header["DESCFINAC"] == 1 ? "Y" : "N";
                    newDoc.Lines.UserFields.Fields.Item("U_apyDDEBAC").Value = (short)header["DESCDEBAC"] == 1 ? "Y" : "N";
                    newDoc.Lines.UserFields.Fields.Item("U_apyDDEBPER").Value = (string)header["DESCDEBPER"];

                    if (BONUS_NAP != 1)
                    {
                        newDoc.Lines.UserFields.Fields.Item("U_apyUDISC").Value = "BONUS";
                        newDoc.Lines.DiscountPercent = 100;
                        newDoc.Lines.UserFields.Fields.Item("U_apyPRCNET").Value = 0;
                        newDoc.Lines.UserFields.Fields.Item("U_apyNETTOT").Value = 0;
                    }
                    else
                    {
                        newDoc.Lines.LineTotal = (double)(decimal)line["LINETOTALBONUS"];
                        newDoc.Lines.UserFields.Fields.Item("U_apyUDISC").Value = (string)line["USER_DISC"];
                        newDoc.Lines.UserFields.Fields.Item("U_apyPRCNET").Value = (double)(decimal)line["NETPRICE"];
                        newDoc.Lines.UserFields.Fields.Item("U_apyNETTOT").Value = (double)(decimal)line["NETTOTALBONUS"];

                        if (newDoc.Lines.ItemCode != "") newDoc.Lines.Add();
                        newDoc.Lines.ItemCode = "BONUS";
                        newDoc.Lines.ItemDescription = (string)line["ITEMNAME"];
                        newDoc.Lines.MeasureUnit = (string)line["InvntryUom"];
                        newDoc.Lines.Quantity = -1 * QTBONUS;
                        newDoc.Lines.UnitPrice = (double)(decimal)line["PRICE"];
                        //newDoc.Lines.VatGroup = (string)line["VATGROUP"];
                        newDoc.Lines.TaxCode = (string)line["VATGROUP"];
                        newDoc.Lines.UserFields.Fields.Item("U_apyUDISC").Value = (string)line["USER_DISC"];
                        newDoc.Lines.LineTotal = -1 * (double)(decimal)line["LINETOTALBONUS"];
                        if (QTSTK != 0)
                            newDoc.Lines.UserFields.Fields.Item("U_apyREFLIN").Value = newDoc.Lines.Count - 3; //manter refrêmncia com a linha principal
                        else
                            newDoc.Lines.UserFields.Fields.Item("U_apyREFLIN").Value = newDoc.Lines.Count - 2; // se é só oferta a linha é a da oferta

                        newDoc.Lines.UserFields.Fields.Item("U_apyINCONF").Value = (short)line["HASINCONF"] == 1 ? "Y" : "N";

                        // Estes campos atualmente estão ao nivel de cabeçalho, mas são guardados no documento nas linhas,
                        // porque preve-se que no futuro esta tenha que ser uma informação linha a linha.
                        newDoc.Lines.UserFields.Fields.Item("U_apyDFIN").Value = (string)header["DESCFIN"];
                        newDoc.Lines.UserFields.Fields.Item("U_apyDDEB").Value = (string)header["DESCDEB"];
                        newDoc.Lines.UserFields.Fields.Item("U_apyDFINAC").Value = (short)header["DESCFINAC"] == 1 ? "Y" : "N";
                        newDoc.Lines.UserFields.Fields.Item("U_apyDDEBAC").Value = (short)header["DESCDEBAC"] == 1 ? "Y" : "N";
                        newDoc.Lines.UserFields.Fields.Item("U_apyDDEBPER").Value = (string)header["DESCDEBPER"];

                        newDoc.Lines.UserFields.Fields.Item("U_apyPRCNET").Value = (double)(decimal)line["NETPRICE"];
                        newDoc.Lines.UserFields.Fields.Item("U_apyNETTOT").Value = -1 * (double)(decimal)line["NETTOTALBONUS"];
                    }

                }

            }

            // ARREDONDAMENTO DE TOTAL
            if ((double)(decimal)header["ROUNDVAL"] != 0)
            {
                newDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES;
                newDoc.RoundingDiffAmount = (double)(decimal)header["ROUNDVAL"];
            }

            if ((string)header["FORCEFIELD"] == "EXTRADISC" ||
                (string)header["FORCEFIELD"] == "EXTRADISCPERC") newDoc.DiscountPercent = (double)(decimal)header["EXTRADISCPERC"];
            if ((string)header["FORCEFIELD"] == "DOCTOTAL") newDoc.DocTotal = (double)(decimal)header["DOCTOTAL"];





            //documentos para revalorização NET
            // *******************************************************************************************************************************
            SAPbobsCOM.Documents invEntry = null;
            SAPbobsCOM.MaterialRevaluation invReval = null;
            SAPbobsCOM.Documents invExit = null;
            bool hasRevalorizacao = false;
            bool hasFakeEntryExit = false;
            if (sujRevalorizacaoNET)
            {
                invEntry = this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                invExit = this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                invReval = this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation);
                invReval.RevalType = "M"; // Débito/Crédito

                if (TAXDATE.Year > 1900)
                {
                    invEntry.TaxDate = TAXDATE;
                    invReval.TaxDate = TAXDATE;
                    invExit.TaxDate = TAXDATE;
                }

                foreach (DataRow line in detailsDtNET.Rows)
                {
                    var transQty = (decimal)line["QTSTK"];
                    var onHand = (decimal)line["OnHand"];
                    var avgPrice = (decimal)line["AvgPrice"];
                    var transCost = (decimal)line["TRANSCOST"];
                    var transCostNet = (decimal)line["TRANSCOSTNET"];
                    var finalOnHand = onHand + transQty;


                    if (onHand < 0)
                    {
                        hasFakeEntryExit = true;

                        decimal finalStkValue = onHand * avgPrice + transCostNet;
                        decimal finalPmc = 0;
                        if (finalOnHand > 0) finalPmc = finalStkValue / finalOnHand;

                        if (onHand <= 0)
                        {
                            decimal docPriceNet = 0;
                            if (transQty != 0) docPriceNet = transCostNet / transQty;
                            finalPmc = docPriceNet;
                        }
                        // add fake stock to make it zero, before
                        if (invEntry.Lines.ItemCode != "") invEntry.Lines.Add();
                        invEntry.Lines.ItemCode = (string)line["ITEMCODE"];
                        invEntry.Lines.WarehouseCode = (string)line["WHSCODE"];
                        invEntry.Lines.InventoryQuantity = (double)onHand * -1;
                        invEntry.Lines.UnitPrice = (double)finalPmc;

                        // remove fake stock, so qty stays correct, after
                        if (invExit.Lines.ItemCode != "") invExit.Lines.Add();
                        invExit.Lines.ItemCode = (string)line["ITEMCODE"];
                        invExit.Lines.WarehouseCode = (string)line["WHSCODE"];
                        invExit.Lines.InventoryQuantity = (double)onHand * -1;
                    }

                    if (transCost != transCostNet && finalOnHand > 0)
                    {
                        hasRevalorizacao = true;
                        if (invReval.Lines.ItemCode != "") invReval.Lines.Add();
                        invReval.Lines.ItemCode = (string)line["ITEMCODE"];
                        invReval.Lines.WarehouseCode = (string)line["WHSCODE"];
                        invReval.Lines.DebitCredit = (double)(transCostNet - transCost);
                        invReval.Lines.Quantity = (double)finalOnHand;
                    }
                }
            }




            try
            {
                this.company.StartTransaction();

                if (hasFakeEntryExit)
                {
                    // invEntry.Reference2 = newDoc.DocNum.ToString();              //preenchido mais abaixo
                    // invEntry.Comments = "Referente a " + newDoc.JournalMemo;     //preenchido mais abaixo
                    if (invEntry.Add() != 0)
                    {
                        var ex = new Exception("Não foi possível gravar fakeEntry em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                        //log the xml to allow easier debug
                        var xml = invEntry.GetAsXML();
                        Logger.Log.Debug(xml, ex);

                        throw ex;
                    }

                    int docentry = 0;
                    int.TryParse(this.company.GetNewObjectKey(), out docentry);

                    if (invEntry.GetByKey(docentry) == false)
                    {
                        throw new Exception("Não foi obter o fakeEntry criada em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    }
                }

                if (newDoc.Add() != 0)
                {
                    var ex = new Exception("Não foi possível gravar em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                    //log the xml to allow easier debug
                    var xml = newDoc.GetAsXML();
                    Logger.Log.Debug(xml, ex);

                    throw ex;
                }
                else
                {
                    int docentry = 0;
                    int.TryParse(this.company.GetNewObjectKey(), out docentry);

                    if (newDoc.GetByKey(docentry) == false)
                    {
                        throw new Exception("Não foi obter o documento criado em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    }
                }

                AddDocResult result = new AddDocResult();
                if (action == DocActions.ADD && newDoc.DocTotal != expectedTotal)
                {
                    //log the xml to allow easier debug
                    var xml = newDoc.GetAsXML();
                    Logger.Log.Debug(xml);

                    result.DocTotal = newDoc.DocTotal;
                    result.message = "(TOTALDIF) O total não é o esperado.";
                    return result;
                }

                if (hasFakeEntryExit)
                {
                    //Atualizar as referências na entrada de stock
                    invEntry.Reference2 = newDoc.DocNum.ToString();
                    invEntry.Comments = "Referente a " + newDoc.JournalMemo;
                    if (invEntry.Update() != 0)
                    {
                        var ex = new Exception("Não foi possível atualizar fakeEntry em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                        //log the xml to allow easier debug
                        var xml = invEntry.GetAsXML();
                        Logger.Log.Debug(xml, ex);

                        throw ex;
                    }
                }

                if (hasRevalorizacao)
                {
                    invReval.Reference2 = newDoc.DocNum.ToString();
                    invReval.Comments = "Referente a " + newDoc.JournalMemo;
                    if (invReval.Add() != 0)
                    {
                        var ex = new Exception("Não foi possível gravar revalorização em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                        //log the xml to allow easier debug
                        var xml = invReval.GetAsXML();
                        Logger.Log.Debug(xml, ex);

                        throw ex;
                    }
                }

                if (hasFakeEntryExit)
                {
                    invExit.Reference2 = newDoc.DocNum.ToString();
                    invExit.Comments = "Referente a " + newDoc.JournalMemo;
                    if (invExit.Add() != 0)
                    {
                        var ex = new Exception("Não foi possível gravar fakeExit em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                        //log the xml to allow easier debug
                        var xml = invExit.GetAsXML();
                        Logger.Log.Debug(xml, ex);

                        throw ex;
                    }
                }

                if (action == DocActions.ADD)
                {
                    this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                    //APGAR OS REGISTOS
                    dataLayer.Execute("DELETE FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC WHERE ID =" + draftId);
                    dataLayer.Execute("DELETE FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES WHERE ID =" + draftId);
                }
                result.DocEntry = newDoc.DocEntry;
                result.DocNum = newDoc.DocNum;

                result.DocTotal = newDoc.DocTotal;
                result.DiscountPercent = newDoc.DiscountPercent;
                result.TotalDiscount = newDoc.TotalDiscount;
                result.VatSum = newDoc.VatSum;
                result.RoundingDiffAmount = newDoc.RoundingDiffAmount;
                return result;
            }
            finally
            {
                if (this.company.InTransaction) this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }
        }
    }

    internal AddDocResult SAPDOC_PATCH_WITH_SAPPY_CHANGES(string objCode, int docEntry)
    {

        int objType = Convert.ToInt32(objCode);
        if (objType.ToString() != objCode) throw new Exception("objCode " + objCode + " is diferent of " + objType.ToString());

        int priceDecimals = 6;
        {
            var s = this.company.GetCompanyService();
            var ai = s.GetAdminInfo();
            priceDecimals = ai.PriceAccuracy;
        }

        var sqlHeader = "SELECT * ";
        sqlHeader += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_EDIT";
        sqlHeader += "\n WHERE OBJTYPE  ='" + objCode + "'";
        sqlHeader += "\n   AND DOCENTRY = " + docEntry;
        using (HelperOdbc dataLayer = new HelperOdbc())
        using (DataTable headerChanges = dataLayer.Execute(sqlHeader))
        {

            if (headerChanges.Rows.Count == 0)
                throw new Exception("Não foi há nenhuma alteração a fazer ao documento");


            SAPbobsCOM.Documents sapDoc = (SAPbobsCOM.Documents)this.company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)objType);

            if (sapDoc.GetByKey(docEntry) == false)
                throw new Exception("Não foi possível obter o documento do SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

            foreach (DataRow change in headerChanges.Rows)
            {
                string field = (string)change["FIELDNAME"];
                string value = (string)change["FIELDVALUE"];

                if (field == "DOCDUEDATE") sapDoc.DocDueDate = DateTime.ParseExact(value, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                if (field == "COMMENTS") sapDoc.Comments = value;
                if (field == "HASINCONF")sapDoc.UserFields.Fields.Item("U_apyINCONF").Value = (value!="" && "true,1".Contains(value.ToLower()) ? "Y" : "N");
            }


            try
            {
                this.company.StartTransaction();

                if (sapDoc.Update() != 0)
                {
                    var ex = new Exception("Não foi possível gravar em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                    //log the xml to allow easier debug
                    var xml = sapDoc.GetAsXML();
                    Logger.Log.Debug(xml, ex);

                    throw ex;
                }
                else
                {
                    if (sapDoc.GetByKey(docEntry) == false)
                    {
                        throw new Exception("Não foi obter o documento atualizado em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    }
                }

                AddDocResult result = new AddDocResult();


                this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                //APGAR OS REGISTOS DE MODIFICAÇÃO
                dataLayer.Execute("DELETE FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_EDIT WHERE OBJTYPE  ='" + objCode + "' AND DOCENTRY = " + docEntry);

                result.DocEntry = sapDoc.DocEntry;
                result.DocNum = sapDoc.DocNum;

                result.DocTotal = sapDoc.DocTotal;
                result.DiscountPercent = sapDoc.DiscountPercent;
                result.TotalDiscount = sapDoc.TotalDiscount;
                result.VatSum = sapDoc.VatSum;
                result.RoundingDiffAmount = sapDoc.RoundingDiffAmount;
                return result;
            }
            finally
            {
                if (this.company.InTransaction) this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }
        }
    }

}