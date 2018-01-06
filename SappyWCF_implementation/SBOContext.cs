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


    internal string GetLayoutCode(string TypeCode)
    {
        SAPbobsCOM.Recordset rec = null;
        try
        {
            SAPbobsCOM.CompanyService oCmpSrv = null;
            SAPbobsCOM.ReportLayoutsService oReportLayoutService = null;
            SAPbobsCOM.ReportParams oReportParam = null;
            SAPbobsCOM.DefaultReportParams oReportParaDefault = null;
            oCmpSrv = this.company.GetCompanyService();
            oReportLayoutService = (SAPbobsCOM.ReportLayoutsService)oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
            oReportParam = (SAPbobsCOM.ReportParams)oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportParams);
            oReportParam.ReportCode = TypeCode;
            oReportParam.UserID = this.company.UserSignature;
            oReportParaDefault = oReportLayoutService.GetDefaultReport(oReportParam);
            return oReportParaDefault.LayoutCode;
        }
        catch (Exception err)
        {
            throw new Exception("GetDocCode" + err.Message);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
            rec = null;
            System.GC.Collect();
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
        sqlDetail += "\n      , BASEDOC.\"CardCode\" AS BASE_CARDCODE";
        sqlDetail += "\n      , BASEDOC.QTYSTK_AVAILABLE_SAP";
        sqlDetail += "\n      , BASEDOC.\"DocStatus\" AS BASE_DOCSTATUS";
        sqlDetail += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES T1";
        sqlDetail += "\n INNER JOIN \"" + this.company.CompanyDB + "\".OITM OITM on T1.ITEMCODE = OITM.\"ItemCode\"";
        sqlDetail += "\n LEFT JOIN \"" + this.company.CompanyDB + "\".OITW OITW on T1.ITEMCODE = OITW.\"ItemCode\" AND T1.WHSCODE = OITW.\"WhsCode\"";
        sqlDetail += "\n LEFT JOIN \"" + this.company.CompanyDB + "\".SAPPY_LINE_LINK_" + objCode + " as BASEDOC ";
        sqlDetail += "\n                ON T1.BASE_OBJTYPE  = BASEDOC.\"ObjType\"";
        sqlDetail += "\n               AND T1.BASE_DOCENTRY = BASEDOC.\"DocEntry\"";
        sqlDetail += "\n               AND T1.BASE_LINENUM  = BASEDOC.\"LineNum\"";
        sqlDetail += "\n WHERE T1.ID =" + draftId;
        sqlDetail += "\n ORDER BY T1.LINENUM";

        var sqlDetailNET = "SELECT T1.ITEMCODE, T1.WHSCODE, SUM(T1.QTSTK) AS QTSTK";
        //O LINETOTAL2 contem iec, ecovalor e ecoree
        sqlDetailNET += "\n , SUM(T1.LINETOTAL2 + COALESCE(CASE WHEN T1.BONUS_NAP=1 THEN T1.LINETOTALBONUS ELSE 0 END,0) ) AS TRANSCOST";
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

        //// Validar os totais por Artigo relacionados com o documento Base
        //var sqlTotalByItem = "SELECT T1.ITEMCODE";
        //sqlTotalByItem += "\n      , max(T1.ITEMNAME) AS ITEMNAME";
        //sqlTotalByItem += "\n      , T1.BASE_OBJTYPE";
        //sqlTotalByItem += "\n      , T1.BASE_DOCENTRY";
        //sqlTotalByItem += "\n      , T1.BASE_LINENUM";
        //sqlTotalByItem += "\n      , BASEDOC.\"CardCode\"";
        //sqlTotalByItem += "\n      , BASEDOC.QTYSTK_AVAILABLE_SAP";
        //sqlTotalByItem += "\n      , SUM(T1.QTSTK) AS QTSTK";
        //sqlTotalByItem += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES T1";
        //sqlTotalByItem += "\n INNER JOIN \"" + this.company.CompanyDB + "\".SAPPY_LINE_LINK_" + objCode + " as BASEDOC ";
        //sqlTotalByItem += "\n                ON T1.BASE_OBJTYPE  = BASEDOC.\"ObjType\"";
        //sqlTotalByItem += "\n               AND T1.BASE_DOCENTRY = BASEDOC.\"DocEntry\"";
        //sqlTotalByItem += "\n               AND T1.BASE_LINENUM  = BASEDOC.\"LineNum\"";
        //sqlTotalByItem += "\n WHERE T1.ID =" + draftId;
        //sqlTotalByItem += "\n GROUP BY T1.ITEMCODE";
        //sqlTotalByItem += "\n      , T1.BASE_OBJTYPE";
        //sqlTotalByItem += "\n      , T1.BASE_DOCENTRY";
        //sqlTotalByItem += "\n      , T1.BASE_LINENUM";
        //sqlTotalByItem += "\n      , BASEDOC.\"CardCode\"";
        //sqlTotalByItem += "\n      , BASEDOC.\"QTYSTK_AVAILABLE_SAP\"";
        //sqlTotalByItem += "\n ORDER BY min(T1.LINENUM)";


        using (HelperOdbc dataLayer = new HelperOdbc())
        using (DataTable headerDt = dataLayer.Execute(sqlHeader))
        using (DataTable detailsDt = dataLayer.Execute(sqlDetail))
        using (DataTable detailsDtNET = dataLayer.Execute(sqlDetailNET))
        //using (DataTable itemTotalsDt = dataLayer.Execute(sqlTotalByItem))
        {
            DataRow header = headerDt.Rows[0];

            int module = Convert.ToInt32(header["MODULE"]);
            if (module != 0 && module != 1) throw new Exception("Esta função só pode ser chamada para MODULE IN (0,1)");

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
            newDoc.UserFields.Fields.Item("U_apyCONTRATO").Value = (int)header["CONTRATO"];

            if ("15,16,21".IndexOf(objCode) > -1)
            {
                if ((string)header["ATDOCTYPE"] != "") { newDoc.ATDocumentType = (string)header["ATDOCTYPE"]; }
                else
                {
                    if (objCode == "15") newDoc.ATDocumentType = "GR"; //Entrega a cliente
                    if (objCode == "16") newDoc.ATDocumentType = "GT"; //Devolução de cliente
                    if (objCode == "21") newDoc.ATDocumentType = "GD"; //Devolução a fornecedor
                }
                if ((string)header["MATRICULA"] != "") newDoc.VehiclePlate = (string)header["MATRICULA"];
                if ((string)header["ATAUTHCODE"] != "")
                {
                    newDoc.AuthorizationCode = (string)header["ATAUTHCODE"];
                    newDoc.ElecCommStatus = SAPbobsCOM.ElecCommStatusEnum.ecsApproved;
                }
                newDoc.StartDeliveryDate = DateTime.Now.AddMinutes(5);
                newDoc.StartDeliveryTime = DateTime.Now.AddMinutes(5);
            }

            if ((string)header["MATRICULA"] != "")
            {
                newDoc.UserFields.Fields.Item("U_apyMATRICULA").Value = (string)header["MATRICULA"];
            }

            newDoc.UserFields.Fields.Item("U_apyUSER").Value = (string)header["CREATED_BY_NAME"];
            newDoc.UserFields.Fields.Item("U_apyINCONF").Value = (short)header["HASINCONF"] == 1 ? "Y" : "N";



            //// preform checks by item total
            //foreach (DataRow itemTotal in itemTotalsDt.Rows)
            //{
            //    var ITEMCODE = (string)itemTotal["ITEMCODE"];
            //    var ITEMNAME = (string)itemTotal["ITEMNAME"];
            //    var QTSTK = (double)(decimal)itemTotal["QTSTK"];

            //    string CardCode = (string)itemTotal["CardCode"];
            //    if (CardCode != "" && newDoc.CardCode != CardCode) throw new Exception("Encontrada referência a documento de outra entidade: " + CardCode);

            //    double QTYSTK_AVAILABLE_SAP = (double)(decimal)itemTotal["QTYSTK_AVAILABLE_SAP"];
            //    if (QTSTK > QTYSTK_AVAILABLE_SAP) throw new Exception("Não pode relacionar mais que " + QTYSTK_AVAILABLE_SAP + " UN do artigo " + ITEMNAME + " com o documento base.");

            //}

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
                    newDoc.Lines.UnitPrice = (double)(decimal)line["PRICE"]
                    + ((double)(decimal)line["IEC"]
                    + (double)(decimal)line["ECOVALOR"]
                    + (double)(decimal)line["ECOREE"]) / QTSTK;


                    newDoc.Lines.WarehouseCode = (string)line["WHSCODE"];
                    newDoc.Lines.VatGroup = (string)line["VATGROUP"];
                    //  newDoc.Lines.TaxCode = (string)line["VATGROUP"];  //Pelos testes que fiz e pela documentação o TaxCode liga á tabela OSTC e não é o que interessa
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
                    newDoc.Lines.UserFields.Fields.Item("U_apyIDPROMO").Value = (int)line["IDPROMO"];

                    newDoc.Lines.UserFields.Fields.Item("U_apyUIEC").Value = (string)line["UIEC"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyUECOVALOR").Value = (string)line["UECOVALOR"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyUECOREE").Value = (string)line["UECOREE"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyIEC").Value = (double)(decimal)line["IEC"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyECOVALOR").Value = (double)(decimal)line["ECOVALOR"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyECOREE").Value = (double)(decimal)line["ECOREE"];

                    //newDoc.Lines.LineTotal = (double)(decimal)line["LINETOTAL"];
                    newDoc.Lines.LineTotal = (double)(decimal)line["LINETOTAL2"]; //includes IEC, ECOVALOR, ECOREE

                    if ((int)line["BASE_DOCENTRY"] != 0)
                    {
                        string BaseCardCode = (string)line["BASE_CARDCODE"];
                        if (newDoc.CardCode != BaseCardCode) throw new Exception("Encontrada referência a documento de outra entidade: " + BaseCardCode);

                        double QTYSTK_AVAILABLE_SAP = (double)(decimal)line["QTYSTK_AVAILABLE_SAP"];
                        if (QTSTK > QTYSTK_AVAILABLE_SAP) throw new Exception("Não pode relacionar mais que " + QTYSTK_AVAILABLE_SAP + " UN do artigo " + (string)line["ITEMNAME"] + " com o documento base.");

                        if ((string)line["BASE_DOCSTATUS"] != "O" || objCode == "14")
                        {
                            // fazer uma referência indirecta
                            newDoc.Lines.UserFields.Fields.Item("U_apyBSTYPE").Value = (int)line["BASE_OBJTYPE"];
                            newDoc.Lines.UserFields.Fields.Item("U_apyBSENTRY").Value = (int)line["BASE_DOCENTRY"];
                            newDoc.Lines.UserFields.Fields.Item("U_apyBSLINE").Value = (int)line["BASE_LINENUM"];
                            newDoc.Lines.UserFields.Fields.Item("U_apyBSNUM").Value = (int)line["BASE_DOCNUM"];
                        }
                        else
                        {
                            // Fazer uma referência directa do SAP B1
                            newDoc.Lines.BaseType = (int)line["BASE_OBJTYPE"];
                            newDoc.Lines.BaseEntry = (int)line["BASE_DOCENTRY"];
                            newDoc.Lines.BaseLine = (int)line["BASE_LINENUM"];
                        }
                    }
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
                    newDoc.Lines.VatGroup = (string)line["VATGROUP"];
                    //  newDoc.Lines.TaxCode = (string)line["VATGROUP"];  //Pelos testes que fiz e pela documentação o TaxCode liga á tabela OSTC e não é o que interessa
                    newDoc.Lines.UserFields.Fields.Item("U_apyINCONF").Value = (short)line["HASINCONF"] == 1 ? "Y" : "N";
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
                        newDoc.Lines.VatGroup = (string)line["VATGROUP"];
                        //  newDoc.Lines.TaxCode = (string)line["VATGROUP"];  //Pelos testes que fiz e pela documentação o TaxCode liga á tabela OSTC e não é o que interessa
                        newDoc.Lines.UserFields.Fields.Item("U_apyINCONF").Value = (short)line["HASINCONF"] == 1 ? "Y" : "N";
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

                if (newDoc != null) Marshal.ReleaseComObject(newDoc);
                if (invEntry != null) Marshal.ReleaseComObject(newDoc);
                if (invReval != null) Marshal.ReleaseComObject(newDoc);
                if (invExit != null) Marshal.ReleaseComObject(newDoc);
                newDoc = null;
                invEntry = null;
                invReval = null;
                invExit = null;

                GC.Collect();
            }
        }
    }

    internal AddDocResult SAPDOC_FROM_SAPPY_DRAFT_POS(DocActions action, string objCode, int draftId, double expectedTotal)
    {
        int priceDecimals = 6;
        {
            var s = this.company.GetCompanyService();
            var ai = s.GetAdminInfo();
            priceDecimals = ai.PriceAccuracy;
        }


        string CFINAL_CARDCODE = "";
        int CFINAL_SERIE13 = 0;
        int DOC_SERIE = 0;
        string COND_ITEMS_WITHOUT_PRICE = "";

        var sql = "SELECT * ";
        sql += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_SETTINGS ";
        sql += "\n WHERE ID IN ('POS.CFINAL.CARDCODE'";
        sql += "\n             ,'POS.CFINAL.SERIE13'";
        sql += "\n             ,'POS.GERAL.SERIE" + objCode + "'";
        sql += "\n             ,'POS.GERAL.COND_ITEMS_WITHOUT_PRICE')";
        using (HelperOdbc dataLayer = new HelperOdbc())
        using (DataTable dt = dataLayer.Execute(sql))
        {
            foreach (DataRow row in dt.Rows)
            {
                if ((string)row["ID"] == "POS.CFINAL.CARDCODE") CFINAL_CARDCODE = (string)row["RAW_VALUE"];
                if ((string)row["ID"] == "POS.CFINAL.SERIE13") CFINAL_SERIE13 = Convert.ToInt32(row["RAW_VALUE"]);
                if ((string)row["ID"] == "POS.GERAL.SERIE" + objCode) DOC_SERIE = Convert.ToInt32(row["RAW_VALUE"]);
                if ((string)row["ID"] == "POS.GERAL.COND_ITEMS_WITHOUT_PRICE") COND_ITEMS_WITHOUT_PRICE = (string)row["RAW_VALUE"];
            }
        }
        if (COND_ITEMS_WITHOUT_PRICE == "") COND_ITEMS_WITHOUT_PRICE = "1=0";// se não houver definição, esta expressão causa que todos tem que ter preço

        // Obter cabeçalho do documento a adicionar
        var sqlHeader = "SELECT T0.* ";
        sqlHeader += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC T0"; ;
        sqlHeader += "\n WHERE T0.ID =" + draftId;

        // Obter linhas do documento a adicionar
        var sqlDetail = "SELECT T1.*";
        sqlDetail += "\n , OITM.\"InvntryUom\"";
        sqlDetail += "\n , OITW.\"OnHand\"";
        sqlDetail += "\n , OITW.\"AvgPrice\"";
        sqlDetail += "\n      , BASEDOC.\"CardCode\" AS BASE_CARDCODE";
        sqlDetail += "\n      , BASEDOC.QTYSTK_AVAILABLE_SAP";
        sqlDetail += "\n      , BASEDOC.\"DocStatus\" AS BASE_DOCSTATUS";
        sqlDetail += "\n , CASE WHEN " + COND_ITEMS_WITHOUT_PRICE + " THEN 'N' ELSE 'Y' END AS MUST_HAVE_PRICE";
        sqlDetail += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES T1";
        sqlDetail += "\n INNER JOIN \"" + this.company.CompanyDB + "\".OITM OITM on T1.ITEMCODE = OITM.\"ItemCode\"";
        sqlDetail += "\n LEFT JOIN \"" + this.company.CompanyDB + "\".OITW OITW on T1.ITEMCODE = OITW.\"ItemCode\" AND T1.WHSCODE = OITW.\"WhsCode\"";
        sqlDetail += "\n LEFT JOIN \"" + this.company.CompanyDB + "\".SAPPY_LINE_LINK_" + objCode + " as BASEDOC ";
        sqlDetail += "\n                ON T1.BASE_OBJTYPE  = BASEDOC.\"ObjType\"";
        sqlDetail += "\n               AND T1.BASE_DOCENTRY = BASEDOC.\"DocEntry\"";
        sqlDetail += "\n               AND T1.BASE_LINENUM  = BASEDOC.\"LineNum\"";
        sqlDetail += "\n WHERE T1.ID =" + draftId;
        sqlDetail += "\n ORDER BY T1.LINENUM";

        //// Validar os totais por Artigo relacionados com o documento Base
        //var sqlTotalByItem = "SELECT T1.ITEMCODE";
        //sqlTotalByItem += "\n      , max(T1.ITEMNAME) AS ITEMNAME";
        //sqlTotalByItem += "\n      , T1.BASE_OBJTYPE";
        //sqlTotalByItem += "\n      , T1.BASE_DOCENTRY";
        //sqlTotalByItem += "\n      , T1.BASE_LINENUM";
        //sqlTotalByItem += "\n      , BASEDOC.\"CardCode\"";
        //sqlTotalByItem += "\n      , BASEDOC.QTYSTK_AVAILABLE_SAP";
        //sqlTotalByItem += "\n      , SUM(T1.QTSTK) AS QTSTK";
        //sqlTotalByItem += "\n FROM \"" + this.company.CompanyDB + "\".SAPPY_DOC_LINES T1";
        //sqlTotalByItem += "\n INNER JOIN \"" + this.company.CompanyDB + "\".SAPPY_LINE_LINK_" + objCode + " as BASEDOC ";
        //sqlTotalByItem += "\n                ON T1.BASE_OBJTYPE  = BASEDOC.\"ObjType\"";
        //sqlTotalByItem += "\n               AND T1.BASE_DOCENTRY = BASEDOC.\"DocEntry\"";
        //sqlTotalByItem += "\n               AND T1.BASE_LINENUM  = BASEDOC.\"LineNum\"";
        //sqlTotalByItem += "\n WHERE T1.ID =" + draftId;
        //sqlTotalByItem += "\n GROUP BY T1.ITEMCODE";
        //sqlTotalByItem += "\n      , T1.BASE_OBJTYPE";
        //sqlTotalByItem += "\n      , T1.BASE_DOCENTRY";
        //sqlTotalByItem += "\n      , T1.BASE_LINENUM";
        //sqlTotalByItem += "\n      , BASEDOC.\"CardCode\"";
        //sqlTotalByItem += "\n      , BASEDOC.\"QTYSTK_AVAILABLE_SAP\"";
        //sqlTotalByItem += "\n ORDER BY min(T1.LINENUM)";

        using (HelperOdbc dataLayer = new HelperOdbc())
        using (DataTable headerDt = dataLayer.Execute(sqlHeader))
        using (DataTable detailsDt = dataLayer.Execute(sqlDetail))
        //using (DataTable itemTotalsDt = dataLayer.Execute(sqlTotalByItem))
        {
            DataRow header = headerDt.Rows[0];

            int module = Convert.ToInt32(header["MODULE"]);
            if (module != 2) throw new Exception("Esta função só pode ser chamada para MODULE = 2 (POS)");

            int objType = Convert.ToInt32(header["OBJTYPE"]);
            if (objType.ToString() != objCode) throw new Exception("objCode " + objCode + " is diferent of " + objType.ToString());

            int serie = 0;

            if (objType == 13 && (string)header["CARDCODE"] == CFINAL_CARDCODE)
            {
                //Factura a CLiente Final = Factura simplificada
                serie = CFINAL_SERIE13;
            }
            else
            {
                serie = DOC_SERIE;
            }
            if (serie == 0) throw new Exception("A Série não está definida nas opções para este documento (POS).");


            int DISTRIBUICAO = (short)header["DISTRIBUICAO"];


            SAPbobsCOM.Documents newDoc = (SAPbobsCOM.Documents)this.company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)objType);
            newDoc.Series = serie;
            // DateTime TAXDATE = (DateTime)header["TAXDATE"];
            // DateTime DOCDUEDATE = (DateTime)header["DOCDUEDATE"];
            // if (TAXDATE.Year > 1900) newDoc.TaxDate = TAXDATE;
            // if (DOCDUEDATE.Year > 1900) newDoc.DocDueDate = DOCDUEDATE;

            if (objType == 17)
            {
                // A Data de entrega é obrigatória
                newDoc.DocDueDate = DateTime.Now;
            }

            newDoc.CardCode = (string)header["CARDCODE"];
            if ((string)header["SHIPADDR"] != "") newDoc.ShipToCode = (string)header["SHIPADDR"];
            if ((string)header["BILLADDR"] != "") newDoc.PayToCode = (string)header["BILLADDR"];
            if ((string)header["NUMATCARD"] != "") newDoc.NumAtCard = (string)header["NUMATCARD"];
            if ((string)header["COMMENTS"] != "") newDoc.Comments = (string)header["COMMENTS"];

            newDoc.UserFields.Fields.Item("U_apyUSER").Value = (string)header["CREATED_BY_NAME"];
            if ((string)header["MATRICULA"] != "") newDoc.UserFields.Fields.Item("U_apyMATRICULA").Value = (string)header["MATRICULA"];



            //// preform checks by item total
            //foreach (DataRow itemTotal in itemTotalsDt.Rows)
            //{
            //    var ITEMCODE = (string)itemTotal["ITEMCODE"];
            //    var ITEMNAME = (string)itemTotal["ITEMNAME"];
            //    var QTSTK = (double)(decimal)itemTotal["QTSTK"];

            //    string CardCode = (string)itemTotal["CardCode"];
            //    if (CardCode != "" && newDoc.CardCode != CardCode) throw new Exception("Encontrada referência a documento de outra entidade: " + CardCode);

            //    double QTYSTK_AVAILABLE_SAP = (double)(decimal)itemTotal["QTYSTK_AVAILABLE_SAP"];
            //    if (QTSTK > QTYSTK_AVAILABLE_SAP) throw new Exception("Não pode relacionar mais que " + QTYSTK_AVAILABLE_SAP + " UN do artigo " + ITEMNAME + " com o documento base.");

            //}

            foreach (DataRow line in detailsDt.Rows)
            {

                var ITEMCODE = (string)line["ITEMCODE"];
                var ITEMNAME = (string)line["ITEMNAME"];
                var QTCX = (double)(decimal)line["QTCX"];   // Num caixas/pack
                var QTPK = (double)(decimal)line["QTPK"];
                var QTSTK = (double)(decimal)line["QTSTK"];
                var QTBONUS = (double)(decimal)line["QTBONUS"];
                var PRICE = (double)(decimal)line["PRICE"];
                bool MUST_HAVE_PRICE = (string)line["MUST_HAVE_PRICE"] == "Y";


                if (MUST_HAVE_PRICE && PRICE <= 0) throw new Exception("O artigo " + ITEMNAME + " tem que ter preço.");


                if (QTSTK != 0)
                {
                    if (newDoc.Lines.ItemCode != "") newDoc.Lines.Add();
                    newDoc.Lines.ItemCode = ITEMCODE;
                    newDoc.Lines.ItemDescription = ITEMNAME;
                    newDoc.Lines.MeasureUnit = (string)line["InvntryUom"];
                    newDoc.Lines.Factor1 = QTCX;   // Num caixas/pack
                    newDoc.Lines.Factor2 = QTPK;   // Qdd por Caixa/pack 
                    //newDoc.Lines.InventoryQuantity = (double)(decimal)line["QTSTK"]; //Definir sobrepoe os fatores 1 e 2
                    newDoc.Lines.UnitPrice = PRICE;
                    newDoc.Lines.WarehouseCode = (string)line["WHSCODE"];
                    newDoc.Lines.VatGroup = (string)line["VATGROUP"];
                    //  newDoc.Lines.TaxCode = (string)line["VATGROUP"];  //Pelos testes que fiz e pela documentação o TaxCode liga á tabela OSTC e não é o que interessa


                    // newDoc.Lines.LineTotal = (double)(decimal)line["LINETOTAL"];
                    newDoc.Lines.DiscountPercent = (double)(decimal)line["DISCOUNT"];

                    newDoc.Lines.UserFields.Fields.Item("U_apyUDISC").Value = (string)line["USER_DISC"];
                    newDoc.Lines.UserFields.Fields.Item("U_apyIDPROMO").Value = (int)line["IDPROMO"];
                    if ((string)line["PRICE_CHANGEDBY"] != "") newDoc.Lines.UserFields.Fields.Item("U_apyPRICECHBY").Value = (string)line["PRICE_CHANGEDBY"];
                    if ((string)line["DISC_CHANGEDBY"] != "") newDoc.Lines.UserFields.Fields.Item("U_apyDISCCHBY").Value = (string)line["DISC_CHANGEDBY"];

                    if ((int)line["BASE_DOCENTRY"] != 0)
                    {
                        string BaseCardCode = (string)line["BASE_CARDCODE"];
                        if (newDoc.CardCode != BaseCardCode) throw new Exception("Encontrada referência a documento de outra entidade: " + BaseCardCode);

                        double QTYSTK_AVAILABLE_SAP = (double)(decimal)line["QTYSTK_AVAILABLE_SAP"];
                        if (QTSTK > QTYSTK_AVAILABLE_SAP) throw new Exception("Não pode relacionar mais que " + QTYSTK_AVAILABLE_SAP + " UN do artigo " + ITEMNAME + " com o documento base.");

                        if ((string)line["BASE_DOCSTATUS"] != "O" || objCode == "14")
                        {
                            // fazer uma referência indirecta
                            newDoc.Lines.UserFields.Fields.Item("U_apyBSTYPE").Value = (int)line["BASE_OBJTYPE"];
                            newDoc.Lines.UserFields.Fields.Item("U_apyBSENTRY").Value = (int)line["BASE_DOCENTRY"];
                            newDoc.Lines.UserFields.Fields.Item("U_apyBSLINE").Value = (int)line["BASE_LINENUM"];
                            newDoc.Lines.UserFields.Fields.Item("U_apyBSNUM").Value = (int)line["BASE_DOCNUM"];
                        }
                        else
                        {
                            // Fazer uma referência directa do SAP B1
                            newDoc.Lines.BaseType = (int)line["BASE_OBJTYPE"];
                            newDoc.Lines.BaseEntry = (int)line["BASE_DOCENTRY"];
                            newDoc.Lines.BaseLine = (int)line["BASE_LINENUM"];
                        }
                    }
                }
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

                if (action == DocActions.ADD)
                {
                    if (DISTRIBUICAO == 1)
                    {
                        SAPbobsCOM.JournalEntries JD = this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        if (!JD.GetByKey(newDoc.TransNum))
                        {
                            throw new Exception("Não foi obter o documento LD criado em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                        }

                        for (int i = 0; i < JD.Lines.Count; i++)
                        {
                            JD.Lines.SetCurrentLine(i);
                            if (JD.Lines.ShortName == newDoc.CardCode)
                            {
                                JD.Lines.UserFields.Fields.Item("U_apyCLASS").Value = "D";
                            }
                        }

                        if (JD.Update() != 0)
                        {
                            var ex = new Exception("Não foi possível marcar LD para Distribuição em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                            throw ex;
                        }
                    }


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
                if (newDoc != null) Marshal.ReleaseComObject(newDoc);
                newDoc = null;
                GC.Collect();
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
                if (field == "HASINCONF") sapDoc.UserFields.Fields.Item("U_apyINCONF").Value = (value != "" && "true,1".Contains(value.ToLower()) ? "Y" : "N");


                if (field == "MATRICULA") sapDoc.UserFields.Fields.Item("U_apyMATRICULA").Value = value;

                if ("15,16,21".IndexOf(objCode) > -1)
                {
                    if (field == "ATDOCTYPE") sapDoc.ATDocumentType = value;
                    if (field == "MATRICULA") sapDoc.VehiclePlate = value;
                    if (field == "ATAUTHCODE") sapDoc.AuthorizationCode = value;
                    if (field == "ATTRYAGAIN" && value == "1")
                    {
                        sapDoc.StartDeliveryDate = DateTime.Now.AddMinutes(5);
                        sapDoc.StartDeliveryTime = DateTime.Now.AddMinutes(5);

                        sapDoc.ElecCommStatus = SAPbobsCOM.ElecCommStatusEnum.ecsPendingApproval;
                    }
                }

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

    internal AddDocResult SAPDOC_CANCELDOC(string objCode, int docEntry)
    {

        int objType = Convert.ToInt32(objCode);
        if (objType.ToString() != objCode) throw new Exception("objCode " + objCode + " is diferent of " + objType.ToString());

        int priceDecimals = 6;
        {
            var s = this.company.GetCompanyService();
            var ai = s.GetAdminInfo();
            priceDecimals = ai.PriceAccuracy;
        }



        SAPbobsCOM.Documents origDoc = null;
        SAPbobsCOM.Documents sapDoc = null;


        try
        {

            origDoc = (SAPbobsCOM.Documents)this.company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)objType);
            if (origDoc.GetByKey(docEntry) == false)
                throw new Exception("Não foi possível obter o documento do SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

            if ("17,23, 22".Contains(objCode))
            {
                // Cotações e encomendas de cleintes e de fornecedores não usam documento de cancelamento                
                this.company.StartTransaction();

                if (origDoc.Cancel() != 0)
                {
                    var ex = new Exception("Não foi possível cancelar em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    throw ex;
                }

                AddDocResult result = new AddDocResult();

                this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                result.DocEntry = origDoc.DocEntry;
                result.DocNum = origDoc.DocNum;

                result.DocTotal = origDoc.DocTotal;
                result.DiscountPercent = origDoc.DiscountPercent;
                result.TotalDiscount = origDoc.TotalDiscount;
                result.VatSum = origDoc.VatSum;
                result.RoundingDiffAmount = origDoc.RoundingDiffAmount;
                return result;
            }
            else
            {
                // Gerar documento de cancelamento
                sapDoc = origDoc.CreateCancellationDocument();

                if (sapDoc == null)
                    throw new Exception("Este documento não pode ser cancelado: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                this.company.StartTransaction();

                if (sapDoc.Add() != 0)
                {
                    var ex = new Exception("Não foi possível cancelar em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    throw ex;
                }

                AddDocResult result = new AddDocResult();

                this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                result.DocEntry = sapDoc.DocEntry;
                result.DocNum = sapDoc.DocNum;

                result.DocTotal = sapDoc.DocTotal;
                result.DiscountPercent = sapDoc.DiscountPercent;
                result.TotalDiscount = sapDoc.TotalDiscount;
                result.VatSum = sapDoc.VatSum;
                result.RoundingDiffAmount = sapDoc.RoundingDiffAmount;
                return result;
            }
        }
        finally
        {
            if (this.company.InTransaction) this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            if (origDoc != null) Marshal.ReleaseComObject(origDoc);
            if (sapDoc != null) Marshal.ReleaseComObject(sapDoc);
            origDoc = null;
            sapDoc = null;
            GC.Collect();
        }

    }

    internal AddDocResult SAPDOC_CLOSEDOC(string objCode, int docEntry)
    {
        int objType = Convert.ToInt32(objCode);
        if (objType.ToString() != objCode) throw new Exception("objCode " + objCode + " is diferent of " + objType.ToString());

        int priceDecimals = 6;
        {
            var s = this.company.GetCompanyService();
            var ai = s.GetAdminInfo();
            priceDecimals = ai.PriceAccuracy;
        }

        SAPbobsCOM.Documents sapDoc = (SAPbobsCOM.Documents)this.company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)objType);

        if (sapDoc.GetByKey(docEntry) == false)
            throw new Exception("Não foi possível obter o documento do SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

        try
        {
            this.company.StartTransaction();

            if (sapDoc.Close() != 0)
            {
                var ex = new Exception("Não foi possível fechar documento em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                throw ex;
            }

            AddDocResult result = new AddDocResult();

            this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

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
            if (sapDoc != null) Marshal.ReleaseComObject(sapDoc);
            sapDoc = null;
            GC.Collect();
        }

    }

    internal AddDocResult ADD_ADIANTAMENTO_PARA_DESPESAS(PostAdiantamentoInput data)
    {
        SAPbobsCOM.Payments sapDoc = (SAPbobsCOM.Payments)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);

        sapDoc.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;
        sapDoc.CardCode = data.CardCode;
        sapDoc.CashAccount = data.CashAccount;
        sapDoc.ContactPersonCode = data.ContactPersonCode;
        sapDoc.CashSum = data.CashSum;
        sapDoc.Remarks = data.Remarks;
        sapDoc.CounterReference = data.CounterReference;

        try
        {
            this.company.StartTransaction();

            if (sapDoc.Add() != 0)
            {
                var ex = new Exception("Não foi possível gravar em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                //log the xml to allow easier debug
                var xml = sapDoc.GetAsXML();
                Logger.Log.Debug(xml, ex);

                throw ex;
            }
            else
            {
                int docEntry = 0;
                int.TryParse(this.company.GetNewObjectKey(), out docEntry);

                if (sapDoc.GetByKey(docEntry) == false)
                {
                    throw new Exception("Não foi obter o documento atualizado em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                }

                // Por algum motivo, a DI API não está a gravar o contacto enviado ao criar o documento. Como permite alterar depois, forçamos um update ao documento
                if (sapDoc.ContactPersonCode != data.ContactPersonCode)
                {
                    sapDoc.ContactPersonCode = data.ContactPersonCode;
                    if (sapDoc.Update() != 0)
                    {
                        var ex = new Exception("Não foi possível atualizar em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                        throw ex;
                    }
                }
            }

            AddDocResult result = new AddDocResult();

            this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

            result.DocEntry = sapDoc.DocEntry;
            result.DocNum = sapDoc.DocNum;

            return result;
        }
        finally
        {
            if (this.company.InTransaction) this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        }

    }

    internal AddPagResult ADD_APGAMENTO_FORNECEDOR(PostPaymentInput data)
    {
        bool addDebito = false;
        bool addEC = false;
        double valorECC = 0;
        double valorECF = 0;

        // Encontro de contas - CLIENTE
        SAPbobsCOM.Payments ECC = (SAPbobsCOM.Payments)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
        ECC.DocType = SAPbobsCOM.BoRcptTypes.rCustomer;
        ECC.CardCode = data.CardCodeEC;
        ECC.Remarks = "EC PAG";

        foreach (var doc in data.PaymentInvoices)
        {
            if (doc.ValorECC != 0)
            {
                addEC = true;
                if (ECC.Invoices.SumApplied != 0) ECC.Invoices.Add();
                if (doc.DocEntry != null) ECC.Invoices.DocEntry = (int)doc.DocEntry;
                if (doc.DocLine != null) ECC.Invoices.DocLine = (int)doc.DocLine;
                ECC.Invoices.InvoiceType = (SAPbobsCOM.BoRcptInvTypes)(Convert.ToInt32(doc.InvoiceType));
                ECC.Invoices.SumApplied = (double)doc.ValorECC;
                //ECC.Invoices.TotalDiscount = (double)doc.ValorDescontoECC; // NÃO HÁ DECONTOS EM ENC CONTAS NO CLIENTE
                if (doc.U_apyCONTRATO != null) ECC.Invoices.UserFields.Fields.Item("U_apyCONTRATO").Value = (int)doc.U_apyCONTRATO;
                ECC.Invoices.UserFields.Fields.Item("U_apyUDISC").Value = doc.U_apyUDISC;
                ECC.Invoices.UserFields.Fields.Item("U_apyUDEBITO").Value = doc.U_apyUDEBITO;

                valorECC += ECC.Invoices.SumApplied;
            }
        }
        ECC.TransferAccount = data.TransferAccountEC;
        ECC.TransferReference = "EC";
        ECC.TransferSum = valorECC;

        // Encontro de contas - FORNECEDOR
        SAPbobsCOM.Payments ECF = (SAPbobsCOM.Payments)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
        ECF.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;
        ECF.CardCode = data.CardCode;
        ECF.Remarks = "EC PAG";

        foreach (var doc in data.PaymentInvoices)
        {
            if (doc.ValorECF != 0)
            {
                if (ECF.Invoices.SumApplied != 0) ECF.Invoices.Add();
                if (doc.DocEntry != null) ECF.Invoices.DocEntry = (int)doc.DocEntry;
                if (doc.DocLine != null) ECF.Invoices.DocLine = (int)doc.DocLine;
                ECF.Invoices.InvoiceType = (SAPbobsCOM.BoRcptInvTypes)(Convert.ToInt32(doc.InvoiceType));
                ECF.Invoices.SumApplied = (double)doc.ValorECF;
                ECF.Invoices.TotalDiscount = (double)doc.ValorDescontoECF;
                if (doc.U_apyCONTRATO != null) ECF.Invoices.UserFields.Fields.Item("U_apyCONTRATO").Value = (int)doc.U_apyCONTRATO;
                ECF.Invoices.UserFields.Fields.Item("U_apyUDISC").Value = doc.U_apyUDISC;
                ECF.Invoices.UserFields.Fields.Item("U_apyUDEBITO").Value = doc.U_apyUDEBITO;

                valorECF += ECF.Invoices.SumApplied;
            }
        }
        ECF.TransferAccount = data.TransferAccountEC;
        ECF.TransferReference = "EC";
        ECF.TransferSum = valorECF;

        if ((decimal)valorECC != (decimal)valorECF)
        {
            var ex = new Exception("Valor de encontro de contas não bate certo entre o cliente e o fornecedor.");
            throw ex;
        }

        // Pagamento ao FORNECEDOR
        SAPbobsCOM.Payments sapDoc = (SAPbobsCOM.Payments)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);

        sapDoc.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;
        sapDoc.CardCode = data.CardCode;
        sapDoc.Remarks = data.Remarks;

        sapDoc.CashAccount = data.CashAccount;
        sapDoc.CashSum = (double)data.CashSum;

        sapDoc.TransferAccount = data.TransferAccount;
        sapDoc.TransferSum = (double)data.TransferSum;
        sapDoc.TransferReference = data.TransferReference;


        sapDoc.CheckAccount = data.CheckAccount;
        foreach (var ch in data.PaymentChecks)
        {
            if (sapDoc.Checks.CheckSum != 0) sapDoc.Checks.Add();
            sapDoc.Checks.AccounttNum = ch.AccounttNum;
            sapDoc.Checks.BankCode = ch.BankCode;
            sapDoc.Checks.CheckAccount = ch.CheckAccount;
            sapDoc.Checks.CheckNumber = ch.CheckNumber;
            sapDoc.Checks.CheckSum = (float)ch.CheckSum;
            sapDoc.Checks.CountryCode = ch.CountryCode;
            sapDoc.Checks.DueDate = DateTime.ParseExact(ch.DueDate, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
            sapDoc.Checks.ManualCheck = SAPbobsCOM.BoYesNoEnum.tYES;
        }

        sapDoc.UserFields.Fields.Item("U_apyNotas5").Value = data.U_apyNotas5;
        sapDoc.UserFields.Fields.Item("U_apyNotas10").Value = data.U_apyNotas10;
        sapDoc.UserFields.Fields.Item("U_apyNotas20").Value = data.U_apyNotas20;
        sapDoc.UserFields.Fields.Item("U_apyNotas50").Value = data.U_apyNotas50;
        sapDoc.UserFields.Fields.Item("U_apyNotas100").Value = data.U_apyNotas100;
        sapDoc.UserFields.Fields.Item("U_apyNotas200").Value = data.U_apyNotas200;
        sapDoc.UserFields.Fields.Item("U_apyNotas500").Value = data.U_apyNotas500;
        sapDoc.UserFields.Fields.Item("U_apyNotas").Value = (float)data.U_apyNotas;
        sapDoc.UserFields.Fields.Item("U_apyMoedas").Value = (float)data.U_apyMoedas;
        sapDoc.UserFields.Fields.Item("U_apyVales").Value = (float)data.U_apyVales;
        sapDoc.UserFields.Fields.Item("U_apyTickets").Value = (float)data.U_apyTickets;
        sapDoc.UserFields.Fields.Item("U_apyTroco").Value = (float)data.U_apyTroco;

        foreach (var doc in data.PaymentInvoices)
        {
            if (doc.ValorDeduzir != 0)
            {
                if (sapDoc.Invoices.SumApplied != 0) sapDoc.Invoices.Add();
                if (doc.DocEntry != null) sapDoc.Invoices.DocEntry = (int)doc.DocEntry;
                if (doc.DocLine != null) sapDoc.Invoices.DocLine = (int)doc.DocLine;
                sapDoc.Invoices.InvoiceType = (SAPbobsCOM.BoRcptInvTypes)(Convert.ToInt32(doc.InvoiceType));
                sapDoc.Invoices.SumApplied = -1*(double)doc.ValorDeduzir;
                //sapDoc.Invoices.TotalDiscount = (double)doc.ValorDescontoDeduzir; 'Não há desconto nos documentos a deduzir
                if (doc.U_apyCONTRATO != null) sapDoc.Invoices.UserFields.Fields.Item("U_apyCONTRATO").Value = (int)doc.U_apyCONTRATO;
                sapDoc.Invoices.UserFields.Fields.Item("U_apyUDISC").Value = doc.U_apyUDISC;
                sapDoc.Invoices.UserFields.Fields.Item("U_apyUDEBITO").Value = doc.U_apyUDEBITO;
            }
        }

        foreach (var doc in data.PaymentInvoices)
        {
            if (doc.ValorPagar != 0)
            {
                if (sapDoc.Invoices.SumApplied != 0) sapDoc.Invoices.Add();
                if (doc.DocEntry != null) sapDoc.Invoices.DocEntry = (int)doc.DocEntry;
                if (doc.DocLine != null) sapDoc.Invoices.DocLine = (int)doc.DocLine;
                sapDoc.Invoices.InvoiceType = (SAPbobsCOM.BoRcptInvTypes)(Convert.ToInt32(doc.InvoiceType));
                sapDoc.Invoices.SumApplied = (double)doc.ValorPagar;
                sapDoc.Invoices.TotalDiscount = (double)doc.ValorDescontoPagar;
                if (doc.U_apyCONTRATO != null) sapDoc.Invoices.UserFields.Fields.Item("U_apyCONTRATO").Value = (int)doc.U_apyCONTRATO;
                sapDoc.Invoices.UserFields.Fields.Item("U_apyUDISC").Value = doc.U_apyUDISC;
                sapDoc.Invoices.UserFields.Fields.Item("U_apyUDEBITO").Value = doc.U_apyUDEBITO;
            }
        }

        // fatura com débito a CLENTE (assoicado ao fornecedor)
        SAPbobsCOM.Documents ftDebito = (SAPbobsCOM.Documents)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
        ftDebito.CardCode = data.CodigoClienteDebito;
        ftDebito.Series = (int)data.SerieDebitos;
        foreach (var doc in data.PaymentInvoices)
        {
            if (doc.ValorDebito != 0)
            {
                addDebito = true;
                if (ftDebito.Lines.ItemCode != "") ftDebito.Lines.Add();
                ftDebito.Lines.ItemCode = data.ArtigoDebitos;
                ftDebito.Lines.Quantity = 1;
                ftDebito.Lines.UnitPrice = (float)doc.ValorDebito;
                if (doc.InvoiceType == "18")
                    ftDebito.Lines.ItemDescription = doc.CONTRATO_DESC + " (Ref v/FT " + doc.Ref2 + ")";
                else
                    ftDebito.Lines.ItemDescription = doc.CONTRATO_DESC + " (Ref v/doc " + doc.Ref2 + ")";
            }
        }

        try
        {
            AddPagResult result = new AddPagResult();
            this.company.StartTransaction();

            if (addEC) { 

                if (ECC.Add() != 0)
                {
                    var ex = new Exception("Não foi possível gravar em SAP1: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                    //log the xml to allow easier debug
                    var xml = sapDoc.GetAsXML();
                    Logger.Log.Debug(xml, ex);

                    throw ex;
                }
                else
                {
                    int docEntry = 0;
                    int.TryParse(this.company.GetNewObjectKey(), out docEntry);

                    if (ECC.GetByKey(docEntry) == false)
                    {
                        throw new Exception("Não foi obter o documento criado em SAP1: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    }
                    sapDoc.UserFields.Fields.Item("U_apyECC").Value = docEntry;
                }

                if (ECF.Add() != 0)
                {
                    var ex = new Exception("Não foi possível gravar em SAP2: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                    //log the xml to allow easier debug
                    var xml = sapDoc.GetAsXML();
                    Logger.Log.Debug(xml, ex);

                    throw ex;
                }
                else
                {
                    int docEntry = 0;
                    int.TryParse(this.company.GetNewObjectKey(), out docEntry);

                    if (ECF.GetByKey(docEntry) == false)
                    {
                        throw new Exception("Não foi obter o documento criado em SAP2: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    }
                    sapDoc.UserFields.Fields.Item("U_apyECF").Value = docEntry;
                }

            }

            if (addDebito)
            {
                if (ftDebito.Add() != 0)
                {
                    var ex = new Exception("Não foi possível gravar em SAP3: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                    //log the xml to allow easier debug
                    var xml = sapDoc.GetAsXML();
                    Logger.Log.Debug(xml, ex);

                    throw ex;
                }
                else
                {
                    int docEntry = 0;
                    int.TryParse(this.company.GetNewObjectKey(), out docEntry);

                    if (ftDebito.GetByKey(docEntry) == false)
                    {
                        throw new Exception("Não foi obter o documento criado em SAP3: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    }
                    sapDoc.UserFields.Fields.Item("U_apyFTDEBITO").Value = docEntry;
                }
                result.FtDebitoDocEntry = ftDebito.DocEntry;
                result.FtDebitoDocNum = ftDebito.DocNum;

            }

            if (sapDoc.Add() != 0)
            {
                var ex = new Exception("Não foi possível gravar em SAP4: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                //log the xml to allow easier debug
                var xml = sapDoc.GetAsXML();
                Logger.Log.Debug(xml, ex);

                throw ex;
            }
            else
            {
                int docEntry = 0;
                int.TryParse(this.company.GetNewObjectKey(), out docEntry);

                if (sapDoc.GetByKey(docEntry) == false)
                {
                    throw new Exception("Não foi obter o documento atualizado em SAP4: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                }
                result.PagDocEntry = sapDoc.DocEntry;
            }
             

            this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
              
            return result;
        }
        finally
        {
            if (this.company.InTransaction) this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        }
    }

    internal AddDocResult CANCELAR_PAGAMENTO(PostCancelarPagamentoInput data)
    {

        SAPbobsCOM.Payments pagamento = (SAPbobsCOM.Payments)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
        SAPbobsCOM.Payments pagamentoEC = (SAPbobsCOM.Payments)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
        SAPbobsCOM.Payments recebimentoEC = (SAPbobsCOM.Payments)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
        SAPbobsCOM.Documents ftDebito = (SAPbobsCOM.Documents)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
        SAPbobsCOM.Documents ftDebitoCancel = (SAPbobsCOM.Documents)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);

        if (pagamento.GetByKey(data.DocEntry) == false)
            throw new Exception("Não foi obter o pagamento do SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());


        int? U_apyECC = pagamento.UserFields.Fields.Item("U_apyECC").Value;
        int? U_apyECF = pagamento.UserFields.Fields.Item("U_apyECF").Value;
        int? U_apyFTDEBITO = pagamento.UserFields.Fields.Item("U_apyFTDEBITO").Value;

        if (U_apyECC > 0)
        {
            if (recebimentoEC.GetByKey((int)U_apyECC) == false)
                throw new Exception("Não foi obter o recebimento EC do SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
        }
        if (U_apyECF > 0)
        {
            if (pagamentoEC.GetByKey((int)U_apyECF) == false)
                throw new Exception("Não foi obter o pagamento EC do SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
        }

        if (U_apyFTDEBITO > 0)
        {
            if (ftDebito.GetByKey((int)U_apyFTDEBITO) == false)
                throw new Exception("Não foi obter a fatura de débito do SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

            ftDebitoCancel = ftDebito.CreateCancellationDocument();
        }

        try
        {
            this.company.StartTransaction();

            pagamento.Remarks = "Cancelado-" + data.Reason;
            if (pagamento.Update() != 0)
            {
                var ex = new Exception("Colocar razão de cancelamento no pagamento: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                throw ex;
            }

            if (pagamento.GetByKey(data.DocEntry) == false)
                throw new Exception("Não foi obter o pagamento do SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
            
            if (pagamento.Cancel() != 0)
            {
                var ex = new Exception("Cancelar pagamento: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                throw ex;
            }




            if (U_apyECC > 0)
            {
                if (recebimentoEC.Cancel() != 0)
                {
                    var ex = new Exception("Cancelar recebimento EC: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    throw ex;
                }
            }

            if (U_apyECF > 0)
            {
                if (pagamentoEC.Cancel() != 0)
                {
                    var ex = new Exception("Cancelar pagamento EC: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    throw ex;
                }
            }

            if (U_apyFTDEBITO > 0)
            {
                if (ftDebitoCancel.Add() != 0)
                {
                    var ex = new Exception("Cancelar fatura débito EC: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    throw ex;
                }
            }


            AddDocResult result = new AddDocResult();

            this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);


            return result;
        }
        finally
        {
            if (this.company.InTransaction) this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        }

    }

    internal AddDocResult FECHAR_ADIANTAMENTO_PARA_DESPESAS(PostFecharAdiantamentoInput data)
    {

        // Documento para devolução do resto do adiantamento para conta de diferenças
        SAPbobsCOM.Payments devAdiant = (SAPbobsCOM.Payments)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
        devAdiant.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;
        devAdiant.CardCode = data.CardCode;
        devAdiant.ContactPersonCode = data.CntctCode;
        devAdiant.CounterReference = data.CounterRef;
        devAdiant.Remarks = data.Comments;
        devAdiant.TransferAccount = data.CAIXA_DIFERENCAS;
        devAdiant.TransferSum = data.VALOR_PENDENTE;
        devAdiant.Invoices.InvoiceType = (SAPbobsCOM.BoRcptInvTypes)data.TransType;
        if (data.TransType == 46)
        {
            devAdiant.Invoices.DocEntry = data.TransId;
            devAdiant.Invoices.DocLine = data.Line_ID;
        }
        else
        {
            devAdiant.Invoices.DocEntry = data.CreatedBy;
        }
        devAdiant.Invoices.SumApplied = data.VALOR_PENDENTE;

        try
        {
            this.company.StartTransaction();

            if (devAdiant.Add() != 0)
            {
                var ex = new Exception("Gravar recebimento: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                //log the xml to allow easier debug
                var xml = devAdiant.GetAsXML();
                Logger.Log.Debug(xml, ex);

                throw ex;
            }

            int docEntry = 0;
            int.TryParse(this.company.GetNewObjectKey(), out docEntry);

            if (devAdiant.GetByKey(docEntry) == false)
            {
                throw new Exception("Não foi obter o documento criado em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
            }

            // Por algum motivo, a DI API não está a gravar o contacto enviado ao criar o documento. Como permite alterar depois, forçamos um update ao documento
            if (devAdiant.ContactPersonCode != data.CntctCode)
            {
                devAdiant.ContactPersonCode = data.CntctCode;
                if (devAdiant.Update() != 0)
                {
                    var ex = new Exception("Atualizar recebimento: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                    throw ex;
                }
            }

            AddDocResult result = new AddDocResult();

            this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

            result.DocEntry = devAdiant.DocEntry;
            result.DocNum = devAdiant.DocNum;

            return result;
        }
        finally
        {
            if (this.company.InTransaction) this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        }

    }

    internal AddDocResult ADD_DESPESA(PostDespesaInput data)
    {
        double totalFactura = 0;

        // Fatura com a despesa real
        SAPbobsCOM.Documents ftDespesa = (SAPbobsCOM.Documents)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
        ftDespesa.Series = data.Series;
        ftDespesa.TaxDate = data.TaxDateParsed();
        ftDespesa.CardCode = data.CardCode;
        ftDespesa.NumAtCard = data.NumAtCard;
        ftDespesa.Comments = data.Comments;
        foreach (var line in data.Lines)
        {
            if (ftDespesa.Lines.ItemCode != "") ftDespesa.Lines.Add();

            ftDespesa.Lines.ItemCode = line.ItemCode;
            ftDespesa.Lines.Quantity = 1;
            ftDespesa.Lines.UnitPrice = 0;
            ftDespesa.Lines.PriceAfterVAT = line.ValorComIva;
            totalFactura += line.ValorComIva;
        }


        // Documento para devolução do adiantamento
        SAPbobsCOM.Payments devAdiant = (SAPbobsCOM.Payments)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
        bool comAdiantamento = data.MeioDePagamento.VALOR_PENDENTE > 0;
        if (comAdiantamento)
        {
            devAdiant.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;
            devAdiant.CardCode = data.MeioDePagamento.CardCode;
            devAdiant.ContactPersonCode = data.MeioDePagamento.CntctCode;
            devAdiant.CounterReference = data.MeioDePagamento.CounterRef;
            devAdiant.Remarks = data.MeioDePagamento.Comments;
            devAdiant.TransferAccount = data.CAIXA_PASSAGEM; //para pagamento da factura
            devAdiant.TransferSum = totalFactura;
            devAdiant.CashAccount = data.CAIXA_PRINCIPAL; //para retorno do troco
            devAdiant.CashSum = data.TrocoRecebido;
            devAdiant.Invoices.InvoiceType = (SAPbobsCOM.BoRcptInvTypes)data.MeioDePagamento.TransType;
            if (data.MeioDePagamento.TransType == 46)
            {
                devAdiant.Invoices.DocEntry = data.MeioDePagamento.TransId;
                devAdiant.Invoices.DocLine = data.MeioDePagamento.Line_ID;
            }
            else
            {
                devAdiant.Invoices.DocEntry = data.MeioDePagamento.CreatedBy;
            }
            devAdiant.Invoices.SumApplied = totalFactura + data.TrocoRecebido;
        }

        // Pagamento da factura real
        SAPbobsCOM.Payments pagFt = (SAPbobsCOM.Payments)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
        pagFt.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;
        pagFt.CardCode = data.CardCode;
        pagFt.CounterReference = data.MeioDePagamento.CounterRef;
        pagFt.Remarks = data.MeioDePagamento.Comments;

        pagFt.TransferAccount = comAdiantamento ? data.CAIXA_PASSAGEM : data.CAIXA_PRINCIPAL; //para pagamento da factura

        pagFt.TransferSum = totalFactura;
        pagFt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice;
        // pagFt.Invoices.DocEntry =  /*só pode ser preenchido de pois de adicionada a fatura de compra*/
        pagFt.Invoices.SumApplied = totalFactura;

        try
        {
            this.company.StartTransaction();

            if (ftDespesa.Add() != 0)
            {
                var ex = new Exception("Gravar fatura: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                //log the xml to allow easier debug
                var xml = ftDespesa.GetAsXML();
                Logger.Log.Debug(xml, ex);

                throw ex;
            }
            else
            {
                int docEntry = 0;
                int.TryParse(this.company.GetNewObjectKey(), out docEntry);

                if (ftDespesa.GetByKey(docEntry) == false)
                {
                    throw new Exception("Obter fatura: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                }
            }


            if (comAdiantamento)
            {
                if (devAdiant.Add() != 0)
                {
                    var ex = new Exception("Gravar recebimento: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                    //log the xml to allow easier debug
                    var xml = devAdiant.GetAsXML();
                    Logger.Log.Debug(xml, ex);

                    throw ex;
                }

                int docEntry = 0;
                int.TryParse(this.company.GetNewObjectKey(), out docEntry);

                if (devAdiant.GetByKey(docEntry) == false)
                {
                    throw new Exception("Não foi obter o recebimento em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                }

                // Por algum motivo, a DI API não está a gravar o contacto enviado ao criar o documento. Como permite alterar depois, forçamos um update ao documento
                if (devAdiant.ContactPersonCode != data.MeioDePagamento.CntctCode)
                {
                    devAdiant.ContactPersonCode = data.MeioDePagamento.CntctCode;

                    if (devAdiant.Update() != 0)
                    {
                        var ex = new Exception("Atualizar recebimento: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());
                        throw ex;
                    }
                }
            }

            pagFt.Invoices.DocEntry = ftDespesa.DocEntry; /*só pode ser preenchido de pois de adicionada a fatura de compra*/
            if (pagFt.Add() != 0)
            {
                var ex = new Exception("Gravar pagamento: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

                //log the xml to allow easier debug
                var xml = pagFt.GetAsXML();
                Logger.Log.Debug(xml, ex);

                throw ex;
            }
            AddDocResult result = new AddDocResult();

            this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

            result.DocEntry = ftDespesa.DocEntry;
            result.DocNum = ftDespesa.DocNum;

            return result;
        }
        finally
        {
            if (this.company.InTransaction) this.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        }

    }

}