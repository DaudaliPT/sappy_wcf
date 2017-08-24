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



    /*
     * 


Public Function GerarRevalorizacaoInventario(PrimeiraVal As Boolean, comment As String) As Boolean
    Dim vDocRV          As SAPbobsCOM.MaterialRevaluation
    Dim vDocStk         As SAPbobsCOM.Documents
    Dim Ret             As Long
    Dim X               As Long
    Dim ObjType
    Dim eRrCode         As Long
    Dim eRrMsg          As String
    Dim sUpdatesCount   As Long
    
    Dim docRVtemlinhas As Boolean
    docRVtemlinhas = False
    
    ObjType = oMaterialRevaluation  'REVALORIZA��O INVENT�RIO
    
    On Error GoTo GerarRevalorizacaoInventario_Error
    
    GerarRevalorizacaoInventario = True
   
    If Plan_Artigo_Inv(0) <> "" Then
        Set vDocRV = this.company.GetBusinessObject(oMaterialRevaluation)
        vDocRV.RevalType = "P"
        vDocRV.DocDate = Plan_DocDate_INV
        vDocRV.TaxDate = Plan_DocDate_INV
        vDocRV.Series = sbo.SerieCorrespondenteNoSAP(LerOutrasOpcoes("SERIE_REVINV"), ObjType)
        vDocRV.Comments = comment
        
        If PrimeiraVal = True Then
            Set vDocStk = this.company.GetBusinessObject(oInventoryGenEntry)
            vDocStk.Comments = "Entrada de stock Autom�tico"
            vDocStk.Series = sbo.SerieCorrespondenteNoSAP(LerOutrasOpcoes("SERIE_STOCKENTRADA"), oInventoryGenEntry)
        Else
            Set vDocStk = this.company.GetBusinessObject(oInventoryGenExit)
            vDocStk.JournalMemo = "Sa�da de Mat�rias (Autom�tico)"
            vDocStk.PaymentGroupCode = -3          'Ser� Pre�o m�dio ???, n�o documentado, mas permitido
            vDocStk.Comments = "Sa�da de stock Autom�tico"
            vDocStk.Series = sbo.SerieCorrespondenteNoSAP(LerOutrasOpcoes("SERIE_STOCKSAIDA"), oInventoryGenExit)
        End If
        vDocStk.DocDate = Plan_DocDate_INV
        vDocStk.TaxDate = Plan_DocDate_INV
        
        sUpdatesCount = 0
        For X = 0 To UBound(Plan_Artigo_Inv)
            If Plan_Artigo_Inv(X) <> "" Then
                If Arredondar(myVal(Plan_Stock_Inv(X)) - myVal(Plan_Artigo_Qtd(X)), 6) < 0 Then
                    If vDocStk.LINES.ItemCode <> "" Then vDocStk.LINES.Add
                    vDocStk.LINES.ItemCode = Plan_Artigo_Inv(X) & ""
                    vDocStk.LINES.WarehouseCode = Plan_Armazem_Inv(X) & ""
                    vDocStk.LINES.Quantity = Abs(myVal(Plan_Stock_Inv(X)) - myVal(Plan_Artigo_Qtd(X))) + 1 'DEVIDO A PROBLEMAS COM DECIMAIS, SOMAMOS (E DEPOIS SUBTRA�MOS) "1" � QUANT.
                    vDocStk.LINES.UnitPrice = GetPrecoMedio(Plan_Artigo_Inv(X), Plan_Armazem_Inv(X))
                    sUpdatesCount = sUpdatesCount + 1
                End If
                
                If PrimeiraVal = True Then
                    If Plan_PrecoNew_Inv(X) <> Plan_PrecoOld_Inv(X) Then
                        docRVtemlinhas = True
                        If vDocRV.LINES.ItemCode <> "" Then vDocRV.LINES.Add
                        vDocRV.LINES.ItemCode = Plan_Artigo_Inv(X) & ""
                        vDocRV.LINES.WarehouseCode = Plan_Armazem_Inv(X) & ""
                        vDocRV.LINES.Price = CDbl(Plan_PrecoNew_Inv(X))
                    End If
                Else
                    If Plan_FazSegundaRevalorizacao(X) Then
                        docRVtemlinhas = True
                        If vDocRV.LINES.ItemCode <> "" Then vDocRV.LINES.Add
                        vDocRV.LINES.ItemCode = Plan_Artigo_Inv(X) & ""
                        vDocRV.LINES.WarehouseCode = Plan_Armazem_Inv(X) & ""
                        vDocRV.LINES.Price = CDbl(Plan_PrecoAct_Inv(X))
                    End If
                End If
            End If
        Next
        
        If PrimeiraVal = True Then
            If vDocStk.LINES.ItemCode <> "" Then
                Ret = vDocStk.Add
                If Ret = 0 Then
                    GerarRevalorizacaoInventario = True
                    vDocStk.GetByKey this.company.GetNewObjectKey   'RETORNA O N�MERO DO DOCUMENTO CRIADO
                Else
                    'Deu erro durante a cria��o do Documento
                    this.company.GetLastError eRrCode, eRrMsg
                    sbo.MyMessageBox Traduz("Erro ao gerar a documento: ", "FrmPlanCarga.MSG014r") & eRrMsg
                    GoTo GerarRevalorizacaoInventario_Error
                End If
            End If
        End If
        
        If docRVtemlinhas Then
            Ret = vDocRV.Add
            If Ret = 0 Then
                GerarRevalorizacaoInventario = True
                vDocRV.GetByKey this.company.GetNewObjectKey   'RETORNA O N�MERO DO DOCUMENTO CRIADO
            Else
                'Deu erro durante a cria��o do Documento
                this.company.GetLastError eRrCode, eRrMsg
                sbo.MyMessageBox Traduz("Erro ao gerar a documento: ", "FrmPlanCarga.MSG014") & eRrMsg
                GoTo GerarRevalorizacaoInventario_Error
            End If
        End If
        
        If PrimeiraVal = False Then
            If vDocStk.LINES.ItemCode <> "" Then
                Ret = vDocStk.Add
                If Ret = 0 Then
                    GerarRevalorizacaoInventario = True
                    vDocStk.GetByKey this.company.GetNewObjectKey   'RETORNA O N�MERO DO DOCUMENTO CRIADO
                Else
                    'Deu erro durante a cria��o do Documento
                    this.company.GetLastError eRrCode, eRrMsg
                    sbo.MyMessageBox Traduz("Erro ao gerar a documento: ", "FrmPlanCarga.MSG014r") & eRrMsg
                    GoTo GerarRevalorizacaoInventario_Error
                End If
            End If
        End If
    End If

    Set vDocRV = Nothing
        
   On Error GoTo 0
   Exit Function

GerarRevalorizacaoInventario_Error:
    GerarRevalorizacaoInventario = False
End Function
     */

    internal AddDocResult Confirmar_SAPPY_DOC(string objCode, int draftId, double expectedTotal)
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

        using (HelperOdbc dataLayer = new HelperOdbc())
        using (DataTable headerDt = dataLayer.Execute(sqlHeader))
        using (DataTable detailsDt = dataLayer.Execute(sqlDetail))
        {
            DataRow header = headerDt.Rows[0];

            int objType = Convert.ToInt32(header["OBJTYPE"]);
            bool sujRevalorizacaoNET = (objType == 18);

            if (objType.ToString() != objCode) throw new Exception("objCode " + objCode + " is diferent of " + objType.ToString());


            SAPbobsCOM.Documents newDoc = (SAPbobsCOM.Documents)this.company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)objType);
            newDoc.Series = (int)header["DOCSERIES"];
            DateTime DOCDATE = (DateTime)header["DOCDATE"];
            DateTime DOCDUEDATE = (DateTime)header["DOCDUEDATE"];
            if (DOCDATE.Year > 1900) newDoc.TaxDate = DOCDATE;
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
                if (newDoc.Lines.ItemCode != "") newDoc.Lines.Add();
                var BONUS_NAP = (short)line["BONUS_NAP"];

                newDoc.Lines.ItemCode = (string)line["ITEMCODE"];
                newDoc.Lines.ItemDescription = (string)line["ITEMNAME"];
                newDoc.Lines.MeasureUnit = (string)line["InvntryUom"];
                newDoc.Lines.Factor1 = (double)(decimal)line["QTCX"];   // Num caixas/pack
                newDoc.Lines.Factor2 = (double)(decimal)line["QTPK"];   // Qdd por Caixa/pack 
                //newDoc.Lines.InventoryQuantity = (double)(decimal)line["QTSTK"]; //Definir sobrepoe os fatores 1 e 2
                newDoc.Lines.UnitPrice = (double)(decimal)line["PRICE"];
                newDoc.Lines.WarehouseCode = (string)line["WHSCODE"];
                newDoc.Lines.DiscountPercent = (double)(decimal)line["DISCOUNT"];
                newDoc.Lines.UserFields.Fields.Item("U_apyUDISC").Value = (string)line["USER_DISC"];
                newDoc.Lines.UserFields.Fields.Item("U_apyPRCNET").Value = (double)(decimal)line["NETPRICE"];
                newDoc.Lines.TaxCode = (string)line["VATGROUP"];
                newDoc.Lines.UserFields.Fields.Item("U_apyINCONF").Value = (short)line["HASINCONF"] == 1 ? "Y" : "N";

                if (BONUS_NAP == 0)
                {
                    newDoc.Lines.LineTotal = (double)(decimal)line["LINETOTAL"];
                }
                else
                {
                    // deixa que o SAP calcule o LineTotal na linha atual e na de Bonus

                    if (newDoc.Lines.ItemCode != "") newDoc.Lines.Add();
                    newDoc.Lines.ItemCode = "BONUS";
                    newDoc.Lines.ItemDescription = (string)line["ITEMNAME"];
                    newDoc.Lines.Factor1 = -1 * (double)(decimal)line["QTCX"];      // Num caixas/pack
                    newDoc.Lines.Factor2 = (double)(decimal)line["QTPK"];           // Qdd por Caixa/pack
                    newDoc.Lines.UnitPrice = (double)(decimal)line["PRICE"];
                    newDoc.Lines.TaxCode = (string)line["VATGROUP"];
                }

            }

            // ARREDONDAMENTO DE TOTAL
            if ((double)(decimal)header["ROUNDVAL"] != 0)
            {
                newDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES;
                newDoc.RoundingDiffAmount = (double)(decimal)header["ROUNDVAL"];
            }






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
                invReval = this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation);
                invExit = this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                if (DOCDATE.Year > 1900)
                {
                    invEntry.TaxDate = DOCDATE;
                    invReval.TaxDate = DOCDATE;
                    invExit.TaxDate = DOCDATE;
                }

                foreach (DataRow line in detailsDt.Rows)
                {
                    var qty = (decimal)line["QTSTK"];
                    var onHand = (decimal)line["OnHand"];
                    var newOnHand = onHand + qty;
                    var BONUS_NAP = (short)line["BONUS_NAP"];
                    var docPrice = (decimal)line["PRICE"];
                    var docPriceNet = (decimal)line["NETPRICE"];
                    var transCost = (decimal)line["LINETOTAL"];
                    var transCostNet = docPriceNet * qty;
                    var avgPrice = (decimal)line["AvgPrice"];

                    if (BONUS_NAP == 1) transCost = docPrice * qty; // caso dos bonus

                    decimal stkValue = onHand * avgPrice + transCost;
                    decimal stkValueNet = onHand * avgPrice + transCostNet;
                    decimal sapPmc = 0;
                    decimal pmcNet = 0;

                    if (newOnHand > 0)
                    {
                        sapPmc = Math.Round(stkValue / newOnHand, priceDecimals);
                        pmcNet = Math.Round(stkValueNet / newOnHand, priceDecimals);
                    }
                    if (onHand <= 0) pmcNet = Math.Round(docPriceNet, priceDecimals);


                    if (onHand < 0)
                    {
                        hasFakeEntryExit = true;

                        // add fake stock to make it zero, before
                        if (invEntry.Lines.ItemCode != "") invEntry.Lines.Add();
                        invEntry.Lines.ItemCode = (string)line["ITEMCODE"];
                        invEntry.Lines.WarehouseCode = (string)line["WHSCODE"];
                        invEntry.Lines.InventoryQuantity = (double)onHand * -1;
                        invEntry.Lines.UnitPrice = (double)pmcNet;

                        // remove fake stock, so qty stays correct, after
                        if (invExit.Lines.ItemCode != "") invExit.Lines.Add();
                        invExit.Lines.ItemCode = (string)line["ITEMCODE"];
                        invExit.Lines.WarehouseCode = (string)line["WHSCODE"];
                        invExit.Lines.InventoryQuantity = (double)onHand * -1;
                    }

                    if (sapPmc != pmcNet && newOnHand > 0)
                    {
                        hasRevalorizacao = true;
                        if (invReval.Lines.ItemCode != "") invReval.Lines.Add();
                        invReval.Lines.ItemCode = (string)line["ITEMCODE"];
                        invReval.Lines.WarehouseCode = (string)line["WHSCODE"];
                        invReval.Lines.Price = (double)pmcNet;
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

                if (hasFakeEntryExit)
                {
                    invEntry.Reference2 = newDoc.DocNum.ToString();
                    invEntry.Comments = "Referente a " + newDoc.JournalMemo;
                    if (invEntry.Add() != 0)
                    {
                        var ex = new Exception("Não foi possível gravar fakeEntry em SAP: " + this.company.GetLastErrorCode() + " - " + this.company.GetLastErrorDescription());

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