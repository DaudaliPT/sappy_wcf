using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


public class PostPaymentInputDoc
{
    public int? DocEntry;  
    public int? DocLine;
    public string CardType;

    public string BaseRef;
    public string Ref2;
    public string InvoiceType;
    public int? U_apyCONTRATO;
    public string U_apyUDISC;
    public string U_apyUDEBITO;

    public decimal ValorDebito;
    public string CONTRATO_DESC;

    public decimal ValorECC;
    public decimal ValorECF;
    public decimal ValorDescontoECF;
    public decimal ValorDeduzir;
    public decimal ValorPagar;
    public decimal ValorDescontoPagar;
  
}

public class PostPaymentInputCheques
{
    public string BankCode;
    public string DueDate;
    public int CheckNumber;
    public decimal CheckSum;
    public string CountryCode;
    public string AccounttNum;
    public string CheckAccount;

}


public class PostPaymentInput
{
    public string CardCode;
    public string DocType;
    public decimal CashSum;
    public string Remarks;
    public string CashAccount;

    public string CardCodeEC;
    public string TransferAccountEC;

    public string ArtigoDebitos;
    public string CodigoClienteDebito;
    public int? SerieDebitos;

    public int? U_apyNotas5;
    public int? U_apyNotas10;
    public int? U_apyNotas20;
    public int? U_apyNotas50;
    public int? U_apyNotas100;
    public int? U_apyNotas200;
    public int? U_apyNotas500;
    public decimal U_apyNotas;
    public decimal U_apyMoedas;
    public decimal U_apyVales;
    public decimal U_apyTickets;
    public decimal U_apyTroco;

    public string TransferAccount;
    public decimal TransferSum;
    public string TransferReference;

    public string CheckAccount;
     
    public PostPaymentInputDoc[] PaymentInvoices; 
    public PostPaymentInputCheques[] PaymentChecks;

}
