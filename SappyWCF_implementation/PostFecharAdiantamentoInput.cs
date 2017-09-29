using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
 
public class PostFecharAdiantamentoInput
{
    public int TransType;
    public int CreatedBy;
    public int TransId;
    public int Line_ID;
    public string TaxDate;
    public double VALOR_ORIGINAL;
    public double VALOR_PENDENTE;
    public string Comments;
    public string CounterRef;
    public string CardCode;
    public int CntctCode;
    public string ContactName;
    internal DateTime TaxDateParsed()
    {
        DateTime data;
        DateTime.TryParse(TaxDate, out data);
        return data;
    }
    public string CAIXA_DIFERENCAS;
}
