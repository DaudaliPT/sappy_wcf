using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

public class PostDespesaLinesInput
{
    public string ItemCode;
    public double ValorComIva;
}
public class PostDespesaMeioPagamentoInput
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
}


public class PostDespesaInput
{
    public string CardCode;
    public string TaxDate;
    public int Series;
    public string Comments;
    public double TrocoRecebido;
    public List<PostDespesaLinesInput> Lines;
    public PostDespesaMeioPagamentoInput MeioDePagamento;
    public string CAIXA_PRINCIPAL;
    public string CAIXA_PASSAGEM;
    public string CAIXA_DIFERENCAS;
    public string NumAtCard;
    internal DateTime TaxDateParsed()
    {
        DateTime data;
        DateTime.TryParse(TaxDate, out data);
        return data;
    }
}

