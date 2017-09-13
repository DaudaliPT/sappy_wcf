using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


public  enum DocActions { ADD, SIMULATE }


public class AddDocResult
{
    public double DocTotal; 
    public int DocEntry;
    public int DocNum;
    public string message;
    public double DiscountPercent;
    public double TotalDiscount;
    public double VatSum;
    public double RoundingDiffAmount; 
}
