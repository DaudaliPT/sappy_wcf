
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

public static class SBOHandler
{
    internal static Dictionary<string, SBOContext> DIAPIConnections = new Dictionary<string, SBOContext>();

    public static void Init()
    {
        var cmpnys = SappyWCF_implementation.Properties.Settings.Default.SAPPY_COMPANYS;

        foreach (var cmp in cmpnys)
        {
            var cn = new SBOContext(cmp);
            DIAPIConnections.Add(cmp, cn);
        }
    }
}
