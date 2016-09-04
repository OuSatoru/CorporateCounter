using System;
using System.Collections.Generic;
using System.Text;

namespace System.Runtime.CompilerServices
{
    //Why this
    public class ExtensionAttribute : Attribute { }
}

namespace CorporateCounter
{
    public static class update
    {

        //http://www.cnblogs.com/linlf03/archive/2011/12/09/2282574.html
        public static void Update(this Dictionary<string,string> d, string k, string v)
        {
            if(d.ContainsKey(k) == true)
            {
                d[k] = v;
            }
            else
            {
                d.Add(k, v);
            }
        }
    }
}
