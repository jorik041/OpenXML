using System;
using System.Data;
using Interfaces;

namespace OpenXmlPrj
{
    public class DataForTest : IDataForTest
    {
        public String A { get; private set; }
        public String B { get; private set; }
        public String C { get; private set; }

        public DataForTest(String a, String b, String c)
        {
            A = a;
            B = b;
            C = c;
        }

        public DataForTest(DataRow item)
        {
            A = item["MyFieldA"].ToString();
            B = item["MyFieldB"].ToString();
            C = item["MyFieldC"].ToString();
        }
    }
}
