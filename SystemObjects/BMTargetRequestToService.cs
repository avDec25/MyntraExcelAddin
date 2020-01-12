using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyntraExcelAddin.SystemObjects
{
    class BMTargetRequestToService
    {
        public string brand;
        public string articleType;
        public string gender;        
        public bool repeated;

        public BMTargetRequestToService(string b, string at, string g, bool r)
        {
            brand = b;
            gender = g;
            articleType = at;
            repeated = r;
        }
    }
}
