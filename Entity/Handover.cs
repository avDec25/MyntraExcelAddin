using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyntraExcelAddin.Entity
{
    class Handover
    {
        public Handover()
        {
            fabrics = new List<Fabric>();
        }
        public String vanId;
        public String brand;
        public String articleType;
        public String gender;
        public Double quantity;
        public String cluster;
        public String subcategory;
        public List<Fabric> fabrics;
        public String dropName;
        public String mrpRange;
        public Double bmTarget;
        public String sizeType;
        public String bodyCode;
        public String dataSource;
        public String dataSourceDetails;
        public String color;
        public Boolean isWashReferenced;
        public String pdpCatalogCallouts;
        public String source;
    }
}