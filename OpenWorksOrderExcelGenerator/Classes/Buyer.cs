using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BomExcelGenerator
{
    public class Buyer
    {
        public string BuyerName { get; set; }
        public int SupplierCount { get; set; }
        public int ShortageCount { get; set; }
        public int NotOnOrderCount { get; set; }
        public decimal NotOnOrderQty { get; set; }
        public int UniquePartsNotOnOrder { get; set; }
        public int NotSupportiveCount { get; set; }
        public decimal NotSupportiveQty { get; set; }   
        public int UniquePartsNotSupportingPODD { get; set; }

    }
}
