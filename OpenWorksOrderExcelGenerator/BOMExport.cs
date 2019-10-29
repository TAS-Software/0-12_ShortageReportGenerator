using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BomExcelGenerator
{
    class BOMExport
    {
        public string SalesOrderTitle { get; set; }
        public Nullable<System.DateTime> PODD { get; set; }
        public Nullable<System.DateTime> CustReqDate { get; set; }
        public string MainPart { get; set; }
        public string MainPartDescription { get; set; }
        public string ComponentPart { get; set; }
        public string ComponentMethod { get; set; }
        public string ComponentPartDesc { get; set; }
        public string Responsibility { get; set; }
        public string ProductGroup { get; set; }
        public string ResourceType { get; set; }
        public string ResourceCode { get; set; }
        public string ResourceGroupName { get; set; }
        public string UnitOfMeasure { get; set; }
        public Nullable<decimal> Quantity { get; set; }
        public Nullable<decimal> TotalBOMQuantity { get; set; }
        public Decimal PriorDemand { get; set; }
        public Nullable<decimal> TotalDemand { get; set; }
        public Nullable<decimal> Stock { get; set; }
        public Nullable<decimal> WoQuantity { get; set; }
        public Nullable<decimal> WoOnTime { get; set; }
        public Nullable<System.DateTime> WoArriving { get; set; }
        public string WoDelayInDays { get; set; }
        public Nullable<decimal> PoQuantity { get; set; }
        public Nullable<decimal> PoOnTime { get; set; }
        public Nullable<System.DateTime> PoArriving { get; set; }
        public string PoDelayInDays { get; set; }
        public Nullable<decimal> Shortage { get; set; }
        public Nullable<decimal> OnTime { get; set; }
        public Nullable<decimal> WOIssued { get; set; }

    }
}
