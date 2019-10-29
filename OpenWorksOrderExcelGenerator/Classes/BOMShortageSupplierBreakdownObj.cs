using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BomExcelGenerator
{
    public class BOMShortageSupplierBreakdownObj
    {
        //public string SalesOrderTitle { get; set; }
        public Nullable<System.DateTime> DespatchDate { get; set; }
        public Nullable<System.DateTime> PODD { get; set; }
        public string ComponentPart { get; set; }
        public string ComponentPartDescription { get; set; }
        public string ComponentMethod { get; set; }
        public string ProductGroup { get; set; }
        public string Responsibility { get; set; }
        public int LeadTime { get; set; }
        public string SupplierName { get; set; }
        public Nullable<decimal> Quantity { get; set; }
        public Decimal PriorDemand { get; set; }
        public Nullable<decimal> TotalDemand { get; set; }
        public Nullable<decimal> Stock { get; set; }
        public Nullable<decimal> WoQuantity { get; set; }
        public Nullable<decimal> WoOnTime { get; set; }
        public Nullable<decimal> PoQuantity { get; set; }
        public Nullable<decimal> PoOnTime { get; set; }
        public Nullable<System.DateTime> POArriving { get; set; }
        public string PODelayInDays { get; set; }
        public Nullable<decimal> Shortage { get; set; }
        public Nullable<decimal> OnTime { get; set; }
        public Nullable<decimal> WOIssued { get; set; }
        public string ReportDate { get; set; }
        public decimal UnitCost { get; set; }
        public string BuyerName { get; set; }       
    }
}