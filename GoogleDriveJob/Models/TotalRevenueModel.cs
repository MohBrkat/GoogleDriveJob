using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace GoogleDriveJob.Models
{
    public class TotalRevenueModel
    {
        [Display(Order = -6)]
        public string InsertDate { get; set; }
        public int OrderId { get; set; }
        public int? BuyerID { get; set; }
        public string OrderTransaction { get; set; }
        public string SourceType { get; set; }
        public string DeliveryOption { get; set; }
        public string SKU { get; set; }
        public decimal ProductPrice { get; set; }
        public decimal? TotalRevenue { get; set; }
    }
}
