using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleDriveJob.Models
{
    public class PopularProductsModel
    {
        public string SKU { get; set; }
        public string ProductTitle { get; set; }
        public string Category { get; set; }
        public decimal Price { get; set; }
        public int ItemsAdded { get; set; }
        public int ItemsBought { get; set; }
    }
}
