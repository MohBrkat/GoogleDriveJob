using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace GoogleDriveJob.Models
{
    public class ListProductsModel
    {
        [Display(Order = -6)]
        public string InsertDate { get; set; }
        [Display(Order = -5)]
        public int ListProductID { get; set; }
        [Display(Order = -4)]
        public int ListID { get; set; }
        [Display(Order = -3)]
        public int ProductID { get; set; }
        [Display(Order = -2)]
        public string ListProductStatus { get; set; }
        [Display(Order = -1)]
        public int Quantity { get; set; }
        [Display(Order = 0)]
        public decimal Price { get; set; }
    }
}
