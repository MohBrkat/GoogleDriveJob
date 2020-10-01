using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace GoogleDriveJob.Models
{
    public class ListsCreatedModel
    {
        [Display(Order = -6)]
        public string CreatedDate { get; set; }
        [Display(Order = -5)]
        public string ListType { get; set; }
        [Display(Order = -4)]
        public string EventDate { get; set; }
        [Display(Order = -3)]
        public string LastPurchasedDate { get; set; }
        [Display(Order = -2)]
        public int ListId { get; set; }
        [Display(Order = -1)]
        public string UniqueURL { get; set; }
        [Display(Order = 0)]
        public int? Products { get; set; }
    }
}
