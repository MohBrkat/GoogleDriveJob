using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace GoogleDriveJob.Models
{
    public class NewUsersModel
    {
        [Display(Order = -4)]
        public string InsertDate { get; set; }
        public int UserId { get; set; }
        public string SourceType { get; set; }
    }
}
