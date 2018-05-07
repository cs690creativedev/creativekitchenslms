using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ck_project.Models
{
    public class CLSA13
    {
        public string Designer { get; set; }
        public DateTime Sold_Date { get; set; }
        public string Responsible_Party { get; set; }
        public string Project_Name { get; set; }
        public string Delivery_Status { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Installed_Total { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Delivered_Total_Before_Taxes { get; set; }

    }
}