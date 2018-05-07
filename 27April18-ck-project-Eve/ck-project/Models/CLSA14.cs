using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ck_project.Models
{
    public class CLSA14
    {
        public string Designer { get; set; }
        public int Assigned_Leads { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_amount_all_leads { get; set; }
        public int Sold_Jobs_Only { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_amount_Sold_Jobs { get; set; }        
        public int Lost_Jobs_Only { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_Amount_Lost_Jobs { get; set; }
        public decimal Closed_Percentage { get; set; }

        public string Avg_Days_Sell { get; set; }

        public string Avg_Amount_Sold_Jobs { get; set; }
    }
}