using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ck_project.Models
{
    public class CLSA6
    {
        public string Designer { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_Open { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_Sold { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_Deferred { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_Lost_Price { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_Lost_Comp { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_Lost_Other { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total { get; set; }
    }
}