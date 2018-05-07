using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ck_project.Models
{
    public class CLSA17
    {
        //Table 1
        public string Branch_Name { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double YTD_Total_Sales { get; set; }
        public int Total1 { get; set; }

        //Table 2
        public string Month { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Price { get; set; }
        public int Total2 { get; set; }

        //Table 3
        public string Month3 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Price3 { get; set; }
        public int Total3 { get; set; }

        //Table 4
        public string Month4 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Price4 { get; set; }
        public int Total4 { get; set; }
    }
}