using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ck_project.Models
{
    public class CLSA11
    {

        //Table 1
        public string Type { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int QTY { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_Amount { get; set; }
        public int Total1 { get; set; }


        //Table 2
        public string Type2 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int QTY2 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_Amount2 { get; set; }
        public int Total2 { get; set; }

        //Table 3
        public string Status { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int QTY3 { get; set; }
        public int Total3 { get; set; }

        //Table 4
        public string YTD_Statistics { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Numerics { get; set; }
        public int Total4 { get; set; }
    }
}