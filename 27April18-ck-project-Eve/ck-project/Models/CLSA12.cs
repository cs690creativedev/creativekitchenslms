using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ck_project.Models
{
    public class CLSA12
    {
        //Table 1
        public string State { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int No_of_Leads { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int No_of_Leads_Sold { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Total_amount_of_Sold_Jobs { get; set; }
        public int Total1 { get; set; }


        //Table 2
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int QTY { get; set; }        
        public string City { get; set; }
        public int Total2 { get; set; }


        //Table 3
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int QTY3 { get; set; }
        public string City3{ get; set; }
        public int Total3 { get; set; }

        //Table 4
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int QTY4 { get; set; }
        public string City4 { get; set; }
        public int Total4 { get; set; }

        //Table 5
        public string State5 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Installed5 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Pickup5 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Delivered5 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double In_City_Installed5 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Out_City_Installed5{ get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Remodel5 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double New_Construction5 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Builder5 { get; set; }
        public int Total5 { get; set; }

        //Table 6
        public string Total { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Installed6 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Pickup6 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Delivered6 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double In_City_Installed6 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Out_City_Installed6 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Remodel6 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double New_Construction6 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Builder6 { get; set; }
        public int Total6 { get; set; }
    }
}