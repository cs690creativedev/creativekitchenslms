using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ck_project.Models
{
    public class CLSA5_1
    {
        public string Project_Type { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Huntington { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Charleston { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Lewisburg { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Total { get; set; }

        public string Delivery_Type { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Huntington2 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Charleston2 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Lewisburg2 { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Total2 { get; set; }

    }
}