using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ck_project.Models
{
    public class CLSA4
    {
        
        public string Type { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Huntington { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Charleston { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Lewisburg { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Companytotal { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public double Percentage { get; set; }


    }
}