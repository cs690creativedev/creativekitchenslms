using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ck_project.Models
{
    public class CLSA
    {
        public string source_name { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Huntington { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Charleston { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int Lewisburg { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public int total { get; set; }

    }
}