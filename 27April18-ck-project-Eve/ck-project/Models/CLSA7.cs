using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ck_project.Models
{
    public class CLSA7
    {
        public string Designer { get; set; }
        
        public int Total_Open { get; set; }
        
        public int Total_Sold { get; set; }
        
        public int Total_Deferred { get; set; }
        
        public int Total_Lost_Price { get; set; }
       
        public int Total_Lost_Comp { get; set; }
        
        public int Total_Lost_Other { get; set; }
     
        public int Total_Closed { get; set; }
      
        public int Total { get; set; }
        [DisplayFormat(DataFormatString = "{0:0.00}")]
        public decimal Closed_Percentage { get; set; }
    }
}