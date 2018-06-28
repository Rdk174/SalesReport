using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace SalesReport.Models
{
    public class Orders
    {
        public int? OrderID { get; set; }
        [DataType(DataType.Date)]
        public System.DateTime? OrderDate { get; set; }
        public int? ProductId { get; set; }
        public string ProductName { get; set; }
        public int? Quntity { get; set; }
        public decimal? Price { get; set; }
    }
}