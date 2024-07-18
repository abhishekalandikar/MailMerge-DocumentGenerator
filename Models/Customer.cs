using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CreateWordSample.Models
{
    public class Customer
    {
        public int CustomerId { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string PhoneNumber { get; set; }
        public string Aadhar { get; set; }
        public string PAN { get; set; }
    }
}