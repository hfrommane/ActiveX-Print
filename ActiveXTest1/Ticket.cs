using System;
using System.Collections.Generic;

namespace ActiveXTest1
{
    public class Ticket
    {   
        public Boolean isKitchen { get; set; }
        public Boolean isPreview { get; set; }
        public String printName { get; set; }
        public String restaurant { get; set; }
        public String orderNo { get; set; }
        public String desk { get; set; }
        public String type { get; set; }
        public List<Menu> menu{ get; set; }
        public String pay { get; set; }
        public String address { get; set; }
        public String telephone { get; set; }
        public String mobilephone { get; set; }

        public class Menu
        {
            public String name { get; set; }
            public String price { get; set; }
            public String count { get; set; }
            public String notes { get; set; }
        }
    }

}
