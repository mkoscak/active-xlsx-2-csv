using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ActiveConverter
{
    public class DataItem
    {
        public DataItem()
        {
            Codes = new List<string>();
            Sizes = new List<string>();
        }

        public List<string> Codes { get; set; }
        public string Description { get; set; }
        public string Rrp { get; set; }
        public string Supp { get; set; }
        public string CatId { get; set; }

        public List<string> Sizes { get; set; }

        public object Picture { get; set; }
    }
}
