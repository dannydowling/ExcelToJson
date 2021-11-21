using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToJson
{
    internal class Location
    {
        public string? name { get; set; }
        public string? city { get; set; }
        public string? state { get; set; }
        public string? country { get; set; }
        public string? icao { get; set; }

        public double? lat { get; set; }
        public double? lon { get; set; }
    }
}
