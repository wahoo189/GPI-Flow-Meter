using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Vironex_GPI_Flow_Meter
{
    [XmlRoot]
    //public class FlowMeterReadings
    //{
    //    public FlowMeterReadings() { }
    //    //public IList<Reading> Items { get; set; }
    //    public Collection<Reading> Items = new Collection<Reading>();
    //}

    public class Reading
    {
        public Reading() {  }

        public decimal Field1 { get; set; }
        public decimal Field2 { get; set; }
        public decimal Field3 { get; set; }
        public decimal Field4 { get; set; }
        public decimal Field5 { get; set; }
        public decimal Field6 { get; set; }
        public decimal Field7 { get; set; }
        public decimal Field8 { get; set; }
        public decimal Field9 { get; set; }
        public decimal Field10 { get; set; }
        public decimal Field11 { get; set; }
        public DateTime Date { get; set; }
    }
}
