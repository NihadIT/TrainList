using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace TrainList.Models
{
    [Serializable()]
    public class Rows
    {
        public string TrainNumber { get; set; }
        public string TrainIndexCombined { get; set; }
        public string FromStationName { get; set; }
        public string ToStationName { get; set; }
        public string LastStationName { get; set; }
        public string WhenLastOperation { get; set; }
        public string LastOperationName { get; set; }
        public string InvoiceNum { get; set; }
        public string PositionInTrain { get; set; }
        public string CarNumber { get; set; }
        public string FreightEtsngName { get; set; }
        public string FreightTotalWeightKg { get; set; }
    }

    [Serializable()]
    [System.Xml.Serialization.XmlRoot("Root")]
    public class RootCollection
    {
        [XmlElement("row")]
        public Rows[] Row { get; set; }
    }

}
