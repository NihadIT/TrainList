using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrainList.Models
{
    public class TrainView
    {
        public int TrainNumber { get; set; }
        public string FromStationName { get; set; }
        public string ToStationName { get; set; }
        public string LastStationName { get; set; }
        public string CompositionNumber { get; set; }
        public string LastOperationName { get; set; }
        public DateTime WhenLastOperation { get; set; }
        public string InvoiceNum { get; set; }
        [Key]
        public int CarNumber { get; set; }
        public int PositionInTrain { get; set; }
        public string FreightEtsngName { get; set; }
        public int FreightTotalWeightKg { get; set; }
    }
}
