using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrainList.Models
{
    public class Train
    {
        [Key]
        public int Id {get;set;}
        public int TrainNumber { get; set; }
    }

    // Таблица движения
    public class Movements
    {
        [Key]
        public int Id { get; set; }
        public int Train_id { get; set; }
        public string Departure { get; set; }
        public string Destination { get; set; }
        public string Dislocation { get; set; }

        [ForeignKey("Train_id")]
        public Train Trains { get; set; }
    }

    // Таблица состава
    public class Compositions
    {
        [Key]
        public int Id { get; set; }
        public int Train_id { get; set; }
        public string Number { get; set; }
        public string Operations { get; set; }
        public DateTime DateTime { get; set; }
        public string InvoiceNum { get; set; }

        [ForeignKey("Train_id")]
        public Train Trains { get; set; }
    }
    // Таблица вагонов
    public class Wagons
    {
        [Key]
        public int Id { get; set; }
        public int Composition_id { get; set; }
        public int CarNumber { get; set; }
        public int PositionInTrain { get; set; }
        public string FreightEtsngName { get; set; }
        public int FreightTotalWeightKg { get; set; }

        [ForeignKey("Composition_id")]
        public Compositions Compositions { get; set; }
    }


}
