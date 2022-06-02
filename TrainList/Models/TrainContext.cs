using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrainList.Models
{
    public class TrainContext : DbContext
    {
        public TrainContext() : base("Data Source=NIHAD\\SQLEXPRESS;Database=TrainList;Trusted_Connection=True;MultipleActiveResultSets=true")
        { }

        public DbSet<Train> Trains { get; set; }
        public DbSet<TrainView> TrainViews { get; set; }
        public DbSet<Movements> Movements { get; set; }
        public DbSet<Compositions> Compositions { get; set; }
        public DbSet<Wagons> Wagons { get; set; }

    }
}
