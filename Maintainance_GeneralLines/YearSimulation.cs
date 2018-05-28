using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Maintainance_GeneralLines
{
    public class YearSimulation
    {
        public int Iteration { get; set; }
        public double InterArrivalTime { get; set; }
        public double ArrivalTime { get; set; }
        public double ServiceStartTime { get; set; }
        public double WaitingTime { get; set; }
        public double ServiceTime { get; set; }
        public double CompleteTime { get; set; }
        public double TimeinSystem { get; set; }
    }
}
