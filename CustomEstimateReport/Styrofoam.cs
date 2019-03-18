using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomEstimateReport
{
    class Styrofoam
    {
        public string name { get; set; }
        public string material { get; set; }
        public double volumeGross { get; set; }
        public int quantity { get; set; }

        public Styrofoam(string name, string material, double volumeGross, int quantity)
        {
            this.name = name;

            this.material = material;

            // Cubic-MM to Cubic-YD
            volumeGross = volumeGross * 0.000000001307950619;
            this.volumeGross = Math.Round(volumeGross, 2);

            this.quantity = quantity;
        }
    }
}

