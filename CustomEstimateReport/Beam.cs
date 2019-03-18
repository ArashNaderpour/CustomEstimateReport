using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomEstimateReport
{
    class Beam
    {
        public string name { get; set; }
        public string material { get; set; }
        public double volumeGross { get; set; }
        public double length { get; set; }
        public double areaBott { get; set; }
        public double areaSide { get; set; }
        public int quantity { get; set; }


        public Beam(string name, string material, double volumeGross, double length, double areaBott, double areaSide, int quantity)
        {
            this.name = name;

            this.material = material;

            // Cubic-MM to Cubic-YD
            volumeGross = volumeGross * 0.000000001307950619;
            this.volumeGross = Math.Round(volumeGross, 2);

            // MM to FT
            length = length * 0.00328084;
            this.length = Math.Round(length, 2);

            this.areaSide = areaSide;

            // Square-MM to Square-FT
            areaBott = areaBott * 0.0000107639;
            this.areaBott = Math.Round(areaBott, 2);

            this.quantity = quantity;
        }
    }
}
