using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomEstimateReport
{
    class Footing
    {
        public string name { get; set; }
        public string material { get; set; }
        public double volumeGross { get; set; }
        public double areaTop { get; set; }
        public double areaBott { get; set; }
        public double areaEdge { get; set; }
        public int quantity { get; set; }

        public Footing(string name, string material, double volumeGross, double areaEdge, double areaTop, double areaBott, int quantity)
        {
            this.name = name;

            this.material = material;

            // Cubic-MM to Cubic-YD
            this.volumeGross = volumeGross * 0.000000001307950619;
            this.volumeGross = Math.Round(this.volumeGross, 2);

            // Square-MM to Square-FT
            this.areaTop = areaTop * 0.0000107639;
            this.areaTop = Math.Round(this.areaTop, 2);

            // Square-MM to Square-FT
            this.areaBott = areaBott * 0.0000107639;
            this.areaBott = Math.Round(this.areaBott, 2);

            // Square-MM to Square-FT
            this.areaEdge = areaEdge * 0.0000107639;
            this.areaEdge = Math.Round(this.areaEdge, 2);

            this.quantity = quantity;
        }
    }
}

