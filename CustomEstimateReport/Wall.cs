using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomEstimateReport
{
    class Wall
    {
        public string name { get; set; }
        public string material { get; set; }
        public double volumeGross { get; set; }
        public double volumeNet { get; set; }
        public double length { get; set; }
        public double areaTop { get; set; }
        public double areaEnd1 { get; set; }
        public double areaEnd2 { get; set; }
        public double areaSide1 { get; set; }
        public double areaSide2 { get; set; }
        public double areaSideGross { get; set; }
        public double areaOpening { get; set; }
        public int quantity { get; set; }

        public Wall(string name, string material, double volumeGross, double volumeNet, double length, double areaTop,
            double areaEnd1, double areaEnd2, double areaSide1, double areaSide2, double areaSideGross,
            double areaOpening, int quantity)
        {
            this.name = name;

            this.material = material;

            // Cubic-MM to Cubic-YD
            this.volumeGross = volumeGross * 0.000000001307950619;
            this.volumeGross = Math.Round(this.volumeGross, 2);

            // Cubic-MM to Cubic-YD 
            this.volumeNet = volumeNet * 0.000000001307950619;
            this.volumeNet = Math.Round(this.volumeNet, 2);

            // MM to FT
            this.length = length * 0.00328084;
            this.length = Math.Round(this.length, 2);

            // Square-MM to Square-FT
            this.areaTop = areaTop * 0.0000107639;
            this.areaTop = Math.Round(this.areaTop, 2);

            // Square-MM to Square-FT
            this.areaEnd1 = areaTop * 0.0000107639;
            this.areaEnd1 = Math.Round(this.areaEnd1, 2);

            // Square-MM to Square-FT
            this.areaEnd2 = areaEnd2 * 0.0000107639;
            this.areaEnd2 = Math.Round(this.areaEnd2, 2);

            // Square-MM to Square-FT
            this.areaSide1 = areaSide1 * 0.0000107639;
            this.areaSide1 = Math.Round(this.areaSide1, 2);

            // Square-MM to Square-FT
            this.areaSide2 = areaSide2 * 0.0000107639;
            this.areaSide2 = Math.Round(this.areaSide2, 2);

            // Square-MM to Square-FT
            this.areaSideGross = areaSideGross * 0.0000107639;
            this.areaSideGross = Math.Round(this.areaSideGross, 2);

            // Square-MM to Square-FT
            this.areaOpening = areaOpening * 0.0000107639;
            this.areaOpening = Math.Round(this.areaOpening, 2);

            this.quantity = quantity;
        }
    }
}
