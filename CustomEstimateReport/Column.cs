using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomEstimateReport
{
    class Column
    {
        public string name { get; set; }
        public string material { get; set; }
        public double volumeGross { get; set; }
        public double height { get; set; }
        public double areaSide { get; set; }
        public int quantity { get; set; }

        public Column(string name, string material, double volumeGross, double height, double areaSide, int quantity)
        {
            this.name = name;

            this.material = material;

            // Cubic-MM to Cubic-YD
            this.volumeGross = volumeGross * 0.000000001307950619;
            this.volumeGross = Math.Round(this.volumeGross, 2);

            // MM to FT
            this.height = height * 0.00328084;
            this.height = Math.Round(this.height, 2);

            // Square-MM to Square-FT
            this.areaSide = areaSide * 0.0000107639;
            this.areaSide = Math.Round(this.areaSide, 2);

            this.quantity = quantity;
        }
    }
}
