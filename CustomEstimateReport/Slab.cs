﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomEstimateReport
{
    class Slab
    {
        public string name { get; set; }
        public string material { get; set; }
        public double volumeGross { get; set; }
        public double volumeNet { get; set; }
        public double areaTop { get; set; }
        public double areaBott { get; set; }
        public double areaEdge { get; set; }
        public double perimeter { get; set; }
        public int quantity { get; set; }

        public Slab(string name, string material, double volumeGross, double volumeNet, double areaEdge, double areaTop, double areaBott, double perimeter, int quantity)
       {
            this.name = name;

            this.material = material;

            // Cubic-MM to Cubic-YD
            this.volumeGross = volumeGross * 0.000000001307950619;
            this.volumeGross = Math.Round(this.volumeGross, 2);

            // Cubic-MM to Cubic-YD 
            this.volumeNet = volumeNet * 0.000000001307950619;
            this.volumeNet = Math.Round(this.volumeNet, 2);

            // Square-MM to Square-FT
            this.areaTop = areaTop * 0.0000107639;
            this.areaTop = Math.Round(this.areaTop, 2);

            // Square-MM to Square-FT
            this.areaBott = areaBott * 0.0000107639;
            this.areaBott = Math.Round(this.areaBott, 2);

            // Square-MM to Square-FT
            this.areaEdge = areaEdge * 0.0000107639;
            this.areaEdge = Math.Round(this.areaEdge, 2);

            // MM to FT
            this.perimeter = perimeter * 0.00328084;
            this.perimeter = Math.Round(this.perimeter, 2);

            this.quantity = quantity;
        }
    }
}
