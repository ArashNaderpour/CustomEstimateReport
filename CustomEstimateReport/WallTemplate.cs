using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomEstimateReport
{
    class WallTemplate
    {
        public static readonly ArrayList stringProperties = new ArrayList
        {
            "MAINPART.NAME",
            "MATERIAL",
        };

        public static readonly ArrayList doubleProperties = new ArrayList
        {
            "VOLUME_GROSS",
            "VOLUME_NET",
            "LENGTH",
            "AREA_PGZ",
            "AREA_PX",
            "AREA_NX",
            "AREA_PZ",
            "AREA_NZ",
            "AREA_PROJECTION_XY_GROSS",
            "AREA_PROJECTION_XY_NET",
        };

        public static readonly ArrayList intProperties = new ArrayList
        {

        };
    }
}

