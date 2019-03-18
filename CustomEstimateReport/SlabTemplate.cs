using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomEstimateReport
{
    class SlabTemplate
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
            "AREA",
            "AREA_PGZ",
            "AREA_NGZ",
            "MAINPART.PERIMETER",
        };

        public static readonly ArrayList intProperties = new ArrayList
        {

        };
    }
}
