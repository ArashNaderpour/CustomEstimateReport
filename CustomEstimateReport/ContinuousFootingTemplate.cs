using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomEstimateReport
{
    class ContinuousFootingTemplate
    {
        public static readonly ArrayList stringProperties = new ArrayList
        {
            "MAINPART.NAME",
            "MATERIAL",
        };

        public static readonly ArrayList doubleProperties = new ArrayList
        {
            "VOLUME_GROSS",
            "LENGTH",
            "AREA_PGZ",
            "AREA_PX",
            "AREA_NX",
            "AREA_PZ",
            "AREA_NZ",
        };

        public static readonly ArrayList intProperties = new ArrayList
        {

        };
    }
}
