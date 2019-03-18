using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomEstimateReport
{
    class BeamTemplate
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
            "WIDTH",
            "AREA_NGZ",
        };

        public static readonly ArrayList intProperties = new ArrayList
        {

        };
    }
}
