using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class ReleaseCollection :List<EnvironmentalRelease>
    {
        public int ToAir
        {
            get
            {
                int retVal = 0;
                foreach (EnvironmentalRelease r in this)
                {
                    if (r.ToAir) retVal++;
                }
                return retVal;
            }
        }

        public int ToWater
        {
            get
            {
                int retVal = 0;
                foreach (EnvironmentalRelease r in this)
                {
                    if (r.ToWater) retVal++;
                }
                return retVal;
            }
        }
        public int ToLand
        {
            get
            {
                int retVal = 0;
                foreach (EnvironmentalRelease r in this)
                {
                    if (r.ToLand) retVal++;
                }
                return retVal;
            }
        }
        public int NotSpecified
        {
            get
            {
                int retVal = 0;
                foreach (EnvironmentalRelease r in this)
                {
                    if (r.NotSpecified) retVal++;
                }
                return retVal;
            }
        }
    }
}
