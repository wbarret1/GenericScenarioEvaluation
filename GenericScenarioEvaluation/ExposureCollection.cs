using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class ExposureCollection : List<OccupationalExposure>
    {
        public int TotalDermal
        {
            get
            {
                int retVal = 0;
                foreach (OccupationalExposure o in this)
                {
                    if (o.Dermal) retVal++;
                }
                return retVal;
            }
        }

        public int DermalLiquid
        {
            get
            {
                int retVal = 0;
                foreach (OccupationalExposure o in this)
                {
                    if (o.DermalLiquid) retVal++;
                }
                return retVal;
            }
        }

        public int DermalSolid
        {
            get
            {
                int retVal = 0;
                foreach (OccupationalExposure o in this)
                {
                    if (o.DermalSolid) retVal++;
                }
                return retVal;
            }
        }

        public int DermalNotCategorized
        {
            get
            {
                int retVal = 0;
                foreach (OccupationalExposure o in this)
                {
                    if (o.Dermal && !o.DermalLiquid && !o.DermalSolid) retVal++;
                }
                return retVal;
            }
        }

        public int TotalInhalation
        {
            get
            {
                int retVal = 0;
                foreach (OccupationalExposure o in this)
                {
                    if (o.Inhalation) retVal++;
                }
                return retVal;
            }
        }
        public int ChemicalOrVapor
        {
            get
            {
                int retVal = 0;
                foreach (OccupationalExposure o in this)
                {
                    if (o.ChemicalOrVapor) retVal++;
                }
                return retVal;
            }
        }
        public int ParticulateInhalation
        {
            get
            {
                int retVal = 0;
                foreach (OccupationalExposure o in this)
                {
                    if (o.Particulate) retVal++;
                }
                return retVal;
            }
        }
        public int InhalationNotSpecified
        {
            get
            {
                int retVal = 0;
                foreach (OccupationalExposure o in this)
                {
                    if (o.Inhalation && !o.ChemicalOrVapor && !o.Particulate) retVal++;
                }
                return retVal;
            }
        }
        public int NotSpecified
        {
            get
            {
                int retVal = 0;
                foreach (OccupationalExposure o in this)
                {
                    if (!o.Dermal && !o.Inhalation) retVal++;
                }
                return retVal;
            }
        }
    }
}
