using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class ReleaseCollectionTypeConverter : System.ComponentModel.ExpandableObjectConverter
    {
        public override bool CanConvertTo(System.ComponentModel.ITypeDescriptorContext context, System.Type destinationType)
        {
            if ((typeof(ReleaseCollectionTypeConverter)).IsAssignableFrom(destinationType))
                return true;

            return base.CanConvertTo(context, destinationType);
        }

        public override Object ConvertTo(System.ComponentModel.ITypeDescriptorContext context, System.Globalization.CultureInfo culture, Object value, System.Type destinationType)
        {
            if ((typeof(System.String)).IsAssignableFrom(destinationType) && (typeof(ReleaseCollectionTypeConverter).IsAssignableFrom(value.GetType())))
            {
                return string.Empty;
                //return ((FunctionalGroupCollection)value).AtomList;
            }

            return base.ConvertTo(context, culture, value, destinationType);
        }
    };
    [System.ComponentModel.TypeConverter(typeof(ReleaseCollectionTypeConverter))]
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
