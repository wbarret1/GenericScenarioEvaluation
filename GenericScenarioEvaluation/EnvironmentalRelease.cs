using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class EnvironmentalReleaseTypeConverter : System.ComponentModel.ExpandableObjectConverter
    {
        public override bool CanConvertTo(System.ComponentModel.ITypeDescriptorContext context, System.Type destinationType)
        {
            if ((typeof(string)).IsAssignableFrom(destinationType))
                return true;

            return base.CanConvertTo(context, destinationType);
        }

        public override Object ConvertTo(System.ComponentModel.ITypeDescriptorContext context, System.Globalization.CultureInfo culture, Object value, System.Type destinationType)
        {
            if ((typeof(System.String)).IsAssignableFrom(destinationType) && (typeof(EnvironmentalRelease).IsAssignableFrom(value.GetType())))
            {
                return ((EnvironmentalRelease)value).Type;
            }

            return base.ConvertTo(context, culture, value, destinationType);
        }
    };

    [System.ComponentModel.TypeConverter(typeof(EnvironmentalReleaseTypeConverter))]
    public class EnvironmentalRelease
    {
        public int Id { get; set; }
        public GenericScenario GenericScenario { get; set; }
        public Source[] sources { get; set; }
        public int ElementNumber { get; set; }
        public string ScenarioName { get; set; }
        public string ElementName { get; set; }
        public string Type { get; set; }
        public string Type2 { get; set; }
        public string ActivitySource { get; set; }
        public string MediaOfRelease { get; set; }
        public string SourceSummary { get; set; }

        public bool RecycledOrReused
        {
            get
            {
                return this.MediaOfRelease.ToLower().Contains("recycl") ||
                    this.MediaOfRelease.ToLower().Contains("reuse");
            }
        }
        public bool ToAir
        {
            get
            {
                return this.MediaOfRelease.ToLower().Contains("air") ||
                    this.MediaOfRelease.ToLower().Contains("incinerat") ||
                    this.MediaOfRelease.ToLower().Contains("evapor") ||
                    this.MediaOfRelease.ToLower().Contains("potw");
            }
        }
        public bool ToWater
        {
            get
            {
                return this.MediaOfRelease.ToLower().Contains("water") ||
                    this.ElementName.ToLower().Contains("water") ||
                    this.MediaOfRelease.ToLower().Contains("wwt") ||
                    this.MediaOfRelease.ToLower().Contains("injection") ||
                    this.MediaOfRelease.ToLower().Contains("potw");
            }
        }
        public bool ToLand
        {
            get
            {
                return this.MediaOfRelease.ToLower().Contains("land") ||
                    this.ActivitySource.ToLower().Contains("solid") ||
                    this.MediaOfRelease.ToLower().Contains("solid") ||
                    this.MediaOfRelease.ToLower().Contains("Hazard");
            }
        }
        public bool NotSpecified
        {
            get
            {
                return !(this.ToAir || this.ToLand || this.ToWater);
            }
        }
    }
}
