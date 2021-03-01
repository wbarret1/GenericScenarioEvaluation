using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class OccupationalExposureTypeConverter : System.ComponentModel.ExpandableObjectConverter
    {
        public override bool CanConvertTo(System.ComponentModel.ITypeDescriptorContext context, System.Type destinationType)
        {
            if ((typeof(string)).IsAssignableFrom(destinationType))
                return true;

            return base.CanConvertTo(context, destinationType);
        }

        public override Object ConvertTo(System.ComponentModel.ITypeDescriptorContext context, System.Globalization.CultureInfo culture, Object value, System.Type destinationType)
        {
            if ((typeof(System.String)).IsAssignableFrom(destinationType) && (typeof(OccupationalExposure).IsAssignableFrom(value.GetType())))
            {
                return ((OccupationalExposure)value).Type;
            }

            return base.ConvertTo(context, culture, value, destinationType);
        }
    };

    [System.ComponentModel.TypeConverter(typeof(OccupationalExposureTypeConverter))]
    public class OccupationalExposure
    {
        public int Id { get; set; }
        public GenericScenario GenericScenario { get; set; }
        public Source[] sources { get; set; }
        public string ScenarioName { get; set; }
        public int ElementNumber { get; set; }
        public string ElementName { get; set; }
        public string Type { get; set; }
        public string ExposureType { get; set; }

        public bool Dermal
        {
            get
            {
                return this.ExposureType.ToLower().Contains("dermal");
            }
        }
        public bool DermalSolid
        {
            get
            {
                return this.ExposureType.ToLower().Contains("dermal") && this.ExposureType.ToLower().Contains("solid");
            }
        }
        public bool DermalLiquid
        {
            get
            {
                return this.ExposureType.ToLower().Contains("dermal") && this.ExposureType.ToLower().Contains("liquid");
            }
        }
        public bool Inhalation
        {
            get
            {
                return this.ExposureType.ToLower().Contains("inhalation");
            }
        }
        public bool ChemicalOrVapor
        {
            get
            {
                return this.ExposureType.ToLower().Contains("vapor") ||
                    this.ExposureType.ToLower().Contains("chem");
            }
        }
        public bool Particulate
        {
            get
            {
                return this.ExposureType.ToLower().Contains("part");
            }
        }
        public string ActivitySource { get; set; }
        public string mediaOfRelease { get; set; }
        public string SourceSummary { get; set; }
    }
}
