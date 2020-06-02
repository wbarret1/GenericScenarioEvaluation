using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class OccupationalExposure
    {
        public int Id { get; set; }
        public GenericScenario GenericScenario { get; set; }
        public Source[] sources { get; set; }
        public string ScenarioName { get; set; }
        public string ElementNumber { get; set; }
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
        public string Activity_Source { get; set; }
        public string mediaOfRelease { get; set; }
        public string sourceSummary { get; set; }
    }
}
