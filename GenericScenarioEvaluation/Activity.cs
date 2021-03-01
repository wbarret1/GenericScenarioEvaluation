using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class Activity
    {
        public Activity()
        {
            //IndustryCodes = new List<string>();
            OccupationalExposures = new List<OccupationalExposure>();
            ProcessDescriptions = new List<ProcessDescription>();
            EnvironmentalReleases = new List<EnvironmentalRelease>();
            ControlTechnologies = new List<ControlTechnology>();
            //ProductionRates = new List<ProductionRate>();
            //Workers = new List<Worker>();
            //DataValues = new List<DataValue>();
            //Parameters = new List<RemainingValue>();
            //Calculations = new List<Calculation>();
            //Concentrations = new List<Concentration>();
            //OperatingDays = new List<OperatingDay>();
            //PPEs = new List<PPE>();
            //Sites = new List<Site>();
            //Shifts = new List<Shift>();
            //UseRates = new List<UseRate>();
        }

        public string Name { get; set; }
        public string ChemSTEERActivity { get; set; }
        public List<ProcessDescription> ProcessDescriptions { get; set; }
        public List<OccupationalExposure> OccupationalExposures { get; set; }
        public List<ControlTechnology> ControlTechnologies { get; set; }
        public List<EnvironmentalRelease> EnvironmentalReleases { get; set; }
    }
}
