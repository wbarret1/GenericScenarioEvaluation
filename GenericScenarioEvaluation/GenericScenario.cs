using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class GenericScenario
    {
        public GenericScenario()
        {
            OccupationalExposures = new List<OccupationalExposure>();
            ProcessDescriptions = new List<ProcessDescription>();
            EnvironmentalReleases = new List<EnvironmentalRelease>();
            ControlTechnologies = new List<ControlTechnology>();
            ProductionRates = new List<ProductionRate>();
            Workers = new List<Worker>();
            DataValues = new List<DataValue>();
            Parameters = new List<RemainingValue>();
            Workers = new List<Worker>();
            Calculations = new List<Calculation>();
            Concentrations = new List<Concentration>();
            OperatingDays = new List<OperatingDay>();
            PPEs = new List<PPE>();
            Sites = new List<Site>();
            Shifts = new List<Shift>();
            UseRates = new List<UseRate>();
        }

        public string this[string index]
        {
            get
            {
                switch (index)
                {
                    case "Name":
                        return this.ESD_GS_Name;
                    case "Date Prepared":
                        return this.DatePrepared;
                    case "Developed By":
                        return this.DevelopedBy;
                    case "Occupational Exposures":
                        return this.OccupationalExposures.Count.ToString();
                    case "Process Descriptions":
                        return this.ProcessDescriptions.Count.ToString();
                    case "Environmental Releases":
                        return this.EnvironmentalReleases.Count.ToString();
                    case "Control Technologies":
                        return this.ControlTechnologies.Count.ToString();
                    case "Production Rates":
                        return this.ProductionRates.Count.ToString();
                    case "Concentrations":
                        return this.Concentrations.Count.ToString();
                    case "Workers":
                        return this.Workers.Count.ToString();
                    case "Calculations":
                        return this.Calculations.Count.ToString();
                    case "Operating Days":
                        return this.OperatingDays.Count.ToString();
                    case "PPEs":
                        return this.PPEs.Count.ToString();
                    case "Sites":
                        return this.Sites.Count.ToString();
                    case "Shifts":
                        return this.Shifts.Count.ToString();
                    case "Use Rates":
                        return this.UseRates.Count.ToString();
                    case "Data Values":
                        return this.DataValues.Count.ToString();
                    case "Parameters":
                        return this.Parameters.Count.ToString();
                    default:
                        return string.Empty;
                }
            }
        }

        public string[] GetColumns()
        {
            return new string[]
            {
                "Name",
                "Date Prepared",
                "Developed By",
                "Process Descriptions",
                "Occupational Exposures",
                "Environmental Releases",
                "Control Technologies",
                "Production Rates",
                "Concentrations",
                "Workers",
                "Workers",
                "Operating Days",
                "PPEs",
                "Sites",
                "Shifts",
                "Use Rates",
                "Calculations",
                "Data Values",
                "Parameters",
            };
        }

        public int Id { get; set; }
        public string Category { get; set; }
        public string DocumentType { get; set; }
        public string DatePrepared { get; set; }
        public string ESD_GS_Name { get; set; }
        public string FullCitation { get; set; }
        public string DevelopedBy { get; set; }
        public string Description { get; set; }
        public string InPaperIndustryDescriptor { get; set; }
        public string IndustryCodeOrDescription { get; set; }
        public string IndustryCodeType { get; set; }

        public List<ProcessDescription> ProcessDescriptions { get; set; }
        public List<OccupationalExposure> OccupationalExposures { get; set; }
        public List<ControlTechnology> ControlTechnologies { get; set; }
        public List<EnvironmentalRelease> EnvironmentalReleases { get; set; }
        public List<ProductionRate> ProductionRates { get; set; }
        public List<DataValue> DataValues { get; set; }
        public List<RemainingValue> Parameters { get; set; }
        public List<Worker> Workers { get; set; }
        public List<Calculation> Calculations { get; set; }
        public List<Concentration> Concentrations { get; set; }
        public List<OperatingDay> OperatingDays { get; set; }
        public List<PPE> PPEs { get; set; }
        public List<Site> Sites { get; set; }
        public List<Shift> Shifts { get; set; }
        public List<UseRate> UseRates { get; set; }


    }
}
