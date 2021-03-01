using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{

    public interface IDataValue
    {
        int ElementNumber {get; set;}
        string ScenarioName { get; set;}
        string ElementName {get; set;}
        string Activity { get; set;}
        string Type { get; set;}
        string Type2 {get; set;}
        string SourceSummary {get; set;}
        Source[] Sources { get;}
    }
    public class RemainingValue: IDataValue
    {
        public int Id { get; set; }
        public int ElementNumber { get; set; }
        public GenericScenario GenericScenario { get; set; }
        public string Activity { get; set;}
        public Source[] Sources { get; set; }

        public string ScenarioName { get; set; }
        public string ElementName { get; set; }
        public string Type { get; set; }
        public string Type2 { get; set; }
        public string ExposureType { get; set; }
        public string Activity_Source { get; set; }
        public string mediaOfRelease { get; set; }
        public string SourceSummary { get; set; }
    }
}
