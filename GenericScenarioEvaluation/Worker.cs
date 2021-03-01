using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class Worker : IDataValue
    {
        public GenericScenario GenericScenario { get; set; }
        public Source[] Sources { get; set; }
        public string ScenarioName { get; set; }
        public int ElementNumber { get; set; }
        public string ElementName { get; set; }
        public string Type { get; set; }
        public string Type2 { get; set; }
        public string ExposureType { get; set; }
        public string Activity { get; set; }
        public string mediaOfRelease { get; set; }
        public string SourceSummary { get; set; }
    }
}
