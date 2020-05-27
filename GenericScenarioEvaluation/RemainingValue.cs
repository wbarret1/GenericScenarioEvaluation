using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class RemainingValue
    {
        public int Id { get; set; }
        public string ElementNumber { get; set; }
        public GenericScenario GenericScenario { get; set; }
        public Source[] sources { get; set; }
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
