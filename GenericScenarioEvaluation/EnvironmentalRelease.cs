using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class EnvironmentalRelease
    {
        public int Id { get; set; }
        public GenericScenario GenericScenario { get; set; }
        public Source[] sources { get; set; }
        public string ElementNumber { get; set; }
        public string ScenarioName { get; set; }
        public string ElementName { get; set; }
        public string Type { get; set; }
        public string Type2 { get; set; }
        public string ActivitySource { get; set; }
        public string MediaOfRelease { get; set; }
        public string SourceSummary { get; set; }
    }
}
