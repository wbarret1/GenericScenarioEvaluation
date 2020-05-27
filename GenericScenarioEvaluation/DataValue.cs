using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class DataValue
    {
        public int Id { get; set; }
        public GenericScenario GenericScenario { get; set; }
        public Source[] sources { get; set; }
        public string ElementNumber { get; set; }
        public string ElementName { get; set; }
        public string SourceSummary { get; set; }
    }
}
