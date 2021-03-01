using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class ExposureInfo
    {
        public ExposureInfo(string Scenario, string Activity, string Type, string Summary)
        {
            sourceSummary = Summary;
            scenarioName = Scenario;
            activity = Activity;
            type = Type;
        }

        public string scenarioName { get; }
        public string activity { get; }
        public string type { get; }
        public string sourceSummary { get; }
    }
}
