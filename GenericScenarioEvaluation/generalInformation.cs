using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class generalInformation
    {
            public string reviewer { get; set; }
            public string name { get; set; }
            public string year { get; set; }
            public string description { get; set; }
            public string flowDiagram { get; set; }
            public string numActvities { get; set; }
            public string numSources { get; set; }
            public string throughput { get; set; }
            public string concCOI { get; set; }
            public string batchSize { get; set; }
            public string batchDuration { get; set; }
            public string batchPerDay { get; set; }
            public string daysOp { get; set; }
            public string NAICS { get; set; }
            public string facSize { get; set; }
            public string MarketShare { get; set; }

        public List<activityInformation> Activities = new List<activityInformation>();
        public List<equationInformation> Equations = new List<equationInformation>();
    }
}
