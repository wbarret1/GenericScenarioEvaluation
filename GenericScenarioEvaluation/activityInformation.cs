using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class activityInformation
    {
            public string name { get; set; }
            public string reviewer { get; set; }
            public string year { get; set; }
            public string activity { get; set; }
            public string chemSteerActivity { get; set; }
            public string Description { get; set; }
            public string ExposureType { get; set; }
            public string exposureValue { get; set; }
            public string expsoureValueUnits { get; set; }
            public string modeled { get; set; }
            public string dataSource { get; set; }
            public string modelName { get; set; }
            public string modelReference { get; set; }

        public generalInformation Scenario { get; set; }
    }
}
