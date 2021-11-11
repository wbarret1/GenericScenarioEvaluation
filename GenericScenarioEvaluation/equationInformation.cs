using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class equationInformation
    {
        public string name { get; set; }
        public string activity { get; set; }
        public string equation { get; set; }
        public string mediaOrRoute { get; set; }
        public string exposureType { get; set; }
        public string exposureComponent { get; set; }
        public string source { get; set; }
        public string variableDescription { get; set; }
        public string variableValue { get; set; }
        public string variableValueUnits { get; set; }
        public string measuredOrEstimated { get; set; }
        public string measurementSource { get; set; }
        public string estimateBasis { get; set; }
        public string equationUsed { get; set; }
        public string reference { get; set; }

        public generalInformation Scenario {get; set;}
    }
}
