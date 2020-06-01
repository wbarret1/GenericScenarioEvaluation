﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class PPE
    {
        public GenericScenario GenericScenario;
        public Source[] sources;
        public string ElementNumber { get; set; }
        public string ScenarioName;
        public string ElementName;
        public string Type;
        public string SourceSummary;
    }
}
