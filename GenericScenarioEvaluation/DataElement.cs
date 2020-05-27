using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class DataElement
    {
        public int Id { get; set; }
        public string Element { get; set; }
        public string ESD_GS_Name { get; set; }
        public string ElementName { get; set; }
        public string Type { get; set; }
        public string Type2 { get; set; }
        public string ExposureType { get; set; }
        public string Activity_Source { get; set; }
        public string mediaOfRelease { get; set; }
        public string SourceSummary { get; set; }
        public string source1;
        public string source2;
        public string source3;
        public string source4;
        public string source5;
        public string source6;
        public string source7;
        public string source8;
        public string Reviewed;
        public string ReferenceCheck;
        public bool accessed = false;
    }
}
