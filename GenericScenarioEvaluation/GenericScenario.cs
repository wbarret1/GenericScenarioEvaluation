using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    public class GenericScenario

    {
        public int Id { get; set; }
        public string Category { get; set; }
        public string DocumentType { get; set; }
        public string DatePrepared { get; set; }
        public string ESD_GS_Name { get; set; }
        public string FullCitation { get; set; }
        public string DevelopedBy { get; set; }
        public string Description { get; set; }
        public string InPaperIndustryDescriptor { get; set; }
        public string IndustryCodeOrDescription { get; set; }
        public string IndustryCodeType { get; set; }
    }
}
