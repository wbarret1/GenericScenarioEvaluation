using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class Activity2 : IComparable<Activity2>
    {
        public Activity2(string s, string c)
        {
            name = s;
            ChemSteerActivity = c;
        }
        public string name { get; set;  }
        public List<string> ScenariosUsedIn = new List<string>();
        public List<string> ModeledUsing = new List<string>();
        public List<string> years = new List<string>();
        public string ChemSteerActivity { get; set; }
 
        public int CompareTo(Activity2 other)
        {
            return name.CompareTo(other.name);
        }
    }
}
