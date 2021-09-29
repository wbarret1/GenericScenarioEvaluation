using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class Activity2 : IComparable<Activity2>
    {
        public Activity2(string s)
        {
            name = s;
        }
        public string name { get; }
        public List<string> ScenariosUsedIn = new List<string>();
 
        public int CompareTo(Activity2 other)
        {
            return name.CompareTo(other.name);
        }
    }
}
