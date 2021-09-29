using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    

    class Model: IComparable<Model>
    {
        public Model(string s)
        {
            name = s;
        }
        public string name { get; }
        public List<string> ScenariosUsedIn = new List<string>();

        public int CompareTo(Model other)
        {
            return name.CompareTo(other.name);
        }
    }
}
