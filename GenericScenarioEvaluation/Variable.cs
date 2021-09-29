using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericScenarioEvaluation
{
    class Variable : IComparable<Variable>
    {
        public Variable(string s)
        {
            name = s;
        }
        public string name { get; }
        public List<string> ScenariosUsedIn = new List<string>();

        public int CompareTo(Variable other)
        {
            return name.CompareTo(other.name);
        }
    }
}
