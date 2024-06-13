using Syncfusion.Office;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MathEquationWord2Latex.Model
{
    public class MathPattern
    {
        public IOfficeMathBaseCollection MathBaseColl { get; set; }
        public string LatextPattern { get; set; }
        public MathPattern(IOfficeMathBaseCollection baseColl)
        {
            MathBaseColl = baseColl;
            LatextPattern = "";
        }
    }
}
