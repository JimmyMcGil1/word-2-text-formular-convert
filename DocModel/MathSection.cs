using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MathEquationWord2Latex.Model
{
    public class MathSection
    {
        public List<MathParagraph> MathParagraphs { get; set; }
        public int Index { get; set; }
        public MathSection(int index)
        {
            MathParagraphs = new List<MathParagraph>();
            Index = index;
        }
        
    }
}
