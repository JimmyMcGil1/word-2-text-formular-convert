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
        public MathSection()
        {
            MathParagraphs = new List<MathParagraph>();
        }
        public MathSection(MathParagraph[] mParagraphs)
        {
            MathParagraphs = new List<MathParagraph>();
            MathParagraphs.AddRange(mParagraphs);
        }
    }
}
