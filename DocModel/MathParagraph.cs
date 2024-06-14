namespace MathEquationWord2Latex.Model
{
    public class MathParagraph
    {
        public int Index { get; set; }
        public List<MathPattern> MathPatts { get; set; }
        public MathParagraph()
        {
            MathPatts = new List<MathPattern>();
        }
        public MathParagraph(MathPattern[] mathBaseCollection, int index)
        {
            MathPatts = new List<MathPattern>();
            MathPatts.AddRange(mathBaseCollection);
            Index = index;
        }
    }
}
