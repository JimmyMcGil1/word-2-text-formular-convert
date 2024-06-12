using DocumentFormat.OpenXml.Wordprocessing;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.Text;
namespace MathEquationWord2Latex
{
    public class Utils
    {
        public static WordDocument GetDocument(string inputFilePath)
        {
            FileStream fileStream = new FileStream(Path.GetFullPath(inputFilePath), FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite);
            WordDocument document = new WordDocument(fileStream, FormatType.Automatic);
            return document;
        }
        public static IOfficeMathBaseCollection[] ReadFormula(WordDocument document)
        {
            try
            {
                List<IOfficeMathBaseCollection> entities = new List<IOfficeMathBaseCollection>();

                for (int i = 0; i < document.Sections.Count; i++)
                {
                    var currSec = document.Sections[i];
                    for (int j = 0; j < currSec.Body.ChildEntities.Count; j++)
                    {
                        var currPara = currSec.Body.ChildEntities[j] as WParagraph;
                        if (currPara != null)
                        {
                            var text = currPara.Text;
                            var entitiesArr = GetMathEntitiesFromParagrapth(currPara);
                            entities.AddRange(entitiesArr);
                        }
                    }
                }
                return entities.ToArray();
            }
            catch (Exception e)
            {
                Console.WriteLine("     [X]Catch error:" + e.Message);
                return [];
            }
        }
        public static string ConvertMathPattern(IOfficeMathBaseCollection collection)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < collection.Count; i++)
            {
                sb.Append(ConvertMathLexical(collection[i]));
            }
            return sb.ToString();
        }
        public static string ConvertMathLexical(IOfficeMathEntity math)
        {
            //Console.WriteLine(math.GetType());
            if (math is IOfficeMathRunElement mathRunElement)
            {
                var runItemText = mathRunElement.Item as WTextRange;
                return runItemText.Text;
            }
            if (math is IOfficeMathFraction mathFrac)
            {
                return ConvertFraction(mathFrac);
            }
            if (math is IOfficeMathRadical mathRad)
            {
                return ConvertRadical(mathRad);
            }
            if (math is IOfficeMathScript script)
            {
                return ConvertScript(script);
            }
            if (math is IOfficeMathAccent accent)
            {
                return ConvertAccent(accent);
            }
            if (math is IOfficeMathLimit limit)
            {
                return ConvertMathLimit(limit);
            }
            if (math is IOfficeMathDelimiter deli)
            {
                return ConvertMathDelimeter(deli);
            }
            return "";
        }
        public static string ConvertFraction(IOfficeMathFraction frac)
        {
            StringBuilder strbuilder = new StringBuilder();
            strbuilder.Append("\\frac{");
            var numerator = frac.Numerator;
            for (int i = 0; i < numerator.Functions.Count; i++)
            {
                var func = numerator.Functions[i];
                strbuilder.Append(ConvertMathLexical(func));
            }
            strbuilder.Append("}{");
            var denominator = frac.Denominator;
            for (int i = 0; i < denominator.Functions.Count; i++)
            {
                var func = denominator.Functions[i];
                strbuilder.Append(ConvertMathLexical(func));
            }
            strbuilder.Append("}");
            return strbuilder.ToString();
        }
        public static string ConvertRadical(IOfficeMathRadical rad)
        {
            //in pratical, need to parse element 
            StringBuilder strBuilder = new StringBuilder();
            if (rad.Degree.Functions.Count > 0)
            {
                strBuilder.Append("\\sqrt[" + ConvertMathLexical(rad.Degree.Functions[0]) + "]{");
            }
            else
            {
                strBuilder.Append("\\sqrt[" + ConvertMathLexical(rad.Degree) + "]{");
            }
            for (int i = 0; i < rad.Equation.Functions.Count; i++)
            {
                strBuilder.Append(ConvertMathLexical(rad.Equation.Functions[i]));
            }
            strBuilder.Append("}");
            return strBuilder.ToString();
        }
        public static string ConvertScript(IOfficeMathScript script)
        {
            StringBuilder strBuilder = new StringBuilder();
            strBuilder.Append(ConvertMathLexical(script.Equation.Functions[0]));
            strBuilder.Append("^{");
            strBuilder.Append(ConvertMathLexical(script.Script.Functions[0]));
            strBuilder.Append("}");
            return strBuilder.ToString();
        }
        public static string ConvertAccent(IOfficeMathAccent mathAccent)
        {
            StringBuilder strBuilder = new StringBuilder();
            //Console.WriteLine("     [-]Accent character:" + mathAccent.AccentCharacter.ToString());
            string accentChar = "widehat";
            //TODO: parse another accent character case
            //switch (mathAccent.AccentCharacter)
            //{
            //    case "?":
            //        accentChar = "bar";
            //        break;
            //    default:
            //        accentChar = "hat";
            //        break;

            //}
            strBuilder.Append($"\\{accentChar}{{");
            for (int i = 0; i < mathAccent.Equation.Functions.Count; i++)
            {
                strBuilder.Append(ConvertMathLexical(mathAccent.Equation.Functions[i]));
            }
            strBuilder.Append("}");
            return strBuilder.ToString();
        }
        public static IOfficeMathBaseCollection[] GetMathEntitiesFromParagrapth(WParagraph paragrapth)
        {
            List<IOfficeMathBaseCollection> entities = new List<IOfficeMathBaseCollection>();
            for (int i = 0; i < paragrapth.ChildEntities.Count; i++)
            {
                var paraItem = paragrapth.ChildEntities[i];
                if (paraItem is WMath)
                {
                    var mathFormular = paraItem as WMath;
                    //for (int j = 0; j < mathFormular.MathParagraph.Maths[0].Functions.Count; j++)
                    //{
                    //    entities.Add(mathFormular.MathParagraph.Maths[0].Functions[j]);
                    //}
                    entities.Add(mathFormular.MathParagraph.Maths[0].Functions);
                }

            }
            return entities.ToArray();
        }
        public static string ConvertMathLimit(IOfficeMathLimit limit)
        {
            StringBuilder sb = new StringBuilder();
            //sb.Append("\\lim_{");
            string lim = ConvertMathPattern(limit.Limit.Functions);
            sb.Append($"\\lim_{{{lim}}}");
            //string equa = ConverktMathPattern(limit.Equation.Functions);
            //sb.Append($"{equa}");
            return sb.ToString();
        }
        public static string ConvertMathDelimeter(IOfficeMathDelimiter delimiter)
        {
            StringBuilder sb = new StringBuilder();
            if (delimiter.BeginCharacter.Equals("(") ||
                delimiter.BeginCharacter.Equals("["))
                sb.Append($"{delimiter.BeginCharacter}");
            else sb.Append($"\\{delimiter.BeginCharacter}");
            sb.Append(ConvertMathLexical(delimiter.Equation));
            if (delimiter.EndCharacter.Equals(")") ||
                delimiter.EndCharacter.Equals("]"))
                sb.Append($"{delimiter.EndCharacter}");
            else sb.Append($"\\{delimiter.EndCharacter}");
            return sb.ToString();
        }
    }

}
