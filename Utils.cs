using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Wordprocessing;
using MathEquationWord2Latex.Model;
using Newtonsoft.Json.Serialization;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.Text;
using System.Text.Json;
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
        public static MathSection[] ReadFormula(WordDocument document)
        {
            try
            {
                List<MathSection> mathSections = new List<MathSection>();

                for (int i = 0; i < document.Sections.Count; i++)
                {
                    var currSec = document.Sections[i];
                    MathSection currMathSec = new MathSection();
                    for (int j = 0; j < currSec.Body.ChildEntities.Count; j++)
                    {
                        var currPara = currSec.Body.ChildEntities[j] as WParagraph;
                        if (currPara != null)
                        {
                            var text = currPara.Text;
                            var entitiesArr = GetMathPatternsFromParagrapth(currPara);
                            MathParagraph mathParagraph = new MathParagraph(entitiesArr, j);
                            if (mathParagraph.MathPatts.Count > 0)
                                currMathSec.MathParagraphs.Add(mathParagraph);
                        }
                    }
                    mathSections.Add(currMathSec);
                }
                return mathSections.ToArray();
            }
            catch (Exception e)
            {
                Console.WriteLine("     [X]Catch error:" + e.Message);
                return [];
            }
        }
        public static string ConvertMathPattern(MathPattern mathPattern)
        {
            StringBuilder sb = new StringBuilder();
            IOfficeMathBaseCollection mathColl = mathPattern.MathBaseColl;
            for (int i = 0; i < mathColl.Count; i++)
            {
                sb.Append(ConvertMathLexical(mathColl[i]));
            }
            return sb.ToString();
        }
        public static string ConvertMathLexicals(IOfficeMathBaseCollection lexicalMaths)
        {
            StringBuilder sbFunc = new StringBuilder();
            for (int i = 0; i < lexicalMaths.Count; i++)
            {
                var currMathLexical = lexicalMaths[i];
                sbFunc.Append(ConvertMathLexical(currMathLexical));
            }

            return sbFunc.ToString();
        }
        public static string ConvertMathLexical(IOfficeMathEntity math)
        {
            Console.WriteLine(math.GetType());
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
                return ConvertLimit(limit);
            }
            if (math is IOfficeMathDelimiter deli)
            {
                return ConvertDelimeter(deli);
            }
            //if (math is IOfficeMathNArray narray)
            //{
            //    return ConvertNArray(narray);
            //}
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

            if (script.ScriptType == MathScriptType.Superscript)
                strBuilder.Append("^{");
            else strBuilder.Append("_{");

            strBuilder.Append(ConvertMathLexical(script.Script.Functions[0]));
            strBuilder.Append("}");
            return strBuilder.ToString();
        }
        public static string ConvertAccent(IOfficeMathAccent mathAccent)
        {
            StringBuilder strBuilder = new StringBuilder();
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
        public static MathPattern[] GetMathPatternsFromParagrapth(WParagraph paragrapth)
        {
            List<MathPattern> entities = new List<MathPattern>();
            for (int i = 0; i < paragrapth.ChildEntities.Count; i++)
            {
                var paraItem = paragrapth.ChildEntities[i];
                if (paraItem is WMath)
                {
                    var mathFormular = paraItem as WMath;
                    entities.Add(new MathPattern(mathFormular.MathParagraph.Maths[0].Functions));
                }

            }
            return entities.ToArray();
        }
        public static string ConvertLimit(IOfficeMathLimit limit)
        {
            StringBuilder sb = new StringBuilder();

            string lim = ConvertMathLexicals(limit.Limit.Functions);
            Console.WriteLine($"Lim: {lim}");
            sb.Append($"\\lim_{{{lim}}}");
            return sb.ToString();
        }
        public static string ConvertDelimeter(IOfficeMathDelimiter delimiter)
        {
            StringBuilder sb = new StringBuilder();
            bool patternHaveSlash = true;
            var beginChar = delimiter.BeginCharacter;
            var endChar = delimiter.EndCharacter;
            StringBuilder sbEqua = new StringBuilder();
            for (int i = 0; i < delimiter.Equation.Count; i++)
            {
                var currEqua = delimiter.Equation[i];
                sbEqua.Append(ConvertMathLexicals(currEqua.Functions));
            }
            string equation = sbEqua.ToString();
            if (beginChar.Equals("(") ||
                beginChar.Equals("["))
                patternHaveSlash = false;
            //sb.Append();
            if (patternHaveSlash)
            {
                sb.Append($"\\{beginChar}{equation}\\{endChar}");
            }
            else
            {
                sb.Append($"{beginChar}{equation}{endChar}");
            }
            return sb.ToString();
        }
        public static void ProcessToFile(string inputFilePath, string outputFilePath)
        {
            WordDocument document = Utils.GetDocument(inputFilePath);
            try
            {
                List<MathSection> mathSections = new List<MathSection>();
                mathSections.AddRange(Utils.ReadFormula(document));
                StringBuilder strBuilder = new StringBuilder();

                foreach (var mathSec in mathSections)
                {
                    foreach (var mathPara in mathSec.MathParagraphs)
                    {
                        foreach (var mathPattern in mathPara.MathPatts)
                        {
                            mathPattern.LatextPattern = Utils.ConvertMathPattern(mathPattern);
                        }
                    }
                }
                int currSec = 0;
                using (StreamWriter writter = new StreamWriter(outputFilePath))
                {
                    foreach (var mathSec in mathSections)
                    {
                        Console.WriteLine();
                        writter.WriteLine($"Current section:{currSec++}");
                        foreach (var mathPara in mathSec.MathParagraphs)
                        {
                            writter.WriteLine($"    [+]Current paragrapth:{mathPara.Index}");
                            foreach (var mathPattern in mathPara.MathPatts)
                            {
                                writter.WriteLine($"         [-]Converted latex math pattern:{mathPattern.LatextPattern}");
                            }
                        }
                    }
                    writter.Close();
                }


                document.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Get exception:" + e.Message);
                throw;
            }

        }
        public static void ProcessToJson(string inputFilePath, string outputFilePath)
        {
            WordDocument document = Utils.GetDocument(inputFilePath);
            try
            {
                List<MathSection> mathSections = new List<MathSection>();
                mathSections.AddRange(Utils.ReadFormula(document));
                StringBuilder strBuilder = new StringBuilder();

                foreach (var mathSec in mathSections)
                {
                    foreach (var mathPara in mathSec.MathParagraphs)
                    {
                        foreach (var mathPattern in mathPara.MathPatts)
                        {
                            mathPattern.LatextPattern = Utils.ConvertMathPattern(mathPattern);
                        }
                    }
                }
                var options = new JsonSerializerOptions() { WriteIndented = true};
                var jsonString = JsonSerializer.Serialize(mathSections, options);
                File.WriteAllText(outputFilePath, jsonString);
            }
            catch (Exception e)
            {
                Console.WriteLine("Get exception:" + e.Message);
                throw;
            }
        }
    }
}