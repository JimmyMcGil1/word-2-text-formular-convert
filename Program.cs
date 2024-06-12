// See https://aka.ms/new-console-template for more information
using MathEquationWord2Latex;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.Office;
using System.Reflection.Metadata;
using System.Text;

WordDocument document = Utils.GetDocument("D:\\_dot net project\\math_eq_word2latex\\MathEquationWord2Latex\\Template_full.docx");
try
{
    IOfficeMathBaseCollection[] officeMathPatterns = Utils.ReadFormula(document);
    StringBuilder strBuilder = new StringBuilder();
    foreach (var item in officeMathPatterns)
    {
        strBuilder.Append(Utils.ConvertMathPattern(item) + "\n");
    }
    Console.WriteLine(strBuilder.ToString());
    document.Close();
}
catch (Exception e)
{
    Console.WriteLine("Get exception:" + e.Message);
    throw;
}
