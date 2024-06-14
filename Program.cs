// See https://aka.ms/new-console-template for more information
using MathEquationWord2Latex;
using Syncfusion.DocIO.DLS;
using System.Text;
using MathEquationWord2Latex.Model;

var document = Utils.GetDocument("D:\\_dot net project\\math_eq_word2latex\\MathEquationWord2Latex\\Template_Full.docx");
List<MathSection> sections = new List<MathSection>();
sections.AddRange(Utils.ReadFormula(document));
Utils.ConvertDocument(sections);
var presentString = Utils.ExtractRawText(document, sections.ToArray());
using (FileStream newFile = new FileStream("D:\\_dot net project\\math_eq_word2latex\\MathEquationWord2Latex\\result.docx", FileMode.OpenOrCreate, FileAccess.Write))
{
    document.Save(newFile, Syncfusion.DocIO.FormatType.Docx);
}
using (StreamWriter writer = new StreamWriter("D:\\_dot net project\\math_eq_word2latex\\MathEquationWord2Latex\\result.txt"))
{
    writer.Write(presentString);
    writer.Close();
}

//Utils.ProcessToFile("D:\\_dot net project\\math_eq_word2latex\\MathEquationWord2Latex\\Template_2.docx",
//    "D:\\_dot net project\\math_eq_word2latex\\MathEquationWord2Latex\\result_2_parse.txt");