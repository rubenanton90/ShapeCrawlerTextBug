// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Bibliography;
using ShapeCrawler;


var pres = new Presentation("./samplepptx_es.pptx");
//var textFrames = pres.Slides.SelectMany(e => e.TextFrames());

foreach (var slide in pres.Slides)
{

    foreach (ITextFrame item in slide.TextFrames())
    {
        string t = item.Text;
        item.Text = ReverseText(t);
    }
}

pres.SaveAs($"./samplepptx_es_{DateTime.Now.ToString("_dd_MM_yyyy_HHmmssfff")}.pptx");

static string ReverseText(string s)
{
    char[] charArray = s.ToCharArray();
    Array.Reverse(charArray);
    return new string(charArray);
}