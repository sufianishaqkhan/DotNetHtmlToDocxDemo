using Aspose.Html;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.IO;

namespace HtmlToOpenXml
{
    public class AsposeHtmlConverter
    {
        public void Run(string[] args)
        {
            string htmlValue = "<!DOCTYPE html><html><body><h1>My First Heading</h1><p>My first paragraph.</p></body></html>";
            string destPath = Path.GetFullPath(@"C:\Users\Hp\Desktop\HtmlToOpenXmlDemo.docx");

            ExportToWord(htmlValue, destPath);
        }

        public void ExportToWord(string htmlValue, string destPath)
        {
            var code = @"<span>Hello World!!</span>";
            System.IO.File.WriteAllText("document.html", code);

            // Initialize an HTML document from the file
            using (var document = new HTMLDocument("document.html"))
            {
                // Initialize DocSaveOptions 
                var options = new Aspose.Html.Saving.DocSaveOptions();

                // Convert HTML webpage to DOCX
                Aspose.Html.Converters.Converter.ConvertHTML(document, options, "output.docx");
            }
        }
    }
}
