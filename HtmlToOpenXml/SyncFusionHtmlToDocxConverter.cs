using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text;

namespace HtmlToOpenXml
{
    public class SyncFusionHtmlToDocxConverter
    {
        public void Run(string[] args)
        {
            byte[] byteArray;

            string htmlValue = "<!DOCTYPE html><html><body><h1>My First Heading</h1><p>My first paragraph.</p></body></html>";

            string destPath = Path.GetFullPath(@"C:\Users\Hp\Desktop\HtmlToOpenXmlDemo.docx");

            MemoryStream mStrm = new MemoryStream(Encoding.UTF8.GetBytes(htmlValue));

            using (WordDocument document = new WordDocument(@"C:\Users\Hp\Desktop\report.html", FormatType.Html, XHTMLValidationType.None))
            {
                MemoryStream stream = new MemoryStream();
                document.Save(stream, FormatType.Docx);

                byteArray = stream.ToArray();
            }

            if (File.Exists(destPath))
            {
                File.Delete(destPath);
            }

            File.WriteAllBytes(destPath, byteArray);
        }
    }
}
