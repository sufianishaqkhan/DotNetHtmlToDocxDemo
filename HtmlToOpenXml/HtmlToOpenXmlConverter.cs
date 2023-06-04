using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.IO;

namespace HtmlToOpenXml
{
    public class HtmlToOpenXmlConverter
    {
        public void Run(string[] args)
        {
            //string htmlValue = "<!DOCTYPE html><html><body><h1>My First Heading</h1><p>My first paragraph.</p></body></html>";
            string htmlValue = "<span>Hello World!!</span>";
            string destPath = Path.GetFullPath(@"C:\Users\Hp\Desktop\HtmlToOpenXmlDemo.docx");

            ExportToWord(htmlValue, destPath);
        }

        public void ExportToWord(string htmlValue, string destPath)
        {
            byte[] bytes;

            using (MemoryStream stream = new MemoryStream())
            {
                WordprocessingDocument package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);

                MainDocumentPart mainPart = package.MainDocumentPart;
                if (mainPart == null)
                {
                    mainPart = package.AddMainDocumentPart();
                    new Document(new Body()).Save(mainPart);
                }

                HtmlConverter converter = new HtmlConverter(mainPart);
                converter.ParseHtml(htmlValue);

                mainPart.Document.Save();

                bytes = stream.ToArray();

                if (File.Exists(destPath))
                {
                    File.Delete(destPath);
                }

                File.WriteAllBytes(destPath, bytes);
            }
        }
    }
}
