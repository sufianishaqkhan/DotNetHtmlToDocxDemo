using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.IO;

namespace HtmlToOpenXml
{
    class Program
    {
        public static void Main(string[] args)
        {
            //HtmlToOpenXmlConverter htmlToOpenXmlConverter = new HtmlToOpenXmlConverter();
            //htmlToOpenXmlConverter.Run(args);

            //AsposeHtmlConverter asposeHtmlConverter = new AsposeHtmlConverter();
            //asposeHtmlConverter.Run(args);

            SyncFusionHtmlToDocxConverter fusionHtmlToDocxConverter = new SyncFusionHtmlToDocxConverter();
            fusionHtmlToDocxConverter.Run(args);
        }
    }
}
