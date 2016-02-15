using System;
using System.IO;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using iTextSharp.tool.xml.html;
using iTextSharp.tool.xml.parser;
using iTextSharp.tool.xml.pipeline.css;
using iTextSharp.tool.xml.pipeline.end;
using iTextSharp.tool.xml.pipeline.html;

namespace SYDQ.Infrastructure.Utilities
{
    //http://demo.itextsupport.com/xmlworker/itextdoc/flatsite.html
    public static class HtmlToPdf
    {
        public enum PdfFont
        {
            //ZhongS,
            Arial
        }

        public static bool GeneratePdf(string htmlText, string fileFullName)
        {
            return GeneratePdf(htmlText, PdfFont.Arial, fileFullName);
        }

        //TODO:HTML中有多字体的情况？
        private static bool GeneratePdf(string htmlText, PdfFont font, string fileFullName)
        {
            if (string.IsNullOrEmpty(htmlText))
            {
                return false;
            }

            htmlText = "<p>" + htmlText + "</p>";

            Document document = new Document();
            MemoryStream outputStream = new MemoryStream();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(fileFullName, FileMode.Create));
            document.Open();

            //pipeline
            HtmlPipelineContext htmlContext = new HtmlPipelineContext(null);
            htmlContext.SetTagFactory(Tags.GetHtmlTagProcessorFactory());
            htmlContext.SetImageProvider(new ChannelImageProvider());

            htmlContext.SetCssAppliers(new CssAppliersImpl(GetFontProviderBy(font)));
            var cssResolver = XMLWorkerHelper.GetInstance().GetDefaultCssResolver(true);
            var pipeline = new CssResolverPipeline(cssResolver,
                new HtmlPipeline(htmlContext, new PdfWriterPipeline(document, writer)));

            //parse
            byte[] data = Encoding.UTF8.GetBytes(htmlText);
            MemoryStream msInput = new MemoryStream(data);
            XMLWorker worker = new XMLWorker(pipeline, true);
            XMLParser parser = new XMLParser(worker);
            parser.Parse(msInput); //XMLWorkerHelper.GetInstance().ParseXHtml(..)
            PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, document.PageSize.Height, 1f);
            PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, writer);
            writer.SetOpenAction(action);

            //close
            document.Close();
            msInput.Close();
            outputStream.Close();

            return true;
        }

        private static IFontProvider GetFontProviderBy(PdfFont font)
        {
            IFontProvider fontProvider = null;
            if (font == PdfFont.Arial)
            {
                fontProvider = new UnicodeFontProvider();
            }

            return fontProvider;
        }
    }

    public class ChannelImageProvider : AbstractImageProvider
    {
        public override string GetImageRootPath()
        {
            return AppDomain.CurrentDomain.BaseDirectory;
        }
    }

    public class UnicodeFontProvider : FontFactoryImp
    {
        private static readonly string ArialFontPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arialuni.ttf");
        //private static readonly string ZhongSFontPath =
        //    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Content\Fonts\STZHONGS.TTF");

        public override Font GetFont(string fontname, string encoding, bool embedded, float size,
            int style, BaseColor color, bool cached)
        {
            BaseFont baseFont = BaseFont.CreateFont(ArialFontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            return new Font(baseFont, size, style, color);
        }
    }
}
