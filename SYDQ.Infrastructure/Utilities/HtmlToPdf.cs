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
    //http://www.cnblogs.com/wit13142/p/3690341.html
    public static class HtmlToPdf
    {
        public enum PdfFont
        {
            //ZhongS,
            Arial
        }

        public static bool GeneratePdf(string htmlText, string fileFullName)
        {
            return GeneratePdf(htmlText, PdfFont.Arial, fileFullName, true);
        }

        public static bool GeneratePdfWithWartermark(string htmlText, string fileFullName, string watermarkText = "")
        {
            string tempPdfFileName = fileFullName + ".temp.pdf";
            bool flag = GeneratePdf(htmlText, PdfFont.Arial, tempPdfFileName);
            if (flag)
            {
                watermarkText = string.IsNullOrEmpty(watermarkText) ? "S  Y  D  Q" : watermarkText;
                SetWatermark(tempPdfFileName, fileFullName, watermarkText);
                File.Delete(tempPdfFileName);
            }

            return flag;
        }

        //TODO:HTML中有多字体的情况？
        private static bool GeneratePdf(string htmlText, PdfFont font, string fileFullName, bool hasWatermark = false)
        {
            if (string.IsNullOrEmpty(htmlText))
            {
                return false;
            }

            htmlText = "<p>" + htmlText + "</p>";

            Document document = new Document();
            //MemoryStream outputStream = new MemoryStream(); PdfWriter.GetInstance(document,outputStream); outputStream.ToArray(); 
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(fileFullName, FileMode.Create));
            if (hasWatermark)
            {
                writer.PageEvent = new PdfWatermarkPageEvent("S  Y  D  Q"); //add default watermark    
            }
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
            //outputStream.Close();

            return true;
        }

        private static void SetWatermark(string originalPdfFileName, string finalFileName, string watermarkText)
        {
            using (var pdfReader = new PdfReader(originalPdfFileName))
            {
                using (var newOutputStream = new FileStream(finalFileName, FileMode.Create))
                {
                    using (PdfStamper stamper = new PdfStamper(pdfReader, newOutputStream))
                    {
                        int pageCount = pdfReader.NumberOfPages;
                        PdfLayer layer = new PdfLayer("WatermarkLayer", stamper.Writer);
                        for (int i = 1; i <= pageCount; i++)
                        {
                            Rectangle rect = pdfReader.GetPageSize(i);
                            PdfContentByte watermarkContent = stamper.GetUnderContent(i);
                            //PdfContentByte watermarkContent = stamper.GetOverContent(i);
                            watermarkContent.BeginLayer(layer);
                            watermarkContent.SetFontAndSize(
                                BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 80);

                            PdfGState gState = new PdfGState {FillOpacity = 0.6f};
                            watermarkContent.SetGState(gState);

                            watermarkContent.SetColorFill(BaseColor.LIGHT_GRAY);
                            watermarkContent.BeginText();
                            watermarkContent.ShowTextAligned(PdfContentByte.ALIGN_CENTER, watermarkText, rect.Width/2,
                                rect.Height/2, 45f);
                            watermarkContent.EndText();

                            watermarkContent.EndLayer();
                        }
                    }
                }
            }
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

    public class PdfWatermarkPageEvent : PdfPageEventHelper
    {
        private readonly string _watermarkText;
        private readonly float _fontSize;
        private readonly float? _xPosition;
        private readonly float? _yPosition;
        private readonly float _angle;

        public PdfWatermarkPageEvent(string watermarkText, float fontSize = 80f,
            float angle = 45f, float? xPosition = null, float? yPosition = null)
        {
            _watermarkText = watermarkText;
            _fontSize = fontSize;
            _xPosition = xPosition;
            _yPosition = yPosition;
            _angle = angle;
        }

        public override void OnStartPage(PdfWriter writer, Document document)
        {
            PdfContentByte watermarkContent = writer.DirectContentUnder;
            BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            watermarkContent.BeginText();
            PdfGState gState = new PdfGState
            {
                FillOpacity = 0.5f,
                OverPrintStroking = false
            };
            watermarkContent.SetGState(gState);
            watermarkContent.SetColorFill(BaseColor.LIGHT_GRAY);
            watermarkContent.SetFontAndSize(baseFont, _fontSize);
            Rectangle rect = document.PageSize;
            watermarkContent.ShowTextAligned(PdfContentByte.ALIGN_CENTER, _watermarkText,
                _xPosition ?? rect.Width/2, _yPosition ?? rect.Height/2, _angle);
            watermarkContent.EndText();
        }
    }
}
