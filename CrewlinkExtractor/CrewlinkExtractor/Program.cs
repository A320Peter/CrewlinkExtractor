using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Filter;

namespace iText.Samples.Sandbox.Parse
{
    public class ParseCustom
    {
        public static readonly String DEST = "results/txt/parse_custom.txt";

        public static readonly String SRC = "dutyplan.pdf";

        public static void Main(String[] args)
        {
            FileInfo file = new FileInfo(DEST);
            file.Directory.Create();

            new ParseCustom().ManipulatePdf(DEST);
        }

        public virtual void ManipulatePdf(String dest)
        {
            PdfDocument pdfDoc = new PdfDocument(new PdfReader(SRC));

            Rectangle leftdutycolumn = new Rectangle(0, 554, 490, 290);
            Rectangle middledutycolumn = new Rectangle(0, 290, 490, 255);
            Rectangle rightdutycolumn = new Rectangle(0, 0, 490, 290);

            TextRegionEventFilter regionFilter = new TextRegionEventFilter(leftdutycolumn);
            ITextExtractionStrategy strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);

            // Note: If you want to re-use the PdfCanvasProcessor, you must call PdfCanvasProcessor.reset()
            new PdfCanvasProcessor(strategy).ProcessPageContent(pdfDoc.GetFirstPage());

            String leftdutycolumnText = strategy.GetResultantText();

            regionFilter = new TextRegionEventFilter(middledutycolumn);
            strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);
            new PdfCanvasProcessor(strategy).ProcessPageContent(pdfDoc.GetFirstPage());
            String middledutycolumnText = strategy.GetResultantText();

            regionFilter = new TextRegionEventFilter(rightdutycolumn);
            strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);
            new PdfCanvasProcessor(strategy).ProcessPageContent(pdfDoc.GetFirstPage());
            String rightdutycolumnText = strategy.GetResultantText();

            pdfDoc.Close();

            using (StreamWriter writer = new StreamWriter(dest))
            {
                writer.Write(leftdutycolumnText);
                writer.Write(middledutycolumnText);
                writer.Write(rightdutycolumnText);
            }
        }

        // The custom filter filters only the text of which the font name ends with Bold or Oblique.
        protected class CustomFontFilter : TextRegionEventFilter
        {
            public CustomFontFilter(Rectangle filterRect)
                : base(filterRect)
            {
            }

            public override bool Accept(IEventData data, EventType type)
            {
                if (type.Equals(EventType.RENDER_TEXT))
                {
                    TextRenderInfo renderInfo = (TextRenderInfo)data;
                    PdfFont font = renderInfo.GetFont();
                    if (null != font)
                    {
                        String fontName = font.GetFontProgram().GetFontNames().GetFontName();
                        return fontName.EndsWith("Bold") || fontName.EndsWith("Oblique");
                    }
                }

                return false;
            }
        }
    }
}




//namespace CrewlinkExtractor
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            Console.Write("Hello World");
//            Console.Read();
//        }
//    }
//}
