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
using System.Runtime.InteropServices.WindowsRuntime;
using Org.BouncyCastle.Security;

namespace CrewlinkExtractor
{
    public class Extractor
    {

        public static readonly String DEST = "results/txt/parse_custom.txt";
        public static readonly String SRC = "dutyplan.pdf";

        public static void Main(String[] args)
        {
            FileInfo file = new FileInfo(DEST);
            file.Directory.Create();

            PDFPlan dutyplan = new PDFPlan(SRC);
            TextDutyPlan txtdutyplan = new TextDutyPlan(dutyplan.duties, dutyplan.period);

            using (StreamWriter writer = new StreamWriter(Extractor.DEST))
            {
                for (int i = 0; i < txtdutyplan.dutyDay.Length; i++)
                {
                    writer.WriteLine(txtdutyplan.dutyDay[i]);
                }
            }
            Console.Write("Duty Plan in txt format successfully created.");
            Console.Read();
        }
    }

    public class TextDutyPlan
    {
        public String startDate;
        public String endDate;
        public String duties;
        public String[] dutyDay;
        public String miscData;
        private static String[] weekdays = { "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"};

        /*Check the dutyplan text for the first occurence of the weekday for the first duty and split it by checking it for the occurence of the following week day*/
        public void ExtractDutyDays(String dutystream)
        {
            int weekDay = Array.IndexOf(weekdays, dutystream.Substring(0, 3));
            List<string> dutyList = new List<string>();
            int dutyCount = 0;
            while(true)
            {
                weekDay = (weekDay == 6) ? 0 : weekDay + 1;
                if (dutystream.IndexOf(weekdays[weekDay]) != -1)
                {
                    dutyList.Add(dutystream.Substring(0, dutystream.IndexOf(weekdays[weekDay])));
                    dutystream = dutystream.Substring(dutystream.IndexOf(weekdays[weekDay]));
                }
                else
                {
                    dutyList.Add(dutystream);
                    break;
                }
                dutyCount++;
            }
            dutyDay = dutyList.ToArray();
        }


        /* Take the two unedited Strings generated and divide them up into the start date, end date, all duties 
         * and all other stuff at the end of the duty plan */
        public TextDutyPlan(String dutystream, String period)
        {
            startDate = period.Substring(0, 7);
            endDate = period.Substring(period.Length - 7);
            duties = dutystream.Substring(0, dutystream.IndexOf("Flight time"));
            miscData = dutystream.Substring(dutystream.IndexOf("Flight time"));
            ExtractDutyDays(duties);
        }
    }

    public class PDFPlan
    {
        public String duties;
        public String period;

        public PDFPlan(String pdfpath)
        {
            PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfpath));
            ExtractPlanPeriodText(pdfDoc);
            ExtractDutyTexts(pdfDoc);
            pdfDoc.Close();
        }

        public virtual void ExtractPlanPeriodText(PdfDocument pdfDoc)
        {
            Rectangle planpriodare = new Rectangle(563, 638, 9, 111);

            TextRegionEventFilter regionFilter = new TextRegionEventFilter(planpriodare);
            ITextExtractionStrategy strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);

            // Note: If you want to re-use the PdfCanvasProcessor, you must call PdfCanvasProcessor.reset()
            new PdfCanvasProcessor(strategy).ProcessPageContent(pdfDoc.GetFirstPage());
            period = strategy.GetResultantText();
        }

        public virtual String ExtractPageDutys(PdfPage pdfPage)
        {
            Rectangle leftdutycolumn = new Rectangle(0, 554, 490, 290);
            Rectangle middledutycolumn = new Rectangle(0, 290, 490, 255);
            Rectangle rightdutycolumn = new Rectangle(0, 0, 490, 290);

            TextRegionEventFilter regionFilter = new TextRegionEventFilter(leftdutycolumn);
            ITextExtractionStrategy strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);
            new PdfCanvasProcessor(strategy).ProcessPageContent(pdfPage);
            String PageDutyText = strategy.GetResultantText();

            regionFilter = new TextRegionEventFilter(middledutycolumn);
            strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);
            new PdfCanvasProcessor(strategy).ProcessPageContent(pdfPage);
            PageDutyText += strategy.GetResultantText();
            
            regionFilter = new TextRegionEventFilter(rightdutycolumn);
            strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);
            new PdfCanvasProcessor(strategy).ProcessPageContent(pdfPage);
            PageDutyText += strategy.GetResultantText();

            return PageDutyText;
        }

        public virtual void ExtractDutyTexts(PdfDocument pdfDoc)
        {
            String DutyTexts = null;

            for(int i = 1; i<= pdfDoc.GetNumberOfPages(); i++)
            {
                DutyTexts += ExtractPageDutys(pdfDoc.GetPage(i));
            }

            duties = DutyTexts;
        }
    }
}