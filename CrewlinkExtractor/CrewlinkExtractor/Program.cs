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
using Org.BouncyCastle.Asn1.Mozilla;
using System.Data.OleDb;
using System.Windows;
using System.ComponentModel;
using System.Diagnostics.Eventing.Reader;

namespace CrewlinkExtractor
{
    public class Extractor
    {

        public static readonly String DUTIES = "parse_duties.txt";
        public static readonly String INFO = "parse_info.txt";
        public static readonly String SRC = "dutyplan.pdf";
        private static string connetionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Flightlog.accdb";
        private static OleDbConnection cnn = new OleDbConnection(connetionString);
        private static OleDbCommand command = new OleDbCommand();
        

        public static void Main(String[] args)
        {
            FileInfo duties_file = new FileInfo(DUTIES);
            duties_file.Directory.Create();

            FileInfo duties_info = new FileInfo(INFO);
            duties_info.Directory.Create();

            PDFPlan dutyplan = new PDFPlan(SRC);
            TextDutyPlan txtdutyplan = new TextDutyPlan(dutyplan.duties, dutyplan.period);

            using (StreamWriter writer = new StreamWriter(Extractor.DUTIES))
            {
                for (int i = 0; i < txtdutyplan.dutyDay.Length; i++)
                {
                    writer.WriteLine(txtdutyplan.dutyDay[i]);
                }
            }

            using (StreamWriter writer = new StreamWriter(Extractor.INFO))
            {
                for (int i = 0; i < txtdutyplan.miscData.Length; i++)
                {
                    writer.Write(txtdutyplan.miscData[i]);
                }
            }
            Console.WriteLine("Duty Plan successfully created in txt format.");

            Console.WriteLine(dutyplan.period);
            DutyPlanParser dutyparser = new DutyPlanParser();
            DateTime startDate = dutyparser.ParseDate(txtdutyplan.startDate);

            command.Connection = cnn;
            cnn.Open();

            /*
            for (int i = 0; i < txtdutyplan.dutyDay.Length; i++)
            {
                if (dutyparser.ContainsFlight(txtdutyplan.dutyDay[i]))
                {
                    Flight[] flight = dutyparser.ParseFlights(txtdutyplan.dutyDay[i]);
                    for (int j = 0; j < flight.Length; j++)
                    {
                        writeToDatabase(flight[j].origin, flight[j].destination, flight[j].flightnumber, startDate.AddDays(i), flight[j].startDate, flight[j].endDate, flight[j].activeCrew);
                    }
                }
            }
            */

            DutyDay[] dayArray = dutyparser.ParseDay(txtdutyplan.dutyDay);
            for (int k = 0; k < dayArray.Length; k++)
            {
                for (int l = 0; l < dayArray[k].flights.Length; l++)
                    {
                        writeToDatabase(dayArray[k].flights[l].origin, dayArray[k].flights[l].destination, dayArray[k].flights[l].flightnumber, startDate.AddDays(k), dayArray[k].flights[l].startDate, dayArray[k].flights[l].endDate, dayArray[k].flights[l].activeCrew, dayArray[k].flights[l].activeTakeoff, dayArray[k].flights[l].activeLanding);
                    }
            }

            cnn.Close();
            Console.WriteLine("Duty Plan successfully exported to Access Database.");

            Console.Read();
        }
        public static void writeToDatabase(string origin, string destination, string flightnumber, DateTime date_, string offblock, string onblock, bool activeCrew, bool activeTakeoff, bool activeLanding)
        {

            try
            {
                command.CommandText = "INSERT INTO table_flights (Origin, Destination, Flightnumber, Date_, Offblock, Onblock, ActiveCrew, ActiveTakeoff, ActiveLanding) VALUES ('" + origin + "','" + destination + "','" + flightnumber + "','" + date_ + "','" + offblock + "','" + onblock + "','" + activeCrew.ToString() + "','" + activeTakeoff.ToString() + "','" + activeLanding.ToString() + "')";
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! " + ex);
            }
        }
    }

    public class Flight
    {
        public String origin;
        public String destination;
        public String flightnumber;
        public String startDate;
        public String endDate;
        public bool activeCrew;
        public bool activeTakeoff;
        public bool activeLanding;
    }

    public class DutyDay
    {
        public DateTime date;
        public Flight[] flights;
        //public int checkinTime;
        //public int checkoutTime;
        //public string checkinLoc;
        //public string checkoutLoc;
    }

    public class DutyPlanParser
    {
        public bool ContainsFlight(string dutyDay)
        {
            try
            {
                return (dutyDay.Substring(dutyDay.IndexOf("OS") - 2, 1) != " ") & (dutyDay.Substring(dutyDay.IndexOf("OS") - 1, 1) != " ") & (dutyDay.Substring(dutyDay.IndexOf("OS") + 2, 1) == " ");
            }
            catch
            {
                return false;
            }
        }

        
        public DutyDay[] ParseDay(string[] dutyDay)
        {
            int dummy;
            List<DutyDay> dayList = new List<DutyDay>();
            for(int i = 0; i < dutyDay.Length; i++)
            {
                List<Flight> flight = new List<Flight>();
                String[] textlines = dutyDay[i].Split(new char[] { '\n' });                
                for(int j = 0; j < textlines.Length; j++)
                {
                    String[] line = textlines[j].Split(new char[] { ' ' });
                    // check  if enough line content sufficient to stand for a flight
                    if (line.Length >= 6)
                    {
                        // check if 2nd and 5th value are numbers (flight number and one of the block times)
                        if (Int32.TryParse(line[1], out dummy) & Int32.TryParse(line[4], out dummy))
                        {
                            try
                            {
                                if (textlines[j].Substring(0, 3) == "DH/")
                                {
                                    if (line[2] == "R")
                                    {
                                        flight.Add(new Flight() { activeCrew = false, flightnumber = line[0].Substring(3) + line[1], origin = line[3], startDate = line[4], endDate = line[5], destination = line[6].Substring(0, 3) });
                                    }
                                    else
                                    {
                                        flight.Add(new Flight() { activeCrew = false, flightnumber = line[0].Substring(3) + line[1], origin = line[2], startDate = line[3], endDate = line[4], destination = line[5].Substring(0, 3) });
                                    }
                                }
                                else
                                {
                                    bool takeoff;
                                    bool landing;
                                    switch (textlines[j+1])
                                    {
                                        case "Take-off, Landing":
                                            takeoff = true;
                                            landing = true;
                                            break;
                                        case "Take-off":
                                            takeoff = true;
                                            landing = false;
                                            break;
                                        case "Landing":
                                            takeoff = false;
                                            landing = true;
                                            break;
                                        default:
                                            takeoff = false;
                                            landing = false;
                                            break;
                                            
                                    }
                                    if (line[2] == "R")
                                    {
                                        flight.Add(new Flight() { activeCrew = true, flightnumber = line[0] + line[1], origin = line[3], startDate = line[4], endDate = line[5], destination = line[6].Substring(0, 3), activeTakeoff = takeoff, activeLanding = landing});
                                    }
                                    else
                                    {
                                        flight.Add(new Flight() { activeCrew = true, flightnumber = line[0] + line[1], origin = line[2], startDate = line[3], endDate = line[4], destination = line[5].Substring(0, 3), activeTakeoff = takeoff, activeLanding = landing});
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
                dayList.Add(new DutyDay() { flights = flight.ToArray() });                
            }
            return dayList.ToArray();
        }

        /*
        public Flight[] ParseFlights(string dutyDay)
        {
            List<Flight> flight = new List<Flight>();
            while (dutyDay.IndexOf("OS") > -1)
            {                
                String flightbuffer = dutyDay.Substring(dutyDay.IndexOf("OS") - 3);  // ... , 28 entfernt aus Substring()
                bool activeCrew = !(flightbuffer.Substring(0, 3) == "DH/");
                flightbuffer = flightbuffer.Substring(3);
                String[] flight_buffer = flightbuffer.Split(new char[] { ' ' });
                if (flight_buffer[2] == "R")
                {
                    flight.Add(new Flight() { activeCrew = activeCrew, flightnumber = flight_buffer[0] + flight_buffer[1], origin = flight_buffer[3], startDate = flight_buffer[4], endDate = flight_buffer[5], destination = flight_buffer[6].Substring(0, 3) });
                }
                else
                {
                    flight.Add(new Flight() { activeCrew = activeCrew, flightnumber = flight_buffer[0] + flight_buffer[1], origin = flight_buffer[2], startDate = flight_buffer[3], endDate = flight_buffer[4], destination = flight_buffer[5].Substring(0, 3) });
                }
                dutyDay = dutyDay.Substring(dutyDay.IndexOf("OS") +25);
            }
            return flight.ToArray();
        }
        */

        public DateTime ParseDate(string date)
        {
            int year = Int32.Parse(date.Substring(5)) + 2000;
            int month = ParseMonth(date.Substring(2, 3));
            int day = Int32.Parse(date.Substring(0, 2));
            return new DateTime(year, month, day);
        }

        public int ParseMonth(string month)
        {
            switch (month)
            {
                case "Jan": return 1;
                case "Feb": return 2;
                case "Mar": return 3;
                case "Apr": return 4;
                case "May": return 5;
                case "Jun": return 6;
                case "Jul": return 7;
                case "Aug": return 8;
                case "Sep": return 9;
                case "Oct": return 10;
                case "Nov": return 11;
                case "Dec": return 12;
                default: return 0;
            }
        }
    }

    public class TextDutyPlan
    {
        public String startDate;
        public String endDate;
        public String duties;
        public String[] dutyDay;
        public String miscData;
        int dutydayCount;
        private static String[] weekdays = { "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"};

        /*Check the dutyplan text for the first occurence of the weekday for the first duty and split it by checking it for the occurence of the following week day*/
        public void ExtractDutyDays(String dutystream)
        {
            int weekDay = Array.IndexOf(weekdays, dutystream.Substring(0, 3));
            List<string> dutyList = new List<string>();
            dutydayCount = 0;
            while(true)
            {
                dutystream = dutystream.Remove(5,1);
                dutystream = dutystream.Insert(5, "\n");
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
                dutydayCount++;
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