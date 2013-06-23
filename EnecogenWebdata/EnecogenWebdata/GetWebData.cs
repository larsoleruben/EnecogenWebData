using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace EnecogenWebdata
{
    class GetWebData
    {
        private Worksheet ws;
        private String type;
        private int refreshInterval;
        private int numberOfDays;
        private Microsoft.Office.Interop.Excel.Application app;
        int endColNumber = 0;
        String[] alphabet = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH" };



        private const String WEBURL = @"http://www.tennet.org/english/operational_management/export_data.aspx?exporttype=";

        public GetWebData(Worksheet ws, String type, int refreshInterval, int numberOfDays, Microsoft.Office.Interop.Excel.Application app)
        {
            this.ws = ws;
            this.type = type;
            this.refreshInterval = refreshInterval;
            this.numberOfDays = numberOfDays;
            this.app = app;
            startSheet(); //initialize the sheet


        }

        //initialize
        public void startSheet()
        {
            DateTime dt = DateTime.Now;
            DateTime dtF = dt.AddDays(-numberOfDays);
            int dayFrom = dtF.Day;
            int monthFrom = dtF.Month;
            int yearFrom = dtF.Year;
            WebClient webClient = new WebClient();

            try
            {
                String dataUrl2 = WEBURL + type + @"&format=csv&datefrom=" + dayFrom + "-" + monthFrom + "-" + yearFrom + @"&dateto=" + dt.Day + "-" + dt.Month + "-" + dt.Year + @"&submit=1";
                System.Byte[] webData = webClient.DownloadData(dataUrl2);
                String webString = Encoding.UTF8.GetString(webData);
                StringReader reader = new StringReader(webString);
                String line;
                int rowCounter = 1;
                Range totalRange = ws.UsedRange;
                int lastRow = totalRange.Rows.Count;
                String lastDate = ws.Cells[lastRow, 1].Text;
                String lastSeq = ws.Cells[lastRow, 2].Text;
                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Replace("\"", ""); //removing all the double qoutes
                    String[] lineContent = line.Split(',');
                    endColNumber = lineContent.Length;
                    Range range = ws.get_Range("A" + rowCounter, alphabet[endColNumber-1] + rowCounter);
                    range.Value = lineContent;
                    if( rowCounter%2 == 0 )
                    {
                        range.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    }
                    rowCounter++;
                }


                totalRange.EntireColumn.AutoFit();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error getting data: " + ex.ToString(), "Data Error");
            }
        
        }


        public void retreiveWebData()
        {


            WebClient webClient = new WebClient();

            while (true)
            {
                DateTime dt = DateTime.Now;
                
               

                try
                {
                    String dataUrl2 = WEBURL + type + @"&format=csv&datefrom=" + dt.Day+ "-" + dt.Month + "-" + dt.Year + @"&dateto=" + dt.Day + "-" + dt.Month + "-" + dt.Year + @"&submit=1";
                    System.Byte[] webData = webClient.DownloadData(dataUrl2);
                    String webString = Encoding.UTF8.GetString(webData);
                    StringReader reader = new StringReader(webString);
                    String line;
                    int rowCounter = 1;
                    Range totalRange = ws.UsedRange;
                    int lastRow = totalRange.Rows.Count;
                    String lastDate = ws.Cells[lastRow, 1].Text;
                    int lastSeq = Convert.ToInt32( ws.Cells[lastRow, 2].Text );
                    CultureInfo provider = CultureInfo.InvariantCulture;
                    DateTime dtLast = new DateTime();
                    //TODO This works but is not very elegant
                    try
                    {
                        dtLast = DateTime.ParseExact( lastDate, "DD/MM/YYYY", provider );
                    }catch( FormatException fe )
                    {
                        Console.WriteLine("Date not in the correct format");
                    }
                    try
                    {
                        dtLast = DateTime.ParseExact(lastDate, "MM/DD/YYYY", provider);
                    }
                    catch (FormatException fe)
                    {
                        Console.WriteLine("Date not in the correct format");
                    }
                    while ((line = reader.ReadLine()) != null)
                    {
                        //skip the first line with text
                        if (rowCounter > 1)
                        {
                            line = line.Replace("\"", ""); //removing all the double qoutes
                            String[] lineContent = line.Split(',');
                            endColNumber = lineContent.Length;
                            DateTime dtLastFile = new DateTime();// = DateTime.Parse(lineContent[0]);
                            //TODO This works but is not very elegant
                            try
                            {
                                dtLastFile = DateTime.ParseExact(lineContent[0], "DD/MM/YYYY", provider);
                            }
                            catch (FormatException fe)
                            {
                                Console.WriteLine("Date not in the correct format");
                            }
                            try
                            {
                                dtLastFile = DateTime.ParseExact(lineContent[0], "MM/DD/YYYY", provider);
                            }
                            catch (FormatException fe)
                            {
                                Console.WriteLine("Date not in the correct format");
                            }



                            int lastSeqFile = Convert.ToInt32(lineContent[1]);

                            if (dtLastFile > dtLast && lastSeq > lastSeqFile)
                            {
                                Range range = ws.get_Range("A" + (lastRow + 1), alphabet[endColNumber-1] + (lastRow + 1));
                                range.Value = lineContent;
                                if (Convert.ToInt32(range.Cells[1, 2].Text) % 2 > 0)
                                {
                                    colorCells(range, System.Drawing.Color.LightGray, System.Drawing.Color.DarkGray);
                                }
                                else
                                {
                                    colorCells(range, System.Drawing.Color.White, System.Drawing.Color.DarkGray);
                                }
                                lastRow++;
                                range = null;
                            }
                            else if (lastSeq < lastSeqFile && dtLastFile == dtLast )
                            {
                                /*ToDo Sometimes it will skip one row*/
                                Range range = ws.get_Range("A" + (lastRow + 1), alphabet[endColNumber-1] + (lastRow + 1));
                                range.Value = lineContent;
                                if (Convert.ToInt32(range.Cells[1, 2].Text) % 2 > 0)
                                {
                                    colorCells(range, System.Drawing.Color.LightGray, System.Drawing.Color.DarkGray);
                                }
                                else
                                {
                                    colorCells(range, System.Drawing.Color.White, System.Drawing.Color.DarkGray);
                                }
                                lastRow++;
                                range = null;
                            }
                        }

                        rowCounter++;
                    }

                    DateTime dtForw = dt.AddMinutes(2);
                    int calcRows = (1+numberOfDays) * 1440;
                    if(lastRow > calcRows )
                    {
                        int rowToBeDeleted = lastRow - calcRows + 1; //there is an offset of 1
                        Range deleteRange = ws.get_Range("A2", alphabet[endColNumber - 1]+rowToBeDeleted.ToString()).EntireRow;
                        deleteRange.Delete(XlDeleteShiftDirection.xlShiftUp);
                    }
                    lastRow = ws.UsedRange.Rows.Count;
                    Range endRange = ws.get_Range("A" + (lastRow), alphabet[endColNumber-1] + (lastRow));
                    endRange.Select();
                    if (Convert.ToInt32(endRange.Cells[1, 2].Text) % 2 > 0)
                    {
                        colorCells(endRange, System.Drawing.Color.LightGray, System.Drawing.Color.DarkGray);
                 
                    }
                    else
                    {
                        colorCells(endRange, System.Drawing.Color.White, System.Drawing.Color.DarkGray);
                    }
                    endRange = null;
                    ws.UsedRange.EntireColumn.AutoFit();

                }
                catch (Exception ex)
                {
                
                    MessageBox.Show("Sheet: " +ws.Name + "Error getting data: " +  ex.ToString(), "Data Error");
                }
                GC.Collect();
                ws.UsedRange.EntireColumn.AutoFit();
                Thread.Sleep(refreshInterval * 60000);
            }

        }

        private void colorCells( Range range, System.Drawing.Color colorBg, System.Drawing.Color colorBorder  )
        {
            range.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(colorBg);
            range.Cells.Borders.Color = System.Drawing.ColorTranslator.ToOle(colorBorder);
        }
    }
}
