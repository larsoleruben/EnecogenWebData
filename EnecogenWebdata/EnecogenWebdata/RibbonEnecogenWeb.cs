using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace EnecogenWebdata
{
    public partial class RibbonEnecogenWeb
    {
        Microsoft.Office.Interop.Excel.Application app;
        Thread webThread;
        Worksheet ws;
        GetWebData gwd = null;
        private void RibbonEnecogenWeb_Load(object sender, RibbonUIEventArgs e)
        {
            editBoxSheetName.Text = DataType.SelectedItem.Tag.ToString();
            app = Globals.ThisAddIn.Application;
            app.SheetChange += app_SheetChange;
        }

        private void buttonStart_Click(object sender, RibbonControlEventArgs e)
        {
            /*Initiate a new sheet and a thread collecting data from Tennet*/
            Boolean existsFlag = false;

            //Does sheet already exists?
            Sheets sheets = app.Worksheets;
            foreach (Worksheet sheet in sheets)
            {
                if (sheet.Name.Equals(editBoxSheetName.Text))
                {
                    existsFlag = true;
                }
            }
            //if it is there, just restart the thread, if it is stopped
            if (existsFlag)
            {
                MessageBox.Show("Sheet is already there. Trying to Restart the collection if it is stopped\n To get a new sheet, delete the old first!", "Info");
                try
                {
                    if (webThread != null)
                    {
                        webThread = null;
                    }
                    if (gwd != null)
                    {
                        webThread = new Thread(new ThreadStart(gwd.retreiveWebData));
                        webThread.Start();
                        MessageBox.Show("Collection succesfully restarted", "Info");
                    }
                    else
                    {
                        MessageBox.Show("Error during restart, delete sheet and start all over", "Error");
                    }
                }
                catch (Exception exp )
                {
                    MessageBox.Show("Error during restart, delete sheet and start all over\n" + exp.ToString(), "Error");
                }
            }
            else //it does not exist
            {
                //check if the hread is running, if it is, try to stop it
                if (webThread != null)
                {
                    if (webThread.IsAlive)
                    {
                        webThread.Abort();
                        webThread.Join();
                    }
                    webThread = null;
                }

                //add the new sheet
                gwd = null;
                try
                {
                    app.Sheets.Add(Type.Missing, Type.Missing, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                    ws = app.ActiveSheet;
                    ws.Name = editBoxSheetName.Text;
                    gwd = new GetWebData(ws, DataType.SelectedItem.Tag.ToString(), Convert.ToInt32(dropDownRefreshInterval.SelectedItem.Tag), Convert.ToInt32(dropDownNumberOfDays.SelectedItem.Tag), app);
                    webThread = new Thread(new ThreadStart(gwd.retreiveWebData));
                    webThread.Start();

                }
                catch (System.Runtime.InteropServices.COMException ce)
                {
                    if (ws != null) { ws.Delete(); }
                    MessageBox.Show(ce.ToString(), "Error");
                }


            }

        }

        private void buttonStop_Click(object sender, RibbonControlEventArgs e)
        {
            if (webThread != null && webThread.IsAlive)
            {
                webThread.Abort();
                webThread.Join();
                MessageBox.Show("Collection has stopped", "Info");
            }
        }

        private void DataTypeParameter_Changes(object sender, RibbonControlEventArgs e)
        {
            editBoxSheetName.Text = DataType.SelectedItem.Tag.ToString();
        }

        private void dropDownNumberOfDays_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void editBoxSheetName_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void app_SheetChange(object Sh, Range Target)
        {
            Worksheet wsChanged = (Worksheet)Sh;


        }


    }
}
