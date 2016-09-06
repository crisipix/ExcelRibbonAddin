using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelRibbonWPF;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop;

namespace ExcelRibbonCL
{
    /*
        Custom Addin Class 
     */
    [ComVisible(true)]
    public class ExcelRibbonAddin : ExcelRibbon
    {
        /// <summary>
        /// Test Function
        /// Go to excel type in =CoolFunction("Name")
        /// </summary>
        [ExcelFunction(Description = "Cool Name Function")]
        public static string CoolFunction(string name)
        {
            return string.Format("Hello {0} You are Cool", name);
        }

        private IRibbonUI ribbon = null;

        public void OnRefresh(IRibbonControl control)
        {
            if (ribbon != null)
            {
                ribbon.InvalidateControl(control.Id);
            }
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            Application xlApp = (Application)ExcelDnaUtil.Application ;

            MainWindow mw = new MainWindow(xlApp);
            mw.Show();
        }    Application xlApp = (Application)ExcelDnaUtil.Application ;
        
        public void OnLoad(IRibbonUI ribbon)
        {
            this.ribbon = ribbon;
        }

        public void OnLoadSampleTemplate(IRibbonControl ribbon)
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            Workbook workbook;
            if (xlApp.Workbooks.Count < 1)
            {
                workbook = xlApp.Workbooks.Add(1);
            }
            else {
                workbook = xlApp.ActiveWorkbook;
            }
            Workbook sourceWorkbook = null;
            try
            {
                Sheets excelSheets = workbook.Worksheets;
                string currentSheet = "Sheet1";
                Worksheet destinationWs = (Worksheet)excelSheets.get_Item(currentSheet);

                var location = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase;
                var directory = Path.GetDirectoryName(location);
                var filepath = Path.Combine(directory, @"Templates\SampleTemplate.xlsx");
                sourceWorkbook = xlApp.Workbooks.Open(filepath);
                sourceWorkbook.Worksheets[1].Copy(Type.Missing, destinationWs);
                //worksheet1.Copy(Type.Missing, sourceWorkbook.Worksheets[1]);

            }
            catch (Exception e)
            {

            }
            finally {
                //x.DisplayAlerts = false; 
                if (sourceWorkbook != null) {
                    sourceWorkbook.Close();

                }
                
            }
            
        }

        /*
        http://stackoverflow.com/questions/30475709/copy-worksheet-excel-vsto-c-sharp
            Microsoft.Office.Interop.Excel.Application x = new Microsoft.Office.Interop.Excel.Application();
        x.Visible = false; x.ScreenUpdating = false;

        x.Workbooks.Open(Properties.Settings.Default.TemplatePath);

        try
        {
            foreach (Worksheet w in x.Worksheets)
                if (w.Name == wsName)
                    w.Copy(Type.Missing, Globals.ThisAddIn.Application.Workbooks[1].Worksheets[1]);
        }
        catch
        { }
        finally
        {
            x.DisplayAlerts = false; x.Workbooks.Close(); x.DisplayAlerts = true;       // close application with disabled alerts
            x.Quit(); x.Visible = true; x.ScreenUpdating = true;
            x = null;

            My problem was similar. I needed to generate Excel Sheet, then convert to PDF. My problem was that the Excel App was displaying and notifying me to save before Exit. The solution was setting .Visible = flase and .DisplayAlerts = False
        }
            */

        public override void OnBeginShutdown(ref Array custom) {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            Workbook workbook;
            if (xlApp.Workbooks.Count < 1)
            {
                workbook = xlApp.Workbooks.Add(1);
            }
            else {
                workbook = xlApp.ActiveWorkbook;
            }


            Sheets excelSheets = workbook.Worksheets;
            string currentSheet = "Sheet1";
            Worksheet worksheet1 = (Worksheet)excelSheets.get_Item(currentSheet);

          
            /*
            SaveFileDialog saveDialog = new SaveFileDialog();

            If saveDialog.ShowDialog() == DialogResult.OK Then
                 wbook.Close(True, saveDialog.fileName, )
            Else
                 wbook.Close(False, , )
            End If

            wapp.Quit()
            */
            Console.WriteLine("Shutting Down");
            base.OnBeginShutdown(ref custom);
        }
    }
}
