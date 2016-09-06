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
        
          public void LoadTempalteIntoWorkSheet(string path)
        {
            try
            {
                var xlApp = (MExcel.Application)ExcelDnaUtil.Application;
                //xlApp.Visible = false;

                MExcel.Workbook activeBook = xlApp.ActiveWorkbook;
                //LogUsage("Fundamentals", control.Id, path);

                // Open in same workbook
                var sourceWorkBook = xlApp.Workbooks.Open(Filename: path, ReadOnly: true);
                try
                {
                    // if the active workbook on on initialization is wiped out and the source workbook
                    // becomes the active workbook we need to reset the active workbook to a new book. 
                    if (xlApp.ActiveWorkbook.FullName == sourceWorkBook.FullName && xlApp.Workbooks.Count == 1)
                    {
                        xlApp.Workbooks.Add();
                        activeBook = xlApp.ActiveWorkbook;
                    }

                    xlApp.DisplayAlerts = false;

                    // copy all sheets from the source workbook at once to maintain references within all of the related worksheets. 
                    var lastSheet = activeBook.Worksheets.Count;
                    sourceWorkBook.Worksheets.Copy(After: activeBook.Worksheets[lastSheet]);
                }
                catch (ApplicationException e)
                {
                    var message = string.Format("Unable to load : '{0}' {1}{2}", path, Environment.NewLine, e.Message);
                    log.Error(message + e.Message);
                    MessageBox.Show(message);
                }
                catch (Exception e)
                {
                    var message = string.Format("Unable to load : '{0}' {1}{2}", path, Environment.NewLine, e.Message);
                    log.Error(message + e.Message);
                    MessageBox.Show(message);

                }
                finally
                {
                    sourceWorkBook.Close(Filename: Type.Missing, RouteWorkbook: Type.Missing);
                    //sourceWorkBook.Close(SaveChanges: false, Filename: Type.Missing, RouteWorkbook: Type.Missing);
                    var isSaved = activeBook.Saved;

                    Marshal.ReleaseComObject(sourceWorkBook);
                }

            }
            catch (Exception e)
            {
                var message = string.Format("Error: File selected could not be loaded. {0}", Environment.NewLine);
                log.Error(message + e.Message);
                MessageBox.Show(message);
            }
        }

        public void LoadReportParameter(IRibbonControl control)
        {
            try
            {
                var path = _fileLookup[control.Id];
                var xlApp = (MExcel.Application)ExcelDnaUtil.Application;
                xlApp.Visible = true;

                MExcel.Workbook activeBook = xlApp.ActiveWorkbook;
                LogUsage(UsageResources.ToolsTemplates, control.Id, path);
                if (_openNew)
                {
                    var sourceWorkBook = xlApp.Workbooks.Open(Filename: path, ReadOnly: true);
                }
                else
                {
                    // Open in same workbook
                    var sourceWorkBook = xlApp.Workbooks.Open(Filename: path, ReadOnly: true);
                    try
                    {
                        // if the active workbook on on initialization is wiped out and the source workbook
                        // becomes the active workbook we need to reset the active workbook to a new book. 
                        if (xlApp.ActiveWorkbook.FullName == sourceWorkBook.FullName && xlApp.Workbooks.Count == 1)
                        {
                            xlApp.Workbooks.Add();
                            activeBook = xlApp.ActiveWorkbook;
                        }

                        xlApp.DisplayAlerts = false;

                        // copy all sheets from the source workbook at once to maintain references within all of the related worksheets. 
                        var lastSheet = activeBook.Worksheets.Count;
                        sourceWorkBook.Worksheets.Copy(After: activeBook.Worksheets[lastSheet]);

                        //copy vbaproject
                        CopyMacros(activeBook, sourceWorkBook);
                        CopyReferences(activeBook, sourceWorkBook);
                    }
                    catch (ApplicationException e)
                    {
                        var message = string.Format("Unable to load : '{0}' {1}{2}", path, Environment.NewLine, e.Message);
                        //log.Error(message + e.Message);
                        //MessageBox.Show(message);
                    }
                    catch (Exception e)
                    {
                        var message = string.Format("Unable to load : '{0}'", path);
                        //log.Error(message + e.Message);
                        //MessageBox.Show(message);

                    }
                    finally
                    {
                        sourceWorkBook.Close(SaveChanges: false, Filename: Type.Missing, RouteWorkbook: Type.Missing);
                        var isSaved = activeBook.Saved;
                        Marshal.ReleaseComObject(sourceWorkBook);
                    }
                }
            }
            catch (Exception e)
            {
                var message = string.Format("Error: File selected could not be loaded. {0}", Environment.NewLine);
               // log.Error(message + e.Message);
                //MessageBox.Show(message);
            }
        }

        /// <summary>
        /// Copy Macros  
        /// https://support.microsoft.com/en-us/kb/282830
        /// </summary>
        private void CopyMacros(MExcel.Workbook activeBook, MExcel.Workbook sourceWorkBook)
        {
            if (!sourceWorkBook.HasVBProject) { return; }

            if (!_vbModules.Any())
            {
                foreach (VBComponent dest in activeBook.VBProject.VBComponents)
                {
                    if (!_vbModules.ContainsKey(dest.Name))
                    {
                        _vbModules.Add(dest.Name, dest);
                    }
                }
            }

            try
            {
                foreach (VBComponent sourceComp in sourceWorkBook.VBProject.VBComponents)
                {
                    var sourceLines = sourceComp.CodeModule.CountOfLines;
                    if (sourceLines == 0) { continue; }

                    if (_vbModules.ContainsKey(sourceComp.Name))
                    {
                        VBComponent destComp = _vbModules[sourceComp.Name];

                        if (destComp.CodeModule.CountOfLines == sourceLines)
                        {
                            var sourceCode = sourceComp.CodeModule.Lines[1, sourceLines];
                            var destCode = destComp.CodeModule.Lines[1, destComp.CodeModule.CountOfLines];
                            if (sourceCode == destCode)
                            {
                                continue; // the code is exactly the same as an existing module break and get out. 
                            }
                        }
                        // append
                        if (destComp.CodeModule.CountOfLines == 0)
                        {
                            destComp.CodeModule.InsertLines(1, sourceComp.CodeModule.Lines[1, sourceLines]);
                            continue; // found the matching codemodule so lets exit
                        }
                    }
                    else
                    {
                        // Add new macro
                        var destComp = activeBook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);

                        destComp.CodeModule.AddFromString(sourceComp.CodeModule.Lines[1, sourceLines]);

                        destComp.Name = sourceComp.Name;
                        _vbModules.Add(destComp.Name, destComp);
                    }
                }
            }
            catch (Exception e)
            {
                throw new ApplicationException("Could not copy source macros");
            }

            activeBook.ChangeLink(Name: sourceWorkBook.Name, NewName: activeBook.Name, Type: MExcel.XlLinkType.xlLinkTypeExcelLinks);
        }

        /// <summary>
        /// https://support.microsoft.com/en-us/kb/282830
        /// Remove if not needed
        /// Copy the references from the source workbook to the destination workbook
        /// </summary>
        private void CopyReferences(MExcel.Workbook activeBook, MExcel.Workbook sourceWorkBook)
        {
            if (!_references.Any())
            {
                foreach (Microsoft.Vbe.Interop.Reference reference in activeBook.VBProject.References)
                {
                    if (!_references.ContainsKey(reference.Guid))
                    {
                        _references.Add(reference.Guid, reference.Description);
                    }
                }
            }

            foreach (Microsoft.Vbe.Interop.Reference reference in sourceWorkBook.VBProject.References)
            {
                if (!_references.ContainsKey(reference.Guid))
                {
                    activeBook.VBProject.References.AddFromGuid(reference.Guid, reference.Major, reference.Minor);
                    _references.Add(reference.Guid, reference.Description);
                }
            }
        }

        /// <summary>
        /// Replace - this will be handedl by the BAM.Fundamentals-AddIn.xll.config
        /// anything in the app config will be picked up by the XLL 
        /// </summary>
        /// <returns></returns>
        private Configuration GetAppConfiguration()
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            var extension = ".dll.config";

            var configMap = new ExeConfigurationFileMap
            {
                ExeConfigFilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, string.Format("{0}{1}", assembly, extension))
            };

            var config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);

            return config;
        }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        private string GetAppPath()
        {
            return AppDomain.CurrentDomain.BaseDirectory;
        }

    }
}
