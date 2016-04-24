using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelRibbonWPF;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop;

namespace ExcelRibbonCL
{
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
        }
        
        public void OnLoad(IRibbonUI ribbon)
        {
            this.ribbon = ribbon;
        }
    }
}
