using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

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

        public void OnLogonPressed(IRibbonControl control)
        {
            if (ribbon != null)
            {
                ribbon.InvalidateControl(control.Id);
            }
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            this.ribbon = ribbon;
        }
    }
}
