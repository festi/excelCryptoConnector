using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Net;

namespace CryptoConnector
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            // up ssl default security to TLS1.2 for compatibilityu with some exchanges (binance and cryptopia when this comment was written)
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public Microsoft.Office.Interop.Excel.Worksheet FindWorksheet(string name, bool createIfNotFound = false)
        {
            bool wasCreated;
            return FindWorksheet(name, createIfNotFound, out wasCreated);
        }

        public Microsoft.Office.Interop.Excel.Worksheet FindWorksheet(string name, bool createIfNotFound , out bool wasCreated)
        {
            Microsoft.Office.Interop.Excel.Worksheet res = null;
            wasCreated = false;

            Microsoft.Office.Interop.Excel.Sheets sheets = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet w in sheets)
            {
                if (w.Name == name)
                {
                    res = w;
                    break;
                }
            }

            if (res == null && createIfNotFound)
            {
                Microsoft.Office.Interop.Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

                res = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                res.Name = name;
                wasCreated = true;

                activeSheet.Activate(); // restore the active sheet
            }

            return res;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
