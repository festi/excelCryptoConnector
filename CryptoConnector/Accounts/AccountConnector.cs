using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace CryptoConnector
{
    /// <summary>
    /// model an exchange or a wallet
    /// </summary>
    abstract class AccountConnector
    {
        // name use to match accounts with excel sheets (need to be unique)
        public abstract string Name { get; }

        // signal which functions are supported
        public abstract bool SupportBalance { get; }
        public abstract bool SupportBalanceHistory { get; }
        public abstract bool SupportFills { get; }

        // functions to implement by each implementation
        protected abstract List<AccountId> RefreshBalance_Internal(Worksheet sheet);
        protected abstract void RefreshBalanceHistory_Internal(Worksheet sheet, AccountId id);
        protected abstract void RefreshFills_Internal(Worksheet sheet);

        // get the name of excel sheet of the account
        // the name of the account is at the end because it can be truncated
        public string BalanceSheetName { get { return LimitLengthForExcel($"balance_{Name}"); } }
        public string BalanceHistorySheetName(string symbol) { return LimitLengthForExcel($"balance_hist_{symbol.ToUpper()}_{Name}"); }
        public string FillsSheetName { get { return LimitLengthForExcel($"fills_{Name}"); } }

        private string LimitLengthForExcel(string s)
        {
            if (s.Length > 31) return s.Substring(0, 31);
            return s;
        }

        /// <summary>
        /// a collection of AccountId is returned when balance is queried
        /// accounts can extend this class to add specific infos needed
        /// </summary>
        protected class AccountId{
            public string currency;
        }

        /// <summary>
        /// look for the first empty line and give the max value of a certain colunm
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="column">name of the column where the we look for the max value (or -1 if the table is empty)</param>
        /// <param name="firstEmptyLine">index of the first empty line</param>
        /// <param name="highestValue">highest value found in column 'column' (or -1 if the table is empty)</param>
        protected void LastLineLookup(Worksheet sheet, string column, out int firstEmptyLine, out long highestValue)
        {
            highestValue = -1;
            firstEmptyLine = 2;

            while (sheet.Range[column + firstEmptyLine].Value2 != null)
            {
                highestValue = Math.Max(highestValue, Convert.ToInt64(sheet.Range[column + firstEmptyLine].Value2));
                firstEmptyLine++;
            }
        }

        /// <summary>
        /// look for the first empty line and give the set of a value in a specific column
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="column">name of the column where to pull values</param>
        /// <param name="firstEmptyLine">index of the first empty line</param>
        /// <param name="values">list of values found in column 'column'</param>
        protected void LastLineLookupList(Worksheet sheet, string column, out int firstEmptyLine, out List<string> values)
        {
            firstEmptyLine = 2;
            values = new List<string>();

            while (sheet.Range[column + firstEmptyLine].Value2 != null)
            {
                values.Add(sheet.Range[column + firstEmptyLine].Value2);
                firstEmptyLine++;
            }
        }

        public void Refresh()
        {
            List<AccountId> accounts = null;
            if (AccountsSheet.SyncBalance && SupportBalance)
            {
                accounts = RefreshBalance();
            }
            if (AccountsSheet.SyncBalanceHistory && accounts != null && SupportBalanceHistory)
            {
                foreach (var a in accounts) RefreshAccountHistory(a);
            }
            if (AccountsSheet.SyncBalanceFills && SupportFills)
            {
                RefreshFills();
            }
        }

        protected string ParseSymbol(string symbol)
        {
            return symbol.ToUpper();
        }

        private void SetupTable(Worksheet sheet)
        {
            //sheet.Activate();

            // Fix first row
            //sheet.Application.ActiveWindow.SplitRow = 1;
            //sheet.Application.ActiveWindow.FreezePanes = true;

            // and then format it as a table

            // define points for selecting a range
            // point 1 is the top, leftmost cell
            Range oRng1 = sheet.Range["A1"];
            // point two is the bottom, rightmost cell
            Range oRng2 = sheet.Range["A1"]
                //.End[XlDirection.xlDown]
                .End[XlDirection.xlToRight];

            // define the actual range we want to select
            var range = sheet.Range[oRng1, oRng2];
            //range.Select(); // and select it

            // add the range to a formatted table
            if (range.Worksheet.ListObjects.Count == 0) {
                var table = range.Worksheet.ListObjects.AddEx(
                    SourceType: XlListObjectSourceType.xlSrcRange,
                    Source: range,
                    XlListObjectHasHeaders: XlYesNoGuess.xlYes);
                table.Name = sheet.Name;
            }
        }

        private void RefreshFills()
        {
            bool wasCreated;
            Worksheet sheet = Globals.ThisAddIn.FindWorksheet(FillsSheetName, true, out wasCreated);

            if (wasCreated) sheet.Visible = XlSheetVisibility.xlSheetHidden;
            
            sheet.Range["A1"].Value = "id";
            sheet.Range["B1"].Value = "date";

            //sheet.Range["C1", "E1"].Merge();
            sheet.Range["C1"].Value = "from";
            sheet.Range["D1"].Value = "from currency";

            //sheet.Range["E1", "H1"].Merge();
            sheet.Range["E1"].Value = "to";
            sheet.Range["F1"].Value = "to currency";

            //sheet.Range["G1", "K1"].Merge();
            sheet.Range["G1"].Value = "fee";
            sheet.Range["H1"].Value = "fee currency";

            SetupTable(sheet);

            RefreshFills_Internal(sheet);
        }

        private List<AccountId> RefreshBalance()
        {
            bool wasCreated;
            Worksheet sheet = Globals.ThisAddIn.FindWorksheet(BalanceSheetName, true, out wasCreated);

            if (wasCreated) sheet.Visible = XlSheetVisibility.xlSheetHidden;

            sheet.Range["A1"].Value = "currency";
            sheet.Range["B1"].Value = "balance";
            sheet.Range["C1"].Value = "available";
            sheet.Range["D1"].Value = "holds";

            SetupTable(sheet);

            return RefreshBalance_Internal(sheet);
        }

        private void RefreshAccountHistory(AccountId id)
        {
            bool wasCreated;
            Worksheet sheet = Globals.ThisAddIn.FindWorksheet(BalanceHistorySheetName(id.currency), true, out wasCreated);

            if (wasCreated) sheet.Visible = XlSheetVisibility.xlSheetHidden;

            sheet.Range["A1"].Value = "id";
            sheet.Range["B1"].Value = "date";
            sheet.Range["C1"].Value = "type";
            sheet.Range["D1"].Value = "amount";
            sheet.Range["E1"].Value = "balance";
            sheet.Range["F1"].Value = "currency";
            sheet.Range["G1"].Value = "description";

            SetupTable(sheet);

            /*
             * type is : transfer, match, fee or rebate
             */

            RefreshBalanceHistory_Internal(sheet, id);
        }

        //https://stackoverflow.com/questions/311165/how-do-you-convert-a-byte-array-to-a-hexadecimal-string-and-vice-versa
        public static byte[] StringToByteArray(String hex)
        {
            int NumberChars = hex.Length;
            byte[] bytes = new byte[NumberChars / 2];
            for (int i = 0; i < NumberChars; i += 2)
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            return bytes;
        }
        public static string ByteArrayToString(byte[] ba)
        {
            string hex = BitConverter.ToString(ba);
            return hex.Replace("-", "");
        }

        protected T Request<T>(string apiAddress, string requestPath) where T : class
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(apiAddress);

                // Add an Accept header for JSON format.
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "C# bot");

                // List data response.
                HttpResponseMessage response = client.GetAsync(requestPath).Result;  // Blocking call!
                if (response.IsSuccessStatusCode)
                {
                    // Parse the response body. Blocking!
                    string result = response.Content.ReadAsStringAsync().Result;
                    return JsonConvert.DeserializeObject<T>(result);
                }
                else
                {
                    throw new Exception(string.Format("{0} ({1} {2})", (int)response.StatusCode, response.ReasonPhrase, response.Content.ReadAsStringAsync().Result));
                }
            }
            //return null;
        }
    }
}
