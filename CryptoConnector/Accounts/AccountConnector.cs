using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CryptoConnector
{
    /// <summary>
    /// model an exchange or a wallet
    /// </summary>
    abstract class AccountConnector
    {
        // name use to match accounts with excel sheets (need to be unique)
        public abstract string UniqueName { get; }
        // name displayed in the tables (does not need to be unique)
        public string ReadableName { get; private set; }

        // signal which functions are supported
        public abstract bool SupportBalance { get; }
        public abstract bool SupportBalanceHistory { get; }
        public abstract bool SupportFills { get; }

        // if true, the balance sheet contains a 'account' colomn describing from which account belong each line
        public virtual bool BalanceContainsAccountName { get { return false; } }

        // functions to implement by each implementation
        protected abstract List<AccountId> RefreshBalance_Internal(Worksheet sheet);
        protected abstract void RefreshBalanceHistory_Internal(Worksheet sheet, AccountId id);
        protected abstract void RefreshFills_Internal(Worksheet sheet);

        // get the name of excel sheet of the account
        // the name of the account is at the end because it can be truncated
        public string BalanceSheetName { get { return LimitLengthForExcel($"balance_{UniqueName}"); } }
        public string BalanceHistorySheetName(string symbol) { return LimitLengthForExcel($"balance_hist_{symbol.ToUpper()}_{UniqueName}"); }
        public string FillsSheetName { get { return LimitLengthForExcel($"fills_{UniqueName}"); } }

        // excel API is not thread-safe, we use a job queue to commit action to a single thread that iinteract with excel
        private BlockingCollection<Task> ExcelJobs = new BlockingCollection<Task>();
        private Thread ExcelJobsThread = null;

        public AccountConnector(string readableName)
        {
            ReadableName = readableName;
        }

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
        /// 
        /// this function is thread-safe
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="column">name of the column where the we look for the max value (or -1 if the table is empty)</param>
        /// <param name="firstEmptyLine">index of the first empty line</param>
        /// <param name="highestValue">highest value found in column 'column' (or -1 if the table is empty)</param>
        protected void LastLineLookup(Worksheet sheet, string column, out int firstEmptyLine, out long highestValue)
        {
            int _highestValue = -1;
            int _firstEmptyLine = 2;

            ExecuteExcelJobSync(delegate ()
            {
                while (sheet.Range[column + _firstEmptyLine].Value2 != null)
                {
                    _highestValue = Math.Max(_highestValue, Convert.ToInt64(sheet.Range[column + _firstEmptyLine].Value2));
                    _firstEmptyLine++;
                }
            });

            firstEmptyLine = _firstEmptyLine;
            highestValue = _highestValue;
        }

        /// <summary>
        /// look for the first empty line and give the set of a value in a specific column
        /// 
        /// this function is thread-safe
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="column">name of the column where to pull values</param>
        /// <param name="firstEmptyLine">index of the first empty line</param>
        /// <param name="values">list of values found in column 'column'</param>
        protected void LastLineLookupList(Worksheet sheet, string column, out int firstEmptyLine, out List<string> values)
        {
            int _firstEmptyLine = 2;
            var _values = new List<string>();

            ExecuteExcelJobSync(delegate ()
            {
                while (sheet.Range[column + _firstEmptyLine].Value2 != null)
                {
                    _values.Add(sheet.Range[column + _firstEmptyLine].Value2);
                    _firstEmptyLine++;
                }
            });
            
            firstEmptyLine = _firstEmptyLine;
            values = _values;
        }

        public void Refresh(IAccountManagerEvents listener)
        {
            try
            {
                List<AccountId> accounts = null;
                if (AccountsSheet.SyncBalance && SupportBalance)
                {
                    listener.SyncAccountStatus(UniqueName, "sync balance");
                    accounts = RefreshBalance();
                }
                if (AccountsSheet.SyncBalanceHistory && accounts != null && SupportBalanceHistory)
                {
                    listener.SyncAccountStatus(UniqueName, "sync balance history");
                    foreach (var a in accounts) RefreshAccountHistory(a);
                }
                if (AccountsSheet.SyncBalanceFills && SupportFills)
                {
                    listener.SyncAccountStatus(UniqueName, "sync fills");
                    RefreshFills();
                }
            }
            catch(Exception e)
            {
                listener.SyncAccountStatus(UniqueName, $"Error {e.Message}");
            }


            listener.SyncAccountStatus(UniqueName, "done");
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

            ExecuteExcelJobSync(delegate ()
            {
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
            });

            RefreshFills_Internal(sheet);
        }

        private List<AccountId> RefreshBalance()
        {
            bool wasCreated;
            Worksheet sheet = Globals.ThisAddIn.FindWorksheet(BalanceSheetName, true, out wasCreated);

            ExecuteExcelJobSync(delegate ()
            {
                if (wasCreated) sheet.Visible = XlSheetVisibility.xlSheetHidden;

                sheet.Range["A1"].Value = "currency";
                sheet.Range["B1"].Value = "balance";
                sheet.Range["C1"].Value = "available";
                sheet.Range["D1"].Value = "holds";

                SetupTable(sheet);
            });

            return RefreshBalance_Internal(sheet);
        }

        private void RefreshAccountHistory(AccountId id)
        {
            bool wasCreated;
            Worksheet sheet = Globals.ThisAddIn.FindWorksheet(BalanceHistorySheetName(id.currency), true, out wasCreated);

            ExecuteExcelJobSync(delegate ()
            {
                if (wasCreated) sheet.Visible = XlSheetVisibility.xlSheetHidden;

                sheet.Range["A1"].Value = "id";
                sheet.Range["B1"].Value = "date";
                sheet.Range["C1"].Value = "type";
                sheet.Range["D1"].Value = "amount";
                sheet.Range["E1"].Value = "balance";
                sheet.Range["F1"].Value = "currency";
                sheet.Range["G1"].Value = "description";

                SetupTable(sheet);
            });

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

        /// <summary>
        /// action will be executed in a safe tread to interact with excel
        /// this function return when the task has been executed
        /// </summary>
        /// <param name="action"></param>
        protected void ExecuteExcelJobSync(Task action)
        {
            // commit job
            ExcelJobs.Add(action);

            StartExcelJobsRunner();

            action.Wait();
        }

        protected void ExecuteExcelJobSync(System.Action p)
        {
            ExecuteExcelJobSync(new Task(p));
        }

        /// <summary>
        /// action will be executed in a safe tread to interact with excel
        /// return an task to wait on
        /// </summary>
        /// <param name="action"></param>
        protected Task ExecuteExcelJobAsync(Task action)
        {
            // commit job
            ExcelJobs.Add(action);

            StartExcelJobsRunner();

            return action;
        }

        protected Task ExecuteExcelJobAsync(System.Action p)
        {
            return ExecuteExcelJobAsync(new Task(p));
        }

        /// <summary>
        /// if the excel job runner thread does not run, launch it
        /// </summary>
        private void StartExcelJobsRunner()
        {
            if (ExcelJobsThread == null)
            {
                // see https://msdn.microsoft.com/fr-fr/library/8sesy69e.aspx
                Thread t = new Thread(ExcelJobsRunner);
                t.SetApartmentState(System.Threading.ApartmentState.STA);
                t.Start();
            }
        }

        private void ExcelJobsRunner()
        {
            Task nextAction;
            while (true)
            {
                if(ExcelJobs.TryTake(out nextAction, 10000))
                {
                    nextAction.RunSynchronously();
                }
            }
        }
    }
}
