using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace CryptoConnector
{
    class GdaxConnector : AccountConnector
    {
        public const string API_ADDRESS = "https://api.gdax.com";

        public override string UniqueName => $"gdax_{Key}";

        public override bool SupportBalance => true;
        public override bool SupportBalanceHistory => true;
        public override bool SupportFills => true;

        private string Passphrase, Key, Secret;

        protected class GdaxAccountId : AccountId{
            public string account_id;
        }

        public GdaxConnector() : base("") { }

        public GdaxConnector(string name, string _Passphrase, string _Key, string _Secret) : base(name)
        {
            Passphrase = _Passphrase;
            Key = _Key;
            Secret = _Secret;
        }

        protected override void RefreshFills_Internal(Worksheet sheet)
        {
            // look for the first empty line to start filling
            int line; long before;
            LastLineLookup(sheet, "A", out line, out before);

            var fills = RequestSecretPaginated<Fill>("/fills", before, x => x.trade_id);
          
            //fill the sheet with the response
            foreach (var f in fills) 
            {
                var currs = f.product_id.Split('-');
                string from, to, fee = currs[1];
                double fromAmount, toAmount;
                if (f.side == "buy")
                {
                    to = currs[0];
                    toAmount = f.size;
                    from = currs[1];
                    fromAmount = f.size * f.price;
                }
                else
                {
                    from = currs[0];
                    fromAmount = f.size;
                    to = currs[1];
                    toAmount = f.size * f.price;
                }

                sheet.Range["A" + line].Value = f.trade_id;
                sheet.Range["B" + line].Value = DateTime.Parse(f.created_at);

                sheet.Range["C" + line].Value = fromAmount;
                sheet.Range["D" + line].Value = ParseSymbol(from);

                sheet.Range["E" + line].Value = toAmount;
                sheet.Range["F" + line].Value = ParseSymbol(to);

                sheet.Range["G" + line].Value = f.fee;
                sheet.Range["H" + line].Value = ParseSymbol(fee);

                line++;
            }
        }

        protected override List<AccountId> RefreshBalance_Internal(Worksheet sheet)
        {
            List<AccountId> res = new List<AccountId>();

            var accounts = RequestSecret<List<Account>>("/accounts");

            int line = 2;
            foreach (var a in accounts)
            {
                sheet.Range["A" + line].Value = ParseSymbol(a.currency);
                sheet.Range["B" + line].Value = a.balance;
                sheet.Range["C" + line].Value = a.available;
                sheet.Range["D" + line].Value = a.holds;

                res.Add(new GdaxAccountId { account_id = a.id, currency = a.currency });

                line++;
            }

            return res;
        }

        protected override void RefreshBalanceHistory_Internal(Worksheet sheet, AccountId id_)
        {
            GdaxAccountId id = id_ as GdaxAccountId;

            // look for the first empty line to start filling
            int line; long before;
            LastLineLookup(sheet, "A", out line, out before);

            var accounts = RequestSecretPaginated<AccountHistory>("/accounts/" + id.account_id + "/ledger", before, x => x.id);

            foreach (var a in accounts)
            {
                sheet.Range["A" + line].Value = a.id;
                sheet.Range["B" + line].Value = DateTime.Parse(a.created_at);
                sheet.Range["C" + line].Value = a.type;
                sheet.Range["D" + line].Value = a.amount;
                sheet.Range["E" + line].Value = a.balance;
                sheet.Range["F" + line].Value = ParseSymbol(id.currency);

                line++;
            }
        }

        //public Worksheet RefreshDeposits(string account_id, string currency)
        //{
        //    var accounts = RequestSecret<List<AccountHistory>>("/accounts/" + account_id + "/ledger");
        //
        //    Worksheet sheet = Globals.ThisAddIn.FindWorksheet("gdax deposits " + currency, true);
        //    sheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
        //
        //    sheet.Range["A1"].Value = "id";
        //    sheet.Range["B1"].Value = "date";
        //    sheet.Range["C1"].Value = "amount";
        //    sheet.Range["C1"].Value = "currency";
        //
        //    int line = 2;
        //    foreach (var a in accounts)
        //    {
        //        if(a.type == "transfer")
        //        {
        //            sheet.Range["A" + line].Value = a.id;
        //            sheet.Range["B" + line].Value = DateTime.Parse(a.created_at);
        //            sheet.Range["C" + line].Value = a.amount;
        //            sheet.Range["D" + line].Value = currency;
        //
        //            line++;
        //        }
        //    }
        //
        //    return sheet;
        //}
        //
        //public void MergeDeposits(List<Worksheet> src)
        //{
        //    Worksheet sheet = Globals.ThisAddIn.FindWorksheet("gdax deposits", true);
        //
        //    sheet.Range["A1"].Value = "date";
        //    sheet.Range["B1"].Value = "amount";
        //    sheet.Range["C1"].Value = "currency";
        //
        //    // Fix first row
        //    sheet.Activate();
        //    sheet.Application.ActiveWindow.SplitRow = 1;
        //    sheet.Application.ActiveWindow.FreezePanes = true;
        //    // Now apply autofilter
        //    Range firstRow = (Range)sheet.Rows[1];
        //    firstRow.AutoFilter(1,
        //                        Type.Missing,
        //                        XlAutoFilterOperator.xlAnd,
        //                        Type.Missing,
        //                        true);
        //
        //    int next = 2;
        //
        //    foreach(var w in src)
        //    {
        //        int lineEnd = 2;
        //        while (w.Range["A" + lineEnd].Value2 != null) lineEnd++;
        //
        //        if(lineEnd > 2)
        //        {
        //            w.Range["B2", "D" + (lineEnd - 1)].Copy(sheet.Range["A" + next]);
        //            next += lineEnd - 2;
        //        }
        //    }
        //}

        private IEnumerable<T> RequestSecretPaginated<T>(string requestPath, long before, Func<T,long> id) where T : class
        {
            HttpHeaders responseHeaders;
            string beforeTxt = before > -1 ? "?before=" + before : "";

            // get last page

            System.Diagnostics.Debug.Print("first request");
            List<T> res = RequestSecret<List<T>>(requestPath + beforeTxt, out responseHeaders);
            string cb_after_str = responseHeaders.Contains("CB-AFTER") ? responseHeaders.GetValues("CB-AFTER").FirstOrDefault() : null;
            long cb_after = cb_after_str != null ? long.Parse(cb_after_str) : -1;
            if(responseHeaders.Contains("CB-BEFORE")) System.Diagnostics.Debug.Print(responseHeaders.GetValues("CB-BEFORE").FirstOrDefault());

            while (cb_after > before+1)
            {
                System.Diagnostics.Debug.Print("additional request");
                var nextRange = RequestSecret<List<T>>(string.Format("{0}?after={1}", requestPath, cb_after), out responseHeaders);
                res.AddRange(nextRange.Where(x => id(x) > before));

                cb_after_str = responseHeaders.Contains("CB-AFTER") ? responseHeaders.GetValues("CB-AFTER").FirstOrDefault() : null;
                cb_after = cb_after_str != null ? long.Parse(cb_after_str) : -1;
                if (responseHeaders.Contains("CB-BEFORE")) System.Diagnostics.Debug.Print(responseHeaders.GetValues("CB-BEFORE").FirstOrDefault());
            }

            // reverse order to sort by assending id
            return res.Reverse<T>();
        }

        private T RequestSecret<T>(string requestPath) where T : class
        {
            HttpHeaders responseHeaders;
            return RequestSecret<T>(requestPath, out responseHeaders);
        }

        private T RequestSecret<T>(string requestPath, out HttpHeaders responseHeaders) where T : class
        {
            long timestamp = (long)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalSeconds;
            string what = timestamp + "GET" + requestPath;
            var decodedKey = Convert.FromBase64String(Secret);
            HMACSHA256 hmac = new HMACSHA256(decodedKey);
            string sign = Convert.ToBase64String(hmac.ComputeHash(Encoding.UTF8.GetBytes(what)));

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(API_ADDRESS);

                // Add an Accept header for JSON format.
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "C# bot");
                client.DefaultRequestHeaders.TryAddWithoutValidation("CB-ACCESS-KEY", Key);
                client.DefaultRequestHeaders.TryAddWithoutValidation("CB-ACCESS-SIGN", sign);
                client.DefaultRequestHeaders.TryAddWithoutValidation("CB-ACCESS-TIMESTAMP", timestamp.ToString());
                client.DefaultRequestHeaders.TryAddWithoutValidation("CB-ACCESS-PASSPHRASE", Passphrase);

                // List data response.
                HttpResponseMessage response = client.GetAsync(requestPath).Result;  // Blocking call!
                if (response.IsSuccessStatusCode)
                {
                    // Parse the response body. Blocking!
                    string result = response.Content.ReadAsStringAsync().Result;
                    responseHeaders = response.Headers;
                    return JsonConvert.DeserializeObject<T>(result);
                }
                else
                {
                    throw new Exception(string.Format("{0} ({1} {2})", (int)response.StatusCode, response.ReasonPhrase, response.Content.ReadAsStringAsync().Result));
                }
            }
            //return null;
        }

        private class Fill
        {
            public long trade_id;
            public string product_id;
            public double price;
            public double size;
            public string order_id;
            public string created_at;
            public string liquidity;
            public double fee;
            public bool settled;
            public string side;
        }

        private class Account
        {
            public string id;
            public string currency;
            public double balance;
            public double available;
            public double holds;
            public string profile_id;
        }

        private class AccountHistory
        {
            public long id;
            public string created_at;
            public double amount;
            public double balance;
            public string type;
        }
    }
}
