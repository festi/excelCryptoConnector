using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;
using System.Net.Http.Headers;
using System.Net.Http;
using Newtonsoft.Json;
using System.Threading;

namespace CryptoConnector
{
    class BitFinexConnector : AccountConnector
    {
        public const string API_ADDRESS = "https://api.bitfinex.com";

        public override string UniqueName => $"bitfinex_{Key}";

        public override bool SupportBalance => true;
        public override bool SupportBalanceHistory => true;
        public override bool SupportFills => true;

        private string Key, Secret;

        public BitFinexConnector() : base("") { }

        public BitFinexConnector(string name, string _Key, string _Secret) : base(name)
        {
            Key = _Key;
            Secret = _Secret;
        }

        protected override void RefreshBalanceHistory_Internal(Worksheet sheet, AccountId id)
        {
            // look for the first empty line to start filling
            int line; long since;
            LastLineLookup(sheet, "A", out line, out since);
            
            Dictionary<string, string> param = new Dictionary<string, string>();
            param.Add("currency", id.currency);
            if (since > -1) param.Add("since", (since+1).ToString());
            var balanceHist = RequestSecret<List<BalanceHistory>>("/v1/history", param);

            foreach (var a in balanceHist.Reverse<BalanceHistory>())
            {
                long timestamp = long.Parse(a.timestamp.Split('.')[0]);

                string type = "ERROR";
                if (a.description.StartsWith("Deposit")) type = "transfer";
                if (a.description.StartsWith("Trading fees")) type = "fee";
                if (a.description.StartsWith("Exchange")) type = "match";

                sheet.Range["A" + line].Value = timestamp;
                sheet.Range["B" + line].Value = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc).AddSeconds(timestamp);
                sheet.Range["C" + line].Value = type;
                sheet.Range["D" + line].Value = a.amount;
                sheet.Range["E" + line].Value = a.balance;
                sheet.Range["F" + line].Value = ParseSymbol(id.currency);
                sheet.Range["G" + line].Value = a.description;
                //sheet.Range["H" + line].Value = "a"+a.timestamp;

                line++;
            }
        }

        protected override List<AccountId> RefreshBalance_Internal(Worksheet sheet)
        {
            List<AccountId> res = new List<AccountId>();

            var balances = RequestSecret<List<Balance>>("/v1/balances");

            int line = 2;
            foreach (var b in balances)
            {
                sheet.Range["A" + line].Value = ParseSymbol(b.currency);
                sheet.Range["B" + line].Value = b.amount;
                sheet.Range["C" + line].Value = b.available;
                sheet.Range["D" + line].Formula = string.Format("= B{0} - C{0}", line);

                res.Add(new AccountId { currency = b.currency });

                line++;
            }

            return res;
        }

        protected override void RefreshFills_Internal(Worksheet sheet)
        {
            // look for the first empty line to start filling
            int line; long since;
            LastLineLookup(sheet, "A", out line, out since);

            var symbols = Request<List<string>>(API_ADDRESS, "/v1/symbols");

            Dictionary <string, string> param = new Dictionary<string, string>();
            if (since > -1) param.Add("timestamp", (since + 1).ToString());
            param.Add("symbol", "");

            foreach(var symbol in symbols)
            {
                param["symbol"] = symbol;
                var trades = RequestSecret<List<Trade>>("/v1/mytrades", param);

                var currs0 = symbol.Substring(0, 3);
                var currs1 = symbol.Substring(3, 3);

                foreach (var trade in trades)
                {
                    long timestamp = long.Parse(trade.timestamp.Split('.')[0]);


                    string from, to;
                    double fromAmount, toAmount;
                    if (trade.type == "Buy")
                    {
                        to = currs0;
                        toAmount = trade.amount;
                        from = currs1;
                        fromAmount = trade.amount * trade.price;
                    }
                    else
                    {
                        from = currs0;
                        fromAmount = trade.amount;
                        to = currs1;
                        toAmount = trade.amount * trade.price;
                    }

                    sheet.Range["A" + line].Value = timestamp;
                    sheet.Range["B" + line].Value = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc).AddSeconds(timestamp);

                    sheet.Range["C" + line].Value = fromAmount;
                    sheet.Range["D" + line].Value = ParseSymbol(from);

                    sheet.Range["E" + line].Value = toAmount;
                    sheet.Range["F" + line].Value = ParseSymbol(to);

                    sheet.Range["G" + line].Value = trade.fee_amount;
                    sheet.Range["H" + line].Value = ParseSymbol(trade.fee_currency);

                    line++;
                }
            }
        }

        private T RequestSecret<T>(string apiPath, Dictionary<string, string> param = null)
        {
            string nonce = ((long)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalMilliseconds).ToString();

            Dictionary<string, string> authPairs = new Dictionary<string, string>();
            authPairs.Add("request", apiPath);
            authPairs.Add("nonce", nonce);

            string body = "";
            foreach(var pair in param != null ? authPairs.Concat(param) : authPairs)
            {
                if (body != "") body += ",";
                body += string.Format("\"{0}\":\"{1}\"", pair.Key, pair.Value);
            }
            body = "{" + body + "}";

            //string body = string.Format("{{\"request\":\"{0}\",\"nonce\":\"{1}\"{2}}}", apiPath, nonce, extra);
            string payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(body));

            HMACSHA384 hmac = new HMACSHA384(Encoding.UTF8.GetBytes(Secret));
            var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(payload));
            string signature = BitConverter.ToString(hash).Replace("-", string.Empty).ToLower();

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(API_ADDRESS);

                // Add an Accept header for JSON format.
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "C# bot");
                client.DefaultRequestHeaders.TryAddWithoutValidation("X-BFX-APIKEY", Key);
                client.DefaultRequestHeaders.TryAddWithoutValidation("X-BFX-PAYLOAD", payload);
                client.DefaultRequestHeaders.TryAddWithoutValidation("X-BFX-SIGNATURE", signature);

                // List data response.
                HttpContent content = new StringContent(body, Encoding.UTF8, "application/json");
                HttpResponseMessage response = client.PostAsync(apiPath, content).Result;  // Blocking call!
                if (response.IsSuccessStatusCode)
                {
                    // Parse the response body. Blocking!
                    string result = response.Content.ReadAsStringAsync().Result;
                    return JsonConvert.DeserializeObject<T>(result);
                }
                //else if (response.ReasonPhrase == "Too Many Requests")
                //{
                //    Thread.Sleep(1000);
                //    return RequestSecret<T>(apiPath, param);
                //}
                else
                {
                    throw new Exception(string.Format("{0} ({1} {2})", (int)response.StatusCode, response.ReasonPhrase, response.Content.ReadAsStringAsync().Result));
                }
            }
        }

        private class Balance
        {
            public string type;
            public string currency;
            public double amount;
            public double available;
        }

        private class BalanceHistory
        {
            public string currency;
            public double amount;
            public double balance;
            public string description;
            public string timestamp;
        }

        private class Trade
        {
            public double price;
            public double amount;
            public string timestamp;
            public string exchange;
            public string type;
            public string fee_currency;
            public double fee_amount;
            public long tid;
            public long order_id;
        }
    }
}
