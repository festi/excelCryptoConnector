using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace CryptoConnector
{
    class BittrexConnector : AccountConnector
    {
        public const string API_ADDRESS = "https://bittrex.com";

        public override string UniqueName => $"bittrex_{Key}";

        public override bool SupportBalance => true;
        public override bool SupportBalanceHistory => false;
        public override bool SupportFills => false;

        private string Key, Secret;

        public BittrexConnector() : base("") { }

        public BittrexConnector(string name, string _Key, string _Secret) : base(name)
        {
            Key = _Key;
            Secret = _Secret;
        }

        protected override void RefreshBalanceHistory_Internal(Worksheet sheet, AccountId id)
        {
            throw new NotImplementedException();
        }

        protected override List<AccountId> RefreshBalance_Internal(Worksheet sheet)
        {
            List<AccountId> res = new List<AccountId>();

            var balances = RequestSecret<Response<List<GetBalance>>>("/api/v1.1/account/getbalances","");

            int line = 2;
            foreach (var b in balances.result)
            {
                sheet.Range["A" + line].Value = ParseSymbol(b.Currency);
                sheet.Range["B" + line].Value = b.Balance;
                sheet.Range["C" + line].Value = b.Available;
                sheet.Range["D" + line].Formula = string.Format("= B{0} - C{0}", line);

                res.Add(new AccountId { currency = b.Currency });

                line++;
            }

            return res;
        }

        protected override void RefreshFills_Internal(Worksheet sheet)
        {
            throw new NotImplementedException();

            // commented for new, the api only give the last 30 days

            //// look for the first empty line to start filling
            //int line; List<string> before;
            //LastLineLookupList(sheet, "A", out line, out before);
            //
            //var orders = RequestSecret<Response<List<Order>>>("/api/v1.1/account/getorderhistory", "");
            //
            //foreach (var f in orders.result)
            //{
            //    if (before.Contains(f.OrderUuid)) continue; // skip existing entries
            //
            //    var currs = f.Exchange.Split('-');
            //    string from, to, fee = currs[1];
            //    double fromAmount, toAmount;
            //    if (f.OrderType == "LIMIT_BUY")
            //    {
            //        to = currs[1];
            //        toAmount = f.Quantity;
            //        from = currs[0];
            //        fromAmount = f.Quantity * f.Price;
            //    }
            //    else if (f.OrderType == "LIMIT_SELL")
            //    {
            //        from = currs[1];
            //        fromAmount = f.Quantity;
            //        to = currs[0];
            //        toAmount = f.Quantity * f.Price;
            //    }
            //    else
            //    {
            //        throw new Exception("unknown type of order type");
            //    }
            //
            //    sheet.Range["A" + line].Value = f.OrderUuid;
            //    sheet.Range["B" + line].Value = DateTime.Parse(f.TimeStamp);
            //
            //    sheet.Range["C" + line].Value = fromAmount;
            //    sheet.Range["D" + line].Value = ParseSymbol(from);
            //
            //    sheet.Range["E" + line].Value = toAmount;
            //    sheet.Range["F" + line].Value = ParseSymbol(to);
            //
            //    sheet.Range["G" + line].Value = f.Commission;
            //    sheet.Range["H" + line].Value = ParseSymbol(fee);
            //
            //    line++;
            //}
        }

        private T RequestSecret<T>(string requestPath, string parameters) where T : class
        {
            string nonce = ((long)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalMilliseconds).ToString();
            string fullRequestPath = $"{requestPath}?apikey={Key}&nonce={nonce}&{parameters}";
            string uri = API_ADDRESS + fullRequestPath;
            var decodedSecret = Encoding.UTF8.GetBytes(Secret);
            HMACSHA512 hmac = new HMACSHA512(decodedSecret);
            string sign = ByteArrayToString(hmac.ComputeHash(Encoding.UTF8.GetBytes(uri)));

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(API_ADDRESS);

                // Add an Accept header for JSON format.
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "C# bot");
                client.DefaultRequestHeaders.TryAddWithoutValidation("apisign", sign);

                // List data response.
                HttpResponseMessage response = client.GetAsync(fullRequestPath).Result;  // Blocking call!
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

        private class Response<T>
        {
            public bool success;
            public string message;
            public T result;
        }

        private class GetBalance
        {
            public string Currency;
            public double Balance;
            public double Available;
            public double Pending;
            public string CryptoAddress;
            public bool Requested;
            public string Uuid;
        }

        private class Order
        {
            public string OrderUuid;
            public string Exchange;
            public string TimeStamp;
            public string OrderType;
            public double Limit;
            public double Quantity;
            public double QuantityRemaining;
            public double Commission;
            public double Price;
            //public string PricePerUnit;
            //public string IsConditional;
            //public string Condition;
            //public string ConditionTarget;
            //public string ImmediateOrCancel;
        }

    }
}
