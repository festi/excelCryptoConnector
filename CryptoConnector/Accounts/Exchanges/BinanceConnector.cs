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
    class BinanceConnector : AccountConnector
    {
        public const string API_ADDRESS = "https://api.binance.com";

        public override string UniqueName => $"binance_{Key}";

        public override bool SupportBalance => true;
        public override bool SupportBalanceHistory => false;
        public override bool SupportFills => false;

        private string Key, Secret;

        public BinanceConnector() : base("") { }

        public BinanceConnector(string name, string _Key, string _Secret) : base(name)
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

            var account = RequestSecret<Account>("/api/v3/account");

            ExecuteExcelJobSync(delegate()
            {
                int line = 2;
                foreach (var b in account.balances)
                {
                    sheet.Range["A" + line].Value = ParseSymbol(b.asset);
                    sheet.Range["B" + line].Value = b.free + b.locked;
                    sheet.Range["C" + line].Value = b.free;
                    sheet.Range["D" + line].Value = b.locked;

                    res.Add(new AccountId { currency = b.asset });

                    line++;
                }
            });
            
            return res;
        }

        protected override void RefreshFills_Internal(Worksheet sheet)
        {
            throw new NotImplementedException();
        }

        private T RequestSecret<T>(string requestPath, string param = "") where T : class
        {
            //long timestamp = (long)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalMilliseconds;
            // take timestamp from binance because it need te be withing a small window
            // and the local computer time may be out of sync
            long timestamp = Request<Time>(API_ADDRESS, "/api/v1/time").serverTime;

            // add timestamp before signing
            if (param != "") param += "&";
            param += $"timestamp={timestamp}";

            var decodedKey = Encoding.UTF8.GetBytes(Secret);
            HMACSHA256 hmac = new HMACSHA256(decodedKey);
            string signature = ByteArrayToString(hmac.ComputeHash(Encoding.UTF8.GetBytes(param)));
            
            // and add signature do the request
            param += $"&signature={signature}";

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(API_ADDRESS);

                // Add an Accept header for JSON format.
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "C# bot");
                client.DefaultRequestHeaders.TryAddWithoutValidation("X-MBX-APIKEY", Key);

                // List data response.
                HttpResponseMessage response = client.GetAsync($"{requestPath}?{param}").Result;  // Blocking call!
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

        private class Account
        {
            public double makerCommission;
            public double takerCommission;
            public double buyerCommission;
            public double sellerCommission;
            public string canTrade;
            public string canWithdraw;
            public string canDeposit;
            public Balance[] balances;
        }

        private class Balance
        {
            public string asset;
            public double free;
            public double locked;
        }

        private class Time
        {
            public long serverTime;
        }
    }
}
