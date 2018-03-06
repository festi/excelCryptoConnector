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
    class CryptopiaConnector : AccountConnector
    {
        public const string API_ADDRESS = "https://www.cryptopia.co.nz";

        public override string UniqueName => $"cryptopia_{Key}";

        public override bool SupportBalance => true;
        public override bool SupportBalanceHistory => false;
        public override bool SupportFills => false;

        private string Key, Secret;

        public CryptopiaConnector() : base("") { }

        public CryptopiaConnector(string name, string _Key, string _Secret) : base(name)
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

            var account = RequestSecret<Balance[]>("/api/GetBalance");

            int line = 2;
            foreach (var b in account.Data)
            {
                sheet.Range["A" + line].Value = ParseSymbol(b.Symbol);
                sheet.Range["B" + line].Value = b.Available + b.HeldForTrades;
                sheet.Range["C" + line].Value = b.Available;
                sheet.Range["D" + line].Value = b.HeldForTrades;

                res.Add(new AccountId { currency = b.Symbol });

                line++;
            }

            return res;
        }

        protected override void RefreshFills_Internal(Worksheet sheet)
        {
            throw new NotImplementedException();
        }


        private Result<T> RequestSecret<T>(string requestPath, Dictionary<string, string> param = null) where T : class
        {

            string post_data = "";
            if(param != null) foreach (var pair in param)
            {
                if (post_data != "") post_data += ",";
                post_data += string.Format("\"{0}\":\"{1}\"", pair.Key, pair.Value);
            }
            post_data = "{" + post_data + "}";


            string nonce = ((long)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalMilliseconds).ToString();
            string url = Uri.EscapeDataString($"{API_ADDRESS}{requestPath}").ToLower();

            string requestContentBase64String = Convert.ToBase64String(MD5.Create().ComputeHash(Encoding.UTF8.GetBytes(post_data)));
            string signature = $"{Key}POST{url}{nonce}{requestContentBase64String}";

            var decodedKey = Convert.FromBase64String(Secret);
            HMACSHA256 hmac = new HMACSHA256(decodedKey);
            string hmacsignature = Convert.ToBase64String(hmac.ComputeHash(Encoding.UTF8.GetBytes(signature)));

            string header_value = $"amx {Key}:{hmacsignature}:{nonce}";
            

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(API_ADDRESS);

                // Add an Accept header for JSON format.
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "C# bot");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", header_value);

                // List data response.
                HttpContent content = new StringContent(post_data, Encoding.UTF8, "application/json");
                HttpResponseMessage response = client.PostAsync(requestPath, content).Result;  // Blocking call!
                if (response.IsSuccessStatusCode)
                {
                    // Parse the response body. Blocking!
                    string result = response.Content.ReadAsStringAsync().Result;
                    return JsonConvert.DeserializeObject<Result<T>>(result);
                }
                else
                {
                    throw new Exception(string.Format("{0} ({1} {2})", (int)response.StatusCode, response.ReasonPhrase, response.Content.ReadAsStringAsync().Result));
                }
            }
            //return null;
        }

        private class Result<T>{
            public string Success;
            public string Error;
            public T Data;
        }

        private class Balance
        {
            public long CurrencyId;
            public string Symbol;
            public double Total;
            public double Available;
            public double Unconfirmed;
            public double HeldForTrades;
            public double PendingWithdraw;
            public string Address;
            public string BaseAddress;
            public string Status;
            public string StatusMessage;
        }
}
}
