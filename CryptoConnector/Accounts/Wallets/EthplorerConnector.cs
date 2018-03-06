using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CryptoConnector
{
    class EthplorerConnector : AccountConnector
    {
        public const string API_ADDRESS = "https://api.ethplorer.io";

        public override string UniqueName => $"eth_{PublicKey}";

        public override bool SupportBalance => true;
        public override bool SupportBalanceHistory => false;
        public override bool SupportFills => false;

        private string PublicKey;

        public EthplorerConnector() : base("") { }

        public EthplorerConnector(string name, string _PublicKey) : base(name)
        {
            PublicKey = _PublicKey;
        }

        protected override void RefreshBalanceHistory_Internal(Worksheet sheet, AccountId id)
        {
            throw new NotImplementedException();
        }

        protected override List<AccountId> RefreshBalance_Internal(Worksheet sheet)
        {
            List<AccountId> res = new List<AccountId>();

            var accounts = Request<AdressInfo>(API_ADDRESS, $"/getAddressInfo/{PublicKey}?apiKey=freekey");

            int line = 2;
            // first take care of the eth balance
            {
                var eth = accounts.ETH;

                sheet.Range["A" + line].Value = ParseSymbol("ETH");
                sheet.Range["B" + line].Value = eth.balance;
                sheet.Range["C" + line].Value = eth.balance;
                sheet.Range["D" + line].Value = 0;

                res.Add(new AccountId { currency = "ETH" });

                line++;
            }

            // ther list all ERC20 tokens
            foreach (var a in accounts.tokens)
            {
                sheet.Range["A" + line].Value = ParseSymbol(a.tokenInfo.symbol);
                sheet.Range["B" + line].Value = a.balance / Math.Pow(10, a.tokenInfo.decimals);
                sheet.Range["C" + line].Value = a.balance / Math.Pow(10, a.tokenInfo.decimals);
                sheet.Range["D" + line].Value = 0;

                res.Add(new AccountId { currency = a.tokenInfo.symbol });

                line++;
            }

            return res;
        }

        protected override void RefreshFills_Internal(Worksheet sheet)
        {
            throw new NotImplementedException();
        }
        
        private class Eth
        {
            public double balance;
            public double totalIn;
            public double totalOut;
        }

        private class Token {
            public TokenInfo tokenInfo;
            public double balance;
            public double totalIn;
            public double totalOut;
        }

        private class TokenInfo
        {
            public string name;
            public string symbol;
            public int decimals;
        }

        private class AdressInfo
        {
            public string address;
            public Eth ETH;
            public Token[] tokens;
        }
    }
}
