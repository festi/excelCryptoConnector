using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CryptoConnector
{
    class NeoscanConnector : AccountConnector
    {
        public const string API_ADDRESS = "https://neoscan.io";

        public override string UniqueName => $"neo_{PublicKey}";

        public override bool SupportBalance => true;
        public override bool SupportBalanceHistory => false;
        public override bool SupportFills => false;
        
        private string PublicKey;

        public NeoscanConnector() : base("") { }

        public NeoscanConnector(string name, string _PublicKey) : base(name)
        {
            PublicKey = _PublicKey;
        }

        protected override List<AccountId> RefreshBalance_Internal(Worksheet sheet)
        {
            List<AccountId> res = new List<AccountId>();

            // list all the balances
            var accounts = Request<Balance>(API_ADDRESS, $"/api/main_net/v1/get_balance/{PublicKey}");

            // and the claimable amount of GAS
            var claimlable = Request<Claimable>(API_ADDRESS, $"/api/main_net/v1/get_claimable/{PublicKey}");

            ExecuteExcelJobSync(delegate ()
            {
                int line = 2;
                foreach (var a in accounts.balance)
                {
                    // amount of clamaible of this currency
                    double this_claimable = a.asset == "GAS" ? claimlable.unclaimed : 0.0d;

                    // we put the claimable balance on hold
                    sheet.Range["A" + line].Value = ParseSymbol(a.asset);
                    sheet.Range["B" + line].Value = a.amount + this_claimable;
                    sheet.Range["C" + line].Value = a.amount;
                    sheet.Range["D" + line].Value = this_claimable;

                    res.Add(new AccountId { currency = a.asset });

                    line++;
                }
            });

            return res;
        }

        protected override void RefreshBalanceHistory_Internal(Worksheet sheet, AccountId id)
        {
            throw new NotImplementedException();
        }

        protected override void RefreshFills_Internal(Worksheet sheet)
        {
            throw new NotImplementedException();
        }

        private class Balance
        {
            public IndividualBalance[] balance;
            public string address;
        }

        private class IndividualBalance
        {
            public string asset;
            public double amount;
        }

        private class Claimable
        {
            public double unclaimed;
            public string address;
        }
    }
}
