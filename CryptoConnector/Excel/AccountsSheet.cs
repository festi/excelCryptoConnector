using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CryptoConnector
{
    class AccountsSheet
    {
        private static string ACCOUNT_SETTINGS_WORKSHEET_NAME = "accounts_settings";

        private static readonly int GDAX_ROW = 11;
        private static readonly int GDAX_PASSPHARSE_ROW = 12;
        private static readonly int GDAX_KEY_ROW = 13;
        private static readonly int GDAX_SECRET_ROW = 14;
        
        private static readonly int BITFINEX_ROW = 16;
        private static readonly int BITFINEX_KEY_ROW = 17;
        private static readonly int BITFINEX_SECRET_ROW = 18;

        private static readonly int BITTREX_ROW = 20;
        private static readonly int BITTREX_KEY_ROW = 21;
        private static readonly int BITTREX_SECRET_ROW = 22;

        private static readonly int BINANCE_ROW = 24;
        private static readonly int BINANCE_KEY_ROW = 25;
        private static readonly int BINANCE_SECRET_ROW = 26;

        private static readonly int CRYPTOPIA_ROW = 28;
        private static readonly int CRYPTOPIA_KEY_ROW = 29;
        private static readonly int CRYPTOPIA_SECRET_ROW = 30;
        
        private static readonly int ETHERSCAN_ROW = 32;
        private static readonly int ETHERSCAN_KEY_ROW = 33;

        private static readonly int SYNC_CONFIG = 1;
        private static readonly int SYNC_BALANCE = 2;
        private static readonly int SYNC_BALANCE_HISTORY = 3;
        private static readonly int SYNC_FILLS = 4;

        // index start at 0
        private static string GetParam(int paramRow, int index)
        {
            var range = (Range)AccountSettingsWorksheet.Cells[paramRow, index + 2];
            string res = range.Value;
            if (res == null) res = "";
            return res;
        }

        private static string SupportText(AccountConnector account)
        {
            string res = "supported operations :";
            if (account.SupportBalance) res += " balance";
            if (account.SupportBalanceHistory) res += " balance_history";
            if (account.SupportFills) res += " fills";
            return res;
        }
        
        public static bool SyncBalance { get { return YesNo(AccountSettingsWorksheet.Range["B" + SYNC_BALANCE]); } }
        public static bool SyncBalanceHistory { get { return YesNo(AccountSettingsWorksheet.Range["B" + SYNC_BALANCE_HISTORY]); } }
        public static bool SyncBalanceFills { get { return YesNo(AccountSettingsWorksheet.Range["B" + SYNC_FILLS]); } }

        private static readonly string[] YesNoArray = { "yes", "no" };

        private static void SetupYesNoDropdown(Range cell, bool defaultValue = false)
        {
            //if(cell.Validation == null)
            {
                cell.Validation.Delete();
                cell.Validation.Add(
                   XlDVType.xlValidateList,
                   XlDVAlertStyle.xlValidAlertInformation,
                   XlFormatConditionOperator.xlBetween,
                   string.Join(";", YesNoArray),
                   Type.Missing);

                cell.Validation.IgnoreBlank = true;
                cell.Validation.InCellDropdown = true;
            }

            if(cell.Value == null || Array.IndexOf(YesNoArray, cell.Value) < 0)
            {
                cell.Value = defaultValue ? YesNoArray[0] : YesNoArray[1];
            }
        }

        private static bool YesNo(Range cell)
        {
            return cell.Value == YesNoArray[0];
        }

        private static Worksheet m_AccountSettingsWorksheet = null;
        public static Worksheet AccountSettingsWorksheet
        {
            get
            {
                if (m_AccountSettingsWorksheet != null) return m_AccountSettingsWorksheet;

                Worksheet res = Globals.ThisAddIn.FindWorksheet(ACCOUNT_SETTINGS_WORKSHEET_NAME, true);
                m_AccountSettingsWorksheet = res;
                return res;
            }
        }

        public static void SetupWorksheet()
        {
            Worksheet s = AccountSettingsWorksheet;
            
            // settings
            s.Range["A" + SYNC_CONFIG].Value = "sync settings";
            s.Range["A" + SYNC_BALANCE].Value = "balance";
            s.Range["A" + SYNC_BALANCE_HISTORY].Value = "balance history";
            s.Range["A" + SYNC_FILLS].Value = "fills";
            SetupYesNoDropdown(s.Range["B" + SYNC_BALANCE]);
            SetupYesNoDropdown(s.Range["B" + SYNC_BALANCE_HISTORY]);
            SetupYesNoDropdown(s.Range["B" + SYNC_FILLS]);
            
            // accounts
            s.Range["A" + GDAX_ROW].Value = "gdax settings";
            s.Range["B" + GDAX_ROW].Value = SupportText(new GdaxConnector());
            s.Range["A" + GDAX_PASSPHARSE_ROW].Value = "passphrase";
            s.Range["A" + GDAX_KEY_ROW].Value = "key";
            s.Range["A" + GDAX_SECRET_ROW].Value = "secret";
            
            s.Range["A" + BITFINEX_ROW].Value = "bitfinex settings";
            s.Range["B" + BITFINEX_ROW].Value = SupportText(new BitFinexConnector());
            s.Range["A" + BITFINEX_KEY_ROW].Value = "key";
            s.Range["A" + BITFINEX_SECRET_ROW].Value = "secret";

            s.Range["A" + BITTREX_ROW].Value = "bittrex settings";
            s.Range["B" + BITTREX_ROW].Value = SupportText(new BittrexConnector());
            s.Range["A" + BITTREX_KEY_ROW].Value = "key";
            s.Range["A" + BITTREX_SECRET_ROW].Value = "secret";

            s.Range["A" + BINANCE_ROW].Value = "binance settings";
            s.Range["B" + BINANCE_ROW].Value = SupportText(new BinanceConnector());
            s.Range["A" + BINANCE_KEY_ROW].Value = "key";
            s.Range["A" + BINANCE_SECRET_ROW].Value = "secret";

            s.Range["A" + CRYPTOPIA_ROW].Value = "cryptopia settings";
            s.Range["B" + CRYPTOPIA_ROW].Value = SupportText(new CryptopiaConnector());
            s.Range["A" + CRYPTOPIA_KEY_ROW].Value = "key";
            s.Range["A" + CRYPTOPIA_SECRET_ROW].Value = "secret";

            s.Range["A" + ETHERSCAN_ROW].Value = "Ethereum address (Powered by Ethplorer.io)";
            s.Range["B" + ETHERSCAN_ROW].Value = SupportText(new EthplorerConnector());
            s.Range["A" + ETHERSCAN_KEY_ROW].Value = "public key";
        }

        public static IEnumerable<AccountConnector> ListAccounts()
        {
            Worksheet s = AccountSettingsWorksheet;
            List<AccountConnector> accounts = new List<AccountConnector>();

            for(int index = 0; GetParam(GDAX_PASSPHARSE_ROW,index) != ""; ++index)
            {
                accounts.Add(new GdaxConnector(GetParam(GDAX_PASSPHARSE_ROW, index), GetParam(GDAX_KEY_ROW, index), GetParam(GDAX_SECRET_ROW, index)));
            }

            for (int index = 0; GetParam(BITFINEX_KEY_ROW, index) != ""; ++index)
            {
                accounts.Add(new BitFinexConnector(GetParam(BITFINEX_KEY_ROW, index), GetParam(BITFINEX_SECRET_ROW, index)));
            }

            for (int index = 0; GetParam(BITTREX_KEY_ROW, index) != ""; ++index)
            {
                accounts.Add(new BittrexConnector(GetParam(BITTREX_KEY_ROW, index), GetParam(BITTREX_SECRET_ROW, index)));
            }

            for (int index = 0; GetParam(BINANCE_KEY_ROW, index) != ""; ++index)
            {
                accounts.Add(new BinanceConnector(GetParam(BINANCE_KEY_ROW, index), GetParam(BINANCE_SECRET_ROW, index)));
            }

            for (int index = 0; GetParam(CRYPTOPIA_KEY_ROW, index) != ""; ++index)
            {
                accounts.Add(new CryptopiaConnector(GetParam(CRYPTOPIA_KEY_ROW, index), GetParam(CRYPTOPIA_SECRET_ROW, index)));
            }

            for (int index = 0; GetParam(ETHERSCAN_KEY_ROW, index) != ""; ++index)
            {
                accounts.Add(new EthplorerConnector(GetParam(ETHERSCAN_KEY_ROW, index)));
            }

            accounts.Add(new ManualConnector());

            return accounts;
        }
    }
}
