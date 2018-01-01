using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;

namespace CryptoConnector
{
    public partial class Ribbon
    {
        AccountManager Exchanges;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Exchanges = new AccountManager();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var s = AccountsSheet.AccountSettingsWorksheet;
            AccountsSheet.SetupWorksheet();
            s.Activate();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<GdaxConnector>();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<BitFinexConnector>();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.Refresh();
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<BittrexConnector>();
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<BinanceConnector>();
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<CryptopiaConnector>();
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<EthplorerConnector>();
        }
    }
}
