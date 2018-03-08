using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using CryptoConnector.ui;

namespace CryptoConnector
{
    public partial class Ribbon
    {
        AccountManager Exchanges;
        SyncStatus SyncStatusUI;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Exchanges = new AccountManager();
            SyncStatusUI = new SyncStatus();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var s = AccountsSheet.AccountSettingsWorksheet;
            AccountsSheet.SetupWorksheet();
            s.Activate();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<GdaxConnector>(SyncStatusUI);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<BitFinexConnector>(SyncStatusUI);
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.Refresh(SyncStatusUI);
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<BittrexConnector>(SyncStatusUI);
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<BinanceConnector>(SyncStatusUI);
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<CryptopiaConnector>(SyncStatusUI);
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<EthplorerConnector>(SyncStatusUI);
        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            Exchanges.RefreshOnly<NeoscanConnector>(SyncStatusUI);
        }
    }
}
