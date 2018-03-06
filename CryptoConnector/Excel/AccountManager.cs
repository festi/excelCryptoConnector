using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace CryptoConnector
{
    class AccountManager
    {
        public static readonly string BALANCE_QUERY = "balance";

        public static readonly string EXCHANGE_COLUMN = "account";

        public void Refresh()
        {
            // refresh everything
            RefreshOnly<AccountConnector>();
        }

        public void RefreshOnly<T>()
        {
            try
            {
                // list all accounts
                var accounts = AccountsSheet.ListAccounts();

                // refresh account of a specific type
                foreach (var ex in accounts)
                {
                    if (ex is T) ex.Refresh();
                }

                // setup queries if necessary
                SetupQueries(accounts);

                // run all queries
                Globals.ThisAddIn.Application.ActiveWorkbook.RefreshAll();
            }
            catch (Exception e)
            {
                //var errorDialog = new ErrorDialog();
                //errorDialog.SetText(e.Message);
                //errorDialog.Show();
                var result = MessageBox.Show($"{e.Message}\n{e.StackTrace}", "",
                                             MessageBoxButtons.OK,
                                             MessageBoxIcon.Error);
            }
        }

        public void SetupQueries(IEnumerable<AccountConnector> accounts)
        {
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (accounts.Count() == 0) return;

            if (AccountsSheet.SyncBalance)
            {
                try
                {
                    // if the workbook contains no queries, this line throw an exception
                    bool exist = workbook.Queries.Cast<WorkbookQuery>().Any(x => x.Name == BALANCE_QUERY);
                    var query = exist ? workbook.Queries[BALANCE_QUERY] : workbook.Queries.Add(BALANCE_QUERY, "");
                    var expectedFormula = BalanceQueryFormula(accounts);

                    if (!exist || query.Formula != expectedFormula)
                    {
                        query.Formula = BalanceQueryFormula(accounts);
                    }
                }
                catch(Exception e)
                {

                    var result = MessageBox.Show("Please create any query in the current excel file. The plugin cannot create one if none exist", "",
                                                 MessageBoxButtons.OK,
                                                 MessageBoxIcon.Error);
                }
            }
        }
        
        private string BalanceQueryFormula(IEnumerable<AccountConnector> balancesQueries)
        {
            return $@"let
Source = Table.Combine({{{string.Join(",", ExchangeConnectorToBalanceQuery(balancesQueries))}}})
in
Source";
        }

        private IEnumerable<string> ExchangeConnectorToBalanceQuery(IEnumerable<AccountConnector> src)
        {
            foreach(var account in src)
            {
                if (account.BalanceContainsAccountName)
                {
                    yield return $"Excel.CurrentWorkbook(){{[Name=\"{account.BalanceSheetName}\"]}}[Content]";
                }
                else
                {
                    yield return AddColumn($"Excel.CurrentWorkbook(){{[Name=\"{account.BalanceSheetName}\"]}}[Content]", EXCHANGE_COLUMN, account.ReadableName);
                }
            }
        }

        private string AddColumn(string src, string columnname, string columnValue)
        {
            return $"Table.AddColumn({src}, \"{columnname}\", each \"{columnValue}\")";
        }
    }
}
