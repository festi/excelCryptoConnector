using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CryptoConnector.ui
{
    /// <summary>
    /// Interaction logic for SyncStatus.xaml
    /// </summary>
    public partial class SyncStatus : Window, IAccountManagerEvents
    {
        Dictionary<string, AccountStatus> Accounts = new Dictionary<string, AccountStatus>();

        public SyncStatus()
        {
            InitializeComponent();
        }

        public void StartSync()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                this.Show();
            }));
        }

        public void EndSync()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                Accounts.Clear();
                CloseButton.IsEnabled = true;
            }));
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        public void StartSyncAccount(string id, string name)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var a = new AccountStatus(name);

                Accounts.Add(id, a);
                AccountList.Children.Add(a);
            }));
        }

        public void SyncAccountStatus(string id, string status)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                AccountStatus a;
                if(Accounts.TryGetValue(id, out a))
                {
                    a.SetStatus(status);
                }
            }));
        }
    }
}
