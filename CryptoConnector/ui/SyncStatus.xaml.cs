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
                CloseButton.IsEnabled = true;
            }));
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }
    }
}
