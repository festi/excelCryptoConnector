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
    /// Interaction logic for AccountStatus.xaml
    /// </summary>
    public partial class AccountStatus : UserControl
    {
        private AccountManagerEventLevel CurrentLevel = AccountManagerEventLevel.Info;

        public AccountStatus()
        {
            InitializeComponent();
        }

        public AccountStatus(string name):this()
        {
            AccName.Content = name;
        }

        internal void SetStatus(string status, AccountManagerEventLevel level)
        {
            if(level >= CurrentLevel)
            {
                AccStatus.Content = status;
                CurrentLevel = level;
            }
        }
    }
}
