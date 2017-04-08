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
using System.Windows.Shapes;

namespace MarketWizard.Windows
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        public LoginObject loginObject;
        public LoginWindow(LoginObject login)
        {
            InitializeComponent();
            loginObject = login;
            this.Focus();
        }

        private void continueButton_Click(object sender, RoutedEventArgs e)
        {
            loginObject.setAddress(addressTextBox.Text);
            loginObject.setPassword(passwordPBox.Password);
            Close();
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            loginObject.setAddress("NULL");
            Close();
        }
    }
}
