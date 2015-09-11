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

namespace MarketWizard
{
    /// <summary>
    /// Interaction logic for PreviewWindow.xaml
    /// </summary>
    public partial class PreviewWindow : Window
    {
        private EmailObject previewEmailObj;
        public PreviewWindow(EmailObject emailToPreview)
        {
            previewEmailObj = emailToPreview;
            InitializeComponent();
            loadPreview();

            // DEBUG ListBox
            addressListBox.SelectionMode = SelectionMode.Multiple;
            for (int test = 500600; test < 500750; test++)
            {
                addressListBox.Items.Add(test.ToString());
            }
            addressListBox.SelectAll();
        }

        private void loadPreview()
        {
            // Get the information from the preview email object, and populate the window
            subjectTextBox.Text = previewEmailObj.getSubject();
            bodyTextBlock.Text = previewEmailObj.getBody();
        }

        private void buttonCancel_Click(object sender, RoutedEventArgs e)
        {
            // Just back out of this window, we aren't sending or saving changes
            this.Close();
        }
    }
}
