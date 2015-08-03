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
using System.IO;


namespace MarketWizard
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string LOG_DIR = "\\batch_logs";
        private const string LOG_START = "Log file for Marketing Wizard v1.1";

        public MainWindow()
        {
            InitializeComponent();
        }
           
        private void sendButton_Click(object sender, RoutedEventArgs e)
        {
            sendButton.IsEnabled = false;
            sendButton.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });

            //Set up progress bar values
            sendingProgressBar.Visibility = System.Windows.Visibility.Visible;
            sendingProgressBar.Minimum = 0;
            sendingProgressBar.Value = 0;

            sendingProgressBar.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
            startNewBatch();
        }

        private void quitButton_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void previewButton_Click(object sender, RoutedEventArgs e)
        {
            String addressLoc = "", bodyLoc = "", subject = "", attachmentLoc = "", currentDir = "";
            gatherEmailStrings(ref addressLoc, ref bodyLoc, ref subject,
                                ref attachmentLoc, ref currentDir);
           
            EmailObject previewEmail = new EmailObject(attachmentLoc, bodyLoc, "previewemail@notanemail.com", subject);

            PreviewWindow emailPreviewWindow = new PreviewWindow(previewEmail);
            emailPreviewWindow.Show();
        }

        private void Textbox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!addressTextbox.GetLineText(0).Equals("") && !bodyTextbox.GetLineText(0).Equals("") && !subjectTextbox.GetLineText(0).Equals(""))
            {
                sendButton.IsEnabled = true;
                previewButton.IsEnabled = true;
            }
            else if (!bodyTextbox.GetLineText(0).Equals("") && !subjectTextbox.GetLineText(0).Equals(""))
            {
                sendButton.IsEnabled = false;
                previewButton.IsEnabled = true;
            }
            else
            {
                sendButton.IsEnabled = false;
            }
        }

        private void startNewBatch()
        {
            bool failureOccurred = false;
            String addressLoc = "", bodyLoc="", subject="", attachmentLoc="", currentDir="";
            gatherEmailStrings(ref addressLoc, ref bodyLoc, ref subject,
                                ref attachmentLoc, ref currentDir);

            // Check to see if there is a log folder
            if (!Directory.Exists(currentDir + LOG_DIR))
            {
                // Make the directory
                Directory.CreateDirectory(currentDir + LOG_DIR);
            }
           // Initiate a new text file to log any issues
            String logFilePath = currentDir + LOG_DIR + "\\error_log_" + DateTime.Now.ToString().Replace(".", "_").Replace(":", "_").Replace("/", "_") + ".txt";
            System.IO.StreamWriter logFile = new System.IO.StreamWriter(logFilePath);
            logFile.WriteLine(LOG_START);
            logFile.WriteLine("****************************");
            
            

            // Create the excel sheet
            ExcelObject excelObj = new ExcelObject(addressLoc);

            // Create the dictionary for the addresses
            Dictionary<int, Utilities.ExcelRowBinder> allAddresses = excelObj.getAddresses();
            sendingProgressBar.Maximum = allAddresses.Count;
            
            // Loop through addresses and make a new email for each one
            foreach (int keyIndex in allAddresses.Keys)
            {
                String newAddr;
                Utilities.ExcelRowBinder currentBinder;
                
                // Need to extract the email value from the ExcelRowBinder

                if (allAddresses.TryGetValue(keyIndex, out currentBinder) == true)
                {
                    newAddr = currentBinder.getValue();
                    EmailObject newEmail = new EmailObject(attachmentLoc, bodyLoc, newAddr, subject);
                    try
                    {
                        newEmail.sendEmail();
                        try
                        {
                            excelObj.logSuccess(currentBinder.getRowNumber());
                        }
                        catch (Exception logException)
                        {
                            // Do nothing
                        }
                        
                    }
                    catch (Exception e)
                    {
                        // There was some reason the email could not be sent.  Log those errors
                        logFile.WriteLine("SEND FAILED: " + newAddr);
                        failureOccurred = true;

                    }
                    
                    sendingProgressBar.Value++;
                    sendingProgressBar.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
                }
                
            }
            sendingProgressBar.Visibility = System.Windows.Visibility.Hidden;

            // Alert the user that the task is done
            MessageBox.Show("All emails have been sent!", "Attention!");
            if (!failureOccurred)
            {
                logFile.WriteLine("ALL MESSAGES SENT SUCCESSFULLY!!");
            }
            logFile.Close();
            excelObj.saveAndClose();
            MessageBox.Show("All logs have been saved and closed", "Attention!");
            
            subjectTextbox.Clear();
        }

        private void gatherEmailStrings(ref String addressLoc, ref String bodyLoc,
                                        ref String subject, ref String attachmentLoc,
                                        ref String currentDir)
        {
            addressLoc = addressTextbox.Text.Replace("\"", "\\");
            bodyLoc = bodyTextbox.Text.Replace("\"", "\\");
            subject = subjectTextbox.Text.Replace("\"", "\\");
            attachmentLoc = attachmentTextbox.Text.Replace("\"", "\\");
            currentDir = Directory.GetCurrentDirectory();
        }

        private void browseBody_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                bodyTextbox.Text = ofd.FileName;
            }
        }

        private void browseAttachment_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.Multiselect = true;
            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                foreach (String attachmentFile in ofd.FileNames)
                {
                    attachmentTextbox.Text = attachmentTextbox.Text + attachmentFile + ";";
                }
            }
        }

        private void browseAddress_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                addressTextbox.Text = ofd.FileName;
            }
        }
    }
}
