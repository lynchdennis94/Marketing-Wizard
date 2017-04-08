using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
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
using System.Net.Mail;


namespace MarketWizard
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string LOG_DIR = "\\batch_logs";
        private const string LOG_START = "Log file for Marketing Wizard v1.5a";

        // Variables for old sessions
        private string addrFile;
        private string tempFile;
        private string attachFile;
        private string subj;
        private int nxtRow;
        private int mailSent;
        private string nxtAddr;
        private DateTime oldDT;

        private bool loadingOldValues = false;

        private Utilities.StatusTracker currentStatus;

        public MainWindow()
        {
            InitializeComponent();

            // Initialize a new status
            currentStatus = new Utilities.StatusTracker();

            // Determine if there were values from the last session
            if (readInSessionInfo(out oldDT, out addrFile, out nxtRow, out nxtAddr,
                                    out tempFile, out attachFile, out subj, out mailSent ))
            {
                loadingOldValues = true;
                // Will need to continue session from previous time
                bodyTextbox.Text = tempFile;
                //bodyTextbox.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
                
                attachmentTextbox.Text = attachFile;
               // attachmentTextbox.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });

                addressTextbox.Text = addrFile;
               // addressTextbox.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });

                subjectTextbox.Text = subj;
               // subjectTextbox.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
                loadingOldValues = false;
                previewButton.IsEnabled = true;
                sendButton.IsEnabled = true;
            }

            
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
            // TODO: Implement new preview functionality
            /*String addressLoc = "", bodyLoc = "", subject = "", attachmentLoc = "", currentDir = "";
            gatherEmailStrings(ref addressLoc, ref bodyLoc, ref subject,
                                ref attachmentLoc, ref currentDir);
           
            //EmailObject previewEmail = new EmailObject(attachmentLoc, bodyLoc, "previewemail@notanemail.com", subject);
            SMTPObject previewEmail = new SMTPObject(attachmentLoc, bodyLoc, "previewemail@notanemail.com", subject, "", "");
            PreviewWindow emailPreviewWindow = new PreviewWindow(previewEmail);
            emailPreviewWindow.Show();*/
        }

        private void Textbox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Catch null exceptions
            if (addressTextbox.Text == null)
            {
                addressTextbox.Text = "";
            }

            if (bodyTextbox.Text == null)
            {
                bodyTextbox.Text = "";
            }

            if (subjectTextbox.Text == null)
            {
                subjectTextbox.Text = "";
            }
            if(!loadingOldValues)
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
            
        }

        private void startNewBatch()
        {
            int sentMsgLimit;
            int messagesSent = 0;
            string stringMsgLimit = limitTextbox.Text;
            bool hitLimit = false;

            // Read in file for login information


            // Get the login information before we start ANY logging/emailing
            /*LoginObject loginObject = new LoginObject();
            Windows.LoginWindow lWindow = new Windows.LoginWindow(loginObject);
            lWindow.Show();

            // Wait for window to close
            while (lWindow != null)
            {
                lWindow.Activate();
                continue;
            }

            // Assuming we have given control up to lWindow, and now have it back
            if (loginObject.getAddress().Equals("NULL"))
            {
                // User cancelled out of login
                return;
            }

            string loginAddress = loginObject.getAddress();
            string loginPassword = loginObject.getPassword();*/

            string loginAddress = emailTextBox.Text;
            string loginPassword = pwdBox.Password;

            if (loginAddress.Equals(""))
            {
                MessageBox.Show("Please enter an email address", "Attention");
                return;
            }

            if (loginPassword.Equals(""))
            {
                MessageBox.Show("Please enter a password", "Attention");
                return;
            }

            // Get the limit on the number of messages to be sent in a session
            if (stringMsgLimit == null || stringMsgLimit.Equals(""))
            {
                sentMsgLimit = Int32.MaxValue;
            }
            else
            {
                sentMsgLimit = Int32.Parse(limitTextbox.Text);
            }

           /* if (mailSent > 0 )
            {
                // We sent some mail last time, so lower the limit
                sentMsgLimit
            }*/

            statusText.Text = "Starting new email batch ...";
            statusText.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
            bool failureOccurred = false;
            String addressLoc = "", bodyLoc="", subject="", attachmentLoc="", currentDir="";
            gatherEmailStrings(ref addressLoc, ref bodyLoc, ref subject,
                                ref attachmentLoc, ref currentDir);
            statusText.Text = "Preparing logs for batch ...";
            statusText.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
            // Check to see if there is a log folder
            if (!Directory.Exists(currentDir + LOG_DIR))
            {
                // Make the directory
                Directory.CreateDirectory(currentDir + LOG_DIR);
            }
           // Initiate a new text file to log any issues
            statusText.Text = "Initializing error logs ...";
            statusText.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
            String logFilePath = currentDir + LOG_DIR + "\\error_log_" + DateTime.Now.ToString().Replace(".", "_").Replace(":", "_").Replace("/", "_") + ".txt";
            System.IO.StreamWriter logFile = new System.IO.StreamWriter(logFilePath);
            logFile.WriteLine(LOG_START);
            logFile.WriteLine("****************************");

            Console.WriteLine("file is at " + logFilePath);

            // Create the excel sheet
            ExcelObject excelObj = new ExcelObject(addressLoc);

            statusText.Text = "Loading recepient addresses...";
            statusText.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
            // Create the dictionary for the addresses
            Dictionary<int, Utilities.ExcelRowBinder> allAddresses = excelObj.getAddresses();
            if (allAddresses.Count == 0)
            {
                MessageBox.Show("The current email cannot be sent.", "Sorry");
                subjectTextbox.Text = "";
                sendButton.IsEnabled = false;
                return;
            }
            currentStatus.setObjectiveTotal(allAddresses.Count);
            currentStatus.setStatus("Sending Messages");
            statusText.Text = currentStatus.getObjectiveStatus();
            statusText.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
            sendingProgressBar.Value = 0;
            sendingProgressBar.Maximum = allAddresses.Count;
            
            int nextRowSent = 0;
            string nextAddressSent = "";
            // Loop through addresses and make a new email for each one
            foreach (int keyIndex in allAddresses.Keys)
            {
                if (messagesSent < sentMsgLimit)
                {
                    String newAddr;
                    Utilities.ExcelRowBinder currentBinder;

                    // Need to extract the email value from the ExcelRowBinder

                    if (allAddresses.TryGetValue(keyIndex, out currentBinder) == true)
                    {
                        newAddr = currentBinder.getValue();
                        //EmailObject newEmail = new EmailObject(attachmentLoc, bodyLoc, newAddr, subject);
                        SMTPObject newEmail = new SMTPObject(attachmentLoc, bodyLoc, newAddr, subject, loginAddress, loginPassword);
                        try
                        {
                            newEmail.sendEmail();

                            if (newEmail.connectionError)
                            {
                                MessageBox.Show("Could not login to specified email account.\nPlease confirm the login details and try again", "Error");
                                logFile.WriteLine("LOGIN FAILED: " + loginAddress);
                                failureOccurred = true;
                                break;
                            }

                            try
                            {
                                currentStatus.incrementCurrentObjective();
                                statusText.Text = currentStatus.getObjectiveStatus();
                                statusText.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
                                excelObj.logSuccess(currentBinder.getRowNumber());
                            }
                            catch (Exception logException)
                            {
                                // Do nothing
                                Console.WriteLine("Error in trying to log stuff");
                            }

                        }
                        catch (SmtpException emailException)
                        {
                            MessageBox.Show("Could not login to specified email account.\nPlease confirm the login details and try again", "Error");
                            logFile.WriteLine("LOGIN FAILED: " + loginAddress);
                            failureOccurred = true;
                            break;
                        }
                        catch (Exception e)
                        {
                            // There was some reason the email could not be sent.  Log those errors
                            Console.WriteLine("Error in trying to send stuff");
                            Console.WriteLine("Trace: \n" + e.Message);
                            Console.WriteLine("Extended Trace: \n" + e.StackTrace);
                            logFile.WriteLine("SEND FAILED: " + newAddr);
                            failureOccurred = true;

                        }

                        sendingProgressBar.Value++;
                        sendingProgressBar.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, (Action)delegate() { });
                    }
                }
                else
                {
                    // We've hit the limit, time to stop this
                    hitLimit = true;
                    Utilities.ExcelRowBinder currentBinder;
                    if (allAddresses.TryGetValue(keyIndex, out currentBinder) == true)
                    {
                        nextRowSent = currentBinder.getRowNumber();
                        nextAddressSent = currentBinder.getValue();
                    }
                    //lastRowSent = 
                    break;
                }
                messagesSent++;
                
            }

            string sessionInfoPath = currentDir + "\\StoredSessionInfo.txt";
                System.IO.StreamWriter sessionInfo = new System.IO.StreamWriter(sessionInfoPath, false); // Overwrite old info
            if (!hitLimit)
            {
                // Alert the user that the task is done
                MessageBox.Show("All emails have been sent!", "Attention!");
                if (!failureOccurred)
                {
                    logFile.WriteLine("ALL MESSAGES SENT SUCCESSFULLY!!");
                }   
 
                // Overwrite the session info
                writeSessionInfo(sessionInfo, "N", "", 0, "", "", "", "", 0);
            }
            else
            {
                MessageBox.Show("The maximum number of emails have been sent for this session", "Attention");
                logFile.WriteLine("THE MAXIMUM {0} MESSAGES HAVE BEEN SENT!!", sentMsgLimit);

                // Amend the StoredSessionInfo.txt file
                

                writeSessionInfo(sessionInfo, "Y", addressLoc, nextRowSent, nextAddressSent, bodyLoc, attachmentLoc, subject, messagesSent);
            }
            sessionInfo.Close();
            sendingProgressBar.Visibility = System.Windows.Visibility.Hidden;
            logFile.Close();
            excelObj.saveAndClose();

            MessageBox.Show("All logs have been saved and closed", "Attention!");

            subjectTextbox.Clear();
            currentStatus.setStatus("");
            statusText.Text = "";
            
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

        private void writeSessionInfo(System.IO.StreamWriter writer, string readValues, string addrFile, int row, string addr, 
                                      string templateFile, string attachments, string subject, int totalMsgsSent)
        {
            // Should we read this file next time?
            writer.WriteLine("READ VALUES: " + readValues);
            // Last Batch:
            writer.WriteLine("LAST BATCH: " + DateTime.Now);

            // Last file:
            writer.WriteLine("LAST FILE: " + addrFile);

            // Last Row:
            writer.WriteLine("NEXT ROW: " + row);

            // Last Address:
            writer.WriteLine("NEXT ADDRESS: " + addr);

            // Last Template:
            writer.WriteLine("LAST TEMPLATE: " + templateFile);

            // Last Attachments:
            writer.WriteLine("LAST ATTACHMENTS: " + attachments); 

            // Last Subject:
            writer.WriteLine("LAST SUBJECT: " + subject);

            // Emails Sent
            writer.WriteLine("EMAILS SENT: " + totalMsgsSent);
        }
    
        private bool readInSessionInfo(out DateTime oldDateTime, out string oldAddressFile, out int oldNextRow,
                                        out string oldNextAddr, out string oldTemplateFile, out string oldAttachmentFile,
                                        out string oldSubject, out int oldEmailsSent)
        {
            string sessionInfoPath = Directory.GetCurrentDirectory() +"\\StoredSessionInfo.txt";
            System.IO.StreamReader reader = new System.IO.StreamReader(sessionInfoPath);

            string readValues = reader.ReadLine();

            // Initialize values
            oldAddressFile = "";
            oldTemplateFile = "";
            oldAttachmentFile = "";
            oldSubject = "";
            oldNextRow = -1;
            oldEmailsSent = -1;
            oldNextAddr = "";
            oldDateTime = DateTime.Now;

            if (readValues.Contains("READ VALUES: Y"))
            {

                MessageBoxResult result = MessageBox.Show("You have a previous session mail session that did not complete.\n\r" +
                                "Would you like to continue that session?\n\r" +
                                "Please note, if you do not continue this session and start a new session,\n\r" +
                                "all information for this stored session will be lost."
                                , "Previous Session Store", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    // Grab the information from the stored session
                    string[] tags = {"LAST BATCH: ", "LAST FILE: ", "NEXT ROW: ", "NEXT ADDRESS: ",
                                    "LAST TEMPLATE: ", "LAST ATTACHMENTS: ", "LAST SUBJECT: ", "EMAILS SENT: "};
                    /*object[] variables = {oldDateTime, oldAddressFile, oldNextRow, oldNextAddr, 
                                         oldTemplateFile, oldAttachmentFile, oldSubject, oldEmailsSent};*/
                    char[] trimChars = { '\n' };
                    string currentLine = reader.ReadLine();
                    oldDateTime = setDateTime(tags[0], currentLine);

                    currentLine = reader.ReadLine();
                    oldAddressFile = setString(tags[1], currentLine);

                    currentLine = reader.ReadLine();
                    oldNextRow = setInt(tags[2], currentLine);

                    currentLine = reader.ReadLine();
                    oldNextAddr = setString(tags[3], currentLine);

                    currentLine = reader.ReadLine();
                    oldTemplateFile = setString(tags[4], currentLine);

                    currentLine = reader.ReadLine();
                    oldAttachmentFile = setString(tags[5], currentLine);

                    currentLine = reader.ReadLine();
                    oldSubject = setString(tags[6], currentLine);

                    currentLine = reader.ReadLine();
                    oldEmailsSent = setInt(tags[7], currentLine);

                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }

        private string setString(string tag, string currentLine)
        {
            char[] trimChars = { '\n' };
            int startIndex = currentLine.IndexOf(tag) + tag.Length;
            return currentLine.Substring(startIndex).Trim(trimChars);
        }

        private int setInt(string tag, string currentLine)
        {
            char[] trimChars = { '\n' };
            int startIndex = currentLine.IndexOf(tag) + tag.Length;
            return Int32.Parse(currentLine.Substring(startIndex).Trim(trimChars));
        }

        private DateTime setDateTime(string tag, string currentLine)
        {
            char[] trimChars = { '\n' };
            int startIndex = currentLine.IndexOf(tag) + tag.Length;
            return DateTime.Parse(currentLine.Substring(startIndex).Trim(trimChars));
        }
    }
}
