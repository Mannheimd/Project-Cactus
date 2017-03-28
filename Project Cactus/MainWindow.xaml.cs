using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
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
using System.Windows.Threading;

namespace Project_Cactus
{
    public partial class MainWindow : Window
    {
        // Values used for logging KCS things
        public string productFamily;

        // Values for working out call duration
        public DateTime startTime;
        public DateTime endTime;
        public bool isTimerRunning = false;
        public TimeSpan callDuration;
        public TimeSpan finalCallDuration;

        // Timer for on-screen display of call duration
        DispatcherTimer durationCounter = new DispatcherTimer();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void productComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Populates the Version drop-down with items depending on what product was selected, and makes the Version drop-down visible. Some products don't require the Version drop-down to appear.
            // Also sets the "Product Family" and "Integration" strings to fill in the KCS stuff.
            if (e.AddedItems.Count > 0)
            {
                string selectedItem = (e.AddedItems[0] as ComboBoxItem).Content as string;
                switch (selectedItem)
                {
                    case "Act! Pro":
                        setVersionDropDown(true, false, "actVersions");
                        productFamily = "Act!";
                        break;

                    case "Act! Premium":
                        setVersionDropDown(true, false, "actVersions");
                        productFamily = "Act!";
                        break;

                    case "Act! Premium for Web":
                        setVersionDropDown(true, false, "actVersions");
                        productFamily = "Act!";
                        break;

                    case "Act! Premium Cloud":
                        setVersionDropDown(false, false, "blankList");
                        productFamily = "Act!";
                        break;

                    case "Act! Emarketing":
                        setVersionDropDown(true, false, "actVersions");
                        productFamily = "Act!";
                        break;

                    case "Swiftpage Emarketing":
                        setVersionDropDown(false, false, "blankList");
                        productFamily = "Emarketing";
                        break;

                    case "Other":
                        setVersionDropDown(false, false, "blankList");
                        productFamily = "Other";
                        break;

                    case "N/A":
                        setVersionDropDown(false, false, "blankList");
                        productFamily = "N/A";
                        break;

                    default:
                        setVersionDropDown(false, false, "blankList");
                        productFamily = null;
                        break;
                }
            }
            else
            {
                setVersionDropDown(false, false, "blankList");
                productFamily = null;
            }
        }

        private void setVersionDropDown(bool dropdownEnable, bool textEnable, string versionList)
        {
            // Sets the Version dropdown and label to active or inactive and chooses a list to apply to it's item source
            // 'textEnable' bool is to allow for future addition of a text box when "other" is selected

            if (dropdownEnable) // Make visible, apply selected list
            {
                productVersion_Label.Visibility = Visibility.Visible;
                productVersion_ComboBox.Visibility = Visibility.Visible;
                productVersion_ComboBox.ItemsSource = VersionLists.listSelector(versionList);
            }
            else // Make invisible, apply blank list
            {
                productVersion_Label.Visibility = Visibility.Collapsed;
                productVersion_ComboBox.Visibility = Visibility.Collapsed;
                productVersion_ComboBox.ItemsSource = VersionLists.listSelector("blankList");
            }
        }

        private void startTimerButton_Click(object sender, RoutedEventArgs e)
        {
            // Enable/disable buttons
            endTimer_Button.IsEnabled = true;
            startTimer_Button.IsEnabled = false;
            resetForm_Button.IsEnabled = false;
            copyToClipboard_Button.IsEnabled = false;

            // Save current time and call duration, start counting new duration
            startTime = DateTime.Now;
            finalCallDuration = finalCallDuration + callDuration;
            callDuration = new TimeSpan();
            isTimerRunning = true;

            // Start on-screen counting timer
            durationCounter.Tick += new EventHandler(durationCounter_Tick);
            durationCounter.Interval = new TimeSpan(0, 0, 1);
            durationCounter.Start();
        }

        private void endTimerButton_Click(object sender, RoutedEventArgs e)
        {
            // Enable/disable buttons
            endTimer_Button.IsEnabled = false;
            startTimer_Button.IsEnabled = true;
            resetForm_Button.IsEnabled = true;
            copyToClipboard_Button.IsEnabled = true;

            // Save final call duration (resets when reset button is pressed)
            endTime = DateTime.Now;
            calculateCallDuration();
            isTimerRunning = false;

            // End on-screen counting timer
            durationCounter.Stop();
        }

        private void resetFormButton_Click(object sender, RoutedEventArgs e)
        {
            // Reset timer
            callDuration = new TimeSpan();
            finalCallDuration = new TimeSpan();
            callDuration_Label.Content = "00:00:00";

            // Reset all fields and drop-downs
            reasonForCall_TextBox.Text = null;
            product_ComboBox.SelectedIndex = -1;
            environmentOs_ComboBox.SelectedIndex = -1;
            environmentBrowser_ComboBox.SelectedIndex = -1;
            environmentSql_ComboBox.SelectedIndex = -1;
            environmentOffice_ComboBox.SelectedIndex = -1;
            environmentOther_TextBox.Text = null;
            errorMessages_TextBox.Text = null;
            stepsTaken_TextBox.Text = null;
            additionalInformation_TextBox.Text = null;
            resolution_TextBox.Text = null;
        }

        private void copyToClipboard_Button_Click(object sender, RoutedEventArgs e)
        {
            string outputString = String.Format(@"Environmental Info:
- Act version: {0} {1}
- Windows version: {2}
- Office version: {3}
- Web browser: {4}
- SQL: {5}
- Other: {6}

---
Reason for Call:
{7}

---
Error messages:
{8}

---
Additional issue/query information:
{9}

---
Steps Taken:
{10}

---
Solution:
{11}

---
Duration: {12}",
            product_ComboBox.Text,
            productVersion_ComboBox.Text,
            environmentOs_ComboBox.Text,
            environmentOffice_ComboBox.Text,
            environmentBrowser_ComboBox.Text,
            environmentSql_ComboBox.Text,
            environmentOther_TextBox.Text,
            reasonForCall_TextBox.Text,
            errorMessages_TextBox.Text,
            additionalInformation_TextBox.Text,
            stepsTaken_TextBox.Text,
            resolution_TextBox.Text,
            calculateCallDuration());

            Clipboard.SetText(outputString);
        }

        private void durationCounter_Tick(object sender, EventArgs e)
        {
            // Get current call duration, 
            callDuration_Label.Content = calculateCallDuration();
        }

        private string calculateCallDuration()
        {
            // Set current call duration, and also return total call duration as a string
            if (isTimerRunning)
            {
                callDuration = DateTime.Now - startTime;
            }
            else
            {
                callDuration = endTime - startTime;
            }
            TimeSpan tempCallDuration = finalCallDuration + callDuration;
            return tempCallDuration.ToString().Substring(0, 8);
        }
    }

    public class VersionLists
    {
        //Defines different lists for things
        static string[] blankList = {""};

        static string[] actVersions = {
            "v19.1",
            "v19.0",
            "v18.2",
            "v18.1",
            "v18.0",
            "v17.2",
            "v17.1",
            "v17.0",
            "v16.3",
            "v16.2",
            "v16.1",
            "v16.0",
            "v15.1",
            "v15.0",
            "v14.2",
            "v14.1",
            "v14.0",
            "v13.1",
            "v13.0",
            "v12.2",
            "v12.1",
            "v12.0",
            "v7-v11",
            "v6 or below",
            "Other",
            "N/A"
        };

        static string[] officeVersions = {
            "Office 2016",
            "Office 2016 (Click to Run)",
            "Office 2013",
            "Office 2013 (Click to Run)",
            "Office 2010",
            "Office 2007",
            "Office 2003",
            "N/A",
            "Unsupported Version"
        };

        static string[] windowsVersions =
        {
            "Windows 10",
            "Windows 8.1",
            "Windows 8",
            "Windows 7",
            "Windows Vista",
            "Windows XP",
            "Windows Server 2012",
            "Windows Server 2008 R2",
            "Windows Server 2008",
            "Windows Server 2003",
            "Windows Home Server",
            "Windows Home Server 2011",
            "N/A",
            "Other"
        };

        //Returns the selected list
        public static string[] listSelector(string list)
        {
            switch (list)
            {
                case "blankList":
                    return blankList;

                case "actVersions":
                    return actVersions;

                case "officeVersions":
                    return officeVersions;

                case "windowsVersions":
                    return windowsVersions;

                default:
                    return blankList;
            };
        }
    }
}
