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

namespace Project_Cactus
{
    public partial class MainWindow : Window
    {
        // Values used for logging KCS things
        public string productFamily;

        // Timestamp for working out call duration
        public DateTime startTime;
        public TimeSpan callDuration;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void copyToClipboard_Button_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText("Hello!");
        }

        private void productComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Populates the Version drop-down with items depending on what product was selected, and makes the Version drop-down visible. Some products don't require the Version drop-down to appear.
            // Also sets the "Product Family" and "Integration" strings to fill in the KCS stuff.
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
            endTimerButton_Button.IsEnabled = true;
            startTimerButton_Button.IsEnabled = false;
            startTime = DateTime.Now;
        }

        private void endTimerButton_Click(object sender, RoutedEventArgs e)
        {
            endTimerButton_Button.IsEnabled = false;
            startTimerButton_Button.IsEnabled = true;
            callDuration = callDuration + (DateTime.Now - startTime);

            MessageBox.Show(callDuration.Minutes.ToString());
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
