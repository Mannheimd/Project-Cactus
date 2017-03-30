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
using System.Xml;

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

        // XML document to contain configuration data
        XmlDocument configurationXml = new XmlDocument();

        // Values for what is required when copying to clipboard
        bool productRequired = false;
        bool productVersionRequired = false;
        bool productUpdateRequired = false;
        bool osRequired = false;
        bool browserRequired = false;
        bool browserVersionRequired = false;
        bool sqlRequired = false;
        bool officeRequired = false;
        bool otherRequired = false;

        public MainWindow()
        {
            if (loadConfigurationXml())
            {
                InitializeComponent();
            }
            else
            {
                MessageBox.Show("Unable to load configuration file. Application will terminate.");
                Application.Current.Shutdown();
            }

            loadProductList();
        }

        private bool loadConfigurationXml()
        {
            // Loads the configuration XML from embedded resources. Later update will also store this locally and check a server for an updated version.
            try
            {
                string tempXml;

                using (Stream stream = this.GetType().Assembly.GetManifestResourceStream("Project_Cactus.Configuration.xml"))
                {
                    using (StreamReader sr = new StreamReader(stream))
                    {
                        tempXml = sr.ReadToEnd();
                    }
                }
                
                configurationXml.LoadXml(tempXml);

                return true;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

                return false;
            }
        }

        private void loadProductList()
        {
            try
            {
                // Empty the current product list
                DropDownLists.productList = null;

                // Create a temp list to store strings
                List<string> tempList = new List<string>();
                
                // Get the 'productlist' XML node
                XmlNode productListNode = configurationXml.SelectSingleNode("configuration/dropdowns/productlist");

                // Check if Product List is a required field
                if (productListNode.Attributes["required"].Value == "true")
                {
                    productRequired = true;
                }
                else
                {
                    productRequired = false;
                }
                
                // Get the list of products
                foreach (XmlNode product in configurationXml.SelectNodes("configuration/dropdowns/productlist/product"))
                {
                    tempList.Add(product.Attributes["text"].Value);
                }
                
                //Update the global product list
                DropDownLists.productList = tempList.ToArray();

                // Throw the list at the UI
                product_ComboBox.ItemsSource = DropDownLists.productList;
            }
            catch (Exception error)
            {
                MessageBox.Show("Unable to load product list. \n\n" + error.Message);
            }
        }

        private void productComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Populates the Version drop-down with items depending on what product was selected, and makes the Version drop-down visible. Some products don't require the Version drop-down to appear.
            // Also sets the "Product Family" and "Integration" strings to fill in the KCS stuff.
            if (e.AddedItems.Count > 0)
            {
                string selectedItem = (sender as ComboBox).SelectedItem.ToString();
                setRequiredRows(selectedItem);
            }
            else
            {
                
            }
        }

        // Method setVersionDropDown is being deprecated, replaced by setRequiredRows
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

        private void setRequiredRows(string selectedProduct)
        {
            try
            {
                // Set the XML path of the selected product from configuration
                string productPath = @"configuration/dropdowns/productlist/product[@text='" + selectedProduct + "']";
                
                // Work on versionlist in the XML, if it exists for the selected product
                if (configurationXml.SelectSingleNode(productPath + "/versionlist") != null)
                {
                    // Enable ProductVersion and ProductUpdate rows
                    productVersion_Row.Visibility = Visibility.Visible;
                    productUpdate_Row.Visibility = Visibility.Visible;

                    // Check if version drop-down is mandatory
                    XmlNode versionListNode = configurationXml.SelectSingleNode(productPath + "/versionlist");
                    if (versionListNode.Attributes["required"].Value == "true")
                    {
                        productVersionRequired = true;
                    }
                    else
                    {
                        productVersionRequired = false;
                    }

                    // Select all version nodes
                    XmlNodeList versionNodes = configurationXml.SelectNodes(productPath + "/versionlist/version");

                    // For each version node, add to the version list
                    DropDownLists.productVersionList = null;
                    List<string> tempVersionList = new List<string>();
                    foreach (XmlNode version in versionNodes)
                    {
                        tempVersionList.Add(version.Attributes["text"].Value);
                    }
                    DropDownLists.productVersionList = tempVersionList.ToArray();

                    // Set the items source on the UI
                    productVersion_ComboBox.ItemsSource = DropDownLists.productVersionList;
                }
                else
                {
                    // Collapse ProductVersion and ProductUpdate rows
                    productVersion_Row.Visibility = Visibility.Collapsed;
                    productUpdate_Row.Visibility = Visibility.Collapsed;

                    // Set the version drop-down to empty
                    productVersion_ComboBox.ItemsSource = DropDownLists.emptyList;

                    // Change productVersion and productUpdate to not mandatory
                    productVersionRequired = false;
                    productUpdateRequired = false;
                }

                // Work on oslist in the XML, if it exists for the selected product
                if (configurationXml.SelectSingleNode(productPath + "/oslist") != null)
                {
                    // Enable OS row
                    os_Row.Visibility = Visibility.Visible;

                    // Check if version drop-down is mandatory
                    XmlNode osListNode = configurationXml.SelectSingleNode(productPath + "/oslist");
                    if (osListNode.Attributes["required"].Value == "true")
                    {
                        osRequired = true;
                    }
                    else
                    {
                        osRequired = false;
                    }

                    // Select all version nodes
                    XmlNodeList osNodes = configurationXml.SelectNodes(productPath + "/oslist/os");

                    // For each version node, add to the version list
                    DropDownLists.osList = null;
                    List<string> tempOsList = new List<string>();
                    foreach (XmlNode os in osNodes)
                    {
                        tempOsList.Add(os.Attributes["text"].Value);
                    }
                    DropDownLists.osList = tempOsList.ToArray();

                    // Set the items source on the UI
                    os_ComboBox.ItemsSource = DropDownLists.osList;
                }
                else
                {
                    // Collapse OS row
                    os_Row.Visibility = Visibility.Collapsed;

                    // Set the OS drop-down to empty
                    os_ComboBox.ItemsSource = DropDownLists.emptyList;

                    // Change OS to not mandatory
                    osRequired = false;
                }

                // Work on browserlist in the XML, if it exists for the selected product
                if (configurationXml.SelectSingleNode(productPath + "/browserlist") != null)
                {
                    // Enable the Browser and BrowserVersion rows
                    browser_Row.Visibility = Visibility.Visible;
                    browserVersion_Row.Visibility = Visibility.Visible;

                    // Check if browser drop-down is mandatory
                    XmlNode browserListNode = configurationXml.SelectSingleNode(productPath + "/browserlist");
                    if (browserListNode.Attributes["required"].Value == "true")
                    {
                        browserRequired = true;
                    }
                    else
                    {
                        browserRequired = false;
                    }

                    // Select all browser nodes
                    XmlNodeList browserNodes = configurationXml.SelectNodes(productPath + "/browserlist/browser");

                    // For each browser node, add to the browser list
                    DropDownLists.browserList = null;
                    List<string> tempBrowserList = new List<string>();
                    foreach (XmlNode browser in browserNodes)
                    {
                        tempBrowserList.Add(browser.Attributes["text"].Value);
                    }
                    DropDownLists.browserList = tempBrowserList.ToArray();

                    // Set the items source on the UI
                    browser_ComboBox.ItemsSource = DropDownLists.browserList;
                }
                else
                {
                    // Collapse the Browser and BrowserVersion rows
                    browser_Row.Visibility = Visibility.Collapsed;
                    browserVersion_Row.Visibility = Visibility.Collapsed;

                    // Set the Browser drop-down to empty
                    browser_ComboBox.ItemsSource = DropDownLists.emptyList;

                    // Change Browser to not mandatory
                    browserRequired = false;
                }

                // Work on sqllist in the XML, if it exists for the selected product
                if (configurationXml.SelectSingleNode(productPath + "/sqllist") != null)
                {
                    // Enable the SQL row
                    sql_Row.Visibility = Visibility.Visible;

                    // Check if sql drop-down is mandatory
                    XmlNode sqlListNode = configurationXml.SelectSingleNode(productPath + "/sqllist");
                    if (sqlListNode.Attributes["required"].Value == "true")
                    {
                        sqlRequired = true;
                    }
                    else
                    {
                        sqlRequired = false;
                    }

                    // Select all sql nodes
                    XmlNodeList sqlNodes = configurationXml.SelectNodes(productPath + "/sqllist/sql");

                    // For each sql node, add to the sql list
                    DropDownLists.sqlList = null;
                    List<string> tempSqlList = new List<string>();
                    foreach (XmlNode sql in sqlNodes)
                    {
                        tempSqlList.Add(sql.Attributes["text"].Value);
                    }
                    DropDownLists.sqlList = tempSqlList.ToArray();

                    // Set the items source on the UI
                    sql_ComboBox.ItemsSource = DropDownLists.sqlList;
                }
                else
                {
                    // Collapse the SQL row
                    sql_Row.Visibility = Visibility.Collapsed;

                    // Set the SQL drop-down to empty
                    sql_ComboBox.ItemsSource = DropDownLists.emptyList;

                    // Change SQL to not mandatory
                    sqlRequired = false;
                }

                // Work on officelist in the XML, if it exists for the selected product
                if (configurationXml.SelectSingleNode(productPath + "/officelist") != null)
                {
                    // Enable the office row
                    office_Row.Visibility = Visibility.Visible;

                    // Check if office drop-down is mandatory
                    XmlNode officeListNode = configurationXml.SelectSingleNode(productPath + "/officelist");
                    if (officeListNode.Attributes["required"].Value == "true")
                    {
                        officeRequired = true;
                    }
                    else
                    {
                        officeRequired = false;
                    }

                    // Select all office nodes
                    XmlNodeList officeNodes = configurationXml.SelectNodes(productPath + "/officelist/office");

                    // For each sql node, add to the sql list
                    DropDownLists.officeList = null;
                    List<string> tempOfficeList = new List<string>();
                    foreach (XmlNode office in officeNodes)
                    {
                        tempOfficeList.Add(office.Attributes["text"].Value);
                    }
                    DropDownLists.officeList = tempOfficeList.ToArray();

                    // Set the items source on the UI
                    office_ComboBox.ItemsSource = DropDownLists.officeList;
                }
                else
                {
                    // Collapse the office row
                    office_Row.Visibility = Visibility.Collapsed;

                    // Set the office drop-down to empty
                    office_ComboBox.ItemsSource = DropDownLists.emptyList;

                    // Change office to not mandatory
                    officeRequired = false;
                }

                // If selected product isn't blank, enable Other row
                if (selectedProduct != null)
                {
                    other_Row.Visibility = Visibility.Visible;
                }
                // Otherwise, hide Other
                else
                {
                    other_Row.Visibility = Visibility.Collapsed;
                }
            }
            catch(Exception error)
            {
                MessageBox.Show("Unable to update context-based options: \n\n" + error.Message);
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
            startTime = new DateTime();
            endTime = new DateTime();
            callDuration = new TimeSpan();
            finalCallDuration = new TimeSpan();
            callDuration_Label.Content = "00:00:00";

            // Reset all fields and drop-downs
            reasonForCall_TextBox.Text = null;
            product_ComboBox.SelectedIndex = -1;
            productUpdate_TextBox.Text = null;
            os_ComboBox.SelectedIndex = -1;
            browser_ComboBox.SelectedIndex = -1;
            browserVersion_TextBox.Text = null;
            sql_ComboBox.SelectedIndex = -1;
            office_ComboBox.SelectedIndex = -1;
            other_TextBox.Text = null;
            errorMessages_TextBox.Text = null;
            stepsTaken_TextBox.Text = null;
            additionalInformation_TextBox.Text = null;
            resolution_TextBox.Text = null;

            // Trigger SetRequired Rows with "nothing" selected, to get rid of all rows
            setRequiredRows(null);
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
            os_ComboBox.Text,
            office_ComboBox.Text,
            browser_ComboBox.Text,
            sql_ComboBox.Text,
            other_TextBox.Text,
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

    public class VersionLists // Being deprecated due to using XML for data, replacing with DropDownLists class
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

    public class DropDownLists
    {
        public static string[] emptyList { get; set; }
        public static string[] productList { get; set; }
        public static string[] productVersionList { get; set; }
        public static string[] osList { get; set; }
        public static string[] browserList { get; set; }
        public static string[] sqlList { get; set; }
        public static string[] officeList { get; set; }
    }
}
