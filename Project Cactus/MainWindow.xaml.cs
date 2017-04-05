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

        // Default values for what is required when copying to clipboard
        // These are altered in setRequiredRows()
        bool reasonForCallRequired = true;
        bool productRequired = false;
        bool productNameRequired = false;
        bool versionTextRequired = false;
        bool productVersionRequired = false;
        bool productUpdateRequired = false;
        bool osRequired = false;
        bool browserRequired = false;
        bool browserVersionRequired = false;
        bool accountNameRequired = false;
        bool sqlRequired = false;
        bool officeRequired = false;
        bool otherRequired = false;
        bool errorMessagesRequired = true;
        bool additionalInformationRequired = true;
        bool stepsTakenRequired = true;
        bool resolutionRequired = true;
        bool apcInfoRequired = false;

        // Default values for what must not be blank when copying to clipboard
        // These are altered in setRequiredRows()
        bool reasonForCallMandatory = true;
        bool productMandatory = true;
        bool productNameMandatory = false;
        bool versionTextMandatory = false;
        bool productVersionMandatory = false;
        bool productUpdateMandatory = false;
        bool osMandatory = false;
        bool browserMandatory = false;
        bool browserVersionMandatory = false;
        bool accountNameMandatory = false;
        bool sqlMandatory = false;
        bool officeMandatory = false;
        bool otherMandatory = false;
        bool errorMessagesMandatory = false;
        bool additionalInformationMandatory = false;
        bool stepsTakenMandatory = true;
        bool resolutionMandatory = true;
        bool urlMandatory = false;
        bool databaseNameMandatory = false;

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
            if (e.AddedItems.Count > 0)
            {
                string selectedItem = (sender as ComboBox).SelectedItem.ToString();
                setRequiredRows(selectedItem);
            }
            else
            {
                // Trigger SetRequired Rows with "nothing" selected, to get rid of all rows
                setRequiredRows(null);
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
                        productVersionMandatory = true;
                    }
                    else
                    {
                        productVersionMandatory = false;
                    }

                    // Flag productVersion and productUpdate for output
                    productVersionRequired = true;
                    productUpdateRequired = true;

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
                    productVersionMandatory = false;
                    productUpdateMandatory = false;

                    // Flag productVersion and productUpdate to not be included in output
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
                        osMandatory = true;
                    }
                    else
                    {
                        osMandatory = false;
                    }

                    // Flag OS for inclusion in output
                    osRequired = true;

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
                    osMandatory = false;

                    // Flag OS as not required in output
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
                        browserMandatory = true;
                    }
                    else
                    {
                        browserMandatory = false;
                    }

                    // Flag browser and browserVersion as required in output
                    browserRequired = true;
                    browserVersionRequired = true;

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
                    browserMandatory = false;

                    // Flag browser and browserVersion as not required in output
                    browserRequired = false;
                    browserVersionRequired = false;
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
                        sqlMandatory = true;
                    }
                    else
                    {
                        sqlMandatory = false;
                    }

                    // Flag sql as required in output
                    sqlRequired = true;

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
                    sqlMandatory = false;

                    // Flag sql as not required in output
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
                        officeMandatory = true;
                    }
                    else
                    {
                        officeMandatory = false;
                    }

                    // Flag office as required in output
                    officeRequired = true;

                    // Select all office nodes
                    XmlNodeList officeNodes = configurationXml.SelectNodes(productPath + "/officelist/office");

                    // For each office node, add to the sql list
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
                    officeMandatory = false;

                    // Flag sql as not required in output
                    officeRequired = false;
                }

                // Work on productname in the XML, if it exists for the selected product
                if (configurationXml.SelectSingleNode(productPath + "/productname") != null)
                {
                    // Enable the product name row
                    productName_Row.Visibility = Visibility.Visible;

                    // Check if product name field is mandatory
                    XmlNode productNameNode = configurationXml.SelectSingleNode(productPath + "/productname");
                    if (productNameNode.Attributes["required"].Value == "true")
                    {
                        productNameMandatory = true;
                    }
                    else
                    {
                        productNameMandatory = false;
                    }

                    // Flag product name as required in output
                    productNameRequired = true;
                }
                else
                {
                    // Collapse the product name row
                    productName_Row.Visibility = Visibility.Collapsed;

                    // Change product name to not mandatory
                    productNameMandatory = false;

                    // Flag product name as not required in output
                    productNameRequired = false;
                }

                // Work on versiontext in the XML, if it exists for the selected product
                if (configurationXml.SelectSingleNode(productPath + "/versiontext") != null)
                {
                    // Enable the version text row
                    versionText_Row.Visibility = Visibility.Visible;

                    // Check if version text field is mandatory
                    XmlNode versionTextNode = configurationXml.SelectSingleNode(productPath + "/versiontext");
                    if (versionTextNode.Attributes["required"].Value == "true")
                    {
                        versionTextMandatory = true;
                    }
                    else
                    {
                        versionTextMandatory = false;
                    }

                    // Flag version text as required in output
                    versionTextRequired = true;
                }
                else
                {
                    // Collapse the version text row
                    versionText_Row.Visibility = Visibility.Collapsed;

                    // Change version text to not mandatory
                    versionTextMandatory = false;

                    // Flag version text as not required in output
                    versionTextRequired = false;
                }

                // Work on accountname in the XML, if it exists for the selected product
                if (configurationXml.SelectSingleNode(productPath + "/accountname") != null)
                {
                    // Enable the account name row
                    accountName_Row.Visibility = Visibility.Visible;

                    // Check if account name field is mandatory
                    XmlNode accountNameNode = configurationXml.SelectSingleNode(productPath + "/accountname");
                    if (accountNameNode.Attributes["required"].Value == "true")
                    {
                        accountNameMandatory = true;
                    }
                    else
                    {
                        accountNameMandatory = false;
                    }

                    // Flag account name as required in output
                    accountNameRequired = true;
                }
                else
                {
                    // Collapse the account name row
                    accountName_Row.Visibility = Visibility.Collapsed;

                    // Change account name to not mandatory
                    accountNameMandatory = false;

                    // Flag account name as not required in output
                    accountNameRequired = false;
                }

                // Work on apcinfo in the XML, if it exists for the selected product
                if (configurationXml.SelectSingleNode(productPath + "/apcinfo") != null)
                {
                    // Enable the APC Info grid
                    apcInfoText_Row.Visibility = Visibility.Visible;
                    url_Row.Visibility = Visibility.Visible;
                    databaseName_Row.Visibility = Visibility.Visible;
                    remoteDatabaseName_Row.Visibility = Visibility.Visible;

                    // Check if APC info fields are mandatory
                    XmlNode apcInfoNode = configurationXml.SelectSingleNode(productPath + "/apcinfo");
                    if (apcInfoNode.Attributes["required"].Value == "true")
                    {
                        urlMandatory = true;
                        databaseNameMandatory = true;
                        // Does not set remoteDatabaseName
                    }
                    else
                    {
                        urlMandatory = false;
                        databaseNameMandatory = false;
                        // Does not set remoteDatabaseName
                    }

                    // Flag APC info fields as required in output
                    apcInfoRequired = true;
                }
                else
                {
                    // Collapse the APC Info rows
                    apcInfoText_Row.Visibility = Visibility.Collapsed;
                    url_Row.Visibility = Visibility.Collapsed;
                    databaseName_Row.Visibility = Visibility.Collapsed;
                    remoteDatabaseName_Row.Visibility = Visibility.Collapsed;

                    // Change APC info fields to not mandatory
                    urlMandatory = false;
                    databaseNameMandatory = false;

                    // Flag APC info as not required in output
                    apcInfoRequired = false;
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
            productName_TextBox.Text = null;
            versionText_TextBox.Text = null;
            productUpdate_TextBox.Text = null;
            os_ComboBox.SelectedIndex = -1;
            browser_ComboBox.SelectedIndex = -1;
            browserVersion_TextBox.Text = null;
            accountName_TextBox.Text = null;
            sql_ComboBox.SelectedIndex = -1;
            office_ComboBox.SelectedIndex = -1;
            other_TextBox.Text = null;
            errorMessages_TextBox.Text = null;
            stepsTaken_TextBox.Text = null;
            additionalInformation_TextBox.Text = null;
            resolution_TextBox.Text = null;
        }

        private void copyToClipboard_Button_Click(object sender, RoutedEventArgs e)
        {
            if (checkMandatoryCriteriaMet())
            {
                Clipboard.SetText(buildTicketOutput());
            }
            else
            {
                MessageBox.Show("Required fields have not been filled in. Please update highlighted fields.");
            }
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

        private bool checkMandatoryCriteriaMet()
        {
            // For every field marked as mandatory, check if it's associated UI element is blank.
            // If mandatory field is empty, make its row orange (Swiftpage Orange, hell yah) and flag failure.
            // if all is well, remove the colour.
            bool criteriaMet = true;

            // reasonForCall
            if (reasonForCallMandatory & reasonForCall_TextBox.Text == "")
            {
                reasonForCall_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                reasonForCall_Grid.ClearValue(BackgroundProperty);
            }

            // product
            if (productMandatory & product_ComboBox.Text == "")
            {
                product_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254,80,0)));
                criteriaMet = false;
            }
            else
            {
                product_Row.ClearValue(BackgroundProperty);
            }

            // productName
            if (productNameMandatory & productName_TextBox.Text == "")
            {
                productName_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                productName_Row.ClearValue(BackgroundProperty);
            }

            // versionText
            if (versionTextMandatory & versionText_TextBox.Text == "")
            {
                versionText_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                versionText_Row.ClearValue(BackgroundProperty);
            }

            // productVersion
            if (productVersionMandatory & productVersion_ComboBox.Text == "")
            {
                productVersion_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                productVersion_Row.ClearValue(BackgroundProperty);
            }

            // productUpdate
            if (productUpdateMandatory & productUpdate_TextBox.Text == "")
            {
                productUpdate_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                productUpdate_Row.ClearValue(BackgroundProperty);
            }

            // os
            if (osMandatory & os_ComboBox.Text == "")
            {
                os_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                os_Row.ClearValue(BackgroundProperty);
            }

            // browser
            if (browserMandatory & browser_ComboBox.Text == "")
            {
                browser_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                browser_Row.ClearValue(BackgroundProperty);
            }

            // browserVersion
            if (browserVersionMandatory & browserVersion_TextBox.Text == "")
            {
                browserVersion_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                browserVersion_Row.ClearValue(BackgroundProperty);
            }

            // accountName
            if (accountNameMandatory & accountName_TextBox.Text == "")
            {
                accountName_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                accountName_Row.ClearValue(BackgroundProperty);
            }

            // sql
            if (sqlMandatory & sql_ComboBox.Text == "")
            {
                sql_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                sql_Row.ClearValue(BackgroundProperty);
            }

            // office
            if (officeMandatory & office_ComboBox.Text == "")
            {
                office_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                office_Row.ClearValue(BackgroundProperty);
            }

            // other
            if (otherMandatory & other_TextBox.Text == "")
            {
                other_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                other_Row.ClearValue(BackgroundProperty);
            }

            // errorMessages
            if (errorMessagesMandatory & errorMessages_TextBox.Text == "")
            {
                errorMessages_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                errorMessages_Grid.ClearValue(BackgroundProperty);
            }

            // additionalInformation
            if (additionalInformationMandatory & additionalInformation_TextBox.Text == "")
            {
                additionalInformation_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                additionalInformation_Grid.ClearValue(BackgroundProperty);
            }

            // stepsTaken
            if (stepsTakenMandatory & stepsTaken_TextBox.Text == "")
            {
                stepsTaken_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                stepsTaken_Grid.ClearValue(BackgroundProperty);
            }

            // resolution
            if (resolutionMandatory & resolution_TextBox.Text == "")
            {
                resolution_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                resolution_Grid.ClearValue(BackgroundProperty);
            }

            // url
            if (urlMandatory & url_TextBox.Text == "")
            {
                url_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                url_Row.ClearValue(BackgroundProperty);
            }

            // databaseName
            if (databaseNameMandatory & url_TextBox.Text == "")
            {
                databaseName_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                databaseName_Row.ClearValue(BackgroundProperty);
            }

            return criteriaMet;
        }

        private string buildTicketOutput()
        {
            // For each required data type, add data to the output string.
            string newLine = Environment.NewLine;
            string outputString = "";

            // == Environmental info heading ==
            outputString = outputString + "Environment and Version Information:" + newLine;

            // product
            if (productRequired)
            {
                outputString = outputString + "Product: " + product_ComboBox.Text + newLine;
            }

            // productName
            if (productNameRequired)
            {
                outputString = outputString + "Product Name: " + productName_TextBox.Text + newLine;
            }

            // versionText
            if (versionTextRequired)
            {
                outputString = outputString + "Version: " + versionText_TextBox.Text + newLine;
            }

            // productVersion
            if (productVersionRequired)
            {
                outputString = outputString + "Version: " + productVersion_ComboBox.Text + newLine;
            }

            // productUpdate
            if (productUpdateRequired)
            {
                outputString = outputString + "Update: " + productUpdate_TextBox.Text + newLine;
            }

            // os
            if (osRequired)
            {
                outputString = outputString + "Operating System: " + os_ComboBox.Text + newLine;
            }

            // browser
            if (browserRequired)
            {
                outputString = outputString + "Browser: " + browser_ComboBox.Text + newLine;
            }

            // browserVersion
            if (browserVersionRequired)
            {
                outputString = outputString + "Browser Version: " + browserVersion_TextBox.Text + newLine;
            }

            // sql
            if (sqlRequired)
            {
                outputString = outputString + "SQL Version: " + sql_ComboBox.Text + newLine;
            }

            // office
            if (officeRequired)
            {
                outputString = outputString + "Office Version: " + office_ComboBox.Text + newLine;
            }

            // other
            if (otherRequired)
            {
                outputString = outputString + "Other:" + newLine + office_ComboBox.Text + newLine;
            }

            // accountName
            if (accountNameRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Account Name:" + newLine + accountName_TextBox.Text + newLine;
            }

            // apcInfo
            if (apcInfoRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "APC Account Info:" + newLine
                    + "URL: " + url_TextBox.Text + newLine
                    + "Database: " + databaseName_TextBox.Text + newLine
                    + "Remote Database: " + remoteDatabaseName_TextBox.Text + newLine;
            }

            // reasonForCall
            if (reasonForCallRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Reason for call:" + newLine + reasonForCall_TextBox.Text + newLine;
            }

            // errorMessages
            if (errorMessagesRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Error messages:" + newLine + errorMessages_TextBox.Text + newLine;
            }

            // additionalInformation
            if (additionalInformationRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Additional Issue/Query Information:" + newLine + errorMessages_TextBox.Text + newLine;
            }

            // stepsTaken
            if (stepsTakenRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Steps taken:" + newLine + stepsTaken_TextBox.Text + newLine;
            }

            // resolution
            if (resolutionRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Resolution:" + newLine + resolution_TextBox.Text + newLine;
            }

            outputString = outputString + newLine + "===" + newLine + "Duration: " + calculateCallDuration();

            return outputString;
        }

        private void exit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void openHowToWindow(object sender, RoutedEventArgs e)
        {
            HowTo howToWindow = new HowTo();
            howToWindow.Show();
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
