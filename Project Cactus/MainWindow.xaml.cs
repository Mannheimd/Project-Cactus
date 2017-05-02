using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Navigation;
using System.Windows.Threading;
using System.Xml;

namespace Cactus
{
    public partial class MainWindow : Window
    {
        // Application ID
        private const string applicationId = "8d378bfa-ad9a-4613-9b3d-d86c7b404c6c";

        // Values for ensuring only one application can run at a time
        private readonly Semaphore instancesAllowedSemaphore = new Semaphore(1, 1, applicationId);
        private bool isApplicationRunning { get; set; }

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

        // String for user's AppData folder
        string appdata = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Swiftpage Support\Cactus";

        // Bool for notes backup loop
        bool isNotesBackupRunning = false;

        // Strings for radio buttons
        string officeArchitecture_Radio_Text = null;

        // Timer for notes backup
        DispatcherTimer backupNotesTimer = new DispatcherTimer();

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
        bool officeArchitectureRequired = false;
        bool otherRequired = false;
        bool errorMessagesRequired = true;
        bool additionalInformationRequired = true;
        bool stepsTakenRequired = true;
        bool apcInfoRequired = false;

        // Bools for things in the Results section - for the love of God, Chris, change this shit. As in, like, yesterday. This code sucks ass. And not the Donkey kind. Though that would still be pretty bad.
        bool transferredRequired = false;
        bool resolutionRequired = false;
        bool nextStepsRequired = false;
        bool cloudOpsEscalationRequired = false;
        bool emarketingTechnicalEscalationRequired = false;
        bool emarketingBillingEscalationRequired = false;
        bool emarketingCancellationEscalationRequired = false;
        bool actTechnicalEscalationRequired = false;

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
        bool officeArchitectureMandatory = false;
        bool otherMandatory = false;
        bool errorMessagesMandatory = false;
        bool additionalInformationMandatory = false;
        bool stepsTakenMandatory = true;
        bool resolutionMandatory = false;
        bool resultMandatory = true;
        bool urlMandatory = false;
        bool databaseNameMandatory = false;
        bool nextStepsMandatory = false;
        bool transferredMandatory = false;

        public MainWindow()
        {
            if (startupTasks())
            {
                InitializeComponent();

                // Load the product list from the configuration XML
                loadProductList();
            }
            else
            {
                Application.Current.Shutdown();
            }
        }

        private bool startupTasks()
        {
            // Load the configuration XML
            if (!loadConfigurationXml())
            {
                MessageBox.Show("Unable to load configuration file. Application will terminate.");
                return false;
            }

            // Check that this is the only running instance
            if (!checkSingleInstance())
            {
                MessageBox.Show("Another instance is already running. Application will terminate.");
                return false;
            }

            // If application didn't close safely, offer the notes log
            if (checkRegRunningState())
            {
                offerNotesLog();
            }

            // Set the running state in the registry
            if (!setRegRunningState(true))
            {
                MessageBox.Show("Unable to set RunningState value in registry. Application will run, however will not auto-load notes after a crash."
                    + "\n\n"
                    + @"You can access the notes log from %AppData%\Swiftpage Support\Cactus\notesLog.txt");
            }

            return true;
        }

        private bool checkRegRunningState()
        {
            return getValueFromRegistry(@"Software\Swiftpage Support\Cactus", "cactusRunState");
        }

        private bool setRegRunningState(bool state)
        {
            if (setValueInRegistry(@"Software\Swiftpage Support\Cactus", "cactusRunState", state))
                return true;
            else return false;
        }

        private bool checkOrCreateRegPath()
        {
            // Check if SubKey HKCU\Software\Swiftpage Support\Cactus exists
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Swiftpage Support\Cactus", false);
            if (key == null)
            {
                // Doesn't exist, let's see if HKCU\Software\Swiftpage Support exists
                key = Registry.CurrentUser.OpenSubKey(@"Software\Swiftpage Support", false);
                if (key == null)
                {
                    // Doesn't exist, try to create 'Swiftpage Support' SubKey
                    key = Registry.CurrentUser.OpenSubKey(@"Software", true);
                    try
                    {
                        key.CreateSubKey("Swiftpage Support");
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show(@"Unable to create SubKey HKCU\Software\Swiftpage Support:\n\n" + error.Message);
                        return false;
                    }
                }

                // 'Swiftpage Support' subkey exists (or has just been created), try creating 'Cactus'
                key = Registry.CurrentUser.OpenSubKey(@"Software\Swiftpage Support", true);
                try
                {
                    key.CreateSubKey("Cactus");
                }
                catch (Exception error)
                {
                    MessageBox.Show(@"Unable to create SubKey HKCU\Software\Swiftpage Support\Cactus:\n\n" + error.Message);
                    return false;
                }
            }
            return true;
        }

        private bool getValueFromRegistry(string path, string valueName)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(path, false);
            if (key != null)
            {
                try
                {
                    return Convert.ToBoolean((string)key.GetValue(valueName));
                }
                catch
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private bool setValueInRegistry(string path, string valueName, bool value)
        {
            if (checkOrCreateRegPath())
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(path, true);
                if (key != null)
                {
                    try
                    {
                        key.SetValue(valueName, value.ToString());

                        if ((string)key.GetValue(valueName) == value.ToString())
                        {
                            return true;
                        }
                        else return false;
                    }
                    catch(Exception error)
                    {
                        MessageBox.Show("Unable to set application run state:\n\n" + error.Message);
                        return false;
                    }
                }
                else return false;
            }
            else return false;
        }

        private void offerNotesLog()
        {
            // Offer the notes log to the user due to application not closing correctly
            if (MessageBox.Show("It looks like the application did not close correctly last time it ran. Would you like to view the backup of your notes?", "Improper Close", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                try
                {
                    Process.Start(appdata + @"\notesLog.txt");
                }
                catch(Exception error)
                {
                    MessageBox.Show("Unable to open notes log. File may not exist, or there may be a problem opening it. You can access it manually here:\n"
                        + @"%AppData%\Swiftpage Support\Cactus\notesLog.txt"
                        + "\n\nError:\n" + error.Message);
                }
            }
        }

        private void openNotesLog(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(appdata + @"\notesLog.txt");
            }
            catch (Exception error)
            {
                MessageBox.Show("Unable to open notes log. File may not exist, or there may be a problem opening it. You can access it manually here:\n"
                    + @"%AppData%\Swiftpage Support\Cactus\notesLog.txt"
                    + "\n\nError:\n" + error.Message);
            }
        }

        private bool saveNotesLog()
        {
            string notesLogPath = appdata + @"\notesLog.txt";

            if (!File.Exists(notesLogPath))
            {
                try
                {
                    if (!Directory.Exists(appdata))
                    {
                        Directory.CreateDirectory(appdata);
                    }

                    File.Create(notesLogPath).Close();
                }
                catch (Exception error)
                {
                    MessageBox.Show("Failed to create file '" + notesLogPath + "'\n\n" + error.Message);
                    return false;
                }
            }

            if (File.Exists(notesLogPath))
            {
                try
                {
                    string output = "Saved at " + DateTime.Now.ToString() + Environment.NewLine +
                        Environment.NewLine +
                        buildTicketOutput();

                    File.WriteAllText(notesLogPath, output);
                    return true;
                }
                catch(Exception error)
                {
                    MessageBox.Show("Unable to save backup notes file.\n\n" + error.Message);
                    return false;
                }
            }

            MessageBox.Show("Application fell outside of expected code path whilst saving backup notes file.");
            return false;
        }

        private bool checkSingleInstance()
        {
            if (instancesAllowedSemaphore.WaitOne(TimeSpan.Zero))
            {
                isApplicationRunning = true;
                return true;
            }
            else return false;
        }

        private bool loadConfigurationXml()
        {
            // Loads the configuration XML from embedded resources. Later update will also store this locally and check a server for an updated version.
            try
            {
                string tempXml;

                using (Stream stream = this.GetType().Assembly.GetManifestResourceStream("Cactus.Configuration.xml"))
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

                // Checking if SelectedItem is null - this is to combat occasional Object Reference errors when changing drop-down boxes
                if (selectedItem != null)
                {
                    try
                    {
                        setRequiredRows(selectedItem);
                        checkMandatoryCriteriaMet(true);
                        startTimer();
                    }
                    catch(Exception error)
                    {
                        MessageBox.Show("Error occurred after changing selection on Product ComboBox:\n\n" + error.Message);
                    }
                }
            }
            else
            {
                // Trigger SetRequired Rows with "nothing" selected, to get rid of all rows
                setRequiredRows(null);
            }
        }

        private void resultComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            escalationType_ComboBox.SelectedIndex = -1;
            if (e.AddedItems.Count > 0)
            {
                string selectedItem = (e.AddedItems[0] as ComboBoxItem).Content.ToString();
                // Checking if SelectedItem is null - this is to combat occasional Object Reference errors when changing drop-down boxes
                if (selectedItem != null)
                {
                    startTimer();
                    try
                    {
                        setRequiredResultsRows(selectedItem);
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show("Error occurred after changing selection on Result ComboBox:\n\n" + error.Message);
                    }
                }
            }
            else
            {
                setRequiredResultsRows(null);
            }
        }

        private void escalationComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                string selectedItem = (e.AddedItems[0] as ComboBoxItem).Content.ToString();

                // Checking if SelectedItem is null - this is to combat occasional Object Reference errors when changing drop-down boxes
                if (selectedItem != null)
                {
                    try
                    {
                        setEscalationTypeRows(selectedItem);
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show("Error occurred after changing selection on Result ComboBox:\n\n" + error.Message);
                    }
                }
            }
            else
            {
                setEscalationTypeRows(null);
            }
        }

        private void officeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                string selectedItem = (sender as ComboBox).SelectedItem.ToString();

                // Checking if SelectedItem is null - this is to combat occasional Object Reference errors when changing drop-down boxes
                if (selectedItem != null)
                {
                    switch (selectedItem)
                    {
                        case "N/A":
                        case "Unsupported Version":
                            officeArchitecture_Row.Visibility = Visibility.Collapsed;
                            officeArchitectureRequired = false;
                            officeArchitectureMandatory = false;

                            break;

                        default:
                            officeArchitecture_Row.Visibility = Visibility.Visible;
                            officeArchitectureRequired = true;
                            officeArchitectureMandatory = true;

                            break;
                    }
                }
            }
            else
            {
                officeArchitecture_Row.Visibility = Visibility.Collapsed;
                officeArchitectureRequired = false;
                officeArchitectureMandatory = false;
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

                // Work on other in the XML, if it exists for the selected product
                if (configurationXml.SelectSingleNode(productPath + "/othertext") != null)
                {
                    // Enable other row
                    other_Row.Visibility = Visibility.Visible;

                    // Check if other field is mandatory
                    XmlNode otherNode = configurationXml.SelectSingleNode(productPath + "/othertext");
                    if (otherNode.Attributes["required"].Value == "true")
                    {
                        otherMandatory = true;
                    }
                    else
                    {
                        otherMandatory = false;
                    }

                    // Flag other field as required in output
                    otherRequired = true;
                }
                else
                {
                    // Collapse the other row
                    other_Row.Visibility = Visibility.Collapsed;

                    // Change other field to not mandatory
                    otherMandatory = false;

                    // Flag other as not required in output
                    otherRequired = false;
                }
            }
            catch(Exception error)
            {
                MessageBox.Show("Unable to update context-based options: \n\n" + error.Message);
            }
        }

        private void setRequiredResultsRows(string selectedResult)
        {
            // Crappy way to handle call result selection
            switch (selectedResult)
            {
                case "Resolved":
                    resolution_Grid.Visibility = Visibility.Visible;
                    nextSteps_Grid.Visibility = Visibility.Collapsed;
                    escalation_Grid.Visibility = Visibility.Collapsed;
                    transfer_Grid.Visibility = Visibility.Collapsed;

                    transferredRequired = false;
                    resolutionRequired = true;
                    nextStepsRequired = false;

                    resolutionMandatory = true;
                    nextStepsMandatory = false;
                    transferredMandatory = false;

                    break;

                case "Next steps":
                    resolution_Grid.Visibility = Visibility.Collapsed;
                    nextSteps_Grid.Visibility = Visibility.Visible;
                    escalation_Grid.Visibility = Visibility.Collapsed;
                    transfer_Grid.Visibility = Visibility.Collapsed;

                    transferredRequired = false;
                    resolutionRequired = false;
                    nextStepsRequired = true;

                    resolutionMandatory = false;
                    nextStepsMandatory = true;
                    transferredMandatory = false;

                    break;

                case "Escalated":
                    resolution_Grid.Visibility = Visibility.Collapsed;
                    nextSteps_Grid.Visibility = Visibility.Collapsed;
                    escalation_Grid.Visibility = Visibility.Visible;
                    transfer_Grid.Visibility = Visibility.Collapsed;

                    transferredRequired = false;
                    resolutionRequired = false;
                    nextStepsRequired = false;

                    resolutionMandatory = false;
                    nextStepsMandatory = false;
                    transferredMandatory = false;

                    break;

                case "Transferred":
                    resolution_Grid.Visibility = Visibility.Collapsed;
                    nextSteps_Grid.Visibility = Visibility.Collapsed;
                    escalation_Grid.Visibility = Visibility.Collapsed;
                    transfer_Grid.Visibility = Visibility.Visible;

                    transferredRequired = true;
                    resolutionRequired = false;
                    nextStepsRequired = false;

                    resolutionMandatory = false;
                    nextStepsMandatory = false;
                    transferredMandatory = true;

                    break;

                default:
                    resolution_Grid.Visibility = Visibility.Collapsed;
                    nextSteps_Grid.Visibility = Visibility.Collapsed;
                    escalation_Grid.Visibility = Visibility.Collapsed;
                    transfer_Grid.Visibility = Visibility.Collapsed;

                    transferredRequired = false;
                    resolutionRequired = false;
                    nextStepsRequired = false;

                    resolutionMandatory = false;
                    nextStepsMandatory = false;
                    transferredMandatory = false;

                    break;
            }
        }

        private void setEscalationTypeRows(string selectedEscalationType)
        {
            // Crappy way to handle escalation type selection
            switch (selectedEscalationType)
            {
                case "Act! Premium Cloud Operations":
                    cloudOpsEscalation_Grid.Visibility = Visibility.Visible;
                    emarketingTechnicalEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingBillingEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingCancellationEscalation_Grid.Visibility = Visibility.Collapsed;

                    cloudOpsEscalationRequired = true;
                    emarketingTechnicalEscalationRequired = false;
                    emarketingBillingEscalationRequired = false;
                    emarketingCancellationEscalationRequired = false;
                    actTechnicalEscalationRequired = false;

                    break;

                case "Emarketing - Technical":
                    cloudOpsEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingTechnicalEscalation_Grid.Visibility = Visibility.Visible;
                    emarketingBillingEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingCancellationEscalation_Grid.Visibility = Visibility.Collapsed;

                    cloudOpsEscalationRequired = false;
                    emarketingTechnicalEscalationRequired = true;
                    emarketingBillingEscalationRequired = false;
                    emarketingCancellationEscalationRequired = false;
                    actTechnicalEscalationRequired = false;

                    break;

                case "Emarketing - Billing":
                    cloudOpsEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingTechnicalEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingBillingEscalation_Grid.Visibility = Visibility.Visible;
                    emarketingCancellationEscalation_Grid.Visibility = Visibility.Collapsed;

                    cloudOpsEscalationRequired = false;
                    emarketingTechnicalEscalationRequired = false;
                    emarketingBillingEscalationRequired = true;
                    emarketingCancellationEscalationRequired = false;
                    actTechnicalEscalationRequired = false;

                    break;

                case "Emarketing - Cancellation":
                    cloudOpsEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingTechnicalEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingBillingEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingCancellationEscalation_Grid.Visibility = Visibility.Visible;

                    cloudOpsEscalationRequired = false;
                    emarketingTechnicalEscalationRequired = false;
                    emarketingBillingEscalationRequired = false;
                    emarketingCancellationEscalationRequired = true;
                    actTechnicalEscalationRequired = false;

                    break;

                case "Act! - Technical":
                    cloudOpsEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingTechnicalEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingBillingEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingCancellationEscalation_Grid.Visibility = Visibility.Collapsed;

                    cloudOpsEscalationRequired = false;
                    emarketingTechnicalEscalationRequired = false;
                    emarketingBillingEscalationRequired = false;
                    emarketingCancellationEscalationRequired = false;
                    actTechnicalEscalationRequired = true;

                    break;

                default:
                    cloudOpsEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingTechnicalEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingBillingEscalation_Grid.Visibility = Visibility.Collapsed;
                    emarketingCancellationEscalation_Grid.Visibility = Visibility.Collapsed;

                    cloudOpsEscalationRequired = false;
                    emarketingTechnicalEscalationRequired = false;
                    emarketingBillingEscalationRequired = false;
                    emarketingCancellationEscalationRequired = false;
                    actTechnicalEscalationRequired = false;

                    break;
            }
        }

        private void startTimerButton_Click(object sender, RoutedEventArgs e)
        {
            startTimer();
        }

        private void endTimerButton_Click(object sender, RoutedEventArgs e)
        {
            endTimer();
        }

        private void resetFormButton_Click(object sender, RoutedEventArgs e)
        {
            resetForm();
        }

        private void resetForm()
        {
            if (MessageBox.Show("This cannot be undone. Are you sure you wish to proceed?", "Form reset confirmation", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
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
                officeArchitecture_RadioButton_32bit.IsChecked = false;
                officeArchitecture_RadioButton_64bit.IsChecked = false;
                officeArchitecture_Radio_Text = null;
                other_TextBox.Text = null;
                errorMessages_TextBox.Text = null;
                stepsTaken_TextBox.Text = null;
                additionalInformation_TextBox.Text = null;
                url_TextBox.Text = null;
                databaseName_TextBox.Text = null;
                remoteDatabaseName_TextBox.Text = null;

                // Just look at how crap this is
                callResult_ComboBox.SelectedIndex = -1;
                escalationType_ComboBox.SelectedIndex = -1;
                emarketingTechnicalEscalation_Platform_ComboBox.SelectedIndex = -1;
                emarketingBillingEscalation_Platform_ComboBox.SelectedIndex = -1;
                emarketingCancellationEscalation_Platform_ComboBox.SelectedIndex = -1;
                resolution_TextBox.Text = null;
                resolution_TextBox.Text = null;
                nextSteps_TextBox.Text = null;
                cloudOpsEscalation_PartnerName_TextBox.Text = null;
                cloudOpsEscalation_PartnerAccount_TextBox.Text = null;
                cloudOpsEscalation_AffectedUsers_TextBox.Text = null;
                cloudOpsEscalation_PrimaryAccountEmail_TextBox.Text = null;
                cloudOpsEscalation_ContactEmail_TextBox.Text = null;
                cloudOpsEscalation_Symptoms_TextBox.Text = null;
                cloudOpsEscalation_HowToReplicate_TextBox.Text = null;
                cloudOpsEscalation_RequestedAction_TextBox.Text = null;
                emarketingTechnicalEscalation_AgentName_TextBox.Text = null;
                emarketingTechnicalEscalation_AccountName_TextBox.Text = null;
                emarketingTechnicalEscalation_CustomerPhoneNumber_TextBox.Text = null;
                emarketingTechnicalEscalation_CustomerEmailAddress_TextBox.Text = null;
                emarketingTechnicalEscalation_UserID_TextBox.Text = null;
                emarketingTechnicalEscalation_Server_TextBox.Text = null;
                emarketingTechnicalEscalation_Integration_TextBox.Text = null;
                emarketingTechnicalEscalation_Symptoms_TextBox.Text = null;
                emarketingTechnicalEscalation_DateTime_TextBox.Text = null;
                emarketingTechnicalEscalation_ActionsAtTime_TextBox.Text = null;
                emarketingTechnicalEscalation_ReplicatedOn_TextBox.Text = null;
                emarketingTechnicalEscalation_HowToReplicate_TextBox.Text = null;
                emarketingTechnicalEscalation_AccountOrGlobal_TextBox.Text = null;
                emarketingBillingEscalation_AccountName_TextBox.Text = null;
                emarketingBillingEscalation_AccountEmail_TextBox.Text = null;
                emarketingBillingEscalation_ContactEmail_TextBox.Text = null;
                emarketingBillingEscalation_ContactPhone_TextBox.Text = null;
                emarketingBillingEscalation_ReasonForEscalation_TextBox.Text = null;
                emarketingCancellationEscalation_AccountName_TextBox.Text = null;
                emarketingCancellationEscalation_AccountEmail_TextBox.Text = null;
                emarketingCancellationEscalation_AccountOwnerContactEmail_TextBox.Text = null;
                emarketingCancellationEscalation_AccountOwnerContactPhone_TextBox.Text = null;
                emarketingCancellationEscalation_ContactEmail_TextBox.Text = null;
                emarketingCancellationEscalation_ContactPhone_TextBox.Text = null;
                emarketingCancellationEscalation_Issue_TextBox.Text = null;
                emarketingCancellationEscalation_ReasonForEscalation_TextBox.Text = null;
                transferDepartment_TextBox.Text = null;
                transferPerson_TextBox.Text = null;

                // Call mandatory field updater with 'Reset' flag set true
                checkMandatoryCriteriaMet(true);

                // Stop backup log
                backupNotesTimer.Stop();
            }
        }

        private void copyToClipboard_Button_Click(object sender, RoutedEventArgs e)
        {
            if (checkMandatoryCriteriaMet(false))
            {
                try
                {
                    Clipboard.SetDataObject(buildTicketOutput());
                }
                catch(Exception error)
                {
                    MessageBox.Show("Unable to copy to clipboard. This can be caused by a an active WebEx session interfering with clipboard operations. Try again after closing your WebEx session.\n\n" + error.Message);
                }
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

        private bool checkMandatoryCriteriaMet(bool reset)
        {
            // For every field marked as mandatory, check if it's associated UI element is blank.
            // If mandatory field is empty, make its row orange (Swiftpage Orange, hell yah) and flag failure.
            // if all is well, remove the colour.
            // If statement reqires Reset bool to be false, otherwise all mantatory statuses are reset
            bool criteriaMet = true;

            // reasonForCall
            if (reasonForCallMandatory & reasonForCall_TextBox.Text == "" & !reset)
            {
                reasonForCall_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                reasonForCall_Grid.ClearValue(BackgroundProperty);
            }

            // product
            if (productMandatory & product_ComboBox.Text == "" & !reset)
            {
                product_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254,80,0)));
                criteriaMet = false;
            }
            else
            {
                product_Row.ClearValue(BackgroundProperty);
            }

            // productName
            if (productNameMandatory & productName_TextBox.Text == "" & !reset)
            {
                productName_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                productName_Row.ClearValue(BackgroundProperty);
            }

            // versionText
            if (versionTextMandatory & versionText_TextBox.Text == "" & !reset)
            {
                versionText_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                versionText_Row.ClearValue(BackgroundProperty);
            }

            // productVersion
            if (productVersionMandatory & productVersion_ComboBox.Text == "" & !reset)
            {
                productVersion_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                productVersion_Row.ClearValue(BackgroundProperty);
            }

            // productUpdate
            if (productUpdateMandatory & productUpdate_TextBox.Text == "" & !reset)
            {
                productUpdate_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                productUpdate_Row.ClearValue(BackgroundProperty);
            }

            // os
            if (osMandatory & os_ComboBox.Text == "" & !reset)
            {
                os_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                os_Row.ClearValue(BackgroundProperty);
            }

            // browser
            if (browserMandatory & browser_ComboBox.Text == "" & !reset)
            {
                browser_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                browser_Row.ClearValue(BackgroundProperty);
            }

            // browserVersion
            if (browserVersionMandatory & browserVersion_TextBox.Text == "" & !reset)
            {
                browserVersion_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                browserVersion_Row.ClearValue(BackgroundProperty);
            }

            // accountName
            if (accountNameMandatory & accountName_TextBox.Text == "" & !reset)
            {
                accountName_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                accountName_Row.ClearValue(BackgroundProperty);
            }

            // sql
            if (sqlMandatory & sql_ComboBox.Text == "" & !reset)
            {
                sql_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                sql_Row.ClearValue(BackgroundProperty);
            }

            // office
            if (officeMandatory & office_ComboBox.Text == "" & !reset)
            {
                office_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                office_Row.ClearValue(BackgroundProperty);
            }

            // officeArchitecture
            if (officeArchitectureMandatory & officeArchitecture_Radio_Text == null & !reset)
            {
                officeArchitecture_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                officeArchitecture_Row.ClearValue(BackgroundProperty);
            }

            // other
            if (otherMandatory & other_TextBox.Text == "" & !reset)
            {
                other_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                other_Row.ClearValue(BackgroundProperty);
            }

            // errorMessages
            if (errorMessagesMandatory & errorMessages_TextBox.Text == "" & !reset)
            {
                errorMessages_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                errorMessages_Grid.ClearValue(BackgroundProperty);
            }

            // additionalInformation
            if (additionalInformationMandatory & additionalInformation_TextBox.Text == "" & !reset)
            {
                additionalInformation_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                additionalInformation_Grid.ClearValue(BackgroundProperty);
            }

            // stepsTaken
            if (stepsTakenMandatory & stepsTaken_TextBox.Text == "" & !reset)
            {
                stepsTaken_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                stepsTaken_Grid.ClearValue(BackgroundProperty);
            }

            // resolution
            if (resolutionMandatory & resolution_TextBox.Text == "" & !reset)
            {
                resolution_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                resolution_Grid.ClearValue(BackgroundProperty);
            }

            // url
            if (urlMandatory & url_TextBox.Text == "" & !reset)
            {
                url_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                url_Row.ClearValue(BackgroundProperty);
            }

            // databaseName
            if (databaseNameMandatory & databaseName_TextBox.Text == "" & !reset)
            {
                databaseName_Row.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                databaseName_Row.ClearValue(BackgroundProperty);
            }

            // result
            if (resultMandatory & callResult_ComboBox.Text == "" & !reset)
            {
                callResult_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                callResult_Grid.ClearValue(BackgroundProperty);
            }

            // nextSteps
            if (nextStepsMandatory & nextSteps_TextBox.Text == "" & !reset)
            {
                nextSteps_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                nextSteps_Grid.ClearValue(BackgroundProperty);
            }

            // transferred
            if (transferredMandatory & transferDepartment_TextBox.Text == "" & transferPerson_TextBox.Text == "" & !reset)
            {
                transferDepartment_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                transferPerson_Grid.SetValue(BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 80, 0)));
                criteriaMet = false;
            }
            else
            {
                transferDepartment_Grid.ClearValue(BackgroundProperty);
                transferPerson_Grid.ClearValue(BackgroundProperty);
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
                outputString = outputString + "Office Version: " + office_ComboBox.Text;
                if (officeArchitectureRequired)
                {
                    outputString = outputString + " " + officeArchitecture_Radio_Text;
                }
                outputString = outputString + newLine;
            }

            // other
            if (otherRequired)
            {
                outputString = outputString + "Other:" + newLine + other_TextBox.Text + newLine;
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
                outputString = outputString + newLine + "===" + newLine + "Additional Issue/Query Information:" + newLine + additionalInformation_TextBox.Text + newLine;
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

            // transfered
            if (transferredRequired)
            {
                outputString = outputString + newLine + "===" + newLine
                    + "Transfer Department: " + transferDepartment_TextBox.Text + newLine
                    + "Transfer Person: " + transferPerson_TextBox.Text + newLine;
            }

            // nextSteps
            if (nextStepsRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Next Steps: " + newLine + nextSteps_TextBox.Text + newLine;
            }

            // cloudOpsEscalation
            if (cloudOpsEscalationRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Escalation:" + newLine
                    + "Type: CloudOps" + newLine
                    + "Partner name: " + cloudOpsEscalation_PartnerName_TextBox.Text + newLine
                    + "Partner account: " + cloudOpsEscalation_PartnerAccount_TextBox.Text + newLine
                    + "Affected users: " + cloudOpsEscalation_AffectedUsers_TextBox.Text + newLine
                    + "Primary account email: " + cloudOpsEscalation_PrimaryAccountEmail_TextBox.Text + newLine
                    + "Contact email: " + cloudOpsEscalation_ContactEmail_TextBox.Text + newLine
                    + "Symptoms: " + cloudOpsEscalation_Symptoms_TextBox.Text + newLine
                    + "How to replicate: " + cloudOpsEscalation_HowToReplicate_TextBox.Text + newLine
                    + "Requested action: " + cloudOpsEscalation_RequestedAction_TextBox.Text + newLine;
            }

            // emarketingTechnicalEscalation
            if (emarketingTechnicalEscalationRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Escalation:" + newLine
                    + "Type: Emarketing Technical" + newLine
                    + "Agent Name: " + emarketingTechnicalEscalation_AgentName_TextBox.Text + newLine
                    + "Account name: " + emarketingTechnicalEscalation_AccountName_TextBox.Text + newLine
                    + "Customer phone number: " + emarketingTechnicalEscalation_CustomerPhoneNumber_TextBox.Text + newLine
                    + "Customer email address: " + emarketingTechnicalEscalation_CustomerEmailAddress_TextBox.Text + newLine
                    + "User ID: " + emarketingTechnicalEscalation_UserID_TextBox.Text + newLine
                    + "Server: " + emarketingTechnicalEscalation_Server_TextBox.Text + newLine
                    + "Integration: " + emarketingTechnicalEscalation_Integration_TextBox.Text + newLine
                    + "Platform: " + emarketingTechnicalEscalation_Platform_ComboBox.Text + newLine
                    + "Symptoms: " + emarketingTechnicalEscalation_Symptoms_TextBox.Text + newLine
                    + "Date and time of occurrence:" + emarketingTechnicalEscalation_DateTime_TextBox.Text + newLine
                    + "What happened when issue occurred: " + emarketingTechnicalEscalation_ActionsAtTime_TextBox.Text + newLine
                    + "Replicated with: " + emarketingTechnicalEscalation_ReplicatedOn_TextBox.Text + newLine
                    + "How to reproduce: " + emarketingTechnicalEscalation_HowToReplicate_TextBox.Text + newLine
                    + "Global or Account Specific: " + emarketingTechnicalEscalation_AccountOrGlobal_TextBox.Text + newLine;
            }

            // emarketingBillingEscalation
            if (emarketingBillingEscalationRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Escalation:" + newLine
                    + "Type: Emarketing Billing" + newLine
                    + "Account name: " + emarketingBillingEscalation_AccountName_TextBox.Text + newLine
                    + "Account email: " + emarketingBillingEscalation_AccountEmail_TextBox.Text + newLine
                    + "Contact email: " + emarketingBillingEscalation_ContactEmail_TextBox.Text + newLine
                    + "Contact phone: " + emarketingBillingEscalation_ContactPhone_TextBox.Text + newLine
                    + "Platform: " + emarketingBillingEscalation_Platform_ComboBox.Text + newLine
                    + "Reason for escalation: " + emarketingBillingEscalation_ReasonForEscalation_TextBox.Text + newLine;
            }

            // emarketingCancellationEscalation
            if (emarketingCancellationEscalationRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Escalation:" + newLine
                    + "Type: Emarketing Cancellation" + newLine
                    + "Account name: " + emarketingCancellationEscalation_AccountName_TextBox.Text + newLine
                    + "Account email: " + emarketingCancellationEscalation_AccountEmail_TextBox.Text + newLine
                    + "Account owner contact email: " + emarketingCancellationEscalation_AccountOwnerContactEmail_TextBox.Text + newLine
                    + "Contact owner contact phone: " + emarketingCancellationEscalation_AccountOwnerContactPhone_TextBox.Text + newLine
                    + "Contact email: " + emarketingCancellationEscalation_ContactEmail_TextBox.Text + newLine
                    + "Contact phone: " + emarketingCancellationEscalation_ContactPhone_TextBox.Text + newLine
                    + "Platform: " + emarketingCancellationEscalation_Platform_ComboBox.Text + newLine
                    + "Description of issue: " + emarketingCancellationEscalation_Issue_TextBox.Text + newLine
                    + "Reason for cancellation request: " + emarketingCancellationEscalation_ReasonForEscalation_TextBox.Text + newLine;
            }

            // actTechnicalEscalation
            if (actTechnicalEscalationRequired)
            {
                outputString = outputString + newLine + "===" + newLine + "Escalation:" + newLine
                    + "Type: Act! Technical" + newLine;
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

        private void openAboutWindow(object sender, RoutedEventArgs e)
        {
            About aboutWindow = new About();
            aboutWindow.Show();
        }

        private void openSixSquareGridWindow(object sender, RoutedEventArgs e)
        {
            SixSquareGrid window = new SixSquareGrid();
            window.Show();
        }

        private void openHelpdeskChecklistWindow(object sender, RoutedEventArgs e)
        {
            HelpdeskChecklist window = new HelpdeskChecklist();
            window.Show();
        }

        private void primaryFieldsEdited(object sender, RoutedEventArgs e)
        {
            startTimer();
        }

        private void startTimer()
        {
            if (!isTimerRunning)
            {
                isTimerRunning = true;

                // Enable/disable buttons
                endTimer_Button.IsEnabled = true;
                startTimer_Button.IsEnabled = false;
                resetForm_Button.IsEnabled = false;
                copyToClipboard_Button.IsEnabled = false;

                // Save current time and call duration, start counting new duration
                startTime = DateTime.Now;
                finalCallDuration = finalCallDuration + callDuration;
                callDuration = new TimeSpan();

                // Start on-screen counting timer
                durationCounter.Tick += new EventHandler(durationCounter_Tick);
                durationCounter.Interval = new TimeSpan(0, 0, 1);
                durationCounter.Start();

                // Start taking notes backups
                if (!isNotesBackupRunning)
                {
                    isNotesBackupRunning = true;

                    backupNotesTimer.Tick += new EventHandler(backupNotes_Tick);
                    backupNotesTimer.Interval = new TimeSpan(0, 0, 30);
                    backupNotesTimer.Start();
                }
            }
        }

        private void endTimer()
        {
            if (isTimerRunning)
            {
                isTimerRunning = false;

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
                isNotesBackupRunning = false;
            }
        }

        private void backupNotes_Tick(object sender, EventArgs e)
        {
            saveNotesLog();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void application_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            saveNotesLog();
            backupNotesTimer.Stop();
            isNotesBackupRunning = false;
            setRegRunningState(false);
            if (isApplicationRunning)
            {
                instancesAllowedSemaphore.Release();
            }
        }

        private void officeArchitecture_RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if ((sender as RadioButton).IsChecked == true)
            {
                officeArchitecture_Radio_Text = (sender as RadioButton).Content.ToString();
            }
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
