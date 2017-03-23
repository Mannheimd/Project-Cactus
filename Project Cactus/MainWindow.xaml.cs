using System;
using System.Collections.Generic;
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
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText("Hello!");
        }

        private void actVersionNumber_OtherSelected(object sender, RoutedEventArgs e)
        {
            actVersion_TextBox.IsEnabled = true;
        }

        private void actVersionNumber_NotOtherSelected(object sender, RoutedEventArgs e)
        {
            actVersion_TextBox.IsEnabled = false;
        }

        private void officeVersion_OtherSelected(object sender, RoutedEventArgs e)
        {
            officeVersion_TextBox.IsEnabled = true;
        }

        private void officeVersion_NotOtherSelected(object sender, RoutedEventArgs e)
        {
            officeVersion_TextBox.IsEnabled = false;
        }

        private void windowsVersion_OtherSelected(object sender, RoutedEventArgs e)
        {
            windowsVersion_TextBox.IsEnabled = true;
        }

        private void windowsVersion_NotOtherSelected(object sender, RoutedEventArgs e)
        {
            windowsVersion_TextBox.IsEnabled = false;
        }

        private void reasonForCallInformation_Image_MouseCheck(object sender, MouseEventArgs e)
        {
            if (reasonForCallInformation_Image.IsMouseOver)
            {
                reasonForCallInformation_Text.Visibility = Visibility.Visible;
            }
            else
            {
                reasonForCallInformation_Text.Visibility = Visibility.Hidden;
            }
        }

        private void errorMessagesInformation_Image_MouseCheck(object sender, MouseEventArgs e)
        {
            if (errorMessagesInformation_Image.IsMouseOver)
            {
                errorMessagesInformation_Text.Visibility = Visibility.Visible;
            }
            else
            {
                errorMessagesInformation_Text.Visibility = Visibility.Hidden;
            }
        }

        private void stepsTakenInformation_Image_MouseCheck(object sender, MouseEventArgs e)
        {
            if (stepsTakenInformation_Image.IsMouseOver)
            {
                stepsTakenInformation_Text.Visibility = Visibility.Visible;
            }
            else
            {
                stepsTakenInformation_Text.Visibility = Visibility.Hidden;
            }
        }

        private void resolutionInformation_Image_MouseCheck(object sender, MouseEventArgs e)
        {
            if (resolutionInformation_Image.IsMouseOver)
            {
                resolutionInformation_Text.Visibility = Visibility.Visible;
            }
            else
            {
                resolutionInformation_Text.Visibility = Visibility.Hidden;
            }
        }

        private void windowViewStyleSelector_ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((e.AddedItems[0] as ComboBoxItem).Content as string == "Wide")
            {
                smallScreenTroubleshootingDetailsRow.Height = new GridLength(0);
                bigScreenSecondPageColumn.Width = new GridLength(1, GridUnitType.Star);
                troubleshootingInformationGrid.SetValue(Grid.RowProperty, 0);
                troubleshootingInformationGrid.SetValue(Grid.ColumnProperty, 1);
                troubleshootingInformationGrid.SetValue(Grid.RowSpanProperty, 2);
            }
            else if ((e.AddedItems[0] as ComboBoxItem).Content as string == "Narrow")
            {
                smallScreenTroubleshootingDetailsRow.Height = new GridLength(1, GridUnitType.Star);
                bigScreenSecondPageColumn.Width = new GridLength(0);
                troubleshootingInformationGrid.SetValue(Grid.RowProperty, 2);
                troubleshootingInformationGrid.SetValue(Grid.ColumnProperty, 0);
                troubleshootingInformationGrid.SetValue(Grid.RowSpanProperty, 1);
            }
        }
    }
}
