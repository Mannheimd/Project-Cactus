using System;
using System.Reflection;
using System.Windows;

namespace Project_Cactus
{
    /// <summary>
    /// Interaction logic for About.xaml
    /// </summary>
    public partial class About : Window
    {
        // Application version
        public Version version = Assembly.GetExecutingAssembly().GetName().Version;

        public About()
        {
            InitializeComponent();

            // Populate the version number
            about_Version_TextBox.Content = "Version " + version.ToString();
        }
    }
}
