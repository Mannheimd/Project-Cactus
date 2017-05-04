using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;

namespace Cactus
{
    public partial class HowTo : Window
    {
        public HowTo()
        {
            InitializeComponent();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }
    }
}
