using System.Windows.Controls;

namespace RF.WpfLibrary
{
    /// <summary>
    /// Interaction logic for WpfUserControl.xaml
    /// </summary>
    public partial class WpfUserControl : UserControl
    {
        public WpfUserControl()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("Hello world.");
        }
    }
}