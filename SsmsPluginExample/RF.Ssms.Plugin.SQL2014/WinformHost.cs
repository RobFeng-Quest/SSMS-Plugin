using System.Windows.Forms;

namespace RF.Ssms.Plugin.SQL2014
{
    public partial class WinformHost : UserControl
    {
        public WinformHost()
        {
            var ctrlHost = new System.Windows.Forms.Integration.ElementHost();
            ctrlHost.Dock = DockStyle.Fill;
            this.Controls.Add(ctrlHost);

            var wpfControl = new RF.WpfLibrary.WpfUserControl();
            wpfControl.InitializeComponent();
            ctrlHost.Child = wpfControl;
        }
    }
}