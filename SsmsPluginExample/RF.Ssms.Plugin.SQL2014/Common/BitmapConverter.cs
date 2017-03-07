using System;
using System.Drawing;
using stdole;
using System.Windows.Forms;

namespace RF.Ssms.Plugin.Common
{
    public class BitmapConverter : AxHost
    {
        public BitmapConverter()
            : base("52D64AAC-29C1-CAC8-BB3A-115F0D3D77CB")
        {
        }

        public static IPictureDisp ToIPicture(System.Drawing.Image image)
        {
            return (IPictureDisp)AxHost.GetIPictureDispFromPicture(image);
        }
    }
}
