using System.Drawing;
using System.Windows.Forms;

namespace BioDraw
{
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter() : base("")
        {
        }

        public static stdole.IPictureDisp ToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }
}
