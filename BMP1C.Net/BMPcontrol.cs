using System.Drawing;
using System.Windows.Forms;

namespace BMP1C.Net
{
    public partial class BmPcontrol : UserControl
    {
        public BmPcontrol()
        {
            InitializeComponent();
            BackColor = Color.Transparent;
            
            SetStyle(ControlStyles.Opaque, true);
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            SetStyle(ControlStyles.UserPaint, true);
        }

    }
}
