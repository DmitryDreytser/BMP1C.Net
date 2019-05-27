using System;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using stdole;

namespace BMP1C.Net
{
    public static class ImageUtilities
    {
        public static IPicture ConvertToIPicture(Image image)
        {
            return ImageOLEConverter.Instance.ConvertToIPicture(image);
        }

        private class ImageOLEConverter : AxHost
        {
            public static readonly ImageOLEConverter Instance = new
            ImageOLEConverter();

            private ImageOLEConverter()
            : base(Guid.Empty.ToString())
            {
            }

            public IPicture ConvertToIPicture(Image image)
            {
                return AxHost.GetIPictureFromPicture(image) as IPicture;
            }
        }
    }
}
