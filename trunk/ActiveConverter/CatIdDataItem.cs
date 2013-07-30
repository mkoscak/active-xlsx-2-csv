using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using OfficeOpenXml.Drawing;

namespace ActiveConverter
{
    public class CatIdDataItem
    {
        DataItem item;

        public CatIdDataItem(DataItem item)
        {
            this.item = item;
        }

        public Image Image
        {
            get
            {
                if (item != null && item.Picture is ExcelPicture)
                {
                    return ResizeImage((item.Picture as ExcelPicture).Image, new Size(100,100));
                }

                return null;
            }
        }

        public string Description
        {
            get
            {
                return item.Description;
            }
        }

        public string CategoryId
        {
            get
            {
                return item.CatId;
            }

            set
            {
                item.CatId = value;
            }
        }

        internal Image ResizeImage(Image imgToResize, Size size)
        {
            return (Image)(new Bitmap(imgToResize, size));
        }
    }
}
