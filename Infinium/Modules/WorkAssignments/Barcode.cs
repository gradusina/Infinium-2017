using System;
using System.Drawing;

namespace Infinium.Modules.WorkAssignments
{
    public class Barcode
    {
        private BarcodeLib.Barcode Barcod;

        private SolidBrush FontBrush;

        public Barcode()
        {
            Barcod = new BarcodeLib.Barcode();

            FontBrush = new System.Drawing.SolidBrush(Color.Black);
        }

        public enum BarcodeLength { Short, Medium, Long };

        public void DrawBarcodeText(BarcodeLength BarcodeLength, Graphics Graphics, string Text, int X, int Y)
        {
            int CharOffset = 0;
            int CharWidth = 0;
            float FontSize = 0;

            if (BarcodeLength == Barcode.BarcodeLength.Short)
            {
                CharWidth = 8;
                CharOffset = 7;
                FontSize = 8.0f;
            }
            if (BarcodeLength == Barcode.BarcodeLength.Medium)
            {
                CharWidth = 18;
                CharOffset = 5;
                FontSize = 12.0f;
            }
            if (BarcodeLength == Barcode.BarcodeLength.Long)
            {
                CharWidth = 26;
                CharOffset = 5;
                FontSize = 14.0f;
            }

            Font F = new Font("Arial", FontSize, FontStyle.Regular);

            for (int i = 0; i < Text.Length; i++)
            {
                Graphics.DrawString(Text[i].ToString(), F, FontBrush, i * CharWidth + CharOffset + X, Y + 2);
            }

            F.Dispose();
        }

        public Image GetBarcode(BarcodeLength BarcodeLength, int BarcodeHeight, string Text)
        {
            //set length and height
            if (BarcodeLength == Barcode.BarcodeLength.Short)
            {
                Barcod.Width = 101 + 12;
            }
            if (BarcodeLength == Barcode.BarcodeLength.Medium)
            {
                Barcod.Width = 202 + 12;
            }
            if (BarcodeLength == Barcode.BarcodeLength.Long)
            {
                Barcod.Width = 303 + 12;
            }

            Barcod.Height = BarcodeHeight;

            //create area
            Bitmap B = new Bitmap(Barcod.Width, BarcodeHeight);
            Graphics G = Graphics.FromImage(B);
            G.Clear(Color.White);

            //create barcode
            Image Bar = Barcod.Encode(BarcodeLib.TYPE.CODE128C, Text);
            G.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            G.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None;
            G.DrawImage(Bar, 0, 2);
            //DrawBarcodeText(Barcode.BarcodeLength.Short, G, Text, 0, 23);

            //create text
            G.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit;

            Bar.Dispose();
            G.Dispose();

            GC.Collect();

            return B;
        }
    }
}