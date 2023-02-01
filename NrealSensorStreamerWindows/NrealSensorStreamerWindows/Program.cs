namespace NrealSensorStreamerWindows
{
    using DirectShowLib;
    using System;
    using System.Drawing;
    using System.Windows.Forms;
    using System.Windows.Media.Imaging;
    using System.Windows.Media;
    using System.IO;
    using System.Windows.Markup;

    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (!DirectShowImageCapture.TryCreateImageCapture("vid_05a9", MediaSubType.YUY2, 640, 481, 30, out var imageCapture))
            {
                Console.WriteLine("Nreal Light glasses are not currently connected. Press any key to exit.");
                Console.ReadKey();
                return;
            }

            imageCapture.Start();

            using (Form form = new Form())
            {
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Size = new Size(1280, 481);

                var pictureBox = new PictureBox();
                pictureBox.Dock = DockStyle.Fill;
                pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
                form.Controls.Add(pictureBox);

                imageCapture.FrameReady += (sender, image) =>
                {
                    using (MemoryStream outStream = new MemoryStream())
                    {
                        var encoder = new BmpBitmapEncoder();                        
                        encoder.Frames.Add(BitmapFrame.Create(BitmapSource.Create(1280, 481, 96, 96, PixelFormats.Rgb24, BitmapPalettes.Gray256, ConvertYUV222To2xGrayRGB(image, 1280, 481), 1280 * 3)));
                        encoder.Save(outStream);
                        pictureBox.Image = new Bitmap(outStream);
                    }
                };

                form.ShowDialog();
            }

            imageCapture.Stop();
        }

        private static byte[] ConvertYUV222To2xGrayRGB(byte[] yuvData, int width, int height)
        {
            var rgbData = new byte[width * height * 3];
            for (var i = 0; i < yuvData.Length; i += 2)
            {
                var index = (i / 2 * 6);
                for (var j = index; j < index + 6; j++)
                    rgbData[j] = yuvData[i];
            }

            return rgbData;
        }
    }
}
