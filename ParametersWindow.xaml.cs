using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;

namespace DocMaker
{
    /// <summary>
    /// Logica di interazione per ParametersWindow.xaml
    /// </summary>
    public partial class ParametersWindow : Window
    {
        public Models.Settings Settings
        {
            get; set;
        }
        public ParametersWindow(Models.Settings _settings)
        {
            Settings = _settings;
            InitializeComponent();
            LoadWindow();
            LoadImage();
        }
        protected void LoadWindow()
        {
            slBorder.Value = Settings.Radius;
            slBorderWidth.Value = Settings.BorderWidth;
            cpBackgroundColor.SelectedColor = System.Windows.Media.Color.FromArgb(Settings.BackgroundColor.A, Settings.BackgroundColor.R, Settings.BackgroundColor.G, Settings.BackgroundColor.B);
            cpBorderColor.SelectedColor = System.Windows.Media.Color.FromArgb(Settings.BorderColor.A, Settings.BorderColor.R, Settings.BorderColor.G, Settings.BorderColor.B);
        }
        protected void LoadImage()
        {
            using Bitmap bmp = new Bitmap((int)100, (int)100);
            using Graphics g = Graphics.FromImage(bmp);
            System.Drawing.Pen borderPen = new System.Drawing.Pen(Settings.BorderColor);
            borderPen.Width = (float)Settings.BorderWidth;
            System.Drawing.Brush brush = new SolidBrush(Settings.BackgroundColor);
            System.Drawing.Point point = new System.Drawing.Point(50, 50);

            g.FillEllipse(brush, point.X - Settings.Radius, point.Y - Settings.Radius,
             Settings.Radius + Settings.Radius, Settings.Radius + Settings.Radius);

            g.DrawEllipse(borderPen, point.X - Settings.Radius, point.Y - Settings.Radius,
            Settings.Radius + Settings.Radius, Settings.Radius + Settings.Radius);

            imgCursor.Source = BitmapToImageSource(bmp);
            Opacity = 1;
        }
        BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapimage = new BitmapImage();
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

                return bitmapimage;
            }
        }

        private void slBorder_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            Settings.Radius = (int)slBorder.Value;
            LoadImage();
        }

        private void cpBackgroundColor_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<System.Windows.Media.Color?> e)
        {
            System.Windows.Media.Color mediacolor = cpBackgroundColor.SelectedColor.Value;
            Settings.BackgroundColor = System.Drawing.Color.FromArgb(mediacolor.A, mediacolor.R, mediacolor.G, mediacolor.B);
            LoadImage();
        }

        private void cpBorderColor_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<System.Windows.Media.Color?> e)
        {
            System.Windows.Media.Color mediacolor = cpBorderColor.SelectedColor.Value;
            Settings.BorderColor = System.Drawing.Color.FromArgb(mediacolor.A, mediacolor.R, mediacolor.G, mediacolor.B);
            LoadImage();
        }

        private void slBorderWidth_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            Settings.BorderWidth = slBorderWidth.Value;
            LoadImage();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            DocMaker.Models.Settings settings = new DocMaker.Models.Settings
            {
                BackgroundColor = Settings.BackgroundColor,
                BorderColor = Settings.BorderColor,
                BorderWidth = Settings.BorderWidth,
                Radius = Settings.Radius
            };

            string filePath = Environment.ExpandEnvironmentVariables(App.SETTINGSPATH);
            string directory = System.IO.Path.GetDirectoryName(filePath);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
            if (File.Exists(filePath))
                File.Delete(filePath);
            File.WriteAllText(filePath, JsonConvert.SerializeObject(settings));
        }
    }
}
