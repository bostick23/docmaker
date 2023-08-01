using DocMaker.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using Gma.System.MouseKeyHook;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.VisualStyles;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DocMaker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private IKeyboardMouseEvents m_GlobalHook;
        public readonly string TEMP_PATH;
        public bool IsRecording;
        public MainWindow()
        {
            InitializeComponent();
            TEMP_PATH = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "DocMaker");
            if (!Directory.Exists(TEMP_PATH))
                Directory.CreateDirectory(TEMP_PATH);
            IsRecording = false;

        }
        public void Subscribe()
        {
            // Note: for the application hook, use the Hook.AppEvents() instead
            m_GlobalHook = Hook.GlobalEvents();

            m_GlobalHook.MouseDownExt += M_GlobalHook_MouseDownExt;
        }

        private void M_GlobalHook_MouseDownExt(object? sender, MouseEventExtArgs e)
        {
            CaptureScreen(e.Location);
        }

        protected void CaptureScreen(System.Drawing.Point point)
        {
            double screenLeft = 0;
            double screenTop = 0;
            double screenWidth = 1920;
            double screenHeight = 1080;

            using Bitmap bmp = new Bitmap((int)screenWidth, (int)screenHeight);
            using Graphics g = Graphics.FromImage(bmp);
            string filename = "ScreenCapture-" + DateTime.Now.ToString("ddMMyyyy-hhmmss") + ".png";
            Opacity = .0;
            g.CopyFromScreen((int)screenLeft, (int)screenTop, 0, 0, bmp.Size);
            int radius = 30;
            System.Drawing.Pen p = new System.Drawing.Pen(System.Drawing.Color.Red);
            g.DrawEllipse(p, point.X - radius, point.Y - radius,
             radius + radius, radius + radius);
            System.Drawing.Brush brush = new SolidBrush(System.Drawing.Color.FromArgb(50, 98, 0, 238));
            g.FillEllipse(brush, point.X - radius, point.Y - radius,
             radius + radius, radius + radius);
            bmp.Save(System.IO.Path.Combine(TEMP_PATH, filename));
            Opacity = 1;
        }
        public static void OpenWithDefaultProgram(string path)
        {
            using Process fileopener = new Process();

            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + path + "\"";
            fileopener.Start();
        }

        private void BtStartStop_Click(object sender, RoutedEventArgs e)
        {
            if (!IsRecording)
            {
                IsRecording = true;
                BtStartStop.Content = "Stop";
                this.WindowState = WindowState.Minimized;
                Subscribe();
            }
            else
            {
                IsRecording = false;
                BtStartStop.Content = "Start";
                m_GlobalHook.Dispose();
                string wordFileName = new WordHelper(TEMP_PATH).CreateAndInsertAllPictures();
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Documento di Word (*.docx)|*.docx";
                if (saveFileDialog.ShowDialog() == true)
                    File.Move(wordFileName, saveFileDialog.FileName);
                OpenWithDefaultProgram(saveFileDialog.FileName);
            }
        }
    }
}
