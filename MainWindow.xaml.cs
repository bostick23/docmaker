using DocMaker.Helpers;
using DocMaker.Models;
using Gma.System.MouseKeyHook;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows;

namespace DocMaker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        protected Models.Settings Settings;
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
            LoadParameters();

        }
        protected void LoadParameters()
        {
            string filePath = Environment.ExpandEnvironmentVariables(App.SETTINGSPATH);
            if (File.Exists(filePath))
            {
                try
                {
                    string fileContent = File.ReadAllText(filePath);
                    Settings = JsonConvert.DeserializeObject<Models.Settings>(fileContent);
                }
                catch (Exception ex)
                {
                    //TODO fare messaggio di errori come si deve
                    MessageBox.Show(ex.ToString());
                }
            }
            else
            {
                Settings = new Models.Settings
                {
                    BackgroundColor = System.Drawing.Color.FromArgb(50, 98, 0, 238),
                    BorderColor = System.Drawing.Color.Red,
                    Radius = 30,
                    BorderWidth = 10
                };
            }
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
            System.Drawing.Rectangle resolution = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
            double screenLeft = 0;
            double screenTop = 0;
            double screenWidth = resolution.Width;
            double screenHeight = resolution.Height;

            using Bitmap bmp = new Bitmap((int)screenWidth, (int)screenHeight);
            using Graphics g = Graphics.FromImage(bmp);
            string filename = "ScreenCapture-" + DateTime.Now.ToString("ddMMyyyy-hhmmss") + ".png";
            Opacity = .0;
            g.CopyFromScreen((int)screenLeft, (int)screenTop, 0, 0, bmp.Size);
            System.Drawing.Pen p = new System.Drawing.Pen(Settings.BorderColor);
            p.Width = (float)Settings.BorderWidth;
            g.DrawEllipse(p, point.X - Settings.Radius, point.Y - Settings.Radius,
             Settings.Radius + Settings.Radius, Settings.Radius + Settings.Radius);
            System.Drawing.Brush brush = new SolidBrush(Settings.BackgroundColor);
            g.FillEllipse(brush, point.X - Settings.Radius, point.Y - Settings.Radius,
             Settings.Radius + Settings.Radius, Settings.Radius + Settings.Radius);
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
                bool fileSaved = false, continueLoop = true;
                while (continueLoop)
                {
                    if (saveFileDialog.ShowDialog() == true)
                    {
                        if (File.Exists(saveFileDialog.FileName))
                        {
                            try
                            {
                                File.Delete(saveFileDialog.FileName);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("File opened by another program. Unable to proceed",
                                    "Save file",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                                continue;
                            }
                        }

                        File.Move(wordFileName, saveFileDialog.FileName);
                        fileSaved = true;
                        continueLoop = false;
                    }
                    else
                        continueLoop = false;
                }
                if (fileSaved)
                    OpenWithDefaultProgram(saveFileDialog.FileName);
            }
        }

        private void BtOpenParameters_Click(object sender, RoutedEventArgs e)
        {
            ParametersWindow parametersWindow = new ParametersWindow(Settings);
            parametersWindow.ShowDialog();
        }
    }
}
