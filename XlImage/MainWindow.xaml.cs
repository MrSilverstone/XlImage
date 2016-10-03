using System;
using System.Collections.Generic;
using System.Drawing;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace XlImage
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private String imagePath;
        private int ImageWidth = 0;
        private int ImageHeight = 0;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openPicker = new Microsoft.Win32.OpenFileDialog();

            openPicker.Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp";
            Nullable<bool> result = openPicker.ShowDialog();

            if (result == true)
            {
                imagePath = openPicker.FileName.ToString();
                FileTextBox.Text = imagePath;

                using (var img = System.Drawing.Image.FromFile(imagePath))
                {
                    ImageWidth = img.Width;
                    ImageHeight = img.Height;

                    WidthTextBox.Text = ImageWidth.ToString();
                    HeightTextBox.Text = ImageHeight.ToString();
                }
            }
        }

        private void WidthTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (ImageWidth == 0 || ImageHeight == 0 || !IsNumber(WidthTextBox.Text))
            {
                WidthTextBox.Text = "";
                return;
            }

            int requestedWidth = int.Parse(WidthTextBox.Text);
            int NewHeight = (ImageHeight / ImageWidth) * requestedWidth;

            Console.WriteLine(NewHeight);

            if (NewHeight == 0)
                return;

            ImageWidth = requestedWidth;
            ImageHeight = NewHeight;
            HeightTextBox.Text = ImageHeight.ToString();
        }

        private void HeightTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (ImageWidth == 0 || ImageHeight == 0 || !IsNumber(HeightTextBox.Text))
            {
                HeightTextBox.Text = "";
                return;
            }

            int requestedHeight = int.Parse(HeightTextBox.Text);
            int NewWidth = (ImageWidth / ImageHeight) * requestedHeight;

            ImageHeight = requestedHeight;
            ImageWidth = NewWidth;
            WidthTextBox.Text = ImageWidth.ToString();
        }

        public static bool IsNumber(string s)
        {
            return s.All(char.IsDigit);
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            var xlApp = new Microsoft.Office.Interop.Excel.Application();

            xlApp.ScreenUpdating = true;
            xlApp.Visible = true;
            xlApp.Interactive = true;
            xlApp.IgnoreRemoteRequests = false;

            if (xlApp == null)
            {
                MessageBoxResult error = MessageBox.Show("Error while loading Excel.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
                return;
            }

            var wb = xlApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            var ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];


            using (var original = new Bitmap(imagePath))
            {
                using (var resized = new Bitmap(original, new System.Drawing.Size(ImageWidth, ImageHeight)))
                {
                    try
                    {
                        for (int y = 1; y < ImageHeight + 1; y++)
                        {
                            ws.Columns[y].ColumnWidth = 2;
                            for (int x = 1; x < ImageWidth + 1; x++)
                            {
                                ws.Cells[y, x].Interior.Color = System.Drawing.ColorTranslator.ToOle(resized.GetPixel(x - 1, y - 1));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        MessageBoxResult error = MessageBox.Show("Error while drawing. Please do not use excel at the same time.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        Application.Current.Shutdown();
                        return;
                    }
                }
            }

        }
    }
}
