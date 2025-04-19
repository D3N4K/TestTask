using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using TestTskApp.Classes;

namespace TestTskApp.Windows
{
    /// <summary>
    /// Логика взаимодействия для DataWindow.xaml
    /// </summary>
    public partial class DataWindow : Window
    {
        string filePath;
        List<InfoObject> infoObjects = new List<InfoObject>();
        List<InfoObject> infoObjectsCsv = new List<InfoObject>();
        List<InfoObject> infoObjectsExcel = new List<InfoObject>();
        InfoObject infoObject = new InfoObject();
        int selectedRow = -1;
        InfoObject currentObject = new InfoObject();
        public DataWindow()
        {
            InitializeComponent();
        }
        private void miImportCsv_Click(object sender, RoutedEventArgs e)
        {
            OpenFile();
            infoObjectsCsv = infoObject.ImportCSV(filePath);
            infoObjects.AddRange(infoObjectsCsv);
            ViewData();
        }
        private void miImportExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFile();
            infoObjectsExcel = infoObject.ImportExcel(filePath);
            infoObjects.AddRange(infoObjectsExcel);
            ViewData();
        }
        void OpenFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            filePath = openFileDialog.FileName;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
        }
        void ViewData()
        {
            dgData.DataContext = infoObjects;
            dgData.SetBinding(System.Windows.Controls.DataGrid.ItemsSourceProperty, new System.Windows.Data.Binding() { Path = new PropertyPath(".") });
        }

        private void dgData_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedRow = dgData.SelectedIndex;
            if (selectedRow > -1)
            {
                currentObject = infoObjects[selectedRow];
                tbName.Text = $"Name: {currentObject.Name}";
                tbDistance.Text = $"Distance: {currentObject.Distance}";
                tbAngle.Text = $"Angle: {currentObject.Angle}";
                tbWidth.Text = $"Width: {currentObject.Width}";
                tbHeight.Text = $"Height: {currentObject.Height}";
                tbIsDefect.Text = $"IsDefect: {currentObject.IsDefect}";
                canvasObject.Children.Clear();
                DrawCanvas(currentObject.Distance, currentObject.Angle, currentObject.Height, currentObject.Width);
            }
        }
        void DrawCanvas(float distance, float angle, float height, float width)
        {
            double x = distance * (canvasObject.ActualWidth / 20);
            double y = (12 - angle) * (canvasObject.ActualHeight / 12);
            double objWidth = width * (canvasObject.ActualWidth / 20);
            double objHeight = height * (canvasObject.ActualHeight / 12);
            Rectangle rectangle = new Rectangle
            {
                Width = objWidth,
                Height = objHeight,
                Stroke = Brushes.Black,
                StrokeThickness = 2,
                Fill = Brushes.LightBlue
            };
            Canvas.SetLeft(rectangle, x);
            Canvas.SetTop(rectangle, y);
            canvasObject.Children.Add(rectangle);
        }
    }
}
