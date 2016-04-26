using System;
using System.Windows;
using System.Windows.Controls;
using GradeParser.BL.Data.Model;
using GradeParser.WPF.ViewModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GradeParser.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private Student CurrentStudent;
        private bool CalcOffset;
        private ReportLoad ReportLoad;

        public MainWindow()
        {
            InitializeComponent();
            ReportLoad = new ReportLoad();
            CurrentStudent = new Student();
        }

        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            var mi = sender as MenuItem;
            var header = mi.Header.ToString();

            switch (header)
            {
                case "Student":

                    ReportLoad.Student = true;
                    break;
                case "Credits":

                    ReportLoad.Credits = true;
                    break;
                default:

                    ReportLoad.Speciality = true;
                    break;
            }
        }

        private void SaveCalcButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void CalculateButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
