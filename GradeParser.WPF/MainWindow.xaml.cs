using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using GradeParser.BL.Data.Interface;
using GradeParser.BL.Data.Model;
using GradeParser.BL.Service;
using GradeParser.WPF.ViewModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GradeParser.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        #region UI variables
        private bool CalcOffset;
        private ReportLoad ReportLoad;
        #endregion

        #region Common variables
        private CalculateService _calculateService;
        #endregion
        

        public MainWindow()
        {
            InitializeComponent();
            InitCommonVariables();
        }

        /// <summary>
        /// Initin common variables & Services
        /// </summary>
        private void InitCommonVariables()
        {
            this.ReportLoad = new ReportLoad();
            this._calculateService = new CalculateService();
        }

        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            //Finding who send event for marking right condition(input file)
            var mi = sender as MenuItem;

            string[] paths = { @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls" };

            var std = this._calculateService.ParseInputExcels(paths, string.Empty);

            switch (mi.Header.ToString())
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
