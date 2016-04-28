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
        private string[] studentPath;
        private string creditPath;
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

            //TODO: open file dialog

            if (mi.Header as string == "Student")
            {
                studentPath = new[] { @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls", @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls" };
                ReportLoad.Student = true;
            }
            else
            {
                creditPath = @"C:\Users\Danil\Desktop\Credits.xlsx";
                ReportLoad.Credits = true;
            }
        }

        private void CalculateButton_Click(object sender, RoutedEventArgs e)
        {
            var std = this._calculateService.ParseInputExcels(studentPath, string.Empty, new CalculationSettings { AllowDiffOffset = true, AllowExam = true, AllowOffset = false });
        }
    }
}
