using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using GradeParser.BL.Data.Model;
using GradeParser.BL.Service;
using GradeParser.WPF.ViewModel;
using Microsoft.Win32;

namespace GradeParser.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        #region Common variables
        private CalculateService _calculateService;
        private string[] _studentPath;
        private string _creditPath;
        private string _savePath;
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
            this._calculateService = new CalculateService();
        }

        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            var buttonName = (sender as Button).Name;

            if (buttonName == StudentReportPathButton.Name)
            {
                
            }
            else
            {
                if (buttonName == CreditsPathButton.Name)
                {
                    
                }
                else
                {
                    
                }
            }


            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx";
            fileDialog.Multiselect = true;
            fileDialog.Title = "Choose Excel files";
            
            var dr = fileDialog.ShowDialog();

            if (dr.HasValue && dr.Value)
            {
                _studentPath = fileDialog.FileNames;
                StudentReportPathtextBox.Text = fileDialog.FileNames.Aggregate((a, b) => a + "\n " + b);
            }

            
        }

        private void CalculateButton_Click(object sender, RoutedEventArgs e)
        {
            //_studentPath = new[] { @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls" };
            _creditPath = @"C:\Users\Danil\Desktop\Credits_545а_545б_КСС.xlsx";
            _savePath = @"C:\Users\Danil\Desktop\Grade\Test\";

            if (string.IsNullOrEmpty(_creditPath) || _studentPath == null || _studentPath.Length < 1 ||
                string.IsNullOrEmpty(_savePath))
            {
                MessageBox.Show(
                    "Please choose files with credits, student reports and directory where to save result. Then continue.",
                    "Choose path", MessageBoxButton.OK);
                return;
            }

            var calculationSettings = new CalculationSettings
            {
                AllowDiffOffset = DiffOffsetCheckBox.IsChecked.Value,
                AllowExam = ExamCheckBox.IsChecked.Value,
                AllowOffset = OffsetCheckBox.IsChecked.Value
            };

            var std = this._calculateService.ParseInputExcels(_studentPath, _creditPath, calculationSettings);
        }
    }
}
