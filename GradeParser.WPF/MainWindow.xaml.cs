using System.Linq;
using System.Windows;
using System.Windows.Forms;
using GradeParser.BL.Data.Model;
using GradeParser.BL.Service;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

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

            #region STUB
            _studentPath = new[] { @"C:\Users\Danil\Desktop\ТОЛМАЧЕВ.xls" };
            _creditPath = @"C:\Users\Danil\Desktop\Credits_545а_545б_КСС.xlsx";
            _savePath = @"C:\Users\Danil\Desktop\Grade\Test\";
            #endregion
        }

        private void CalculateButton_Click(object sender, RoutedEventArgs e)
        {
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

            if (calculationSettings.AllowDiffOffset || calculationSettings.AllowExam || calculationSettings.AllowOffset)
            {
                var std = this._calculateService.ParseInputExcels(_studentPath, _creditPath, _savePath, calculationSettings);
            }
        }

        #region Browse buttons
        private void SaveToPathButton_Click(object sender, RoutedEventArgs e)
        {
            var folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();

            if (string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                return;

            _savePath = folderBrowserDialog.SelectedPath;
            SaveToPathtextBox.Text = _savePath;
        }

        private void CreditsPathButton_Click(object sender, RoutedEventArgs e)
        {
            var fileDialog = new OpenFileDialog
            {
                Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx",
                Title = "Choose Excel files"
            };
            fileDialog.ShowDialog();

            if (string.IsNullOrWhiteSpace(fileDialog.FileName))
                return;

            _creditPath = fileDialog.FileName;
            CreditsPathtextBox.Text = _creditPath;
        }

        private void StudentReportPathButton_Click(object sender, RoutedEventArgs e)
        {
            var fileDialog = new OpenFileDialog
            {
                Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx",
                Multiselect = true,
                Title = "Choose Excel files"
            };
            fileDialog.ShowDialog();

            if (string.IsNullOrWhiteSpace(fileDialog.FileName))
                return;

            _studentPath = fileDialog.FileNames;
            StudentReportPathtextBox.Text = _studentPath.Aggregate((a, b) => a + "\n " + b);
        }
        #endregion
    }
}
