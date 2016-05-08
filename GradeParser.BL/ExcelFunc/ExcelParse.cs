using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using GradeParser.BL.Data.Model;
using GradeParser.BL.Data.Model.Subjects;
using Excel = Microsoft.Office.Interop.Excel;

namespace GradeParser.BL.ExcelFunc
{
    internal class ExcelParse
    {
        #region Const variables

        private const string SubjectTypeNameExam = "екзамен";
        private const string SubjectTypeNameOffset = "залік";
        private const string SubjectTypeNameDiffOffset = "диф. залік";
        private const string StudentIsFree = "зв.";
        private const string StudentIsMiss = "нея.";

        #endregion

        public Student ParseStudentExcel(string path)
        {
            var excelApp = new Excel.Application();
            var xlWorkBook = excelApp.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];

            var usedRange = xlWorkSheet.UsedRange;

            var student = new Student
            {
                Name = xlWorkSheet.Range["A1"].Value2.ToString(),
                StudyGroup = xlWorkSheet.Range["A2"].Value2.ToString().Split('(', ')')[0].Trim(),
                Subjects = new List<Subject>()
            };

            foreach (Excel.Range row in usedRange.Rows.Offset[5, 0])
            {
                if (row.Cells[1, 2].Value2 != null)
                {
                    var classicGradeExcel = row.Cells[1, 4].Value2.ToString();
                    var bologneGradeExcel = row.Cells[1, 5].Value2.ToString();

                    if (classicGradeExcel == StudentIsFree || classicGradeExcel == StudentIsMiss)
                    {
                        continue;
                    }

                    var subject = new Subject
                    {
                        Name = row.Cells[1, 2].Value2.ToString(),
                        Grade = new Grade
                        {
                            ClassicGrade = Int32.Parse(classicGradeExcel),
                            BolognaGrade = Int32.Parse(bologneGradeExcel),
                            ESTCGrade = row.Cells[1, 6].Value2.ToString()
                        },
                        Years = row.Cells[1, 7].Value2.ToString(),
                        Term = row.Cells[1, 8].Value2.ToString(),
                        Type = RecognizeSubjectType(row.Cells[1, 3].Value2 as string)
                    };

                    student.Subjects.Add(subject);
                }
                else
                {
                    break;
                }
            }

            xlWorkBook.Close(true, null, null);
            excelApp.Quit();
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(excelApp);

            return student;
        }

        public List<SubjectCredit> ParseCreditsExcelFile(string path)
        {
            var excelApp = new Excel.Application();
            var xlWorkBook = excelApp.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];

            var usedRange = xlWorkSheet.UsedRange;

            var subjectList = new List<SubjectCredit>();

            foreach (Excel.Range row in usedRange.Rows)
            {
                if (row.Cells[1, 2].Value2 != null)
                {
                    var creditsExcel = row.Cells[1, 5].Value2.ToString();
                    var subjectCredit = new SubjectCredit
                    {
                        Name = row.Cells[1, 1].Value2.ToString(),
                        Years = row.Cells[1, 2].Value2.ToString(),
                        Term = row.Cells[1, 3].Value2.ToString(),
                        Type = RecognizeSubjectType(row.Cells[1, 4].Value2 as string),
                        Credit = Double.Parse(creditsExcel)
                    };

                    subjectList.Add(subjectCredit);
                }
                else
                {
                    break;
                }
            }

            xlWorkBook.Close(true, null, null);
            excelApp.Quit();
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(excelApp);

            return subjectList;
        }

        public bool SaveStudentExcel(Student student, string savePath)
        {
            var xlApp = new Excel.Application();

            object misValue = System.Reflection.Missing.Value;

            var xlWorkBook = xlApp.Workbooks.Add(misValue);
            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];
            Excel.Range chartRange;

            //add data
            // Set name row
            xlWorkSheet.Range["A1", "D1"].Merge(false);
            
            chartRange = xlWorkSheet.Range["A1", "D1"];
            chartRange.FormulaR1C1 = student.Name;
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 14;
            chartRange.Font.Bold = true;

            //Set group row
            xlWorkSheet.Range["A2", "D2"].Merge(false);
            chartRange = xlWorkSheet.Range["A2", "D2"];
            chartRange.FormulaR1C1 = student.StudyGroup;
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 14;
            chartRange.Font.Bold = true;

            chartRange = xlWorkSheet.Range["B3"];
            chartRange.FormulaR1C1 = "ср. Бал";
            chartRange.Font.Size = 12;
            chartRange.Font.Bold = true;

            chartRange = xlWorkSheet.Range["C3","D3"];
            chartRange.Font.Size = 16;
            chartRange.Font.Bold = true;
            chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            xlWorkSheet.Cells[3, 3] = Math.Round(student.AvgClassicAllYears, 2);
            xlWorkSheet.Cells[3, 4] = Math.Round(student.AvgBologneAllYears, MidpointRounding.AwayFromZero);

            chartRange = xlWorkSheet.Range["A4", "D4"];
            chartRange.Font.Size = 12;
            chartRange.Font.Bold = true;
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGoldenrodYellow);

            xlWorkSheet.Cells[4, 1] = "№";
            xlWorkSheet.Cells[4, 2] = "Дисципліна";
            xlWorkSheet.Cells[4, 3] = "Оцінка";
            xlWorkSheet.Cells[4, 4] = "Бал";

            for (int i = 0; i < student.Subjects.Count; i++)
            {
                xlWorkSheet.Cells[i + 5, 1] = i + 1;
                xlWorkSheet.Cells[i + 5, 2] = student.Subjects.ElementAt(i).Name;
                xlWorkSheet.Cells[i + 5, 3] = student.Subjects.ElementAt(i).Grade.ClassicGrade;
                xlWorkSheet.Cells[i + 5, 4] = student.Subjects.ElementAt(i).Grade.BolognaGrade;
            }

            chartRange = xlWorkSheet.UsedRange;
            chartRange.Rows.AutoFit();
            chartRange.Columns.AutoFit();

            savePath += "\\" + student.StudyGroup;
            if (!Directory.Exists(savePath))
            {
                Directory.CreateDirectory(savePath);
            }
            
            xlWorkBook.SaveAs(savePath + "\\"+ student.Name + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            
            return true;
        }
        
        #region Help methods

        private SubjectType RecognizeSubjectType(string typeName)
        {
            switch (typeName)
            {
                case SubjectTypeNameExam:
                    return SubjectType.Exam;
                case SubjectTypeNameOffset:
                    return SubjectType.Offset;
                default:
                    return SubjectType.DiffOffset;
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        #endregion
    }
}