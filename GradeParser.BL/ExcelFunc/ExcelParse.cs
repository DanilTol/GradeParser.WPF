using System;
using System.Collections.Generic;
using GradeParser.BL.Data.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace GradeParser.BL.ExcelFunc
{
    internal class ExcelParse
    {
        #region Const variables

        private const string SubjectTypeNameExam = "екзамен";
        private const string SubjectTypeNameOffset = "залік";
        private const string SubjectTypeNameDiffOffset = "диф. залік";

        #endregion

        public Student ParseStudentExcel(string path)
        {
            var _excelApp = new Excel.Application();
            var xlWorkBook = _excelApp.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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
                    var subject = new Subject
                    {
                        Name = row.Cells[1, 2].Value2.ToString(),
                        Grade = new Grade
                        {
                            ClassicGrade = row.Cells[1, 4].Value2.ToString(),
                            BolognaGrade = row.Cells[1, 5].Value2.ToString(),
                            ESTCGrade = row.Cells[1, 6].Value2.ToString()
                        },
                        Years = row.Cells[1, 7].Value2.ToString(),
                        Term = row.Cells[1, 8].Value2.ToString()
                    };

                    switch (row.Cells[1, 3].Value2 as string)
                    {
                        case SubjectTypeNameExam:
                            subject.Type = SubjectType.Exam;
                            break;
                        case SubjectTypeNameOffset:
                            subject.Type = SubjectType.Offset;
                            break;
                        default:
                            subject.Type = SubjectType.DiffOffset;
                            break;
                    }

                    student.Subjects.Add(subject);
                }
                else
                {
                    break;
                }
            }
            xlWorkBook.Close(true, null, null);
            _excelApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(_excelApp);


            return student;
        }

        public List<SubjectCredit> ParseCredirExcelFile(string path)
        {
            var _excelApp = new Excel.Application();
            var xlWorkBook = _excelApp.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];

            var usedRange = xlWorkSheet.UsedRange;

            var subjectList = new List<SubjectCredit>();

            foreach (Excel.Range row in usedRange.Rows.Offset[5, 0])
            {
                if (row.Cells[1, 2].Value2 != null)
                {
                    var subjectCredit = new SubjectCredit
                    {
                        Name = row.Cells[1, 1].Value2.ToString(),
                        Years = row.Cells[1, 2].Value2.ToString(),
                        Term = row.Cells[1, 3].Value2.ToString(),
                        Type = row.Cells[1, 4].Value2.ToString(),
                        Credit = row.Cells[1, 5].Value2 == null ? default(int) : int.Parse(row.Cells[1, 1].Value2 as string)
                    };

                    subjectList.Add(subjectCredit);
                }
                else
                {
                    break;
                }
            }
            xlWorkBook.Close(true, null, null);
            _excelApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(_excelApp);

            return subjectList;
        } 

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}