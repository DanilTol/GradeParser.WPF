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

            foreach (Excel.Range row in usedRange.Rows.Offset[5, 0])
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

        public bool SaveStudentExcel(Student student)
        {

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