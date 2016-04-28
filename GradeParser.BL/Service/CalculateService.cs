using System.Collections.Generic;
using System.Linq;
using GradeParser.BL.Data.Interface;
using GradeParser.BL.Data.Model;
using GradeParser.BL.ExcelFunc;

namespace GradeParser.BL.Service
{
    public class CalculateService : IService
    {
        private readonly ExcelParse _excelParse;
        private ExcelSave _excelSave;

        public CalculateService()
        {
            this._excelSave = new ExcelSave();
            this._excelParse = new ExcelParse();
        }

        public List<Student> ParseInputExcels(string[] studentPaths, string creditsPath, CalculationSettings calculationSettings)
        {
            var studentSource = studentPaths.AsParallel().Select(path => this._excelParse.ParseStudentExcel(path)).ToList();
            var subjectCredits = this._excelParse.ParseCredirExcelFile(creditsPath);
            //studentSource.Select(st => st.Subjects.Where(sub => sub.Type == SubjectType.Offset)).First()
            //var withoutRussian = studentSource.Select(st => st.Subjects.Where(sub => !sub.Name.Contains("Russian")));

            // Remove Russian language
            studentSource = studentSource.Select(st =>
            {
                var stud = new Student
                {
                    Name = st.Name,
                    StudyGroup = st.StudyGroup,
                    Subjects = st.Subjects.Where(sub => !sub.Name.Contains("Russian")).ToList()
                };

                return stud;
            }).ToList();
            

            #region Filter by allowed subjects for calculation(diff/offset/exam)
            if (!calculationSettings.AllowExam)
            {
                // Skip exams
                studentSource = studentSource.Select(st =>
                {
                    var stud = new Student
                    {
                        Name = st.Name,
                        StudyGroup = st.StudyGroup,
                        Subjects = st.Subjects.Where(sub => sub.Type != SubjectType.Exam).ToList()
                    };

                    return stud;
                }).ToList();

                subjectCredits = subjectCredits.Where(sub => sub.Type != SubjectType.Exam).ToList();
            }

            if (!calculationSettings.AllowDiffOffset)
            {
                //Skip diff offset
                studentSource = studentSource.Select(st =>
                {
                    var stud = new Student
                    {
                        Name = st.Name,
                        StudyGroup = st.StudyGroup,
                        Subjects = st.Subjects.Where(sub => sub.Type != SubjectType.DiffOffset).ToList()
                    };

                    return stud;
                }).ToList();

                subjectCredits = subjectCredits.Where(sub => sub.Type != SubjectType.DiffOffset).ToList();
            }

            if (!calculationSettings.AllowOffset)
            {
                //Skip offset
                studentSource = studentSource.Select(st =>
                {
                    var stud = new Student
                    {
                        Name = st.Name,
                        StudyGroup = st.StudyGroup,
                        Subjects = st.Subjects.Where(sub => sub.Type != SubjectType.Offset).ToList()
                    };

                    return stud;
                }).ToList();

                subjectCredits = subjectCredits.Where(sub => sub.Type != SubjectType.Offset).ToList();
            }
            #endregion



            return studentSource;
        }
    }
}