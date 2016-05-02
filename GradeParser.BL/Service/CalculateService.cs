using System;
using System.Collections.Generic;
using System.Linq;
using GradeParser.BL.BL;
using GradeParser.BL.Data.Interface;
using GradeParser.BL.Data.Model;
using GradeParser.BL.ExcelFunc;

namespace GradeParser.BL.Service
{
    public class CalculateService : IService
    {
        private readonly ExcelParse _excelParse;
        private CalculateGrade _calculateGrade;

        public CalculateService()
        {
            this._calculateGrade = new CalculateGrade();
            this._excelParse = new ExcelParse();
        }

        public List<Student> ParseInputExcels(string[] studentPaths, string creditsPath, CalculationSettings calculationSettings)
        {
            var studentSource = studentPaths.AsParallel().Select(path => this._excelParse.ParseStudentExcel(path)).ToList();
            var subjectCredits = this._excelParse.ParseCreditsExcelFile(creditsPath);
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
                //studentSource.ForEach(student =>
                //{
                //    student.Subjects = _calculateGrade.RemoveUnneedSubject(student.Subjects, SubjectType.Exam).ToList();
                //});
                //subjectCredits = _calculateGrade.RemoveUnneedSubject(subjectCredits, SubjectType.Exam).ToList();

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








            var AllCredits =
                subjectCredits.DistinctBy(sub => sub.Name)
                    .Select(sub => new SubjectCreditsOnly {Credit = sub.Credit, Name = sub.Name});



            return studentSource;
        }


    }

    public static class Helper
    {

        #region helpers

        public static IEnumerable<TSource> DistinctBy<TSource, TKey>
            (this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            HashSet<TKey> seenKeys = new HashSet<TKey>();

            return source.Where(element => seenKeys.Add(keySelector(element)));
        }

        #endregion
    }
}