using System;
using System.Collections.Generic;
using System.Linq;
using GradeParser.BL.Data.Model;
using Microsoft.Office.Interop.Excel;

namespace GradeParser.BL.BL
{
    internal class CalculateGrade
    {
        public List<BaseSubject> RemoveUnneedSubjectTypes(List<BaseSubject> subjects, SubjectType subjectType)
        {
            return subjects.Where(sub => sub.Type != subjectType).ToList();
        }

        public List<BaseSubject> RemoveUnnedSubject(List<BaseSubject> subjects, string subjectName)
        {
            return subjects.Where(sub => !sub.Name.Contains(subjectName)).ToList();
        }

        public IEnumerable<SubjectCreditsOnly> CountCreditsForSubject(IEnumerable<SubjectCredit> subjectCredits)
        {
            //var allCredits =
            //    subjectCredits.DistinctBy(sub => sub.Name)
            //        .Select(sub => new SubjectCreditsOnly { Credit = 0, Name = sub.Name });


            var allCredits =
                subjectCredits.GroupBy(p => p.Name)
                    .Select(g => g.First())
                    .Select(subject => new SubjectCreditsOnly { Name = subject.Name, Credit = 0 });

            foreach (var subCre in subjectCredits)
            {
                allCredits.ToList().ForEach(allCre =>
                {
                    //TODO: Sum all credits
                    if (allCre.Name == subCre.Name)
                    {
                        allCre.Credit += subCre.Credit;
                    }
                });
            }

            return allCredits;
        }

        public Student AverageStudentGrade(Student student, IEnumerable<SubjectCredit> subjectCredits, IEnumerable<SubjectCreditsOnly> subjectCreditsOnlies)
        {
            return new Student
            {
                Name = student.Name,
                StudyGroup = student.StudyGroup,
                Subjects = student.Subjects.Select(stdSubject =>
                {
                    // Get credit for subject for a certain year & term
                    var termYearSubCredit = subjectCredits.FirstOrDefault(
                        creSubject =>
                            creSubject.Name == stdSubject.Name && creSubject.Term == stdSubject.Term &&
                            creSubject.Years == stdSubject.Years);
                    var allCreditForSubject =
                        subjectCreditsOnlies.FirstOrDefault(allcre => allcre.Name == stdSubject.Name).Credit;

                    return new Subject
                    {
                        Name = stdSubject.Name,
                        Grade = new Grade
                        {
                            BolognaGrade = GradeForSubjectAllTime(stdSubject.Grade.BolognaGrade, termYearSubCredit.Credit, allCreditForSubject)
                        }
                    };
                }).ToList()
            };
        }

        private int GradeForSubjectAllTime(int curentGrade, double creditsTerm, double creditsAll)
        {
            return Convert.ToInt32(Math.Ceiling(curentGrade * creditsTerm / creditsAll));
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