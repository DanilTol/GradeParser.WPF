using System;
using System.Collections.Generic;
using System.Linq;
using GradeParser.BL.Data.Model;
using GradeParser.BL.Data.Model.Subjects;
using Microsoft.Office.Interop.Excel;

namespace GradeParser.BL.BL
{
    internal class CalculateGrade
    {
        public IEnumerable<BaseSubject> RemoveUnneedSubjectTypes(List<BaseSubject> subjects, SubjectType subjectType)
        {
            return subjects.Where(sub => sub.Type != subjectType);
        }

        public IEnumerable<BaseSubject> RemoveUnnedSubject(List<BaseSubject> subjects, string subjectName)
        {
            return subjects.Where(sub => !sub.Name.Contains(subjectName)).ToList();
        }

#region Temp methods
        public IEnumerable<SubjectCredit> RemoveUnneedSubjectTypes(List<SubjectCredit> subjects, SubjectType subjectType)
        {
            return subjects.Where(sub => sub.Type != subjectType);
        }

        public IEnumerable<Subject> RemoveUnneedSubjectTypes(List<Subject> subjects, SubjectType subjectType)
        {
            return subjects.Where(sub => sub.Type != subjectType);
        }

        #endregion

        public IEnumerable<SubjectCreditsOnly> CountCreditsForSubject(List<SubjectCredit> subjectCredits)
        {
            //Get distinct subjects
            var allCredits =
                subjectCredits.GroupBy(p => p.Name)
                    .Select(g => g.First())
                    .Select(subject => new SubjectCreditsOnly { Name = subject.Name, Credit = 0 }).ToList();
            
            subjectCredits.ForEach(subCre =>
            {
                allCredits.ForEach(allCre =>
                {
                    if(allCre.Name == subCre.Name)
                    {
                        allCre.Credit += subCre.Credit;
                    }
                });
            });
            
            // Remove all custom curses because they was count like one curse but it was different
            allCredits.Remove(allCredits.FirstOrDefault(allCre => allCre.Name == "Курс на вибір"));

            // add custom curses with a unique name
            allCredits.AddRange(subjectCredits.Where(subCre => subCre.Name == "Курс на вибір").Select(curses => new SubjectCreditsOnly {Name = curses.Name + "_" + curses.Term + "_" + curses.Years, Credit = curses.Credit}));

            //foreach (var subCre in subjectCredits)
            //{
            //    allCredits = allCredits.Select(allCre =>
            //    {
            //        if (allCre.Name == subCre.Name)
            //        {
            //            allCre.Credit += subCre.Credit;
            //        }

            //        return allCre;
            //    }).ToList();
            //}

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