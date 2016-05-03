using System;
using System.Collections.Generic;
using System.Linq;
using GradeParser.BL.BL;
using GradeParser.BL.Data.Interface;
using GradeParser.BL.Data.Model;
using GradeParser.BL.Data.Model.Subjects;
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
            

            #region Filter by allowed subjects for calculation(diff/offset/exam)
            if (!calculationSettings.AllowExam)
            {
                studentSource.ForEach(student =>
                {
                    student.Subjects = _calculateGrade.RemoveUnneedSubjectTypes(student.Subjects, SubjectType.Exam).ToList();
                });
                subjectCredits = _calculateGrade.RemoveUnneedSubjectTypes(subjectCredits, SubjectType.Exam).ToList();

                // Skip exams
                //studentSource = studentSource.Select(st =>
                //{
                //    var stud = new Student
                //    {
                //        Name = st.Name,
                //        StudyGroup = st.StudyGroup,
                //        Subjects = st.Subjects.Where(sub => sub.Type != SubjectType.Exam).ToList()
                //    };

                //    return stud;
                //}).ToList();

                //subjectCredits = subjectCredits.Where(sub => sub.Type != SubjectType.Exam).ToList();
            }

            if (!calculationSettings.AllowDiffOffset)
            {
                studentSource.ForEach(student =>
                {
                    student.Subjects = _calculateGrade.RemoveUnneedSubjectTypes(student.Subjects, SubjectType.DiffOffset).ToList();
                });
                subjectCredits = _calculateGrade.RemoveUnneedSubjectTypes(subjectCredits, SubjectType.DiffOffset).ToList();
            }

            if (!calculationSettings.AllowOffset)
            {
                studentSource.ForEach(student =>
                {
                    student.Subjects = _calculateGrade.RemoveUnneedSubjectTypes(student.Subjects, SubjectType.DiffOffset).ToList();
                });
                subjectCredits = _calculateGrade.RemoveUnneedSubjectTypes(subjectCredits, SubjectType.DiffOffset).ToList();
            }
            #endregion
            
            var allCredits = _calculateGrade.CountCreditsForSubject(subjectCredits);

            var studentOut =
                studentSource.Select(student => _calculateGrade.AverageStudentGrade(student, subjectCredits, allCredits));


            return studentOut.ToList();
        }


    }

   
}