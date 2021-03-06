﻿using System.Collections.Generic;
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
        private const string PhCult = "Фізична культура";

        private readonly ExcelParse _excelParse;
        private CalculateGrade _calculateGrade;

        public CalculateService()
        {
            this._calculateGrade = new CalculateGrade();
            this._excelParse = new ExcelParse();
        }

        public List<Student> ParseInputExcels(string[] studentPaths, string creditsPath, string savePath, CalculationSettings calculationSettings)
        {
            // Get source data of students
            var studentSource = studentPaths.AsParallel().Select(path => this._excelParse.ParseStudentExcel(path)).ToList();
            // Get source data for subject for every term
            var subjectCredits = this._excelParse.ParseCreditsExcelFile(creditsPath);

            #region Remove unneed subjects
            studentSource.ForEach(student =>
            {
                student.Subjects = _calculateGrade.RemoveUnnedSubject(student.Subjects, PhCult).ToList();
            });
            subjectCredits = _calculateGrade.RemoveUnnedSubject(subjectCredits, PhCult).ToList();

            #endregion

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
                    student.Subjects = _calculateGrade.RemoveUnneedSubjectTypes(student.Subjects, SubjectType.Offset).ToList();
                });
                subjectCredits = _calculateGrade.RemoveUnneedSubjectTypes(subjectCredits, SubjectType.Offset).ToList();
            }
            #endregion

            // Get value of credits for each subject
            var allCredits = _calculateGrade.CountCreditsForSubject(subjectCredits);

            // Calc proportional mark for subject in every term
            var studentOut =
                studentSource.Select(student => _calculateGrade.AverageStudentGrade(student, subjectCredits, allCredits)).ToList();

            studentOut = studentOut.Select(student =>
            {
                student.Subjects = _calculateGrade.GradeForSubjectAllTime(student.Subjects).ToList();
                student.AvgBologneAllYears = student.Subjects.Select(subj => subj.Grade.BolognaGrade).Sum() /
                                           student.Subjects.Count;
                student.AvgClassicAllYears = student.Subjects.Select(subj => subj.Grade.ClassicGrade).Sum()/
                                           student.Subjects.Count;


                return student;
            }).ToList();

            var saveResults = studentOut.AsParallel().Select(student => _excelParse.SaveStudentExcel(student, savePath));
            var answer = saveResults.Aggregate(true, (a, b) => a && b);
            
            return studentOut;
        }


    }


}