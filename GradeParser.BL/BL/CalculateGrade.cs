using System.Collections.Generic;
using System.Linq;
using GradeParser.BL.Data.Model;
using Microsoft.Office.Interop.Excel;

namespace GradeParser.BL.BL
{
    public class CalculateGrade
    {
        public List<BaseSubject> RemoveUnneedSubjectTypes(List<BaseSubject> subjects, SubjectType subjectType)
        {
            return subjects.Where(sub => sub.Type != subjectType).ToList();
        }

        public List<BaseSubject> RemoveUnnedSubject(List<BaseSubject> subjects, string subjectName)
        {
            return subjects.Where(sub => !sub.Name.Contains(subjectName)).ToList();
        }

    }
}