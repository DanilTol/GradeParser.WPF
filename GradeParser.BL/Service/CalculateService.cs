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

        public List<Student> ParseInputExcels(string[] studentPaths, string creditsPath)
        {
            var studentSource = studentPaths.AsParallel().Select(path => this._excelParse.ParseStudentExcel(path)).ToList();
            //studentSource.Select(st => st.Subjects.Where(sub => sub.Type == SubjectType.Offset)).First()

            var withoutRussian = studentSource.Select(st => st.Subjects.Where(sub => !sub.Name.Contains("Russian")));

            return studentSource;
        } 
    }
}