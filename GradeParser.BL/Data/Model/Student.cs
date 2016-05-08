using System.Collections.Generic;
using GradeParser.BL.Data.Model.Subjects;

namespace GradeParser.BL.Data.Model
{
    public class Student : BaseEntity
    {
        public string StudyGroup { get; set; }
        public List<Subject> Subjects { get; set; } 
        public double AvgBologneAllYears { get; set; }
        public double AvgClassicAllYears { get; set; }
    }
}