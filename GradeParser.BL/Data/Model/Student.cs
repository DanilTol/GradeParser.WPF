using System.Collections.Generic;

namespace GradeParser.BL.Data.Model
{
    public class Student : BaseEntity
    {
        public string StudyGroup { get; set; }
        public List<Subject> Subjects { get; set; } 
        public double AvgForAllYears { get; set; }
        public double AvgForSpeciality { get; set; }
    }
}