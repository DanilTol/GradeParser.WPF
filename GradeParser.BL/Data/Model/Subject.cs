namespace GradeParser.BL.Data.Model
{
    public class Subject : BaseEntity
    {
        public SubjectType Type { get; set; }
        public Grade Grade { get; set; }
        public string Years { get; set; }
        public string Term { get; set; }
    }
}