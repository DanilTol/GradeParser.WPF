namespace GradeParser.BL.Data.Model
{
    public class BaseSubject : BaseEntity
    {
        public string Years { get; set; }
        public string Term { get; set; }
        public SubjectType Type { get; set; }
    }
}