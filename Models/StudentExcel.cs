namespace ExcelToDatabase.Models
{
    public class StudentExcel
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public DateOnly DOB { get; set; }
        public int marks { get; set; }

    }
}
