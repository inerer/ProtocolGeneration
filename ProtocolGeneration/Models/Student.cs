using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace ProtocolGeneration.Models;

public class Student
{
    public string? LastName { get; set; }
    public string? FirstName { get; set; }
    public string? MiddleName { get; set; }
    public decimal? Ball { get; set; }

    public int? Grade
    {
        get
        {
            if (Ball <= (decimal?)19.99)
                 return Grade = 2;
            
            if (Ball >= 20 && Ball <= (decimal?)39.99)
                return Grade = 3;
            
            if (Ball >= 40 && Ball <= (decimal?)69.99)
                return Grade = 4;
            
            if (Ball is >= 70 and <= 100)
                return Grade = 5;

            return Grade = null;
        }
        set
        {
            
        }
    }
}