using System;
using System.Runtime.InteropServices.JavaScript;
using Cyriller;
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

    public string? FullName => $"{ LastName} {FirstName} {MiddleName}";
    
    public string? Theme { get; set; }
    
    public string? MainTeacher { get; set; }
    
    public int? CountList { get; set; }
    
    public int? CountGrap { get; set; }
    
    public string? Opinion { get; set; }
    
    public string? Review { get; set; }
    
    public int? Group { get; set; }
    
    public string? FirstQuestion { get; set; }
    
    public string? SecondQuestion { get; set; }
    
    public string? ThirdQuestion { get; set; }
    
    public string? SpecialOpinion { get; set; }
    
    public int Gender { get; set; }
    
    public DateTime? Date { get; set; }
    
    public int VKRGrade { get; set; }
    
    public DateTime? DemoDate { get; set; }
    
    public string? Qualification { get; set; }
    
    public string? DiplomCathegory { get; set; }

    public string? WordGrade
    {
        get
        {
            if (Grade == 5)
                return $"{Grade}(отлично)";
            if (Grade == 4)
                return $"{Grade}(хорошо)";
            if (Grade == 3)
                return $"{Grade}(удолетворительно)";
            return null;

        }
    }
    
    public int voteYes { get; set; }
    
    public int voteNo { get; set; }
    
    public int voteMaybe { get; set; }
    
}