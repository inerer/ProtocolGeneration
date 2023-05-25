namespace ProtocolGeneration.Models;

public class People
{
    public string? LastName { get; set; }
    public string? FirstName { get; set; }
    public string? MiddleName { get; set; }
    public string? Role { get; set; }
    public string? FullName => $"{LastName} {FirstName} {MiddleName}, {Role}";
}