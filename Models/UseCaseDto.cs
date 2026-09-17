namespace SRSPARSER.Models;

public class UseCaseDto
{
    public string Id { get; set; } = "";
    public string Code { get; set; } = "";
    public string Name { get; set; } = "";
    public string Actors { get; set; } = "";
    public string Description { get; set; } = "";
}

public class ParseResultDto
{
    public List<UseCaseDto> UseCases { get; set; } = new();
    public ParseMetadata Metadata { get; set; } = new();
}

public class ParseMetadata
{
    public int TotalUC { get; set; }
    public string SourceFile { get; set; } = "";
    public DateTime ParsedAt { get; set; }
    public bool TableFound { get; set; }
    public List<string> Warnings { get; set; } = new();
}