using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SRSPARSER.Models;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddCors(o => o.AddDefaultPolicy(p =>
    p.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader()));

var app = builder.Build();
app.UseCors();

// ─── Health check (Flutter dùng để đợi API ready) ───
app.MapGet("/health", () => Results.Ok(new { status = "ok" }));

// ─── Parse Word ───
app.MapPost("/api/parse-word", (IFormFile file) =>
{
    try
    {
        using var stream = file.OpenReadStream();
        using var doc = WordprocessingDocument.Open(stream, false);
        
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null)
            return Results.BadRequest(new { error = "File Word không hợp lệ" });

        var useCases = ExtractUseCases(body, out var tableFound);

        var result = new ParseResultDto
        {
            UseCases = useCases,
            Metadata = new ParseMetadata
            {
                TotalUC = useCases.Count,
                SourceFile = file.FileName,
                ParsedAt = DateTime.UtcNow,
                TableFound = tableFound,
                Warnings = tableFound ? new List<string>() 
                    : new List<string> { "Không tìm thấy bảng Use Case trong SRS" }
            }
        };

        return Results.Ok(result);
    }
    catch (Exception ex)
    {
        return Results.Problem(ex.Message);
    }
})
.DisableAntiforgery();

app.Run("http://localhost:5000");

// ══════════════════════════════════════════════
// PARSE LOGIC
// ══════════════════════════════════════════════

static List<UseCaseDto> ExtractUseCases(Body body, out bool tableFound)
{
    tableFound = false;
    var result = new List<UseCaseDto>();

    // Tìm bảng có header chứa "Use Case"
    foreach (var table in body.Descendants<Table>())
    {
        var rows = table.Descendants<TableRow>().ToList();
        if (rows.Count < 2) continue;

        var headerCells = rows[0].Descendants<TableCell>()
            .Select(c => c.InnerText.Trim().ToLower())
            .ToList();

        // Có phải bảng UC?
        var isUCTable = headerCells.Any(h => 
            h.Contains("use case") || h.Contains("usecase"));

        if (!isUCTable) continue;

        tableFound = true;

        // Xác định cột nào là ID, Name, Actors, Description
        int idCol = FindColumn(headerCells, "id");
        int nameCol = FindColumn(headerCells, "use case");
        int actorsCol = FindColumn(headerCells, "actor");
        int descCol = FindColumn(headerCells, "description");

        // Parse từng row (bỏ header)
        foreach (var row in rows.Skip(1))
        {
            var cells = row.Descendants<TableCell>()
                .Select(c => c.InnerText.Trim())
                .ToList();

            if (cells.Count == 0) continue;
            if (string.IsNullOrWhiteSpace(cells[0])) continue;
            if (cells[0].ToLower() == "id") continue;  // dòng header lặp

            var uc = new UseCaseDto
            {
                Id = GetCell(cells, idCol),
                Name = GetCell(cells, nameCol),
                Actors = GetCell(cells, actorsCol),
                Description = GetCell(cells, descCol),
            };
            
            uc.Code = NormalizeUCId(uc.Id);
            
            if (!string.IsNullOrWhiteSpace(uc.Name))
                result.Add(uc);
        }
        
        break;  // Chỉ lấy bảng UC đầu tiên
    }

    return result;
}

static int FindColumn(List<string> headers, string keyword)
{
    for (int i = 0; i < headers.Count; i++)
        if (headers[i].Contains(keyword)) return i;
    return -1;
}

static string GetCell(List<string> cells, int index)
{
    if (index < 0 || index >= cells.Count) return "";
    return cells[index];
}

static string NormalizeUCId(string raw)
{
    var digits = new string(raw.Where(char.IsDigit).ToArray());
    if (string.IsNullOrEmpty(digits)) return raw.ToUpper();
    return $"UC{int.Parse(digits):D2}";
}