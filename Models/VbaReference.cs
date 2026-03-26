using System.Text.Json.Serialization;

namespace MsAccessExtract.Models;

internal sealed class VbaReference
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("description")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Description { get; set; }

    [JsonPropertyName("guid")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Guid { get; set; }

    [JsonPropertyName("major")]
    public int Major { get; set; }

    [JsonPropertyName("minor")]
    public int Minor { get; set; }

    [JsonPropertyName("fullPath")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FullPath { get; set; }

    [JsonPropertyName("isBroken")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool IsBroken { get; set; }

    [JsonPropertyName("builtIn")]
    public bool BuiltIn { get; set; }
}

internal sealed class DatabaseProperties
{
    [JsonPropertyName("databaseName")]
    public string DatabaseName { get; set; } = string.Empty;

    [JsonPropertyName("exportDate")]
    public string ExportDate { get; set; } = string.Empty;

    [JsonPropertyName("accessVersion")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AccessVersion { get; set; }

    [JsonPropertyName("vbaProjectName")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? VbaProjectName { get; set; }

    [JsonPropertyName("vbaReferences")]
    public List<VbaReference> VbaReferences { get; set; } = [];
}
