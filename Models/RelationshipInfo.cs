using System.Text.Json.Serialization;

namespace MsAccessExtract.Models;

internal sealed class RelationshipCollection
{
    [JsonPropertyName("relationships")]
    public List<RelationshipInfo> Relationships { get; set; } = [];
}

internal sealed class RelationshipInfo
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("table")]
    public string Table { get; set; } = string.Empty;

    [JsonPropertyName("foreignTable")]
    public string ForeignTable { get; set; } = string.Empty;

    [JsonPropertyName("fields")]
    public List<RelationField> Fields { get; set; } = [];

    [JsonPropertyName("enforceIntegrity")]
    public bool EnforceIntegrity { get; set; }

    [JsonPropertyName("cascadeUpdate")]
    public bool CascadeUpdate { get; set; }

    [JsonPropertyName("cascadeDelete")]
    public bool CascadeDelete { get; set; }
}

internal sealed class RelationField
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("foreignName")]
    public string ForeignName { get; set; } = string.Empty;
}
