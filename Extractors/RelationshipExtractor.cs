using System.Text;
using System.Text.Json;
using MsAccessExtract.Helpers;
using MsAccessExtract.Models;

namespace MsAccessExtract.Extractors;

/// <summary>
/// Extracts database relationships from DAO Relations collection.
/// </summary>
internal sealed class RelationshipExtractor
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly ConsoleLogger _logger;

    public RelationshipExtractor(ConsoleLogger logger)
    {
        _logger = logger;
    }

    public void Extract(dynamic accessApp, string outputFolder)
    {
        _logger.Section("Relationships");
        Directory.CreateDirectory(outputFolder);

        dynamic? currentDb = null;
        dynamic? relations = null;

        try
        {
            currentDb = accessApp.CurrentDb();
            relations = currentDb.Relations;
            int count = (int)relations.Count;

            var collection = new RelationshipCollection();

            for (int i = 0; i < count; i++)
            {
                dynamic? rel = null;
                try
                {
                    rel = relations[i];
                    string name = (string)rel.Name;

                    // Skip system relationships
                    if (name.StartsWith("MSys", StringComparison.OrdinalIgnoreCase))
                        continue;

                    int attributes = (int)rel.Attributes;
                    bool enforceIntegrity = (attributes & AccessConstants.DbRelationDontEnforce) == 0;

                    var relInfo = new RelationshipInfo
                    {
                        Name = name,
                        Table = (string)rel.Table,
                        ForeignTable = (string)rel.ForeignTable,
                        EnforceIntegrity = enforceIntegrity,
                        CascadeUpdate = (attributes & AccessConstants.DbRelationUpdateCascade) != 0,
                        CascadeDelete = (attributes & AccessConstants.DbRelationDeleteCascade) != 0,
                    };

                    // Extract fields
                    dynamic? relFields = null;
                    try
                    {
                        relFields = rel.Fields;
                        int fieldCount = (int)relFields.Count;

                        for (int f = 0; f < fieldCount; f++)
                        {
                            dynamic? relField = null;
                            try
                            {
                                relField = relFields[f];
                                relInfo.Fields.Add(new RelationField
                                {
                                    Name = (string)relField.Name,
                                    ForeignName = (string)relField.ForeignName
                                });
                            }
                            finally
                            {
                                ComHelper.ReleaseComObject(relField);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.ReleaseComObject(relFields);
                    }

                    collection.Relationships.Add(relInfo);
                    _logger.Success("Relationships", $"{relInfo.Table} → {relInfo.ForeignTable} ({name})");
                }
                catch (Exception ex)
                {
                    _logger.Error($"Failed to export relation at index {i}: {ex.Message}");
                }
                finally
                {
                    ComHelper.ReleaseComObject(rel);
                }
            }

            if (collection.Relationships.Count > 0)
            {
                // Sort for stable output (prevents false diffs)
                collection.Relationships = collection.Relationships
                    .OrderBy(r => r.Table)
                    .ThenBy(r => r.ForeignTable)
                    .ThenBy(r => r.Name)
                    .ToList();

                string json = JsonSerializer.Serialize(collection, JsonOptions);
                string filePath = Path.Combine(outputFolder, "relations.json");
                File.WriteAllText(filePath, json + "\n", new UTF8Encoding(false));
            }
            else
            {
                _logger.Info("No relationships found");
            }
        }
        finally
        {
            ComHelper.ReleaseAll(relations, currentDb);
        }
    }
}
