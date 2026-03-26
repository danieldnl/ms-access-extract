using System.Text;
using System.Text.Json;
using MsAccessExtract.Helpers;
using MsAccessExtract.Models;

namespace MsAccessExtract.Extractors;

/// <summary>
/// Extracts table structure (DDL) as JSON schema using DAO TableDefs.
/// </summary>
internal sealed class TableExtractor
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    private readonly ConsoleLogger _logger;

    public TableExtractor(ConsoleLogger logger)
    {
        _logger = logger;
    }

    public void Extract(dynamic accessApp, string outputFolder)
    {
        _logger.Section("Tables");
        Directory.CreateDirectory(outputFolder);

        dynamic? currentDb = null;
        dynamic? tableDefs = null;

        try
        {
            currentDb = accessApp.CurrentDb();
            tableDefs = currentDb.TableDefs;
            int count = (int)tableDefs.Count;

            for (int i = 0; i < count; i++)
            {
                dynamic? tdf = null;
                try
                {
                    tdf = tableDefs[i];
                    string name = (string)tdf.Name;

                    // Skip system and temporary tables
                    if (name.StartsWith("MSys", StringComparison.OrdinalIgnoreCase)
                        || name.StartsWith("~")
                        || name.StartsWith("USys", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // Skip linked tables (they don't contain local structure)
                    string connect = string.Empty;
                    try { connect = (string)tdf.Connect; } catch { }
                    if (!string.IsNullOrEmpty(connect))
                    {
                        _logger.Skip(name, "linked table");
                        continue;
                    }

                    var schema = ExtractTableSchema(tdf, name);
                    string json = JsonSerializer.Serialize(schema, JsonOptions);
                    string filePath = Path.Combine(outputFolder, $"{SanitizeFileName(name)}.json");
                    File.WriteAllText(filePath, json + "\n", new UTF8Encoding(false));
                    _logger.Success("Tables", name);
                }
                catch (Exception ex)
                {
                    _logger.Error($"Failed to export table at index {i}: {ex.Message}");
                }
                finally
                {
                    ComHelper.ReleaseComObject(tdf);
                }
            }
        }
        finally
        {
            ComHelper.ReleaseAll(tableDefs, currentDb);
        }
    }

    private TableSchema ExtractTableSchema(dynamic tdf, string tableName)
    {
        var schema = new TableSchema { Name = tableName };

        // Extract fields
        dynamic? fields = null;
        try
        {
            fields = tdf.Fields;
            int fieldCount = (int)fields.Count;

            for (int f = 0; f < fieldCount; f++)
            {
                dynamic? field = null;
                try
                {
                    field = fields[f];
                    var fieldInfo = new FieldInfo
                    {
                        Name = (string)field.Name,
                        Type = AccessConstants.GetFieldTypeName((int)field.Type),
                        Size = (int)field.Size,
                        Required = (bool)field.Required,
                        OrdinalPosition = (int)field.OrdinalPosition,
                    };

                    // Attributes
                    int attributes = (int)field.Attributes;
                    if ((attributes & AccessConstants.DbAutoIncrField) != 0)
                    {
                        fieldInfo.Type = "AutoNumber";
                        fieldInfo.Attributes = "dbAutoIncrField";
                    }

                    // Optional properties
                    try { fieldInfo.AllowZeroLength = (bool)field.AllowZeroLength; } catch { }
                    try
                    {
                        string dv = Convert.ToString(field.DefaultValue) ?? "";
                        if (!string.IsNullOrEmpty(dv)) fieldInfo.DefaultValue = dv;
                    }
                    catch { }
                    try
                    {
                        string vr = (string)field.ValidationRule;
                        if (!string.IsNullOrEmpty(vr)) fieldInfo.ValidationRule = vr;
                    }
                    catch { }
                    try
                    {
                        string vt = (string)field.ValidationText;
                        if (!string.IsNullOrEmpty(vt)) fieldInfo.ValidationText = vt;
                    }
                    catch { }

                    // Description (stored as property)
                    try
                    {
                        dynamic props = field.Properties;
                        dynamic descProp = props["Description"];
                        string desc = (string)descProp.Value;
                        if (!string.IsNullOrEmpty(desc)) fieldInfo.Description = desc;
                        ComHelper.ReleaseAll(descProp, props);
                    }
                    catch { }

                    schema.Fields.Add(fieldInfo);
                }
                finally
                {
                    ComHelper.ReleaseComObject(field);
                }
            }
        }
        finally
        {
            ComHelper.ReleaseComObject(fields);
        }

        // Extract indexes
        dynamic? indexes = null;
        try
        {
            indexes = tdf.Indexes;
            int indexCount = (int)indexes.Count;

            for (int ix = 0; ix < indexCount; ix++)
            {
                dynamic? idx = null;
                try
                {
                    idx = indexes[ix];
                    string idxName = (string)idx.Name;

                    var indexInfo = new IndexInfo
                    {
                        Name = idxName,
                        Primary = (bool)idx.Primary,
                        Unique = (bool)idx.Unique,
                        Foreign = (bool)idx.Foreign,
                    };

                    try { indexInfo.IgnoreNulls = (bool)idx.IgnoreNulls; } catch { }
                    try { indexInfo.Clustered = (bool)idx.Clustered; } catch { }

                    // Index fields
                    dynamic? idxFields = null;
                    try
                    {
                        idxFields = idx.Fields;
                        int idxFieldCount = (int)idxFields.Count;
                        for (int fi = 0; fi < idxFieldCount; fi++)
                        {
                            dynamic? idxField = null;
                            try
                            {
                                idxField = idxFields[fi];
                                indexInfo.Fields.Add((string)idxField.Name);
                            }
                            finally
                            {
                                ComHelper.ReleaseComObject(idxField);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.ReleaseComObject(idxFields);
                    }

                    schema.Indexes.Add(indexInfo);
                }
                finally
                {
                    ComHelper.ReleaseComObject(idx);
                }
            }
        }
        catch
        {
            // Index collection may not be accessible (e.g. some linked tables)
        }
        finally
        {
            ComHelper.ReleaseComObject(indexes);
        }

        return schema;
    }

    private static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
