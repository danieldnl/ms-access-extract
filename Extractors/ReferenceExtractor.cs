using System.Text;
using System.Text.Json;
using MsAccessExtract.Helpers;
using MsAccessExtract.Models;

namespace MsAccessExtract.Extractors;

/// <summary>
/// Extracts VBA project references and database properties.
/// </summary>
internal sealed class ReferenceExtractor
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    private readonly ConsoleLogger _logger;

    public ReferenceExtractor(ConsoleLogger logger)
    {
        _logger = logger;
    }

    public void Extract(dynamic accessApp, string outputFolder, string databaseName)
    {
        _logger.Section("Database Properties & VBA References");

        var dbProps = new DatabaseProperties
        {
            DatabaseName = databaseName,
            ExportDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
        };

        // Access version
        try
        {
            dbProps.AccessVersion = (string)accessApp.Version;
        }
        catch { }

        // VBA references
        dynamic? vbProject = null;
        dynamic? references = null;

        try
        {
            vbProject = accessApp.VBE.ActiveVBProject;
            dbProps.VbaProjectName = (string)vbProject.Name;
            references = vbProject.References;
            int count = (int)references.Count;

            for (int i = 1; i <= count; i++)
            {
                dynamic? refItem = null;
                try
                {
                    refItem = references.Item(i);

                    var vbaRef = new VbaReference
                    {
                        Name = (string)refItem.Name,
                    };

                    try { vbaRef.Description = (string)refItem.Description; } catch { }
                    try { vbaRef.Guid = (string)refItem.GUID; } catch { }
                    try { vbaRef.Major = (int)refItem.Major; } catch { }
                    try { vbaRef.Minor = (int)refItem.Minor; } catch { }
                    try { vbaRef.FullPath = (string)refItem.FullPath; } catch { }
                    try { vbaRef.IsBroken = (bool)refItem.IsBroken; } catch { }
                    try { vbaRef.BuiltIn = (bool)refItem.BuiltIn; } catch { }

                    dbProps.VbaReferences.Add(vbaRef);
                    _logger.Success("References", $"{vbaRef.Name}{(vbaRef.IsBroken ? " (BROKEN)" : "")}");
                }
                catch (Exception ex)
                {
                    _logger.Error($"Failed to read reference at index {i}: {ex.Message}");
                }
                finally
                {
                    ComHelper.ReleaseComObject(refItem);
                }
            }
        }
        catch (Exception ex) when (ex.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase)
                                    || ex.HResult == unchecked((int)0x800A0046))
        {
            _logger.Error("Cannot access VBA references — 'Trust access to VBA project' is disabled");
        }
        catch (Exception ex)
        {
            _logger.Warn($"Could not read VBA references: {ex.Message}");
        }
        finally
        {
            ComHelper.ReleaseAll(references, vbProject);
        }

        // Write database properties
        string json = JsonSerializer.Serialize(dbProps, JsonOptions);
        string filePath = Path.Combine(outputFolder, "database-properties.json");
        File.WriteAllText(filePath, json + "\n", new UTF8Encoding(false));
        _logger.Info("database-properties.json written");
    }
}
