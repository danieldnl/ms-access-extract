using System.Text;
using MsAccessExtract.Helpers;
using MsAccessExtract.Sanitizers;

namespace MsAccessExtract.Extractors;

/// <summary>
/// Extracts queries via SaveAsText + raw SQL extraction via DAO QueryDefs.
/// </summary>
internal sealed class QueryExtractor
{
    private readonly ConsoleLogger _logger;

    public QueryExtractor(ConsoleLogger logger)
    {
        _logger = logger;
    }

    public void Extract(dynamic accessApp, string outputFolder)
    {
        _logger.Section("Queries");
        Directory.CreateDirectory(outputFolder);

        dynamic? allQueries = null;
        dynamic? currentDb = null;
        dynamic? queryDefs = null;

        try
        {
            // SaveAsText export
            allQueries = accessApp.CurrentProject.AllQueries;
            int count = (int)allQueries.Count;

            if (count == 0)
            {
                _logger.Info("No queries found");
                return;
            }

            // Get DAO for SQL extraction
            currentDb = accessApp.CurrentDb();
            queryDefs = currentDb.QueryDefs;

            for (int i = 0; i < count; i++)
            {
                dynamic? query = null;
                try
                {
                    query = allQueries[i];
                    string name = (string)query.Name;

                    // Skip system queries
                    if (name.StartsWith("~") || name.StartsWith("MSys", StringComparison.OrdinalIgnoreCase))
                    {
                        _logger.Skip(name, "system query");
                        continue;
                    }

                    string safeName = SanitizeFileName(name);

                    // Export SaveAsText version (full definition)
                    string basFilePath = Path.Combine(outputFolder, $"{safeName}.bas");
                    accessApp.SaveAsText(AccessConstants.AcQuery, name, basFilePath);
                    SaveAsTextSanitizer.SanitizeFile(basFilePath);

                    // Export raw SQL
                    try
                    {
                        dynamic queryDef = queryDefs[name];
                        string sql = (string)queryDef.SQL;
                        if (!string.IsNullOrWhiteSpace(sql))
                        {
                            string sqlFilePath = Path.Combine(outputFolder, $"{safeName}.sql");
                            File.WriteAllText(sqlFilePath, sql.TrimEnd() + "\n", new UTF8Encoding(false));
                        }
                        ComHelper.ReleaseComObject(queryDef);
                    }
                    catch
                    {
                        // SQL extraction is a bonus — don't fail the whole query
                    }

                    _logger.Success("Queries", name);
                }
                catch (Exception ex)
                {
                    _logger.Error($"Failed to export query at index {i}: {ex.Message}");
                }
                finally
                {
                    ComHelper.ReleaseComObject(query);
                }
            }
        }
        finally
        {
            ComHelper.ReleaseAll(queryDefs, currentDb, allQueries);
        }
    }

    private static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
