using MsAccessExtract.Helpers;
using MsAccessExtract.Sanitizers;

namespace MsAccessExtract.Extractors;

/// <summary>
/// Extracts reports via Application.SaveAsText with sanitization.
/// </summary>
internal sealed class ReportExtractor
{
    private readonly ConsoleLogger _logger;

    public ReportExtractor(ConsoleLogger logger)
    {
        _logger = logger;
    }

    public void Extract(dynamic accessApp, string outputFolder)
    {
        _logger.Section("Reports");
        Directory.CreateDirectory(outputFolder);

        dynamic? allReports = null;
        try
        {
            allReports = accessApp.CurrentProject.AllReports;
            int count = (int)allReports.Count;

            if (count == 0)
            {
                _logger.Info("No reports found");
                return;
            }

            for (int i = 0; i < count; i++)
            {
                dynamic? report = null;
                try
                {
                    report = allReports[i];
                    string name = (string)report.Name;
                    string filePath = Path.Combine(outputFolder, $"{SanitizeFileName(name)}.bas");

                    accessApp.SaveAsText(AccessConstants.AcReport, name, filePath);
                    SaveAsTextSanitizer.SanitizeFile(filePath);
                    _logger.Success("Reports", name);
                }
                catch (Exception ex)
                {
                    _logger.Error($"Failed to export report at index {i}: {ex.Message}");
                }
                finally
                {
                    ComHelper.ReleaseComObject(report);
                }
            }
        }
        finally
        {
            ComHelper.ReleaseComObject(allReports);
        }
    }

    private static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
