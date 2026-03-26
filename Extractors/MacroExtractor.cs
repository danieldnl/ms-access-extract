using MsAccessExtract.Helpers;
using MsAccessExtract.Sanitizers;

namespace MsAccessExtract.Extractors;

/// <summary>
/// Extracts macros via Application.SaveAsText with sanitization.
/// </summary>
internal sealed class MacroExtractor
{
    private readonly ConsoleLogger _logger;

    public MacroExtractor(ConsoleLogger logger)
    {
        _logger = logger;
    }

    public void Extract(dynamic accessApp, string outputFolder)
    {
        _logger.Section("Macros");
        Directory.CreateDirectory(outputFolder);

        dynamic? allMacros = null;
        try
        {
            allMacros = accessApp.CurrentProject.AllMacros;
            int count = (int)allMacros.Count;

            if (count == 0)
            {
                _logger.Info("No macros found");
                return;
            }

            for (int i = 0; i < count; i++)
            {
                dynamic? macro = null;
                try
                {
                    macro = allMacros[i];
                    string name = (string)macro.Name;
                    string filePath = Path.Combine(outputFolder, $"{SanitizeFileName(name)}.bas");

                    accessApp.SaveAsText(AccessConstants.AcMacro, name, filePath);
                    SaveAsTextSanitizer.SanitizeFile(filePath);
                    _logger.Success("Macros", name);
                }
                catch (Exception ex)
                {
                    _logger.Error($"Failed to export macro at index {i}: {ex.Message}");
                }
                finally
                {
                    ComHelper.ReleaseComObject(macro);
                }
            }
        }
        finally
        {
            ComHelper.ReleaseComObject(allMacros);
        }
    }

    private static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
