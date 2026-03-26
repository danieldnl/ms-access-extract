using MsAccessExtract.Helpers;
using MsAccessExtract.Sanitizers;

namespace MsAccessExtract.Extractors;

/// <summary>
/// Extracts forms via Application.SaveAsText with sanitization.
/// </summary>
internal sealed class FormExtractor
{
    private readonly ConsoleLogger _logger;

    public FormExtractor(ConsoleLogger logger)
    {
        _logger = logger;
    }

    public void Extract(dynamic accessApp, string outputFolder)
    {
        _logger.Section("Forms");
        Directory.CreateDirectory(outputFolder);

        dynamic? allForms = null;
        try
        {
            allForms = accessApp.CurrentProject.AllForms;
            int count = (int)allForms.Count;

            if (count == 0)
            {
                _logger.Info("No forms found");
                return;
            }

            for (int i = 0; i < count; i++)
            {
                dynamic? form = null;
                try
                {
                    form = allForms[i];
                    string name = (string)form.Name;
                    string filePath = Path.Combine(outputFolder, $"{SanitizeFileName(name)}.bas");

                    accessApp.SaveAsText(AccessConstants.AcForm, name, filePath);
                    SaveAsTextSanitizer.SanitizeFile(filePath);
                    _logger.Success("Forms", name);
                }
                catch (Exception ex)
                {
                    _logger.Error($"Failed to export form at index {i}: {ex.Message}");
                }
                finally
                {
                    ComHelper.ReleaseComObject(form);
                }
            }
        }
        finally
        {
            ComHelper.ReleaseComObject(allForms);
        }
    }

    private static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
