using MsAccessExtract.Helpers;

namespace MsAccessExtract.Extractors;

/// <summary>
/// Extracts VBA standard modules and class modules from the VBProject.
/// </summary>
internal sealed class ModuleExtractor
{
    private readonly ConsoleLogger _logger;

    public ModuleExtractor(ConsoleLogger logger)
    {
        _logger = logger;
    }

    public void Extract(dynamic accessApp, string modulesFolder, string classesFolder)
    {
        _logger.Section("VBA Modules");

        dynamic? vbProject = null;
        dynamic? vbComponents = null;

        try
        {
            vbProject = accessApp.VBE.ActiveVBProject;
            vbComponents = vbProject.VBComponents;
        }
        catch (Exception ex) when (ex.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase)
                                   || ex.HResult == unchecked((int)0x800A0046))
        {
            _logger.Error("Cannot access VBA Project. Please enable:");
            _logger.Error("  File → Options → Trust Center → Trust Center Settings");
            _logger.Error("  → Macro Settings → Trust access to the VBA project object model");
            return;
        }
        catch (Exception ex)
        {
            _logger.Error($"Failed to access VBA Project: {ex.Message}");
            return;
        }

        if (vbComponents == null)
        {
            _logger.Warn("No VBA components found");
            return;
        }

        Directory.CreateDirectory(modulesFolder);
        Directory.CreateDirectory(classesFolder);

        try
        {
            int count = (int)vbComponents.Count;

            for (int i = 1; i <= count; i++)
            {
                dynamic? component = null;
                try
                {
                    component = vbComponents.Item(i);
                    int componentType = (int)component.Type;
                    string name = (string)component.Name;

                    switch (componentType)
                    {
                        case AccessConstants.VbextCtStdModule:
                        {
                            string filePath = Path.Combine(modulesFolder, $"{SanitizeFileName(name)}.bas");
                            component.Export(filePath);
                            _logger.Success("Modules", name);
                            break;
                        }

                        case AccessConstants.VbextCtClassModule:
                        {
                            string filePath = Path.Combine(classesFolder, $"{SanitizeFileName(name)}.cls");
                            component.Export(filePath);
                            _logger.Success("Classes", name);
                            break;
                        }

                        case AccessConstants.VbextCtDocument:
                        {
                            // Document modules (behind forms/reports) — check if they have code
                            dynamic codeModule = component.CodeModule;
                            int lineCount = (int)codeModule.CountOfLines;
                            if (lineCount > 2) // More than just Option statements
                            {
                                string filePath = Path.Combine(classesFolder, $"{SanitizeFileName(name)}.cls");
                                component.Export(filePath);
                                _logger.Success("Classes", $"{name} (document module)");
                            }
                            ComHelper.ReleaseComObject(codeModule);
                            break;
                        }

                        default:
                            _logger.Skip(name, $"component type {componentType}");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error($"Failed to export component at index {i}: {ex.Message}");
                }
                finally
                {
                    ComHelper.ReleaseComObject(component);
                }
            }
        }
        finally
        {
            ComHelper.ReleaseAll(vbComponents, vbProject);
        }
    }

    private static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
