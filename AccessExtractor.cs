using System.Runtime.InteropServices;
using MsAccessExtract.Extractors;
using MsAccessExtract.Helpers;

namespace MsAccessExtract;

/// <summary>
/// Main orchestrator — opens Access via COM and coordinates all extractors.
/// </summary>
internal sealed class AccessExtractor : IDisposable
{
    private dynamic? _accessApp;
    private readonly ConsoleLogger _logger;
    private bool _disposed;

    public AccessExtractor(ConsoleLogger logger)
    {
        _logger = logger;
    }

    public void Extract(string dbPath)
    {
        string dbName = Path.GetFileNameWithoutExtension(dbPath);
        string dbDir = Path.GetDirectoryName(dbPath)!;
        string outputRoot = Path.Combine(dbDir, $"{dbName}_src");

        _logger.Header($"Extracting: {Path.GetFileName(dbPath)}");
        _logger.Info($"Output folder: {outputRoot}");
        _logger.StartTimer();

        // Clean and recreate output folder
        if (Directory.Exists(outputRoot))
        {
            _logger.Info("Cleaning previous export...");
            Directory.Delete(outputRoot, recursive: true);
        }
        Directory.CreateDirectory(outputRoot);

        // Start Access via COM
        _logger.Info("Starting Microsoft Access...");
        StartAccess(dbPath);

        try
        {
            // Define output folders
            string modulesFolder = Path.Combine(outputRoot, "modules");
            string classesFolder = Path.Combine(outputRoot, "classes");
            string formsFolder = Path.Combine(outputRoot, "forms");
            string reportsFolder = Path.Combine(outputRoot, "reports");
            string queriesFolder = Path.Combine(outputRoot, "queries");
            string macrosFolder = Path.Combine(outputRoot, "macros");
            string tablesFolder = Path.Combine(outputRoot, "tables");
            string relationsFolder = Path.Combine(outputRoot, "relations");

            // Execute extractors in order
            new ModuleExtractor(_logger).Extract(_accessApp!, modulesFolder, classesFolder);
            new FormExtractor(_logger).Extract(_accessApp!, formsFolder);
            new ReportExtractor(_logger).Extract(_accessApp!, reportsFolder);
            new QueryExtractor(_logger).Extract(_accessApp!, queriesFolder);
            new MacroExtractor(_logger).Extract(_accessApp!, macrosFolder);
            new TableExtractor(_logger).Extract(_accessApp!, tablesFolder);
            new RelationshipExtractor(_logger).Extract(_accessApp!, relationsFolder);
            new ReferenceExtractor(_logger).Extract(_accessApp!, outputRoot, dbName);

            // Clean up empty folders
            CleanEmptyFolders(outputRoot);
        }
        finally
        {
            _logger.StopTimer();
            ShutdownAccess();
        }

        _logger.PrintSummary();

        if (!_logger.HasErrors)
        {
            _logger.Header("Export completed successfully!");
            _logger.Info($"Files written to: {outputRoot}");
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\nExport completed with errors. Check the log above.");
            Console.ResetColor();
        }
    }

    private void StartAccess(string dbPath)
    {
        try
        {
            var accessType = Type.GetTypeFromProgID("Access.Application");
            if (accessType == null)
            {
                throw new InvalidOperationException(
                    "Microsoft Access is not installed or not registered on this machine.\n" +
                    "This tool requires MS Access to extract database objects.");
            }

            _accessApp = Activator.CreateInstance(accessType);
            _accessApp!.Visible = false;
            _accessApp!.UserControl = false;

            // Open database exclusively for reliable export
            _accessApp!.OpenCurrentDatabase(dbPath, Exclusive: false);
            _logger.Info("Database opened successfully");
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x80080005))
        {
            throw new InvalidOperationException(
                "Failed to start Access. The application may be busy or blocked.\n" +
                "Please close any running Access instances and try again.", ex);
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw new InvalidOperationException(
                $"Failed to open database: {ex.Message}\n" +
                "Ensure the file is not locked by another application.", ex);
        }
    }

    private void ShutdownAccess()
    {
        if (_accessApp == null) return;

        try
        {
            _logger.Info("Closing Microsoft Access...");
            try { _accessApp!.CloseCurrentDatabase(); } catch { }
            try { _accessApp!.Quit(); } catch { }
        }
        catch { }
        finally
        {
            ComHelper.ReleaseComObject(_accessApp);
            _accessApp = null;

            // Force COM cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }

    private static void CleanEmptyFolders(string rootPath)
    {
        foreach (string dir in Directory.GetDirectories(rootPath))
        {
            if (Directory.GetFileSystemEntries(dir).Length == 0)
            {
                Directory.Delete(dir);
            }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        ShutdownAccess();
    }
}
