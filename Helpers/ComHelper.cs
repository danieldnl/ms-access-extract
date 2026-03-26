using System.Runtime.InteropServices;

namespace MsAccessExtract.Helpers;

/// <summary>
/// Manages COM object lifecycle to prevent memory leaks and orphaned processes.
/// </summary>
internal static class ComHelper
{
    public static void ReleaseComObject(object? comObject)
    {
        if (comObject == null) return;

        try
        {
            if (Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch (Exception)
        {
            // Swallow — object may already be released
        }
    }

    public static void ReleaseAll(params object?[] comObjects)
    {
        foreach (var obj in comObjects)
        {
            ReleaseComObject(obj);
        }
    }

    /// <summary>
    /// Safely get a COM property that might not exist or throw.
    /// Returns default(T) on failure.
    /// </summary>
    public static T? SafeGetProperty<T>(dynamic comObject, string propertyName, T? fallback = default)
    {
        try
        {
            var value = comObject.GetType().InvokeMember(
                propertyName,
                System.Reflection.BindingFlags.GetProperty,
                null,
                comObject,
                null);

            if (value == null) return fallback;
            return (T)Convert.ChangeType(value, typeof(T));
        }
        catch
        {
            return fallback;
        }
    }

    /// <summary>
    /// Iterate a COM collection safely, yielding each item.
    /// Handles Count/Item pattern used by DAO and VBA collections.
    /// </summary>
    public static IEnumerable<dynamic> EnumerateComCollection(dynamic collection)
    {
        int count;
        try
        {
            count = (int)collection.Count;
        }
        catch
        {
            yield break;
        }

        for (int i = 0; i < count; i++)
        {
            dynamic? item = null;
            try
            {
                item = collection[i];
            }
            catch
            {
                // Some collections are 1-based
                try
                {
                    item = collection[i + 1];
                }
                catch
                {
                    continue;
                }
            }

            if (item != null)
            {
                yield return item;
            }
        }
    }
}
