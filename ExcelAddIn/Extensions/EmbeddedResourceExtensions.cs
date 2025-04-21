using System.Reflection;

namespace HcExcelAddIn.Extensions;
internal static class EmbeddedResourceExtensions
{
    /// <summary>
    /// Gets the embedded resource as a string.
    /// </summary>
    /// <param name="resourceName"></param>
    /// <returns></returns>
    public static string GetEmbeddedResource(this Assembly assembly, string resourceName)
    {
        var fullResourceName = assembly.GetManifestResourceNames().Single(str => str.EndsWith(resourceName));
        using var stream = assembly.GetManifestResourceStream(fullResourceName);
        if (stream == null) throw new InvalidOperationException($"Resource {fullResourceName} not found.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
