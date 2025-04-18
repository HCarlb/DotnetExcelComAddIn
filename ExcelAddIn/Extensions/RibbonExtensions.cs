using System.Reflection;


namespace HcExcelAddIn.Extensions;
internal static class RibbonExtensions
{
    public static string GetRibbonXML(string name, string resourcePath)
    {
        return name.GetRibbonResourceName(resourcePath).GetEmbeddedResource();
    }

    private static string GetRibbonResourceName(this string name, string resourcePath)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceNames = assembly.GetManifestResourceNames();
        return resourceNames.Single(str => str.EndsWith(name) && str.Contains(resourcePath));
    }

    private static string GetEmbeddedResource(this string resourceName)
    {
        // Utility to retrieve the embedded resource as a string

        Log.Debug($"Loading embedded resource: {resourceName}");
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
        if (stream == null)
        {
            Log.Error("Resource '{0}' not found.", resourceName);
            return string.Empty;
        }

        Log.Debug("Resource '{0}' found.", resourceName);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
