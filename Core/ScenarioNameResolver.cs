using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public static class ScenarioNameResolver
{
    public static string Resolve(RequestContract request, string fallback)
    {
        return string.IsNullOrWhiteSpace(request.ScenarioName)
            ? fallback
            : request.ScenarioName.Trim();
    }
}
