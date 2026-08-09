using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public static class CoverModeResolver
{
    public static bool ResolveHasCover(RequestContract request, Func<bool> detectCover)
    {
        return request.HasCoverOverride ?? detectCover();
    }
}
