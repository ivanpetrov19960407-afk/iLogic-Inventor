using StoneDrawingAuto.Core.Abstractions;
using StoneDrawingAuto.Core.Models;

namespace StoneDrawingAuto.Core.Services;

/// <summary>
/// Рекомендуемые пресеты текстур/стилей для изометрии и проекционных видов.
/// </summary>
public sealed class StoneAppearanceService : IStoneAppearanceService
{
    public StoneTextureAssignment BuildTextureAssignment(string materialName)
    {
        var normalized = (materialName ?? string.Empty).Trim().ToLowerInvariant();

        if (normalized.Contains("marble") || normalized.Contains("мрамор"))
        {
            return new StoneTextureAssignment(
                MaterialName: "Marble",
                TextureAssetPath: "Assets/Textures/Stone/Marble_4k.jpg",
                IsoVisualStyle: "kShadedWithEdgesVisualStyle",
                ProjectionVisualStyle: "kHiddenLineRemovedDrawingViewStyle");
        }

        return new StoneTextureAssignment(
            MaterialName: "Granite",
            TextureAssetPath: "Assets/Textures/Stone/Granite_4k.jpg",
            IsoVisualStyle: "kShadedWithEdgesVisualStyle",
            ProjectionVisualStyle: "kHiddenLineRemovedDrawingViewStyle");
    }
}
