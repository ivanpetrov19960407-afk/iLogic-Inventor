namespace StoneDrawingAuto.Core.Models;

/// <summary>
/// Целевая текстура/стиль отображения для подготовки листа альбома.
/// </summary>
public sealed record StoneTextureAssignment(
    string MaterialName,
    string TextureAssetPath,
    string IsoVisualStyle,
    string ProjectionVisualStyle);
