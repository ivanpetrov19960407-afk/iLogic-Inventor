namespace StoneDrawingAuto.Core.Abstractions;
using StoneDrawingAuto.Core.Models;

/// <summary>
/// Подбирает визуальные настройки для камня перед генерацией изометрии.
/// </summary>
public interface IStoneAppearanceService
{
    StoneTextureAssignment BuildTextureAssignment(string materialName);
}
