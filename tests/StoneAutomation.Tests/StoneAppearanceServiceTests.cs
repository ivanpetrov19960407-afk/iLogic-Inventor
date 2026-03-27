using StoneDrawingAuto.Core.Services;

namespace StoneAutomation.Tests;

public class StoneAppearanceServiceTests
{
    [Fact]
    public void BuildTextureAssignment_ForMarble_ReturnsMarblePreset()
    {
        var service = new StoneAppearanceService();

        var assignment = service.BuildTextureAssignment("Белый мрамор");

        Assert.Equal("Marble", assignment.MaterialName);
        Assert.Contains("Marble_4k", assignment.TextureAssetPath, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("kShadedWithEdgesVisualStyle", assignment.IsoVisualStyle);
    }

    [Fact]
    public void BuildTextureAssignment_ForUnknown_ReturnsGranitePreset()
    {
        var service = new StoneAppearanceService();

        var assignment = service.BuildTextureAssignment("Quartzite");

        Assert.Equal("Granite", assignment.MaterialName);
        Assert.Contains("Granite_4k", assignment.TextureAssetPath, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("kHiddenLineRemovedDrawingViewStyle", assignment.ProjectionVisualStyle);
    }
}
