using StoneDrawingAuto.Core.Services;

namespace StoneAutomation.Tests;

public class StoneMathServiceTests
{
    private readonly StoneMathService _service = new();

    [Fact]
    public void NearlyEqual_NegativeTolerance_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.NearlyEqual(1d, 1d, -0.001d));
    }

    [Fact]
    public void ValidateA3Landscape_ValidLandscapeA3_ReturnsSuccess()
    {
        var result = _service.ValidateA3Landscape(42.0d, 29.7d);

        Assert.True(result.IsValid);
        Assert.Null(result.Message);
    }

    [Fact]
    public void ValidateA3Landscape_PortraitA3_ReturnsFailureWithPortraitMessage()
    {
        var result = _service.ValidateA3Landscape(29.7d, 42.0d);

        Assert.False(result.IsValid);
        Assert.NotNull(result.Message);
        Assert.Contains("portrait", result.Message!, StringComparison.OrdinalIgnoreCase);
    }
}
