namespace StoneDrawingAuto.Core.Services;

/// <summary>
/// Математические и геометрические утилиты для чертежей камнеобработки.
/// Основано на <c>RKM_Utils.bas</c>: рабочая единица производства — мм,
/// а внутренние единицы Inventor Drawing API — см.
/// </summary>
public sealed class StoneMathService
{
    public const double MmToCmFactor = 0.1d;
    public const double CmToMmFactor = 10d;

    public const double A3WidthMm = 420d;
    public const double A3HeightMm = 297d;
    public const double FrameLeftMm = 20d;
    public const double FrameOtherMm = 5d;
    public const double TitleBlockWidthMm = 185d;
    public const double TitleBlockHeightMm = 55d;
    public const double DimensionToleranceMm = 0.05d;

    public double MmToCm(double valueMm) => valueMm * MmToCmFactor;

    public double CmToMm(double valueCm) => valueCm * CmToMmFactor;

    /// <summary>
    /// Создаёт точку в единицах Inventor (см) из входных значений в мм.
    /// </summary>
    public Point2dCm CreatePointFromMm(double xMm, double yMm) =>
        new(MmToCm(xMm), MmToCm(yMm));

    /// <summary>
    /// Создаёт точку напрямую в см (например, если данные уже в Inventor-единицах).
    /// </summary>
    public Point2dCm CreatePointFromCm(double xCm, double yCm) => new(xCm, yCm);

    /// <summary>
    /// Проверяет, что лист соответствует формату A3 Landscape (420x297 мм)
    /// с учётом допуска из legacy VBA.
    /// </summary>
    public SheetValidationResult ValidateA3Landscape(double widthCm, double heightCm)
    {
        var widthMm = CmToMm(widthCm);
        var heightMm = CmToMm(heightCm);

        var isMatch = NearlyEqual(widthMm, A3WidthMm, DimensionToleranceMm)
                      && NearlyEqual(heightMm, A3HeightMm, DimensionToleranceMm);

        if (isMatch)
        {
            return SheetValidationResult.Success(widthMm, heightMm);
        }

        var isPortrait = NearlyEqual(widthMm, A3HeightMm, DimensionToleranceMm)
                         && NearlyEqual(heightMm, A3WidthMm, DimensionToleranceMm);

        return isPortrait
            ? SheetValidationResult.Fail(widthMm, heightMm, "Sheet is A3 but in portrait orientation; expected landscape 420x297 mm.")
            : SheetValidationResult.Fail(widthMm, heightMm, "Sheet size is not A3 (420x297 mm).");
    }

    public bool NearlyEqual(double a, double b, double tolerance)
    {
        if (tolerance < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(tolerance), "Tolerance cannot be negative.");
        }

        return Math.Abs(a - b) <= tolerance;
    }
}

public readonly record struct Point2dCm(double X, double Y);

public sealed record SheetValidationResult(bool IsValid, double WidthMm, double HeightMm, string? Message)
{
    public static SheetValidationResult Success(double widthMm, double heightMm) =>
        new(true, widthMm, heightMm, null);

    public static SheetValidationResult Fail(double widthMm, double heightMm, string message) =>
        new(false, widthMm, heightMm, message);
}
