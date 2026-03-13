namespace StoneAutomation.Core;

/// <summary>
/// Чистые математические утилиты для домена камнеобработки.
/// Inventor в чертежах использует сантиметры, поэтому базовый пересчёт мм -> см = 0.1.
/// </summary>
public static class StoneMath
{
    /// <summary>
    /// Коэффициент пересчёта мм в см для Inventor API.
    /// </summary>
    public const decimal MmToCmFactor = 0.1m;

    /// <summary>
    /// Коэффициент пересчёта см в мм.
    /// </summary>
    public const decimal CmToMmFactor = 10m;

    /// <summary>
    /// Стандартная допусковая погрешность (мм), использовавшаяся в VBA.
    /// </summary>
    public const decimal DimensionToleranceMm = 0.05m;

    public static decimal MmToCm(decimal valueMm) => valueMm * MmToCmFactor;

    public static double MmToCm(double valueMm) => valueMm * (double)MmToCmFactor;

    public static decimal CmToMm(decimal valueCm) => valueCm * CmToMmFactor;

    public static double CmToMm(double valueCm) => valueCm * (double)CmToMmFactor;

    /// <summary>
    /// Сравнение линейных размеров в мм с учётом допуска.
    /// </summary>
    public static bool NearlyEqualMm(decimal aMm, decimal bMm, decimal toleranceMm = DimensionToleranceMm)
    {
        if (toleranceMm < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(toleranceMm), "Tolerance cannot be negative.");
        }

        return Math.Abs(aMm - bMm) <= toleranceMm;
    }

    /// <summary>
    /// Безопасный расчёт масштаба для fit-логики: доступный размер / размер модели.
    /// Возвращает 0, если размер модели меньше или равен 0.
    /// </summary>
    public static decimal SafeScale(decimal availableSizeMm, decimal modelSizeMm)
    {
        if (availableSizeMm <= 0 || modelSizeMm <= 0)
        {
            return 0;
        }

        return availableSizeMm / modelSizeMm;
    }
}
