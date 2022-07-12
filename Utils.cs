using System.Globalization;

namespace Unidigital.Cobros;

public static class Util
{
    public static string formatNumber(decimal number)
    {
        return number.ToString("F2", new CultureInfo("es-ES"));
    }
}