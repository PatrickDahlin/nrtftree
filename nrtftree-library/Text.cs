namespace Net.Sgoliver.NRtfTree.Util;

public static class Text
{
    
    /// <summary>
    /// Gets the hexadecimal code of an integer.
    /// </summary>
    /// <param name="code">Integer.</param>
    /// <returns>Hexadecimal code of the integer passed as a parameter.</returns>
    public static string GetHex(byte code)
    {
        var hex = Convert.ToString(code, 16);

        if (hex.Length == 1)
        {
            hex = "0" + hex;
        }

        return hex;
    }
    
}