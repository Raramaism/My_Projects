using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public static class MovementIdGenerator
{
    private static readonly Random random = new Random();

    /// <summary>
    /// Generates a highly unique ID using string.Format() for backward compatibility.
    /// This combines a millisecond-precise timestamp, random component, and GUID fragment.
    /// </summary>
    /// <returns>A unique string.</returns>
    public static string GenerateMovementId()
    {
        // 1. Get a highly precise, chronologically unique timestamp (17 digits)
        string timestampPart = DateTime.UtcNow.ToString("yyyyMMddHHmmssfff");

        // 2. Generate a small random number (4 digits)
        int randomPart = random.Next(1000, 10000);

        // 3. Get a tiny unique GUID fragment (3 characters)
        Guid guid = Guid.NewGuid();
        string uniqueFragment = guid.ToString("N").Substring(29, 3);

        // 4. Combine them using string.Format()
        // {0} replaces timestampPart, {1} replaces randomPart, {2} replaces uniqueFragment
        return string.Format("{0}-{1}-{2}", timestampPart, randomPart, uniqueFragment);
    }
}
