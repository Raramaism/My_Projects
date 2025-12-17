using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace MPHBSMS
{
    /// <summary>
    /// Helper methods to build file paths that respect traditional MAX_PATH limits.
    /// Add this file to your project and rebuild.
    /// </summary>
    public static class FilePathHelper
    {
        // Traditional Windows MAX_PATH limitation
        private const int MAX_PATH = 260;

        // Short fallback directories to try when user-specified directory is too long
        private static readonly string[] FallbackDirs = new[]
        {
            Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        };

        // Remove invalid filename chars and collapse whitespace
        public static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder();
            foreach (char c in name)
            {
                if (Array.IndexOf(invalid, c) >= 0) continue;
                sb.Append(c);
            }
            var s = sb.ToString().Trim();
            if (string.IsNullOrEmpty(s)) return null;
            s = Regex.Replace(s, @"\s+", " ");
            return s;
        }

        /// <summary>
        /// Build a safe file path by truncating the base name if necessary so full path length does not exceed MAX_PATH.
        /// extension should include the leading dot (e.g. ".txt").
        /// </summary>
        public static string BuildSafeFilePath(string directory, string baseName, string extension)
        {
            if (string.IsNullOrEmpty(extension))
                extension = string.Empty;
            if (!string.IsNullOrEmpty(extension) && !extension.StartsWith("."))
                extension = "." + extension;

            string safeBase = SanitizeFileName(baseName) ?? string.Empty;

            // Choose working directory
            string chosenDir = directory;
            if (string.IsNullOrWhiteSpace(chosenDir))
                chosenDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            chosenDir = chosenDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            // If chosenDir leaves no room, try fallbacks
            if (ComputeAllowedBaseLength(chosenDir, extension) <= 0)
            {
                foreach (var d in FallbackDirs)
                {
                    if (string.IsNullOrEmpty(d)) continue;
                    if (ComputeAllowedBaseLength(d, extension) > 0)
                    {
                        chosenDir = d;
                        break;
                    }
                }
            }

            // Final safety fallback to first fallback dir
            if (ComputeAllowedBaseLength(chosenDir, extension) <= 0)
            {
                chosenDir = FallbackDirs[0];
            }

            int allowedLength = ComputeAllowedBaseLength(chosenDir, extension);
            if (allowedLength < 1)
            {
                // As last resort, use timestamp-only name
                string tsName = "Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                return Path.Combine(chosenDir, tsName + extension);
            }

            string finalBase;
            if (safeBase.Length > allowedLength)
            {
                finalBase = safeBase.Substring(0, allowedLength);
            }
            else if (safeBase.Length == 0)
            {
                finalBase = "Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                if (finalBase.Length > allowedLength)
                    finalBase = finalBase.Substring(0, allowedLength);
            }
            else
            {
                finalBase = safeBase;
            }

            string finalPath = Path.Combine(chosenDir, finalBase + extension);

            // Extra safety: if still too long, fallback to timestamp
            if (finalPath.Length >= MAX_PATH)
            {
                string tsName = "Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                finalPath = Path.Combine(chosenDir, tsName + extension);
                if (finalPath.Length >= MAX_PATH)
                {
                    string temp = FallbackDirs[0];
                    finalPath = Path.Combine(temp, tsName + extension);
                }
            }

            return finalPath;
        }

        // Compute how many characters are available for the base filename given a directory and extension
        private static int ComputeAllowedBaseLength(string dir, string extension)
        {
            if (string.IsNullOrEmpty(dir)) return -1;
            // account for directory + separator + extension and leave 1 char headroom
            int used = dir.Length + 1 + extension.Length;
            int allowedForBase = MAX_PATH - used - 1;
            return allowedForBase;
        }
    }
}