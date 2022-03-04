using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace ExcelUtilities
{
  /// <summary>
  /// Class containing helper methods for imports purposes.
  /// </summary>
  public static class ImportHelper
  {
    /// <summary>
    /// Compresses the object into a Base64 gzipped string.
    /// </summary>
    /// <typeparam name="T">The type of the input object.</typeparam>
    /// <param name="objectToCompress">The object to compress.</param>
    /// <returns>The compressed object.</returns>
    public static string CompressObject<T>(T objectToCompress)
    {
      return Convert.ToBase64String(ObjectCompressionHelper.CompressObject(objectToCompress));
    }

    /// <summary>
    /// Unzip the object from a Base64 gzipped string.
    /// </summary>
    /// <typeparam name="T">The type of the outputted object.</typeparam>
    /// <param name="objectToUnzip">The string to unzip.</param>
    /// <returns>The unzip object.</returns>
    public static T DecompressObject<T>(string objectToUnzip)
    {
      return ObjectCompressionHelper.DecompressObject<T>(Convert.FromBase64String(objectToUnzip));
    }

    /// <summary>
    /// Computes a string's hash.
    /// </summary>
    /// <param name="stringToHash">The string to hash.</param>
    /// <returns>The SHA-256 hash.</returns>
    public static string ComputeHash(string stringToHash)
    {
      StringBuilder stringBuilder = new StringBuilder();
      using (var hash = SHA256.Create())
      {
        Encoding encoding = Encoding.UTF8;
        byte[] result = hash.ComputeHash(encoding.GetBytes(stringToHash));
        foreach (byte b in result)
        {
          stringBuilder.Append(b.ToString("x2", CultureInfo.InvariantCulture));
        }
      }

      return stringBuilder.ToString();
    }
  }
}