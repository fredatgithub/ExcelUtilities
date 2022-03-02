using GroBuf;
using GroBuf.DataMembersExtracters;
using System.IO;
using System.IO.Compression;

namespace ExcelUtilities
{
  /// <summary>
  /// Class for compressing/uncompressing objects.
  /// </summary>
  public static class ObjectCompressionHelper
  {
    /// <summary>
    /// The serializer.
    /// </summary>
    private static readonly Serializer InternalSerializer = new Serializer(new PropertiesExtractor(), options: GroBufOptions.WriteEmptyObjects);

    /// <summary>
    /// Compress the object into a byte array.
    /// </summary>
    /// <typeparam name="T">The type of the input object.</typeparam>
    /// <param name="objectToCompress">The object to compress.</param>
    /// <returns>The compressed object.</returns>
    public static byte[] CompressObject<T>(T objectToCompress)
    {
      byte[] flow = InternalSerializer.Serialize(objectToCompress);
      byte[] compressed;
      using (var outputStream = new MemoryStream())
      {
        using (var gzipStream = new DeflateStream(outputStream, CompressionLevel.Fastest))
        {
          gzipStream.Write(flow, 0, flow.Length);
        }

        compressed = outputStream.ToArray();
      }

      return compressed;
    }

    /// <summary>
    /// Decompresses the object from byte array.
    /// </summary>
    /// <typeparam name="T">The type of the outputted object.</typeparam>
    /// <param name="objectToDecompress">The byte array to decompress.</param>
    /// <returns>The decompressed object.</returns>
    public static T DecompressObject<T>(byte[] objectToDecompress)
    {
      T flow = default(T);

      using (var inputStream = new MemoryStream(objectToDecompress))
      {
        inputStream.Position = 0;
        using (var outputStream = new MemoryStream())
        {
          using (var gzipStream = new DeflateStream(inputStream, CompressionMode.Decompress))
          {
            gzipStream.CopyTo(outputStream);
          }

          flow = InternalSerializer.Deserialize<T>(outputStream.ToArray());
        }
      }

      return flow;
    }
  }
}