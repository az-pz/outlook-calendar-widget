using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.IO.Compression;
using System.Text;

namespace OutlookCalendarWidget.Core;

/// <summary>
/// The colored free/busy bar drawn next to each event, rendered as a tiny PNG data URI. Adaptive Card hosts pad
/// styled containers, so an image is the only way to get a consistent thin bar on every renderer.
/// </summary>
public static class StatusBar
{
    private const int Width = 8;
    private const int Height = 72;

    private static readonly ConcurrentDictionary<string, string> Cache = new(StringComparer.Ordinal);
    private static readonly uint[] CrcTable = BuildCrcTable();

    /// <summary>Outlook-like ARGB color for an event's <see cref="CardEvent.StatusStyle"/>.</summary>
    public static uint Color(string statusStyle) => statusStyle switch
    {
        "warning" => 0xFF0F6CBD,   // tentative: striped blue
        "attention" => 0xFF8764B8, // out of office
        "good" => 0xFF038387,      // working elsewhere
        "emphasis" => 0xFFB3B0AD,  // free
        "default" => 0xFF8A8886,   // declined
        _ => 0xFF0F6CBD,           // busy
    };

    public static string DataUri(string statusStyle) => Cache.GetOrAdd(
        statusStyle,
        style => "data:image/png;base64," + Convert.ToBase64String(RenderPng(Color(style), striped: style == "warning")));

    internal static byte[] RenderPng(uint argb, bool striped)
    {
        const double radius = Width / 2.0;
        var alpha = (byte)(argb >> 24);
        var red = (byte)(argb >> 16);
        var green = (byte)(argb >> 8);
        var blue = (byte)argb;

        // Rows of RGBA pixels, each prefixed with PNG filter type 0 (none). The shape is an anti-aliased capsule.
        var raw = new byte[Height * (1 + (Width * 4))];
        var i = 0;
        for (var y = 0; y < Height; y++)
        {
            raw[i++] = 0;
            for (var x = 0; x < Width; x++)
            {
                var px = x + 0.5;
                var py = y + 0.5;
                var cy = Math.Clamp(py, radius, Height - radius);
                var distance = Math.Sqrt(((px - radius) * (px - radius)) + ((py - cy) * (py - cy)));
                var coverage = Math.Clamp(radius - distance + 0.5, 0, 1);
                if (striped && (x + y) / 4 % 2 == 1)
                {
                    coverage *= 0.4;
                }

                raw[i++] = red;
                raw[i++] = green;
                raw[i++] = blue;
                raw[i++] = (byte)Math.Round(alpha * coverage);
            }
        }

        using var png = new MemoryStream();
        png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

        var header = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(header, Width);
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), Height);
        header[8] = 8; // bit depth
        header[9] = 6; // color type: RGBA
        WriteChunk(png, "IHDR", header);

        using (var compressed = new MemoryStream())
        {
            using (var zlib = new ZLibStream(compressed, CompressionLevel.SmallestSize, leaveOpen: true))
            {
                zlib.Write(raw);
            }

            WriteChunk(png, "IDAT", compressed.ToArray());
        }

        WriteChunk(png, "IEND", []);
        return png.ToArray();
    }

    private static void WriteChunk(Stream stream, string type, byte[] data)
    {
        Span<byte> buffer = stackalloc byte[4];
        BinaryPrimitives.WriteInt32BigEndian(buffer, data.Length);
        stream.Write(buffer);

        var typeBytes = Encoding.ASCII.GetBytes(type);
        stream.Write(typeBytes);
        stream.Write(data);

        var crc = Crc32(Crc32(0xFFFFFFFF, typeBytes), data) ^ 0xFFFFFFFF;
        BinaryPrimitives.WriteUInt32BigEndian(buffer, crc);
        stream.Write(buffer);
    }

    private static uint Crc32(uint crc, byte[] data)
    {
        foreach (var b in data)
        {
            crc = CrcTable[(crc ^ b) & 0xFF] ^ (crc >> 8);
        }

        return crc;
    }

    private static uint[] BuildCrcTable()
    {
        var table = new uint[256];
        for (uint n = 0; n < 256; n++)
        {
            var c = n;
            for (var k = 0; k < 8; k++)
            {
                c = (c & 1) != 0 ? 0xEDB88320 ^ (c >> 1) : c >> 1;
            }

            table[n] = c;
        }

        return table;
    }
}
