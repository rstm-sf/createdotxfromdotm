using System;
using System.IO;
using System.Linq;

namespace CreateDocxFromDotx
{
    public static class DetectImageType
    {
        private static readonly byte[] Bmp = System.Text.Encoding.ASCII.GetBytes("BM");
        private static readonly byte[] Gif = System.Text.Encoding.ASCII.GetBytes("GIF");
        private static readonly byte[] Png = { 137, 80, 78, 71 };
        private static readonly byte[] Jpeg = { 255, 216, 255 };

        private static readonly byte[] Tiff = Utility.ConvertHexStringToByteArray(string.Concat(
            "49", "20", "49"));
        private static readonly byte[] Tiff2 = Utility.ConvertHexStringToByteArray(string.Concat(
            "49", "49", "2A", "00")); // little endian
        private static readonly byte[] Tiff3 = Utility.ConvertHexStringToByteArray(string.Concat(
            "4D", "4D", "00", "2A")); // big endian
        private static readonly byte[] Tiff4 = Utility.ConvertHexStringToByteArray(string.Concat(
            "4D", "4D", "00", "2B")); // BigTIFF

        public static ImageFileType GetImageType(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            var buffer = new byte[4];
            stream.Read(buffer, 0, buffer.Length);
            stream.Position = 0;
            return GetImageType(buffer);
        }

        public static ImageFileType GetImageType(byte[] buffer)
        {
            if (Jpeg.SequenceEqual(buffer.Take(Jpeg.Length)))
            {
                return ImageFileType.Jpeg;
            }
            else if (Png.SequenceEqual(buffer.Take(Png.Length)))
            {
                return ImageFileType.Png;
            }
            else if (Gif.SequenceEqual(buffer.Take(Gif.Length)))
            {
                return ImageFileType.Gif;
            }
            else if (Bmp.SequenceEqual(buffer.Take(Bmp.Length)))
            {
                return ImageFileType.Bmp;
            }
            else if (Tiff.SequenceEqual(buffer.Take(Tiff.Length))
                     || Tiff2.SequenceEqual(buffer.Take(Tiff2.Length))
                     || Tiff3.SequenceEqual(buffer.Take(Tiff3.Length))
                     || Tiff4.SequenceEqual(buffer.Take(Tiff4.Length)))
            {
                return ImageFileType.Tiff;
            }
            else
            {
                throw new InvalidCastException();
            }
        }
    }
}
