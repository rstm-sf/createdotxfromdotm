using System;
using System.IO;
using System.Linq;

namespace CreateDocxFromDotx
{
    public static class DetectImageType
    {
        private static readonly byte[] Bmp = Utility.ConvertHexStringToByteArray(string.Concat(
            "42", "4D")); // BM

        private static readonly byte[] Gif = Utility.ConvertHexStringToByteArray(string.Concat(
            "47", "49", "46", "38")); // first 4 bytes of GIF87a or GIF89a

        private static readonly byte[] Png = Utility.ConvertHexStringToByteArray(string.Concat(
            "89", "50", "4E", "47", "0D", "0A", "1A", "0A"));

        private static readonly byte[] Tiff = Utility.ConvertHexStringToByteArray(string.Concat(
            "49", "20", "49"));
        private static readonly byte[] Tiff2 = Utility.ConvertHexStringToByteArray(string.Concat(
            "49", "49", "2A", "00")); // little endian
        private static readonly byte[] Tiff3 = Utility.ConvertHexStringToByteArray(string.Concat(
            "4D", "4D", "00", "2A")); // big endian
        private static readonly byte[] Tiff4 = Utility.ConvertHexStringToByteArray(string.Concat(
            "4D", "4D", "00", "2B")); // BigTIFF

        private static readonly byte[] Icon = Utility.ConvertHexStringToByteArray(string.Concat(
            "00", "00", "01", "00")); // Windows icon file

        private static readonly byte[] Pcx0 = Utility.ConvertHexStringToByteArray(string.Concat(
            "0A", "00", "01")); // Version 00
        private static readonly byte[] Pcx2 = Utility.ConvertHexStringToByteArray(string.Concat(
            "0A", "02", "01")); // Version 02
        private static readonly byte[] Pcx3 = Utility.ConvertHexStringToByteArray(string.Concat(
            "0A", "03", "01")); // Version 03
        private static readonly byte[] Pcx4 = Utility.ConvertHexStringToByteArray(string.Concat(
            "0A", "04", "01")); // Version 04
        private static readonly byte[] Pcx5 = Utility.ConvertHexStringToByteArray(string.Concat(
            "0A", "05", "01")); // Version 05

        private static readonly byte[] Jpeg = Utility.ConvertHexStringToByteArray(string.Concat(
            "FF", "D8"));

        private static readonly byte[] Emf = Utility.ConvertHexStringToByteArray(string.Concat(
            "01", "00", "00", "00"));

        public static ImageFileType GetImageType(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            var buffer = new byte[8];
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
            else if (Emf.SequenceEqual(buffer.Take(Emf.Length)))
            {
                return ImageFileType.Emf;
            }
            else if (Tiff.SequenceEqual(buffer.Take(Tiff.Length))
                     || Tiff2.SequenceEqual(buffer.Take(Tiff2.Length))
                     || Tiff3.SequenceEqual(buffer.Take(Tiff3.Length))
                     || Tiff4.SequenceEqual(buffer.Take(Tiff4.Length)))
            {
                return ImageFileType.Tiff;
            }
            else if (Icon.SequenceEqual(buffer.Take(Icon.Length)))
            {
                return ImageFileType.Icon;
            }
            else if (Pcx0.SequenceEqual(buffer.Take(Pcx0.Length))
                     || Pcx2.SequenceEqual(buffer.Take(Pcx2.Length))
                     || Pcx3.SequenceEqual(buffer.Take(Pcx3.Length))
                     || Pcx4.SequenceEqual(buffer.Take(Pcx4.Length))
                     || Pcx5.SequenceEqual(buffer.Take(Pcx5.Length)))
            {
                return ImageFileType.Pcx;
            }
            else
            {
                throw new InvalidCastException();
            }
        }
    }
}
