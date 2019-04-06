using System;
using DocumentFormat.OpenXml.Packaging;
using System.Data;
using System.Drawing.Imaging;

namespace CreateDocxFromDotx
{
    public static class Utility
    {
        public static DataTable CreateTestDataTable(string tableName)
        {
            var table = new DataTable(tableName);
            var column = new DataColumn("id", typeof(int))
            {
                AutoIncrement = true
            };
            table.Columns.Add(column);

            column = new DataColumn("item", typeof(string));
            table.Columns.Add(column);

            const int rowSize = 10;
            for (var i = 0; i < rowSize; ++i)
            {
                var row = table.NewRow();
                row["item"] = "item " + i;
                table.Rows.Add(row);
            }

            table.AcceptChanges();
            return table;
        }

        public static DataTable CreateTestDataTableWith5Column(string tableName)
        {
            var table = new DataTable(tableName);
            table.Columns.Add(
                new DataColumn("id", typeof(int))
                {
                    AutoIncrement = true
                });

            for (var i = 1; i < 5; ++i)
            {
                var column = new DataColumn("column " + i, typeof(string));
                table.Columns.Add(column);
            }

            const int rowSize = 10;
            for (var i = 0; i < rowSize; ++i)
            {
                var row = table.NewRow();
                for (var j = 1; j < 5; ++j)
                    row[j] = "item_" + j + " " + i;
                table.Rows.Add(row);
            }

            table.AcceptChanges();
            return table;
        }


        private static ImagePartType ImageFormatToImagePartType(ImageFormat rawFormat)
        {
            if (ImageFormat.Jpeg.Equals(rawFormat))
            {
                return ImagePartType.Jpeg;
            }
            else if (ImageFormat.Png.Equals(rawFormat))
            {
                return ImagePartType.Png;
            }
            else if (ImageFormat.Gif.Equals(rawFormat))
            {
                return ImagePartType.Gif;
            }
            else if (ImageFormat.Bmp.Equals(rawFormat))
            {
                return ImagePartType.Bmp;
            }
            else if (ImageFormat.Emf.Equals(rawFormat))
            {
                return ImagePartType.Emf;
            }
            else if (ImageFormat.Icon.Equals(rawFormat))
            {
                return ImagePartType.Icon;
            }
            else if (ImageFormat.Tiff.Equals(rawFormat))
            {
                return ImagePartType.Tiff;
            }
            else if (ImageFormat.Wmf.Equals(rawFormat))
            {
                return ImagePartType.Wmf;
            }
            else
            {
                throw new InvalidCastException();
            }
        }
    }
}
