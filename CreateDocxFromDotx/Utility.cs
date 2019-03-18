using System.Data;

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
    }
}
