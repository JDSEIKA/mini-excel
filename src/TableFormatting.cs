using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace CalcRapido
{
    /// <summary>
    /// Estilo visual de tabla leído de styles.xml. Es una capa de
    /// presentación: no modifica los estilos de celda del libro.
    /// </summary>
    public sealed class TableStyleDefinition
    {
        public CellStyle WholeTable;
        public CellStyle HeaderRow;
        public CellStyle TotalRow;
        public CellStyle FirstRowStripe;
        public CellStyle FirstColumn;
        public CellStyle LastColumn;
        public CellStyle LastHeaderCell;
        public CellStyle LastTotalCell;
    }

    public sealed class TableFormat
    {
        public int Row1;
        public int Col1;
        public int Row2;
        public int Col2;
        public int HeaderRows = 1;
        public int TotalsRows;
        public bool ShowRowStripes;
        public bool ShowFirstColumn;
        public bool ShowLastColumn;
        public string StyleName;
        public Sheet Sheet;
        public string Name;
        public string SourcePart;
        public TableStyleDefinition Style;

        public TableFormat Clone()
        {
            return (TableFormat)MemberwiseClone();
        }
    }

    internal static class TableFormatting
    {
        public static TableFormat Find(Sheet sheet, int row, int col)
        {
            if (sheet == null) return null;
            foreach (TableFormat t in sheet.TableFormats)
                if (row >= t.Row1 && row <= t.Row2 && col >= t.Col1 && col <= t.Col2)
                    return t;
            return null;
        }

        public static void ExpandForCell(Sheet sheet, int row, int col)
        {
            if (sheet == null) return;
            foreach (TableFormat t in sheet.TableFormats)
            {
                if (row == t.Row2 + 1 && col >= t.Col1 && col <= t.Col2) t.Row2 = row;
                if (col == t.Col2 + 1 && row >= t.Row1 && row <= t.Row2) t.Col2 = col;
            }
        }

        public static Color Fill(TableFormat table, int row, int col)
        {
            if (table == null || table.Style == null) return Color.Empty;
            TableStyleDefinition s = table.Style;
            int headerEnd = table.Row1 + Math.Max(0, table.HeaderRows);
            int dataStart = headerEnd;
            int dataEnd = table.Row2 - Math.Max(0, table.TotalsRows);
            bool lastCol = table.ShowLastColumn && col == table.Col2;
            bool firstCol = table.ShowFirstColumn && col == table.Col1;
            if (row < headerEnd && s.HeaderRow != null && !s.HeaderRow.FillColor.IsEmpty)
                return s.HeaderRow.FillColor;
            if (row > dataEnd && s.TotalRow != null && !s.TotalRow.FillColor.IsEmpty)
                return s.TotalRow.FillColor;
            if (s.WholeTable != null && !s.WholeTable.FillColor.IsEmpty)
                return s.WholeTable.FillColor;
            if (row >= headerEnd && row <= dataEnd)
            {
                if (lastCol && s.LastColumn != null && !s.LastColumn.FillColor.IsEmpty)
                    return s.LastColumn.FillColor;
                if (firstCol && s.FirstColumn != null && !s.FirstColumn.FillColor.IsEmpty)
                    return s.FirstColumn.FillColor;
                if (table.ShowRowStripes && s.FirstRowStripe != null &&
                    !s.FirstRowStripe.FillColor.IsEmpty &&
                    ((row - dataStart) % 2) == 0)
                    return s.FirstRowStripe.FillColor;
            }
            if (lastCol && s.LastColumn != null && !s.LastColumn.FillColor.IsEmpty)
                return s.LastColumn.FillColor;
            if (firstCol && s.FirstColumn != null && !s.FirstColumn.FillColor.IsEmpty)
                return s.FirstColumn.FillColor;
            return Color.Empty;
        }
    }

    internal static class PivotEngine
    {
        public static bool Refresh(Sheet sheet, PivotRange pivot)
        {
            return RefreshCore(sheet, pivot, -1, null);
        }

        public static bool RefreshFiltered(Sheet sheet, PivotRange pivot, int field, string value)
        {
            return RefreshCore(sheet, pivot, field, value);
        }

        private static bool RefreshCore(Sheet sheet, PivotRange pivot, int filterField, string filterValue)
        {
            if (sheet == null || pivot == null || pivot.Cache == null) return false;
            if (pivot.ColField >= 0) return RefreshColumns(sheet, pivot, filterField, filterValue);
            List<int> rowFields = new List<int>();
            if (pivot.RowFields.Count > 0) rowFields.AddRange(pivot.RowFields);
            else rowFields.Add(pivot.RowField);
            int visibleRowFields = Math.Min(rowFields.Count, Math.Max(1, pivot.FirstDataCol));
            if (visibleRowFields > 0 && rowFields.Count > visibleRowFields)
                rowFields.RemoveRange(visibleRowFields, rowFields.Count - visibleRowFields);
            foreach (int f in rowFields)
                if (f < 0 || f >= pivot.Cache.FieldNames.Count) return false;
            List<string> order = new List<string>();
            Dictionary<string, string[]> keys = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, double[]> groups =
                new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);
            foreach (List<CellValue> record in pivot.Cache.Rows)
            {
                bool valid = true;
                foreach (int f in rowFields) if (f >= record.Count) { valid = false; break; }
                if (!valid) continue;
                if (filterField >= 0 && filterField < record.Count &&
                    !string.Equals(record[filterField].ToDisplayString(), filterValue,
                        StringComparison.OrdinalIgnoreCase)) continue;
                StringBuilder keyBuilder = new StringBuilder();
                string[] values = new string[rowFields.Count];
                for (int i = 0; i < rowFields.Count; i++)
                {
                    values[i] = record[rowFields[i]].ToDisplayString();
                    keyBuilder.Append(values[i]).Append('\u001f');
                }
                string key = keyBuilder.ToString();
                double[] sums;
                if (!groups.TryGetValue(key, out sums))
                {
                    sums = new double[Math.Max(1, pivot.DataFields.Count)];
                    groups[key] = sums;
                    keys[key] = values;
                    order.Add(key);
                }
                for (int i = 0; i < pivot.DataFields.Count; i++)
                {
                    int field = pivot.DataFields[i];
                    if (field < 0 || field >= record.Count || record[field].Kind != ValueKind.Number)
                        continue;
                    sums[i] += record[field].Number;
                }
            }
            for (int r = pivot.Row1; r <= pivot.Row2; r++)
                for (int c = pivot.Col1; c <= pivot.Col2; c++) sheet.ClearContents(r, c);
            int row = pivot.Row1;
            int dataStart = pivot.Col1 + rowFields.Count;
            if (pivot.ShowColHeaders)
            {
                for (int i = 0; i < rowFields.Count; i++)
                {
                    string caption = i == 0 && !string.IsNullOrEmpty(pivot.RowCaption)
                        ? pivot.RowCaption : pivot.Cache.FieldNames[rowFields[i]];
                    sheet.SetTextLiteral(row, pivot.Col1 + i, caption);
                }
                for (int i = 0; i < pivot.DataNames.Count && dataStart + i <= pivot.Col2; i++)
                    sheet.SetTextLiteral(row, dataStart + i, pivot.DataNames[i]);
                row++;
            }
            foreach (string key in order)
            {
                if (row > pivot.Row2) break;
                string[] values = keys[key];
                for (int i = 0; i < values.Length && pivot.Col1 + i <= pivot.Col2; i++)
                    sheet.SetTextLiteral(row, pivot.Col1 + i, values[i]);
                double[] sums = groups[key];
                for (int i = 0; i < sums.Length && dataStart + i <= pivot.Col2; i++)
                    sheet.SetNumber(row, dataStart + i, sums[i]);
                row++;
            }
            if (row <= pivot.Row2)
            {
                sheet.SetTextLiteral(row, pivot.Col1, "Total general");
                for (int i = 0; i < pivot.DataFields.Count && dataStart + i <= pivot.Col2; i++)
                {
                    double total = 0;
                    foreach (double[] sums in groups.Values) total += sums[i];
                    sheet.SetNumber(row, dataStart + i, total);
                }
            }
            sheet.RecalculateUsedRange();
            return true;
        }

        private static bool RefreshColumns(Sheet sheet, PivotRange pivot, int filterField, string filterValue)
        {
            List<int> rowFields = new List<int>();
            if (pivot.RowFields.Count > 0) rowFields.AddRange(pivot.RowFields);
            else rowFields.Add(pivot.RowField);
            int visible = Math.Min(rowFields.Count, Math.Max(1, pivot.FirstDataCol));
            if (visible <= 0) return false;
            if (rowFields.Count > visible) rowFields.RemoveRange(visible, rowFields.Count - visible);
            if (pivot.ColField < 0 || pivot.ColField >= pivot.Cache.FieldNames.Count) return false;
            List<string> colOrder = new List<string>();
            HashSet<string> colSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (List<CellValue> record in pivot.Cache.Rows)
                if (pivot.ColField < record.Count)
                {
                    string v = record[pivot.ColField].ToDisplayString();
                    if (colSeen.Add(v)) colOrder.Add(v);
                }
            List<string> rowOrder = new List<string>();
            Dictionary<string, string[]> rowKeys = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, double[]> cells = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);
            foreach (List<CellValue> record in pivot.Cache.Rows)
            {
                if (filterField >= 0 && filterField < record.Count &&
                    !string.Equals(record[filterField].ToDisplayString(), filterValue,
                        StringComparison.OrdinalIgnoreCase)) continue;
                bool valid = true;
                foreach (int f in rowFields) if (f >= record.Count) { valid = false; break; }
                if (!valid || pivot.ColField >= record.Count) continue;
                StringBuilder rk = new StringBuilder();
                string[] values = new string[rowFields.Count];
                for (int i = 0; i < rowFields.Count; i++)
                {
                    values[i] = record[rowFields[i]].ToDisplayString();
                    rk.Append(values[i]).Append('\u001f');
                }
                string rowKey = rk.ToString();
                if (!rowKeys.ContainsKey(rowKey))
                {
                    rowKeys[rowKey] = values;
                    rowOrder.Add(rowKey);
                }
                string col = record[pivot.ColField].ToDisplayString();
                string key = rowKey + "\u001e" + col;
                double[] sums;
                if (!cells.TryGetValue(key, out sums))
                {
                    sums = new double[Math.Max(1, pivot.DataFields.Count)];
                    cells[key] = sums;
                }
                for (int i = 0; i < pivot.DataFields.Count; i++)
                {
                    int field = pivot.DataFields[i];
                    if (field >= 0 && field < record.Count && record[field].Kind == ValueKind.Number)
                        sums[i] += record[field].Number;
                }
            }
            for (int r = pivot.Row1; r <= pivot.Row2; r++)
                for (int c = pivot.Col1; c <= pivot.Col2; c++) sheet.ClearContents(r, c);
            int dataStart = pivot.Col1 + rowFields.Count;
            int stride = Math.Max(1, pivot.DataNames.Count);
            int row = pivot.Row1;
            if (pivot.ShowColHeaders)
            {
                for (int i = 0; i < rowFields.Count; i++)
                    sheet.SetTextLiteral(row, pivot.Col1 + i,
                        i == 0 && !string.IsNullOrEmpty(pivot.RowCaption)
                            ? pivot.RowCaption : pivot.Cache.FieldNames[rowFields[i]]);
                for (int ci = 0; ci < colOrder.Count; ci++)
                    for (int di = 0; di < pivot.DataNames.Count; di++)
                    {
                        int c = dataStart + ci * stride + di;
                        if (c <= pivot.Col2)
                            sheet.SetTextLiteral(row, c, colOrder[ci] + " " + pivot.DataNames[di]);
                    }
                row++;
            }
            foreach (string rowKey in rowOrder)
            {
                if (row > pivot.Row2) break;
                string[] values = rowKeys[rowKey];
                for (int i = 0; i < values.Length && pivot.Col1 + i <= pivot.Col2; i++)
                    sheet.SetTextLiteral(row, pivot.Col1 + i, values[i]);
                for (int ci = 0; ci < colOrder.Count; ci++)
                {
                    double[] sums;
                    if (!cells.TryGetValue(rowKey + "\u001e" + colOrder[ci], out sums)) continue;
                    for (int di = 0; di < sums.Length; di++)
                    {
                        int c = dataStart + ci * stride + di;
                        if (c <= pivot.Col2) sheet.SetNumber(row, c, sums[di]);
                    }
                }
                row++;
            }
            sheet.RecalculateUsedRange();
            return true;
        }
    }

    internal static class PivotFormatting
    {
        public static PivotRange Find(Sheet sheet, int row, int col)
        {
            if (sheet == null) return null;
            foreach (PivotRange p in sheet.PivotRanges)
                if (row >= p.Row1 && row <= p.Row2 && col >= p.Col1 && col <= p.Col2) return p;
            return null;
        }

        public static Color Fill(PivotRange pivot, int row, int col)
        {
            if (pivot == null) return Color.Empty;
            if (pivot.ShowColHeaders && row == pivot.Row1 && pivot.Col2 > pivot.Col1)
                return Color.FromArgb(189, 215, 238);
            if (pivot.ShowRowHeaders && col == pivot.Col1 && row > pivot.Row1 &&
                pivot.Row2 > pivot.Row1)
                return Color.FromArgb(217, 226, 243);
            if (pivot.ShowRowStripes && row > pivot.Row1 && ((row - pivot.Row1) % 2) == 0)
                return Color.FromArgb(242, 242, 242);
            return Color.Empty;
        }
    }
}
