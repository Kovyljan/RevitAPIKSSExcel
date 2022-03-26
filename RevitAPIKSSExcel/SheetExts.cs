using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIKSSExcel
{
    public static class SheetExts
    {
        public static void SetCellValue<T>(this ISheet sheet, int rowIndex, int columnIndex, T value)
        {
            var cellReference = new CellReference(rowIndex, columnIndex);
            var row = sheet.GetRow(cellReference.Row);
            if (row == null)
            {
                row = sheet.CreateRow(cellReference.Row);
            }

            var cell = row.GetCell(cellReference.Col);
            if (cell == null)
            {
                cell = row.CreateCell(cellReference.Col);
            }

            if (value is string)
            {
                cell.SetCellValue((string)(object)value);
            }
            else if (value is double)
            {
                cell.SetCellValue((double)(object)value);
            }
            else if (value is int)
            {
                cell.SetCellValue((int)(object)value);
            }
        }
    }
}
