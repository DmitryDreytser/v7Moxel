using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Moxel.Moxel;

using ClosedXML.Excel;

namespace Moxel
{
    public static class ExcelWriter
    {
        public static bool SaveToExcel(Moxel moxel, string filename, int formatVersion = 7)
        {
            using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
            {
                var worksheet = workbook.Worksheets.Add("Лист 1");
                for(int columnNumber = 0; columnNumber < moxel.nAllColumnCount; columnNumber++)
                {
                    double columnwidth = 40.0d;

                    if (moxel.Columns.ContainsKey(columnNumber))
                        columnwidth = (double)moxel.Columns[columnNumber].FormatCell.wWidth;
                    else
                        if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                        columnwidth = (double)moxel.DefFormat.wWidth;

                    worksheet.Column(columnNumber + 1).Width = columnwidth * 0.135;
                }   


                for (int rowNumber = 0; rowNumber < moxel.nAllRowCount; rowNumber++)
                {
                    MoxelRow Row = null;
                    if (moxel.Rows.ContainsKey(rowNumber))
                        Row = moxel.Rows[rowNumber];

                    for (int columnNumber = 0; columnNumber < moxel.nAllColumnCount; columnNumber++)
                    {
                        if(Row != null)
                            worksheet.Cell(rowNumber + 1, columnNumber + 1).Value = Row[columnNumber].Text;
                    }

                }
                workbook.SaveAs(filename);
            }
            return true;
        }
    }
}
