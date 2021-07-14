using System;
using System.IO;
using SpreadsheetLight;

namespace XlsxCreator
{
    public class Worksheet
    {
        private SLDocument sl;

        public Worksheet()
        {
            sl = new SLDocument();
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, uint Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, double Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, int Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, ushort Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, short Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, long Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, byte Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, decimal Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, ulong Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, DateTime Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, float Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, bool Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool SetCellValue(int RowIndex, int ColumnIndex, string Data)
        {
            return sl.SetCellValue(RowIndex, ColumnIndex, Data);
        }

        public bool MergeCells(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            return sl.MergeWorksheetCells(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex);
        }

        public bool SetColumnWidth(int ColumnIndex, int ColumnWidth)
        {
            return sl.SetColumnWidth(ColumnIndex, ColumnWidth);
        }

        public bool SetRowHeight(int RowIndex, int RowHeight)
        {
            return sl.SetRowHeight(RowIndex, RowHeight);
        }

        public void SaveAs(Stream OutputStream)
        {
            sl.SaveAs(OutputStream);
        }

        public void SaveAs(string FileName)
        {
            sl.SaveAs(FileName);
        }
    }
}