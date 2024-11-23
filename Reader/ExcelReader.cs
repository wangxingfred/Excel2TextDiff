using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader;

namespace Excel2TextDiff
{
    internal class ExcelReader : IReader
    {
        private readonly string _excelFile;

        private readonly List<string> _rowCells = new();

        internal ExcelReader(string excelFile)
        {
            _excelFile = excelFile;
        }

        public void Read(IVisitor visitor)
        {
            using var excelFileStream = new FileStream(_excelFile,
                FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            var ext = Path.GetExtension(_excelFile);
            using var reader = ext switch
            {
                ".csv" => ExcelReaderFactory.CreateCsvReader(excelFileStream),
                _ => ExcelReaderFactory.CreateReader(excelFileStream)
            };

            do
            {
                visitor.BeginSheet(reader.Name);
                ReadRows(reader, visitor);
                visitor.EndSheet();

            } while (reader.NextResult());
        }

        private void ReadRows(IExcelDataReader dataReader, IVisitor visitor)
        {
            while (dataReader.Read())
            {
                try
                {
                    _rowCells.Clear();
                    for (int i = 0, n = dataReader.FieldCount; i < n; i++)
                    {
                        var cell = dataReader.GetValue(i);
                        _rowCells.Add(cell.ToString());
                    }

                    visitor.VisitRow(_rowCells);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }
    }
}