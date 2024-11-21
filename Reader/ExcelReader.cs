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

        public void Read(TextWriter writer)
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
                writer.BeginSheet(reader.Name);
                ReadRows(reader, writer);
                writer.EndSheet();

            } while (reader.NextResult());
        }

        private void ReadRows(IExcelDataReader reader, TextWriter writer)
        {
            while (reader.Read())
            {
                try
                {
                    _rowCells.Clear();
                    for (int i = 0, n = reader.FieldCount; i < n; i++)
                    {
                        var cell = reader.GetValue(i);
                        _rowCells.Add(cell.ToString());
                    }

                    writer.WriteRow(_rowCells);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }
    }
}