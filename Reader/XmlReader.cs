using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;


namespace Excel2TextDiff
{
    internal class XmlReader : IReader
    {
        private XmlElements.Workbook Workbook { get; set; }

        private XmlDocument XmlDocument { get; set; }

        internal XmlReader(string filePath)
        {
            var xml = new XmlDocument();
            using Stream s = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            xml.Load(s);

            var workbook = new XmlElements.Workbook();
            if (!workbook.Parse(xml.DocumentElement)) throw new Exception("Failed to parse workbook");

            Workbook = workbook;
            XmlDocument = xml;
        }

        public void Read(TextWriter writer)
        {
            var rowCells = new List<string>();

            foreach (var sheet in Workbook.Worksheets)
            {
                writer.BeginSheet(sheet.Name);

                foreach (var row in sheet.Table.Rows)
                {
                    rowCells.Clear();
                    foreach (var cell in row.FilledCells)
                    {
                        rowCells.Add(cell.ReadStr());
                    }

                    writer.WriteRow(rowCells);
                }

                writer.EndSheet();
            }
        }
    }
}