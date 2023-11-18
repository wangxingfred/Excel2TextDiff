using System;
using ExcelDataReader;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Excel2TextDiff
{
    class Excel2TextWriter
    {
        public void TransformToTextAndSave(string excelFile, string outputTextFile)
        {
            var lines = new List<string>();
            using var excelFileStream = new FileStream(excelFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            string ext = Path.GetExtension(excelFile);
            using (var reader = ext != ".csv" ? ExcelReaderFactory.CreateReader(excelFileStream) : ExcelReaderFactory.CreateCsvReader(excelFileStream))
            {
                do
                {
                    lines.Add($"==========  [{reader.Name ?? ""}]   ==========");
                    LoadRows(reader, lines);
                } while (reader.NextResult());
            }

            File.WriteAllLines(outputTextFile, lines, System.Text.Encoding.UTF8);
        }

        private void LoadRows(IExcelDataReader reader, List<string> lines)
        {
            var row = new List<string>();
            bool excelHeader = true;

            var rowMaxWidth = new List<int>();
            List<List<string>> rows = new List<List<string>>();
            while (reader.Read())
            {
                try
                {
                    row.Clear();
                    for (int i = 0, n = reader.FieldCount; i < n; i++)
                    {
                        object cell = reader.GetValue(i);
                        row.Add(cell != null ? cell.ToString() : "");
                    }

                    // 只保留到最后一个非空白单元格
                    int lastNotEmptyIndex = row.FindLastIndex(s => !string.IsNullOrEmpty(s));
                    if (lastNotEmptyIndex >= 0)
                    {
                        var listStr = row.GetRange(0, lastNotEmptyIndex + 1);
                        if (excelHeader)
                        {
                            excelHeader = false;
                        }

                        if (rowMaxWidth.Count <= 0)
                        {
                            for (int i = 0; i < listStr.Count; i++)
                            {
                                var str = listStr[i];
                                rowMaxWidth.Add(str.Length);
                            }
                        }

                        List<string> result = new List<string>();
                        for (int i = 0; i < listStr.Count; i++)
                        {
                            var str = listStr[i];
                            var totalCnt = str.Length;
                            var chineseCnt = GetChineseCnt(str);
                            var englishCnt = totalCnt - chineseCnt;
                            var cnt = englishCnt + 2 * chineseCnt;
                            if (i < rowMaxWidth.Count)
                            {
                                int currentRowMaxWidth = 10;
                                currentRowMaxWidth = rowMaxWidth[i];
                                rowMaxWidth[i] = cnt > currentRowMaxWidth ? cnt : currentRowMaxWidth;
                            }

                            result.Add(str);
                        }

                        rows.Add(result);
                    }
                    else
                    {
                        // 忽略空白行，没必要diff这个
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    continue;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                try
                {
                    var rowList = rows[i];
                    List<string> result = new List<string>();
                    for (int j = 0; j < rowList.Count; j++)
                    {
                        var str = rowList[j];

                        var chineseCnt = GetChineseCnt(str);
                        // if (chineseCnt > 12)
                        // {
                        //     chineseCnt = 12;
                        // }
                        int currentRowMaxWidth = 8;
                        if (j < rowMaxWidth.Count)
                        {
                            currentRowMaxWidth = rowMaxWidth[j];
                        }

                        int padding = currentRowMaxWidth;
                        if (padding < 8)
                        {
                            padding = 8;
                        }

                        str = str.PadRight(padding - chineseCnt < 0 ? 0 : padding - chineseCnt, ' ');
                        result.Add(str);
                    }

                    lines.Add(string.Join("|", result));
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    continue;
                }
            }
        }

        private int GetChineseCnt(string str)
        {
            using CharEnumerator cEnumerator = str.GetEnumerator();
            Regex regex = new Regex("^[\u4E00-\u9FA5]{0,}$");
            int cnt = 0;
            while (cEnumerator.MoveNext())
            {
                if (regex.IsMatch(cEnumerator.Current.ToString(), 0))
                    cnt++;
            }

            return cnt;
        }
    }
}