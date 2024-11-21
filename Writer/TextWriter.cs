using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel2TextDiff
{
    public class TextWriter
    {
        public int PaddingBeginRow { get; init; }

        public char Separator { get; init; }


        private readonly StringBuilder _sb = new();

        private readonly List<int> _rowMaxWidth = new();

        private readonly List<List<string>> _rows = new();

        public void BeginSheet(string sheetName)
        {
            _sb.Append($"\n==========  [{sheetName}]   ==========");
            _rows.Clear();
            _rowMaxWidth.Clear();
        }

        public void WriteRow(List<string> row)
        {
            // 只保留到最后一个非空白单元格
            int lastNotEmptyIndex = row.FindLastIndex(s => !string.IsNullOrEmpty(s));
            if (lastNotEmptyIndex >= 0)
            {
                var listStr = row.GetRange(0, lastNotEmptyIndex + 1);
                // excelHeader = false;

                if (PaddingBeginRow == _rows.Count + 1 && _rowMaxWidth.Count <= 0)
                {
                    foreach (var str in listStr)
                    {
                        _rowMaxWidth.Add(str.Length);
                    }
                }

                var result = new List<string>();
                for (var i = 0; i < listStr.Count; i++)
                {
                    var str = listStr[i];

                    if (i < _rowMaxWidth.Count)
                    {
                        var totalCnt = str.Length;
                        var chineseCnt = GetChineseCnt(str);
                        var englishCnt = totalCnt - chineseCnt;
                        var cnt = englishCnt + 2 * chineseCnt;
                        _rowMaxWidth[i] = Math.Max(cnt, _rowMaxWidth[i]);
                    }

                    result.Add(str);
                }

                _rows.Add(result);
            }
            // else
            // {
            //     // 忽略空白行，没必要diff这个
            // }
        }

        public void EndSheet()
        {
            foreach (var t in _rows)
            {
                try
                {
                    _sb.Append('\n');

                    for (var j = 0; j < t.Count; j++)
                    {
                        var str = t[j];
                        if (j < _rowMaxWidth.Count)
                        {
                            var maxWidth = _rowMaxWidth[j];

                            var paddingCnt = Math.Max(maxWidth - str.Length, 0);

                            str = str.PadRight(paddingCnt, ' ');
                        }

                        if (j > 0) _sb.Append(Separator);
                        _sb.Append(str);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }

            _sb.Append('\n');
        }

        public void Save(string outputTextFile)
        {
            System.IO.File.WriteAllText(outputTextFile, _sb.ToString(), Encoding.UTF8);
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