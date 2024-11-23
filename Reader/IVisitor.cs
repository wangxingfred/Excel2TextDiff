using System.Collections.Generic;

namespace Excel2TextDiff
{
    public interface IVisitor
    {
        public void BeginSheet(string sheetName);

        public void VisitRow(List<string> row);

        public void EndSheet();
    }
}