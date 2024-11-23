namespace Excel2TextDiff
{
    internal interface IReader
    {
        public void Read(IVisitor visitor);
    }
}