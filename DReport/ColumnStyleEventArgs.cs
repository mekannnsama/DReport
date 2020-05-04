namespace DReport
{
    internal class ColumnStyleEventArgs
    {
        public object Column { get; internal set; }
        public object Appearance { get; internal set; }
        public int RowHandle { get; internal set; }
    }
}