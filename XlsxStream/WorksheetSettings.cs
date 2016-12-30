namespace XlsxStream
{
    public class WorksheetSettings
    {
        public int DefaultRowHeight { get; set; }
        public static WorksheetSettings Default
        {
            get
            {
                return new WorksheetSettings
                {
                    DefaultRowHeight = 15
                };
            }
        }
    }
}