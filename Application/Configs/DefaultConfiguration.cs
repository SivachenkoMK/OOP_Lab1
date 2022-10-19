namespace Application.Configs
{
    public class DefaultConfiguration
    {
        public int DefaultColumnsAmount { get; set; }
        public int DefaultRowsAmount { get; set; }
        public string CalculateMessage { get; set; } = default!;
        public string Rows { get; set; } = default!;
        public string Columns { get; set; } = default!;
        public string Save { get; set; } = default!;
        public string Open { get; set; } = default!;
        public string ApplicationName { get; set; } = default!;
    }
}