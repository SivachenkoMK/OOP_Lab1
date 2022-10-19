namespace Application.Configs
{
    public class ErrorMessages
    {
        public string DivisionByZero { get; set; } = default!;
        public string Loop { get; set; } = default!;
        public string MessageAndName { get; set; } = default!;
        public string Name { get; set; } = default!;
        public string NonEmptyCellsPresent { get; set; } = default!;
        public string ReferencesToCurrentRowPresent { get; set; } = default!;
        public string ReferencesToCurrentColumnPresent { get; set; } = default!;
        public string ConfirmDeletingThisRow { get; set; } = default!;
        public string ConfirmDeletingThisColumn { get; set; } = default!;
        public string Warning { get; set; } = default!;
    }
}