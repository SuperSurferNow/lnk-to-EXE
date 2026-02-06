namespace lnk_to_EXE
{
    public class ShortcutInfo
    {
        public required string TargetPath { get; init; }
        public string Arguments { get; init; } = string.Empty;
        public string WorkingDirectory { get; init; } = string.Empty;
        public string IconPath { get; init; } = string.Empty;
        public int IconIndex { get; init; }
    }
}
