namespace BioDraw
{
    internal sealed class ImageReplacePreset
    {
        public string Name { get; set; }
        public int SortOrder { get; set; }
        public double FuzzPercent { get; set; }
        public string TargetColor { get; set; }
        public string Mode { get; set; }
        public string ReplacementColor { get; set; }
    }

    internal sealed class MaterialEntry
    {
        public string Name { get; set; }
        public string FilePath { get; set; }
    }
}
