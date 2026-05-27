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

    internal sealed class AiImageApiSettings
    {
        public string DisplayName { get; set; }
        public string EndpointUrl { get; set; }
        public string ApiToken { get; set; }
        public string Model { get; set; }
        public int DefaultWidth { get; set; }
        public int DefaultHeight { get; set; }
        public string DefaultQuality { get; set; }
        public string DefaultFormat { get; set; }
        public string IconPath { get; set; }
        public bool LockAspectRatio { get; set; }
    }

    internal sealed class AiImageGlobalSettings
    {
        public bool OverridePerModel { get; set; }
        public int ModelPreviewCount { get; set; }
        public int DefaultWidth { get; set; }
        public int DefaultHeight { get; set; }
        public string DefaultQuality { get; set; }
        public string DefaultFormat { get; set; }

        // Independent settings for 图生图
        public bool ImgOverridePerModel { get; set; }
        public int ImgModelPreviewCount { get; set; }
        public int ImgDefaultWidth { get; set; }
        public int ImgDefaultHeight { get; set; }
        public string ImgDefaultQuality { get; set; }
        public string ImgDefaultFormat { get; set; }
    }
}
