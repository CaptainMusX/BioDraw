using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using static BioDraw.NativeMethods;

// TODO:   按照以下步骤启用功能区(XML)项:

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意: 如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。


namespace BioDraw
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private readonly List<string> level1Items;
        private readonly Dictionary<string, List<string>> level2Items;
        private readonly Dictionary<string, List<string>> level3Items;
        private readonly Dictionary<string, stdole.IPictureDisp> materialPreviewCache;
        private int selectedLevel1Index;
        private int selectedLevel2Index;
        private int selectedLevel3Index;
        private int materialPageIndex = 0;
        private const int MaterialPreviewButtonCount = 20;
        private int materialPreviewCount = 5;
        private const int MaterialThumbnailWidth = 132;
        private const int MaterialThumbnailHeight = 100;
        private const float MaterialLabelWidthRatio = 1.0f;
        private const int MaterialLabelMaxLines = 2;
        private const string TransparentPlaceholderResourceName = "BioDraw.BioDrawIcon.blank-image-200x200.png";
        private const string PickerColorResourceName = "BioDraw.BioDrawIcon.picker-color.png";
        private const string SettingsGearResourceName = "BioDraw.BioDrawIcon.settings-gear.png";
        private const string BrandPngResourceName = "BioDraw.BioDrawIcon.BioDraw.png";
        private const string ImageRecolorResourceName = "BioDraw.BioDrawIcon.image-recolor.png";
        private const string PlusPngResourceName = "BioDraw.BioDrawIcon.plus.png";
        private const string PickerColorFileName = "picker-color.png";
        private const string LegacyPickerColorFileName = "Picker_color .png";
        private const string BrandPngFileName = "BioDraw.png";
        private const string SettingsGearPngFileName = "settings-gear.png";
        private const string ImageRecolorPngFileName = "image-recolor.png";
        private const string AddToLibraryIcoFileName = "add-to-library.ico";
        private const int MaterialPreviewCacheLimit = 300;
        private const string RibbonEmptyInputToken = "\u2060";
        private string level1InputText;
        private string level2InputText;
        private string materialSearchText;
        private double imageReplaceFuzzInput;
        private stdole.IPictureDisp brandImageLarge;
        private stdole.IPictureDisp brandImageSmall;
        private stdole.IPictureDisp transparentPlaceholderImage;
        private stdole.IPictureDisp pickerButtonImage;
        private stdole.IPictureDisp settingsButtonImage;
        private stdole.IPictureDisp imageRecolorButtonImage;
        private stdole.IPictureDisp aiImageButtonImage;
        private stdole.IPictureDisp plusButtonImage;
        private stdole.IPictureDisp addToLibraryContextMenuImage;
        private stdole.IPictureDisp pageUpButtonImage;
        private stdole.IPictureDisp pageDownButtonImage;
        private readonly List<ImageReplacePreset> imageReplacePresets;
        private readonly string presetStorePath;
        private const string projectAddressUrl = "https://github.com/CaptainMusX/BioDraw";
        private string currentPresetName;
        private string defaultPresetName;
        private string materialLibraryPath;
        private string imageMagickPath;
        private string materialSearchCacheRootPath;
        private List<MaterialEntry> materialSearchCacheEntries;
        private string imageReplaceSourceColorInput;
        private string imageReplaceNewColorInput;
        private readonly List<string> imageReplaceSourceColorOptions;
        private readonly List<string> imageReplaceNewColorOptions;
        private bool presetEditorSaveAsDefaultChecked;
        private Rectangle presetEditorBounds;
        private bool hasPresetEditorBounds;
        private bool embedContextAddToLibraryEnabled;
        private List<AiImageApiSettings> aiImageSettingsList;
        private AiImageGlobalSettings aiGlobalSettings;
        private int aiModelPageIndex;
        private int aiModelPreviewCount;
        private int selectedAiModelIndex;
        private string aiWidthInputText;
        private string aiHeightInputText;
        private string aiQualityInputText;
        private readonly Dictionary<string, stdole.IPictureDisp> aiModelIconCache;
        private Rectangle aiSettingsDialogBounds;
        private bool hasAiSettingsDialogBounds;
        private Rectangle aiGlobalSettingsDialogBounds;
        private bool hasAiGlobalSettingsDialogBounds;
        private Rectangle imgGlobalSettingsDialogBounds;
        private bool hasImgGlobalSettingsDialogBounds;
        private int imgModelPageIndex;
        private int imgModelPreviewCount;
        private int selectedImgModelIndex;
        private string imgWidthInputText;
        private string imgHeightInputText;
        private string imgQualityInputText;

        public Ribbon1()
        {
            imageReplacePresets = new List<ImageReplacePreset>();
            imageReplaceSourceColorOptions = new List<string>();
            imageReplaceNewColorOptions = new List<string>();
            materialPreviewCache = new Dictionary<string, stdole.IPictureDisp>(StringComparer.OrdinalIgnoreCase);
            presetStorePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BioDraw",
                "image_replace_presets.xml");
            presetEditorBounds = Rectangle.Empty;
            hasPresetEditorBounds = false;
            presetEditorSaveAsDefaultChecked = false;
            embedContextAddToLibraryEnabled = true;
            aiImageSettingsList = AiImageService.LoadModelSettings();
            aiGlobalSettings = AiImageService.LoadGlobalSettings();
            aiModelPageIndex = 0;
            aiModelPreviewCount = AiImageService.ClampModelPreviewCount(aiGlobalSettings.ModelPreviewCount);
            selectedAiModelIndex = 0;
            aiWidthInputText = string.Empty;
            aiHeightInputText = string.Empty;
            aiQualityInputText = string.Empty;
            imgModelPageIndex = 0;
            imgModelPreviewCount = AiImageService.ClampModelPreviewCount(aiGlobalSettings.ImgModelPreviewCount);
            selectedImgModelIndex = 0;
            imgWidthInputText = string.Empty;
            imgHeightInputText = string.Empty;
            imgQualityInputText = string.Empty;
            aiModelIconCache = new Dictionary<string, stdole.IPictureDisp>(StringComparer.OrdinalIgnoreCase);
            aiSettingsDialogBounds = Rectangle.Empty;
            hasAiSettingsDialogBounds = false;
            aiGlobalSettingsDialogBounds = Rectangle.Empty;
            hasAiGlobalSettingsDialogBounds = false;
            imgGlobalSettingsDialogBounds = Rectangle.Empty;
            hasImgGlobalSettingsDialogBounds = false;
            AiImageService.TryParseSettingsBounds(out aiSettingsDialogBounds, out aiGlobalSettingsDialogBounds, out imgGlobalSettingsDialogBounds);
            hasAiSettingsDialogBounds = aiSettingsDialogBounds != Rectangle.Empty;
            hasAiGlobalSettingsDialogBounds = aiGlobalSettingsDialogBounds != Rectangle.Empty;
            hasImgGlobalSettingsDialogBounds = imgGlobalSettingsDialogBounds != Rectangle.Empty;
            materialLibraryPath = string.Empty;
            imageMagickPath = string.Empty;
            level1InputText = string.Empty;
            level2InputText = string.Empty;
            imageReplaceSourceColorInput = string.Empty;
            imageReplaceNewColorInput = string.Empty;
            imageReplaceFuzzInput = 5.0;
            materialPreviewCount = 5;
            ResetImageReplaceColorOptions();
            materialSearchText = string.Empty;
            level1Items = new List<string>
            {
                "Cell Types",
                "Proteins",
                "Nucleic Acids",
                "Human Anatomy",
                "Lab and Objects",
                "Species",
                "Agriculture",
                "Membranes",
                "Cell Structures",
                "Epithelium",
                "Lipids and Carbs",
                "Chemistry"
            };

            level2Items = new Dictionary<string, List<string>>
            {
                {
                    "Cell Types",
                    new List<string>
                    {
                        "Epithelial Cells",
                        "Generic Cells",
                        "Immune Blood Cells",
                        "Mitosis and Meiosis",
                        "Muscle Cells",
                        "Neural Cells",
                        "Plant Cells",
                        "Reproductive Cells",
                        "Secretory Cells",
                        "Stromal Cells"
                    }
                },
                {
                    "Proteins",
                    new List<string>
                    {
                        "Antibodies",
                        "Enzymes",
                        "Generic Proteins",
                        "Intercellular Proteins",
                        "Pathway Proteins",
                        "Receptors and Ligands",
                        "Soluble Proteins",
                        "Transporters"
                    }
                },
                {
                    "Nucleic Acids",
                    new List<string>
                    {
                        "DNA",
                        "DNA (with Nucleotides)",
                        "DNA Ministring",
                        "Nucleic acid motifs",
                        "Nucleotide Bases",
                        "Plasmids",
                        "RNA"
                    }
                },
                {
                    "Human Anatomy",
                    new List<string>
                    {
                        "Cardiovascular System",
                        "Dental",
                        "Digestive System",
                        "Embryology",
                        "Endocrine and Exocrine System",
                        "Head and Neuroanatomy",
                        "Human Figure",
                        "Lymphatic System",
                        "Muscular System",
                        "Reproductive System",
                        "Respiratory System",
                        "Skeletal System",
                        "Skin",
                        "Urogenital System"
                    }
                },
                {
                    "Lab and Objects",
                    new List<string>
                    {
                        "Animal Housing",
                        "Beakers, Bottles, Flasks",
                        "Environment and Ecology",
                        "Food",
                        "Machinery and Tech",
                        "Medical Equipment",
                        "Microscope and Optics",
                        "Nanoparticles",
                        "Other Items",
                        "Tools",
                        "Tubes and Vials",
                        "Wells, Plates, and Cultures"
                    }
                },
                {
                    "Species",
                    new List<string>
                    {
                        "Amphibians",
                        "Arthropods",
                        "Bacteria",
                        "Birds",
                        "Fish",
                        "Fungi",
                        "Mammals",
                        "Other Organisms",
                        "Plants",
                        "Reptiles",
                        "Rodents",
                        "Viruses",
                        "Worms"
                    }
                },
                {
                    "Agriculture",
                    new List<string>
                    {
                        "Agricultural Plants",
                        "Produce",
                        "Plant Anatomy",
                        "Plant Pathology",
                        "Landscapes and Soil",
                        "Animal Agriculture",
                        "Equipment and Objects",
                        "Agricultural Symbols"
                    }
                },
                {
                    "Membranes",
                    new List<string>
                    {
                        "Bacterial Membranes",
                        "Neural Membranes",
                        "Nuclear Membranes",
                        "Phospholipid Bilayer Membranes",
                        "Simplified Bilayer Membranes"
                    }
                },
                {
                    "Cell Structures",
                    new List<string>
                    {
                        "Cytoskeleton and ECM",
                        "Organelles"
                    }
                },
                {
                    "Epithelium",
                    new List<string>
                    {
                        "Glandular Epithelia",
                        "Intestinal Epithelia",
                        "Skin Epithelia"
                    }
                },
                {
                    "Lipids and Carbs",
                    new List<string>
                    {
                        "Carbohydrates",
                        "Glycans",
                        "Lipids"
                    }
                },
                {
                    "Chemistry",
                    new List<string>
                    {
                        "Amino Acids",
                        "Biochemistry",
                        "Molecular Model Kit",
                        "Nanoparticles",
                        "Structural Formulas"
                    }
                }
            };

            level3Items = new Dictionary<string, List<string>>
            {
            };

            selectedLevel1Index = 0;
            selectedLevel2Index = 0;
            selectedLevel3Index = 0;

            LoadImageReplacePresets();
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            var xml = GetResourceText("BioDraw.Ribbon1.xml");
            if (!string.IsNullOrWhiteSpace(xml))
            {
                return xml;
            }

            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            foreach (var resourceName in asm.GetManifestResourceNames())
            {
                if (resourceName.EndsWith("Ribbon1.xml", System.StringComparison.OrdinalIgnoreCase))
                {
                    var fallback = GetResourceText(resourceName);
                    if (!string.IsNullOrWhiteSpace(fallback))
                    {
                        return fallback;
                    }
                }
            }

            return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab id='TabBioDrawFallback' label='BioDraw'><group id='GroupFallback' label='BioDraw'><button id='BtnFallback' label='BioDraw' onAction='OnAbout' imageMso='HappyFace' size='large'/></group></tab></tabs></ribbon></customUI>";
        }

        #endregion

        #region 功能区回调
        //在此处创建回叫方法。有关添加回叫方法的详细信息，请访问 https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            EnsureBrandImages();
        }

        public stdole.IPictureDisp GetBrandImage(Office.IRibbonControl control)
        {
            EnsureBrandImages();
            return brandImageLarge ?? brandImageSmall;
        }

        public stdole.IPictureDisp GetPickerButtonImage(Office.IRibbonControl control)
        {
            EnsureBrandImages();
            return pickerButtonImage ?? brandImageSmall ?? brandImageLarge;
        }

        public stdole.IPictureDisp GetSettingsButtonImage(Office.IRibbonControl control)
        {
            EnsureBrandImages();
            return settingsButtonImage ?? brandImageSmall ?? brandImageLarge;
        }

        public stdole.IPictureDisp GetImageRecolorButtonImage(Office.IRibbonControl control)
        {
            EnsureBrandImages();
            return imageRecolorButtonImage ?? brandImageLarge ?? brandImageSmall;
        }

        public stdole.IPictureDisp GetAiModelImage1(Office.IRibbonControl control) { return GetAiModelImageForButton(0); }
        public stdole.IPictureDisp GetAiModelImage2(Office.IRibbonControl control) { return GetAiModelImageForButton(1); }
        public stdole.IPictureDisp GetAiModelImage3(Office.IRibbonControl control) { return GetAiModelImageForButton(2); }
        public stdole.IPictureDisp GetAiModelImage4(Office.IRibbonControl control) { return GetAiModelImageForButton(3); }
        public stdole.IPictureDisp GetAiModelImage5(Office.IRibbonControl control) { return GetAiModelImageForButton(4); }
        public stdole.IPictureDisp GetAiModelImage6(Office.IRibbonControl control) { return GetAiModelImageForButton(5); }
        public stdole.IPictureDisp GetAiModelImage7(Office.IRibbonControl control) { return GetAiModelImageForButton(6); }
        public stdole.IPictureDisp GetAiModelImage8(Office.IRibbonControl control) { return GetAiModelImageForButton(7); }
        public stdole.IPictureDisp GetAiModelImage9(Office.IRibbonControl control) { return GetAiModelImageForButton(8); }
        public stdole.IPictureDisp GetAiModelImage10(Office.IRibbonControl control) { return GetAiModelImageForButton(9); }
        public stdole.IPictureDisp GetAiModelImage11(Office.IRibbonControl control) { return GetAiModelImageForButton(10); }
        public stdole.IPictureDisp GetAiModelImage12(Office.IRibbonControl control) { return GetAiModelImageForButton(11); }

        private stdole.IPictureDisp GetAiModelImageForButton(int buttonOffset)
        {
            EnsureBrandImages();
            var entry = GetAiModelEntryForButton(buttonOffset);
            if (entry == null)
                return plusButtonImage ?? aiImageButtonImage ?? brandImageLarge ?? brandImageSmall;
            return GetAiModelIcon(entry);
        }

        public stdole.IPictureDisp GetAddToLibraryContextMenuImage(Office.IRibbonControl control)
        {
            EnsureBrandImages();
            return addToLibraryContextMenuImage ?? settingsButtonImage ?? brandImageSmall ?? brandImageLarge;
        }

        public stdole.IPictureDisp GetPageButtonImage(Office.IRibbonControl control)
        {
            if (control != null && string.Equals(control.Id, "BtnPageUp", StringComparison.Ordinal))
            {
                if (pageUpButtonImage == null)
                {
                    pageUpButtonImage = CreateSvgChevronButtonImage(true);
                }
                return pageUpButtonImage ?? brandImageSmall ?? brandImageLarge;
            }

            if (pageDownButtonImage == null)
            {
                pageDownButtonImage = CreateSvgChevronButtonImage(false);
            }
            return pageDownButtonImage ?? brandImageSmall ?? brandImageLarge;
        }

        private static stdole.IPictureDisp CreateSvgChevronButtonImage(bool isDown)
        {
            const int size = 32;
            const float scale = size / 24f;
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                using (var pen = new Pen(Color.FromArgb(64, 64, 64), 2f))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    pen.LineJoin = LineJoin.Round;
                    PointF p1;
                    PointF p2;
                    PointF p3;
                    if (isDown)
                    {
                        p1 = new PointF(4.5f * scale, 15.75f * scale);
                        p2 = new PointF(12f * scale, 8.25f * scale);
                        p3 = new PointF(19.5f * scale, 15.75f * scale);
                    }
                    else
                    {
                        p1 = new PointF(4.5f * scale, 8.25f * scale);
                        p2 = new PointF(12f * scale, 15.75f * scale);
                        p3 = new PointF(19.5f * scale, 8.25f * scale);
                    }
                    g.DrawLines(pen, new[] { p1, p2, p3 });
                }
            }
            return PictureConverter.ToPictureDisp(bmp);
        }

        public stdole.IPictureDisp GetAiPageButtonImage(Office.IRibbonControl control)
        {
            if (control != null && string.Equals(control.Id, "BtnAiModelPageUp", StringComparison.Ordinal))
            {
                if (pageUpButtonImage == null)
                    pageUpButtonImage = CreateSvgChevronButtonImage(true);
                return pageUpButtonImage ?? brandImageSmall ?? brandImageLarge;
            }
            if (pageDownButtonImage == null)
                pageDownButtonImage = CreateSvgChevronButtonImage(false);
            return pageDownButtonImage ?? brandImageSmall ?? brandImageLarge;
        }

        public int GetLevel1Count(Office.IRibbonControl control)
        {
            return GetLevel1List().Count;
        }

        public string GetLevel1Label(Office.IRibbonControl control, int index)
        {
            var list = GetLevel1List();
            return list[index];
        }

        public int GetLevel1SelectedIndex(Office.IRibbonControl control)
        {
            var list = GetLevel1List();
            return MaterialLibraryService.NormalizeIndex(selectedLevel1Index, list.Count);
        }

        public string GetLevel1Text(Office.IRibbonControl control)
        {
            if (!string.IsNullOrWhiteSpace(level1InputText))
            {
                return level1InputText;
            }

            var list = GetLevel1List();
            if (list.Count == 0)
            {
                return string.Empty;
            }
            var index = MaterialLibraryService.NormalizeIndex(selectedLevel1Index, list.Count);
            return list[index];
        }

        public void OnLevel1Changed(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var list = GetLevel1List();
            selectedLevel1Index = MaterialLibraryService.NormalizeIndex(selectedIndex, list.Count);
            selectedLevel2Index = 0;
            if (list.Count > 0)
            {
                level1InputText = list[selectedLevel1Index];
            }

            var level2List = GetLevel2List();
            if (level2List.Count > 0)
            {
                level2InputText = level2List[MaterialLibraryService.NormalizeIndex(selectedLevel2Index, level2List.Count)];
            }
            else
            {
                level2InputText = string.Empty;
            }

            materialPageIndex = 0;
            ribbon?.InvalidateControl("DdLevel2");
            InvalidateMaterialPreview();
        }

        public void OnLevel1TextChanged(Office.IRibbonControl control, string text)
        {
            level1InputText = (text ?? string.Empty).Trim();
            var list = GetLevel1List();
            var index = PresetManager.FindExactMatchIndex(list, text);
            if (index < 0)
            {
                return;
            }
            OnLevel1Changed(control, string.Empty, index);
        }

        public int GetLevel2Count(Office.IRibbonControl control)
        {
            var list = GetLevel2List();
            return list.Count;
        }

        public string GetLevel2Label(Office.IRibbonControl control, int index)
        {
            var list = GetLevel2List();
            return list[index];
        }

        public int GetLevel2SelectedIndex(Office.IRibbonControl control)
        {
            var list = GetLevel2List();
            return MaterialLibraryService.NormalizeIndex(selectedLevel2Index, list.Count);
        }

        public string GetLevel2Text(Office.IRibbonControl control)
        {
            if (!string.IsNullOrWhiteSpace(level2InputText))
            {
                return level2InputText;
            }

            var list = GetLevel2List();
            if (list.Count == 0)
            {
                return string.Empty;
            }
            var index = MaterialLibraryService.NormalizeIndex(selectedLevel2Index, list.Count);
            return list[index];
        }

        public void OnLevel2Changed(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var list = GetLevel2List();
            selectedLevel2Index = MaterialLibraryService.NormalizeIndex(selectedIndex, list.Count);
            if (list.Count > 0)
            {
                level2InputText = list[selectedLevel2Index];
            }
            materialPageIndex = 0;
            InvalidateMaterialPreview();
        }

        public void OnLevel2TextChanged(Office.IRibbonControl control, string text)
        {
            level2InputText = (text ?? string.Empty).Trim();
            var list = GetLevel2List();
            var index = PresetManager.FindExactMatchIndex(list, text);
            if (index < 0)
            {
                return;
            }
            OnLevel2Changed(control, string.Empty, index);
        }

        public string GetMaterialSearchText(Office.IRibbonControl control)
        {
            return materialSearchText ?? string.Empty;
        }

        public void OnMaterialSearchChanged(Office.IRibbonControl control, string text)
        {
            materialSearchText = text ?? string.Empty;
            materialPageIndex = 0;
            InvalidateMaterialPreview();
        }

        public int GetLevel3Count(Office.IRibbonControl control)
        {
            var list = GetLevel3List();
            return list.Count;
        }

        public string GetLevel3Label(Office.IRibbonControl control, int index)
        {
            var list = GetLevel3List();
            return list[index];
        }

        public int GetLevel3SelectedIndex(Office.IRibbonControl control)
        {
            var list = GetLevel3List();
            return MaterialLibraryService.NormalizeIndex(selectedLevel3Index, list.Count);
        }

        public void OnLevel3Changed(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var list = GetLevel3List();
            selectedLevel3Index = MaterialLibraryService.NormalizeIndex(selectedIndex, list.Count);
            materialPageIndex = 0;
            InvalidateMaterialPreview();
        }

        private int GetMaterialPageSize()
        {
            return Math.Max(1, Math.Min(MaterialPreviewButtonCount, materialPreviewCount));
        }

        private void EnsureMaterialPageIndexRange(List<MaterialEntry> list)
        {
            var pageSize = GetMaterialPageSize();
            var totalPages = Math.Max(1, (int)Math.Ceiling((double)list.Count / pageSize));
            materialPageIndex = Math.Max(0, Math.Min(materialPageIndex, totalPages - 1));
        }

        private void InvalidateMaterialPreview()
        {
            if (ribbon == null) return;
            ribbon.InvalidateControl("BtnMaterial1");
            ribbon.InvalidateControl("BtnMaterial2");
            ribbon.InvalidateControl("BtnMaterial3");
            ribbon.InvalidateControl("BtnMaterial4");
            ribbon.InvalidateControl("BtnMaterial5");
            ribbon.InvalidateControl("BtnMaterial6");
            ribbon.InvalidateControl("BtnMaterial7");
            ribbon.InvalidateControl("BtnMaterial8");
            ribbon.InvalidateControl("BtnMaterial9");
            ribbon.InvalidateControl("BtnMaterial10");
            ribbon.InvalidateControl("BtnMaterial11");
            ribbon.InvalidateControl("BtnMaterial12");
            ribbon.InvalidateControl("BtnMaterial13");
            ribbon.InvalidateControl("BtnMaterial14");
            ribbon.InvalidateControl("BtnMaterial15");
            ribbon.InvalidateControl("BtnMaterial16");
            ribbon.InvalidateControl("BtnMaterial17");
            ribbon.InvalidateControl("BtnMaterial18");
            ribbon.InvalidateControl("BtnMaterial19");
            ribbon.InvalidateControl("BtnMaterial20");
            ribbon.InvalidateControl("BtnPageUp");
            ribbon.InvalidateControl("BtnPageDown");
        }

        private MaterialEntry GetMaterialEntryForButton(int buttonOffset)
        {
            var list = GetMaterialEntries();
            var pageSize = GetMaterialPageSize();
            EnsureMaterialPageIndexRange(list);
            if (buttonOffset < 0 || buttonOffset >= pageSize)
            {
                return null;
            }
            int index = materialPageIndex * pageSize + buttonOffset;
            if (index >= 0 && index < list.Count)
            {
                return list[index];
            }
            return null;
        }

        private bool IsMaterialButtonVisible(int buttonOffset)
        {
            return buttonOffset >= 0 && buttonOffset < GetMaterialPageSize();
        }

        public bool GetMaterialVisible1(Office.IRibbonControl control) { return IsMaterialButtonVisible(0); }
        public bool GetMaterialVisible2(Office.IRibbonControl control) { return IsMaterialButtonVisible(1); }
        public bool GetMaterialVisible3(Office.IRibbonControl control) { return IsMaterialButtonVisible(2); }
        public bool GetMaterialVisible4(Office.IRibbonControl control) { return IsMaterialButtonVisible(3); }
        public bool GetMaterialVisible5(Office.IRibbonControl control) { return IsMaterialButtonVisible(4); }
        public bool GetMaterialVisible6(Office.IRibbonControl control) { return IsMaterialButtonVisible(5); }
        public bool GetMaterialVisible7(Office.IRibbonControl control) { return IsMaterialButtonVisible(6); }
        public bool GetMaterialVisible8(Office.IRibbonControl control) { return IsMaterialButtonVisible(7); }
        public bool GetMaterialVisible9(Office.IRibbonControl control) { return IsMaterialButtonVisible(8); }
        public bool GetMaterialVisible10(Office.IRibbonControl control) { return IsMaterialButtonVisible(9); }
        public bool GetMaterialVisible11(Office.IRibbonControl control) { return IsMaterialButtonVisible(10); }
        public bool GetMaterialVisible12(Office.IRibbonControl control) { return IsMaterialButtonVisible(11); }
        public bool GetMaterialVisible13(Office.IRibbonControl control) { return IsMaterialButtonVisible(12); }
        public bool GetMaterialVisible14(Office.IRibbonControl control) { return IsMaterialButtonVisible(13); }
        public bool GetMaterialVisible15(Office.IRibbonControl control) { return IsMaterialButtonVisible(14); }
        public bool GetMaterialVisible16(Office.IRibbonControl control) { return IsMaterialButtonVisible(15); }
        public bool GetMaterialVisible17(Office.IRibbonControl control) { return IsMaterialButtonVisible(16); }
        public bool GetMaterialVisible18(Office.IRibbonControl control) { return IsMaterialButtonVisible(17); }
        public bool GetMaterialVisible19(Office.IRibbonControl control) { return IsMaterialButtonVisible(18); }
        public bool GetMaterialVisible20(Office.IRibbonControl control) { return IsMaterialButtonVisible(19); }

        public bool GetMaterialEnabled1(Office.IRibbonControl control) { return GetMaterialEntryForButton(0) != null; }
        public bool GetMaterialEnabled2(Office.IRibbonControl control) { return GetMaterialEntryForButton(1) != null; }
        public bool GetMaterialEnabled3(Office.IRibbonControl control) { return GetMaterialEntryForButton(2) != null; }
        public bool GetMaterialEnabled4(Office.IRibbonControl control) { return GetMaterialEntryForButton(3) != null; }
        public bool GetMaterialEnabled5(Office.IRibbonControl control) { return GetMaterialEntryForButton(4) != null; }
        public bool GetMaterialEnabled6(Office.IRibbonControl control) { return GetMaterialEntryForButton(5) != null; }
        public bool GetMaterialEnabled7(Office.IRibbonControl control) { return GetMaterialEntryForButton(6) != null; }
        public bool GetMaterialEnabled8(Office.IRibbonControl control) { return GetMaterialEntryForButton(7) != null; }
        public bool GetMaterialEnabled9(Office.IRibbonControl control) { return GetMaterialEntryForButton(8) != null; }
        public bool GetMaterialEnabled10(Office.IRibbonControl control) { return GetMaterialEntryForButton(9) != null; }
        public bool GetMaterialEnabled11(Office.IRibbonControl control) { return GetMaterialEntryForButton(10) != null; }
        public bool GetMaterialEnabled12(Office.IRibbonControl control) { return GetMaterialEntryForButton(11) != null; }
        public bool GetMaterialEnabled13(Office.IRibbonControl control) { return GetMaterialEntryForButton(12) != null; }
        public bool GetMaterialEnabled14(Office.IRibbonControl control) { return GetMaterialEntryForButton(13) != null; }
        public bool GetMaterialEnabled15(Office.IRibbonControl control) { return GetMaterialEntryForButton(14) != null; }
        public bool GetMaterialEnabled16(Office.IRibbonControl control) { return GetMaterialEntryForButton(15) != null; }
        public bool GetMaterialEnabled17(Office.IRibbonControl control) { return GetMaterialEntryForButton(16) != null; }
        public bool GetMaterialEnabled18(Office.IRibbonControl control) { return GetMaterialEntryForButton(17) != null; }
        public bool GetMaterialEnabled19(Office.IRibbonControl control) { return GetMaterialEntryForButton(18) != null; }
        public bool GetMaterialEnabled20(Office.IRibbonControl control) { return GetMaterialEntryForButton(19) != null; }

        public string GetMaterialLabel1(Office.IRibbonControl control) { return GetMaterialDisplayLabel(0); }
        public string GetMaterialLabel2(Office.IRibbonControl control) { return GetMaterialDisplayLabel(1); }
        public string GetMaterialLabel3(Office.IRibbonControl control) { return GetMaterialDisplayLabel(2); }
        public string GetMaterialLabel4(Office.IRibbonControl control) { return GetMaterialDisplayLabel(3); }
        public string GetMaterialLabel5(Office.IRibbonControl control) { return GetMaterialDisplayLabel(4); }
        public string GetMaterialLabel6(Office.IRibbonControl control) { return GetMaterialDisplayLabel(5); }
        public string GetMaterialLabel7(Office.IRibbonControl control) { return GetMaterialDisplayLabel(6); }
        public string GetMaterialLabel8(Office.IRibbonControl control) { return GetMaterialDisplayLabel(7); }
        public string GetMaterialLabel9(Office.IRibbonControl control) { return GetMaterialDisplayLabel(8); }
        public string GetMaterialLabel10(Office.IRibbonControl control) { return GetMaterialDisplayLabel(9); }
        public string GetMaterialLabel11(Office.IRibbonControl control) { return GetMaterialDisplayLabel(10); }
        public string GetMaterialLabel12(Office.IRibbonControl control) { return GetMaterialDisplayLabel(11); }
        public string GetMaterialLabel13(Office.IRibbonControl control) { return GetMaterialDisplayLabel(12); }
        public string GetMaterialLabel14(Office.IRibbonControl control) { return GetMaterialDisplayLabel(13); }
        public string GetMaterialLabel15(Office.IRibbonControl control) { return GetMaterialDisplayLabel(14); }
        public string GetMaterialLabel16(Office.IRibbonControl control) { return GetMaterialDisplayLabel(15); }
        public string GetMaterialLabel17(Office.IRibbonControl control) { return GetMaterialDisplayLabel(16); }
        public string GetMaterialLabel18(Office.IRibbonControl control) { return GetMaterialDisplayLabel(17); }
        public string GetMaterialLabel19(Office.IRibbonControl control) { return GetMaterialDisplayLabel(18); }
        public string GetMaterialLabel20(Office.IRibbonControl control) { return GetMaterialDisplayLabel(19); }

        public string GetMaterialScreentip1(Office.IRibbonControl control) { return GetMaterialTooltip(0); }
        public string GetMaterialScreentip2(Office.IRibbonControl control) { return GetMaterialTooltip(1); }
        public string GetMaterialScreentip3(Office.IRibbonControl control) { return GetMaterialTooltip(2); }
        public string GetMaterialScreentip4(Office.IRibbonControl control) { return GetMaterialTooltip(3); }
        public string GetMaterialScreentip5(Office.IRibbonControl control) { return GetMaterialTooltip(4); }
        public string GetMaterialScreentip6(Office.IRibbonControl control) { return GetMaterialTooltip(5); }
        public string GetMaterialScreentip7(Office.IRibbonControl control) { return GetMaterialTooltip(6); }
        public string GetMaterialScreentip8(Office.IRibbonControl control) { return GetMaterialTooltip(7); }
        public string GetMaterialScreentip9(Office.IRibbonControl control) { return GetMaterialTooltip(8); }
        public string GetMaterialScreentip10(Office.IRibbonControl control) { return GetMaterialTooltip(9); }
        public string GetMaterialScreentip11(Office.IRibbonControl control) { return GetMaterialTooltip(10); }
        public string GetMaterialScreentip12(Office.IRibbonControl control) { return GetMaterialTooltip(11); }
        public string GetMaterialScreentip13(Office.IRibbonControl control) { return GetMaterialTooltip(12); }
        public string GetMaterialScreentip14(Office.IRibbonControl control) { return GetMaterialTooltip(13); }
        public string GetMaterialScreentip15(Office.IRibbonControl control) { return GetMaterialTooltip(14); }
        public string GetMaterialScreentip16(Office.IRibbonControl control) { return GetMaterialTooltip(15); }
        public string GetMaterialScreentip17(Office.IRibbonControl control) { return GetMaterialTooltip(16); }
        public string GetMaterialScreentip18(Office.IRibbonControl control) { return GetMaterialTooltip(17); }
        public string GetMaterialScreentip19(Office.IRibbonControl control) { return GetMaterialTooltip(18); }
        public string GetMaterialScreentip20(Office.IRibbonControl control) { return GetMaterialTooltip(19); }

        public stdole.IPictureDisp GetMaterialImage1(Office.IRibbonControl control) { return GetMaterialImageForButton(0); }
        public stdole.IPictureDisp GetMaterialImage2(Office.IRibbonControl control) { return GetMaterialImageForButton(1); }
        public stdole.IPictureDisp GetMaterialImage3(Office.IRibbonControl control) { return GetMaterialImageForButton(2); }
        public stdole.IPictureDisp GetMaterialImage4(Office.IRibbonControl control) { return GetMaterialImageForButton(3); }
        public stdole.IPictureDisp GetMaterialImage5(Office.IRibbonControl control) { return GetMaterialImageForButton(4); }
        public stdole.IPictureDisp GetMaterialImage6(Office.IRibbonControl control) { return GetMaterialImageForButton(5); }
        public stdole.IPictureDisp GetMaterialImage7(Office.IRibbonControl control) { return GetMaterialImageForButton(6); }
        public stdole.IPictureDisp GetMaterialImage8(Office.IRibbonControl control) { return GetMaterialImageForButton(7); }
        public stdole.IPictureDisp GetMaterialImage9(Office.IRibbonControl control) { return GetMaterialImageForButton(8); }
        public stdole.IPictureDisp GetMaterialImage10(Office.IRibbonControl control) { return GetMaterialImageForButton(9); }
        public stdole.IPictureDisp GetMaterialImage11(Office.IRibbonControl control) { return GetMaterialImageForButton(10); }
        public stdole.IPictureDisp GetMaterialImage12(Office.IRibbonControl control) { return GetMaterialImageForButton(11); }
        public stdole.IPictureDisp GetMaterialImage13(Office.IRibbonControl control) { return GetMaterialImageForButton(12); }
        public stdole.IPictureDisp GetMaterialImage14(Office.IRibbonControl control) { return GetMaterialImageForButton(13); }
        public stdole.IPictureDisp GetMaterialImage15(Office.IRibbonControl control) { return GetMaterialImageForButton(14); }
        public stdole.IPictureDisp GetMaterialImage16(Office.IRibbonControl control) { return GetMaterialImageForButton(15); }
        public stdole.IPictureDisp GetMaterialImage17(Office.IRibbonControl control) { return GetMaterialImageForButton(16); }
        public stdole.IPictureDisp GetMaterialImage18(Office.IRibbonControl control) { return GetMaterialImageForButton(17); }
        public stdole.IPictureDisp GetMaterialImage19(Office.IRibbonControl control) { return GetMaterialImageForButton(18); }
        public stdole.IPictureDisp GetMaterialImage20(Office.IRibbonControl control) { return GetMaterialImageForButton(19); }

        private string GetMaterialDisplayLabel(int buttonOffset)
        {
            var item = GetMaterialEntryForButton(buttonOffset);
            return MaterialLibraryService.ToFixedLengthMaterialLabel(item == null ? string.Empty : item.Name);
        }

        private string GetMaterialTooltip(int buttonOffset)
        {
            var item = GetMaterialEntryForButton(buttonOffset);
            if (item == null || string.IsNullOrWhiteSpace(item.Name))
            {
                return "当前列无素材";
            }

            var title = item.Name.Trim();
            var fileName = string.Empty;
            if (!string.IsNullOrWhiteSpace(item.FilePath))
            {
                try
                {
                    fileName = Path.GetFileName(item.FilePath.Trim());
                }
                catch
                {
                    fileName = string.Empty;
                }
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                return title;
            }

            var titleNoExt = Path.GetFileNameWithoutExtension(fileName);
            if (string.Equals(title, titleNoExt, StringComparison.OrdinalIgnoreCase))
            {
                return fileName;
            }
            return title + " (" + fileName + ")";
        }




        private stdole.IPictureDisp GetMaterialImageForButton(int buttonOffset)
        {
            EnsureBrandImages();
            var item = GetMaterialEntryForButton(buttonOffset);
            if (item == null) return transparentPlaceholderImage ?? brandImageLarge ?? brandImageSmall;
            return GetMaterialPreviewImage(item);
        }

        public void OnMaterialClick1(Office.IRibbonControl control) { InsertMaterialAtOffset(0); }
        public void OnMaterialClick2(Office.IRibbonControl control) { InsertMaterialAtOffset(1); }
        public void OnMaterialClick3(Office.IRibbonControl control) { InsertMaterialAtOffset(2); }
        public void OnMaterialClick4(Office.IRibbonControl control) { InsertMaterialAtOffset(3); }
        public void OnMaterialClick5(Office.IRibbonControl control) { InsertMaterialAtOffset(4); }
        public void OnMaterialClick6(Office.IRibbonControl control) { InsertMaterialAtOffset(5); }
        public void OnMaterialClick7(Office.IRibbonControl control) { InsertMaterialAtOffset(6); }
        public void OnMaterialClick8(Office.IRibbonControl control) { InsertMaterialAtOffset(7); }
        public void OnMaterialClick9(Office.IRibbonControl control) { InsertMaterialAtOffset(8); }
        public void OnMaterialClick10(Office.IRibbonControl control) { InsertMaterialAtOffset(9); }
        public void OnMaterialClick11(Office.IRibbonControl control) { InsertMaterialAtOffset(10); }
        public void OnMaterialClick12(Office.IRibbonControl control) { InsertMaterialAtOffset(11); }
        public void OnMaterialClick13(Office.IRibbonControl control) { InsertMaterialAtOffset(12); }
        public void OnMaterialClick14(Office.IRibbonControl control) { InsertMaterialAtOffset(13); }
        public void OnMaterialClick15(Office.IRibbonControl control) { InsertMaterialAtOffset(14); }
        public void OnMaterialClick16(Office.IRibbonControl control) { InsertMaterialAtOffset(15); }
        public void OnMaterialClick17(Office.IRibbonControl control) { InsertMaterialAtOffset(16); }
        public void OnMaterialClick18(Office.IRibbonControl control) { InsertMaterialAtOffset(17); }
        public void OnMaterialClick19(Office.IRibbonControl control) { InsertMaterialAtOffset(18); }
        public void OnMaterialClick20(Office.IRibbonControl control) { InsertMaterialAtOffset(19); }

        private void InsertMaterialAtOffset(int buttonOffset)
        {
            var item = GetMaterialEntryForButton(buttonOffset);
            if (item == null) return;

            if ((Control.ModifierKeys & Keys.Alt) == Keys.Alt)
            {
                RenameMaterial(item);
                return;
            }

            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                DeleteMaterial(item);
                return;
            }

            InsertMaterial(item);
        }

        private void DeleteMaterial(MaterialEntry item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.FilePath))
            {
                ImageReplacePipeline.SetStatusText("BioDraw：当前素材不可删除。");
                return;
            }

            if (!File.Exists(item.FilePath))
            {
                materialPreviewCache.Remove(item.FilePath);
                materialSearchCacheRootPath = null;
                materialSearchCacheEntries = null;
                InvalidateMaterialPreview();
                ImageReplacePipeline.SetStatusText("BioDraw：素材文件不存在，已刷新列表。");
                return;
            }

            if (UseCustomMaterialLibrary() &&
                (!IsPathWithinRoot(materialLibraryPath, item.FilePath) ||
                 HasReparsePointBetween(materialLibraryPath, item.FilePath)))
            {
                MessageBox.Show("素材路径超出当前素材库或经过链接目录，已拒绝删除。", "BioDraw",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                "确认永久删除素材 \"" + (item.Name ?? Path.GetFileName(item.FilePath)) + "\" 吗？",
                "BioDraw",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes)
            {
                return;
            }

            try
            {
                materialPreviewCache.Remove(item.FilePath);
                File.Delete(item.FilePath);
                materialSearchCacheRootPath = null;
                materialSearchCacheEntries = null;
                InvalidateMaterialPreview();
                ImageReplacePipeline.SetStatusText("BioDraw：已删除素材 - " + item.Name);
            }
            catch (Exception ex)
            {
                MessageBox.Show("删除素材失败：" + ex.Message, "BioDraw");
            }
        }

        private void RenameMaterial(MaterialEntry item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.FilePath))
            {
                ImageReplacePipeline.SetStatusText("BioDraw：当前素材不可重命名。");
                return;
            }

            if (!File.Exists(item.FilePath))
            {
                materialPreviewCache.Remove(item.FilePath);
                materialSearchCacheRootPath = null;
                materialSearchCacheEntries = null;
                InvalidateMaterialPreview();
                ImageReplacePipeline.SetStatusText("BioDraw：素材文件不存在，已刷新列表。");
                return;
            }

            if (UseCustomMaterialLibrary() &&
                (!IsPathWithinRoot(materialLibraryPath, item.FilePath) ||
                 HasReparsePointBetween(materialLibraryPath, item.FilePath)))
            {
                MessageBox.Show("素材路径超出当前素材库或经过链接目录，已拒绝重命名。", "BioDraw",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var currentName = Path.GetFileNameWithoutExtension(item.FilePath) ?? string.Empty;
            string requestedName;
            if (!TryShowMaterialRenameDialog(currentName, out requestedName))
            {
                return;
            }

            var sanitizedName = SanitizeFileNameWithoutExtension(requestedName);
            if (string.IsNullOrWhiteSpace(sanitizedName))
            {
                MessageBox.Show("名称不能为空。", "BioDraw");
                return;
            }

            var directoryPath = Path.GetDirectoryName(item.FilePath);
            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                MessageBox.Show("素材路径无效，无法重命名。", "BioDraw");
                return;
            }

            var extension = Path.GetExtension(item.FilePath);
            if (string.IsNullOrWhiteSpace(extension))
            {
                extension = ".png";
            }

            var desiredPath = Path.Combine(directoryPath, sanitizedName + extension);
            var targetPath = desiredPath;
            if (File.Exists(targetPath) && !string.Equals(targetPath, item.FilePath, StringComparison.OrdinalIgnoreCase))
            {
                targetPath = BuildUniqueFilePath(directoryPath, sanitizedName, extension);
            }

            if (string.Equals(targetPath, item.FilePath, StringComparison.Ordinal))
            {
                ImageReplacePipeline.SetStatusText("BioDraw：素材名称未变化。");
                return;
            }

            try
            {
                // Windows 文件系统大小写不敏感，纯大小写改名需借助中间文件。
                if (string.Equals(targetPath, item.FilePath, StringComparison.OrdinalIgnoreCase))
                {
                    var tempPath = Path.Combine(directoryPath, "_biodraw_rename_" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture) + extension);
                    File.Move(item.FilePath, tempPath);
                    File.Move(tempPath, targetPath);
                }
                else
                {
                    File.Move(item.FilePath, targetPath);
                }

                materialPreviewCache.Remove(item.FilePath);
                materialPreviewCache.Remove(targetPath);
                materialSearchCacheRootPath = null;
                materialSearchCacheEntries = null;
                InvalidateMaterialPreview();
                ImageReplacePipeline.SetStatusText("BioDraw：已重命名素材 - " + Path.GetFileNameWithoutExtension(targetPath));
            }
            catch (Exception ex)
            {
                MessageBox.Show("重命名素材失败：" + ex.Message, "BioDraw");
            }
        }

        private static bool TryShowMaterialRenameDialog(string currentName, out string newName)
        {
            newName = currentName ?? string.Empty;
            using (var dialog = new Form())
            using (var layout = new TableLayoutPanel())
            using (var lblName = new Label())
            using (var txtName = new TextBox())
            using (var buttonRow = new FlowLayoutPanel())
            using (var btnOk = new Button())
            using (var btnCancel = new Button())
            {
                dialog.Text = "重命名素材";
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterScreen;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.ShowInTaskbar = false;
                dialog.AutoScaleMode = AutoScaleMode.Dpi;
                dialog.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Regular, GraphicsUnit.Point);
                dialog.ClientSize = new Size(640, 250);
                dialog.Padding = new Padding(24, 22, 24, 20);

                lblName.Text = "请输入新名称";
                lblName.AutoSize = true;
                lblName.Margin = new Padding(0);
                lblName.Anchor = AnchorStyles.Left | AnchorStyles.Top;

                txtName.Text = currentName ?? string.Empty;
                txtName.Font = dialog.Font;
                txtName.BorderStyle = BorderStyle.FixedSingle;
                txtName.Multiline = true;
                txtName.Margin = new Padding(0);
                txtName.Dock = DockStyle.Fill;
                txtName.MinimumSize = new Size(0, 40);

                var buttonHeight = 36;
                var buttonWidth = 104;
                var buttonGap = 12;
                btnOk.Text = "确定";
                btnOk.DialogResult = DialogResult.OK;
                btnOk.Size = new Size(buttonWidth, buttonHeight);
                btnOk.Margin = new Padding(0, 0, buttonGap, 0);

                btnCancel.Text = "取消";
                btnCancel.DialogResult = DialogResult.Cancel;
                btnCancel.Size = new Size(buttonWidth, buttonHeight);
                btnCancel.Margin = new Padding(0);

                buttonRow.FlowDirection = FlowDirection.LeftToRight;
                buttonRow.WrapContents = false;
                buttonRow.AutoSize = true;
                buttonRow.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                buttonRow.Margin = new Padding(0);
                buttonRow.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                buttonRow.Controls.Add(btnOk);
                buttonRow.Controls.Add(btnCancel);

                layout.Dock = DockStyle.Fill;
                layout.Margin = new Padding(0);
                layout.Padding = new Padding(0);
                layout.ColumnCount = 1;
                layout.RowCount = 5;
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 标签
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26F)); // 固定间距，防止挤占
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F)); // 输入框
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 弹性空白
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 按钮
                layout.Controls.Add(lblName, 0, 0);
                layout.Controls.Add(txtName, 0, 2);
                layout.Controls.Add(buttonRow, 0, 4);

                dialog.Controls.Add(layout);
                dialog.AcceptButton = btnOk;
                dialog.CancelButton = btnCancel;
                dialog.Shown += (_, __) =>
                {
                    txtName.Focus();
                    txtName.SelectAll();
                };

                var result = dialog.ShowDialog();
                if (result != DialogResult.OK)
                {
                    return false;
                }

                newName = (txtName.Text ?? string.Empty).Trim();
                return true;
            }
        }

        public void OnMaterialPageUp(Office.IRibbonControl control)
        {
            if (materialPageIndex > 0)
            {
                materialPageIndex--;
                InvalidateMaterialPreview();
            }
        }

        public void OnMaterialPageDown(Office.IRibbonControl control)
        {
            var list = GetMaterialEntries();
            var pageSize = GetMaterialPageSize();
            int totalPages = (int)Math.Ceiling((double)list.Count / pageSize);
            if (materialPageIndex < totalPages - 1)
            {
                materialPageIndex++;
                InvalidateMaterialPreview();
            }
        }

        public void OnEditMaterialPreviewSettings(Office.IRibbonControl control)
        {
            int newCount;
            if (!ShowMaterialPreviewSettingsDialog(out newCount))
            {
                return;
            }

            materialPreviewCount = Math.Max(1, Math.Min(MaterialPreviewButtonCount, newCount));
            materialPageIndex = 0;
            SaveImageReplacePresets();
            InvalidateMaterialPreview();
        }

        private bool ShowMaterialPreviewSettingsDialog(out int newCount)
        {
            newCount = GetMaterialPageSize();
            using (var form = new Form())
            using (var lblPreviewCount = new Label())
            using (var numPreviewCount = new NumericUpDown())
            using (var tbPreviewCount = new TrackBar())
            using (var btnOk = new Button())
            {
                form.Text = "预览设置";
                form.FormBorderStyle = FormBorderStyle.Sizable;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Regular, GraphicsUnit.Point);
                form.BackColor = Color.FromArgb(244, 247, 252);
                form.ForeColor = Color.FromArgb(32, 41, 57);
                form.AutoScaleMode = AutoScaleMode.Dpi;
                form.MinimizeBox = false;
                form.MaximizeBox = true;
                form.MinimumSize = new Size(520, 240);
                form.ClientSize = new Size(620, 280);

                lblPreviewCount.Text = "素材数量";
                lblPreviewCount.TextAlign = ContentAlignment.MiddleLeft;

                numPreviewCount.Minimum = 1;
                numPreviewCount.Maximum = MaterialPreviewButtonCount;
                numPreviewCount.DecimalPlaces = 0;
                numPreviewCount.Value = Convert.ToDecimal(GetMaterialPageSize(), CultureInfo.InvariantCulture);
                numPreviewCount.BorderStyle = BorderStyle.FixedSingle;
                numPreviewCount.TextAlign = HorizontalAlignment.Right;
                numPreviewCount.BackColor = Color.White;
                numPreviewCount.ForeColor = Color.FromArgb(32, 41, 57);

                tbPreviewCount.Minimum = 1;
                tbPreviewCount.Maximum = MaterialPreviewButtonCount;
                tbPreviewCount.TickFrequency = 1;
                tbPreviewCount.SmallChange = 1;
                tbPreviewCount.LargeChange = 1;
                tbPreviewCount.AutoSize = false;
                tbPreviewCount.Value = Convert.ToInt32(numPreviewCount.Value, CultureInfo.InvariantCulture);

                btnOk.Text = "保存";
                btnOk.DialogResult = DialogResult.OK;
                btnOk.FlatStyle = FlatStyle.Flat;
                btnOk.FlatAppearance.BorderSize = 1;
                btnOk.FlatAppearance.BorderColor = Color.FromArgb(24, 118, 242);
                btnOk.BackColor = Color.FromArgb(24, 118, 242);
                btnOk.ForeColor = Color.White;
                btnOk.UseVisualStyleBackColor = false;
                btnOk.Cursor = Cursors.Hand;

                var syncingPreviewCount = false;
                void SyncPreviewCountToTrackBar()
                {
                    if (syncingPreviewCount)
                    {
                        return;
                    }
                    syncingPreviewCount = true;
                    var next = Convert.ToInt32(numPreviewCount.Value, CultureInfo.InvariantCulture);
                    tbPreviewCount.Value = Math.Max(tbPreviewCount.Minimum, Math.Min(tbPreviewCount.Maximum, next));
                    syncingPreviewCount = false;
                }

                void SyncPreviewCountToNumeric()
                {
                    if (syncingPreviewCount)
                    {
                        return;
                    }
                    syncingPreviewCount = true;
                    numPreviewCount.Value = Convert.ToDecimal(tbPreviewCount.Value, CultureInfo.InvariantCulture);
                    syncingPreviewCount = false;
                }

                numPreviewCount.ValueChanged += (_, __) => SyncPreviewCountToTrackBar();
                tbPreviewCount.Scroll += (_, __) => SyncPreviewCountToNumeric();
                tbPreviewCount.MouseEnter += (_, __) => tbPreviewCount.Focus();
                tbPreviewCount.MouseWheel += (_, e) =>
                {
                    var delta = e.Delta > 0 ? 1 : -1;
                    var next = Math.Max(tbPreviewCount.Minimum, Math.Min(tbPreviewCount.Maximum, tbPreviewCount.Value + delta));
                    if (next == tbPreviewCount.Value)
                    {
                        return;
                    }
                    tbPreviewCount.Value = next;
                    SyncPreviewCountToNumeric();
                };

                void ApplyDialogLayout()
                {
                    var margin = 24;
                    var labelWidth = 160;
                    var fieldGap = 12;
                    var rowHeight = 40;
                    var buttonWidth = 128;
                    var buttonHeight = 40;

                    var fieldX = margin + labelWidth + fieldGap;
                    var rightEdge = form.ClientSize.Width - margin;
                    var top = margin + 20;

                    lblPreviewCount.Location = new Point(margin, top);
                    lblPreviewCount.Size = new Size(labelWidth, rowHeight);
                    numPreviewCount.Location = new Point(fieldX, top);
                    numPreviewCount.Size = new Size(180, rowHeight);
                    tbPreviewCount.Location = new Point(numPreviewCount.Right + 14, top + 4);
                    tbPreviewCount.Size = new Size(Math.Max(220, rightEdge - tbPreviewCount.Left), rowHeight - 8);

                    var bottomY = Math.Max(top + rowHeight + 24, form.ClientSize.Height - margin - buttonHeight);
                    btnOk.Location = new Point(rightEdge - buttonWidth, bottomY);
                    btnOk.Size = new Size(buttonWidth, buttonHeight);
                    ApplyRoundedRegion(btnOk, 7);
                }

                form.Controls.Add(lblPreviewCount);
                form.Controls.Add(numPreviewCount);
                form.Controls.Add(tbPreviewCount);
                form.Controls.Add(btnOk);
                form.AcceptButton = btnOk;
                form.Resize += (_, __) => ApplyDialogLayout();
                ApplyDialogLayout();

                if (form.ShowDialog() != DialogResult.OK)
                {
                    return false;
                }

                newCount = Convert.ToInt32(numPreviewCount.Value, CultureInfo.InvariantCulture);
                return true;
            }
        }

        private void InsertMaterial(MaterialEntry item)
        {
            if (string.IsNullOrWhiteSpace(item.FilePath))
            {
                ImageReplacePipeline.SetStatusText("BioDraw：当前素材仅为占位项。");
                return;
            }

            string error;
            if (!MaterialLibraryService.TryInsertMaterialToCurrentSlide(item.FilePath, out error))
            {
                if (!string.IsNullOrWhiteSpace(error))
                {
                    MessageBox.Show("插入素材失败：" + error, "BioDraw");
                }
                return;
            }

            ImageReplacePipeline.SetStatusText("BioDraw：已插入素材 - " + item.Name);
        }

        public void OnAbout(Office.IRibbonControl control)
        {
            using (var dialog = new AboutDialog())
            {
                dialog.ShowDialog();
            }
        }

        public void OnOpenProjectAddress(Office.IRibbonControl control)
        {
            if (string.IsNullOrWhiteSpace(projectAddressUrl))
            {
                MessageBox.Show("项目地址暂未配置。", "BioDraw");
                return;
            }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = projectAddressUrl,
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开项目地址失败：" + ex.Message, "BioDraw");
            }
        }

        public void OnSetMaterialLibraryPath(Office.IRibbonControl control)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "请选择素材库文件夹";
                if (!string.IsNullOrWhiteSpace(materialLibraryPath) && Directory.Exists(materialLibraryPath))
                {
                    dialog.SelectedPath = materialLibraryPath;
                }

                if (dialog.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    return;
                }

                materialLibraryPath = dialog.SelectedPath.Trim();
                selectedLevel1Index = 0;
                selectedLevel2Index = 0;
                level1InputText = string.Empty;
                level2InputText = string.Empty;
                materialPageIndex = 0;
                materialSearchText = string.Empty;
                materialPreviewCache.Clear();
                materialSearchCacheRootPath = null;
                materialSearchCacheEntries = null;
                SaveImageReplacePresets();
                ribbon?.InvalidateControl("DdLevel1");
                ribbon?.InvalidateControl("DdLevel2");
                ribbon?.InvalidateControl("TxtMaterialSearch");
                InvalidateMaterialPreview();
                ImageReplacePipeline.SetStatusText("BioDraw：素材库路径已更新。");
            }
        }

        public void OnManageLevel1Directory(Office.IRibbonControl control)
        {
            string message;
            if (!TryManageLevel1Directory(out message))
            {
                MessageBox.Show(message, "BioDraw");
                return;
            }

            ImageReplacePipeline.SetStatusText(message);
        }

        public void OnManageLevel2Directory(Office.IRibbonControl control)
        {
            string message;
            if (!TryManageLevel2Directory(out message))
            {
                MessageBox.Show(message, "BioDraw");
                return;
            }

            ImageReplacePipeline.SetStatusText(message);
        }

        public void OnExecuteMaterialSearch(Office.IRibbonControl control)
        {
            OnMaterialSearchChanged(control, materialSearchText ?? string.Empty);
        }

        private bool TryManageLevel1Directory(out string message)
        {
            message = string.Empty;
            if (!UseCustomMaterialLibrary())
            {
                message = "请先在“关于 -> 素材库”中设置素材库目录。";
                return false;
            }

            var name = (level1InputText ?? string.Empty).Trim();
            string fullPath;
            string pathError;
            if (!TryBuildManagedDirectoryPath(
                    materialLibraryPath, name, out fullPath, out pathError))
            {
                message = pathError;
                return false;
            }
            var isDelete = (Control.ModifierKeys & Keys.Control) == Keys.Control;

            try
            {
                if (isDelete)
                {
                    if (!Directory.Exists(fullPath))
                    {
                        message = "要删除的类别目录不存在。";
                        return false;
                    }

                    if (HasReparsePointBetween(materialLibraryPath, fullPath))
                    {
                        message = "该目录包含链接或挂载点，为避免越界删除，已拒绝操作。";
                        return false;
                    }

                    var confirm = MessageBox.Show(
                        "确认永久删除类别 \"" + name + "\" 及其中全部素材吗？",
                        "BioDraw",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    if (confirm != DialogResult.Yes)
                    {
                        message = "BioDraw：已取消删除。";
                        return true;
                    }

                    Directory.Delete(fullPath, true);
                    RefreshMaterialDirectoryView(name, string.Empty, false);
                    message = "BioDraw：已删除类别目录 - " + name;
                    return true;
                }

                if (Directory.Exists(fullPath))
                {
                    RefreshMaterialDirectoryView(name, string.Empty, false);
                    message = "BioDraw：类别目录已存在 - " + name;
                    return true;
                }

                Directory.CreateDirectory(fullPath);
                RefreshMaterialDirectoryView(name, string.Empty, false);
                message = "BioDraw：已创建类别目录 - " + name;
                return true;
            }
            catch (Exception ex)
            {
                message = "操作失败：" + ex.Message;
                return false;
            }
        }

        private bool TryManageLevel2Directory(out string message)
        {
            message = string.Empty;
            if (!UseCustomMaterialLibrary())
            {
                message = "请先在“关于 -> 素材库”中设置素材库目录。";
                return false;
            }

            var level1List = GetLevel1List();
            if (level1List.Count == 0)
            {
                message = "请先创建类别目录。";
                return false;
            }

            var level1Name = level1List[MaterialLibraryService.NormalizeIndex(selectedLevel1Index, level1List.Count)];
            var level1Path = Path.Combine(materialLibraryPath, level1Name);
            if (!Directory.Exists(level1Path))
            {
                message = "当前类别目录不存在，请先创建类别。";
                return false;
            }
            if (HasReparsePointBetween(materialLibraryPath, level1Path))
            {
                message = "当前类别是链接或挂载目录，已拒绝管理其子目录。";
                return false;
            }

            var level2Name = (level2InputText ?? string.Empty).Trim();
            string level2Path;
            string pathError;
            if (!TryBuildManagedDirectoryPath(
                    level1Path, level2Name, out level2Path, out pathError) ||
                !IsPathWithinRoot(materialLibraryPath, level2Path))
            {
                message = string.IsNullOrWhiteSpace(pathError)
                    ? "子类目录超出素材库范围，已拒绝操作。"
                    : pathError;
                return false;
            }
            var isDelete = (Control.ModifierKeys & Keys.Control) == Keys.Control;
            try
            {
                if (isDelete)
                {
                    if (!Directory.Exists(level2Path))
                    {
                        message = "要删除的子类目录不存在。";
                        return false;
                    }

                    if (HasReparsePointBetween(materialLibraryPath, level2Path))
                    {
                        message = "该目录包含链接或挂载点，为避免越界删除，已拒绝操作。";
                        return false;
                    }

                    var confirm = MessageBox.Show(
                        "确认永久删除子类 \"" + level2Name + "\" 及其中全部素材吗？",
                        "BioDraw",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    if (confirm != DialogResult.Yes)
                    {
                        message = "BioDraw：已取消删除。";
                        return true;
                    }

                    Directory.Delete(level2Path, true);
                    RefreshMaterialDirectoryView(level1Name, level2Name, false);
                    message = "BioDraw：已删除子类目录 - " + level2Name;
                    return true;
                }

                if (Directory.Exists(level2Path))
                {
                    RefreshMaterialDirectoryView(level1Name, level2Name, true);
                    message = "BioDraw：子类目录已存在 - " + level2Name;
                    return true;
                }

                Directory.CreateDirectory(level2Path);
                RefreshMaterialDirectoryView(level1Name, level2Name, true);
                message = "BioDraw：已创建子类目录 - " + level2Name;
                return true;
            }
            catch (Exception ex)
            {
                message = "操作失败：" + ex.Message;
                return false;
            }
        }

        private void RefreshMaterialDirectoryView(string preferredLevel1, string preferredLevel2, bool keepLevel2)
        {
            materialPreviewCache.Clear();
            materialSearchCacheRootPath = null;
            materialSearchCacheEntries = null;
            materialPageIndex = 0;

            var level1List = GetLevel1List();
            var level1Index = PresetManager.FindExactMatchIndex(level1List, preferredLevel1);
            selectedLevel1Index = level1Index >= 0 ? level1Index : 0;

            var level2List = GetLevel2List();
            var level2Index = keepLevel2 ? PresetManager.FindExactMatchIndex(level2List, preferredLevel2) : -1;
            selectedLevel2Index = level2Index >= 0 ? level2Index : 0;

            level1InputText = level1List.Count > 0 ? level1List[MaterialLibraryService.NormalizeIndex(selectedLevel1Index, level1List.Count)] : string.Empty;
            level2InputText = level2List.Count > 0 ? level2List[MaterialLibraryService.NormalizeIndex(selectedLevel2Index, level2List.Count)] : string.Empty;

            ribbon?.InvalidateControl("DdLevel1");
            ribbon?.InvalidateControl("DdLevel2");
            InvalidateMaterialPreview();
        }

        public bool GetEmbedContextAddToLibraryPressed(Office.IRibbonControl control)
        {
            return embedContextAddToLibraryEnabled;
        }

        public void OnToggleEmbedContextAddToLibrary(Office.IRibbonControl control, bool pressed)
        {
            embedContextAddToLibraryEnabled = pressed;
            SaveImageReplacePresets();
            ribbon?.Invalidate();
            ImageReplacePipeline.SetStatusText(pressed
                ? "BioDraw：已启用右键菜单项 - 添加到 BioDraw 素材库。"
                : "BioDraw：已停用右键菜单项 - 添加到 BioDraw 素材库。");
        }

        public bool GetAddToLibraryContextMenuVisible(Office.IRibbonControl control)
        {
            if (!embedContextAddToLibraryEnabled)
            {
                return false;
            }

            return HasSelectedShapes();
        }

        public void OnAddToLibraryFromContextMenu(Office.IRibbonControl control)
        {
            string message;
            if (!TryAddSelectionToCurrentMaterialFolder(out message))
            {
                MessageBox.Show(message, "BioDraw");
                return;
            }

            ImageReplacePipeline.SetStatusText(message);
        }

        public string GetApplyPresetLabel(Office.IRibbonControl control)
        {
            var preset = GetCurrentPreset();
            if (preset == null)
            {
                return "颜色替换";
            }
            return $"颜色替换({preset.Name})";
        }

        public string GetImageReplaceSourceColorText(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            return PresetManager.ToRibbonColorInputText(imageReplaceSourceColorInput);
        }

        public int GetImageReplaceSourceColorItemCount(Office.IRibbonControl control)
        {
            return imageReplaceSourceColorOptions.Count + 1;
        }

        public string GetImageReplaceSourceColorItemLabel(Office.IRibbonControl control, int index)
        {
            if (index == 0)
            {
                return RibbonEmptyInputToken;
            }
            int optIndex = index - 1;
            if (optIndex >= 0 && optIndex < imageReplaceSourceColorOptions.Count)
            {
                return imageReplaceSourceColorOptions[optIndex];
            }
            return string.Empty;
        }

        public int GetImageReplaceSourceColorSelectedIndex(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            var normalized = PresetManager.NormalizeColorInputText(imageReplaceSourceColorInput);
            if (string.IsNullOrEmpty(normalized))
            {
                return 0;
            }
            int idx = PresetManager.FindColorOptionIndex(imageReplaceSourceColorOptions, normalized);
            return idx >= 0 ? idx + 1 : -1;
        }

        public string GetImageReplaceNewColorText(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            return PresetManager.ToRibbonColorInputText(imageReplaceNewColorInput);
        }

        public int GetImageReplaceNewColorItemCount(Office.IRibbonControl control)
        {
            return imageReplaceNewColorOptions.Count + 1;
        }

        public string GetImageReplaceNewColorItemLabel(Office.IRibbonControl control, int index)
        {
            if (index == 0)
            {
                return RibbonEmptyInputToken;
            }
            int optIndex = index - 1;
            if (optIndex >= 0 && optIndex < imageReplaceNewColorOptions.Count)
            {
                return imageReplaceNewColorOptions[optIndex];
            }
            return string.Empty;
        }

        public int GetImageReplaceNewColorSelectedIndex(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            var normalized = PresetManager.NormalizeColorInputText(imageReplaceNewColorInput);
            if (string.IsNullOrEmpty(normalized))
            {
                return 0;
            }
            int idx = PresetManager.FindColorOptionIndex(imageReplaceNewColorOptions, normalized);
            return idx >= 0 ? idx + 1 : -1;
        }

        public void OnImageReplaceSourceColorChanged(Office.IRibbonControl control, string text)
        {
            imageReplaceSourceColorInput = PresetManager.ToStorageColorInputText(text);
            PersistImageReplaceInputMemory();
            ribbon?.InvalidateControl("TxtImageReplaceSourceColor");
        }

        public void OnImageReplaceNewColorChanged(Office.IRibbonControl control, string text)
        {
            imageReplaceNewColorInput = PresetManager.ToStorageColorInputText(text);
            PersistImageReplaceInputMemory();
            ribbon?.InvalidateControl("TxtImageReplaceNewColor");
        }

        public void OnEditImageReplaceSourceColor(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            string selectedColor;
            if (!ShowColorOptionManagerDialog("原色预设", imageReplaceSourceColorOptions, imageReplaceSourceColorInput, false, out selectedColor))
            {
                return;
            }

            imageReplaceSourceColorInput = PresetManager.ToStorageColorInputText(selectedColor);
            PersistImageReplaceInputMemory();
            ribbon?.InvalidateControl("TxtImageReplaceSourceColor");
        }

        public void OnEditImageReplaceNewColor(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            string selectedColor;
            if (!ShowColorOptionManagerDialog("新色预设", imageReplaceNewColorOptions, imageReplaceNewColorInput, true, out selectedColor))
            {
                return;
            }

            imageReplaceNewColorInput = PresetManager.ToStorageColorInputText(selectedColor);
            PersistImageReplaceInputMemory();
            ribbon?.InvalidateControl("TxtImageReplaceNewColor");
        }

        public void OnPickImageReplaceSourceColor(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            string colorToken;
            string errorMessage;
            if (!TryPickColorWithPowerPoint(false, imageReplaceSourceColorInput, out colorToken, out errorMessage))
            {
                if (!string.IsNullOrWhiteSpace(errorMessage))
                {
                    MessageBox.Show(errorMessage, "BioDraw");
                }
                return;
            }

            imageReplaceSourceColorInput = PresetManager.ToStorageColorInputText(colorToken);
            PersistImageReplaceInputMemory();
            ribbon?.InvalidateControl("TxtImageReplaceSourceColor");
        }

        public void OnPickImageReplaceNewColor(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            string colorToken;
            string errorMessage;
            if (!TryPickColorWithPowerPoint(false, imageReplaceNewColorInput, out colorToken, out errorMessage))
            {
                if (!string.IsNullOrWhiteSpace(errorMessage))
                {
                    MessageBox.Show(errorMessage, "BioDraw");
                }
                return;
            }

            imageReplaceNewColorInput = PresetManager.ToStorageColorInputText(colorToken);
            PersistImageReplaceInputMemory();
            ribbon?.InvalidateControl("TxtImageReplaceNewColor");
        }

        public int GetImageReplacePresetItemCount(Office.IRibbonControl control)
        {
            return imageReplacePresets.Count;
        }

        public string GetImageReplacePresetItemLabel(Office.IRibbonControl control, int index)
        {
            var ordered = PresetManager.GetPresetsInDisplayOrder(imageReplacePresets).ToList();
            if (index >= 0 && index < ordered.Count)
            {
                return ordered[index].Name;
            }
            return string.Empty;
        }

        public int GetImageReplacePresetSelectedIndex(Office.IRibbonControl control)
        {
            var ordered = PresetManager.GetPresetsInDisplayOrder(imageReplacePresets).ToList();
            if (ordered.Count == 0)
            {
                return -1;
            }
            var currentIndex = ordered.FindIndex(x => string.Equals(x.Name, currentPresetName, StringComparison.OrdinalIgnoreCase));
            return currentIndex >= 0 ? currentIndex : 0;
        }

        public string GetImageReplacePresetText(Office.IRibbonControl control)
        {
            var ordered = PresetManager.GetPresetsInDisplayOrder(imageReplacePresets).ToList();
            if (ordered.Count == 0)
            {
                return string.Empty;
            }
            var index = GetImageReplacePresetSelectedIndex(control);
            if (index < 0 || index >= ordered.Count)
            {
                return ordered[0].Name;
            }
            return ordered[index].Name;
        }

        public void OnImageReplacePresetDropDownChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var ordered = PresetManager.GetPresetsInDisplayOrder(imageReplacePresets).ToList();
            if (selectedIndex >= 0 && selectedIndex < ordered.Count)
            {
                currentPresetName = ordered[selectedIndex].Name;
                SyncImageReplaceInputValuesFromCurrentPreset();
                InvalidateImageReplaceRibbonControls();
            }
        }

        public void OnImageReplacePresetTextChanged(Office.IRibbonControl control, string text)
        {
            var ordered = PresetManager.GetPresetsInDisplayOrder(imageReplacePresets).ToList();
            var names = ordered.Select(x => x.Name).ToList();
            var index = PresetManager.FindExactMatchIndex(names, text);
            if (index < 0)
            {
                return;
            }
            OnImageReplacePresetDropDownChanged(control, string.Empty, index);
        }

        public void OnApplyImageReplace(Office.IRibbonControl control)
        {
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                OnEditImageReplacePreset(control);
                return;
            }

            var preset = GetCurrentPreset();

            var application = Globals.ThisAddIn?.Application;
            if (application == null)
            {
                MessageBox.Show("未能获取 PowerPoint 应用实例。", "BioDraw");
                return;
            }

            dynamic selection = null;
            try
            {
                selection = application.ActiveWindow?.Selection;
            }
            catch
            {
            }

            if (selection == null)
            {
                MessageBox.Show("请先选中一张或多张图片。", "BioDraw");
                return;
            }

            List<dynamic> rawShapes;
            if (!ImageReplacePipeline.TryGetSelectedShapes(selection, out rawShapes))
            {
                MessageBox.Show("请先选中一张或多张图片。", "BioDraw");
                return;
            }

            var shapes = ImageReplacePipeline.FlattenToPictures(rawShapes);
            if (shapes.Count == 0)
            {
                MessageBox.Show("选中组合中未找到图片。", "BioDraw");
                return;
            }

            EnsureImageReplaceInputValues();
            string sourceColor;
            if (!PresetManager.TryNormalizeImageMagickColor(
                    imageReplaceSourceColorInput, false, out sourceColor))
            {
                MessageBox.Show(
                    "原色格式无效。请使用颜色名称、十六进制颜色或 rgb/rgba 表达式。",
                    "BioDraw");
                return;
            }

            string newColor;
            if (!PresetManager.TryNormalizeImageMagickColor(
                    imageReplaceNewColorInput, true, out newColor))
            {
                MessageBox.Show(
                    "新色格式无效。请使用颜色名称、十六进制颜色或 rgb/rgba 表达式；留空表示透明。",
                    "BioDraw");
                return;
            }

            var applyPreset = new ImageReplacePreset
            {
                Name = preset?.Name ?? "临时预设",
                SortOrder = preset?.SortOrder ?? 1,
                FuzzPercent = PresetManager.NormalizeFuzzPercent(preset?.FuzzPercent ?? imageReplaceFuzzInput),
                TargetColor = sourceColor,
                Mode = string.IsNullOrWhiteSpace(newColor) ? "transparent" : "fill",
                ReplacementColor = string.IsNullOrWhiteSpace(newColor) ? "black" : newColor
            };

            var replacedCount = 0;
            var failedCount = 0;
            var lastError = string.Empty;
            var replacedShapes = new List<dynamic>();
            string snapshotPath;
            if (!ImageReplacePipeline.TryCreatePresentationSnapshot(application, out snapshotPath, out lastError))
            {
                MessageBox.Show(lastError, "BioDraw");
                return;
            }

            foreach (dynamic shape in shapes)
            {
                dynamic replacedShape;
                string error;
                if (!ImageReplacePipeline.TryReplaceShapePictureWithMagick(shape, imageMagickPath, applyPreset, snapshotPath, out replacedShape, out error))
                {
                    lastError = error;
                    failedCount++;
                    continue;
                }
                if (replacedShape != null)
                {
                    replacedShapes.Add(replacedShape);
                }
                replacedCount++;
            }

            ImageReplacePipeline.TryReselectShapes(replacedShapes);
            ImageReplacePipeline.TryDeleteFile(snapshotPath);

            if (replacedCount == 0)
            {
                MessageBox.Show(string.IsNullOrWhiteSpace(lastError) ? "处理失败：未找到可处理的图片。" : lastError, "BioDraw");
                return;
            }

            if (failedCount > 0)
            {
                ImageReplacePipeline.SetStatusText($"BioDraw：已替换 {replacedCount} 张，{failedCount} 张未处理。");
                return;
            }
            ImageReplacePipeline.SetStatusText($"BioDraw：已替换 {replacedCount} 张图片。");
        }

        public void OnEditImageReplacePreset(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            var preset = GetCurrentPreset();
            if (preset == null)
            {
                preset = PresetManager.CreateDefaultPreset();
                preset.Name = PresetManager.GenerateNewPresetName(imageReplacePresets);
                preset.SortOrder = Math.Max(1, imageReplacePresets.Count + 1);
                preset.TargetColor = PresetManager.NormalizeColorInputText(imageReplaceSourceColorInput);
                preset.Mode = PresetManager.HasVisibleColorText(imageReplaceNewColorInput) ? "fill" : "transparent";
                preset.ReplacementColor = PresetManager.HasVisibleColorText(imageReplaceNewColorInput) ? PresetManager.NormalizeColorInputText(imageReplaceNewColorInput) : "black";
                preset.FuzzPercent = PresetManager.NormalizeFuzzPercent(imageReplaceFuzzInput);
            }

            var canDelete = imageReplacePresets.Any(x => string.Equals(x.Name, preset.Name, StringComparison.OrdinalIgnoreCase));
            ImageReplacePreset editedPreset;
            bool setAsDefault;
            bool deleteRequested;
            if (!ShowPresetEditorDialog(preset, canDelete, out editedPreset, out setAsDefault, out deleteRequested))
            {
                return;
            }

            if (deleteRequested)
            {
                DeletePresetByName(preset.Name);
                return;
            }

            editedPreset.TargetColor = PresetManager.NormalizeColorInputText(imageReplaceSourceColorInput);
            editedPreset.Mode = PresetManager.HasVisibleColorText(imageReplaceNewColorInput) ? "fill" : "transparent";
            editedPreset.ReplacementColor = PresetManager.HasVisibleColorText(imageReplaceNewColorInput) ? PresetManager.NormalizeColorInputText(imageReplaceNewColorInput) : "black";
            editedPreset.FuzzPercent = PresetManager.NormalizeFuzzPercent(editedPreset.FuzzPercent);
            imageReplaceFuzzInput = editedPreset.FuzzPercent;
            var isSameName = string.Equals(editedPreset.Name, preset.Name, StringComparison.OrdinalIgnoreCase);
            var nameAlreadyExists = imageReplacePresets.Any(x => string.Equals(x.Name, editedPreset.Name, StringComparison.OrdinalIgnoreCase));
            var replaceOriginalPreset = isSameName || nameAlreadyExists;
            UpsertPresetBySortOrder(replaceOriginalPreset ? preset.Name : string.Empty, editedPreset, editedPreset.SortOrder);

            currentPresetName = editedPreset.Name;
            if (setAsDefault || (replaceOriginalPreset && string.Equals(defaultPresetName, preset.Name, StringComparison.OrdinalIgnoreCase)))
            {
                defaultPresetName = editedPreset.Name;
            }

            EnsurePresetSelectionNames();
            SaveImageReplacePresets();
            SyncImageReplaceInputValuesFromCurrentPreset();
            InvalidateImageReplaceRibbonControls();
        }

        public void OnDeleteImageReplacePreset(Office.IRibbonControl control)
        {
            var preset = GetCurrentPreset();
            if (preset == null)
            {
                return;
            }

            DeletePresetByName(preset.Name);
        }

        private void DeletePresetByName(string presetName)
        {
            if (string.IsNullOrWhiteSpace(presetName))
            {
                return;
            }

            if (!imageReplacePresets.Any(x => string.Equals(x.Name, presetName, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            imageReplacePresets.RemoveAll(x => string.Equals(x.Name, presetName, StringComparison.OrdinalIgnoreCase));
            PresetManager.NormalizePresetSortOrders(imageReplacePresets);
            EnsurePresetSelectionNames();
            SaveImageReplacePresets();
            SyncImageReplaceInputValuesFromCurrentPreset();
            InvalidateImageReplaceRibbonControls();
        }

        public string GetPresetMenuContent(Office.IRibbonControl control)
        {
            var sb = new StringBuilder();
            sb.Append("<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>");
            foreach (var preset in PresetManager.GetPresetsInDisplayOrder(imageReplacePresets))
            {
                var mark = string.Equals(preset.Name, currentPresetName, StringComparison.OrdinalIgnoreCase) ? " ✓" : string.Empty;
                sb.Append("<button id='Preset_")
                    .Append(XmlEscape(preset.Name))
                    .Append("' label='")
                    .Append(XmlEscape(preset.Name + mark))
                    .Append("' onAction='OnSelectImageReplacePreset'/>");
            }
            sb.Append("</menu>");
            return sb.ToString();
        }

        public void OnSelectImageReplacePreset(Office.IRibbonControl control)
        {
            if (control?.Id == null || !control.Id.StartsWith("Preset_", StringComparison.Ordinal))
            {
                return;
            }

            var presetName = control.Id.Substring("Preset_".Length);
            if (imageReplacePresets.Any(x => string.Equals(x.Name, presetName, StringComparison.OrdinalIgnoreCase)))
            {
                currentPresetName = presetName;
                SyncImageReplaceInputValuesFromCurrentPreset();
                SaveImageReplacePresets();
                InvalidateImageReplaceRibbonControls();
            }
        }

        #region AI 模型分页与回调

        private static readonly string[] AiSizeOptions =
        {
            "512", "768", "1024", "1536", "1792", "1920",
            "2048", "2160", "2560", "3072", "3840", "4096"
        };

        private int GetAiModelPageSize()
        {
            return Math.Max(1, Math.Min(AiImageService.AiModelButtonCount, aiModelPreviewCount));
        }

        private void EnsureAiModelPageIndexRange()
        {
            var list = aiImageSettingsList ?? new List<AiImageApiSettings>();
            var pageSize = GetAiModelPageSize();
            var totalPages = Math.Max(1, (int)Math.Ceiling((double)list.Count / pageSize));
            aiModelPageIndex = Math.Max(0, Math.Min(aiModelPageIndex, totalPages - 1));
        }

        private AiImageApiSettings GetAiModelEntryForButton(int buttonOffset)
        {
            var list = aiImageSettingsList ?? new List<AiImageApiSettings>();
            var pageSize = GetAiModelPageSize();
            EnsureAiModelPageIndexRange();
            if (buttonOffset < 0 || buttonOffset >= pageSize)
                return null;
            int index = aiModelPageIndex * pageSize + buttonOffset;
            if (index >= 0 && index < list.Count)
                return list[index];
            return null;
        }

        private bool IsAiModelButtonVisible(int buttonOffset)
        {
            return buttonOffset >= 0 && buttonOffset < GetAiModelPageSize();
        }

        private void InvalidateAiModelControls()
        {
            if (ribbon == null) return;
            ribbon.InvalidateControl("BtnAiModel1");
            ribbon.InvalidateControl("BtnAiModel2");
            ribbon.InvalidateControl("BtnAiModel3");
            ribbon.InvalidateControl("BtnAiModel4");
            ribbon.InvalidateControl("BtnAiModel5");
            ribbon.InvalidateControl("BtnAiModel6");
            ribbon.InvalidateControl("BtnAiModel7");
            ribbon.InvalidateControl("BtnAiModel8");
            ribbon.InvalidateControl("BtnAiModel9");
            ribbon.InvalidateControl("BtnAiModel10");
            ribbon.InvalidateControl("BtnAiModel11");
            ribbon.InvalidateControl("BtnAiModel12");
            ribbon.InvalidateControl("DdAiImageWidth");
            ribbon.InvalidateControl("DdAiImageHeight");
            ribbon.InvalidateControl("EdAiImageQuality");
            ribbon.InvalidateControl("BtnAiModelPageUp");
            ribbon.InvalidateControl("BtnAiModelPageDown");
        }

        private AiImageApiSettings GetEffectiveAiSettings()
        {
            if (aiGlobalSettings != null && aiGlobalSettings.OverridePerModel)
            {
                return new AiImageApiSettings
                {
                    DisplayName = "全局",
                    EndpointUrl = aiImageSettingsList?.FirstOrDefault()?.EndpointUrl ?? string.Empty,
                    ApiToken = aiImageSettingsList?.FirstOrDefault()?.ApiToken ?? string.Empty,
                    Model = aiImageSettingsList?.FirstOrDefault()?.Model ?? string.Empty,
                    DefaultWidth = aiGlobalSettings.DefaultWidth,
                    DefaultHeight = aiGlobalSettings.DefaultHeight,
                    DefaultQuality = aiGlobalSettings.DefaultQuality,
                    DefaultFormat = aiGlobalSettings.DefaultFormat,
                    IconPath = string.Empty,
                    LockAspectRatio = false
                };
            }
            if (selectedAiModelIndex >= 0
                && (aiImageSettingsList?.Count ?? 0) > selectedAiModelIndex)
            {
                return aiImageSettingsList[selectedAiModelIndex];
            }
            return aiImageSettingsList?.FirstOrDefault()
                ?? AiImageService.CreateDefaultSettings();
        }

        private AiImageApiSettings GetEffectiveImgSettings()
        {
            if (aiGlobalSettings != null && aiGlobalSettings.ImgOverridePerModel)
            {
                return new AiImageApiSettings
                {
                    DisplayName = "全局",
                    EndpointUrl = aiImageSettingsList?.FirstOrDefault()?.EndpointUrl ?? string.Empty,
                    ApiToken = aiImageSettingsList?.FirstOrDefault()?.ApiToken ?? string.Empty,
                    Model = aiImageSettingsList?.FirstOrDefault()?.Model ?? string.Empty,
                    DefaultWidth = aiGlobalSettings.ImgDefaultWidth,
                    DefaultHeight = aiGlobalSettings.ImgDefaultHeight,
                    DefaultQuality = aiGlobalSettings.ImgDefaultQuality,
                    DefaultFormat = aiGlobalSettings.ImgDefaultFormat,
                    IconPath = string.Empty,
                    LockAspectRatio = false
                };
            }
            if (selectedImgModelIndex >= 0
                && (aiImageSettingsList?.Count ?? 0) > selectedImgModelIndex)
            {
                return aiImageSettingsList[selectedImgModelIndex];
            }
            return aiImageSettingsList?.FirstOrDefault()
                ?? AiImageService.CreateDefaultSettings();
        }

        // ---- 可见性 ----

        public bool GetAiModelVisible1(Office.IRibbonControl control) { return IsAiModelButtonVisible(0); }
        public bool GetAiModelVisible2(Office.IRibbonControl control) { return IsAiModelButtonVisible(1); }
        public bool GetAiModelVisible3(Office.IRibbonControl control) { return IsAiModelButtonVisible(2); }
        public bool GetAiModelVisible4(Office.IRibbonControl control) { return IsAiModelButtonVisible(3); }
        public bool GetAiModelVisible5(Office.IRibbonControl control) { return IsAiModelButtonVisible(4); }
        public bool GetAiModelVisible6(Office.IRibbonControl control) { return IsAiModelButtonVisible(5); }
        public bool GetAiModelVisible7(Office.IRibbonControl control) { return IsAiModelButtonVisible(6); }
        public bool GetAiModelVisible8(Office.IRibbonControl control) { return IsAiModelButtonVisible(7); }
        public bool GetAiModelVisible9(Office.IRibbonControl control) { return IsAiModelButtonVisible(8); }
        public bool GetAiModelVisible10(Office.IRibbonControl control) { return IsAiModelButtonVisible(9); }
        public bool GetAiModelVisible11(Office.IRibbonControl control) { return IsAiModelButtonVisible(10); }
        public bool GetAiModelVisible12(Office.IRibbonControl control) { return IsAiModelButtonVisible(11); }

        // ---- 启用 ----

        public bool GetAiModelEnabled1(Office.IRibbonControl control) { return IsAiModelButtonVisible(0); }
        public bool GetAiModelEnabled2(Office.IRibbonControl control) { return IsAiModelButtonVisible(1); }
        public bool GetAiModelEnabled3(Office.IRibbonControl control) { return IsAiModelButtonVisible(2); }
        public bool GetAiModelEnabled4(Office.IRibbonControl control) { return IsAiModelButtonVisible(3); }
        public bool GetAiModelEnabled5(Office.IRibbonControl control) { return IsAiModelButtonVisible(4); }
        public bool GetAiModelEnabled6(Office.IRibbonControl control) { return IsAiModelButtonVisible(5); }
        public bool GetAiModelEnabled7(Office.IRibbonControl control) { return IsAiModelButtonVisible(6); }
        public bool GetAiModelEnabled8(Office.IRibbonControl control) { return IsAiModelButtonVisible(7); }
        public bool GetAiModelEnabled9(Office.IRibbonControl control) { return IsAiModelButtonVisible(8); }
        public bool GetAiModelEnabled10(Office.IRibbonControl control) { return IsAiModelButtonVisible(9); }
        public bool GetAiModelEnabled11(Office.IRibbonControl control) { return IsAiModelButtonVisible(10); }
        public bool GetAiModelEnabled12(Office.IRibbonControl control) { return IsAiModelButtonVisible(11); }

        // ---- 提示 ----

        public string GetAiModelScreentip1(Office.IRibbonControl control) { return GetAiModelTooltip(0); }
        public string GetAiModelScreentip2(Office.IRibbonControl control) { return GetAiModelTooltip(1); }
        public string GetAiModelScreentip3(Office.IRibbonControl control) { return GetAiModelTooltip(2); }
        public string GetAiModelScreentip4(Office.IRibbonControl control) { return GetAiModelTooltip(3); }
        public string GetAiModelScreentip5(Office.IRibbonControl control) { return GetAiModelTooltip(4); }
        public string GetAiModelScreentip6(Office.IRibbonControl control) { return GetAiModelTooltip(5); }
        public string GetAiModelScreentip7(Office.IRibbonControl control) { return GetAiModelTooltip(6); }
        public string GetAiModelScreentip8(Office.IRibbonControl control) { return GetAiModelTooltip(7); }
        public string GetAiModelScreentip9(Office.IRibbonControl control) { return GetAiModelTooltip(8); }
        public string GetAiModelScreentip10(Office.IRibbonControl control) { return GetAiModelTooltip(9); }
        public string GetAiModelScreentip11(Office.IRibbonControl control) { return GetAiModelTooltip(10); }
        public string GetAiModelScreentip12(Office.IRibbonControl control) { return GetAiModelTooltip(11); }

        // ---- 文生图 标签 ----

        public string GetAiModelLabel1(Office.IRibbonControl control) { return GetAiModelLabel(0); }
        public string GetAiModelLabel2(Office.IRibbonControl control) { return GetAiModelLabel(1); }
        public string GetAiModelLabel3(Office.IRibbonControl control) { return GetAiModelLabel(2); }
        public string GetAiModelLabel4(Office.IRibbonControl control) { return GetAiModelLabel(3); }
        public string GetAiModelLabel5(Office.IRibbonControl control) { return GetAiModelLabel(4); }
        public string GetAiModelLabel6(Office.IRibbonControl control) { return GetAiModelLabel(5); }
        public string GetAiModelLabel7(Office.IRibbonControl control) { return GetAiModelLabel(6); }
        public string GetAiModelLabel8(Office.IRibbonControl control) { return GetAiModelLabel(7); }
        public string GetAiModelLabel9(Office.IRibbonControl control) { return GetAiModelLabel(8); }
        public string GetAiModelLabel10(Office.IRibbonControl control) { return GetAiModelLabel(9); }
        public string GetAiModelLabel11(Office.IRibbonControl control) { return GetAiModelLabel(10); }
        public string GetAiModelLabel12(Office.IRibbonControl control) { return GetAiModelLabel(11); }

        private string GetAiModelLabel(int buttonOffset)
        {
            var entry = GetAiModelEntryForButton(buttonOffset);
            if (entry == null) return "+";
            return string.IsNullOrWhiteSpace(entry.DisplayName)
                ? (entry.Model ?? "未命名")
                : entry.DisplayName;
        }

        private string GetAiModelTooltip(int buttonOffset)
        {
            var entry = GetAiModelEntryForButton(buttonOffset);
            if (entry == null)
                return "文生图 — 无模型（Ctrl+单击添加）";
            return "文生图 — " + (string.IsNullOrWhiteSpace(entry.DisplayName)
                ? (entry.Model ?? "未知模型")
                : entry.DisplayName + " (" + (entry.Model ?? "未知") + ")");
        }

        // ---- 点击 ----

        public void OnAiModelClick1(Office.IRibbonControl control) { OnAiModelClick(0); }
        public void OnAiModelClick2(Office.IRibbonControl control) { OnAiModelClick(1); }
        public void OnAiModelClick3(Office.IRibbonControl control) { OnAiModelClick(2); }
        public void OnAiModelClick4(Office.IRibbonControl control) { OnAiModelClick(3); }
        public void OnAiModelClick5(Office.IRibbonControl control) { OnAiModelClick(4); }
        public void OnAiModelClick6(Office.IRibbonControl control) { OnAiModelClick(5); }
        public void OnAiModelClick7(Office.IRibbonControl control) { OnAiModelClick(6); }
        public void OnAiModelClick8(Office.IRibbonControl control) { OnAiModelClick(7); }
        public void OnAiModelClick9(Office.IRibbonControl control) { OnAiModelClick(8); }
        public void OnAiModelClick10(Office.IRibbonControl control) { OnAiModelClick(9); }
        public void OnAiModelClick11(Office.IRibbonControl control) { OnAiModelClick(10); }
        public void OnAiModelClick12(Office.IRibbonControl control) { OnAiModelClick(11); }

        private void OnAiModelClick(int buttonOffset)
        {
            var entry = GetAiModelEntryForButton(buttonOffset);

            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                if (entry == null)
                    entry = AiImageService.CreateDefaultSettings();
                OpenAiModelSettingsDialog(entry);
            }
            else
            {
                if (entry == null) return;
                selectedAiModelIndex = aiModelPageIndex * GetAiModelPageSize() + buttonOffset;
                OpenAiGenerationDialog(entry);
            }
        }

        private void OpenAiGenerationDialog(AiImageApiSettings entry)
        {
            var settingsList = aiImageSettingsList ?? AiImageService.LoadModelSettings();
            aiImageSettingsList = settingsList;

            var effective = GetEffectiveAiSettings();
            int w = effective.DefaultWidth;
            int h = effective.DefaultHeight;
            if (!string.IsNullOrWhiteSpace(aiWidthInputText)
                && int.TryParse(aiWidthInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedW)
                && parsedW > 0)
                w = parsedW;
            if (!string.IsNullOrWhiteSpace(aiHeightInputText)
                && int.TryParse(aiHeightInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedH)
                && parsedH > 0)
                h = parsedH;

            var quality = effective.DefaultQuality;
            if (!string.IsNullOrWhiteSpace(aiQualityInputText))
            {
                // Try to map display text back to a value
                for (int qi = 0; qi < AiQualityOptions.Length; qi++)
                {
                    if (string.Equals(AiQualityOptions[qi], aiQualityInputText, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(AiQualityValues[qi], aiQualityInputText, StringComparison.OrdinalIgnoreCase))
                    {
                        quality = AiQualityValues[qi];
                        break;
                    }
                }
            }

            using (var dialog = new AiImageWebDialog(entry, settingsList, w, h,
                quality, effective.DefaultFormat))
            {
                dialog.ShowDialog();
            }
        }

        private void OpenAiModelSettingsDialog(AiImageApiSettings entry)
        {
            var settingsList = aiImageSettingsList ?? AiImageService.LoadModelSettings();
            aiImageSettingsList = settingsList;
            aiGlobalSettings = aiGlobalSettings ?? AiImageService.LoadGlobalSettings();

            using (var dialog = new ModelSettingsWebDialog(entry, settingsList))
            {
                if (hasAiSettingsDialogBounds)
                {
                    dialog.StartPosition = FormStartPosition.Manual;
                    dialog.Bounds = aiSettingsDialogBounds;
                }
                dialog.FormClosed += (s, e) =>
                {
                    aiSettingsDialogBounds = dialog.Bounds;
                    hasAiSettingsDialogBounds = true;
                };
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    aiModelIconCache.Clear();
                    AiImageService.SaveSettings(settingsList, aiGlobalSettings);
                    AiImageService.SaveSettingsBounds(aiSettingsDialogBounds, null);
                    InvalidateAiModelControls();
                }
            }
        }

        // ---- 宽 ----

        public int GetAiWidthItemCount(Office.IRibbonControl control)
        {
            return AiSizeOptions.Length;
        }

        public string GetAiWidthItemLabel(Office.IRibbonControl control, int index)
        {
            if (index >= 0 && index < AiSizeOptions.Length)
                return AiSizeOptions[index];
            return string.Empty;
        }

        public string GetAiWidthText(Office.IRibbonControl control)
        {
            if (!string.IsNullOrWhiteSpace(aiWidthInputText))
                return aiWidthInputText;
            var effective = GetEffectiveAiSettings();
            return effective.DefaultWidth.ToString(CultureInfo.InvariantCulture);
        }

        public void OnAiWidthChanged(Office.IRibbonControl control, string text)
        {
            aiWidthInputText = (text ?? string.Empty).Trim();
            var effective = GetEffectiveAiSettings();
            if (effective != null && effective.LockAspectRatio
                && int.TryParse(aiWidthInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int w)
                && w > 0)
            {
                if (!int.TryParse(aiHeightInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                {
                    double ratio = (double)effective.DefaultHeight / effective.DefaultWidth;
                    aiHeightInputText = ((int)Math.Round(w * ratio)).ToString(CultureInfo.InvariantCulture);
                    if (ribbon != null)
                        ribbon.InvalidateControl("DdAiImageHeight");
                }
            }
        }

        // ---- 高 ----

        public int GetAiHeightItemCount(Office.IRibbonControl control)
        {
            return AiSizeOptions.Length;
        }

        public string GetAiHeightItemLabel(Office.IRibbonControl control, int index)
        {
            if (index >= 0 && index < AiSizeOptions.Length)
                return AiSizeOptions[index];
            return string.Empty;
        }

        public string GetAiHeightText(Office.IRibbonControl control)
        {
            if (!string.IsNullOrWhiteSpace(aiHeightInputText))
                return aiHeightInputText;
            var effective = GetEffectiveAiSettings();
            return effective.DefaultHeight.ToString(CultureInfo.InvariantCulture);
        }

        public void OnAiHeightChanged(Office.IRibbonControl control, string text)
        {
            aiHeightInputText = (text ?? string.Empty).Trim();
            var effective = GetEffectiveAiSettings();
            if (effective != null && effective.LockAspectRatio
                && int.TryParse(aiHeightInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int h)
                && h > 0)
            {
                if (!int.TryParse(aiWidthInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                {
                    double ratio = (double)effective.DefaultWidth / effective.DefaultHeight;
                    aiWidthInputText = ((int)Math.Round(h * ratio)).ToString(CultureInfo.InvariantCulture);
                    if (ribbon != null)
                        ribbon.InvalidateControl("DdAiImageWidth");
                }
            }
        }

        // ---- 质量 ----

        private static readonly string[] AiQualityOptions = { "auto (自动)", "low (低)", "medium (中)", "high (高)" };
        private static readonly string[] AiQualityValues = { "auto", "low", "medium", "high" };

        public int GetAiQualityItemCount(Office.IRibbonControl control)
        {
            return AiQualityOptions.Length;
        }

        public string GetAiQualityItemLabel(Office.IRibbonControl control, int index)
        {
            if (index >= 0 && index < AiQualityOptions.Length)
                return AiQualityOptions[index];
            return string.Empty;
        }

        public string GetAiQualityText(Office.IRibbonControl control)
        {
            if (!string.IsNullOrWhiteSpace(aiQualityInputText))
                return aiQualityInputText;
            var effective = GetEffectiveAiSettings();
            var quality = effective.DefaultQuality ?? "auto";
            var idx = Array.IndexOf(AiQualityValues, quality.ToLowerInvariant());
            return idx >= 0 ? AiQualityOptions[idx] : AiQualityOptions[0];
        }

        public void OnAiQualityChanged(Office.IRibbonControl control, string text)
        {
            aiQualityInputText = (text ?? string.Empty).Trim();
        }

        // ---- 翻页 ----

        public void OnAiModelPageUp(Office.IRibbonControl control)
        {
            if (aiModelPageIndex > 0)
            {
                aiModelPageIndex--;
                InvalidateAiModelControls();
            }
        }

        public void OnAiModelPageDown(Office.IRibbonControl control)
        {
            var list = aiImageSettingsList ?? new List<AiImageApiSettings>();
            var pageSize = GetAiModelPageSize();
            int totalPages = Math.Max(1, (int)Math.Ceiling((double)list.Count / pageSize));
            if (aiModelPageIndex < totalPages - 1)
            {
                aiModelPageIndex++;
                InvalidateAiModelControls();
            }
        }

        // ---- 全局设置 ----

        public void OnAiGlobalSettingsClick(Office.IRibbonControl control)
        {
            var settingsList = aiImageSettingsList ?? AiImageService.LoadModelSettings();
            aiImageSettingsList = settingsList;
            aiGlobalSettings = aiGlobalSettings ?? AiImageService.LoadGlobalSettings();

            using (var dialog = new GlobalSettingsWebDialog(aiGlobalSettings))
            {
                if (hasAiGlobalSettingsDialogBounds)
                {
                    dialog.StartPosition = FormStartPosition.Manual;
                    dialog.Bounds = aiGlobalSettingsDialogBounds;
                }
                dialog.FormClosed += (s, e) =>
                {
                    aiGlobalSettingsDialogBounds = dialog.Bounds;
                    hasAiGlobalSettingsDialogBounds = true;
                };
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    aiModelPreviewCount = AiImageService.ClampModelPreviewCount(aiGlobalSettings.ModelPreviewCount);
                    aiModelPageIndex = 0;
                    aiWidthInputText = string.Empty;
                    aiHeightInputText = string.Empty;
                    aiQualityInputText = string.Empty;
                    AiImageService.SaveSettings(settingsList, aiGlobalSettings);
                    AiImageService.SaveSettingsBounds(null, aiGlobalSettingsDialogBounds);
                    InvalidateAiModelControls();
                }
            }
        }

        public void OnImgGlobalSettingsClick(Office.IRibbonControl control)
        {
            var settingsList = aiImageSettingsList ?? AiImageService.LoadModelSettings();
            aiImageSettingsList = settingsList;
            aiGlobalSettings = aiGlobalSettings ?? AiImageService.LoadGlobalSettings();

            using (var dialog = new ImgGlobalSettingsWebDialog(aiGlobalSettings))
            {
                if (hasImgGlobalSettingsDialogBounds)
                {
                    dialog.StartPosition = FormStartPosition.Manual;
                    dialog.Bounds = imgGlobalSettingsDialogBounds;
                }
                dialog.FormClosed += (s, e) =>
                {
                    imgGlobalSettingsDialogBounds = dialog.Bounds;
                    hasImgGlobalSettingsDialogBounds = true;
                };
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    imgModelPreviewCount = AiImageService.ClampModelPreviewCount(aiGlobalSettings.ImgModelPreviewCount);
                    imgModelPageIndex = 0;
                    imgWidthInputText = string.Empty;
                    imgHeightInputText = string.Empty;
                    imgQualityInputText = string.Empty;
                    AiImageService.SaveSettings(settingsList, aiGlobalSettings);
                    AiImageService.SaveSettingsBounds(null, null, imgGlobalSettingsDialogBounds);
                    InvalidateImgModelControls();
                }
            }
        }

        // ---- 模型图标 ----

        private stdole.IPictureDisp GetAiModelIcon(AiImageApiSettings entry)
        {
            var cacheKey = entry.Model ?? string.Empty;
            if (aiModelIconCache.TryGetValue(cacheKey, out var cached))
                return cached;

            stdole.IPictureDisp icon = null;
            if (!string.IsNullOrWhiteSpace(entry.IconPath) && File.Exists(entry.IconPath))
            {
                try
                {
                    icon = LoadFileImageAsPicture(entry.IconPath);
                }
                catch { }
            }

            if (icon == null)
            {
                EnsureBrandImages();
                icon = plusButtonImage ?? aiImageButtonImage ?? brandImageLarge ?? brandImageSmall;
            }

            aiModelIconCache[cacheKey] = icon;
            return icon;
        }

        #endregion

        #region 图生图 分页与回调

        private int GetImgModelPageSize()
        {
            return Math.Max(1, Math.Min(AiImageService.AiModelButtonCount, imgModelPreviewCount));
        }

        private void EnsureImgModelPageIndexRange()
        {
            var list = aiImageSettingsList ?? new List<AiImageApiSettings>();
            var pageSize = GetImgModelPageSize();
            var totalPages = Math.Max(1, (int)Math.Ceiling((double)list.Count / pageSize));
            imgModelPageIndex = Math.Max(0, Math.Min(imgModelPageIndex, totalPages - 1));
        }

        private AiImageApiSettings GetImgModelEntryForButton(int buttonOffset)
        {
            var list = aiImageSettingsList ?? new List<AiImageApiSettings>();
            var pageSize = GetImgModelPageSize();
            EnsureImgModelPageIndexRange();
            if (buttonOffset < 0 || buttonOffset >= pageSize) return null;
            int index = imgModelPageIndex * pageSize + buttonOffset;
            return index >= 0 && index < list.Count ? list[index] : null;
        }

        private bool IsImgModelButtonVisible(int buttonOffset)
        {
            return buttonOffset >= 0 && buttonOffset < GetImgModelPageSize();
        }

        private bool IsImageSelected()
        {
            try
            {
                dynamic app = Globals.ThisAddIn?.Application;
                if (app == null) return false;
                dynamic sel = null;
                try { sel = app.ActiveWindow?.Selection; } catch { return false; }
                if (sel == null || sel.Type != 2) return false; // ppSelectionShapes = 2
                if (sel.ShapeRange.Count < 1) return false;
                dynamic shape = sel.ShapeRange[1];
                if (shape.Type == 13 || shape.Type == 11) return true; // msoPicture, msoLinkedPicture
                if (shape.Type == 6) // msoGroup — check first picture child
                {
                    try
                    {
                        foreach (var child in shape.GroupItems)
                        {
                            if (child.Type == 13 || child.Type == 11) return true;
                        }
                    }
                    catch { }
                }
                return false;
            }
            catch { return false; }
        }

        internal void InvalidateImgModelControls()
        {
            if (ribbon == null) return;
            ribbon.InvalidateControl("BtnImgModel1");
            ribbon.InvalidateControl("BtnImgModel2");
            ribbon.InvalidateControl("BtnImgModel3");
            ribbon.InvalidateControl("BtnImgModel4");
            ribbon.InvalidateControl("BtnImgModel5");
            ribbon.InvalidateControl("BtnImgModel6");
            ribbon.InvalidateControl("BtnImgModel7");
            ribbon.InvalidateControl("BtnImgModel8");
            ribbon.InvalidateControl("BtnImgModel9");
            ribbon.InvalidateControl("BtnImgModel10");
            ribbon.InvalidateControl("BtnImgModel11");
            ribbon.InvalidateControl("BtnImgModel12");
            ribbon.InvalidateControl("DdImgImageWidth");
            ribbon.InvalidateControl("DdImgImageHeight");
            ribbon.InvalidateControl("EdImgImageQuality");
            ribbon.InvalidateControl("BtnImgModelPageUp");
            ribbon.InvalidateControl("BtnImgModelPageDown");
        }

        private void GetImgDefaultDimensions(out int width, out int height)
        {
            AiImageService.GetSelectedImageDimensions(out int srcW, out int srcH);
            if (srcW <= 0 || srcH <= 0)
            {
                width = 1024;
                height = 1024;
                return;
            }
            if (srcW >= srcH)
            {
                width = 2048;
                height = Math.Max(1, (int)Math.Round(2048.0 * srcH / srcW));
            }
            else
            {
                height = 2048;
                width = Math.Max(1, (int)Math.Round(2048.0 * srcW / srcH));
            }
        }

        // ---- 图生图 可见性 ----

        public bool GetImgModelVisible1(Office.IRibbonControl control) { return IsImgModelButtonVisible(0); }
        public bool GetImgModelVisible2(Office.IRibbonControl control) { return IsImgModelButtonVisible(1); }
        public bool GetImgModelVisible3(Office.IRibbonControl control) { return IsImgModelButtonVisible(2); }
        public bool GetImgModelVisible4(Office.IRibbonControl control) { return IsImgModelButtonVisible(3); }
        public bool GetImgModelVisible5(Office.IRibbonControl control) { return IsImgModelButtonVisible(4); }
        public bool GetImgModelVisible6(Office.IRibbonControl control) { return IsImgModelButtonVisible(5); }
        public bool GetImgModelVisible7(Office.IRibbonControl control) { return IsImgModelButtonVisible(6); }
        public bool GetImgModelVisible8(Office.IRibbonControl control) { return IsImgModelButtonVisible(7); }
        public bool GetImgModelVisible9(Office.IRibbonControl control) { return IsImgModelButtonVisible(8); }
        public bool GetImgModelVisible10(Office.IRibbonControl control) { return IsImgModelButtonVisible(9); }
        public bool GetImgModelVisible11(Office.IRibbonControl control) { return IsImgModelButtonVisible(10); }
        public bool GetImgModelVisible12(Office.IRibbonControl control) { return IsImgModelButtonVisible(11); }

        // ---- 图生图 启用（需选中图片） ----

        public bool GetImgModelEnabled1(Office.IRibbonControl control) { return IsImgModelButtonVisible(0) && IsImageSelected(); }
        public bool GetImgModelEnabled2(Office.IRibbonControl control) { return IsImgModelButtonVisible(1) && IsImageSelected(); }
        public bool GetImgModelEnabled3(Office.IRibbonControl control) { return IsImgModelButtonVisible(2) && IsImageSelected(); }
        public bool GetImgModelEnabled4(Office.IRibbonControl control) { return IsImgModelButtonVisible(3) && IsImageSelected(); }
        public bool GetImgModelEnabled5(Office.IRibbonControl control) { return IsImgModelButtonVisible(4) && IsImageSelected(); }
        public bool GetImgModelEnabled6(Office.IRibbonControl control) { return IsImgModelButtonVisible(5) && IsImageSelected(); }
        public bool GetImgModelEnabled7(Office.IRibbonControl control) { return IsImgModelButtonVisible(6) && IsImageSelected(); }
        public bool GetImgModelEnabled8(Office.IRibbonControl control) { return IsImgModelButtonVisible(7) && IsImageSelected(); }
        public bool GetImgModelEnabled9(Office.IRibbonControl control) { return IsImgModelButtonVisible(8) && IsImageSelected(); }
        public bool GetImgModelEnabled10(Office.IRibbonControl control) { return IsImgModelButtonVisible(9) && IsImageSelected(); }
        public bool GetImgModelEnabled11(Office.IRibbonControl control) { return IsImgModelButtonVisible(10) && IsImageSelected(); }
        public bool GetImgModelEnabled12(Office.IRibbonControl control) { return IsImgModelButtonVisible(11) && IsImageSelected(); }

        // ---- 图生图 图片 ----

        public stdole.IPictureDisp GetImgModelImage1(Office.IRibbonControl control) { return GetImgModelImageForButton(0); }
        public stdole.IPictureDisp GetImgModelImage2(Office.IRibbonControl control) { return GetImgModelImageForButton(1); }
        public stdole.IPictureDisp GetImgModelImage3(Office.IRibbonControl control) { return GetImgModelImageForButton(2); }
        public stdole.IPictureDisp GetImgModelImage4(Office.IRibbonControl control) { return GetImgModelImageForButton(3); }
        public stdole.IPictureDisp GetImgModelImage5(Office.IRibbonControl control) { return GetImgModelImageForButton(4); }
        public stdole.IPictureDisp GetImgModelImage6(Office.IRibbonControl control) { return GetImgModelImageForButton(5); }
        public stdole.IPictureDisp GetImgModelImage7(Office.IRibbonControl control) { return GetImgModelImageForButton(6); }
        public stdole.IPictureDisp GetImgModelImage8(Office.IRibbonControl control) { return GetImgModelImageForButton(7); }
        public stdole.IPictureDisp GetImgModelImage9(Office.IRibbonControl control) { return GetImgModelImageForButton(8); }
        public stdole.IPictureDisp GetImgModelImage10(Office.IRibbonControl control) { return GetImgModelImageForButton(9); }
        public stdole.IPictureDisp GetImgModelImage11(Office.IRibbonControl control) { return GetImgModelImageForButton(10); }
        public stdole.IPictureDisp GetImgModelImage12(Office.IRibbonControl control) { return GetImgModelImageForButton(11); }

        private stdole.IPictureDisp GetImgModelImageForButton(int buttonOffset)
        {
            EnsureBrandImages();
            var entry = GetImgModelEntryForButton(buttonOffset);
            if (entry == null)
                return plusButtonImage ?? aiImageButtonImage ?? brandImageLarge ?? brandImageSmall;
            return GetAiModelIcon(entry);
        }

        // ---- 图生图 提示 ----

        public string GetImgModelScreentip1(Office.IRibbonControl control) { return GetImgModelTooltip(0); }
        public string GetImgModelScreentip2(Office.IRibbonControl control) { return GetImgModelTooltip(1); }
        public string GetImgModelScreentip3(Office.IRibbonControl control) { return GetImgModelTooltip(2); }
        public string GetImgModelScreentip4(Office.IRibbonControl control) { return GetImgModelTooltip(3); }
        public string GetImgModelScreentip5(Office.IRibbonControl control) { return GetImgModelTooltip(4); }
        public string GetImgModelScreentip6(Office.IRibbonControl control) { return GetImgModelTooltip(5); }
        public string GetImgModelScreentip7(Office.IRibbonControl control) { return GetImgModelTooltip(6); }
        public string GetImgModelScreentip8(Office.IRibbonControl control) { return GetImgModelTooltip(7); }
        public string GetImgModelScreentip9(Office.IRibbonControl control) { return GetImgModelTooltip(8); }
        public string GetImgModelScreentip10(Office.IRibbonControl control) { return GetImgModelTooltip(9); }
        public string GetImgModelScreentip11(Office.IRibbonControl control) { return GetImgModelTooltip(10); }
        public string GetImgModelScreentip12(Office.IRibbonControl control) { return GetImgModelTooltip(11); }

        // ---- 图生图 标签 ----

        public string GetImgModelLabel1(Office.IRibbonControl control) { return GetImgModelLabel(0); }
        public string GetImgModelLabel2(Office.IRibbonControl control) { return GetImgModelLabel(1); }
        public string GetImgModelLabel3(Office.IRibbonControl control) { return GetImgModelLabel(2); }
        public string GetImgModelLabel4(Office.IRibbonControl control) { return GetImgModelLabel(3); }
        public string GetImgModelLabel5(Office.IRibbonControl control) { return GetImgModelLabel(4); }
        public string GetImgModelLabel6(Office.IRibbonControl control) { return GetImgModelLabel(5); }
        public string GetImgModelLabel7(Office.IRibbonControl control) { return GetImgModelLabel(6); }
        public string GetImgModelLabel8(Office.IRibbonControl control) { return GetImgModelLabel(7); }
        public string GetImgModelLabel9(Office.IRibbonControl control) { return GetImgModelLabel(8); }
        public string GetImgModelLabel10(Office.IRibbonControl control) { return GetImgModelLabel(9); }
        public string GetImgModelLabel11(Office.IRibbonControl control) { return GetImgModelLabel(10); }
        public string GetImgModelLabel12(Office.IRibbonControl control) { return GetImgModelLabel(11); }

        private string GetImgModelLabel(int buttonOffset)
        {
            var entry = GetImgModelEntryForButton(buttonOffset);
            if (entry == null) return "+";
            return string.IsNullOrWhiteSpace(entry.DisplayName)
                ? (entry.Model ?? "未命名")
                : entry.DisplayName;
        }

        private string GetImgModelTooltip(int buttonOffset)
        {
            var entry = GetImgModelEntryForButton(buttonOffset);
            if (entry == null) return "图生图 — 无模型（Ctrl+单击添加）";
            return "图生图 — " + (string.IsNullOrWhiteSpace(entry.DisplayName)
                ? (entry.Model ?? "未知模型")
                : entry.DisplayName + " (" + (entry.Model ?? "未知") + ")");
        }

        // ---- 图生图 点击 ----

        public void OnImgModelClick1(Office.IRibbonControl control) { OnImgModelClick(0); }
        public void OnImgModelClick2(Office.IRibbonControl control) { OnImgModelClick(1); }
        public void OnImgModelClick3(Office.IRibbonControl control) { OnImgModelClick(2); }
        public void OnImgModelClick4(Office.IRibbonControl control) { OnImgModelClick(3); }
        public void OnImgModelClick5(Office.IRibbonControl control) { OnImgModelClick(4); }
        public void OnImgModelClick6(Office.IRibbonControl control) { OnImgModelClick(5); }
        public void OnImgModelClick7(Office.IRibbonControl control) { OnImgModelClick(6); }
        public void OnImgModelClick8(Office.IRibbonControl control) { OnImgModelClick(7); }
        public void OnImgModelClick9(Office.IRibbonControl control) { OnImgModelClick(8); }
        public void OnImgModelClick10(Office.IRibbonControl control) { OnImgModelClick(9); }
        public void OnImgModelClick11(Office.IRibbonControl control) { OnImgModelClick(10); }
        public void OnImgModelClick12(Office.IRibbonControl control) { OnImgModelClick(11); }

        private void OnImgModelClick(int buttonOffset)
        {
            if (!IsImageSelected()) return;
            var entry = GetImgModelEntryForButton(buttonOffset);

            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                if (entry == null) entry = AiImageService.CreateDefaultSettings();
                OpenAiModelSettingsDialog(entry);
            }
            else
            {
                if (entry == null) return;
                selectedImgModelIndex = imgModelPageIndex * GetImgModelPageSize() + buttonOffset;
                OpenImgToImgDialog(entry);
            }
        }

        private void OpenImgToImgDialog(AiImageApiSettings entry)
        {
            var settingsList = aiImageSettingsList ?? AiImageService.LoadModelSettings();
            aiImageSettingsList = settingsList;

            var effective = GetEffectiveImgSettings();
            // Default size from selected image: long edge = 2048
            GetImgDefaultDimensions(out int w, out int h);
            if (!string.IsNullOrWhiteSpace(imgWidthInputText)
                && int.TryParse(imgWidthInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedW)
                && parsedW > 0)
                w = parsedW;
            if (!string.IsNullOrWhiteSpace(imgHeightInputText)
                && int.TryParse(imgHeightInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedH)
                && parsedH > 0)
                h = parsedH;

            var quality = effective.DefaultQuality;
            if (!string.IsNullOrWhiteSpace(imgQualityInputText))
            {
                for (int qi = 0; qi < AiQualityOptions.Length; qi++)
                {
                    if (string.Equals(AiQualityOptions[qi], imgQualityInputText, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(AiQualityValues[qi], imgQualityInputText, StringComparison.OrdinalIgnoreCase))
                    {
                        quality = AiQualityValues[qi];
                        break;
                    }
                }
            }

            using (var dialog = new ImgToImgWebDialog(entry, settingsList, w, h, quality, effective.DefaultFormat))
            {
                dialog.ShowDialog();
            }
        }

        // ---- 图生图 宽度 ----

        public int GetImgWidthItemCount(Office.IRibbonControl control) => AiSizeOptions.Length;
        public string GetImgWidthItemLabel(Office.IRibbonControl control, int index)
        {
            return index >= 0 && index < AiSizeOptions.Length ? AiSizeOptions[index] : string.Empty;
        }

        public string GetImgWidthText(Office.IRibbonControl control)
        {
            if (!string.IsNullOrWhiteSpace(imgWidthInputText)) return imgWidthInputText;
            GetImgDefaultDimensions(out int w, out _);
            return w.ToString(CultureInfo.InvariantCulture);
        }

        public void OnImgWidthChanged(Office.IRibbonControl control, string text)
        {
            imgWidthInputText = (text ?? string.Empty).Trim();
            var effective = GetEffectiveImgSettings();
            if (effective != null && effective.LockAspectRatio
                && int.TryParse(imgWidthInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int w)
                && w > 0)
            {
                if (!int.TryParse(imgHeightInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                {
                    GetImgDefaultDimensions(out _, out int defH);
                    double ratio = (double)defH / Math.Max(1, w);
                    imgHeightInputText = ((int)Math.Round(w * ratio)).ToString(CultureInfo.InvariantCulture);
                    if (ribbon != null) ribbon.InvalidateControl("DdImgImageHeight");
                }
            }
        }

        // ---- 图生图 高度 ----

        public int GetImgHeightItemCount(Office.IRibbonControl control) => AiSizeOptions.Length;
        public string GetImgHeightItemLabel(Office.IRibbonControl control, int index)
        {
            return index >= 0 && index < AiSizeOptions.Length ? AiSizeOptions[index] : string.Empty;
        }

        public string GetImgHeightText(Office.IRibbonControl control)
        {
            if (!string.IsNullOrWhiteSpace(imgHeightInputText)) return imgHeightInputText;
            GetImgDefaultDimensions(out _, out int h);
            return h.ToString(CultureInfo.InvariantCulture);
        }

        public void OnImgHeightChanged(Office.IRibbonControl control, string text)
        {
            imgHeightInputText = (text ?? string.Empty).Trim();
            var effective = GetEffectiveImgSettings();
            if (effective != null && effective.LockAspectRatio
                && int.TryParse(imgHeightInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int h)
                && h > 0)
            {
                if (!int.TryParse(imgWidthInputText, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                {
                    GetImgDefaultDimensions(out int defW, out _);
                    double ratio = (double)defW / Math.Max(1, h);
                    imgWidthInputText = ((int)Math.Round(h * ratio)).ToString(CultureInfo.InvariantCulture);
                    if (ribbon != null) ribbon.InvalidateControl("DdImgImageWidth");
                }
            }
        }

        // ---- 图生图 质量 ----

        public int GetImgQualityItemCount(Office.IRibbonControl control) => AiQualityOptions.Length;
        public string GetImgQualityItemLabel(Office.IRibbonControl control, int index)
        {
            return index >= 0 && index < AiQualityOptions.Length ? AiQualityOptions[index] : string.Empty;
        }

        public string GetImgQualityText(Office.IRibbonControl control)
        {
            if (!string.IsNullOrWhiteSpace(imgQualityInputText)) return imgQualityInputText;
            var effective = GetEffectiveImgSettings();
            var quality = effective.DefaultQuality ?? "auto";
            var idx = Array.IndexOf(AiQualityValues, quality.ToLowerInvariant());
            return idx >= 0 ? AiQualityOptions[idx] : AiQualityOptions[0];
        }

        public void OnImgQualityChanged(Office.IRibbonControl control, string text)
        {
            imgQualityInputText = (text ?? string.Empty).Trim();
        }

        // ---- 图生图 翻页 ----

        public void OnImgModelPageUp(Office.IRibbonControl control)
        {
            if (imgModelPageIndex > 0) { imgModelPageIndex--; InvalidateImgModelControls(); }
        }

        public void OnImgModelPageDown(Office.IRibbonControl control)
        {
            var list = aiImageSettingsList ?? new List<AiImageApiSettings>();
            var pageSize = GetImgModelPageSize();
            int totalPages = Math.Max(1, (int)Math.Ceiling((double)list.Count / pageSize));
            if (imgModelPageIndex < totalPages - 1) { imgModelPageIndex++; InvalidateImgModelControls(); }
        }

        // ---- 图生图 页按钮图片 ----

        public stdole.IPictureDisp GetImgPageButtonImage(Office.IRibbonControl control)
        {
            if (control != null && string.Equals(control.Id, "BtnImgModelPageUp", StringComparison.Ordinal))
            {
                if (pageUpButtonImage == null)
                    pageUpButtonImage = CreateSvgChevronButtonImage(true);
                return pageUpButtonImage ?? brandImageSmall ?? brandImageLarge;
            }
            if (pageDownButtonImage == null)
                pageDownButtonImage = CreateSvgChevronButtonImage(false);
            return pageDownButtonImage ?? brandImageSmall ?? brandImageLarge;
        }

        #endregion

        #endregion

        #region 帮助器

        private static string GetResourceText(string resourceName)
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], System.StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var resourceReader = new System.IO.StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        private void EnsureBrandImages()
        {
            if (brandImageLarge == null)
            {
                var brandLargePath = ResolveBioDrawIconFilePath(BrandPngFileName);
                brandImageLarge = LoadFileImageAsPicture(brandLargePath)
                    ?? LoadEmbeddedPngAsPicture(BrandPngResourceName, BrandPngFileName);
            }
            if (brandImageSmall == null)
            {
                var brandSmallPath = ResolveBioDrawIconFilePath(BrandPngFileName);
                brandImageSmall = LoadFileImageAsPicture(brandSmallPath)
                    ?? LoadEmbeddedPngAsPicture(BrandPngResourceName, BrandPngFileName);
            }
            if (brandImageLarge == null)
            {
                brandImageLarge = brandImageSmall;
            }
            if (brandImageSmall == null)
            {
                brandImageSmall = brandImageLarge;
            }
            if (transparentPlaceholderImage == null)
            {
                transparentPlaceholderImage = LoadEmbeddedPngAsPicture(TransparentPlaceholderResourceName, "blank-image-200x200.png");
            }
            if (pickerButtonImage == null)
            {
                var pickerFilePath = ResolveBioDrawIconFilePath(PickerColorFileName) ?? ResolveBioDrawIconFilePath(LegacyPickerColorFileName);
                pickerButtonImage = LoadFileImageAsPicture(pickerFilePath)
                    ?? LoadEmbeddedPngAsPicture(PickerColorResourceName, PickerColorFileName)
                    ?? LoadEmbeddedPngAsPicture("BioDraw.BioDrawIcon.Picker_color .png", LegacyPickerColorFileName);
            }
            if (settingsButtonImage == null)
            {
                var settingsPngPath = ResolveBioDrawIconFilePath(SettingsGearPngFileName);
                settingsButtonImage = LoadFileImageAsPicture(settingsPngPath)
                    ?? LoadEmbeddedPngAsPicture(SettingsGearResourceName, SettingsGearPngFileName);
            }
            if (imageRecolorButtonImage == null)
            {
                var imageRecolorPath = ResolveBioDrawIconFilePath(ImageRecolorPngFileName);
                imageRecolorButtonImage = LoadFileImageAsPicture(imageRecolorPath)
                    ?? LoadEmbeddedPngAsPicture(ImageRecolorResourceName, ImageRecolorPngFileName);
            }
            if (aiImageButtonImage == null)
            {
                aiImageButtonImage = CreateAiIconImage();
            }
            if (plusButtonImage == null)
            {
                plusButtonImage = LoadEmbeddedPngAsPicture(PlusPngResourceName, "plus.png");
            }
            if (addToLibraryContextMenuImage == null)
            {
                var addToLibraryPath = ResolveBioDrawIconFilePath(AddToLibraryIcoFileName);
                addToLibraryContextMenuImage = LoadFileImageAsPicture(addToLibraryPath)
                    ?? settingsButtonImage
                    ?? brandImageSmall
                    ?? brandImageLarge;
            }
        }

        private static stdole.IPictureDisp LoadEmbeddedPngAsPicture(string exactName, string suffixName)
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            var resourceName = asm.GetManifestResourceNames()
                .FirstOrDefault(name => string.Equals(name, exactName, System.StringComparison.OrdinalIgnoreCase));

            if (resourceName == null)
            {
                resourceName = asm.GetManifestResourceNames()
                    .FirstOrDefault(name => name.EndsWith(suffixName, System.StringComparison.OrdinalIgnoreCase));
            }

            if (resourceName == null)
            {
                return null;
            }

            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }

                using (var image = Image.FromStream(stream))
                {
                    return PictureConverter.ToPictureDisp(new Bitmap(image));
                }
            }
        }

        private static stdole.IPictureDisp LoadFileImageAsPicture(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return null;
            }

            try
            {
                using (var image = Image.FromFile(filePath))
                {
                    return PictureConverter.ToPictureDisp(new Bitmap(image));
                }
            }
            catch
            {
                return null;
            }
        }

        private static string ResolveBioDrawIconFilePath(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return null;
            }

            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            if (string.IsNullOrWhiteSpace(baseDir))
            {
                return null;
            }

            try
            {
                var current = new DirectoryInfo(baseDir);
                for (int i = 0; i < 8 && current != null; i++)
                {
                    var candidate = Path.Combine(current.FullName, "BioDrawIcon", fileName);
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                    current = current.Parent;
                }
            }
            catch
            {
            }

            return null;
        }

        private static string XmlEscape(string value)
        {
            return SecurityElement.Escape(value) ?? string.Empty;
        }

        private ImageReplacePreset GetCurrentPreset()
        {
            if (imageReplacePresets.Count == 0)
            {
                return null;
            }

            var preset = imageReplacePresets.FirstOrDefault(x => string.Equals(x.Name, currentPresetName, StringComparison.OrdinalIgnoreCase));
            if (preset != null)
            {
                return preset;
            }

            preset = imageReplacePresets.FirstOrDefault(x => string.Equals(x.Name, defaultPresetName, StringComparison.OrdinalIgnoreCase));
            if (preset != null)
            {
                currentPresetName = preset.Name;
                return preset;
            }

            currentPresetName = imageReplacePresets[0].Name;
            return imageReplacePresets[0];
        }

        private void EnsureImageReplaceInputValues()
        {
            EnsureImageReplaceColorOptions();
        }

        private void SyncImageReplaceInputValuesFromCurrentPreset()
        {
            var preset = GetCurrentPreset();
            if (preset == null)
            {
                return;
            }

            imageReplaceSourceColorInput = preset.TargetColor ?? string.Empty;
            imageReplaceFuzzInput = PresetManager.NormalizeFuzzPercent(preset.FuzzPercent);
            if (string.Equals(preset.Mode, "fill", StringComparison.OrdinalIgnoreCase))
            {
                imageReplaceNewColorInput = preset.ReplacementColor ?? string.Empty;
                return;
            }

            imageReplaceNewColorInput = string.Empty;
        }

        private void ResetImageReplaceColorOptions()
        {
            imageReplaceSourceColorOptions.Clear();
            imageReplaceNewColorOptions.Clear();
            PresetManager.AddColorOption(imageReplaceSourceColorOptions, "white");
            PresetManager.AddColorOption(imageReplaceNewColorOptions, "white");
        }


        private void EnsureImageReplaceColorOptions()
        {
            if (imageReplaceSourceColorOptions.Count == 0)
            {
                ResetImageReplaceColorOptions();
            }
        }

        private bool ShowColorOptionManagerDialog(string title, List<string> options, string currentValue, bool allowEmpty, out string selectedValue)
        {
            selectedValue = PresetManager.NormalizeColorInputText(currentValue);
            if (options == null)
            {
                return false;
            }

            using (var dialog = new ColorPresetWebDialog(
                title, options, selectedValue, allowEmpty,
                onPickColor: (initial) =>
                {
                    string colorToken;
                    string errorMessage;
                    if (TryPickColorWithPowerPoint(false, initial, out colorToken, out errorMessage))
                        return colorToken;
                    if (!string.IsNullOrWhiteSpace(errorMessage))
                        MessageBox.Show(errorMessage, "BioDraw");
                    return null;
                },
                onPersist: () => PersistImageReplaceInputMemory(),
                onInvalidate: () => InvalidateImageReplaceRibbonControls()))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return false;
                }
                selectedValue = dialog.SelectedValue;
                return true;
            }
        }

        private void PersistImageReplaceInputMemory()
        {
            SaveImageReplacePresets();
        }

        private bool TryAddSelectionToCurrentMaterialFolder(out string message)
        {
            message = string.Empty;
            var app = Globals.ThisAddIn?.Application;
            if (app == null)
            {
                message = "未能获取 PowerPoint 应用实例。";
                return false;
            }

            if (!UseCustomMaterialLibrary())
            {
                message = "请先在“关于 -> 素材库”中设置素材库目录。";
                return false;
            }

            dynamic selection = null;
            try
            {
                selection = app.ActiveWindow?.Selection;
            }
            catch
            {
            }

            List<dynamic> shapes;
            if (!ImageReplacePipeline.TryGetSelectedShapes(selection, out shapes) || shapes.Count == 0)
            {
                message = "请先选中一个或多个对象。";
                return false;
            }

            string targetDirectory;
            string targetDirectoryError;
            if (!TryResolveCurrentMaterialTargetDirectory(out targetDirectory, out targetDirectoryError))
            {
                message = targetDirectoryError;
                return false;
            }

            string snapshotPath = null;
            string snapshotError = string.Empty;
            if (!ImageReplacePipeline.TryCreatePresentationSnapshot(app, out snapshotPath, out snapshotError))
            {
                snapshotPath = null;
            }

            var savedCount = 0;
            var failedCount = 0;
            string lastError = string.Empty;
            foreach (var shape in shapes)
            {
                string savedPath;
                string saveError;
                if (TrySaveShapeToMaterialFolder(shape, snapshotPath, targetDirectory, out savedPath, out saveError))
                {
                    savedCount++;
                }
                else
                {
                    failedCount++;
                    lastError = saveError;
                }
            }

            ImageReplacePipeline.TryDeleteFile(snapshotPath);
            materialSearchCacheRootPath = null;
            materialSearchCacheEntries = null;
            materialPageIndex = 0;
            ribbon?.InvalidateControl("DdLevel1");
            ribbon?.InvalidateControl("DdLevel2");
            InvalidateMaterialPreview();

            if (savedCount == 0)
            {
                if (!string.IsNullOrWhiteSpace(lastError))
                {
                    message = "保存失败：" + lastError;
                }
                else if (!string.IsNullOrWhiteSpace(snapshotError))
                {
                    message = "保存失败：" + snapshotError;
                }
                else
                {
                    message = "保存失败：未找到可保存的对象。";
                }
                return false;
            }

            message = failedCount > 0
                ? $"BioDraw：已保存 {savedCount} 个对象到当前素材目录，{failedCount} 个失败。"
                : $"BioDraw：已保存 {savedCount} 个对象到当前素材目录。";
            return true;
        }

        private bool HasSelectedShapes()
        {
            var app = Globals.ThisAddIn?.Application;
            if (app == null)
            {
                return false;
            }

            dynamic selection = null;
            try
            {
                selection = app.ActiveWindow?.Selection;
            }
            catch
            {
            }

            List<dynamic> shapes;
            return ImageReplacePipeline.TryGetSelectedShapes(selection, out shapes) && shapes.Count > 0;
        }

        private bool TryResolveCurrentMaterialTargetDirectory(out string targetDirectory, out string errorMessage)
        {
            targetDirectory = null;
            errorMessage = string.Empty;
            if (!UseCustomMaterialLibrary())
            {
                errorMessage = "请先在“关于 -> 素材库”中设置素材库目录。";
                return false;
            }

            var level1List = GetLevel1List();
            if (level1List.Count == 0)
            {
                errorMessage = "素材库目录结构不可用。";
                return false;
            }

            var level1 = level1List[MaterialLibraryService.NormalizeIndex(selectedLevel1Index, level1List.Count)];
            var level1Path = Path.Combine(materialLibraryPath, level1);
            Directory.CreateDirectory(level1Path);

            bool hasSubDirectories;
            try
            {
                hasSubDirectories = Directory.EnumerateDirectories(level1Path).Any();
            }
            catch
            {
                hasSubDirectories = false;
            }

            if (!hasSubDirectories)
            {
                targetDirectory = level1Path;
                return true;
            }

            var level2List = GetLevel2List();
            if (level2List.Count == 0)
            {
                targetDirectory = level1Path;
                return true;
            }

            var level2 = level2List[MaterialLibraryService.NormalizeIndex(selectedLevel2Index, level2List.Count)];
            var level2Path = Path.Combine(level1Path, level2);
            Directory.CreateDirectory(level2Path);
            targetDirectory = level2Path;
            return true;
        }

        private bool TrySaveShapeToMaterialFolder(dynamic shape, string snapshotPath, string targetDirectory, out string savedPath, out string errorMessage)
        {
            savedPath = null;
            errorMessage = string.Empty;
            if (shape == null)
            {
                errorMessage = "对象为空。";
                return false;
            }

            var shapeName = GetShapeDisplayName(shape);
            var type = 0;
            try
            {
                type = (int)shape.Type;
            }
            catch
            {
            }

            // 图片对象尽量保持原始媒体格式（png/jpg/svg/...）
            if (type == 13 || type == 11)
            {
                if (string.IsNullOrWhiteSpace(snapshotPath))
                {
                    errorMessage = "无法读取演示文稿快照。";
                    return false;
                }

                string sourcePath;
                if (!ImageReplacePipeline.TryExtractOriginalImageFromPptx(shape, snapshotPath, out sourcePath, out errorMessage))
                {
                    return false;
                }

                try
                {
                    var ext = Path.GetExtension(sourcePath);
                    if (string.IsNullOrWhiteSpace(ext))
                    {
                        ext = ".png";
                    }

                    savedPath = BuildUniqueFilePath(targetDirectory, shapeName, ext);
                    File.Copy(sourcePath, savedPath, false);
                    return true;
                }
                catch (Exception ex)
                {
                    errorMessage = ex.Message;
                    return false;
                }
                finally
                {
                    ImageReplacePipeline.TryDeleteFile(sourcePath);
                }
            }

            try
            {
                // 原生形状、文本框、组合等优先导出为 SVG。
                savedPath = BuildUniqueFilePath(targetDirectory, shapeName, ".svg");
                shape.Export(savedPath, 6);
                if (File.Exists(savedPath) && new FileInfo(savedPath).Length > 0)
                {
                    return true;
                }
            }
            catch
            {
            }

            try
            {
                // 兜底使用 PNG，保证至少可保存。
                savedPath = BuildUniqueFilePath(targetDirectory, shapeName, ".png");
                shape.Export(savedPath, 2);
                if (File.Exists(savedPath) && new FileInfo(savedPath).Length > 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }

            errorMessage = "导出对象失败。";
            return false;
        }

        private static string GetShapeDisplayName(dynamic shape)
        {
            try
            {
                var value = (string)shape.Name;
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }
            catch
            {
            }

            return "Object";
        }

        private static string BuildUniqueFilePath(string directoryPath, string preferredName, string extension)
        {
            var stem = SanitizeFileNameWithoutExtension(preferredName);
            if (string.IsNullOrWhiteSpace(stem))
            {
                stem = "Object";
            }

            if (string.IsNullOrWhiteSpace(extension))
            {
                extension = ".png";
            }
            if (!extension.StartsWith(".", StringComparison.Ordinal))
            {
                extension = "." + extension;
            }

            var path = Path.Combine(directoryPath, stem + extension);
            if (!File.Exists(path))
            {
                return path;
            }

            for (int i = 1; i < 10000; i++)
            {
                path = Path.Combine(directoryPath, stem + "_" + i.ToString(CultureInfo.InvariantCulture) + extension);
                if (!File.Exists(path))
                {
                    return path;
                }
            }

            return Path.Combine(
                directoryPath,
                stem + "_" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture) + extension);
        }

        private static string SanitizeFileNameWithoutExtension(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var invalidChars = Path.GetInvalidFileNameChars();
            var buffer = new StringBuilder(value.Length);
            for (int i = 0; i < value.Length; i++)
            {
                var ch = value[i];
                if (invalidChars.Contains(ch))
                {
                    buffer.Append('_');
                }
                else
                {
                    buffer.Append(ch);
                }
            }

            var result = buffer.ToString().Trim();
            return string.IsNullOrWhiteSpace(result) ? "Object" : result;
        }

        private static bool TryBuildManagedDirectoryPath(
            string parentPath,
            string requestedName,
            out string fullPath,
            out string errorMessage)
        {
            fullPath = string.Empty;
            errorMessage = string.Empty;
            var name = (requestedName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                errorMessage = "请输入目录名称。";
                return false;
            }
            if (name.Length > 100 || name == "." || name == ".." ||
                name.EndsWith(".", StringComparison.Ordinal) ||
                name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                IsReservedWindowsName(name))
            {
                errorMessage = "目录名称无效；不能包含路径分隔符、保留名称或非法字符。";
                return false;
            }

            try
            {
                var parent = Path.GetFullPath(parentPath ?? string.Empty);
                var candidate = Path.GetFullPath(Path.Combine(parent, name));
                if (!IsPathWithinRoot(parent, candidate))
                {
                    errorMessage = "目录路径超出素材库范围，已拒绝操作。";
                    return false;
                }
                fullPath = candidate;
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = "目录路径无效：" + ex.Message;
                return false;
            }
        }

        private static bool IsReservedWindowsName(string name)
        {
            var stem = (name ?? string.Empty).Split('.')[0].ToUpperInvariant();
            if (stem == "CON" || stem == "PRN" || stem == "AUX" || stem == "NUL")
                return true;
            if (stem.Length == 4 &&
                (stem.StartsWith("COM", StringComparison.Ordinal) ||
                 stem.StartsWith("LPT", StringComparison.Ordinal)) &&
                stem[3] >= '1' && stem[3] <= '9')
                return true;
            return false;
        }

        private static bool IsPathWithinRoot(string rootPath, string candidatePath)
        {
            try
            {
                var root = Path.GetFullPath(rootPath ?? string.Empty)
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                var candidate = Path.GetFullPath(candidatePath ?? string.Empty);
                var prefix = root + Path.DirectorySeparatorChar;
                return candidate.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static bool HasReparsePointBetween(string rootPath, string candidatePath)
        {
            try
            {
                var root = Path.GetFullPath(rootPath ?? string.Empty)
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                var candidate = Path.GetFullPath(candidatePath ?? string.Empty)
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                if (!IsPathWithinRoot(root, candidate)) return true;

                var relative = candidate.Substring(root.Length)
                    .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                var current = root;
                foreach (var segment in relative.Split(
                    new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
                    StringSplitOptions.RemoveEmptyEntries))
                {
                    current = Path.Combine(current, segment);
                    if (Directory.Exists(current) &&
                        (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    {
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                return true;
            }
        }

        private void InvalidateImageReplaceRibbonControls()
        {
            ribbon?.InvalidateControl("ApplyImageReplace");
            ribbon?.InvalidateControl("DdImageReplacePreset");
            ribbon?.InvalidateControl("TxtImageReplaceSourceColor");
            ribbon?.InvalidateControl("TxtImageReplaceNewColor");
        }

        private void EnsurePresetSelectionNames()
        {
            if (imageReplacePresets.Count == 0)
            {
                defaultPresetName = string.Empty;
                currentPresetName = string.Empty;
                return;
            }

            if (!imageReplacePresets.Any(x => string.Equals(x.Name, defaultPresetName, StringComparison.OrdinalIgnoreCase)))
            {
                defaultPresetName = imageReplacePresets[0].Name;
            }

            if (!imageReplacePresets.Any(x => string.Equals(x.Name, currentPresetName, StringComparison.OrdinalIgnoreCase)))
            {
                currentPresetName = defaultPresetName;
            }
        }

        private void UpsertPresetBySortOrder(string originalName, ImageReplacePreset editedPreset, int desiredSortOrder)
        {
            if (editedPreset == null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(originalName))
            {
                imageReplacePresets.RemoveAll(x => string.Equals(x.Name, originalName, StringComparison.OrdinalIgnoreCase));
            }
            imageReplacePresets.RemoveAll(x => string.Equals(x.Name, editedPreset.Name, StringComparison.OrdinalIgnoreCase));

            var ordered = PresetManager.GetPresetsInDisplayOrder(imageReplacePresets).ToList();
            var insertIndex = Math.Max(0, Math.Min(desiredSortOrder - 1, ordered.Count));
            ordered.Insert(insertIndex, editedPreset);
            for (int i = 0; i < ordered.Count; i++)
            {
                ordered[i].SortOrder = i + 1;
            }

            imageReplacePresets.Clear();
            imageReplacePresets.AddRange(ordered);
        }

        private void LoadImageReplacePresets()
        {
            imageReplacePresets.Clear();
            ResetImageReplaceColorOptions();
            defaultPresetName = string.Empty;
            currentPresetName = string.Empty;
            imageReplaceSourceColorInput = string.Empty;
            imageReplaceNewColorInput = string.Empty;
            imageReplaceFuzzInput = 5.0;

            if (!File.Exists(presetStorePath))
            {
                var defaultPreset = PresetManager.CreateDefaultPreset();
                imageReplacePresets.Add(defaultPreset);
                defaultPresetName = defaultPreset.Name;
                currentPresetName = defaultPresetName;
                PresetManager.NormalizePresetSortOrders(imageReplacePresets);
                EnsurePresetSelectionNames();
                SyncImageReplaceInputValuesFromCurrentPreset();
                return;
            }

            var doc = XDocument.Load(presetStorePath);
            var root = doc.Root;
            if (root == null)
            {
                return;
            }
            else
            {
                defaultPresetName = (string)root.Attribute("Default");
                currentPresetName = (string)root.Attribute("Current");
                materialLibraryPath = (string)root.Attribute("MaterialLibraryPath") ?? string.Empty;
                imageMagickPath = (string)root.Attribute("ImageMagickPath") ?? string.Empty;
                embedContextAddToLibraryEnabled = PresetManager.ParseBool((string)root.Attribute("EmbedContextAddToLibraryEnabled"));
                imageReplaceFuzzInput = PresetManager.ParseFuzz((string)root.Attribute("ImageReplaceFuzzInput"));
                materialPreviewCount = PresetManager.ParseMaterialPreviewCount((string)root.Attribute("MaterialPreviewCount"), MaterialPreviewButtonCount);
                hasPresetEditorBounds = PresetManager.TryParseEditorBounds(root, out presetEditorBounds);
                presetEditorSaveAsDefaultChecked = PresetManager.ParseBool((string)root.Attribute("EditorSaveAsDefault"));

                foreach (var sourceOption in root.Elements("SourceColorOption"))
                {
                    PresetManager.AddColorOption(imageReplaceSourceColorOptions, (string)sourceOption.Attribute("Value"));
                }

                foreach (var newOption in root.Elements("NewColorOption"))
                {
                    PresetManager.AddColorOption(imageReplaceNewColorOptions, (string)newOption.Attribute("Value"));
                }

                PresetManager.TrimLegacyDefaultColorOptions(imageReplaceSourceColorOptions);
                PresetManager.TrimLegacyDefaultColorOptions(imageReplaceNewColorOptions);

                foreach (var xPreset in root.Elements("Preset"))
                {
                    var preset = new ImageReplacePreset
                    {
                        Name = (string)xPreset.Attribute("Name") ?? "Default",
                        TargetColor = (string)xPreset.Attribute("TargetColor") ?? "white",
                        Mode = (string)xPreset.Attribute("Mode") ?? "transparent",
                        ReplacementColor = (string)xPreset.Attribute("ReplacementColor") ?? "black",
                        FuzzPercent = PresetManager.ParseFuzz((string)xPreset.Attribute("FuzzPercent")),
                        SortOrder = PresetManager.ParseSortOrder((string)xPreset.Attribute("SortOrder"), imageReplacePresets.Count + 1)
                    };
                    imageReplacePresets.Add(preset);
                }
            }

            if (imageReplacePresets.Count > 0 && string.IsNullOrWhiteSpace(defaultPresetName))
            {
                defaultPresetName = imageReplacePresets[0].Name;
            }
            if (imageReplacePresets.Count > 0 && string.IsNullOrWhiteSpace(currentPresetName))
            {
                currentPresetName = defaultPresetName;
            }
            if (imageReplacePresets.Count == 0)
            {
                var defaultPreset = PresetManager.CreateDefaultPreset();
                imageReplacePresets.Add(defaultPreset);
                defaultPresetName = defaultPreset.Name;
                currentPresetName = defaultPresetName;
            }

            PresetManager.NormalizePresetSortOrders(imageReplacePresets);
            EnsurePresetSelectionNames();
            imageReplaceFuzzInput = PresetManager.NormalizeFuzzPercent(imageReplaceFuzzInput);
            materialPreviewCount = PresetManager.ParseMaterialPreviewCount(materialPreviewCount.ToString(CultureInfo.InvariantCulture), MaterialPreviewButtonCount);
            EnsureImageReplaceColorOptions();
            SyncImageReplaceInputValuesFromCurrentPreset();
            // 始终默认启用“添加到 BioDraw 素材库”右键项，避免旧配置将其静默关闭。
            embedContextAddToLibraryEnabled = true;
        }

        private void SaveImageReplacePresets()
        {
            var dir = Path.GetDirectoryName(presetStorePath);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            PresetManager.NormalizePresetSortOrders(imageReplacePresets);
            EnsurePresetSelectionNames();
            imageReplaceFuzzInput = PresetManager.NormalizeFuzzPercent(imageReplaceFuzzInput);
            EnsureImageReplaceColorOptions();
            var root = new XElement(
                "Presets",
                new XAttribute("Default", defaultPresetName ?? string.Empty),
                new XAttribute("Current", currentPresetName ?? string.Empty),
                new XAttribute("MaterialLibraryPath", materialLibraryPath ?? string.Empty),
                new XAttribute("ImageMagickPath", imageMagickPath ?? string.Empty),
                new XAttribute("EmbedContextAddToLibraryEnabled", embedContextAddToLibraryEnabled),
                new XAttribute("MaterialPreviewCount", GetMaterialPageSize()),
                new XAttribute("ImageReplaceSourceInput", PresetManager.NormalizeColorInputText(imageReplaceSourceColorInput)),
                new XAttribute("ImageReplaceNewInput", PresetManager.NormalizeColorInputText(imageReplaceNewColorInput)),
                new XAttribute("ImageReplaceFuzzInput", imageReplaceFuzzInput.ToString("0.0", CultureInfo.InvariantCulture)),
                new XAttribute("EditorSaveAsDefault", presetEditorSaveAsDefaultChecked),
                imageReplaceSourceColorOptions.Select(x => new XElement(
                    "SourceColorOption",
                    new XAttribute("Value", x))),
                imageReplaceNewColorOptions.Select(x => new XElement(
                    "NewColorOption",
                    new XAttribute("Value", x))),
                imageReplacePresets.Select(p => new XElement(
                    "Preset",
                    new XAttribute("Name", p.Name),
                    new XAttribute("SortOrder", p.SortOrder),
                    new XAttribute("FuzzPercent", PresetManager.NormalizeFuzzPercent(p.FuzzPercent).ToString("0.0", CultureInfo.InvariantCulture)),
                    new XAttribute("TargetColor", p.TargetColor),
                    new XAttribute("Mode", p.Mode),
                    new XAttribute("ReplacementColor", p.ReplacementColor ?? "black"))));

            if (hasPresetEditorBounds)
            {
                root.SetAttributeValue("EditorX", presetEditorBounds.X);
                root.SetAttributeValue("EditorY", presetEditorBounds.Y);
                root.SetAttributeValue("EditorWidth", presetEditorBounds.Width);
                root.SetAttributeValue("EditorHeight", presetEditorBounds.Height);
            }

            var doc = new XDocument(root);
            doc.Save(presetStorePath);
        }

        private bool ShowPresetEditorDialog(ImageReplacePreset source, bool canDelete, out ImageReplacePreset result, out bool setAsDefault, out bool deleteRequested)
        {
            result = null;
            setAsDefault = false;
            deleteRequested = false;

            using (var dialog = new PresetEditorWebDialog(
                source,
                canDelete,
                Math.Max(1, imageReplacePresets.Count + 1),
                presetEditorSaveAsDefaultChecked))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return false;

                if (dialog.DeleteRequested)
                {
                    deleteRequested = true;
                    return true;
                }

                if (dialog.Result == null)
                    return false;

                result = dialog.Result;
                setAsDefault = dialog.SetAsDefault;
                presetEditorSaveAsDefaultChecked = dialog.SetAsDefault;
                return true;
            }
        }

        private static bool TryParseColorTokenToOleRgb(string colorToken, out int oleRgb)
        {
            oleRgb = 0;
            if (string.IsNullOrWhiteSpace(colorToken))
            {
                return false;
            }

            try
            {
                var color = ColorTranslator.FromHtml(colorToken.Trim());
                oleRgb = color.R + (color.G << 8) + (color.B << 16);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string ToHexColor(int oleRgb)
        {
            var r = oleRgb & 0xFF;
            var g = (oleRgb >> 8) & 0xFF;
            var b = (oleRgb >> 16) & 0xFF;
            return "#" + r.ToString("X2", CultureInfo.InvariantCulture) +
                g.ToString("X2", CultureInfo.InvariantCulture) +
                b.ToString("X2", CultureInfo.InvariantCulture);
        }

        private bool TryPickColorWithPowerPoint(bool useMoreColorsDialog, string initialColor, out string colorToken, out string errorMessage)
        {
            colorToken = string.Empty;
            errorMessage = string.Empty;

            if (!useMoreColorsDialog)
            {
                return TryPickColorFromSelectedPictureSource(initialColor, out colorToken, out errorMessage);
            }

            dynamic tempShape = null;
            var previousShapes = new List<dynamic>();

            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null)
                {
                    errorMessage = "未能获取 PowerPoint 应用实例。";
                    return false;
                }

                dynamic selection = null;
                try
                {
                    selection = app.ActiveWindow?.Selection;
                }
                catch
                {
                }

                List<dynamic> shapes = new List<dynamic>();
                if (selection != null && ImageReplacePipeline.TryGetSelectedShapes(selection, out shapes))
                {
                    previousShapes.AddRange(shapes);
                }

                dynamic slide = null;
                try
                {
                    slide = app.ActiveWindow?.View?.Slide;
                }
                catch
                {
                }

                if (slide == null)
                {
                    errorMessage = "请先切换到普通编辑视图后再取色。";
                    return false;
                }

                tempShape = slide.Shapes.AddShape(1, -200, -200, 10, 10);
                tempShape.Fill.Visible = -1;
                tempShape.Line.Visible = 0;
                tempShape.Fill.Solid();

                int initialOleRgb;
                if (TryParseColorTokenToOleRgb(initialColor, out initialOleRgb))
                {
                    tempShape.Fill.ForeColor.RGB = initialOleRgb;
                }

                tempShape.Select();

                var commandIds = useMoreColorsDialog
                    ? new[] { "ShapeFillColorMoreColorsDialog", "ShapeFillMoreColorsDialog", "ObjectFillMoreColorsDialog" }
                    : new[] { "ShapeFillColorPicker", "ObjectFillColorPicker", "TextFillColorPicker" };
                if (!TryExecuteMso(app, commandIds))
                {
                    int fallbackOleRgb = (int)tempShape.Fill.ForeColor.RGB;
                    string fallbackColor;
                    if (TryPickColorWithSystemDialog(fallbackOleRgb, useMoreColorsDialog, out fallbackColor))
                    {
                        colorToken = fallbackColor;
                        return true;
                    }

                    errorMessage = string.Empty;
                    return false;
                }

                colorToken = ToHexColor((int)tempShape.Fill.ForeColor.RGB);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
            finally
            {
                try
                {
                    if (tempShape != null)
                    {
                        tempShape.Delete();
                    }
                }
                catch
                {
                }

                ImageReplacePipeline.TryReselectShapes(previousShapes);
            }
        }

        private bool TryPickColorFromSelectedPictureSource(string initialColor, out string colorToken, out string errorMessage)
        {
            colorToken = string.Empty;
            errorMessage = string.Empty;
            var application = Globals.ThisAddIn?.Application;
            if (application == null)
            {
                errorMessage = "未能获取 PowerPoint 应用实例。";
                return false;
            }

            return TryPickColorFromPptSurface(application, initialColor, out colorToken, out errorMessage);
        }

        private static bool TryPickColorFromPptSurface(dynamic application, string initialColor, out string colorToken, out string errorMessage)
        {
            colorToken = string.Empty;
            errorMessage = string.Empty;
            if (application == null)
            {
                errorMessage = "未能初始化取色上下文。";
                return false;
            }

            using (var preview = new Form())
            using (var swatch = new PictureBox())
            using (var code = new Label())
            {
                preview.FormBorderStyle = FormBorderStyle.None;
                preview.ShowInTaskbar = false;
                preview.StartPosition = FormStartPosition.Manual;
                preview.TopMost = true;
                preview.BackColor = Color.FromArgb(35, 35, 35);
                preview.Opacity = 0.95;
                preview.AutoScaleMode = AutoScaleMode.Dpi;
                preview.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);

                var swatchSize = 18;
                var radius = 7;
                var paddingX = 8;
                var gap = 6;

                swatch.Size = new Size(swatchSize, swatchSize);
                swatch.BackColor = Color.Transparent;
                swatch.SizeMode = PictureBoxSizeMode.Normal;

                code.AutoSize = false;
                code.ForeColor = Color.White;
                code.Text = "#000000";
                code.TextAlign = ContentAlignment.MiddleCenter;

                var measured = TextRenderer.MeasureText(code.Text, preview.Font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding);
                var codeWidth = measured.Width + 4;
                var contentWidth = swatchSize + gap + codeWidth;
                var previewWidth = contentWidth + paddingX * 2;
                var previewHeight = Math.Max((int)Math.Ceiling(swatchSize * 1.1), measured.Height + 6);
                preview.Size = new Size(previewWidth, previewHeight);

                var centerY = previewHeight / 2;
                swatch.Location = new Point(paddingX, centerY - swatchSize / 2);
                code.Size = new Size(codeWidth, measured.Height + 2);
                code.Location = new Point(swatch.Right + gap, centerY - code.Height / 2);

                ApplyRoundedRegion(preview, radius);

                preview.Controls.Add(swatch);
                preview.Controls.Add(code);
                preview.Show();

                Cursor pickerCursor;
                bool ownsPickerCursor;
                TryCreatePickerCursor(out pickerCursor, out ownsPickerCursor);
                preview.Cursor = pickerCursor;

                try
                {
                    Color currentColor = Color.Black;
                    int initialOleRgb;
                    if (TryParseColorTokenToOleRgb(initialColor, out initialOleRgb))
                    {
                        currentColor = Color.FromArgb(
                            initialOleRgb & 0xFF,
                            (initialOleRgb >> 8) & 0xFF,
                            (initialOleRgb >> 16) & 0xFF);
                    }

                    bool hasHoverColor = false;
                    bool leftPressed = false;
                    while (true)
                    {
                        Application.DoEvents();
                        Thread.Sleep(10);
                        Cursor.Current = pickerCursor;

                        var cursor = Cursor.Position;
                        preview.Location = new Point(cursor.X + 18, Math.Max(0, cursor.Y - preview.Height - 10));

                        var hoverColor = Color.Empty;
                        if (TryGetScreenPixelColor(cursor, out hoverColor))
                        {
                            currentColor = hoverColor;
                            hasHoverColor = true;
                        }
                        else
                        {
                            hasHoverColor = false;
                        }

                        var oldImage = swatch.Image;
                        swatch.Image = CreateSwatchCircleImage(currentColor, swatchSize);
                        if (oldImage != null)
                        {
                            oldImage.Dispose();
                        }
                        code.Text = "#" + currentColor.R.ToString("X2", CultureInfo.InvariantCulture) +
                            currentColor.G.ToString("X2", CultureInfo.InvariantCulture) +
                            currentColor.B.ToString("X2", CultureInfo.InvariantCulture);

                        if (IsVirtualKeyDown(0x1B))
                        {
                            errorMessage = string.Empty;
                            return false;
                        }

                        var down = IsVirtualKeyDown(0x01);
                        if (down && !leftPressed)
                        {
                            leftPressed = true;
                        }

                        if (!down && leftPressed)
                        {
                            if (!hasHoverColor)
                            {
                                leftPressed = false;
                                continue;
                            }

                            colorToken = code.Text;
                            return true;
                        }
                    }
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                    var image = swatch.Image;
                    swatch.Image = null;
                    if (image != null)
                    {
                        image.Dispose();
                    }

                    if (ownsPickerCursor && pickerCursor != null)
                    {
                        pickerCursor.Dispose();
                    }
                }
            }
        }

        private static void TryCreatePickerCursor(out Cursor cursor, out bool ownsCursor)
        {
            cursor = Cursors.Cross;
            ownsCursor = false;
            try
            {
                var customCursorPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "BioDraw",
                    "eyedropper.cur");
                if (!File.Exists(customCursorPath))
                {
                    cursor = CreateGeneratedPickerCursor();
                    ownsCursor = cursor != null;
                    if (!ownsCursor)
                    {
                        cursor = Cursors.Cross;
                    }
                    return;
                }

                cursor = new Cursor(customCursorPath);
                ownsCursor = true;
            }
            catch
            {
                cursor = Cursors.Cross;
                ownsCursor = false;
            }
        }

        private static Cursor CreateGeneratedPickerCursor()
        {
            IntPtr sourceIcon = IntPtr.Zero;
            IntPtr cursorHandle = IntPtr.Zero;
            IntPtr colorBitmap = IntPtr.Zero;
            IntPtr maskBitmap = IntPtr.Zero;
            try
            {
                using (var bitmap = new Bitmap(32, 32))
                using (var graphics = Graphics.FromImage(bitmap))
                using (var bodyBrush = new SolidBrush(Color.FromArgb(245, 245, 245)))
                using (var metalBrush = new SolidBrush(Color.FromArgb(170, 170, 170)))
                using (var outlinePen = new Pen(Color.FromArgb(40, 40, 40), 2f))
                using (var detailPen = new Pen(Color.FromArgb(85, 85, 85), 1.2f))
                {
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    graphics.Clear(Color.Transparent);
                    graphics.TranslateTransform(8f, 2f);
                    graphics.RotateTransform(40f);
                    graphics.FillRectangle(bodyBrush, 2f, 2f, 7f, 16f);
                    graphics.FillRectangle(metalBrush, 3f, 0f, 5f, 3f);
                    graphics.FillRectangle(bodyBrush, 2f, 18f, 7f, 7f);
                    graphics.DrawRectangle(outlinePen, 2f, 2f, 7f, 23f);
                    graphics.DrawLine(detailPen, 2f, 16f, 9f, 16f);
                    graphics.FillEllipse(Brushes.White, 3f, 20f, 5f, 5f);
                    graphics.ResetTransform();
                    graphics.FillEllipse(Brushes.Black, 3f, 26f, 3f, 3f);
                    graphics.FillEllipse(Brushes.White, 2f, 25f, 5f, 5f);
                    graphics.FillEllipse(Brushes.Black, 3f, 26f, 3f, 3f);
                    sourceIcon = bitmap.GetHicon();
                }

                ICONINFO iconInfo;
                if (!GetIconInfo(sourceIcon, out iconInfo))
                {
                    return null;
                }

                colorBitmap = iconInfo.hbmColor;
                maskBitmap = iconInfo.hbmMask;
                iconInfo.fIcon = false;
                iconInfo.xHotspot = 4;
                iconInfo.yHotspot = 28;
                cursorHandle = CreateIconIndirect(ref iconInfo);
                if (cursorHandle == IntPtr.Zero)
                {
                    return null;
                }

                return new Cursor(cursorHandle);
            }
            catch
            {
                if (cursorHandle != IntPtr.Zero)
                {
                    DestroyIcon(cursorHandle);
                }
                return null;
            }
            finally
            {
                if (colorBitmap != IntPtr.Zero)
                {
                    DeleteObject(colorBitmap);
                }
                if (maskBitmap != IntPtr.Zero)
                {
                    DeleteObject(maskBitmap);
                }
                if (sourceIcon != IntPtr.Zero)
                {
                    DestroyIcon(sourceIcon);
                }
            }
        }

        private static Bitmap CreateSwatchCircleImage(Color color, int diameter)
        {
            var safeDiameter = Math.Max(10, diameter);
            var bitmap = new Bitmap(safeDiameter, safeDiameter);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphics.Clear(Color.Transparent);
                using (var brush = new SolidBrush(color))
                {
                    graphics.FillEllipse(brush, 1, 1, safeDiameter - 2, safeDiameter - 2);
                }

                using (var pen = new Pen(Color.FromArgb(210, 255, 255, 255), 1f))
                {
                    graphics.DrawEllipse(pen, 1, 1, safeDiameter - 3, safeDiameter - 3);
                }
            }
            return bitmap;
        }

        private static bool TryGetScreenPixelColor(Point point, out Color color)
        {
            color = Color.Empty;
            IntPtr hdc = IntPtr.Zero;
            try
            {
                hdc = GetDC(IntPtr.Zero);
                if (hdc == IntPtr.Zero)
                {
                    return false;
                }

                var rgb = GetPixel(hdc, point.X, point.Y);
                if (rgb == 0xFFFFFFFF)
                {
                    return false;
                }

                var r = (int)(rgb & 0xFF);
                var g = (int)((rgb >> 8) & 0xFF);
                var b = (int)((rgb >> 16) & 0xFF);
                color = Color.FromArgb(r, g, b);
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (hdc != IntPtr.Zero)
                {
                    ReleaseDC(IntPtr.Zero, hdc);
                }
            }
        }

        private static bool TryIsPointInsidePptWindow(dynamic application, Point cursor)
        {
            try
            {
                if (application == null)
                {
                    return false;
                }

                var hwnd = (IntPtr)(int)application.HWND;
                if (hwnd == IntPtr.Zero)
                {
                    return false;
                }

                var hit = WindowFromPoint(new POINT { X = cursor.X, Y = cursor.Y });
                if (hit == IntPtr.Zero)
                {
                    return false;
                }

                uint pptPid;
                GetWindowThreadProcessId(hwnd, out pptPid);
                uint hitPid;
                GetWindowThreadProcessId(hit, out hitPid);
                if (pptPid != 0 && hitPid == pptPid)
                {
                    return true;
                }

                var hitRoot = GetAncestor(hit, 2);
                var pptRoot = GetAncestor(hwnd, 2);
                if (hitRoot != IntPtr.Zero && pptRoot != IntPtr.Zero && hitRoot == pptRoot)
                {
                    return true;
                }

                RECT rect;
                if (!GetWindowRect(hwnd, out rect))
                {
                    return false;
                }
                return cursor.X >= rect.Left && cursor.X <= rect.Right &&
                    cursor.Y >= rect.Top && cursor.Y <= rect.Bottom;
            }
            catch
            {
                return false;
            }
        }

        private static void ApplyRoundedRegion(Control control, int radius)
        {
            if (control == null || control.Width <= 0 || control.Height <= 0 || radius <= 0)
            {
                return;
            }

            var diameter = radius * 2;
            var rect = new Rectangle(0, 0, control.Width, control.Height);
            using (var path = new GraphicsPath())
            {
                path.AddArc(rect.Left, rect.Top, diameter, diameter, 180, 90);
                path.AddArc(rect.Right - diameter, rect.Top, diameter, diameter, 270, 90);
                path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(rect.Left, rect.Bottom - diameter, diameter, diameter, 90, 90);
                path.CloseFigure();
                control.Region = new Region(path);
            }
        }

        internal static bool TryExecuteMso(dynamic app, IEnumerable<string> commandIds)
        {
            if (app == null || commandIds == null)
            {
                return false;
            }

            foreach (var commandId in commandIds)
            {
                if (string.IsNullOrWhiteSpace(commandId))
                {
                    continue;
                }

                try
                {
                    app.CommandBars.ExecuteMso(commandId);
                    return true;
                }
                catch
                {
                }
            }

            return false;
        }

        private static bool TryPickColorWithSystemDialog(int initialOleRgb, bool fullOpen, out string colorToken)
        {
            colorToken = string.Empty;
            using (var dialog = new ColorDialog())
            {
                var red = initialOleRgb & 0xFF;
                var green = (initialOleRgb >> 8) & 0xFF;
                var blue = (initialOleRgb >> 16) & 0xFF;
                dialog.Color = Color.FromArgb(red, green, blue);
                dialog.FullOpen = fullOpen;
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return false;
                }

                colorToken = "#" + dialog.Color.R.ToString("X2", CultureInfo.InvariantCulture) +
                    dialog.Color.G.ToString("X2", CultureInfo.InvariantCulture) +
                    dialog.Color.B.ToString("X2", CultureInfo.InvariantCulture);
                return true;
            }
        }

        private List<string> GetLevel1List()
        {
            if (!UseCustomMaterialLibrary())
            {
                return level1Items.Count > 0 ? level1Items : new List<string> { "默认" };
            }

            return MaterialLibraryService.GetSubDirectoryNames(materialLibraryPath);
        }

        private List<string> GetLevel2List()
        {
            if (UseCustomMaterialLibrary())
            {
                var level1List = GetLevel1List();
                if (level1List.Count == 0)
                {
                    return new List<string> { "默认" };
                }

                var level1 = level1List[MaterialLibraryService.NormalizeIndex(selectedLevel1Index, level1List.Count)];
                var level1Path = Path.Combine(materialLibraryPath, level1);
                return MaterialLibraryService.GetSubDirectoryNames(level1Path);
            }

            var fallbackLevel1List = GetLevel1List();
            var fallbackLevel1 = fallbackLevel1List[MaterialLibraryService.NormalizeIndex(selectedLevel1Index, fallbackLevel1List.Count)];
            List<string> list;
            if (level2Items.TryGetValue(fallbackLevel1, out list) && list.Count > 0)
            {
                return list;
            }
            return new List<string> { "默认" };
        }

        private List<string> GetLevel3List()
        {
            if (UseCustomMaterialLibrary())
            {
                var level1List = GetLevel1List();
                var level2List = GetLevel2List();
                if (level1List.Count == 0 || level2List.Count == 0)
                {
                    return new List<string> { "默认" };
                }

                var level1 = level1List[MaterialLibraryService.NormalizeIndex(selectedLevel1Index, level1List.Count)];
                var level2 = level2List[MaterialLibraryService.NormalizeIndex(selectedLevel2Index, level2List.Count)];
                var level2Path = Path.Combine(materialLibraryPath, level1, level2);
                return MaterialLibraryService.GetSubDirectoryNames(level2Path);
            }

            var fallbackLevel2List = GetLevel2List();
            var fallbackLevel2 = fallbackLevel2List[MaterialLibraryService.NormalizeIndex(selectedLevel2Index, fallbackLevel2List.Count)];
            List<string> list;
            if (level3Items.TryGetValue(fallbackLevel2, out list) && list.Count > 0)
            {
                return list;
            }
            return new List<string> { "默认" };
        }

        private List<MaterialEntry> GetMaterialEntries()
        {
            if (!UseCustomMaterialLibrary())
            {
                return GetLevel3List()
                    .Select(x => new MaterialEntry { Name = x, FilePath = string.Empty })
                    .ToList();
            }

            if (!string.IsNullOrWhiteSpace(materialSearchText))
            {
                return SearchMaterialEntries(materialSearchText);
            }

            var level1List = GetLevel1List();
            var level2List = GetLevel2List();
            if (level1List.Count == 0 || level2List.Count == 0)
            {
                return new List<MaterialEntry> { new MaterialEntry { Name = "默认", FilePath = string.Empty } };
            }

            var level1 = level1List[MaterialLibraryService.NormalizeIndex(selectedLevel1Index, level1List.Count)];
            var level1Path = Path.Combine(materialLibraryPath, level1);
            if (!Directory.Exists(level1Path))
            {
                return new List<MaterialEntry> { new MaterialEntry { Name = "默认", FilePath = string.Empty } };
            }

            // 兼容“一级目录直接放素材文件”的结构。
            var hasSubDirectories = false;
            try
            {
                hasSubDirectories = Directory.EnumerateDirectories(level1Path).Any();
            }
            catch
            {
                hasSubDirectories = false;
            }

            if (!hasSubDirectories)
            {
                return MaterialLibraryService.GetMaterialEntriesFromFolder(level1Path);
            }

            var level2 = level2List[MaterialLibraryService.NormalizeIndex(selectedLevel2Index, level2List.Count)];
            var level2Path = Path.Combine(level1Path, level2);
            return MaterialLibraryService.GetMaterialEntriesFromFolder(level2Path);
        }

        private List<MaterialEntry> SearchMaterialEntries(string keyword)
        {
            if (string.IsNullOrWhiteSpace(materialLibraryPath) || !Directory.Exists(materialLibraryPath))
            {
                return new List<MaterialEntry> { new MaterialEntry { Name = "默认", FilePath = string.Empty } };
            }

            var searchKey = (keyword ?? string.Empty).Trim();
            if (searchKey.Length == 0)
            {
                return new List<MaterialEntry> { new MaterialEntry { Name = "默认", FilePath = string.Empty } };
            }

            try
            {
                var allEntries = GetMaterialSearchEntriesCache();
                var entries = allEntries
                    .Where(x => x.Name.IndexOf(searchKey, StringComparison.OrdinalIgnoreCase) >= 0)
                    .OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(x => x.FilePath, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (entries.Count > 0)
                {
                    return entries;
                }
            }
            catch
            {
            }

            return new List<MaterialEntry> { new MaterialEntry { Name = "默认", FilePath = string.Empty } };
        }

        private List<MaterialEntry> GetMaterialSearchEntriesCache()
        {
            if (string.IsNullOrWhiteSpace(materialLibraryPath) || !Directory.Exists(materialLibraryPath))
            {
                materialSearchCacheRootPath = null;
                materialSearchCacheEntries = null;
                return new List<MaterialEntry>();
            }

            if (materialSearchCacheEntries != null &&
                string.Equals(materialSearchCacheRootPath, materialLibraryPath, StringComparison.OrdinalIgnoreCase))
            {
                return materialSearchCacheEntries;
            }

            try
            {
                materialSearchCacheEntries = Directory.GetFiles(materialLibraryPath, "*", SearchOption.AllDirectories)
                    .Where(MaterialLibraryService.IsSupportedMaterialFile)
                    .Select(path => new MaterialEntry
                    {
                        Name = Path.GetFileNameWithoutExtension(path),
                        FilePath = path
                    })
                    .ToList();
                materialSearchCacheRootPath = materialLibraryPath;
                return materialSearchCacheEntries;
            }
            catch
            {
                materialSearchCacheRootPath = materialLibraryPath;
                materialSearchCacheEntries = new List<MaterialEntry>();
                return materialSearchCacheEntries;
            }
        }

        private stdole.IPictureDisp GetMaterialPreviewImage(MaterialEntry entry)
        {
            EnsureBrandImages();
            if (entry == null)
            {
                return transparentPlaceholderImage ?? brandImageLarge ?? brandImageSmall;
            }

            var filePath = entry.FilePath;
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return transparentPlaceholderImage ?? brandImageLarge ?? brandImageSmall;
            }

            stdole.IPictureDisp picture;
            if (materialPreviewCache.TryGetValue(filePath, out picture) && picture != null)
            {
                return picture;
            }

            Bitmap bitmap;
            if (!MaterialLibraryService.TryBuildMaterialThumbnail(filePath, entry.Name, MaterialThumbnailWidth, MaterialThumbnailHeight, out bitmap))
            {
                return transparentPlaceholderImage ?? brandImageLarge ?? brandImageSmall;
            }

            using (bitmap)
            {
                picture = PictureConverter.ToPictureDisp(new Bitmap(bitmap));
            }

            if (materialPreviewCache.Count >= MaterialPreviewCacheLimit)
            {
                try
                {
                    var keysToRemove = materialPreviewCache.Keys.Take(materialPreviewCache.Count / 2).ToList();
                    foreach (var key in keysToRemove)
                    {
                        materialPreviewCache.Remove(key);
                    }
                }
                catch
                {
                }
            }

            materialPreviewCache[filePath] = picture;
            return picture;
        }

        private bool UseCustomMaterialLibrary()
        {
            return !string.IsNullOrWhiteSpace(materialLibraryPath) && Directory.Exists(materialLibraryPath);
        }

        private static stdole.IPictureDisp CreateAiIconImage()
        {
            const int size = 32;
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                var fillColor = Color.FromArgb(0, 122, 204);
                var accentColor = Color.FromArgb(220, 240, 255);

                using (var outerBrush = new SolidBrush(fillColor))
                using (var innerBrush = new SolidBrush(accentColor))
                using (var pen = new Pen(fillColor, 1.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                {
                    // Central circle — a "brain/neuron" node
                    g.FillEllipse(outerBrush, 10, 10, 12, 12);

                    // Three sparkle rays
                    g.DrawLine(pen, 16, 4, 16, 8);
                    g.DrawLine(pen, 28, 10, 25, 13);
                    g.DrawLine(pen, 4, 22, 7, 20);

                    // Small outer nodes
                    g.FillEllipse(innerBrush, 14, 2, 4, 4);
                    g.FillEllipse(innerBrush, 27, 8, 4, 4);
                    g.FillEllipse(innerBrush, 2, 21, 4, 4);

                    // Connecting arcs to outer nodes
                    g.DrawLine(pen, 16, 6, 16, 10);
                    g.DrawLine(pen, 24, 12, 27, 10);
                    g.DrawLine(pen, 8, 21, 10, 20);
                }
            }

            return PictureConverter.ToPictureDisp(bmp);
        }

    }
}
