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
        private const string PickerColorFileName = "picker-color.png";
        private const string LegacyPickerColorFileName = "Picker_color .png";
        private const string BrandPngFileName = "BioDraw.png";
        private const string SettingsGearPngFileName = "settings-gear.png";
        private const string ImageRecolorPngFileName = "image-recolor.png";
        private const string RibbonEmptyInputToken = "\u2060";
        private string materialSearchText;
        private double imageReplaceFuzzInput;
        private stdole.IPictureDisp brandImageLarge;
        private stdole.IPictureDisp brandImageSmall;
        private stdole.IPictureDisp transparentPlaceholderImage;
        private stdole.IPictureDisp pickerButtonImage;
        private stdole.IPictureDisp settingsButtonImage;
        private stdole.IPictureDisp imageRecolorButtonImage;
        private stdole.IPictureDisp pageUpButtonImage;
        private stdole.IPictureDisp pageDownButtonImage;
        private readonly List<ImageReplacePreset> imageReplacePresets;
        private readonly string presetStorePath;
        private const string projectAddressUrl = "https://github.com/CaptainMusX/BioDraw";
        private string currentPresetName;
        private string defaultPresetName;
        private string materialLibraryPath;
        private string materialSearchCacheRootPath;
        private List<MaterialEntry> materialSearchCacheEntries;
        private string imageReplaceSourceColorInput;
        private string imageReplaceNewColorInput;
        private readonly List<string> imageReplaceSourceColorOptions;
        private readonly List<string> imageReplaceNewColorOptions;
        private bool presetEditorSaveAsDefaultChecked;
        private Rectangle presetEditorBounds;
        private bool hasPresetEditorBounds;
        private static readonly HashSet<string> materialFileExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg",
            ".jpeg",
            ".png",
            ".bmp",
            ".gif",
            ".tif",
            ".tiff",
            ".webp",
            ".svg",
            ".emf",
            ".wmf"
        };

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
            materialLibraryPath = string.Empty;
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
            return NormalizeIndex(selectedLevel1Index, list.Count);
        }

        public string GetLevel1Text(Office.IRibbonControl control)
        {
            var list = GetLevel1List();
            if (list.Count == 0)
            {
                return string.Empty;
            }
            var index = NormalizeIndex(selectedLevel1Index, list.Count);
            return list[index];
        }

        public void OnLevel1Changed(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var list = GetLevel1List();
            selectedLevel1Index = NormalizeIndex(selectedIndex, list.Count);
            selectedLevel2Index = 0;
            materialPageIndex = 0;
            ribbon?.InvalidateControl("DdLevel2");
            InvalidateMaterialPreview();
        }

        public void OnLevel1TextChanged(Office.IRibbonControl control, string text)
        {
            var list = GetLevel1List();
            var index = FindExactMatchIndex(list, text);
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
            return NormalizeIndex(selectedLevel2Index, list.Count);
        }

        public string GetLevel2Text(Office.IRibbonControl control)
        {
            var list = GetLevel2List();
            if (list.Count == 0)
            {
                return string.Empty;
            }
            var index = NormalizeIndex(selectedLevel2Index, list.Count);
            return list[index];
        }

        public void OnLevel2Changed(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var list = GetLevel2List();
            selectedLevel2Index = NormalizeIndex(selectedIndex, list.Count);
            materialPageIndex = 0;
            InvalidateMaterialPreview();
        }

        public void OnLevel2TextChanged(Office.IRibbonControl control, string text)
        {
            var list = GetLevel2List();
            var index = FindExactMatchIndex(list, text);
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
            return NormalizeIndex(selectedLevel3Index, list.Count);
        }

        public void OnLevel3Changed(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var list = GetLevel3List();
            selectedLevel3Index = NormalizeIndex(selectedIndex, list.Count);
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
            return ToFixedLengthMaterialLabel(item == null ? string.Empty : item.Name);
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



        private static int GetDisplayLength(string text)
        {
            int len = 0;
            foreach (char c in text)
            {
                len += c > 255 ? 2 : 1;
            }
            return len;
        }

        private static string ToFixedLengthMaterialLabel(string text)
        {
            const int maxVisibleLength = 16;
            const int totalLength = 20;
            const char padChar = '\u00A0';

            var normalized = (text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(normalized))
            {
                return new string(padChar, totalLength);
            }

            string label = "";
            int currentLen = 0;
            bool truncated = false;

            foreach (char c in normalized)
            {
                int charLen = c > 255 ? 2 : 1;
                if (currentLen + charLen > maxVisibleLength)
                {
                    truncated = true;
                    break;
                }
                label += c;
                currentLen += charLen;
            }

            if (truncated)
            {
                label += "…";
                currentLen += 2; // '…' is usually full-width
            }

            int padCount = totalLength - currentLen;
            if (padCount > 0)
            {
                label += new string(padChar, padCount);
            }

            return label;
        }

        private static string ToEllipsisLabel(string text, int maxWidthPixels, int maxLines)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            if (maxWidthPixels <= 8)
            {
                return "…";
            }

            var sourceText = text.Trim();
            using (var font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point))
            {
                const TextFormatFlags flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine;
                if (MeasureTextWidth(sourceText, font, flags) <= maxWidthPixels)
                {
                    return sourceText;
                }

                var safeLines = Math.Max(1, maxLines);
                if (safeLines >= 2)
                {
                    var firstLineLength = GetLineBreakLength(sourceText, maxWidthPixels, font, flags);
                    if (firstLineLength > 0)
                    {
                        var firstLine = sourceText.Substring(0, firstLineLength).TrimEnd();
                        var remaining = sourceText.Substring(firstLineLength).TrimStart();
                        if (!string.IsNullOrWhiteSpace(remaining))
                        {
                            var secondLine = BuildEllipsisLine(remaining, maxWidthPixels, font, flags);
                            return string.Concat(firstLine, "\n", secondLine);
                        }
                    }
                }

                return BuildEllipsisLine(sourceText, maxWidthPixels, font, flags);
            }
        }

        private static int MeasureTextWidth(string text, Font font, TextFormatFlags flags)
        {
            return TextRenderer.MeasureText(text, font, new Size(int.MaxValue, int.MaxValue), flags).Width;
        }

        private static int GetLineBreakLength(string text, int maxWidthPixels, Font font, TextFormatFlags flags)
        {
            var maxLength = GetMaxFittingLength(text, maxWidthPixels, font, flags);
            if (maxLength <= 0)
            {
                return 0;
            }

            var breakLength = maxLength;
            for (int index = maxLength - 1; index >= 1; index--)
            {
                if (char.IsWhiteSpace(text[index]))
                {
                    breakLength = index;
                    break;
                }
            }

            while (breakLength > 0 && char.IsWhiteSpace(text[breakLength - 1]))
            {
                breakLength--;
            }

            return breakLength > 0 ? breakLength : maxLength;
        }

        private static int GetMaxFittingLength(string text, int maxWidthPixels, Font font, TextFormatFlags flags)
        {
            var low = 1;
            var high = text.Length;
            var best = 0;
            while (low <= high)
            {
                var mid = low + ((high - low) / 2);
                var candidate = text.Substring(0, mid);
                if (MeasureTextWidth(candidate, font, flags) <= maxWidthPixels)
                {
                    best = mid;
                    low = mid + 1;
                }
                else
                {
                    high = mid - 1;
                }
            }

            return best;
        }

        private static string BuildEllipsisLine(string text, int maxWidthPixels, Font font, TextFormatFlags flags)
        {
            if (MeasureTextWidth(text, font, flags) <= maxWidthPixels)
            {
                return text;
            }

            for (int length = text.Length - 1; length >= 1; length--)
            {
                var candidate = text.Substring(0, length).TrimEnd() + "…";
                if (MeasureTextWidth(candidate, font, flags) <= maxWidthPixels)
                {
                    return candidate;
                }
            }

            return "…";
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
            InsertMaterial(item);
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
                SetStatusText("BioDraw：当前素材仅为占位项。");
                return;
            }

            string error;
            if (!TryInsertMaterialToCurrentSlide(item.FilePath, out error))
            {
                if (!string.IsNullOrWhiteSpace(error))
                {
                    MessageBox.Show("插入素材失败：" + error, "BioDraw");
                }
                return;
            }

            SetStatusText("BioDraw：已插入素材 - " + item.Name);
        }

        public void OnAbout(Office.IRibbonControl control)
        {
            MessageBox.Show(
                "由 CaptainMus 开发的一款用于科研绘图的 PowerPoint 插件，欢迎使用！",
                "关于 BioDraw",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
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
                SetStatusText("BioDraw：素材库路径已更新。");
            }
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
            return ToRibbonColorInputText(imageReplaceSourceColorInput);
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
            var normalized = NormalizeColorInputText(imageReplaceSourceColorInput);
            if (string.IsNullOrEmpty(normalized))
            {
                return 0;
            }
            int idx = FindColorOptionIndex(imageReplaceSourceColorOptions, normalized);
            return idx >= 0 ? idx + 1 : -1;
        }

        public string GetImageReplaceNewColorText(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            return ToRibbonColorInputText(imageReplaceNewColorInput);
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
            var normalized = NormalizeColorInputText(imageReplaceNewColorInput);
            if (string.IsNullOrEmpty(normalized))
            {
                return 0;
            }
            int idx = FindColorOptionIndex(imageReplaceNewColorOptions, normalized);
            return idx >= 0 ? idx + 1 : -1;
        }

        public void OnImageReplaceSourceColorChanged(Office.IRibbonControl control, string text)
        {
            imageReplaceSourceColorInput = ToStorageColorInputText(text);
            PersistImageReplaceInputMemory();
            ribbon?.InvalidateControl("TxtImageReplaceSourceColor");
        }

        public void OnImageReplaceNewColorChanged(Office.IRibbonControl control, string text)
        {
            imageReplaceNewColorInput = ToStorageColorInputText(text);
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

            imageReplaceSourceColorInput = ToStorageColorInputText(selectedColor);
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

            imageReplaceNewColorInput = ToStorageColorInputText(selectedColor);
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

            imageReplaceSourceColorInput = ToStorageColorInputText(colorToken);
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

            imageReplaceNewColorInput = ToStorageColorInputText(colorToken);
            PersistImageReplaceInputMemory();
            ribbon?.InvalidateControl("TxtImageReplaceNewColor");
        }

        public int GetImageReplacePresetItemCount(Office.IRibbonControl control)
        {
            return imageReplacePresets.Count;
        }

        public string GetImageReplacePresetItemLabel(Office.IRibbonControl control, int index)
        {
            var ordered = GetPresetsInDisplayOrder().ToList();
            if (index >= 0 && index < ordered.Count)
            {
                return ordered[index].Name;
            }
            return string.Empty;
        }

        public int GetImageReplacePresetSelectedIndex(Office.IRibbonControl control)
        {
            var ordered = GetPresetsInDisplayOrder().ToList();
            if (ordered.Count == 0)
            {
                return -1;
            }
            var currentIndex = ordered.FindIndex(x => string.Equals(x.Name, currentPresetName, StringComparison.OrdinalIgnoreCase));
            return currentIndex >= 0 ? currentIndex : 0;
        }

        public string GetImageReplacePresetText(Office.IRibbonControl control)
        {
            var ordered = GetPresetsInDisplayOrder().ToList();
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
            var ordered = GetPresetsInDisplayOrder().ToList();
            if (selectedIndex >= 0 && selectedIndex < ordered.Count)
            {
                currentPresetName = ordered[selectedIndex].Name;
                SyncImageReplaceInputValuesFromCurrentPreset();
                InvalidateImageReplaceRibbonControls();
            }
        }

        public void OnImageReplacePresetTextChanged(Office.IRibbonControl control, string text)
        {
            var ordered = GetPresetsInDisplayOrder().ToList();
            var names = ordered.Select(x => x.Name).ToList();
            var index = FindExactMatchIndex(names, text);
            if (index < 0)
            {
                return;
            }
            OnImageReplacePresetDropDownChanged(control, string.Empty, index);
        }

        public void OnApplyImageReplace(Office.IRibbonControl control)
        {
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

            List<dynamic> shapes;
            if (!TryGetSelectedShapes(selection, out shapes))
            {
                MessageBox.Show("请先选中一张或多张图片。", "BioDraw");
                return;
            }

            EnsureImageReplaceInputValues();
            var sourceColor = NormalizeColorInputText(imageReplaceSourceColorInput);
            var newColor = NormalizeColorInputText(imageReplaceNewColorInput);
            if (string.IsNullOrWhiteSpace(sourceColor))
            {
                MessageBox.Show("原色不能为空。", "BioDraw");
                return;
            }

            var applyPreset = new ImageReplacePreset
            {
                Name = preset?.Name ?? "临时预设",
                SortOrder = preset?.SortOrder ?? 1,
                FuzzPercent = NormalizeFuzzPercent(preset?.FuzzPercent ?? imageReplaceFuzzInput),
                TargetColor = sourceColor,
                Mode = string.IsNullOrWhiteSpace(newColor) ? "transparent" : "fill",
                ReplacementColor = string.IsNullOrWhiteSpace(newColor) ? "black" : newColor
            };

            var replacedCount = 0;
            var failedCount = 0;
            var lastError = string.Empty;
            var replacedShapes = new List<dynamic>();
            string snapshotPath;
            if (!TryCreatePresentationSnapshot(application, out snapshotPath, out lastError))
            {
                MessageBox.Show(lastError, "BioDraw");
                return;
            }

            foreach (dynamic shape in shapes)
            {
                dynamic replacedShape;
                string error;
                if (!TryReplaceShapePictureWithMagick(shape, applyPreset, snapshotPath, out replacedShape, out error))
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

            TryReselectShapes(replacedShapes);
            TryDeleteFile(snapshotPath);

            if (replacedCount == 0)
            {
                MessageBox.Show(string.IsNullOrWhiteSpace(lastError) ? "处理失败：未找到可处理的图片。" : lastError, "BioDraw");
                return;
            }

            if (failedCount > 0)
            {
                SetStatusText($"BioDraw：已替换 {replacedCount} 张，{failedCount} 张未处理。");
                return;
            }
            SetStatusText($"BioDraw：已替换 {replacedCount} 张图片。");
        }

        public void OnEditImageReplacePreset(Office.IRibbonControl control)
        {
            EnsureImageReplaceInputValues();
            var preset = GetCurrentPreset();
            if (preset == null)
            {
                preset = CreateDefaultPreset();
                preset.Name = GenerateNewPresetName();
                preset.SortOrder = Math.Max(1, imageReplacePresets.Count + 1);
                preset.TargetColor = NormalizeColorInputText(imageReplaceSourceColorInput);
                preset.Mode = HasVisibleColorText(imageReplaceNewColorInput) ? "fill" : "transparent";
                preset.ReplacementColor = HasVisibleColorText(imageReplaceNewColorInput) ? NormalizeColorInputText(imageReplaceNewColorInput) : "black";
                preset.FuzzPercent = NormalizeFuzzPercent(imageReplaceFuzzInput);
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

            editedPreset.TargetColor = NormalizeColorInputText(imageReplaceSourceColorInput);
            editedPreset.Mode = HasVisibleColorText(imageReplaceNewColorInput) ? "fill" : "transparent";
            editedPreset.ReplacementColor = HasVisibleColorText(imageReplaceNewColorInput) ? NormalizeColorInputText(imageReplaceNewColorInput) : "black";
            editedPreset.FuzzPercent = NormalizeFuzzPercent(editedPreset.FuzzPercent);
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
            NormalizePresetSortOrders();
            EnsurePresetSelectionNames();
            SaveImageReplacePresets();
            SyncImageReplaceInputValuesFromCurrentPreset();
            InvalidateImageReplaceRibbonControls();
        }

        public string GetPresetMenuContent(Office.IRibbonControl control)
        {
            var sb = new StringBuilder();
            sb.Append("<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>");
            foreach (var preset in GetPresetsInDisplayOrder())
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

        private sealed class PictureConverter : AxHost
        {
            private PictureConverter() : base("")
            {
            }

            public static stdole.IPictureDisp ToPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
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
            imageReplaceFuzzInput = NormalizeFuzzPercent(preset.FuzzPercent);
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
            AddColorOption(imageReplaceSourceColorOptions, "white");
            AddColorOption(imageReplaceNewColorOptions, "white");
        }

        private static bool IsLegacyDefaultColorOption(string value)
        {
            return string.Equals(value, "white", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "black", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "red", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "blue", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "green", StringComparison.OrdinalIgnoreCase);
        }

        private static void TrimLegacyDefaultColorOptions(List<string> options)
        {
            if (options == null || options.Count == 0)
            {
                return;
            }

            if (!options.All(IsLegacyDefaultColorOption))
            {
                return;
            }

            options.Clear();
            options.Add("white");
        }

        private void EnsureImageReplaceColorOptions()
        {
            if (imageReplaceSourceColorOptions.Count == 0)
            {
                ResetImageReplaceColorOptions();
            }
        }

        private static string NormalizeColorInputText(string text)
        {
            var normalized = (text ?? string.Empty).Trim();
            if (string.Equals(normalized, RibbonEmptyInputToken, StringComparison.Ordinal))
            {
                return string.Empty;
            }
            return normalized;
        }

        private static string ToStorageColorInputText(string text)
        {
            var normalized = NormalizeColorInputText(text);
            return string.IsNullOrEmpty(normalized) ? RibbonEmptyInputToken : normalized;
        }

        private static bool HasVisibleColorText(string text)
        {
            return !string.IsNullOrWhiteSpace(NormalizeColorInputText(text));
        }

        private static string ToRibbonColorInputText(string text)
        {
            return ToStorageColorInputText(text);
        }

        private static int FindColorOptionIndex(List<string> options, string value)
        {
            if (options == null || options.Count == 0 || string.IsNullOrWhiteSpace(value))
            {
                return -1;
            }

            for (int i = 0; i < options.Count; i++)
            {
                if (string.Equals(options[i], value, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            return -1;
        }

        private static int FindExactMatchIndex(List<string> options, string value)
        {
            return FindColorOptionIndex(options, value);
        }

        private static bool AddColorOption(List<string> options, string value)
        {
            if (options == null)
            {
                return false;
            }

            var normalized = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            if (options.Any(x => string.Equals(x, normalized, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            options.Add(normalized);
            return true;
        }

        private static bool UpsertColorOptionAtPosition(List<string> options, string value, int oneBasedPosition)
        {
            if (options == null)
            {
                return false;
            }

            var normalized = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            var existingIndex = FindColorOptionIndex(options, normalized);
            if (existingIndex >= 0)
            {
                options.RemoveAt(existingIndex);
            }

            var insertIndex = Math.Max(0, Math.Min(options.Count, oneBasedPosition - 1));
            options.Insert(insertIndex, normalized);
            return true;
        }

        private static bool RemoveColorOption(List<string> options, string value)
        {
            if (options == null || options.Count == 0 || string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var index = FindColorOptionIndex(options, value);
            if (index < 0)
            {
                return false;
            }

            options.RemoveAt(index);
            return true;
        }

        private bool ShowColorOptionManagerDialog(string title, List<string> options, string currentValue, bool allowEmpty, out string selectedValue)
        {
            selectedValue = NormalizeColorInputText(currentValue);
            if (options == null)
            {
                return false;
            }
            var dialogSelectedValue = selectedValue;
            var currentNormalizedValue = NormalizeColorInputText(currentValue);
            var latestCommittedValue = currentNormalizedValue;

            using (var dialog = new Form())
            using (var listBox = new ListBox())
            using (var lblSortOrder = new Label())
            using (var numSortOrder = new NumericUpDown())
            using (var inputBox = new TextBox())
            using (var pickButton = new Button())
            using (var saveButton = new Button())
            using (var deleteButton = new Button())
            {
                dialog.Text = title;
                dialog.FormBorderStyle = FormBorderStyle.Sizable;
                dialog.StartPosition = FormStartPosition.CenterScreen;
                dialog.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Regular, GraphicsUnit.Point);
                dialog.BackColor = Color.FromArgb(244, 247, 252);
                dialog.ForeColor = Color.FromArgb(32, 41, 57);
                dialog.AutoScaleMode = AutoScaleMode.Dpi;
                dialog.MinimizeBox = true;
                dialog.MaximizeBox = true;
                dialog.ShowInTaskbar = false;
                dialog.MinimumSize = new Size(620, 280);
                dialog.ClientSize = new Size(680, 300);

                listBox.BorderStyle = BorderStyle.FixedSingle;
                listBox.IntegralHeight = false;
                listBox.BackColor = Color.White;
                listBox.ForeColor = Color.FromArgb(32, 41, 57);
                listBox.AllowDrop = true;

                lblSortOrder.Text = "位置";
                lblSortOrder.TextAlign = ContentAlignment.MiddleLeft;
                lblSortOrder.ForeColor = Color.FromArgb(32, 41, 57);

                numSortOrder.Minimum = 1;
                numSortOrder.Maximum = Math.Max(1, options.Count + 1);
                numSortOrder.DecimalPlaces = 0;
                numSortOrder.TextAlign = HorizontalAlignment.Right;
                numSortOrder.BorderStyle = BorderStyle.FixedSingle;
                numSortOrder.BackColor = Color.White;
                numSortOrder.ForeColor = Color.FromArgb(32, 41, 57);
                numSortOrder.Value = 1;

                inputBox.BorderStyle = BorderStyle.FixedSingle;
                inputBox.BackColor = Color.White;
                inputBox.ForeColor = Color.FromArgb(32, 41, 57);
                inputBox.Text = selectedValue;

                pickButton.Text = "取色";
                saveButton.Text = "保存";
                deleteButton.Text = "删除";

                void StyleActionButton(Button button, bool primary, bool danger)
                {
                    button.FlatStyle = FlatStyle.Flat;
                    button.FlatAppearance.BorderSize = 1;
                    if (danger)
                    {
                        button.FlatAppearance.BorderColor = Color.FromArgb(220, 53, 69);
                        button.BackColor = Color.FromArgb(220, 53, 69);
                        button.ForeColor = Color.White;
                    }
                    else
                    {
                        button.FlatAppearance.BorderColor = primary ? Color.FromArgb(24, 118, 242) : Color.FromArgb(189, 198, 213);
                        button.BackColor = primary ? Color.FromArgb(24, 118, 242) : Color.White;
                        button.ForeColor = primary ? Color.White : Color.FromArgb(43, 52, 69);
                    }
                    button.UseVisualStyleBackColor = false;
                    button.Cursor = Cursors.Hand;
                }

                StyleActionButton(pickButton, false, false);
                StyleActionButton(saveButton, true, false);
                StyleActionButton(deleteButton, false, true);

                Action syncSortOrderMaximum = () =>
                {
                    numSortOrder.Maximum = Math.Max(1, options.Count + 1);
                    if (numSortOrder.Value > numSortOrder.Maximum)
                    {
                        numSortOrder.Value = numSortOrder.Maximum;
                    }
                };

                Action syncSortOrderFromInput = () =>
                {
                    var selected = (inputBox.Text ?? string.Empty).Trim();
                    var index = FindColorOptionIndex(options, selected);
                    if (index >= 0)
                    {
                        numSortOrder.Value = Math.Min(numSortOrder.Maximum, index + 1);
                        return;
                    }
                    numSortOrder.Value = numSortOrder.Maximum;
                };

                Action refreshList = () =>
                {
                    var selected = (inputBox.Text ?? string.Empty).Trim();
                    syncSortOrderMaximum();
                    listBox.BeginUpdate();
                    listBox.Items.Clear();
                    foreach (var option in options)
                    {
                        listBox.Items.Add(option);
                    }
                    listBox.EndUpdate();
                    var index = FindColorOptionIndex(options, selected);
                    if (index >= 0)
                    {
                        listBox.SelectedIndex = index;
                    }
                    syncSortOrderFromInput();
                };

                var dragStartIndex = -1;
                var dragStartPoint = Point.Empty;

                listBox.SelectedIndexChanged += (sender, args) =>
                {
                    if (listBox.SelectedIndex >= 0 && listBox.SelectedIndex < listBox.Items.Count)
                    {
                        inputBox.Text = listBox.Items[listBox.SelectedIndex].ToString() ?? string.Empty;
                        syncSortOrderFromInput();
                    }
                };

                listBox.MouseDown += (sender, args) =>
                {
                    if (args.Button != MouseButtons.Left)
                    {
                        dragStartIndex = -1;
                        return;
                    }

                    dragStartPoint = args.Location;
                    dragStartIndex = listBox.IndexFromPoint(args.Location);
                    if (dragStartIndex >= 0 && dragStartIndex < listBox.Items.Count)
                    {
                        listBox.SelectedIndex = dragStartIndex;
                    }
                };

                listBox.MouseMove += (sender, args) =>
                {
                    if (args.Button != MouseButtons.Left || dragStartIndex < 0 || dragStartIndex >= listBox.Items.Count)
                    {
                        return;
                    }

                    var dragSize = SystemInformation.DragSize;
                    var dragRect = new Rectangle(
                        dragStartPoint.X - dragSize.Width / 2,
                        dragStartPoint.Y - dragSize.Height / 2,
                        dragSize.Width,
                        dragSize.Height);
                    if (dragRect.Contains(args.Location))
                    {
                        return;
                    }

                    var dragValue = listBox.Items[dragStartIndex].ToString() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(dragValue))
                    {
                        return;
                    }

                    listBox.DoDragDrop(dragValue, DragDropEffects.Move);
                    dragStartIndex = -1;
                };

                listBox.DragOver += (sender, args) =>
                {
                    if (!args.Data.GetDataPresent(typeof(string)))
                    {
                        args.Effect = DragDropEffects.None;
                        return;
                    }
                    args.Effect = DragDropEffects.Move;
                };

                listBox.DragDrop += (sender, args) =>
                {
                    if (!args.Data.GetDataPresent(typeof(string)))
                    {
                        return;
                    }

                    var draggedValue = args.Data.GetData(typeof(string)) as string ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(draggedValue))
                    {
                        return;
                    }

                    var sourceIndex = FindColorOptionIndex(options, draggedValue);
                    if (sourceIndex < 0)
                    {
                        return;
                    }

                    var dropPoint = listBox.PointToClient(new Point(args.X, args.Y));
                    int targetInsertIndex;
                    if (dropPoint.Y < 0)
                    {
                        targetInsertIndex = 0;
                    }
                    else if (options.Count > 0 && dropPoint.Y > listBox.GetItemRectangle(options.Count - 1).Bottom)
                    {
                        targetInsertIndex = options.Count;
                    }
                    else
                    {
                        var targetIndex = listBox.IndexFromPoint(dropPoint);
                        targetInsertIndex = targetIndex < 0 ? options.Count : targetIndex;
                    }
                    targetInsertIndex = Math.Max(0, Math.Min(targetInsertIndex, options.Count));

                    if (targetInsertIndex == sourceIndex || targetInsertIndex == sourceIndex + 1)
                    {
                        return;
                    }

                    options.RemoveAt(sourceIndex);
                    if (sourceIndex < targetInsertIndex)
                    {
                        targetInsertIndex--;
                    }
                    targetInsertIndex = Math.Max(0, Math.Min(targetInsertIndex, options.Count));
                    options.Insert(targetInsertIndex, draggedValue);

                    inputBox.Text = draggedValue;
                    dialogSelectedValue = draggedValue;
                    latestCommittedValue = draggedValue;
                    refreshList();
                    PersistImageReplaceInputMemory();
                    InvalidateImageReplaceRibbonControls();
                };

                pickButton.Click += (sender, args) =>
                {
                    string colorToken;
                    string errorMessage;
                    if (TryPickColorWithPowerPoint(false, inputBox.Text, out colorToken, out errorMessage))
                    {
                        inputBox.Text = colorToken;
                        return;
                    }

                    if (!string.IsNullOrWhiteSpace(errorMessage))
                    {
                        MessageBox.Show(errorMessage, "BioDraw");
                    }
                };

                saveButton.Click += (sender, args) =>
                {
                    var value = (inputBox.Text ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        return;
                    }

                    var position = Convert.ToInt32(numSortOrder.Value, CultureInfo.InvariantCulture);
                    UpsertColorOptionAtPosition(options, value, position);
                    inputBox.Text = value;
                    dialogSelectedValue = value;
                    latestCommittedValue = value;
                    refreshList();
                    PersistImageReplaceInputMemory();
                    InvalidateImageReplaceRibbonControls();
                };

                deleteButton.Click += (sender, args) =>
                {
                    var value = (inputBox.Text ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        return;
                    }

                    if (!RemoveColorOption(options, value))
                    {
                        return;
                    }

                    if (string.Equals(currentNormalizedValue, value, StringComparison.OrdinalIgnoreCase))
                    {
                        dialogSelectedValue = allowEmpty ? string.Empty : (options.FirstOrDefault() ?? currentNormalizedValue);
                    }
                    else
                    {
                        dialogSelectedValue = currentNormalizedValue;
                    }

                    latestCommittedValue = dialogSelectedValue;
                    inputBox.Text = dialogSelectedValue;
                    refreshList();
                    PersistImageReplaceInputMemory();
                    InvalidateImageReplaceRibbonControls();
                };

                dialog.FormClosing += (sender, args) =>
                {
                    var value = (inputBox.Text ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        dialogSelectedValue = allowEmpty ? string.Empty : latestCommittedValue;
                    }
                    else
                    {
                        dialogSelectedValue = value;
                    }
                };

                void ApplyDialogLayout()
                {
                    var margin = 24;
                    var hGap = 16;
                    var leftWidth = 180;
                    var rightWidth = 300;
                    var rightX = Math.Max(margin + leftWidth + hGap, dialog.ClientSize.Width - margin - rightWidth);
                    var rowHeight = 40;
                    var buttonGap = 10;
                    var buttonWidth = (rightWidth - buttonGap) / 2;
                    var pickWidth = 96;
                    var inputWidth = Math.Max(96, rightWidth - pickWidth - buttonGap);
                    var leftX = margin;
                    var top = margin;
                    var rowGap = 10;
                    var saveDeleteY = dialog.ClientSize.Height - margin - rowHeight;

                    listBox.Location = new Point(leftX, top);
                    listBox.Size = new Size(leftWidth, Math.Max(120, saveDeleteY - top - buttonGap));

                    inputBox.Location = new Point(rightX, top);
                    inputBox.Size = new Size(inputWidth, rowHeight);

                    pickButton.Location = new Point(inputBox.Right + buttonGap, top);
                    pickButton.Size = new Size(pickWidth, rowHeight);

                    var positionY = top + rowHeight + rowGap;
                    lblSortOrder.Location = new Point(rightX, positionY);
                    lblSortOrder.Size = new Size(120, rowHeight);
                    numSortOrder.Location = new Point(lblSortOrder.Right + 12, positionY);
                    numSortOrder.Size = new Size(120, rowHeight);

                    saveButton.Location = new Point(rightX, saveDeleteY);
                    saveButton.Size = new Size(buttonWidth, rowHeight);

                    deleteButton.Location = new Point(saveButton.Right + buttonGap, saveDeleteY);
                    deleteButton.Size = new Size(buttonWidth, rowHeight);

                    ApplyRoundedRegion(pickButton, 7);
                    ApplyRoundedRegion(saveButton, 7);
                    ApplyRoundedRegion(deleteButton, 7);
                }

                dialog.Controls.Add(listBox);
                dialog.Controls.Add(inputBox);
                dialog.Controls.Add(pickButton);
                dialog.Controls.Add(lblSortOrder);
                dialog.Controls.Add(numSortOrder);
                dialog.Controls.Add(saveButton);
                dialog.Controls.Add(deleteButton);
                dialog.Resize += (_, __) => ApplyDialogLayout();
                ApplyDialogLayout();

                EnsureImageReplaceColorOptions();
                refreshList();
                dialog.ShowDialog();
                selectedValue = dialogSelectedValue;
                return true;
            }
        }

        private void PersistImageReplaceInputMemory()
        {
            SaveImageReplacePresets();
        }

        private void InvalidateImageReplaceRibbonControls()
        {
            ribbon?.InvalidateControl("ApplyImageReplace");
            ribbon?.InvalidateControl("DdImageReplacePreset");
            ribbon?.InvalidateControl("TxtImageReplaceSourceColor");
            ribbon?.InvalidateControl("TxtImageReplaceNewColor");
        }

        private IEnumerable<ImageReplacePreset> GetPresetsInDisplayOrder()
        {
            return imageReplacePresets
                .OrderBy(x => x.SortOrder)
                .ThenBy(x => x.Name, StringComparer.OrdinalIgnoreCase);
        }

        private void NormalizePresetSortOrders()
        {
            var ordered = GetPresetsInDisplayOrder().ToList();
            for (int i = 0; i < ordered.Count; i++)
            {
                ordered[i].SortOrder = i + 1;
            }

            imageReplacePresets.Clear();
            imageReplacePresets.AddRange(ordered);
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

            var ordered = GetPresetsInDisplayOrder().ToList();
            var insertIndex = Math.Max(0, Math.Min(desiredSortOrder - 1, ordered.Count));
            ordered.Insert(insertIndex, editedPreset);
            for (int i = 0; i < ordered.Count; i++)
            {
                ordered[i].SortOrder = i + 1;
            }

            imageReplacePresets.Clear();
            imageReplacePresets.AddRange(ordered);
        }

        private string GenerateNewPresetName()
        {
            var baseName = "新预设";
            var index = 1;
            string name;
            do
            {
                name = baseName + index.ToString(CultureInfo.InvariantCulture);
                index++;
            }
            while (imageReplacePresets.Any(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase)));
            return name;
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
                imageReplaceFuzzInput = ParseFuzz((string)root.Attribute("ImageReplaceFuzzInput"));
                materialPreviewCount = ParseMaterialPreviewCount((string)root.Attribute("MaterialPreviewCount"));
                hasPresetEditorBounds = TryParseEditorBounds(root, out presetEditorBounds);
                presetEditorSaveAsDefaultChecked = ParseBool((string)root.Attribute("EditorSaveAsDefault"));

                foreach (var sourceOption in root.Elements("SourceColorOption"))
                {
                    AddColorOption(imageReplaceSourceColorOptions, (string)sourceOption.Attribute("Value"));
                }

                foreach (var newOption in root.Elements("NewColorOption"))
                {
                    AddColorOption(imageReplaceNewColorOptions, (string)newOption.Attribute("Value"));
                }

                TrimLegacyDefaultColorOptions(imageReplaceSourceColorOptions);
                TrimLegacyDefaultColorOptions(imageReplaceNewColorOptions);

                foreach (var xPreset in root.Elements("Preset"))
                {
                    var preset = new ImageReplacePreset
                    {
                        Name = (string)xPreset.Attribute("Name") ?? "Default",
                        TargetColor = (string)xPreset.Attribute("TargetColor") ?? "white",
                        Mode = (string)xPreset.Attribute("Mode") ?? "transparent",
                        ReplacementColor = (string)xPreset.Attribute("ReplacementColor") ?? "black",
                        FuzzPercent = ParseFuzz((string)xPreset.Attribute("FuzzPercent")),
                        SortOrder = ParseSortOrder((string)xPreset.Attribute("SortOrder"), imageReplacePresets.Count + 1)
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

            NormalizePresetSortOrders();
            EnsurePresetSelectionNames();
            imageReplaceSourceColorInput = string.Empty;
            imageReplaceNewColorInput = string.Empty;
            imageReplaceFuzzInput = NormalizeFuzzPercent(imageReplaceFuzzInput);
            materialPreviewCount = ParseMaterialPreviewCount(materialPreviewCount.ToString(CultureInfo.InvariantCulture));
            EnsureImageReplaceColorOptions();
        }

        private void SaveImageReplacePresets()
        {
            var dir = Path.GetDirectoryName(presetStorePath);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            NormalizePresetSortOrders();
            EnsurePresetSelectionNames();
            imageReplaceFuzzInput = NormalizeFuzzPercent(imageReplaceFuzzInput);
            EnsureImageReplaceColorOptions();
            var root = new XElement(
                "Presets",
                new XAttribute("Default", defaultPresetName ?? string.Empty),
                new XAttribute("Current", currentPresetName ?? string.Empty),
                new XAttribute("MaterialLibraryPath", materialLibraryPath ?? string.Empty),
                new XAttribute("MaterialPreviewCount", GetMaterialPageSize()),
                new XAttribute("ImageReplaceSourceInput", NormalizeColorInputText(imageReplaceSourceColorInput)),
                new XAttribute("ImageReplaceNewInput", NormalizeColorInputText(imageReplaceNewColorInput)),
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
                    new XAttribute("FuzzPercent", NormalizeFuzzPercent(p.FuzzPercent).ToString("0.0", CultureInfo.InvariantCulture)),
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

        private static ImageReplacePreset CreateDefaultPreset()
        {
            return new ImageReplacePreset
            {
                Name = "默认预设",
                SortOrder = 1,
                FuzzPercent = 5,
                TargetColor = "white",
                Mode = "transparent",
                ReplacementColor = "black"
            };
        }

        private static int ParseSortOrder(string sortText, int fallbackValue)
        {
            int sortOrder;
            if (int.TryParse(sortText, NumberStyles.Integer, CultureInfo.InvariantCulture, out sortOrder) && sortOrder > 0)
            {
                return sortOrder;
            }
            return Math.Max(1, fallbackValue);
        }

        private static int ParseMaterialPreviewCount(string countText)
        {
            int count;
            if (int.TryParse(countText, NumberStyles.Integer, CultureInfo.InvariantCulture, out count))
            {
                return Math.Max(1, Math.Min(MaterialPreviewButtonCount, count));
            }
            return 5;
        }

        private static double ParseFuzz(string fuzzText)
        {
            double fuzz;
            if (double.TryParse(fuzzText, NumberStyles.Float, CultureInfo.InvariantCulture, out fuzz))
            {
                return NormalizeFuzzPercent(fuzz);
            }
            return 5;
        }

        private static double NormalizeFuzzPercent(double fuzz)
        {
            if (double.IsNaN(fuzz) || double.IsInfinity(fuzz))
            {
                return 5.0;
            }
            if (fuzz < 0)
            {
                fuzz = 0;
            }
            if (fuzz > 100)
            {
                fuzz = 100;
            }
            return Math.Round(fuzz, 1, MidpointRounding.AwayFromZero);
        }

        private static bool ParseBool(string boolText)
        {
            bool value;
            if (bool.TryParse(boolText, out value))
            {
                return value;
            }
            return false;
        }

        private static bool TryParseEditorBounds(XElement root, out Rectangle bounds)
        {
            bounds = Rectangle.Empty;
            if (root == null)
            {
                return false;
            }

            int x;
            int y;
            int w;
            int h;
            if (!int.TryParse((string)root.Attribute("EditorX"), NumberStyles.Integer, CultureInfo.InvariantCulture, out x) ||
                !int.TryParse((string)root.Attribute("EditorY"), NumberStyles.Integer, CultureInfo.InvariantCulture, out y) ||
                !int.TryParse((string)root.Attribute("EditorWidth"), NumberStyles.Integer, CultureInfo.InvariantCulture, out w) ||
                !int.TryParse((string)root.Attribute("EditorHeight"), NumberStyles.Integer, CultureInfo.InvariantCulture, out h))
            {
                return false;
            }

            if (w < 620 || h < 360)
            {
                return false;
            }

            bounds = new Rectangle(x, y, w, h);
            return true;
        }

        private bool RunImageMagickReplace(string sourcePath, string outputPath, ImageReplacePreset preset, out string errorMessage)
        {
            errorMessage = string.Empty;

            var arguments = new StringBuilder();
            arguments.Append(QuoteArg(sourcePath));
            arguments.Append(" -fuzz ");
            arguments.Append(preset.FuzzPercent.ToString("0.##", CultureInfo.InvariantCulture));
            arguments.Append("% ");

            if (string.Equals(preset.Mode, "fill", StringComparison.OrdinalIgnoreCase))
            {
                arguments.Append("-fill ");
                arguments.Append(QuoteArg(preset.ReplacementColor));
                arguments.Append(" -opaque ");
                arguments.Append(QuoteArg(preset.TargetColor));
                arguments.Append(" ");
            }
            else
            {
                arguments.Append("-transparent ");
                arguments.Append(QuoteArg(preset.TargetColor));
                arguments.Append(" ");
            }

            arguments.Append(QuoteArg(outputPath));

            try
            {
                var processStartInfo = new ProcessStartInfo
                {
                    FileName = "magick",
                    Arguments = arguments.ToString(),
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(processStartInfo))
                {
                    if (process == null)
                    {
                        errorMessage = "无法启动 ImageMagick。";
                        return false;
                    }

                    var stdOut = process.StandardOutput.ReadToEnd();
                    var stdErr = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        errorMessage = string.IsNullOrWhiteSpace(stdErr) ? stdOut : stdErr;
                        if (string.IsNullOrWhiteSpace(errorMessage))
                        {
                            errorMessage = "ImageMagick 执行失败。";
                        }
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static string QuoteArg(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\"", "\\\"") + "\"";
        }

        private bool TryGetSelectedShapes(dynamic selection, out List<dynamic> shapes)
        {
            shapes = new List<dynamic>();
            try
            {
                var shapeRange = selection.ShapeRange;
                if (shapeRange == null)
                {
                    return false;
                }

                var count = 0;
                try
                {
                    count = (int)shapeRange.Count;
                }
                catch
                {
                }

                if (count <= 0)
                {
                    return false;
                }

                for (int i = 1; i <= count; i++)
                {
                    shapes.Add(shapeRange[i]);
                }

                return shapes.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        private bool TryReplaceShapePictureWithMagick(dynamic shape, ImageReplacePreset preset, string snapshotPath, out dynamic replacedShape, out string errorMessage)
        {
            replacedShape = null;
            errorMessage = string.Empty;
            if (shape == null)
            {
                errorMessage = "未找到可处理的图片。";
                return false;
            }

            try
            {
                var type = 0;
                try
                {
                    type = (int)shape.Type;
                }
                catch
                {
                }

                if (type != 13)
                {
                    errorMessage = "选中对象不是图片。";
                    return false;
                }

                string sourcePath;
                if (!TryExtractOriginalImageFromPptx(shape, snapshotPath, out sourcePath, out errorMessage))
                {
                    return false;
                }

                var outputPath = BuildMagickOutputPath(sourcePath, preset);

                if (!RunImageMagickReplace(sourcePath, outputPath, preset, out errorMessage))
                {
                    TryDeleteFile(sourcePath);
                    return false;
                }

                if (!TryReplaceShapeImageInPlace(shape, outputPath, out replacedShape, out errorMessage))
                {
                    TryDeleteFile(sourcePath);
                    TryDeleteFile(outputPath);
                    return false;
                }

                TryDeleteFile(sourcePath);
                TryDeleteFile(outputPath);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static string BuildMagickOutputPath(string sourcePath, ImageReplacePreset preset)
        {
            var outputExtension = Path.GetExtension(sourcePath);
            if (string.IsNullOrWhiteSpace(outputExtension))
            {
                outputExtension = ".png";
            }

            if (string.Equals(preset.Mode, "transparent", StringComparison.OrdinalIgnoreCase) &&
                IsTransparencyUnsupportedFormat(outputExtension))
            {
                outputExtension = ".png";
            }

            return Path.Combine(
                Path.GetDirectoryName(sourcePath) ?? Path.GetTempPath(),
                Path.GetFileNameWithoutExtension(sourcePath) + "_magick" + outputExtension);
        }

        private static bool IsTransparencyUnsupportedFormat(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension))
            {
                return true;
            }

            switch (extension.Trim().ToLowerInvariant())
            {
                case ".jpg":
                case ".jpeg":
                case ".jpe":
                case ".jfif":
                case ".bmp":
                case ".dib":
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryGetPictureCropValues(dynamic shape, out float cropLeft, out float cropTop, out float cropRight, out float cropBottom)
        {
            cropLeft = 0f;
            cropTop = 0f;
            cropRight = 0f;
            cropBottom = 0f;
            try
            {
                cropLeft = (float)shape.PictureFormat.CropLeft;
                cropTop = (float)shape.PictureFormat.CropTop;
                cropRight = (float)shape.PictureFormat.CropRight;
                cropBottom = (float)shape.PictureFormat.CropBottom;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool TryExtractOriginalImageFromPptx(dynamic shape, string snapshotPath, out string filePath, out string errorMessage)
        {
            filePath = null;
            errorMessage = string.Empty;

            try
            {
                if (string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
                {
                    errorMessage = "无法读取当前演示文稿快照。";
                    return false;
                }

                var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw");
                Directory.CreateDirectory(tempDir);

                var slideIndex = (int)shape.Parent.SlideIndex;
                var shapeId = (int)shape.Id;

                using (var stream = new FileStream(snapshotPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, false))
                {
                    string slidePartPath;
                    if (!TryResolveSlidePartPath(archive, slideIndex, out slidePartPath, out errorMessage))
                    {
                        return false;
                    }

                    string mediaPartPath;
                    if (!TryResolveMediaPathForShape(archive, slidePartPath, shapeId, out mediaPartPath, out errorMessage))
                    {
                        return false;
                    }

                    var mediaEntry = archive.GetEntry(mediaPartPath);
                    if (mediaEntry == null)
                    {
                        errorMessage = "未找到选中图片对应的媒体文件。";
                        return false;
                    }

                    var ext = Path.GetExtension(mediaPartPath);
                    if (string.IsNullOrWhiteSpace(ext))
                    {
                        ext = ".png";
                    }

                    var targetFilePath = Path.Combine(tempDir, "ppt_media_" + Guid.NewGuid().ToString("N") + ext);
                    using (var entryStream = mediaEntry.Open())
                    using (var output = new FileStream(targetFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        entryStream.CopyTo(output);
                    }

                    if (!File.Exists(targetFilePath) || new FileInfo(targetFilePath).Length <= 0)
                    {
                        errorMessage = "读取原始图片失败。";
                        return false;
                    }

                    filePath = targetFilePath;
                    return true;
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private bool TryCreatePresentationSnapshot(dynamic application, out string snapshotPath, out string errorMessage)
        {
            snapshotPath = null;
            errorMessage = string.Empty;
            try
            {
                var presentation = application?.ActivePresentation;
                if (presentation == null)
                {
                    errorMessage = "未找到当前演示文稿。";
                    return false;
                }

                var tempDir = Path.Combine(Path.GetTempPath(), "BioDraw");
                Directory.CreateDirectory(tempDir);
                snapshotPath = Path.Combine(tempDir, "ppt_snapshot_" + Guid.NewGuid().ToString("N") + ".pptx");

                presentation.SaveCopyAs(snapshotPath);
                if (!File.Exists(snapshotPath) || new FileInfo(snapshotPath).Length <= 0)
                {
                    errorMessage = "无法创建演示文稿快照。";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = "无法读取当前演示文稿，请确认文档已正常打开。 " + ex.Message;
                snapshotPath = null;
                return false;
            }
        }

        private static void SetStatusText(string text)
        {
            _ = text;
        }

        private static bool TryResolveSlidePartPath(ZipArchive archive, int slideIndex, out string slidePartPath, out string errorMessage)
        {
            slidePartPath = null;
            errorMessage = string.Empty;
            try
            {
                var presentationEntry = archive.GetEntry("ppt/presentation.xml");
                var presentationRelsEntry = archive.GetEntry("ppt/_rels/presentation.xml.rels");
                if (presentationEntry == null || presentationRelsEntry == null)
                {
                    errorMessage = "PPTX 结构异常，找不到演示文稿索引。";
                    return false;
                }

                XDocument presentationDoc;
                XDocument relsDoc;
                using (var stream = presentationEntry.Open())
                {
                    presentationDoc = XDocument.Load(stream);
                }
                using (var stream = presentationRelsEntry.Open())
                {
                    relsDoc = XDocument.Load(stream);
                }

                var p = (XNamespace)"http://schemas.openxmlformats.org/presentationml/2006/main";
                var r = (XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                var slideIdNodes = presentationDoc.Descendants(p + "sldId").ToList();
                if (slideIndex <= 0 || slideIndex > slideIdNodes.Count)
                {
                    errorMessage = "无法定位选中图片所在幻灯片。";
                    return false;
                }

                var slideRid = (string)slideIdNodes[slideIndex - 1].Attribute(r + "id");
                if (string.IsNullOrWhiteSpace(slideRid))
                {
                    errorMessage = "幻灯片关系索引缺失。";
                    return false;
                }

                var rel = relsDoc.Root?
                    .Elements()
                    .FirstOrDefault(x => string.Equals((string)x.Attribute("Id"), slideRid, StringComparison.Ordinal));
                var target = (string)rel?.Attribute("Target");
                if (string.IsNullOrWhiteSpace(target))
                {
                    errorMessage = "幻灯片关系目标缺失。";
                    return false;
                }

                slidePartPath = ResolveZipPartPath("ppt/presentation.xml", target);
                if (archive.GetEntry(slidePartPath) == null)
                {
                    errorMessage = "未找到幻灯片数据。";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static bool TryResolveMediaPathForShape(ZipArchive archive, string slidePartPath, int shapeId, out string mediaPartPath, out string errorMessage)
        {
            mediaPartPath = null;
            errorMessage = string.Empty;
            try
            {
                var slideEntry = archive.GetEntry(slidePartPath);
                var slideRelsPath = GetRelationshipPartPath(slidePartPath);
                var slideRelsEntry = archive.GetEntry(slideRelsPath);
                if (slideEntry == null || slideRelsEntry == null)
                {
                    errorMessage = "找不到幻灯片图片关系文件。";
                    return false;
                }

                XDocument slideDoc;
                XDocument relsDoc;
                using (var stream = slideEntry.Open())
                {
                    slideDoc = XDocument.Load(stream);
                }
                using (var stream = slideRelsEntry.Open())
                {
                    relsDoc = XDocument.Load(stream);
                }

                var p = (XNamespace)"http://schemas.openxmlformats.org/presentationml/2006/main";
                var a = (XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";
                var r = (XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                var targetPic = slideDoc.Descendants(p + "pic")
                    .FirstOrDefault(pic =>
                    {
                        var idAttr = (string)pic
                            .Element(p + "nvPicPr")?
                            .Element(p + "cNvPr")?
                            .Attribute("id");
                        int idValue;
                        return int.TryParse(idAttr, out idValue) && idValue == shapeId;
                    });

                if (targetPic == null)
                {
                    errorMessage = "无法定位选中图片对应的原始资源。";
                    return false;
                }

                var embedRid = (string)targetPic
                    .Element(p + "blipFill")?
                    .Element(a + "blip")?
                    .Attribute(r + "embed");
                if (string.IsNullOrWhiteSpace(embedRid))
                {
                    errorMessage = "该图片不包含可提取的嵌入资源。";
                    return false;
                }

                var rel = relsDoc.Root?
                    .Elements()
                    .FirstOrDefault(x => string.Equals((string)x.Attribute("Id"), embedRid, StringComparison.Ordinal));
                var target = (string)rel?.Attribute("Target");
                if (string.IsNullOrWhiteSpace(target))
                {
                    errorMessage = "未找到图片关系映射。";
                    return false;
                }

                mediaPartPath = ResolveZipPartPath(slidePartPath, target);
                if (archive.GetEntry(mediaPartPath) == null)
                {
                    errorMessage = "媒体文件不存在。";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static string ResolveZipPartPath(string basePartPath, string relativeTarget)
        {
            var normalizedBase = basePartPath.Replace("\\", "/");
            var normalizedTarget = relativeTarget.Replace("\\", "/");

            if (normalizedTarget.StartsWith("/", StringComparison.Ordinal))
            {
                return normalizedTarget.TrimStart('/');
            }

            var baseUri = new Uri("http://local/" + normalizedBase, UriKind.Absolute);
            var resolvedUri = new Uri(baseUri, normalizedTarget);
            return resolvedUri.AbsolutePath.TrimStart('/');
        }

        private static string GetRelationshipPartPath(string partPath)
        {
            var normalized = partPath.Replace("\\", "/");
            var lastSlash = normalized.LastIndexOf('/');
            if (lastSlash < 0)
            {
                return "_rels/" + normalized + ".rels";
            }

            var dir = normalized.Substring(0, lastSlash);
            var file = normalized.Substring(lastSlash + 1);
            return dir + "/_rels/" + file + ".rels";
        }

        private bool TryReplaceShapeImageInPlace(dynamic shape, string outputPath, out dynamic newShape, out string errorMessage)
        {
            newShape = null;
            errorMessage = string.Empty;
            try
            {
                if (!File.Exists(outputPath))
                {
                    errorMessage = "输出文件不存在。";
                    return false;
                }

                var left = (float)shape.Left;
                var top = (float)shape.Top;
                var width = (float)shape.Width;
                var height = (float)shape.Height;
                var rotation = (float)shape.Rotation;
                var zOrderPosition = (int)shape.ZOrderPosition;
                var shapeName = string.Empty;
                try
                {
                    shapeName = (string)shape.Name;
                }
                catch
                {
                }
                var lockAspectRatio = 0;
                try
                {
                    lockAspectRatio = (int)shape.LockAspectRatio;
                }
                catch
                {
                }

                var cropLeft = 0f;
                var cropTop = 0f;
                var cropRight = 0f;
                var cropBottom = 0f;
                var hasCrop = TryGetPictureCropValues(shape, out cropLeft, out cropTop, out cropRight, out cropBottom) &&
                    (Math.Abs(cropLeft) > 0.01f || Math.Abs(cropTop) > 0.01f || Math.Abs(cropRight) > 0.01f || Math.Abs(cropBottom) > 0.01f);
                var insertionLeft = left;
                var insertionTop = top;
                var insertionWidth = width;
                var insertionHeight = height;
                if (hasCrop)
                {
                    insertionLeft = left - cropLeft;
                    insertionTop = top - cropTop;
                    insertionWidth = Math.Max(1f, width + cropLeft + cropRight);
                    insertionHeight = Math.Max(1f, height + cropTop + cropBottom);
                }

                try
                {
                    var shapes = shape.Parent.Shapes;
                    shape.Delete();

                    newShape = shapes.AddPicture(
                        outputPath,
                        Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue,
                        insertionLeft,
                        insertionTop,
                        insertionWidth,
                        insertionHeight);

                    try
                    {
                        if (!string.IsNullOrWhiteSpace(shapeName))
                        {
                            newShape.Name = shapeName;
                        }
                    }
                    catch
                    {
                    }

                    try
                    {
                        newShape.Rotation = rotation;
                    }
                    catch
                    {
                    }
                    try
                    {
                        newShape.LockAspectRatio = lockAspectRatio;
                    }
                    catch
                    {
                    }

                    if (hasCrop)
                    {
                        try
                        {
                            newShape.PictureFormat.CropLeft = cropLeft;
                            newShape.PictureFormat.CropTop = cropTop;
                            newShape.PictureFormat.CropRight = cropRight;
                            newShape.PictureFormat.CropBottom = cropBottom;
                        }
                        catch
                        {
                        }
                    }

                    try
                    {
                        newShape.Left = left;
                        newShape.Top = top;
                        newShape.Width = width;
                        newShape.Height = height;
                    }
                    catch
                    {
                    }

                    try
                    {
                        var guard = 0;
                        while ((int)newShape.ZOrderPosition > zOrderPosition && guard < 2048)
                        {
                            newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward);
                            guard++;
                        }
                    }
                    catch
                    {
                    }
                }
                catch
                {
                    errorMessage = "替换图片失败。";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                newShape = null;
                return false;
            }
        }

        private static void TryReselectShapes(List<dynamic> shapes)
        {
            if (shapes == null || shapes.Count == 0)
            {
                return;
            }

            try
            {
                if (shapes.Count == 1)
                {
                    shapes[0].Select();
                    return;
                }

                var selectedCount = 0;
                for (int i = 0; i < shapes.Count; i++)
                {
                    var shape = shapes[i];
                    if (shape == null)
                    {
                        continue;
                    }

                    if (selectedCount == 0)
                    {
                        shape.Select(Microsoft.Office.Core.MsoTriState.msoTrue);
                    }
                    else
                    {
                        shape.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
                    }
                    selectedCount++;
                }

                if (selectedCount > 1)
                {
                    return;
                }

                if (selectedCount == 1)
                {
                    return;
                }

                var parentShapes = shapes[0].Parent.Shapes;
                var ids = new int[shapes.Count];
                for (int i = 0; i < shapes.Count; i++)
                {
                    ids[i] = (int)shapes[i].Id;
                }
                var range = parentShapes.Range(ids);
                range.Select();
            }
            catch
            {
                try
                {
                    shapes[0].Select();
                }
                catch
                {
                }
            }
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private bool ShowPresetEditorDialog(ImageReplacePreset source, bool canDelete, out ImageReplacePreset result, out bool setAsDefault, out bool deleteRequested)
        {
            result = null;
            setAsDefault = false;
            deleteRequested = false;

            using (var form = new Form())
            using (var lblPresetName = new Label())
            using (var txtPresetName = new TextBox())
            using (var lblSortOrder = new Label())
            using (var numSortOrder = new NumericUpDown())
            using (var lblFuzz = new Label())
            using (var numFuzz = new NumericUpDown())
            using (var tbFuzz = new TrackBar())
            using (var chkDefault = new CheckBox())
            using (var btnDelete = new Button())
            using (var btnOk = new Button())
            {
                form.Text = "颜色替换参数";
                form.FormBorderStyle = FormBorderStyle.Sizable;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Regular, GraphicsUnit.Point);
                form.BackColor = Color.FromArgb(244, 247, 252);
                form.ForeColor = Color.FromArgb(32, 41, 57);
                form.AutoScaleMode = AutoScaleMode.Dpi;
                form.MinimizeBox = false;
                form.MaximizeBox = true;
                form.MinimumSize = new Size(760, 360);
                form.ClientSize = new Size(820, 400);
                if (hasPresetEditorBounds)
                {
                    form.StartPosition = FormStartPosition.Manual;
                    form.Bounds = presetEditorBounds;
                }

                lblPresetName.Text = "名称";
                lblPresetName.TextAlign = ContentAlignment.MiddleLeft;
                txtPresetName.Text = source.Name;
                txtPresetName.BorderStyle = BorderStyle.FixedSingle;

                lblSortOrder.Text = "位置";
                lblSortOrder.TextAlign = ContentAlignment.MiddleLeft;
                numSortOrder.Minimum = 1;
                numSortOrder.Maximum = Math.Max(1, imageReplacePresets.Count + 1);
                numSortOrder.DecimalPlaces = 0;
                numSortOrder.Value = Convert.ToDecimal(Math.Max(1, Math.Min((int)numSortOrder.Maximum, source.SortOrder)), CultureInfo.InvariantCulture);

                lblFuzz.Text = "近似度";
                lblFuzz.TextAlign = ContentAlignment.MiddleLeft;
                numFuzz.Minimum = 0;
                numFuzz.Maximum = 100;
                numFuzz.DecimalPlaces = 1;
                numFuzz.Increment = 0.1m;
                numFuzz.Value = Convert.ToDecimal(NormalizeFuzzPercent(source.FuzzPercent), CultureInfo.InvariantCulture);
                numFuzz.BorderStyle = BorderStyle.FixedSingle;
                numFuzz.TextAlign = HorizontalAlignment.Right;

                tbFuzz.Minimum = 0;
                tbFuzz.Maximum = 1000;
                tbFuzz.TickFrequency = 50;
                tbFuzz.SmallChange = 5;
                tbFuzz.LargeChange = 5;
                tbFuzz.AutoSize = false;
                tbFuzz.Value = Math.Max(tbFuzz.Minimum, Math.Min(tbFuzz.Maximum, (int)Math.Round(Convert.ToDouble(numFuzz.Value, CultureInfo.InvariantCulture) * 10, MidpointRounding.AwayFromZero)));

                chkDefault.Text = "保存为默认预设";
                chkDefault.Checked = presetEditorSaveAsDefaultChecked;
                chkDefault.AutoSize = true;

                btnDelete.Text = "删除";
                btnDelete.Enabled = canDelete;
                btnOk.Text = "保存";
                btnOk.DialogResult = DialogResult.OK;

                void StyleInputControl(Control control)
                {
                    control.BackColor = Color.White;
                    control.ForeColor = Color.FromArgb(32, 41, 57);
                }

                void StyleActionButton(Button button, bool primary, bool danger)
                {
                    button.FlatStyle = FlatStyle.Flat;
                    button.FlatAppearance.BorderSize = 1;
                    if (danger)
                    {
                        button.FlatAppearance.BorderColor = Color.FromArgb(220, 53, 69);
                        button.BackColor = Color.FromArgb(220, 53, 69);
                        button.ForeColor = Color.White;
                    }
                    else
                    {
                        button.FlatAppearance.BorderColor = primary ? Color.FromArgb(24, 118, 242) : Color.FromArgb(189, 198, 213);
                        button.BackColor = primary ? Color.FromArgb(24, 118, 242) : Color.White;
                        button.ForeColor = primary ? Color.White : Color.FromArgb(43, 52, 69);
                    }
                    button.UseVisualStyleBackColor = false;
                    button.Cursor = Cursors.Hand;
                }

                StyleInputControl(txtPresetName);
                StyleInputControl(numSortOrder);
                StyleInputControl(numFuzz);
                chkDefault.ForeColor = Color.FromArgb(43, 52, 69);
                StyleActionButton(btnOk, true, false);
                StyleActionButton(btnDelete, false, true);

                var syncingFuzz = false;
                void SyncFuzzToTrackBar()
                {
                    if (syncingFuzz)
                    {
                        return;
                    }
                    syncingFuzz = true;
                    var value = Convert.ToDouble(numFuzz.Value, CultureInfo.InvariantCulture);
                    var trackValue = Math.Max(tbFuzz.Minimum, Math.Min(tbFuzz.Maximum, (int)Math.Round(value * 10, MidpointRounding.AwayFromZero)));
                    tbFuzz.Value = trackValue;
                    syncingFuzz = false;
                }

                void SyncFuzzToNumeric()
                {
                    if (syncingFuzz)
                    {
                        return;
                    }
                    syncingFuzz = true;
                    numFuzz.Value = Convert.ToDecimal(tbFuzz.Value / 10.0, CultureInfo.InvariantCulture);
                    syncingFuzz = false;
                }

                numFuzz.ValueChanged += (_, __) => SyncFuzzToTrackBar();
                tbFuzz.Scroll += (_, __) => SyncFuzzToNumeric();
                tbFuzz.MouseEnter += (_, __) => tbFuzz.Focus();
                tbFuzz.MouseWheel += (_, e) =>
                {
                    var delta = e.Delta > 0 ? 5 : -5;
                    var next = Math.Max(tbFuzz.Minimum, Math.Min(tbFuzz.Maximum, tbFuzz.Value + delta));
                    if (next == tbFuzz.Value)
                    {
                        return;
                    }
                    tbFuzz.Value = next;
                    SyncFuzzToNumeric();
                };

                void ApplyDialogLayout()
                {
                    var margin = 24;
                    var labelWidth = 120;
                    var fieldGap = 12;
                    var rowHeight = 40;
                    var rowGap = 18;
                    var buttonWidth = 128;
                    var buttonHeight = 40;

                    var fieldX = margin + labelWidth + fieldGap;
                    var rightEdge = form.ClientSize.Width - margin;
                    var top = margin + 8;

                    lblPresetName.Location = new Point(margin, top);
                    lblPresetName.Size = new Size(labelWidth, rowHeight);
                    txtPresetName.Location = new Point(fieldX, top);
                    txtPresetName.Size = new Size(rightEdge - fieldX, rowHeight);

                    var row2Y = top + rowHeight + rowGap;
                    lblSortOrder.Location = new Point(margin, row2Y);
                    lblSortOrder.Size = new Size(labelWidth, rowHeight);
                    numSortOrder.Location = new Point(fieldX, row2Y);
                    numSortOrder.Size = new Size(180, rowHeight);

                    var row3Y = row2Y + rowHeight + rowGap;
                    lblFuzz.Location = new Point(margin, row3Y);
                    lblFuzz.Size = new Size(labelWidth, rowHeight);
                    numFuzz.Location = new Point(fieldX, row3Y);
                    numFuzz.Size = new Size(180, rowHeight);
                    tbFuzz.Location = new Point(numFuzz.Right + 14, row3Y + 4);
                    tbFuzz.Size = new Size(Math.Max(200, rightEdge - tbFuzz.Left), rowHeight - 8);

                    var contentBottom = row3Y + rowHeight;
                    var preferredBottomY = contentBottom + 22;
                    var bottomY = Math.Max(preferredBottomY, form.ClientSize.Height - margin - buttonHeight);

                    chkDefault.Location = new Point(margin, bottomY + Math.Max(0, (buttonHeight - chkDefault.Height) / 2));
                    var buttonsStartX = rightEdge - (buttonWidth * 2) - 10;
                    btnOk.Location = new Point(buttonsStartX, bottomY);
                    btnOk.Size = new Size(buttonWidth, buttonHeight);
                    btnDelete.Location = new Point(btnOk.Right + 10, bottomY);
                    btnDelete.Size = new Size(buttonWidth, buttonHeight);
                    ApplyRoundedRegion(btnOk, 7);
                    ApplyRoundedRegion(btnDelete, 7);
                }

                var localDeleteRequested = false;
                btnDelete.Click += (_, __) =>
                {
                    if (!canDelete)
                    {
                        return;
                    }
                    localDeleteRequested = true;
                    form.DialogResult = DialogResult.OK;
                    form.Close();
                };

                form.Controls.Add(lblPresetName);
                form.Controls.Add(txtPresetName);
                form.Controls.Add(lblSortOrder);
                form.Controls.Add(numSortOrder);
                form.Controls.Add(lblFuzz);
                form.Controls.Add(numFuzz);
                form.Controls.Add(tbFuzz);
                form.Controls.Add(chkDefault);
                form.Controls.Add(btnDelete);
                form.Controls.Add(btnOk);

                form.AcceptButton = btnOk;
                form.Resize += (_, __) => ApplyDialogLayout();
                ApplyDialogLayout();
                form.FormClosed += (_, __) =>
                {
                    presetEditorBounds = form.Bounds;
                    hasPresetEditorBounds = true;
                    presetEditorSaveAsDefaultChecked = chkDefault.Checked;
                    SaveImageReplacePresets();
                };

                if (form.ShowDialog() != DialogResult.OK)
                {
                    return false;
                }

                if (localDeleteRequested)
                {
                    deleteRequested = true;
                    return true;
                }

                var name = txtPresetName.Text?.Trim();
                if (string.IsNullOrWhiteSpace(name))
                {
                    MessageBox.Show("名称不能为空。", "BioDraw");
                    return false;
                }

                result = new ImageReplacePreset
                {
                    Name = name,
                    SortOrder = Convert.ToInt32(numSortOrder.Value, CultureInfo.InvariantCulture),
                    TargetColor = source.TargetColor,
                    Mode = source.Mode,
                    ReplacementColor = source.ReplacementColor,
                    FuzzPercent = NormalizeFuzzPercent(Convert.ToDouble(numFuzz.Value, CultureInfo.InvariantCulture))
                };
                setAsDefault = chkDefault.Checked;
                presetEditorSaveAsDefaultChecked = chkDefault.Checked;
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
                if (selection != null && TryGetSelectedShapes(selection, out shapes))
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

                TryReselectShapes(previousShapes);
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

        private static bool TryExecuteMso(dynamic app, IEnumerable<string> commandIds)
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

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        private static extern uint GetPixel(IntPtr hdc, int nXPos, int nYPos);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetIconInfo(IntPtr hIcon, out ICONINFO pIconInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CreateIconIndirect(ref ICONINFO icon);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool DeleteObject(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ICONINFO
        {
            [MarshalAs(UnmanagedType.Bool)]
            public bool fIcon;
            public int xHotspot;
            public int yHotspot;
            public IntPtr hbmMask;
            public IntPtr hbmColor;
        }

        private static bool IsVirtualKeyDown(int keyCode)
        {
            return (GetAsyncKeyState(keyCode) & 0x8000) != 0;
        }

        private sealed class ImageReplacePreset
        {
            public string Name { get; set; }
            public int SortOrder { get; set; }
            public double FuzzPercent { get; set; }
            public string TargetColor { get; set; }
            public string Mode { get; set; }
            public string ReplacementColor { get; set; }
        }

        private sealed class MaterialEntry
        {
            public string Name { get; set; }
            public string FilePath { get; set; }
        }

        private List<string> GetLevel1List()
        {
            if (!UseCustomMaterialLibrary())
            {
                return level1Items.Count > 0 ? level1Items : new List<string> { "默认" };
            }

            return GetSubDirectoryNames(materialLibraryPath);
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

                var level1 = level1List[NormalizeIndex(selectedLevel1Index, level1List.Count)];
                var level1Path = Path.Combine(materialLibraryPath, level1);
                return GetSubDirectoryNames(level1Path);
            }

            var fallbackLevel1List = GetLevel1List();
            var fallbackLevel1 = fallbackLevel1List[NormalizeIndex(selectedLevel1Index, fallbackLevel1List.Count)];
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

                var level1 = level1List[NormalizeIndex(selectedLevel1Index, level1List.Count)];
                var level2 = level2List[NormalizeIndex(selectedLevel2Index, level2List.Count)];
                var level2Path = Path.Combine(materialLibraryPath, level1, level2);
                return GetSubDirectoryNames(level2Path);
            }

            var fallbackLevel2List = GetLevel2List();
            var fallbackLevel2 = fallbackLevel2List[NormalizeIndex(selectedLevel2Index, fallbackLevel2List.Count)];
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

            var level1 = level1List[NormalizeIndex(selectedLevel1Index, level1List.Count)];
            var level2 = level2List[NormalizeIndex(selectedLevel2Index, level2List.Count)];
            var level2Path = Path.Combine(materialLibraryPath, level1, level2);
            return GetMaterialEntriesFromFolder(level2Path);
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
                    .Where(IsSupportedMaterialFile)
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

        private static List<MaterialEntry> GetMaterialEntriesFromFolder(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                return new List<MaterialEntry> { new MaterialEntry { Name = "默认", FilePath = string.Empty } };
            }

            try
            {
                var entries = Directory.GetFiles(folderPath, "*", SearchOption.TopDirectoryOnly)
                    .Where(IsSupportedMaterialFile)
                    .Select(path => new MaterialEntry
                    {
                        Name = Path.GetFileNameWithoutExtension(path),
                        FilePath = path
                    })
                    .OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
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
            if (!TryBuildMaterialThumbnail(filePath, entry.Name, MaterialThumbnailWidth, MaterialThumbnailHeight, out bitmap))
            {
                return transparentPlaceholderImage ?? brandImageLarge ?? brandImageSmall;
            }

            using (bitmap)
            {
                picture = PictureConverter.ToPictureDisp(new Bitmap(bitmap));
            }

            materialPreviewCache[filePath] = picture;
            return picture;
        }

        private static bool TryBuildMaterialThumbnail(string filePath, string label, int width, int height, out Bitmap bitmap)
        {
            bitmap = null;
            var safeWidth = Math.Max(24, width);
            var safeHeight = Math.Max(24, height);

            try
            {
                using (var image = Image.FromFile(filePath))
                {
                    bitmap = BuildThumbnailBitmap(image, safeWidth, safeHeight, label);
                    return bitmap != null;
                }
            }
            catch
            {
            }

            return TryBuildMaterialThumbnailByPowerPoint(filePath, label, safeWidth, safeHeight, out bitmap);
        }

        private static Bitmap BuildThumbnailBitmap(Image image, int width, int height, string label)
        {
            var bitmap = new Bitmap(width, height);
            var framePadding = 2f;
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                var frameRect = new RectangleF(0.5f, 0.5f, width - 1f, height - 1f);
                var mediaRect = new RectangleF(
                    framePadding,
                    framePadding,
                    width - (framePadding * 2f),
                    height - (framePadding * 2f));

                using (var pen = new Pen(Color.FromArgb(180, 180, 180), 1f))
                {
                    graphics.FillRectangle(Brushes.White, mediaRect);

                    var scale = Math.Min(mediaRect.Width / Math.Max(1f, image.Width), mediaRect.Height / Math.Max(1f, image.Height));
                    var drawWidth = Math.Max(1f, image.Width * scale);
                    var drawHeight = Math.Max(1f, image.Height * scale);
                    var x = mediaRect.Left + (mediaRect.Width - drawWidth) / 2f;
                    var y = mediaRect.Top + (mediaRect.Height - drawHeight) / 2f;
                    graphics.DrawImage(image, x, y, drawWidth, drawHeight);

                    graphics.DrawRectangle(pen, frameRect.X, frameRect.Y, frameRect.Width, frameRect.Height);
                }
            }

            return bitmap;
        }

        private static bool TryBuildMaterialThumbnailByPowerPoint(string filePath, string label, int width, int height, out Bitmap bitmap)
        {
            bitmap = null;
            dynamic shape = null;
            string tempPngPath = null;

            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null)
                {
                    return false;
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
                    return false;
                }

                shape = slide.Shapes.AddPicture(
                    filePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    -5000f,
                    -5000f,
                    -1f,
                    -1f);

                tempPngPath = Path.Combine(Path.GetTempPath(), "BioDraw", "material_thumb_" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture) + ".png");
                Directory.CreateDirectory(Path.GetDirectoryName(tempPngPath));
                shape.Export(tempPngPath, 2);

                using (var image = Image.FromFile(tempPngPath))
                {
                    bitmap = BuildThumbnailBitmap(image, width, height, label);
                }

                return bitmap != null;
            }
            catch
            {
                return false;
            }
            finally
            {
                try
                {
                    if (shape != null)
                    {
                        shape.Delete();
                    }
                }
                catch
                {
                }

                try
                {
                    if (!string.IsNullOrWhiteSpace(tempPngPath) && File.Exists(tempPngPath))
                    {
                        File.Delete(tempPngPath);
                    }
                }
                catch
                {
                }
            }
        }

        private bool TryInsertMaterialToCurrentSlide(string filePath, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                errorMessage = "素材文件不存在。";
                return false;
            }

            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null)
                {
                    errorMessage = "未能获取 PowerPoint 应用实例。";
                    return false;
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
                    errorMessage = "请先切换到普通编辑视图。";
                    return false;
                }

                dynamic newShape = slide.Shapes.AddPicture(
                    filePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    0f,
                    0f,
                    -1f,
                    -1f);

                var pageSetup = app.ActivePresentation?.PageSetup;
                if (pageSetup != null)
                {
                    float slideWidth = (float)pageSetup.SlideWidth;
                    float slideHeight = (float)pageSetup.SlideHeight;
                    newShape.Left = (slideWidth - (float)newShape.Width) / 2f;
                    newShape.Top = (slideHeight - (float)newShape.Height) / 2f;
                }

                if (string.Equals(Path.GetExtension(filePath), ".svg", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        newShape.Select();
                        TryExecuteMso(app, new[] { "SVGEdit" });
                    }
                    catch
                    {
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static bool IsSupportedMaterialFile(string filePath)
        {
            return materialFileExtensions.Contains(Path.GetExtension(filePath) ?? string.Empty);
        }

        private bool UseCustomMaterialLibrary()
        {
            return !string.IsNullOrWhiteSpace(materialLibraryPath) && Directory.Exists(materialLibraryPath);
        }

        private static List<string> GetSubDirectoryNames(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                return new List<string> { "默认" };
            }

            try
            {
                var names = Directory.GetDirectories(folderPath)
                    .Select(Path.GetFileName)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (names.Count > 0)
                {
                    return names;
                }
            }
            catch
            {
            }

            return new List<string> { "默认" };
        }

        private static int NormalizeIndex(int index, int count)
        {
            if (count <= 0)
            {
                return 0;
            }
            if (index < 0 || index >= count)
            {
                return 0;
            }
            return index;
        }
    }
}
