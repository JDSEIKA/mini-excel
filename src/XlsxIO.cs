using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace CalcRapido
{
    public static class XlsxIO
    {
        private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const long MaxFileBytes = 120L * 1024 * 1024;
        private const long MaxEntryBytes = 60L * 1024 * 1024;
        private const long MaxTotalUncompressedBytes = 256L * 1024 * 1024;
        private const long MaxCompressionRatio = 500;
        private const long MaxImageBytes = 20L * 1024 * 1024;
        private const long MaxImagePixels = 25000000;
        private const int MaxImageDimension = 12000;
        private const int MaxXmlCharacters = 64 * 1024 * 1024;
        private const int MaxZipEntries = 10000;
        private const int MaxImages = 2000;
        private const int MaxCharts = 500;
        private const int MaxShapes = 2000;
        private const int MaxSlicers = 500;
        private const int MaxComments = 100000;
        private const int MaxSheets = 512;
        private const int MaxCellsPerSheet = 2000000;
        private const int MaxCellsTotal = 4000000;
        private const int MaxSharedStrings = 1000000;
        private const int MaxStringChars = 32767;
        private static long _readBudgetRemaining = -1;

        // ---------------------------------------------------------------- LOAD

        public static Workbook Load(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (fi.Length > MaxFileBytes) throw new FileFormatException("El archivo .xlsx es demasiado grande.");
            _readBudgetRemaining = MaxTotalUncompressedBytes;
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    CheckZipBudget(zip);
                    string workbookPath = FindWorkbookPath(zip);
                    XDocument wbDoc = ReadXml(zip, workbookPath);
                    string wbDir = PartDirectory(workbookPath);
                    string relPath = PartDirectory(workbookPath) + "_rels/" + Path.GetFileName(workbookPath) + ".rels";
                    XDocument relDoc = ReadXmlIfExists(zip, relPath);
                    Dictionary<string, string> rels = ParseRelationships(relDoc);
                    string sharedPath = ResolvePart(wbDir, RelTargetByType(relDoc, "/sharedStrings") ?? "sharedStrings.xml");
                    string stylesPath = ResolvePart(wbDir, RelTargetByType(relDoc, "/styles") ?? "styles.xml");
                    string themePath = ResolvePart(wbDir, RelTargetByType(relDoc, "/theme") ?? "theme/theme1.xml");
                    XDocument sharedDoc = ReadXmlIfExists(zip, sharedPath);
                    List<string> shared = ParseSharedStrings(sharedDoc);
                    ThemeColors theme = ParseTheme(ReadXmlIfExists(zip, themePath));
                    Color[] themePalette = theme.AccentPalette();
                    ChartData.ThemePalette = themePalette;
                    XDocument stylesDoc = ReadXmlIfExists(zip, stylesPath);
                    List<CellStyle> styles = ParseStyles(stylesDoc, theme);
                    List<CellStyle> dxfs = ParseDxfs(stylesDoc, theme);
                    Dictionary<string, TableStyleDefinition> tableStyles =
                        ParseTableStyles(stylesDoc, dxfs);
                    Workbook wb = new Workbook();
                    wb.Sheets.Clear();
                    wb.ThemePalette = themePalette;
                    AddPackageWarnings(zip, wb);
                    wb.Date1904 = IsTrue(Attr(FirstDescendant(wbDoc.Root, "workbookPr"), "date1904"));
                    ParseDefinedNames(wbDoc, wb);
                    int activeTab = 0;
                    XElement bookView = FirstDescendant(wbDoc.Root, "workbookView");
                    if (bookView != null)
                    {
                        int a;
                        if (int.TryParse(Attr(bookView, "activeTab"), out a)) activeTab = a;
                    }
                    int sheetIndex = 0;
                    int totalCells = 0;
                    List<TableInfo> tables = new List<TableInfo>();
                    foreach (XElement sheetEl in Descendants(wbDoc.Root, "sheet"))
                    {
                        if (sheetIndex >= MaxSheets)
                            throw new FileFormatException("El archivo .xlsx contiene demasiadas hojas.");
                        string name = Attr(sheetEl, "name");
                        if (string.IsNullOrEmpty(name)) name = "Hoja" + (sheetIndex + 1);
                        string rid = RelAttr(sheetEl, "id");
                        string target = null;
                        if (!string.IsNullOrEmpty(rid)) rels.TryGetValue(rid, out target);
                        if (string.IsNullOrEmpty(target))
                        {
                            target = "xl/worksheets/sheet" + (sheetIndex + 1) + ".xml";
                        }
                        string part = ResolvePart(wbDir, target);
                        XDocument sheetDoc = ReadXmlIfExists(zip, part);
                        Sheet sheet = new Sheet(wb, MakeUniqueName(wb, name));
                        string srelPath = PartDirectory(part) + "_rels/" + Path.GetFileName(part) + ".rels";
                        Dictionary<string, string> sheetRels = ParseRelationships(ReadXmlIfExists(zip, srelPath), true);
                        sheet.SourcePart = part;
                        if (sheetDoc != null) ParseSheet(sheetDoc, sheet, shared, styles, dxfs, theme, sheetRels);
                        if (sheetDoc != null) ParseSheetTables(zip, sheet, sheetDoc, sheetRels, tables, tableStyles);
                        if (sheetDoc != null) ParseSheetCharts(zip, sheet, sheetDoc, sheetRels);
                        ParseComments(zip, sheet, sheetRels);
                        ParsePivotRanges(zip, sheet, sheetRels);
                        sheet.RecalculateUsedRange();
                        totalCells += sheet.Cells.Count;
                        if (totalCells > MaxCellsTotal)
                            throw new FileFormatException("El archivo .xlsx contiene demasiadas celdas en total.");
                        wb.Sheets.Add(sheet);
                        sheetIndex++;
                    }
                    if (wb.Sheets.Count == 0) wb.AddSheet("Hoja1");
                    wb.ActiveSheetIndex = Math.Max(0, Math.Min(wb.Sheets.Count - 1, activeTab));
                    wb.XlsxSource = new XlsxSourceInfo();
                    wb.XlsxSource.Path = Path.GetFullPath(path);
                    wb.XlsxSource.Length = fi.Length;
                    wb.XlsxSource.LastWriteUtc = fi.LastWriteTimeUtc;
                    wb.XlsxSource.WorkbookPart = workbookPath;
                    wb.XlsxSource.StylesPart = stylesPath;
                    wb.XlsxSource.ThemePart = themePath;
                    wb.XlsxSource.SharedStringsPart = sharedPath;
                    wb.XlsxSource.Date1904 = wb.Date1904;
                    if (wb.Date1904) AdjustDate1904(wb, 1462);
                    TranslateStructuredFormulas(wb, tables);
                    foreach (DefinedName dn in wb.DefinedNames)
                    {
                        if (!string.Equals(dn.Name, "_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(dn.Name, "Print_Area", StringComparison.OrdinalIgnoreCase)) continue;
                        int index = dn.LocalSheetId;
                        if (index < 0 || index >= wb.Sheets.Count) continue;
                        wb.Sheets[index].PrintArea = dn.Formula;
                    }
                    foreach (DefinedName dn in wb.DefinedNames)
                    {
                        if (!string.Equals(dn.Name, "_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(dn.Name, "Print_Titles", StringComparison.OrdinalIgnoreCase)) continue;
                        int index = dn.LocalSheetId;
                        if (index < 0 || index >= wb.Sheets.Count) continue;
                        wb.Sheets[index].PrintTitleRows = dn.Formula;
                        wb.Sheets[index].PrintTitleCols = dn.Formula;
                    }
                    int loadedCells = 0;
                    foreach (Sheet sh in wb.Sheets) loadedCells += sh.Cells.Count;
                    if (loadedCells > 500000)
                        wb.AddLoadWarning("Libro grande: " + loadedCells.ToString("N0", CultureInfo.CurrentCulture) +
                            " celdas. El uso de memoria puede superar 1 GB.");
                    wb.Recalculate();
                    _readBudgetRemaining = -1;
                    return wb;
                }
            }
            catch (FileFormatException) { throw; }
            catch (InvalidDataException) { throw new FileFormatException("El archivo .xlsx no es un paquete válido."); }
            catch (Exception ex) { throw new FileFormatException("No se pudo abrir el archivo .xlsx: " + ex.Message); }
        }

        private static string FindWorkbookPath(ZipArchive zip)
        {
            XDocument rootRels = ReadXmlIfExists(zip, "_rels/.rels");
            if (rootRels != null)
            {
                foreach (XElement rel in Descendants(rootRels.Root, "Relationship"))
                {
                    string type = Attr(rel, "Type");
                    if (type != null && type.EndsWith("/officeDocument", StringComparison.OrdinalIgnoreCase))
                    {
                        string target = Attr(rel, "Target");
                        if (!string.IsNullOrEmpty(target))
                        {
                            string p = target.Replace('\\', '/').TrimStart('/');
                            if (p == ".." || p.StartsWith("../", StringComparison.Ordinal) ||
                                p.Contains("/../"))
                                throw new FileFormatException("Ruta interna no válida en el .xlsx.");
                            return p;
                        }
                    }
                }
            }
            return "xl/workbook.xml";
        }

        private static string RelTargetByType(XDocument relDoc, string suffix)
        {
            if (relDoc == null || relDoc.Root == null) return null;
            foreach (XElement rel in Descendants(relDoc.Root, "Relationship"))
            {
                string type = Attr(rel, "Type");
                if (type != null && type.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    string target = Attr(rel, "Target");
                    if (!string.IsNullOrEmpty(target)) return target;
                }
            }
            return null;
        }

        private static void ParseDefinedNames(XDocument wbDoc, Workbook wb)
        {
            if (wbDoc == null || wbDoc.Root == null) return;
            foreach (XElement dn in Descendants(wbDoc.Root, "definedName"))
            {
                if (wb.DefinedNames.Count >= 100000)
                    throw new FileFormatException("El .xlsx contiene demasiados nombres definidos.");
                string name = Attr(dn, "name");
                if (string.IsNullOrEmpty(name)) continue;
                DefinedName d = new DefinedName();
                d.Name = name;
                d.Formula = (dn.Value ?? "").Trim();
                if (d.Formula.StartsWith("=", StringComparison.Ordinal)) d.Formula = d.Formula.Substring(1);
                int local;
                if (int.TryParse(Attr(dn, "localSheetId"), out local)) d.LocalSheetId = local;
                d.Hidden = IsTrue(Attr(dn, "hidden"));
                d.Comment = Attr(dn, "comment");
                wb.DefinedNames.Add(d);
            }
        }

        private static bool IsTrue(string s)
        {
            return s == "1" || string.Equals(s, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static void AddPackageWarnings(ZipArchive zip, Workbook wb)
        {
            int charts = 0, images = 0, tables = 0, pivots = 0, externalLinks = 0;
            bool macros = false;
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                string n = entry.FullName.Replace('\\', '/');
                if (n.StartsWith("xl/charts/", StringComparison.OrdinalIgnoreCase) &&
                    n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) charts++;
                else if (n.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase) &&
                         !n.EndsWith("/", StringComparison.Ordinal)) images++;
                else if (n.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase) &&
                         n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) tables++;
                else if (n.StartsWith("xl/pivotCache/", StringComparison.OrdinalIgnoreCase)) pivots++;
                else if (n.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase)) macros = true;
                else if (n.StartsWith("xl/externalLinks/", StringComparison.OrdinalIgnoreCase) &&
                         n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) externalLinks++;
            }
            if (charts > 0)
                wb.AddLoadWarning("El libro contiene " + charts +
                    " gráfico(s). Se muestran los tipos básicos; los gráficos no compatibles se conservan al guardar.");
            if (images > 0)
                wb.AddLoadWarning("El libro contiene " + images +
                    " imagen(es). Se muestran y se pueden mover, recortar o eliminar.");
            if (tables > 0)
                wb.AddLoadWarning("El libro contiene tablas. Se conservan al guardar, pero no se editan como tabla.");
            if (pivots > 0)
                wb.AddLoadWarning("El libro contiene tablas dinámicas. No se recalculan ni se editan.");
            if (macros)
                wb.AddLoadWarning("El libro contiene macros. Mini Excel no las ejecuta ni las edita.");
            if (externalLinks > 0)
                wb.AddLoadWarning("El libro contiene vínculos externos. Se conservan, pero no se actualizan.");
        }

        private static void CheckZipBudget(ZipArchive zip)
        {
            if (zip.Entries.Count > MaxZipEntries)
                throw new FileFormatException("El archivo .xlsx contiene demasiadas partes internas.");
            long total = 0;
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                string entryName = entry.FullName.Replace('\\', '/');
                if (entryName.StartsWith("/", StringComparison.Ordinal) ||
                    entryName.IndexOf("../", StringComparison.Ordinal) >= 0)
                    throw new FileFormatException("El paquete .xlsx contiene una ruta no válida.");
                if (entry.Length < 0 || entry.Length > MaxEntryBytes)
                    throw new FileFormatException("Una parte interna del .xlsx es demasiado grande.");
                if (entry.Length > 1024 * 1024 &&
                    entry.CompressedLength > 0 &&
                    entry.Length / Math.Max(1, entry.CompressedLength) > MaxCompressionRatio)
                    throw new FileFormatException("Una parte del .xlsx tiene una compresión sospechosa.");
                if (entry.Length > 1024 * 1024 && entry.CompressedLength == 0)
                    throw new FileFormatException("Una parte del .xlsx tiene una compresión no válida.");
                try { total = checked(total + entry.Length); }
                catch (OverflowException) { throw new FileFormatException("El paquete .xlsx es demasiado grande."); }
                if (total > MaxTotalUncompressedBytes)
                    throw new FileFormatException("El contenido descomprimido del .xlsx es demasiado grande.");
            }
        }

        private static XDocument ReadXml(ZipArchive zip, string path)
        {
            XDocument doc = ReadXmlIfExists(zip, path);
            if (doc == null) throw new FileFormatException("Falta una parte interna del archivo .xlsx: " + path);
            return doc;
        }

        private static XDocument ReadXmlIfExists(ZipArchive zip, string path)
        {
            ZipArchiveEntry entry = GetEntry(zip, path);
            if (entry == null) return null;
            if (entry.Length > MaxEntryBytes) throw new FileFormatException("Una parte interna del .xlsx es demasiado grande.");
            try
            {
                byte[] bytes = ReadEntryBytes(entry, MaxEntryBytes);
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Prohibit;
                settings.XmlResolver = null;
                settings.MaxCharactersInDocument = MaxXmlCharacters;
                settings.MaxCharactersFromEntities = 0;
                using (MemoryStream s = new MemoryStream(bytes, false))
                using (XmlReader reader = XmlReader.Create(s, settings))
                    return XDocument.Load(reader, LoadOptions.None);
            }
            catch (XmlException ex)
            {
                throw new FileFormatException("XML dañado dentro del .xlsx: " + ex.Message);
            }
        }

        private static ZipArchiveEntry GetEntry(ZipArchive zip, string path)
        {
            string p = path.Replace('\\', '/').TrimStart('/');
            ZipArchiveEntry e = zip.GetEntry(p);
            if (e != null) return e;
            foreach (ZipArchiveEntry entry in zip.Entries)
                if (string.Equals(entry.FullName.Replace('\\', '/'), p, StringComparison.OrdinalIgnoreCase)) return entry;
            return null;
        }

        private static Dictionary<string, string> ParseRelationships(XDocument doc)
        {
            return ParseRelationships(doc, false);
        }

        private static Dictionary<string, string> ParseRelationships(XDocument doc, bool includeExternal)
        {
            Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (doc == null || doc.Root == null) return d;
            foreach (XElement rel in Descendants(doc.Root, "Relationship"))
            {
                string id = Attr(rel, "Id");
                string target = Attr(rel, "Target");
                string mode = Attr(rel, "TargetMode");
                if (id != null && target != null &&
                    (includeExternal || !string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase)))
                    d[id] = target;
            }
            return d;
        }

        private static List<string> ParseSharedStrings(XDocument doc)
        {
            List<string> list = new List<string>();
            if (doc == null || doc.Root == null) return list;
            foreach (XElement si in Descendants(doc.Root, "si"))
            {
                if (list.Count >= MaxSharedStrings)
                    throw new FileFormatException("El archivo .xlsx contiene demasiadas cadenas.");
                StringBuilder sb = new StringBuilder();
                foreach (XElement t in Descendants(si, "t"))
                {
                    if (t.Value == null) continue;
                    int take = Math.Min(t.Value.Length, MaxStringChars - sb.Length);
                    if (take > 0) sb.Append(t.Value, 0, take);
                    if (sb.Length >= MaxStringChars) break;
                }
                list.Add(sb.ToString());
            }
            return list;
        }

        private sealed class ThemeColors
        {
            public readonly Color[] Colors = new Color[12];

            public ThemeColors()
            {
                // Índices usados por el atributo theme de styles.xml:
                // 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9=accent1..6, 10=hlink, 11=folHlink.
                Colors[0] = Color.White;
                Colors[1] = Color.Black;
                Colors[2] = Color.FromArgb(238, 236, 225);
                Colors[3] = Color.FromArgb(31, 73, 125);
                Colors[4] = Color.FromArgb(79, 129, 189);
                Colors[5] = Color.FromArgb(192, 80, 77);
                Colors[6] = Color.FromArgb(155, 187, 89);
                Colors[7] = Color.FromArgb(128, 100, 162);
                Colors[8] = Color.FromArgb(75, 172, 198);
                Colors[9] = Color.FromArgb(247, 150, 70);
                Colors[10] = Color.Blue;
                Colors[11] = Color.Purple;
            }

            public Color Resolve(int index, double tint)
            {
                if (index < 0 || index >= Colors.Length) return Color.Black;
                Color c = Colors[index];
                if (Math.Abs(tint) < 0.0001) return c;
                return ApplyTint(c, tint);
            }

            public Color[] AccentPalette()
            {
                return new Color[]
                {
                    Colors[4], Colors[5], Colors[6], Colors[7], Colors[8], Colors[9]
                };
            }
        }

        private static ThemeColors ParseTheme(XDocument doc)
        {
            ThemeColors theme = new ThemeColors();
            if (doc == null || doc.Root == null) return theme;
            XElement scheme = FirstDescendant(doc.Root, "clrScheme");
            if (scheme == null) return theme;
            string[] xmlOrder = new string[] { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };
            int[] styleMap = new int[] { 1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11 };
            for (int i = 0; i < xmlOrder.Length; i++)
            {
                XElement slot = FirstDescendant(scheme, xmlOrder[i]);
                if (slot == null) continue;
                XElement c = null;
                foreach (XElement child in slot.Elements()) { c = child; break; }
                if (c == null) continue;
                string rgb = Attr(c, "val");
                if (rgb == null) rgb = Attr(c, "lastClr");
                Color color = ParseColor(rgb);
                if (!color.IsEmpty) theme.Colors[styleMap[i]] = color;
            }
            return theme;
        }

        private static readonly int[] IndexedPalette = new int[]
        {
            0x000000,0xFFFFFF,0xFF0000,0x00FF00,0x0000FF,0xFFFF00,0xFF00FF,0x00FFFF,
            0x000000,0xFFFFFF,0xFF0000,0x00FF00,0x0000FF,0xFFFF00,0xFF00FF,0x00FFFF,
            0x800000,0x008000,0x000080,0x808000,0x800080,0x008080,0xC0C0C0,0x808080,
            0x9999FF,0x993366,0xFFFFCC,0xCCFFFF,0x660066,0xFF8080,0x0066CC,0xCCCCFF,
            0x000080,0xFF00FF,0xFFFF00,0x00FFFF,0x800080,0x800000,0x008080,0x0000FF,
            0x00CCFF,0xCCFFFF,0xCCFFCC,0xFFFF99,0x99CCFF,0xFF99CC,0xCC99FF,0xFFCC99,
            0x3366FF,0x33CCCC,0x99CC00,0xFFCC00,0xFF9900,0xFF6600,0x666699,0x969696,
            0x003366,0x339966,0x003300,0x333300,0x993300,0x993366,0x333399,0x333333
        };

        private static Color ApplyTint(Color c, double tint)
        {
            int a = c.A;
            int r = TintChannel(c.R, tint);
            int g = TintChannel(c.G, tint);
            int b = TintChannel(c.B, tint);
            return Color.FromArgb(a, r, g, b);
        }

        private static int TintChannel(int channel, double tint)
        {
            double d = channel / 255.0;
            if (tint < 0) d = d * (1.0 + tint);
            else d = d + (1.0 - d) * tint;
            return Math.Max(0, Math.Min(255, (int)Math.Round(d * 255.0)));
        }

        private static List<CellStyle> ParseStyles(XDocument doc, ThemeColors theme)
        {
            List<CellStyle> result = new List<CellStyle>();
            if (doc == null || doc.Root == null) return result;

            Dictionary<int, string> customFormats = new Dictionary<int, string>();
            XElement numFmts = FirstDescendant(doc.Root, "numFmts");
            if (numFmts != null)
                foreach (XElement nf in Descendants(numFmts, "numFmt"))
                {
                    int id;
                    if (int.TryParse(Attr(nf, "numFmtId"), out id)) customFormats[id] = Attr(nf, "formatCode") ?? "";
                }

            List<XElement> fonts = new List<XElement>();
            XElement fontsEl = FirstDescendant(doc.Root, "fonts");
            if (fontsEl != null) foreach (XElement f in Descendants(fontsEl, "font")) fonts.Add(f);
            List<XElement> fills = new List<XElement>();
            XElement fillsEl = FirstDescendant(doc.Root, "fills");
            if (fillsEl != null) foreach (XElement f in Descendants(fillsEl, "fill")) fills.Add(f);
            List<XElement> borders = new List<XElement>();
            XElement bordersEl = FirstDescendant(doc.Root, "borders");
            if (bordersEl != null) foreach (XElement b in Descendants(bordersEl, "border")) borders.Add(b);

            XElement cellXfs = FirstDescendant(doc.Root, "cellXfs");
            if (cellXfs == null) return result;
            foreach (XElement xf in Descendants(cellXfs, "xf"))
            {
                CellStyle st = new CellStyle();
                int numFmtId = ParseInt(Attr(xf, "numFmtId"));
                int fontId = ParseInt(Attr(xf, "fontId"));
                int fillId = ParseInt(Attr(xf, "fillId"));
                int borderId = ParseInt(Attr(xf, "borderId"));
                string code;
                if (!customFormats.TryGetValue(numFmtId, out code)) code = BuiltInFormat(numFmtId);
                ApplyNumberFormat(st, code);
                if (fontId >= 0 && fontId < fonts.Count) ApplyFont(st, fonts[fontId], theme);
                if (fillId >= 0 && fillId < fills.Count) ApplyFill(st, fills[fillId], theme);
                if (borderId >= 0 && borderId < borders.Count) ApplyBorder(st, borders[borderId]);
                XElement align = FirstDescendant(xf, "alignment");
                if (align != null)
                {
                    string h = Attr(align, "horizontal");
                    if (string.Equals(h, "left", StringComparison.OrdinalIgnoreCase)) st.Align = 0;
                    else if (string.Equals(h, "center", StringComparison.OrdinalIgnoreCase)) st.Align = 1;
                    else if (string.Equals(h, "right", StringComparison.OrdinalIgnoreCase)) st.Align = 2;
                    st.WrapText = string.Equals(Attr(align, "wrapText"), "1", StringComparison.OrdinalIgnoreCase) ||
                                  string.Equals(Attr(align, "wrapText"), "true", StringComparison.OrdinalIgnoreCase);
                }
                result.Add(st);
            }
            return result;
        }

        private static List<CellStyle> ParseDxfs(XDocument doc, ThemeColors theme)
        {
            List<CellStyle> result = new List<CellStyle>();
            if (doc == null || doc.Root == null) return result;
            XElement box = FirstDescendant(doc.Root, "dxfs");
            if (box == null) return result;
            foreach (XElement dxf in Descendants(box, "dxf"))
            {
                CellStyle st = new CellStyle();
                XElement font = FirstDescendant(dxf, "font");
                if (font != null) ApplyFont(st, font, theme);
                XElement fill = FirstDescendant(dxf, "fill");
                if (fill != null)
                {
                    ApplyFill(st, fill, theme);
                    if (st.FillColor.IsEmpty)
                    {
                        XElement bg = FirstDescendant(fill, "bgColor");
                        if (bg == null) bg = FirstDescendant(fill, "fgColor");
                        Color fillColor = ResolveColor(bg, theme);
                        if (!fillColor.IsEmpty) st.FillColor = fillColor;
                    }
                }
                XElement border = FirstDescendant(dxf, "border");
                if (border != null) ApplyBorder(st, border);
                result.Add(st);
            }
            return result;
        }

        private static Dictionary<string, TableStyleDefinition> ParseTableStyles(
            XDocument doc, List<CellStyle> dxfs)
        {
            Dictionary<string, TableStyleDefinition> result =
                new Dictionary<string, TableStyleDefinition>(StringComparer.OrdinalIgnoreCase);
            if (doc == null || doc.Root == null) return result;
            XElement box = FirstDescendant(doc.Root, "tableStyles");
            if (box == null) return result;
            foreach (XElement ts in Descendants(box, "tableStyle"))
            {
                string name = Attr(ts, "name");
                if (string.IsNullOrEmpty(name)) continue;
                TableStyleDefinition def = new TableStyleDefinition();
                foreach (XElement el in Descendants(ts, "tableStyleElement"))
                {
                    int id;
                    if (!int.TryParse(Attr(el, "dxfId"), out id) || id < 0 || id >= dxfs.Count) continue;
                    string type = (Attr(el, "type") ?? "").ToLowerInvariant();
                    CellStyle style = dxfs[id].Clone();
                    if (type == "wholetable") def.WholeTable = style;
                    else if (type == "headerrow") def.HeaderRow = style;
                    else if (type == "totalrow") def.TotalRow = style;
                    else if (type == "firstrowstripe") def.FirstRowStripe = style;
                    else if (type == "firstcolumn") def.FirstColumn = style;
                    else if (type == "lastcolumn") def.LastColumn = style;
                    else if (type == "lastheadercell") def.LastHeaderCell = style;
                    else if (type == "lasttotalcell") def.LastTotalCell = style;
                }
                result[name] = def;
            }
            return result;
        }

        internal static void ApplyNumberFormat(CellStyle st, string code)
        {
            if (string.IsNullOrEmpty(code)) { st.Format = NumberFormatKind.General; return; }
            string rawCode = code;
            string classification = StripBrackets(code);
            string lower = classification.ToLowerInvariant();
            if (lower == "general" || lower.IndexOf("general", StringComparison.Ordinal) >= 0)
            {
                st.Format = NumberFormatKind.General;
                st.CustomFormat = null;
                return;
            }
            st.CustomFormat = rawCode;
            bool hasNumericPlaceholder = lower.IndexOf('0') >= 0 || lower.IndexOf('#') >= 0;
            bool hasDate = lower.IndexOf("yy", StringComparison.Ordinal) >= 0 ||
                           lower.IndexOf("dd", StringComparison.Ordinal) >= 0 ||
                           lower.IndexOf("mmm", StringComparison.Ordinal) >= 0 ||
                           lower.IndexOf("d/m", StringComparison.Ordinal) >= 0 ||
                           lower.IndexOf("m/d", StringComparison.Ordinal) >= 0 ||
                           lower.Trim() == "d" ||
                           lower.IndexOf("d ", StringComparison.Ordinal) >= 0;
            bool hasTime = lower.IndexOf('h') >= 0 || lower.IndexOf("ss", StringComparison.Ordinal) >= 0;
            // Un formato como "d/m/yy;@" contiene @ para la sección de texto,
            // pero sigue siendo una fecha. El @ solo decide Texto cuando no
            // hay ningún componente de fecha u hora.
            if (lower.IndexOf('@') >= 0 && !hasNumericPlaceholder && !hasDate && !hasTime)
            {
                st.Format = NumberFormatKind.Text;
                return;
            }
            int decimals = CountDecimals(classification);
            st.Decimals = decimals;
            if (lower.IndexOf('%') >= 0) st.Format = NumberFormatKind.Percent;
            else if (lower.IndexOf('$') >= 0 || lower.IndexOf('\u20AC') >= 0 || lower.IndexOf('\u00A3') >= 0)
                st.Format = NumberFormatKind.Currency;
            else
            {
                if (hasDate) st.Format = NumberFormatKind.Date;
                else if (hasTime) st.Format = NumberFormatKind.Time;
                else if (lower.IndexOf('0') >= 0 || lower.IndexOf('#') >= 0) st.Format = NumberFormatKind.Number;
                else st.Format = NumberFormatKind.General;
            }
            bool complex = rawCode.IndexOf('"') >= 0 || rawCode.IndexOf(';') >= 0 ||
                           rawCode.IndexOf('\\') >= 0 || rawCode.IndexOf('_') >= 0 ||
                           rawCode.IndexOf('*') >= 0 || rawCode.IndexOf('[') >= 0 ||
                           rawCode.IndexOf("E+", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           rawCode.IndexOf("E-", StringComparison.OrdinalIgnoreCase) >= 0;
            if (st.Format == NumberFormatKind.Date && lower.IndexOf('h') >= 0) complex = true;
            if (complex && st.Format != NumberFormatKind.Text) st.Format = NumberFormatKind.Custom;
        }

        private static string StripBrackets(string code)
        {
            StringBuilder sb = new StringBuilder();
            int depth = 0;
            foreach (char c in code)
            {
                if (c == '[') { depth++; continue; }
                if (c == ']' && depth > 0) { depth--; continue; }
                if (depth == 0) sb.Append(c);
            }
            return sb.ToString();
        }

        private static int CountDecimals(string code)
        {
            int dot = code.IndexOf('.');
            if (dot < 0) return 0;
            int n = 0;
            for (int i = dot + 1; i < code.Length; i++)
            {
                if (code[i] == '0' || code[i] == '#') n++;
                else if (code[i] == ' ') continue;
                else break;
            }
            return Math.Max(0, Math.Min(10, n));
        }

        internal static string BuiltInFormat(int id)
        {
            switch (id)
            {
                case 1: return "0";
                case 2: return "0.00";
                case 3: return "#,##0";
                case 4: return "#,##0.00";
                case 9: return "0%";
                case 10: return "0.00%";
                case 11: return "0.00E+00";
                case 14: return "dd/mm/yyyy";
                case 15: return "d-mmm-yy";
                case 16: return "d-mmm";
                case 17: return "mmm-yy";
                case 18: return "h:mm AM/PM";
                case 19: return "h:mm:ss AM/PM";
                case 20: return "h:mm";
                case 21: return "h:mm:ss";
                case 22: return "dd/mm/yyyy h:mm";
                case 45: return "mm:ss";
                case 46: return "[h]:mm:ss";
                case 47: return "mm:ss.0";
                case 49: return "@";
                default: return "";
            }
        }

        private static void ApplyFont(CellStyle st, XElement font, ThemeColors theme)
        {
            if (font == null) return;
            st.Bold = HasOn(font, "b");
            st.Italic = HasOn(font, "i");
            st.Underline = HasOn(font, "u");
            XElement sz = FirstDescendant(font, "sz");
            if (sz != null)
            {
                float f;
                if (float.TryParse(Attr(sz, "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out f) && f >= 6 && f <= 72)
                    st.FontSize = f;
            }
            XElement name = FirstDescendant(font, "name");
            if (name != null && !string.IsNullOrEmpty(Attr(name, "val"))) st.FontName = Attr(name, "val");
            XElement color = FirstDescendant(font, "color");
            if (color != null)
            {
                Color c = ResolveColor(color, theme);
                if (!c.IsEmpty) st.FontColor = c;
            }
        }

        private static void ApplyFill(CellStyle st, XElement fill, ThemeColors theme)
        {
            if (fill == null) return;
            XElement pattern = FirstDescendant(fill, "patternFill");
            if (pattern == null) return;
            string type = Attr(pattern, "patternType");
            if (string.IsNullOrEmpty(type) || string.Equals(type, "none", StringComparison.OrdinalIgnoreCase)) return;
            if (string.Equals(type, "gray125", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(type, "gray0625", StringComparison.OrdinalIgnoreCase)) return;
            XElement fg = FirstDescendant(pattern, "fgColor");
            if (fg == null) fg = FirstDescendant(pattern, "bgColor");
            if (fg != null)
            {
                Color c = ResolveColor(fg, theme);
                if (!c.IsEmpty) st.FillColor = c;
            }
        }

        private static void ApplyBorder(CellStyle st, XElement border)
        {
            if (border == null) return;
            if (HasBorderSide(border, "left")) st.Borders |= BorderFlags.Left;
            if (HasBorderSide(border, "top")) st.Borders |= BorderFlags.Top;
            if (HasBorderSide(border, "right")) st.Borders |= BorderFlags.Right;
            if (HasBorderSide(border, "bottom")) st.Borders |= BorderFlags.Bottom;
        }

        private static bool HasBorderSide(XElement border, string name)
        {
            XElement side = FirstDescendant(border, name);
            if (side == null) return false;
            string style = Attr(side, "style");
            return !string.IsNullOrEmpty(style) && !string.Equals(style, "none", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasOn(XElement parent, string name)
        {
            XElement el = FirstDescendant(parent, name);
            if (el == null) return false;
            string val = Attr(el, "val");
            return val == null || val == "1" || string.Equals(val, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static Color ParseColor(string rgb)
        {
            if (string.IsNullOrEmpty(rgb) || rgb.Length < 6) return Color.Empty;
            string s = rgb.Length == 8 ? rgb : "FF" + rgb;
            uint v;
            if (!uint.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v)) return Color.Empty;
            return Color.FromArgb((int)v);
        }

        private static Color ResolveColor(XElement el, ThemeColors theme)
        {
            if (el == null) return Color.Empty;
            string rgb = Attr(el, "rgb");
            Color c = ParseColor(rgb);
            if (!c.IsEmpty) return c;
            string indexed = Attr(el, "indexed");
            int idx;
            if (!string.IsNullOrEmpty(indexed) && int.TryParse(indexed, out idx) &&
                idx >= 0 && idx < IndexedPalette.Length)
                return Color.FromArgb(unchecked((int)0xFF000000) | IndexedPalette[idx]);
            string themeIndex = Attr(el, "theme");
            if (!string.IsNullOrEmpty(themeIndex) && theme != null)
            {
                int ti;
                double tint = 0;
                double.TryParse(Attr(el, "tint"), NumberStyles.Float, CultureInfo.InvariantCulture, out tint);
                if (int.TryParse(themeIndex, out ti)) return theme.Resolve(ti, tint);
            }
            return Color.Empty;
        }

        private static void ParseSheet(XDocument doc, Sheet sheet, List<string> shared, List<CellStyle> styles,
                                       List<CellStyle> dxfs, ThemeColors theme,
                                       Dictionary<string, string> sheetRels)
        {
            if (doc == null || doc.Root == null) return;

            // Anchos de columna
            XElement cols = FirstDescendant(doc.Root, "cols");
            if (cols != null)
            {
                foreach (XElement col in Descendants(cols, "col"))
                {
                    int min = ParseInt(Attr(col, "min")) - 1;
                    int max = ParseInt(Attr(col, "max")) - 1;
                    float width;
                    if (!float.TryParse(Attr(col, "width"), NumberStyles.Float, CultureInfo.InvariantCulture, out width)) continue;
                    if (min < 0) min = 0;
                    if (max < min) max = min;
                    int w = (int)Math.Round(width * 7.5f);
                    w = Math.Max(28, Math.Min(800, w));
                    bool hidden = IsTrue(Attr(col, "hidden"));
                    for (int c = min; c <= max && c < Limits.MaxCols; c++)
                    {
                        sheet.SetColWidthImported(c, w);
                        sheet.SetColumnHiddenImported(c, hidden);
                    }
                }
            }

            List<XlsxCell> records = new List<XlsxCell>();
            int rowIndex = 0;
            foreach (XElement rowEl in Descendants(doc.Root, "row"))
            {
                int rAttr;
                if (int.TryParse(Attr(rowEl, "r"), out rAttr) && rAttr > 0) rowIndex = rAttr - 1;
                if (rowIndex < 0 || rowIndex >= Limits.MaxRows) { rowIndex++; continue; }
                float rowHeight;
                if (float.TryParse(Attr(rowEl, "ht"), NumberStyles.Float, CultureInfo.InvariantCulture, out rowHeight) && rowHeight > 0)
                    sheet.SetRowHeightPointsImported(rowIndex, rowHeight);
                if (IsTrue(Attr(rowEl, "hidden"))) sheet.SetRowHiddenImported(rowIndex, true);
                int colIndex = 0;
                foreach (XElement cEl in Descendants(rowEl, "c"))
                {
                    XlsxCell rec = new XlsxCell();
                    rec.Row = rowIndex;
                    rec.Col = colIndex;
                    string r = Attr(cEl, "r");
                    int rr, cc; bool ar, ac;
                    if (!string.IsNullOrEmpty(r) && AddressUtil.TryParseCell(r, out rr, out cc, out ar, out ac))
                    {
                        rec.Row = rr; rec.Col = cc;
                    }
                    colIndex = rec.Col + 1;
                    string t = Attr(cEl, "t");
                    rec.Type = t;
                    string sAttr = Attr(cEl, "s");
                    int sIdx;
                    if (!string.IsNullOrEmpty(sAttr) && int.TryParse(sAttr, out sIdx)) rec.StyleIndex = sIdx;
                    XElement fEl = FirstDescendant(cEl, "f");
                    if (fEl != null)
                    {
                        rec.HasFormula = true;
                        rec.Formula = fEl.Value;
                        rec.SharedIndex = Attr(fEl, "si");
                        rec.FormulaType = Attr(fEl, "t");
                        rec.FormulaRef = Attr(fEl, "ref");
                        rec.FormulaAca = IsTrue(Attr(fEl, "aca"));
                        rec.FormulaBx = IsTrue(Attr(fEl, "bx"));
                        rec.Shared = string.Equals(rec.FormulaType, "shared", StringComparison.OrdinalIgnoreCase);
                    }
                    XElement vEl = FirstDescendant(cEl, "v");
                    if (vEl != null) rec.Value = vEl.Value;
                    XElement isEl = FirstDescendant(cEl, "is");
                    if (isEl != null)
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (XElement te in Descendants(isEl, "t"))
                        {
                            if (te.Value == null) continue;
                            int take = Math.Min(te.Value.Length, MaxStringChars - sb.Length);
                            if (take > 0) sb.Append(te.Value, 0, take);
                            if (sb.Length >= MaxStringChars) break;
                        }
                        rec.Inline = sb.ToString();
                    }
                    if (records.Count >= MaxCellsPerSheet)
                        throw new FileFormatException("La hoja .xlsx contiene demasiadas celdas.");
                    records.Add(rec);
                }
                rowIndex++;
            }

            foreach (XElement mc in Descendants(doc.Root, "mergeCell"))
            {
                string r = Attr(mc, "ref");
                int r1, c1, r2, c2;
                if (AddressUtil.TryParseRange(r, out r1, out c1, out r2, out c2))
                    sheet.Merges.Add(new MergeRange(r1, c1, r2, c2));
            }
            sheet.RebuildMergeLookup();

            XElement pane = FirstDescendant(doc.Root, "pane");
            if (pane != null)
            {
                string state = Attr(pane, "state");
                if (!string.IsNullOrEmpty(state) && state.IndexOf("frozen", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    double xs, ys;
                    double.TryParse(Attr(pane, "xSplit"), NumberStyles.Float, CultureInfo.InvariantCulture, out xs);
                    double.TryParse(Attr(pane, "ySplit"), NumberStyles.Float, CultureInfo.InvariantCulture, out ys);
                    sheet.FreezeCols = Math.Max(0, (int)Math.Floor(xs));
                    sheet.FreezeRows = Math.Max(0, (int)Math.Floor(ys));
                }
            }

            foreach (XElement link in Descendants(doc.Root, "hyperlink"))
            {
                string r = Attr(link, "ref");
                if (string.IsNullOrEmpty(r)) continue;
                int row, col; bool ar, ac;
                string first = r;
                int colon = first.IndexOf(':');
                if (colon > 0) first = first.Substring(0, colon);
                if (!AddressUtil.TryParseCell(first, out row, out col, out ar, out ac)) continue;
                string rid = RelAttr(link, "id");
                string location = Attr(link, "location");
                string target = location;
                if (!string.IsNullOrEmpty(rid) && sheetRels != null)
                {
                    string relTarget;
                    if (sheetRels.TryGetValue(rid, out relTarget) && !string.IsNullOrEmpty(relTarget))
                        target = relTarget;
                }
                if (string.IsNullOrEmpty(target)) continue;
                int rr1, cc1, rr2, cc2;
                if (AddressUtil.TryParseRange(r, out rr1, out cc1, out rr2, out cc2) &&
                    (long)(rr2 - rr1 + 1) * (cc2 - cc1 + 1) <= 10000)
                {
                    for (int rr = rr1; rr <= rr2; rr++)
                        for (int cc = cc1; cc <= cc2; cc++)
                        {
                            if (!ValidCellRange(rr, cc)) continue;
                            Cell c = sheet.GetOrCreate(rr, cc);
                            c.Hyperlink = target;
                            c.HyperlinkTooltip = Attr(link, "tooltip");
                            c.HyperlinkInternal = !string.IsNullOrEmpty(location);
                            c.Dirty = false;
                        }
                }
                else
                {
                    Cell cell = sheet.GetOrCreate(row, col);
                    cell.Hyperlink = target;
                    cell.HyperlinkTooltip = Attr(link, "tooltip");
                    cell.HyperlinkInternal = !string.IsNullOrEmpty(location);
                    cell.Dirty = false;
                }
            }

            Dictionary<string, SharedFormula> masters = new Dictionary<string, SharedFormula>();
            foreach (XlsxCell rec in records)
            {
                if (rec.HasFormula && rec.Shared && !string.IsNullOrEmpty(rec.Formula) && rec.SharedIndex != null)
                    masters[rec.SharedIndex] = new SharedFormula(rec.Formula, rec.Row, rec.Col);
            }
            foreach (XlsxCell rec in records)
            {
                if (rec.HasFormula && rec.Shared && string.IsNullOrEmpty(rec.Formula) && rec.SharedIndex != null)
                {
                    SharedFormula master;
                    if (masters.TryGetValue(rec.SharedIndex, out master))
                        rec.Formula = FormulaEngine.TranslateFormula("=" + master.Formula, rec.Row - master.Row, rec.Col - master.Col).TrimStart('=');
                }
            }

            foreach (XlsxCell rec in records)
            {
                if (rec.Row < 0 || rec.Col < 0 || rec.Row >= Limits.MaxRows || rec.Col >= Limits.MaxCols) continue;
                string formula = rec.Formula;
                if (rec.HasFormula && !string.IsNullOrEmpty(formula))
                {
                    string f = formula.Trim();
                    if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1);
                    f = ConvertExcelSeparators(f);
                    sheet.SetInput(rec.Row, rec.Col, "=" + f);
                    Cell formulaCell = sheet.GetCell(rec.Row, rec.Col);
                    CellValue cached = CachedValueOf(rec, shared);
                    if (formulaCell != null && cached.Kind != ValueKind.Empty)
                    {
                        formulaCell.CachedValue = cached;
                        formulaCell.HasCachedValue = true;
                    }
                }
                else if (!rec.HasFormula)
                {
                    string type = rec.Type ?? "";
                    string text = ValueToText(rec, shared);
                    if (string.Equals(type, "e", StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.SetError(rec.Row, rec.Col, text);
                    }
                    else if (string.IsNullOrEmpty(type) || string.Equals(type, "n", StringComparison.OrdinalIgnoreCase))
                    {
                        double number;
                        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                            sheet.SetNumber(rec.Row, rec.Col, number);
                        else if (!string.IsNullOrEmpty(text))
                            sheet.SetTextLiteral(rec.Row, rec.Col, text);
                        else if (rec.StyleIndex > 0) sheet.GetOrCreate(rec.Row, rec.Col);
                    }
                    else if (string.Equals(type, "b", StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.SetInput(rec.Row, rec.Col, text);
                    }
                    else if (string.Equals(type, "d", StringComparison.OrdinalIgnoreCase))
                    {
                        DateTime dt;
                        if (DateTime.TryParse(text, CultureInfo.InvariantCulture,
                            DateTimeStyles.RoundtripKind, out dt))
                        {
                            double serial = dt.ToOADate();
                            if (sheet.Workbook != null && sheet.Workbook.Date1904) serial -= 1462;
                            sheet.SetNumber(rec.Row, rec.Col, serial);
                            CellStyle dateStyle = new CellStyle();
                            dateStyle.Format = NumberFormatKind.Date;
                            sheet.SetStyle(rec.Row, rec.Col, dateStyle);
                            rec.ForceDateStyle = true;
                        }
                        else if (!string.IsNullOrEmpty(text)) sheet.SetTextLiteral(rec.Row, rec.Col, text);
                    }
                    else if (!string.IsNullOrEmpty(text))
                    {
                        // s, inlineStr, str y tipos desconocidos son texto literal:
                        // nunca se reinterpretan como fórmula, número o booleano.
                        sheet.SetTextLiteral(rec.Row, rec.Col, text);
                    }
                    else if (rec.StyleIndex > 0) sheet.GetOrCreate(rec.Row, rec.Col);
                }
                int validStyle = (rec.StyleIndex > 0 && rec.StyleIndex < styles.Count) ? rec.StyleIndex : 0;
                if (validStyle > 0)
                {
                    CellStyle st = styles[validStyle];
                    if (st != null && !st.IsDefault()) sheet.SetStyle(rec.Row, rec.Col, st.Clone());
                }
                Cell cell = sheet.GetCell(rec.Row, rec.Col);
                if (cell != null)
                {
                    cell.SourceStyleIndex = rec.ForceDateStyle ? -1 : validStyle;
                    cell.StyleDirty = rec.ForceDateStyle;
                    cell.FormulaDirty = false;
                    cell.Dirty = false;
                    if (rec.HasFormula)
                    {
                        string source = rec.Formula == null ? "" : rec.Formula.Trim();
                        if (source.StartsWith("=", StringComparison.Ordinal)) source = source.Substring(1);
                        cell.SourceFormulaText = source;
                        // Las fórmulas compartidas se expanden a fórmulas normales;
                        // escribir t="shared" sin su atributo si produciría un archivo inválido.
                        cell.SourceFormulaType = string.Equals(rec.FormulaType, "shared", StringComparison.OrdinalIgnoreCase)
                            ? null : rec.FormulaType;
                        cell.SourceFormulaRef = string.Equals(rec.FormulaType, "shared", StringComparison.OrdinalIgnoreCase)
                            ? null : rec.FormulaRef;
                        cell.SourceFormulaAca = rec.FormulaAca;
                        cell.SourceFormulaBx = rec.FormulaBx;
                    }
                }
            }
            ParsePageSetup(sheet, doc);
            ParseConditionalFormats(sheet, doc, dxfs, theme);
            ParseDataValidations(sheet, doc);
            XElement autoFilter = FirstDescendant(doc.Root, "autoFilter");
            if (autoFilter != null)
            {
                sheet.AutoFilterRef = Attr(autoFilter, "ref");
                int ar1, ac1, ar2, ac2;
                if (AddressUtil.TryParseRange(sheet.AutoFilterRef, out ar1, out ac1, out ar2, out ac2))
                    ParseFilterSelections(sheet, autoFilter, ac1);
            }
            sheet.WidthsDirty = false;
            sheet.RowsDirty = false;
            sheet.MergesDirty = false;
            sheet.HyperlinksDirty = false;
            sheet.StructureDirty = false;
        }

        private static void ParseConditionalFormats(Sheet sheet, XDocument doc,
                                                    List<CellStyle> dxfs, ThemeColors theme)
        {
            if (doc == null || doc.Root == null) return;
            foreach (XElement cfEl in Descendants(doc.Root, "conditionalFormatting"))
            {
                string sqref = Attr(cfEl, "sqref") ?? "";
                if (sqref.Length == 0) continue;
                List<ConditionalRule> rules = new List<ConditionalRule>();
                foreach (XElement ruleEl in Descendants(cfEl, "cfRule"))
                {
                    ConditionalRule rule = new ConditionalRule();
                    int priority;
                    if (int.TryParse(Attr(ruleEl, "priority"), out priority)) rule.Priority = priority;
                    rule.StopIfTrue = IsTrue(Attr(ruleEl, "stopIfTrue"));
                    int dxfId;
                    if (int.TryParse(Attr(ruleEl, "dxfId"), out dxfId) &&
                        dxfId >= 0 && dxfId < dxfs.Count)
                        rule.Style = dxfs[dxfId].Clone();
                    string type = Attr(ruleEl, "type") ?? "";
                    rule.Kind = string.IsNullOrEmpty(type) ? "expression" : type;
                    rule.Operator = Attr(ruleEl, "operator");
                    rule.TextValue = Attr(ruleEl, "text");
                    rule.TimePeriod = Attr(ruleEl, "timePeriod");
                    int rank;
                    if (int.TryParse(Attr(ruleEl, "rank"), out rank)) rule.Rank = rank;
                    rule.Percent = IsTrue(Attr(ruleEl, "percent"));
                    rule.Bottom = IsTrue(Attr(ruleEl, "bottom"));
                    rule.IconSetName = Attr(FirstDescendant(ruleEl, "iconSet"), "val");
                    XElement formula = FirstDescendant(ruleEl, "formula");
                    if (formula != null) rule.Formula = formula.Value;

                    if (string.Equals(type, "dataBar", StringComparison.OrdinalIgnoreCase))
                    {
                        rule.IsDataBar = true;
                        rule.DataBarColor = Color.FromArgb(166, 166, 166);
                        XElement color = FirstDescendant(FirstDescendant(ruleEl, "dataBar"), "color");
                        Color parsed = ResolveColor(color, theme);
                        if (!parsed.IsEmpty) rule.DataBarColor = parsed;
                    }
                    else if (string.Equals(type, "colorScale", StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (XElement vo in Descendants(ruleEl, "cfvo"))
                        {
                            string vt = Attr(vo, "type") ?? "num";
                            double vv;
                            double.TryParse(Attr(vo, "val"), NumberStyles.Float,
                                CultureInfo.InvariantCulture, out vv);
                            rule.ThresholdTypes.Add(vt);
                            rule.Thresholds.Add(vv);
                        }
                        foreach (XElement colorEl in Descendants(ruleEl, "color"))
                        {
                            Color c = ResolveColor(colorEl, theme);
                            if (!c.IsEmpty) rule.ScaleColors.Add(c);
                        }
                    }
                    else if (string.Equals(type, "iconSet", StringComparison.OrdinalIgnoreCase))
                    {
                        XElement iconSet = FirstDescendant(ruleEl, "iconSet");
                        if (iconSet != null)
                        {
                            string name = Attr(iconSet, "iconSet");
                            if (!string.IsNullOrEmpty(name)) rule.IconSetName = name;
                        }
                        foreach (XElement vo in Descendants(ruleEl, "cfvo"))
                        {
                            string vt = Attr(vo, "type") ?? "percent";
                            double vv;
                            double.TryParse(Attr(vo, "val"), NumberStyles.Float,
                                CultureInfo.InvariantCulture, out vv);
                            rule.ThresholdTypes.Add(vt);
                            rule.Thresholds.Add(vv);
                        }
                    }
                    else if (string.Equals(type, "expression", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(type, "cellIs", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(type, "containsText", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(type, "notContainsText", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(type, "beginsWith", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(type, "endsWith", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(type, "timePeriod", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(type, "top10", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(type, "duplicateValues", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(type, "uniqueValues", StringComparison.OrdinalIgnoreCase))
                    {
                        // La fórmula o los atributos ya fueron capturados.
                    }
                    else continue;
                    rules.Add(rule);
                }
                if (rules.Count == 0) continue;
                foreach (string part in sqref.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    int r1, c1, r2, c2;
                    if (!AddressUtil.TryParseRange(part, out r1, out c1, out r2, out c2)) continue;
                    ConditionalFormat cf = new ConditionalFormat
                    {
                        Row1 = r1, Col1 = c1, Row2 = r2, Col2 = c2
                    };
                    foreach (ConditionalRule rule in rules) cf.Rules.Add(rule.Clone());
                    sheet.ConditionalFormats.Add(cf);
                }
            }
        }

        private static void ParsePageSetup(Sheet sheet, XDocument doc)
        {
            if (doc == null || doc.Root == null) return;
            XElement view = FirstDescendant(doc.Root, "sheetView");
            if (view != null)
            {
                string grid = Attr(view, "showGridLines");
                if (grid != null) sheet.ShowGridLines = IsTrue(grid);
            }
            XElement protection = FirstDescendant(doc.Root, "sheetProtection");
            if (protection != null)
            {
                sheet.IsProtected = IsTrue(Attr(protection, "sheet"));
                sheet.ProtectionPasswordHash = Attr(protection, "password");
            }
            XElement headerFooter = FirstDescendant(doc.Root, "headerFooter");
            if (headerFooter != null)
            {
                XElement header = FirstDescendant(headerFooter, "oddHeader");
                XElement footer = FirstDescendant(headerFooter, "oddFooter");
                if (header != null) sheet.PrintHeader = header.Value;
                if (footer != null) sheet.PrintFooter = footer.Value;
            }
            XElement rowBreaks = FirstDescendant(doc.Root, "rowBreaks");
            if (rowBreaks != null)
            {
                foreach (XElement br in Descendants(rowBreaks, "brk"))
                {
                    int id;
                    if (int.TryParse(Attr(br, "id"), out id) && id >= 0)
                        sheet.PrintRowBreaks.Add(id);
                }
            }
            XElement printOptions = FirstDescendant(doc.Root, "printOptions");
            if (printOptions != null)
            {
                sheet.PrintGridLines = IsTrue(Attr(printOptions, "gridLines"));
                sheet.PrintHeadings = IsTrue(Attr(printOptions, "headings"));
            }
            XElement setup = FirstDescendant(doc.Root, "pageSetup");
            if (setup != null)
            {
                string orientation = Attr(setup, "orientation");
                sheet.PrintLandscape = string.Equals(orientation, "landscape", StringComparison.OrdinalIgnoreCase);
                int value;
                if (int.TryParse(Attr(setup, "paperSize"), out value)) sheet.PrintPaperSize = value;
                if (int.TryParse(Attr(setup, "scale"), out value) && value > 0)
                    sheet.PrintScalePercent = Math.Max(10, Math.Min(400, value));
                if (int.TryParse(Attr(setup, "fitToWidth"), out value)) sheet.PrintFitToWidth = value > 0;
                if (int.TryParse(Attr(setup, "fitToHeight"), out value)) sheet.PrintFitToHeight = value > 0;
            }
            XElement margins = FirstDescendant(doc.Root, "pageMargins");
            if (margins != null)
            {
                float left, right, top, bottom;
                float.TryParse(Attr(margins, "left"), NumberStyles.Float, CultureInfo.InvariantCulture, out left);
                float.TryParse(Attr(margins, "right"), NumberStyles.Float, CultureInfo.InvariantCulture, out right);
                float.TryParse(Attr(margins, "top"), NumberStyles.Float, CultureInfo.InvariantCulture, out top);
                float.TryParse(Attr(margins, "bottom"), NumberStyles.Float, CultureInfo.InvariantCulture, out bottom);
                float mm = Math.Max(Math.Max(left, right), Math.Max(top, bottom)) * 25.4f;
                if (mm > 0) sheet.PrintMarginMm = Math.Max(5f, Math.Min(40f, mm));
            }
        }

        private static void ParseDataValidations(Sheet sheet, XDocument doc)
        {
            if (doc == null || doc.Root == null) return;
            XElement box = FirstDescendant(doc.Root, "dataValidations");
            if (box == null) return;
            foreach (XElement dv in Descendants(box, "dataValidation"))
            {
                XElement formula = FirstDescendant(dv, "formula1");
                if (formula == null) continue;
                string listText = formula.Value ?? "";
                bool literalList = listText.Length >= 2 && listText[0] == '"';
                if (literalList)
                    listText = listText.Substring(1, listText.Length - 2).Replace("\"\"", "\"");
                string[] items = literalList ? listText.Split(',') : new string[0];
                string sqref = Attr(dv, "sqref") ?? "";
                foreach (string part in sqref.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    int r1, c1, r2, c2;
                    if (!AddressUtil.TryParseRange(part, out r1, out c1, out r2, out c2)) continue;
                    DataValidationRule rule = new DataValidationRule();
                    rule.Row1 = Math.Min(r1, r2); rule.Col1 = Math.Min(c1, c2);
                    rule.Row2 = Math.Max(r1, r2); rule.Col2 = Math.Max(c1, c2);
                    rule.AllowBlank = string.Equals(Attr(dv, "allowBlank"), "1", StringComparison.OrdinalIgnoreCase);
                    rule.Message = Attr(dv, "error") ?? "El valor no está en la lista.";
                    if (!literalList) rule.SourceRange = listText;
                    foreach (string item in items) if (item.Trim().Length > 0) rule.Items.Add(item.Trim());
                    if (rule.Items.Count > 0 || !string.IsNullOrEmpty(rule.SourceRange))
                        sheet.Validations.Add(rule);
                }
            }
        }

        private static bool ValidCellRange(int row, int col)
        {
            return row >= 0 && row < Limits.MaxRows && col >= 0 && col < Limits.MaxCols;
        }

        private static void ParseSheetTables(ZipArchive zip, Sheet sheet, XDocument sheetDoc,
                                             Dictionary<string, string> sheetRels, List<TableInfo> tables,
                                             Dictionary<string, TableStyleDefinition> tableStyles)
        {
            if (sheetDoc == null || sheetDoc.Root == null) return;
            string baseDir = PartDirectory(sheet.SourcePart ?? "xl/");
            foreach (XElement tp in Descendants(sheetDoc.Root, "tablePart"))
            {
                string rid = RelAttr(tp, "id");
                string target;
                if (rid == null || sheetRels == null || !sheetRels.TryGetValue(rid, out target)) continue;
                string part = ResolvePart(baseDir, target);
                XDocument td = ReadXmlIfExists(zip, part);
                if (td == null || td.Root == null) continue;
                XElement table = string.Equals(td.Root.Name.LocalName, "table", StringComparison.OrdinalIgnoreCase)
                    ? td.Root : FirstDescendant(td.Root, "table");
                if (table == null) continue;
                string name = Attr(table, "displayName");
                if (string.IsNullOrEmpty(name)) name = Attr(table, "name");
                int r1, c1, r2, c2;
                if (!AddressUtil.TryParseRange(Attr(table, "ref"), out r1, out c1, out r2, out c2)) continue;
                TableInfo info = new TableInfo();
                info.Name = name;
                info.Sheet = sheet;
                info.Row1 = r1; info.Col1 = c1; info.Row2 = r2; info.Col2 = c2;
                int headerRows;
                if (int.TryParse(Attr(table, "headerRowCount"), out headerRows)) info.HeaderRows = headerRows;
                int totalsRows;
                if (int.TryParse(Attr(table, "totalsRowCount"), out totalsRows)) info.TotalsRows = totalsRows;
                XElement cols = FirstDescendant(table, "tableColumns");
                if (cols != null)
                {
                    int idx = 0;
                    foreach (XElement col in Descendants(cols, "tableColumn"))
                    {
                        string cn = Attr(col, "name");
                        if (!string.IsNullOrEmpty(cn)) info.Columns[cn] = idx;
                        idx++;
                    }
                }
                tables.Add(info);
                XElement styleInfo = FirstDescendant(table, "tableStyleInfo");
                TableFormat tf = new TableFormat
                {
                    Row1 = info.Row1, Col1 = info.Col1, Row2 = info.Row2, Col2 = info.Col2,
                    HeaderRows = info.HeaderRows, TotalsRows = info.TotalsRows,
                    Sheet = sheet, Name = info.Name, SourcePart = part
                };
                if (styleInfo != null)
                {
                    tf.StyleName = Attr(styleInfo, "name");
                    tf.ShowRowStripes = IsTrue(Attr(styleInfo, "showRowStripes"));
                    tf.ShowFirstColumn = IsTrue(Attr(styleInfo, "showFirstColumn"));
                    tf.ShowLastColumn = IsTrue(Attr(styleInfo, "showLastColumn"));
                    TableStyleDefinition def;
                    if (!string.IsNullOrEmpty(tf.StyleName) &&
                        tableStyles.TryGetValue(tf.StyleName, out def))
                        tf.Style = def;
                }
                sheet.TableFormats.Add(tf);
                XElement tableFilter = FirstDescendant(table, "autoFilter");
                if (tableFilter != null) ParseFilterSelections(sheet, tableFilter, tf.Col1);
            }
        }

        private static void ParseComments(ZipArchive zip, Sheet sheet, Dictionary<string, string> sheetRels)
        {
            if (zip == null || sheet == null || sheetRels == null) return;
            string sheetDir = PartDirectory(sheet.SourcePart ?? "xl/worksheets/sheet1.xml");
            int commentCount = 0;
            foreach (string target in sheetRels.Values)
            {
                if (string.IsNullOrEmpty(target) ||
                    target.IndexOf("comments", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                string part = ResolvePart(sheetDir, target);
                XDocument comments = ReadXmlIfExists(zip, part);
                if (comments == null || comments.Root == null) continue;
                foreach (XElement comment in Descendants(comments.Root, "comment"))
                {
                    if (++commentCount > MaxComments) return;
                    string reference = Attr(comment, "ref");
                    if (string.IsNullOrEmpty(reference)) continue;
                    int row, col, row2, col2;
                    if (!AddressUtil.TryParseRange(reference, out row, out col, out row2, out col2)) continue;
                    StringBuilder text = new StringBuilder();
                    foreach (XElement t in Descendants(comment, "t")) text.Append(t.Value);
                    if (text.Length == 0) continue;
                    Cell cell = sheet.GetOrCreate(row, col);
                    cell.Comment = text.ToString();
                }
            }
        }

        private static void ParsePivotRanges(ZipArchive zip, Sheet sheet, Dictionary<string, string> sheetRels)
        {
            if (zip == null || sheet == null || sheetRels == null) return;
            string sheetDir = PartDirectory(sheet.SourcePart ?? "xl/worksheets/sheet1.xml");
            foreach (string target in sheetRels.Values)
            {
                if (string.IsNullOrEmpty(target) ||
                    target.IndexOf("pivotTables", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                string pivotPart = ResolvePart(sheetDir, target);
                XDocument pivot = ReadXmlIfExists(zip, pivotPart);
                if (pivot == null || pivot.Root == null) continue;
                XElement location = FirstDescendant(pivot.Root, "location");
                string reference = Attr(location, "ref");
                int r1, c1, r2, c2;
                if (string.IsNullOrEmpty(reference) ||
                    !AddressUtil.TryParseRange(reference, out r1, out c1, out r2, out c2)) continue;
                PivotRange range = new PivotRange
                {
                    Row1 = r1, Col1 = c1, Row2 = r2, Col2 = c2
                };
                int firstDataCol;
                if (int.TryParse(Attr(location, "firstDataCol"), out firstDataCol))
                    range.FirstDataCol = Math.Max(1, firstDataCol);
                XElement style = FirstDescendant(pivot.Root, "pivotTableStyleInfo");
                if (style != null)
                {
                    range.StyleName = Attr(style, "name");
                    range.ShowRowHeaders = !string.Equals(Attr(style, "showRowHeaders"), "0");
                    range.ShowColHeaders = !string.Equals(Attr(style, "showColHeaders"), "0");
                    range.ShowRowStripes = !string.Equals(Attr(style, "showRowStripes"), "0");
                }
                range.RowCaption = Attr(pivot.Root, "rowHeaderCaption");
                XElement rowFields = FirstDescendant(pivot.Root, "rowFields");
                if (rowFields != null)
                {
                    foreach (XElement rf in Descendants(rowFields, "field"))
                    {
                        int rowField;
                        if (int.TryParse(Attr(rf, "x"), out rowField)) range.RowFields.Add(rowField);
                    }
                }
                if (range.RowFields.Count > 0) range.RowField = range.RowFields[0];
                XElement colFields = FirstDescendant(pivot.Root, "colFields");
                if (colFields != null)
                {
                    foreach (XElement cf in Descendants(colFields, "field"))
                    {
                        int colField;
                        if (int.TryParse(Attr(cf, "x"), out colField) && colField >= 0)
                        {
                            range.ColField = colField;
                            break;
                        }
                    }
                }
                XElement dataFields = FirstDescendant(pivot.Root, "dataFields");
                if (dataFields != null)
                {
                    foreach (XElement df in Descendants(dataFields, "dataField"))
                    {
                        int fld;
                        if (int.TryParse(Attr(df, "fld"), out fld)) range.DataFields.Add(fld);
                        string name = Attr(df, "name");
                        range.DataNames.Add(string.IsNullOrEmpty(name) ? "Valor" : name);
                    }
                }
                range.Cache = ParsePivotCache(zip, pivotPart);
                sheet.PivotRanges.Add(range);
            }
        }

        private static PivotCache ParsePivotCache(ZipArchive zip, string pivotPart)
        {
            string dir = PartDirectory(pivotPart);
            string relsPath = dir + "_rels/" + Path.GetFileName(pivotPart) + ".rels";
            XDocument rels = ReadXmlIfExists(zip, relsPath);
            string cachePart = null;
            if (rels != null && rels.Root != null)
            {
                foreach (XElement rel in Descendants(rels.Root, "Relationship"))
                {
                    if ((Attr(rel, "Type") ?? "").IndexOf("pivotCacheDefinition",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        cachePart = ResolvePart(dir, Attr(rel, "Target"));
                        break;
                    }
                }
            }
            if (cachePart == null) return null;
            XDocument def = ReadXmlIfExists(zip, cachePart);
            if (def == null || def.Root == null) return null;
            PivotCache cache = new PivotCache();
            foreach (XElement field in Descendants(def.Root, "cacheField"))
            {
                cache.FieldNames.Add(Attr(field, "name") ?? "");
            }
            XDocument records = ReadXmlIfExists(zip,
                ResolvePart(PartDirectory(cachePart), RelTargetByType(
                    ReadXmlIfExists(zip, PartDirectory(cachePart) + "_rels/" +
                        Path.GetFileName(cachePart) + ".rels"), "/pivotCacheRecords") ??
                    "pivotCacheRecords1.xml"));
            if (records == null || records.Root == null) return cache;
            foreach (XElement r in Descendants(records.Root, "r"))
            {
                List<CellValue> row = new List<CellValue>();
                foreach (XElement v in r.Elements())
                {
                    string name = v.Name.LocalName;
                    if (name == "m") row.Add(CellValue.Empty);
                    string raw = Attr(v, "v") ?? v.Value;
                    if (name == "n")
                    {
                        double d;
                        row.Add(double.TryParse(raw, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out d)
                            ? CellValue.FromNumber(d) : CellValue.Empty);
                    }
                    else if (name == "s" || name == "str") row.Add(CellValue.FromText(raw));
                    else if (name == "b") row.Add(CellValue.FromBool(raw == "1"));
                    else if (name == "x")
                    {
                        int index;
                        if (int.TryParse(raw, out index)) row.Add(RowFieldSharedValue(def,
                            row.Count, index));
                        else row.Add(CellValue.Empty);
                    }
                    else row.Add(CellValue.Empty);
                }
                cache.Rows.Add(row);
            }
            return cache;
        }

        private static CellValue RowFieldSharedValue(XDocument def, int fieldIndex, int itemIndex)
        {
            if (def == null) return CellValue.Empty;
            int i = 0;
            foreach (XElement field in Descendants(def.Root, "cacheField"))
            {
                if (i++ != fieldIndex) continue;
                XElement shared = FirstDescendant(field, "sharedItems");
                if (shared == null) return CellValue.Empty;
                int j = 0;
                foreach (XElement item in shared.Elements())
                {
                    if (j++ != itemIndex) continue;
                    string name = item.Name.LocalName;
                    string raw = Attr(item, "v") ?? item.Value;
                    if (name == "s" || name == "str") return CellValue.FromText(raw);
                    if (name == "n")
                    {
                        double d;
                        if (double.TryParse(raw, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out d)) return CellValue.FromNumber(d);
                    }
                    if (name == "b") return CellValue.FromBool(raw == "1");
                    return CellValue.Empty;
                }
            }
            return CellValue.Empty;
        }

        private static void ParseSheetCharts(ZipArchive zip, Sheet sheet, XDocument sheetDoc,
                                             Dictionary<string, string> sheetRels)
        {
            XElement drawing = FirstDescendant(sheetDoc.Root, "drawing");
            if (drawing == null) return;
            string drawingRid = RelAttr(drawing, "id");
            string drawingTarget;
            if (drawingRid == null || sheetRels == null ||
                !sheetRels.TryGetValue(drawingRid, out drawingTarget)) return;
            string sheetDir = PartDirectory(sheet.SourcePart ?? "xl/worksheets/sheet1.xml");
            string drawingPart = ResolvePart(sheetDir, drawingTarget);
            XDocument drawingDoc = ReadXmlIfExists(zip, drawingPart);
            if (drawingDoc == null || drawingDoc.Root == null) return;
            string drawingDir = PartDirectory(drawingPart);
            XDocument relDoc = ReadXmlIfExists(zip,
                drawingDir + "_rels/" + Path.GetFileName(drawingPart) + ".rels");
            Dictionary<string, string> drawingRels = ParseRelationships(relDoc, true);

            foreach (XElement anchor in drawingDoc.Root.Descendants())
            {
                string local = anchor.Name.LocalName;
                if (!string.Equals(local, "twoCellAnchor", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(local, "oneCellAnchor", StringComparison.OrdinalIgnoreCase)) continue;
                XElement slicerEl = FirstDescendant(anchor, "slicer");
                if (slicerEl != null)
                {
                    sheet.HasSlicerPlaceholders = true;
                    ParseSlicer(sheet, anchor, slicerEl);
                }
                if (FirstDescendant(anchor, "pic") != null) sheet.HasPicturePlaceholders = true;
                XElement chartEl = FirstDescendant(anchor, "chart");
                if (chartEl != null) sheet.HasChartPlaceholders = true;
                if (chartEl != null && sheet.Charts.Count >= MaxCharts) continue;
                if (chartEl == null)
                {
                    ParsePicture(zip, sheet, anchor, drawingRels, drawingDir, drawingPart);
                    ParseShape(sheet, anchor, drawingRels);
                    continue;
                }
                string rid = RelAttr(chartEl, "id");
                string chartTarget;
                if (rid == null || !drawingRels.TryGetValue(rid, out chartTarget)) continue;
                string chartPart = ResolvePart(drawingDir, chartTarget);
                XDocument chartDoc = ReadXmlIfExists(zip, chartPart);
                if (chartDoc == null) continue;
                SheetChart chart;
                try { chart = ChartXml.Parse(chartDoc); }
                catch { chart = null; }
                if (chart == null) continue;
                chart.ThemePalette = ChartData.ThemePalette == null
                    ? null : (Color[])ChartData.ThemePalette.Clone();

                XElement from = FirstDescendant(anchor, "from");
                int fromRow, fromCol;
                if (!ReadAnchorPoint(from, out fromCol, out fromRow)) continue;
                chart.FromRow = fromRow;
                chart.FromCol = fromCol;
                chart.FromColOff = ReadAnchorOffset(from, "colOff");
                chart.FromRowOff = ReadAnchorOffset(from, "rowOff");
                XElement to = FirstDescendant(anchor, "to");
                if (to != null)
                {
                    int toRow, toCol;
                    if (ReadAnchorPoint(to, out toCol, out toRow))
                    {
                        chart.ToRow = toRow;
                        chart.ToCol = toCol;
                        chart.ToColOff = ReadAnchorOffset(to, "colOff");
                        chart.ToRowOff = ReadAnchorOffset(to, "rowOff");
                    }
                }
                XElement ext = FirstDescendant(anchor, "ext");
                if (ext != null)
                {
                    long cx = ParseLong(Attr(ext, "cx"), 0);
                    long cy = ParseLong(Attr(ext, "cy"), 0);
                    if (cx > 0) chart.WidthPx = (int)Math.Max(40, Math.Min(4000, cx / 9525));
                    if (cy > 0) chart.HeightPx = (int)Math.Max(30, Math.Min(4000, cy / 9525));
                }
                chart.SourceChartPart = chartPart;
                chart.SourceDrawingPart = drawingPart;
                chart.Imported = true;
                sheet.Charts.Add(chart);
            }
        }

        private static void ParsePicture(ZipArchive zip, Sheet sheet, XElement anchor,
                                         Dictionary<string, string> drawingRels, string drawingDir,
                                         string drawingPart)
        {
            if (sheet.Images.Count >= MaxImages) return;
            XElement pic = FirstDescendant(anchor, "pic");
            if (pic == null) return;
            XElement blip = FirstDescendant(pic, "blip");
            if (blip == null) return;
            string rid = RelAttr(blip, "embed");
            string target;
            if (rid == null || drawingRels == null || !drawingRels.TryGetValue(rid, out target)) return;
            string imagePart = ResolvePart(drawingDir, target);
            ZipArchiveEntry entry = GetEntry(zip, imagePart);
            if (entry == null) return;
            try
            {
                byte[] bytes = ReadEntryBytes(entry, MaxImageBytes);
                if (!IsSafeRasterImage(bytes)) return;
                Image image;
                using (MemoryStream ms = new MemoryStream(bytes))
                using (Image source = Image.FromStream(ms))
                {
                    if (source.Width <= 0 || source.Height <= 0 ||
                        source.Width > MaxImageDimension || source.Height > MaxImageDimension ||
                        (long)source.Width * source.Height > MaxImagePixels) return;
                    image = new Bitmap(source);
                }
                SheetImage si = new SheetImage();
                si.Image = image;
                XElement from = FirstDescendant(anchor, "from");
                int fromRow, fromCol;
                if (!ReadAnchorPoint(from, out fromCol, out fromRow)) { image.Dispose(); return; }
                si.FromRow = fromRow;
                si.FromCol = fromCol;
                si.FromColOff = ReadAnchorOffset(from, "colOff");
                si.FromRowOff = ReadAnchorOffset(from, "rowOff");
                XElement to = FirstDescendant(anchor, "to");
                if (to != null)
                {
                    int toRow, toCol;
                    if (ReadAnchorPoint(to, out toCol, out toRow))
                    {
                        si.ToRow = toRow; si.ToCol = toCol;
                        si.ToColOff = ReadAnchorOffset(to, "colOff");
                        si.ToRowOff = ReadAnchorOffset(to, "rowOff");
                    }
                }
                XElement ext = FirstDescendant(anchor, "ext");
                if (ext != null)
                {
                    long cx = ParseLong(Attr(ext, "cx"), 0);
                    long cy = ParseLong(Attr(ext, "cy"), 0);
                    if (cx > 0) si.WidthPx = (int)Math.Max(20, Math.Min(8000, cx / 9525));
                    if (cy > 0) si.HeightPx = (int)Math.Max(20, Math.Min(8000, cy / 9525));
                }
                si.SourcePart = imagePart;
                si.SourceDrawingPart = drawingPart;
                si.SourceRelId = rid;
                sheet.Images.Add(si);
            }
            catch { }
        }

        private static void ParseSlicer(Sheet sheet, XElement anchor, XElement slicer)
        {
            if (sheet.Slicers.Count >= MaxSlicers) return;
            XElement from = FirstDescendant(anchor, "from");
            int fromRow, fromCol;
            if (!ReadAnchorPoint(from, out fromCol, out fromRow)) return;
            SheetSlicer s = new SheetSlicer
            {
                Name = Attr(slicer, "name"),
                Caption = Attr(slicer, "caption"),
                FieldName = Attr(slicer, "name"),
                FromRow = fromRow,
                FromCol = fromCol,
                FromColOff = ReadAnchorOffset(from, "colOff"),
                FromRowOff = ReadAnchorOffset(from, "rowOff")
            };
            XElement to = FirstDescendant(anchor, "to");
            if (to != null)
            {
                int toRow, toCol;
                if (ReadAnchorPoint(to, out toCol, out toRow))
                {
                    s.ToRow = toRow; s.ToCol = toCol;
                    s.ToColOff = ReadAnchorOffset(to, "colOff");
                    s.ToRowOff = ReadAnchorOffset(to, "rowOff");
                }
            }
            if (string.IsNullOrEmpty(s.Name)) s.Name = "Segmentador";
            sheet.Slicers.Add(s);
        }

        private static void ParseShape(Sheet sheet, XElement anchor, Dictionary<string, string> drawingRels)
        {
            if (sheet.Shapes.Count >= MaxShapes) return;
            XElement sp = FirstDescendant(anchor, "sp");
            if (sp == null) return;
            XElement hlink = FirstDescendant(sp, "hlinkClick");
            if (hlink == null || drawingRels == null) return;
            string rid = RelAttr(hlink, "id");
            string target;
            if (rid == null || !drawingRels.TryGetValue(rid, out target) || string.IsNullOrEmpty(target))
                return;
            XElement from = FirstDescendant(anchor, "from");
            int fromRow, fromCol;
            if (!ReadAnchorPoint(from, out fromCol, out fromRow)) return;
            SheetShape shape = new SheetShape();
            shape.Name = Attr(FirstDescendant(sp, "cNvPr"), "name");
            shape.FromRow = fromRow;
            shape.FromCol = fromCol;
            shape.FromColOff = ReadAnchorOffset(from, "colOff");
            shape.FromRowOff = ReadAnchorOffset(from, "rowOff");
            XElement to = FirstDescendant(anchor, "to");
            if (to != null)
            {
                int toRow, toCol;
                if (ReadAnchorPoint(to, out toCol, out toRow))
                {
                    shape.ToRow = toRow; shape.ToCol = toCol;
                    shape.ToColOff = ReadAnchorOffset(to, "colOff");
                    shape.ToRowOff = ReadAnchorOffset(to, "rowOff");
                }
            }
            if (target.StartsWith("#", StringComparison.Ordinal))
                shape.HyperlinkInternal = target.Substring(1);
            else
                shape.Hyperlink = target;
            StringBuilder text = new StringBuilder();
            foreach (XElement t in Descendants(sp, "t"))
                text.Append(t.Value);
            shape.Text = text.ToString();
            XElement fill = FirstDescendant(sp, "solidFill");
            if (fill != null)
            {
                XElement color = FirstDescendant(fill, "srgbClr");
                if (color != null)
                {
                    Color c = ParseColor(Attr(color, "rgb"));
                    if (!c.IsEmpty) shape.FillColor = c;
                }
            }
            XElement rPr = FirstDescendant(sp, "rPr");
            if (rPr != null)
            {
                int size;
                string sz = Attr(rPr, "sz");
                if (!string.IsNullOrEmpty(sz) && int.TryParse(sz, out size) && size > 0)
                    shape.FontSize = Math.Max(6f, Math.Min(36f, size / 100f));
            }
            sheet.Shapes.Add(shape);
        }

        private static bool ReadAnchorPoint(XElement e, out int col, out int row)
        {
            col = row = 0;
            if (e == null) return false;
            XElement c = FirstDescendant(e, "col");
            XElement r = FirstDescendant(e, "row");
            if (c == null || r == null) return false;
            int.TryParse(c.Value, out col);
            int.TryParse(r.Value, out row);
            return col >= 0 && row >= 0;
        }

        private static long ReadAnchorOffset(XElement e, string name)
        {
            XElement off = FirstDescendant(e, name);
            return off == null ? 0 : ParseLong(off.Value, 0);
        }

        private static long ParseLong(string text, long fallback)
        {
            long value;
            return long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
                ? value : fallback;
        }

        private static void TranslateStructuredFormulas(Workbook wb, List<TableInfo> tables)
        {
            if (tables.Count == 0) return;
            foreach (Sheet sheet in wb.Sheets)
            {
                foreach (KeyValuePair<long, Cell> kv in sheet.Cells)
                {
                    Cell cell = kv.Value;
                    if (!cell.IsFormula || cell.Formula.IndexOf('[') < 0) continue;
                    int row = (int)(kv.Key >> 14);
                    string converted = TranslateStructuredRefs(cell.Formula, sheet, row, tables);
                    if (!string.Equals(converted, cell.Formula, StringComparison.Ordinal))
                    {
                        cell.Formula = converted;
                        cell.Raw = converted;
                        cell.Ast = null;
                        cell.AstReady = false;
                        cell.FormulaDirty = false;
                        cell.Value = CellValue.Empty;
                    }
                }
            }
            // Los nombres definidos también pueden usar referencias
            // estructuradas: Total = SUM(tabla[Columna]).
            foreach (DefinedName name in wb.DefinedNames)
            {
                if (string.IsNullOrEmpty(name.Formula) || name.Formula.IndexOf('[') < 0) continue;
                string converted = TranslateStructuredRefs(name.Formula, tables[0].Sheet, 0, tables);
                if (!string.Equals(converted, name.Formula, StringComparison.Ordinal))
                    name.Formula = converted;
            }
        }

        private static string TranslateStructuredRefs(string formula, Sheet sheet, int formulaRow, List<TableInfo> tables)
        {
            if (string.IsNullOrEmpty(formula) || tables.Count == 0) return formula;
            StringBuilder sb = new StringBuilder();
            int i = 0;
            while (i < formula.Length)
            {
                char ch = formula[i];
                if (ch == '"')
                {
                    int j = i + 1;
                    while (j < formula.Length)
                    {
                        if (formula[j] == '"')
                        {
                            if (j + 1 < formula.Length && formula[j + 1] == '"') { j += 2; continue; }
                            j++;
                            break;
                        }
                        j++;
                    }
                    int end = Math.Min(j, formula.Length);
                    sb.Append(formula.Substring(i, end - i));
                    i = end;
                    continue;
                }
                bool matched = false;
                foreach (TableInfo table in tables)
                {
                    if (string.Compare(formula, i, table.Name, 0, table.Name.Length, StringComparison.OrdinalIgnoreCase) != 0)
                        continue;
                    int after = i + table.Name.Length;
                    if (after >= formula.Length || formula[after] != '[') continue;
                    if (i > 0 && (char.IsLetterOrDigit(formula[i - 1]) || formula[i - 1] == '_' || formula[i - 1] == '.')) continue;
                    int close = FindMatchingBracket(formula, after);
                    if (close < 0) continue;
                    string inner = formula.Substring(after + 1, close - after - 1);
                    string range = BuildTableRef(table, inner, sheet, formulaRow);
                    if (range == null) continue;
                    sb.Append(range);
                    i = close + 1;
                    matched = true;
                    break;
                }
                if (!matched)
                {
                    sb.Append(ch);
                    i++;
                }
            }
            return sb.ToString();
        }

        private static int FindMatchingBracket(string s, int open)
        {
            int depth = 0;
            for (int i = open; i < s.Length; i++)
            {
                if (s[i] == '[') depth++;
                else if (s[i] == ']')
                {
                    depth--;
                    if (depth == 0) return i;
                }
            }
            return -1;
        }

        private static string BuildTableRef(TableInfo table, string inner, Sheet formulaSheet, int formulaRow)
        {
                        string normalized = inner.Replace("[", "").Replace("]", "");
            string[] parts = normalized.Split(new char[] { ',', ';' });
            string area = "Data";
            string column = null;
            foreach (string raw in parts)
            {
                string p = raw.Trim();
                if (p.Length == 0) continue;
                if (p.StartsWith("#", StringComparison.Ordinal))
                {
                    string u = p.ToUpperInvariant();
                    if (u == "#ALL") area = "All";
                    else if (u == "#HEADERS") area = "Headers";
                    else if (u == "#THIS ROW") area = "ThisRow";
                    else if (u == "#DATA") area = "Data";
                }
                else column = p;
            }
            int r1 = table.Row1, r2 = table.Row2;
            if (area == "Data")
            {
                r1 = table.Row1 + Math.Max(0, table.HeaderRows);
                r2 = table.Row2 - Math.Max(0, table.TotalsRows);
            }
            else if (area == "Headers")
                r2 = table.Row1 + Math.Max(0, table.HeaderRows) - 1;
            else if (area == "ThisRow")
            {
                if (formulaRow < table.Row1 || formulaRow > table.Row2) return null;
                r1 = r2 = formulaRow;
            }
            int c1 = table.Col1, c2 = table.Col2;
            if (!string.IsNullOrEmpty(column))
            {
                int idx;
                if (!table.Columns.TryGetValue(column, out idx)) return null;
                c1 = table.Col1 + idx;
                c2 = c1;
            }
            if (r1 > r2 || c1 > c2) return null;
            string sheetName = table.Sheet.Name.Replace("'", "''");
            return "'" + sheetName + "'!$" + AddressUtil.ColumnName(c1) + "$" + (r1 + 1) +
                   (r1 == r2 && c1 == c2 ? "" : ":$" + AddressUtil.ColumnName(c2) + "$" + (r2 + 1));
        }

        private static CellValue CachedValueOf(XlsxCell rec, List<string> shared)
        {
            string t = rec.Type ?? "";
            string text = ValueToText(rec, shared);
            if (string.Equals(t, "e", StringComparison.OrdinalIgnoreCase))
                return CellValue.FromError(string.IsNullOrEmpty(text) ? Errors.Error : text);
            if (string.Equals(t, "b", StringComparison.OrdinalIgnoreCase))
                return CellValue.FromBool(text == "1");
            if (string.Equals(t, "str", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(t, "s", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(t, "inlineStr", StringComparison.OrdinalIgnoreCase))
                return CellValue.FromText(text);
            if (string.Equals(t, "d", StringComparison.OrdinalIgnoreCase))
            {
                DateTime dt;
                if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out dt))
                    return CellValue.FromNumber(dt.ToOADate());
                return CellValue.Empty;
            }
            if (string.IsNullOrEmpty(text)) return CellValue.Empty;
            double number;
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                return CellValue.FromNumber(number);
            return CellValue.FromText(text);
        }

        private static string ValueToText(XlsxCell rec, List<string> shared)
        {
            string t = rec.Type;
            if (string.Equals(t, "s", StringComparison.OrdinalIgnoreCase))
            {
                int idx;
                if (int.TryParse(rec.Value, out idx) && idx >= 0 && idx < shared.Count) return shared[idx];
                return "";
            }
            if (string.Equals(t, "inlineStr", StringComparison.OrdinalIgnoreCase)) return rec.Inline ?? "";
            if (string.Equals(t, "str", StringComparison.OrdinalIgnoreCase)) return rec.Value ?? "";
            if (string.Equals(t, "b", StringComparison.OrdinalIgnoreCase))
                return rec.Value == "1" ? "VERDADERO" : "FALSO";
            if (string.Equals(t, "e", StringComparison.OrdinalIgnoreCase)) return rec.Value ?? Errors.Error;
            if (string.IsNullOrEmpty(rec.Value)) return "";
            return rec.Value;
        }

        private static string ConvertExcelSeparators(string formula)
        {
            StringBuilder sb = new StringBuilder(formula.Length);
            bool quoted = false;
            for (int i = 0; i < formula.Length; i++)
            {
                char ch = formula[i];
                if (ch == '"')
                {
                    sb.Append(ch);
                    if (quoted && i + 1 < formula.Length && formula[i + 1] == '"') { sb.Append('"'); i++; }
                    else quoted = !quoted;
                }
                else if (ch == ',' && !quoted) sb.Append(';');
                else sb.Append(ch);
            }
            return sb.ToString();
        }

        private sealed class XlsxCell
        {
            public int Row;
            public int Col;
            public string Type;
            public int StyleIndex = -1;
            public bool HasFormula;
            public bool Shared;
            public string SharedIndex;
            public string Formula;
            public string FormulaType;
            public string FormulaRef;
            public bool FormulaAca;
            public bool FormulaBx;
            public bool ForceDateStyle;
            public string Value;
            public string Inline;
        }

        private sealed class SharedFormula
        {
            public string Formula;
            public int Row;
            public int Col;
            public SharedFormula(string f, int r, int c) { Formula = f; Row = r; Col = c; }
        }

        private sealed class TableInfo
        {
            public string Name;
            public Sheet Sheet;
            public int Row1;
            public int Col1;
            public int Row2;
            public int Col2;
            public int HeaderRows = 1;
            public int TotalsRows;
            public Dictionary<string, int> Columns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        }

        // ---------------------------------------------------------------- SAVE

        public static bool LastSavePreserved { get; private set; }
        public static string LastSaveError { get; private set; }
        private static readonly object SaveGate = new object();

        public static void Save(Workbook wb, string path)
        {
            lock (SaveGate)
            {
            _readBudgetRemaining = -1;
            LastSavePreserved = false;
            LastSaveError = null;
            if (wb != null && wb.XlsxSource != null)
            {
                if (!CanPreserve(wb))
                    throw new FileFormatException(
                        "El archivo original cambió desde que se abrió. No se sobrescribió; reabra el archivo original antes de guardar.");
                bool preserved;
                try { preserved = TrySavePreserving(wb, path); }
                catch (Exception ex)
                {
                    throw new FileFormatException(
                        "No se pudo guardar conservando el contenido avanzado del .xlsx: " + ex.Message);
                }
                if (!preserved)
                    throw new FileFormatException(
                        "No se pudo guardar conservando el contenido avanzado del .xlsx. No se sobrescribió el original.");
                LastSavePreserved = true;
            }
            else
            {
                SaveNew(wb, path);
            }
            RefreshSourceInfo(wb, path);
            }
        }

        private sealed class SheetSaveInfo
        {
            public Sheet Sheet;
            public string Part;
            public string Rid;
            public string SheetId;
            public string State;
            public bool New;
        }

        private static bool CanPreserve(Workbook wb)
        {
            if (wb == null || wb.XlsxSource == null || string.IsNullOrEmpty(wb.XlsxSource.Path)) return false;
            try
            {
                FileInfo fi = new FileInfo(wb.XlsxSource.Path);
                return fi.Exists && fi.Length == wb.XlsxSource.Length &&
                       fi.LastWriteTimeUtc == wb.XlsxSource.LastWriteUtc;
            }
            catch { return false; }
        }

        private sealed class ImageSavePlan
        {
            public Sheet Sheet;
            public string DrawingPart;
            public List<SheetImage> Images = new List<SheetImage>();
            public List<SheetImage> Deleted = new List<SheetImage>();
        }

        private sealed class ChartSavePlan
        {
            public Sheet Sheet;
            public string DrawingPart;
            public string DrawingRelsPart;
            public bool NewDrawing;
            public List<SheetChart> Charts = new List<SheetChart>();
            public List<SheetChart> Rewrites = new List<SheetChart>();
            public List<int> ChartNumbers = new List<int>();
            public int NextRid = 1;
            public int NextShapeId = 2;
        }

        private static int ExistingPartNumber(ZipArchive zip, string prefix, string suffix)
        {
            int max = 0;
            foreach (ZipArchiveEntry e in zip.Entries)
            {
                string n = e.FullName.Replace('\\', '/');
                if (!n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ||
                    !n.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) continue;
                string mid = n.Substring(prefix.Length, n.Length - prefix.Length - suffix.Length);
                int v;
                if (int.TryParse(mid, out v) && v > max) max = v;
            }
            return max;
        }

        private static string RelativePart(string fromDir, string target)
        {
            if (!fromDir.EndsWith("/", StringComparison.Ordinal)) fromDir += "/";
            Uri from = new Uri("http://local/" + fromDir.TrimStart('/'));
            Uri to = new Uri("http://local/" + target.TrimStart('/'));
            return from.MakeRelativeUri(to).ToString();
        }

        private static void WriteChartPlan(ZipArchive source, ZipArchive destination,
                                           ChartSavePlan plan, string sheetPart, XDocument sheetDoc,
                                           Dictionary<string, string> srels,
                                           ref XDocument srelsDoc)
        {
            if (plan == null || (plan.Charts.Count == 0 && plan.Rewrites.Count == 0)) return;
            foreach (SheetChart rewrite in plan.Rewrites)
            {
                if (string.IsNullOrEmpty(rewrite.SourceChartPart)) continue;
                WriteXml(destination, rewrite.SourceChartPart, ChartWriter.ChartXml(rewrite));
                rewrite.Dirty = false;
            }
            if (plan.Charts.Count == 0) return;
            XDocument drawingDoc;
            XDocument drawingRelsDoc;
            XElement drawingRoot;
            if (!plan.NewDrawing)
            {
                drawingDoc = ReadXmlIfExists(source, plan.DrawingPart);
                drawingRelsDoc = ReadXmlIfExists(source, plan.DrawingRelsPart);
                drawingRoot = drawingDoc == null ? null : drawingDoc.Root;
                if (drawingRoot == null)
                {
                    drawingDoc = new XDocument(new XElement(ChartWriter.Xdr + "wsDr"));
                    drawingRoot = drawingDoc.Root;
                }
                drawingRoot.SetAttributeValue(XNamespace.Xmlns + "xdr", ChartWriter.Xdr.NamespaceName);
                drawingRoot.SetAttributeValue(XNamespace.Xmlns + "a", ChartWriter.A.NamespaceName);
                drawingRoot.SetAttributeValue(XNamespace.Xmlns + "c", ChartWriter.C.NamespaceName);
                drawingRoot.SetAttributeValue(XNamespace.Xmlns + "r", ChartWriter.R.NamespaceName);
                if (drawingRelsDoc == null || drawingRelsDoc.Root == null)
                    drawingRelsDoc = new XDocument(new XElement(PkgRel + "Relationships"));
                foreach (XElement r in Descendants(drawingRelsDoc.Root, "Relationship"))
                {
                    string id = Attr(r, "Id");
                    if (id != null && id.StartsWith("rId", StringComparison.OrdinalIgnoreCase))
                    {
                        int n;
                        if (int.TryParse(id.Substring(3), out n) && n >= plan.NextRid) plan.NextRid = n + 1;
                    }
                }
                foreach (XElement pr in Descendants(drawingRoot, "cNvPr"))
                {
                    int n;
                    if (int.TryParse(Attr(pr, "id"), out n) && n >= plan.NextShapeId) plan.NextShapeId = n + 1;
                }
            }
            else
            {
                drawingDoc = new XDocument(new XElement(ChartWriter.Xdr + "wsDr"));
                drawingRelsDoc = new XDocument(new XElement(PkgRel + "Relationships"));
                drawingRoot = drawingDoc.Root;
                drawingRoot.SetAttributeValue(XNamespace.Xmlns + "xdr", ChartWriter.Xdr.NamespaceName);
                drawingRoot.SetAttributeValue(XNamespace.Xmlns + "a", ChartWriter.A.NamespaceName);
                drawingRoot.SetAttributeValue(XNamespace.Xmlns + "c", ChartWriter.C.NamespaceName);
                drawingRoot.SetAttributeValue(XNamespace.Xmlns + "r", ChartWriter.R.NamespaceName);
            }

            for (int i = 0; i < plan.Charts.Count; i++)
            {
                SheetChart chart = plan.Charts[i];
                int chartNo = plan.ChartNumbers[i];
                string chartPart = "xl/charts/chart" + chartNo + ".xml";
                WriteXml(destination, chartPart, ChartWriter.ChartXml(chart));
                chart.SourceChartPart = chartPart;
                chart.SourceDrawingPart = plan.DrawingPart;
                chart.Imported = true;
                chart.Dirty = false;
                string rid = "rId" + plan.NextRid++;
                drawingRoot.Add(ChartWriter.DrawingAnchorXml(chart, plan.NextShapeId++, rid));
                drawingRelsDoc.Root.Add(new XElement(PkgRel + "Relationship",
                    new XAttribute("Id", rid),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                    new XAttribute("Target", RelativePart(PartDirectory(plan.DrawingPart), chartPart))));
            }
            WriteXml(destination, plan.DrawingPart, drawingDoc);
            WriteXml(destination, plan.DrawingRelsPart, drawingRelsDoc);

            if (plan.NewDrawing && sheetDoc != null && sheetDoc.Root != null)
            {
                AddDrawingRelationship(sheetDoc, ref srelsDoc,
                    RelativePart(PartDirectory(sheetPart), plan.DrawingPart), "rIdDrawing");
                if (srelsDoc != null && srelsDoc.Root != null)
                {
                    foreach (XElement r in srelsDoc.Root.Elements())
                    {
                        string id = Attr(r, "Id");
                        if (!string.IsNullOrEmpty(id)) srels[id] = Attr(r, "Target");
                    }
                }
            }
        }

        private static void WriteImagePlan(ZipArchive source, ZipArchive destination, ImageSavePlan plan)
        {
            XDocument doc = ReadXmlIfExists(source, plan.DrawingPart);
            if (doc == null || doc.Root == null) return;
            HashSet<string> deleted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SheetImage img in plan.Deleted)
                if (!string.IsNullOrEmpty(img.SourceRelId)) deleted.Add(img.SourceRelId);
            foreach (XElement anchor in doc.Root.Descendants().ToList())
            {
                string local = anchor.Name.LocalName;
                if (!string.Equals(local, "twoCellAnchor", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(local, "oneCellAnchor", StringComparison.OrdinalIgnoreCase)) continue;
                XElement blip = FirstDescendant(anchor, "blip");
                string rid = blip == null ? null : RelAttr(blip, "embed");
                if (rid == null) continue;
                if (deleted.Contains(rid)) { anchor.Remove(); continue; }
                SheetImage img = null;
                foreach (SheetImage candidate in plan.Images)
                    if (string.Equals(candidate.SourceRelId, rid, StringComparison.OrdinalIgnoreCase))
                    { img = candidate; break; }
                if (img == null) continue;
                XElement from = FirstDescendant(anchor, "from");
                SetAnchorPoint(from, img.FromCol, img.FromRow, img.FromColOff, img.FromRowOff);
                XElement to = FirstDescendant(anchor, "to");
                if (to != null && img.ToRow >= 0 && img.ToCol >= 0)
                    SetAnchorPoint(to, img.ToCol, img.ToRow, img.ToColOff, img.ToRowOff);
                XElement ext = FirstDescendant(anchor, "ext");
                if (ext != null)
                {
                    ext.SetAttributeValue("cx", (long)img.WidthPx * 9525);
                    ext.SetAttributeValue("cy", (long)img.HeightPx * 9525);
                }
                img.Dirty = false;
            }
            WriteXml(destination, plan.DrawingPart, doc);

            foreach (SheetImage img in plan.Images)
            {
                if (string.IsNullOrEmpty(img.SourcePart) || img.Image == null) continue;
                ZipArchiveEntry entry = destination.CreateEntry(img.SourcePart, CompressionLevel.Optimal);
                using (Stream s = entry.Open())
                {
                    string ext = Path.GetExtension(img.SourcePart).ToLowerInvariant();
                    System.Drawing.Imaging.ImageFormat format =
                        ext == ".jpg" || ext == ".jpeg" ? System.Drawing.Imaging.ImageFormat.Jpeg :
                        ext == ".gif" ? System.Drawing.Imaging.ImageFormat.Gif :
                        System.Drawing.Imaging.ImageFormat.Png;
                    using (Bitmap bmp = new Bitmap(img.Image)) bmp.Save(s, format);
                }
            }
            plan.Sheet.ImagesDirty = false;
            plan.Sheet.DeletedImages.Clear();
        }

        private static void SetAnchorPoint(XElement point, int col, int row, long colOff, long rowOff)
        {
            if (point == null) return;
            XElement c = FirstDescendant(point, "col");
            XElement r = FirstDescendant(point, "row");
            XElement co = FirstDescendant(point, "colOff");
            XElement ro = FirstDescendant(point, "rowOff");
            if (c == null) { c = new XElement(point.Name.Namespace + "col"); point.Add(c); }
            if (r == null) { r = new XElement(point.Name.Namespace + "row"); point.Add(r); }
            if (co == null) { co = new XElement(point.Name.Namespace + "colOff"); point.Add(co); }
            if (ro == null) { ro = new XElement(point.Name.Namespace + "rowOff"); point.Add(ro); }
            c.Value = Math.Max(0, col).ToString(CultureInfo.InvariantCulture);
            r.Value = Math.Max(0, row).ToString(CultureInfo.InvariantCulture);
            co.Value = Math.Max(0, colOff).ToString(CultureInfo.InvariantCulture);
            ro.Value = Math.Max(0, rowOff).ToString(CultureInfo.InvariantCulture);
        }

        private static bool TrySavePreserving(Workbook wb, string path)
        {
            XlsxSourceInfo src = wb.XlsxSource;
            if (src == null || string.IsNullOrEmpty(src.Path)) return false;
            FileInfo sfi = new FileInfo(src.Path);
            if (!sfi.Exists || sfi.Length != src.Length || sfi.LastWriteTimeUtc != src.LastWriteUtc) return false;
            string dir = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!Directory.Exists(dir)) throw new DirectoryNotFoundException("No existe la carpeta de destino.");
            string temp = Path.Combine(dir, "~calcrapido_" + Guid.NewGuid().ToString("N") + ".xlsx");
            StylesAppender appender = null;
            try
            {
                using (FileStream sfs = new FileStream(src.Path, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (ZipArchive sz = new ZipArchive(sfs, ZipArchiveMode.Read))
                {
                    string wbPart = src.WorkbookPart;
                    if (string.IsNullOrEmpty(wbPart)) wbPart = FindWorkbookPath(sz);
                    string wbDir = PartDirectory(wbPart);
                    XDocument wbDoc = ReadXml(sz, wbPart);
                    string wbRelPath = wbDir + "_rels/" + Path.GetFileName(wbPart) + ".rels";
                    XDocument wbRelDoc = ReadXmlIfExists(sz, wbRelPath);
                    XDocument ctDoc = ReadXmlIfExists(sz, "[Content_Types].xml");
                    string stylesPart = src.StylesPart;
                    if (string.IsNullOrEmpty(stylesPart)) stylesPart = wbDir + "styles.xml";
                    XDocument stylesDoc = ReadXmlIfExists(sz, stylesPart);
                    bool needStyles = false;
                    foreach (Sheet sh in wb.Sheets)
                    {
                        foreach (Cell c in sh.Cells.Values)
                        {
                            if (c.StyleDirty || (c.SourceStyleIndex < 0 && !c.EffectiveStyle.IsDefault()))
                            {
                                needStyles = true;
                                break;
                            }
                        }
                        if (needStyles) break;
                    }
                    if (stylesDoc == null) needStyles = true;
                    appender = needStyles ? new StylesAppender(stylesDoc) : null;
                    IStyleIndexer styleIndexer = appender != null ? (IStyleIndexer)appender : new SourceStyleIndexer();

                    Dictionary<string, string> rels = ParseRelationships(wbRelDoc);
                    Dictionary<string, string> partToRid = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (KeyValuePair<string, string> kv in rels)
                        partToRid[ResolvePart(wbDir, kv.Value)] = kv.Key;

                    Dictionary<string, XElement> sheetElByRid = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
                    int maxSheetId = 0;
                    foreach (XElement se in Descendants(wbDoc.Root, "sheet"))
                    {
                        string rid = RelAttr(se, "id");
                        if (rid != null) sheetElByRid[rid] = se;
                        int sid;
                        if (int.TryParse(Attr(se, "sheetId"), out sid) && sid > maxSheetId) maxSheetId = sid;
                    }
                    int maxRid = 0;
                    foreach (string rid in rels.Keys)
                    {
                        if (rid != null && rid.StartsWith("rId", StringComparison.OrdinalIgnoreCase))
                        {
                            int n;
                            if (int.TryParse(rid.Substring(3), out n) && n > maxRid) maxRid = n;
                        }
                    }
                    int nextRid = maxRid + 1;
                    int nextSheetNum = 1;
                    foreach (ZipArchiveEntry e in sz.Entries)
                    {
                        string name = e.FullName.Replace('\\', '/');
                        if (name.StartsWith(wbDir + "worksheets/sheet", StringComparison.OrdinalIgnoreCase) &&
                            name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                        {
                            string num = name.Substring((wbDir + "worksheets/sheet").Length);
                            num = num.Substring(0, num.Length - 4);
                            int n;
                            if (int.TryParse(num, out n) && n >= nextSheetNum) nextSheetNum = n + 1;
                        }
                    }

                    List<SheetSaveInfo> infos = new List<SheetSaveInfo>();
                    foreach (Sheet sh in wb.Sheets)
                    {
                        SheetSaveInfo info = new SheetSaveInfo();
                        info.Sheet = sh;
                        info.New = false;
                        info.Part = sh.SourcePart;
                        if (string.IsNullOrEmpty(info.Part) || GetEntry(sz, info.Part) == null)
                        {
                            info.New = true;
                            while (GetEntry(sz, wbDir + "worksheets/sheet" + nextSheetNum + ".xml") != null) nextSheetNum++;
                            info.Part = wbDir + "worksheets/sheet" + nextSheetNum + ".xml";
                            nextSheetNum++;
                        }
                        string rid;
                        if (!info.New && partToRid.TryGetValue(info.Part, out rid)) info.Rid = rid;
                        else info.Rid = "rId" + nextRid++;
                        XElement old;
                        if (sheetElByRid.TryGetValue(info.Rid, out old))
                        {
                            info.SheetId = Attr(old, "sheetId");
                            info.State = Attr(old, "state");
                        }
                        if (string.IsNullOrEmpty(info.SheetId))
                        {
                            maxSheetId++;
                            info.SheetId = maxSheetId.ToString(CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            int sid;
                            if (int.TryParse(info.SheetId, out sid) && sid > maxSheetId) maxSheetId = sid;
                        }
                        infos.Add(info);
                    }

                    int maxChartPart = ExistingPartNumber(sz, "xl/charts/chart", ".xml");
                    int maxDrawingPart = ExistingPartNumber(sz, "xl/drawings/drawing", ".xml");
                    List<ChartSavePlan> chartPlans = new List<ChartSavePlan>();
                    foreach (SheetSaveInfo info in infos)
                    {
                        List<SheetChart> newCharts = new List<SheetChart>();
                        List<SheetChart> dirtyCharts = new List<SheetChart>();
                        string existingDrawing = null;
                        foreach (SheetChart ch in info.Sheet.Charts)
                        {
                            if (string.IsNullOrEmpty(ch.SourceChartPart)) newCharts.Add(ch);
                            else
                            {
                                if (ch.Dirty) dirtyCharts.Add(ch);
                                if (existingDrawing == null) existingDrawing = ch.SourceDrawingPart;
                            }
                        }
                        if (newCharts.Count == 0 && dirtyCharts.Count == 0) continue;
                        if (string.IsNullOrEmpty(existingDrawing) && !string.IsNullOrEmpty(info.Sheet.SourcePart))
                        {
                            XDocument existingSheet = ReadXmlIfExists(sz, info.Sheet.SourcePart);
                            XElement oldDrawing = existingSheet == null ? null :
                                FirstDescendant(existingSheet.Root, "drawing");
                            if (oldDrawing != null)
                            {
                                string did = RelAttr(oldDrawing, "id");
                                string dtarget;
                                Dictionary<string, string> srmap = ParseRelationships(
                                    ReadXmlIfExists(sz, SheetRelsPath(info.Sheet.SourcePart)), true);
                                if (did != null && srmap.TryGetValue(did, out dtarget))
                                    existingDrawing = ResolvePart(PartDirectory(info.Sheet.SourcePart), dtarget);
                            }
                        }
                        ChartSavePlan plan = new ChartSavePlan();
                        plan.Sheet = info.Sheet;
                        plan.NewDrawing = string.IsNullOrEmpty(existingDrawing);
                        plan.DrawingPart = plan.NewDrawing
                            ? "xl/drawings/drawing" + (++maxDrawingPart) + ".xml" : existingDrawing;
                        plan.DrawingRelsPart = PartDirectory(plan.DrawingPart) + "_rels/" +
                            Path.GetFileName(plan.DrawingPart) + ".rels";
                        plan.Charts.AddRange(newCharts);
                        plan.Rewrites.AddRange(dirtyCharts);
                        for (int i = 0; i < newCharts.Count; i++)
                            plan.ChartNumbers.Add(++maxChartPart);
                        chartPlans.Add(plan);

                        if (ctDoc != null && ctDoc.Root != null)
                        {
                            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
                            HashSet<string> existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            foreach (XElement ov in ctDoc.Root.Elements(ct + "Override"))
                                existing.Add(Attr(ov, "PartName") ?? "");
                            string dp = "/" + plan.DrawingPart.TrimStart('/');
                            if (!existing.Contains(dp))
                                ctDoc.Root.Add(new XElement(ct + "Override", new XAttribute("PartName", dp),
                                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")));
                            foreach (int num in plan.ChartNumbers)
                            {
                                string cp = "/xl/charts/chart" + num + ".xml";
                                if (!existing.Contains(cp))
                                    ctDoc.Root.Add(new XElement(ct + "Override", new XAttribute("PartName", cp),
                                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")));
                            }
                        }
                    }

                    List<ImageSavePlan> imagePlans = new List<ImageSavePlan>();
                    foreach (SheetSaveInfo info in infos)
                    {
                        if (!info.Sheet.ImagesDirty && info.Sheet.DeletedImages.Count == 0) continue;
                        ImageSavePlan ip = new ImageSavePlan();
                        ip.Sheet = info.Sheet;
                        foreach (SheetImage img in info.Sheet.Images)
                            if (img.Dirty && !string.IsNullOrEmpty(img.SourceDrawingPart))
                            {
                                ip.DrawingPart = img.SourceDrawingPart;
                                ip.Images.Add(img);
                            }
                        foreach (SheetImage img in info.Sheet.DeletedImages)
                            if (!string.IsNullOrEmpty(img.SourceDrawingPart))
                            {
                                if (string.IsNullOrEmpty(ip.DrawingPart)) ip.DrawingPart = img.SourceDrawingPart;
                                ip.Deleted.Add(img);
                            }
                        if (!string.IsNullOrEmpty(ip.DrawingPart)) imagePlans.Add(ip);
                    }

                    PatchWorkbook(wbDoc, wbRelDoc, ctDoc, wb, infos, wbDir);
                    string calcChainPart = RemoveCalcChain(wbRelDoc, ctDoc, wbDir);

                    using (FileStream dfs = new FileStream(temp, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                    using (ZipArchive dz = new ZipArchive(dfs, ZipArchiveMode.Create))
                    {
                        HashSet<string> skip = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        skip.Add(wbPart);
                        skip.Add(wbRelPath);
                        skip.Add("[Content_Types].xml");
                        if (stylesPart != null && appender != null) skip.Add(stylesPart);
                        if (!string.IsNullOrEmpty(calcChainPart)) skip.Add(calcChainPart);
                        foreach (SheetSaveInfo info in infos)
                        {
                            if (!info.New)
                            {
                                skip.Add(info.Part);
                                skip.Add(SheetRelsPath(info.Part));
                            }
                        }
                        foreach (ChartSavePlan plan in chartPlans)
                        {
                            foreach (SheetChart rw in plan.Rewrites)
                                if (!string.IsNullOrEmpty(rw.SourceChartPart)) skip.Add(rw.SourceChartPart);
                            if (plan.NewDrawing || plan.Charts.Count == 0) continue;
                            skip.Add(plan.DrawingPart);
                            skip.Add(plan.DrawingRelsPart);
                        }
                        foreach (ImageSavePlan ip in imagePlans)
                        {
                            skip.Add(ip.DrawingPart);
                            foreach (SheetImage img in ip.Images)
                                if (img.Dirty && !string.IsNullOrEmpty(img.SourcePart)) skip.Add(img.SourcePart);
                        }

                        long copiedTotal = 0;
                        foreach (ZipArchiveEntry entry in sz.Entries)
                        {
                            string name = entry.FullName.Replace('\\', '/');
                            if (skip.Contains(name)) continue;
                            if (entry.Length > MaxTotalUncompressedBytes ||
                                copiedTotal + entry.Length > MaxTotalUncompressedBytes)
                                throw new FileFormatException("El contenido del .xlsx es demasiado grande para guardarlo.");
                            TableFormat tableFormat = name.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase)
                                ? FindTableFormat(wb, name) : null;
                            if (tableFormat != null)
                            {
                                XDocument tableDoc = ReadXml(sz, name);
                                UpdateTablePart(tableDoc, tableFormat);
                                WriteXml(dz, name, tableDoc);
                                continue;
                            }
                            ZipArchiveEntry de = dz.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                            using (Stream input = entry.Open())
                            using (Stream output = de.Open())
                                copiedTotal += CopyStreamLimited(input, output, MaxTotalUncompressedBytes - copiedTotal);
                        }

                        foreach (SheetSaveInfo info in infos)
                        {
                            XDocument sheetDoc;
                            XDocument srels = null;
                            bool linkChanged;
                            if (info.New)
                            {
                                XElement links = BuildHyperlinks(info.Sheet, ref srels, out linkChanged);
                                sheetDoc = BuildFreshSheet(info.Sheet, styleIndexer, links, wb.Date1904);
                            }
                            else
                            {
                                sheetDoc = ReadXml(sz, info.Part);
                                srels = ReadXmlIfExists(sz, SheetRelsPath(info.Part));
                                XElement links = BuildHyperlinks(info.Sheet, ref srels, out linkChanged);
                                PatchSheetXml(sheetDoc, info.Sheet, styleIndexer, links, wb.Date1904);
                            }
                            ChartSavePlan plan = null;
                            foreach (ChartSavePlan cp in chartPlans)
                                if (ReferenceEquals(cp.Sheet, info.Sheet)) { plan = cp; break; }
                            if (plan != null)
                            {
                                Dictionary<string, string> sheetRelMap =
                                    ParseRelationships(srels, true);
                                WriteChartPlan(sz, dz, plan, info.Part, sheetDoc, sheetRelMap, ref srels);
                            }
                            WriteXml(dz, info.Part, sheetDoc);
                            if (srels != null && srels.Root != null && srels.Root.HasElements)
                                WriteXml(dz, SheetRelsPath(info.Part), srels);
                        }

                        foreach (ImageSavePlan ip in imagePlans)
                            WriteImagePlan(sz, dz, ip);

                        if (stylesPart != null && appender != null)
                        {
                            if (appender.Changed)
                            {
                                if (stylesDoc != null && stylesDoc.Root != null &&
                                    stylesDoc.Root.Name.Namespace != Ns)
                                    ConvertNamespace(appender.Document.Root, stylesDoc.Root.Name.Namespace);
                                WriteXml(dz, stylesPart, appender.Document);
                            }
                            else if (stylesDoc != null) WriteXml(dz, stylesPart, stylesDoc);
                        }

                        WriteXml(dz, wbPart, wbDoc);
                        WriteXml(dz, wbRelPath, wbRelDoc);
                        if (ctDoc != null) WriteXml(dz, "[Content_Types].xml", ctDoc);
                    }
                }
                ReplaceAtomically(temp, path);
                if (appender != null) appender.Commit();
                MarkSaved(wb);
                return true;
            }
            catch (Exception ex)
            {
                LastSaveError = ex.GetType().Name + ": " + ex.Message;
                if (File.Exists(temp)) { try { File.Delete(temp); } catch { } }
                return false;
            }
        }

        private static void PatchWorkbook(XDocument wbDoc, XDocument wbRelDoc, XDocument ctDoc,
                                          Workbook wb, List<SheetSaveInfo> infos, string wbDir)
        {
            XElement root = wbDoc.Root;
            XNamespace ns = root.Name.Namespace;
            bool strict = ns != Ns;
            string relNs = strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
                : Rns.NamespaceName;
            XElement sheetsEl = ChildByLocalName(root, "sheets");
            if (sheetsEl == null)
            {
                sheetsEl = new XElement(ns + "sheets");
                root.AddFirst(sheetsEl);
            }
            sheetsEl.RemoveNodes();
            foreach (SheetSaveInfo info in infos)
            {
                XElement se = new XElement(ns + "sheet",
                    new XAttribute("name", info.Sheet.Name),
                    new XAttribute("sheetId", info.SheetId),
                    new XAttribute(XNamespace.Get(relNs) + "id", info.Rid));
                if (!string.IsNullOrEmpty(info.State)) se.SetAttributeValue("state", info.State);
                sheetsEl.Add(se);
            }

            XElement view = FirstDescendant(root, "workbookView");
            if (view != null)
                view.SetAttributeValue("activeTab", Math.Max(0, Math.Min(wb.Sheets.Count - 1, wb.ActiveSheetIndex)));

            XElement oldNames = ChildByLocalName(root, "definedNames");
            if (oldNames != null) oldNames.Remove();
            if (wb.DefinedNames.Count > 0)
            {
                XElement names = new XElement(ns + "definedNames");
                foreach (DefinedName d in wb.DefinedNames)
                {
                    XElement el = new XElement(ns + "definedName", new XAttribute("name", d.Name));
                    if (d.LocalSheetId >= 0) el.Add(new XAttribute("localSheetId", d.LocalSheetId));
                    if (d.Hidden) el.Add(new XAttribute("hidden", "1"));
                    if (!string.IsNullOrEmpty(d.Comment)) el.Add(new XAttribute("comment", d.Comment));
                    string formula = d.Formula ?? "";
                    if (formula.StartsWith("=", StringComparison.Ordinal)) formula = formula.Substring(1);
                    el.Value = formula;
                    names.Add(el);
                }
                XElement calc = ChildByLocalName(root, "calcPr");
                if (calc != null) calc.AddBeforeSelf(names); else root.Add(names);
            }

            XElement calcPr = ChildByLocalName(root, "calcPr");
            if (calcPr == null)
            {
                calcPr = new XElement(ns + "calcPr");
                root.Add(calcPr);
            }
            calcPr.SetAttributeValue("fullCalcOnLoad", "1");

            if (wbRelDoc == null)
                wbRelDoc = new XDocument(new XElement(PkgRel + "Relationships"));
            XElement relRoot = wbRelDoc.Root;
            if (relRoot == null)
            {
                relRoot = new XElement(PkgRel + "Relationships");
                wbRelDoc.Add(relRoot);
            }
            HashSet<string> wanted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SheetSaveInfo info in infos) wanted.Add(info.Part);
            List<XElement> removeRels = new List<XElement>();
            foreach (XElement rel in relRoot.Elements(PkgRel + "Relationship"))
            {
                string type = Attr(rel, "Type");
                if (type == null || !type.EndsWith("/worksheet", StringComparison.OrdinalIgnoreCase)) continue;
                string target = Attr(rel, "Target");
                if (target == null) continue;
                string part = ResolvePart(wbDir, target);
                if (!wanted.Contains(part)) removeRels.Add(rel);
            }
            foreach (XElement rel in removeRels) rel.Remove();

            foreach (SheetSaveInfo info in infos)
            {
                XElement existing = null;
                foreach (XElement rel in relRoot.Elements(PkgRel + "Relationship"))
                {
                    if (string.Equals(Attr(rel, "Id"), info.Rid, StringComparison.OrdinalIgnoreCase)) { existing = rel; break; }
                }
                string relTarget = info.Part.StartsWith(wbDir, StringComparison.OrdinalIgnoreCase)
                    ? info.Part.Substring(wbDir.Length) : info.Part;
                if (existing == null)
                {
                    relRoot.Add(new XElement(PkgRel + "Relationship",
                        new XAttribute("Id", info.Rid),
                        new XAttribute("Type", strict
                            ? "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
                            : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                        new XAttribute("Target", relTarget)));
                }
                else
                {
                    existing.SetAttributeValue("Type", strict
                        ? "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
                        : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                    existing.SetAttributeValue("Target", relTarget);
                }
            }

            if (ctDoc != null)
            {
                XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
                XElement ctRoot = ctDoc.Root;
                foreach (SheetSaveInfo info in infos)
                {
                    string partName = "/" + info.Part.TrimStart('/');
                    bool found = false;
                    foreach (XElement ov in ctRoot.Elements(ct + "Override"))
                        if (string.Equals(Attr(ov, "PartName"), partName, StringComparison.OrdinalIgnoreCase)) { found = true; break; }
                    if (!found)
                        ctRoot.Add(new XElement(ct + "Override",
                            new XAttribute("PartName", partName),
                            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
                }
            }
        }

        private static string RemoveCalcChain(XDocument wbRelDoc, XDocument ctDoc, string wbDir)
        {
            if (wbRelDoc == null || wbRelDoc.Root == null) return null;
            XElement chain = null;
            foreach (XElement rel in wbRelDoc.Root.Elements())
            {
                string type = Attr(rel, "Type");
                if (type != null && type.EndsWith("/calcChain", StringComparison.OrdinalIgnoreCase)) { chain = rel; break; }
            }
            if (chain == null) return null;
            string target = Attr(chain, "Target");
            chain.Remove();
            if (string.IsNullOrEmpty(target)) return null;
            string part = ResolvePart(wbDir, target);
            if (ctDoc != null && ctDoc.Root != null)
            {
                XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
                string partName = "/" + part.TrimStart('/');
                List<XElement> remove = new List<XElement>();
                foreach (XElement ov in ctDoc.Root.Elements(ct + "Override"))
                    if (string.Equals(Attr(ov, "PartName"), partName, StringComparison.OrdinalIgnoreCase)) remove.Add(ov);
                foreach (XElement ov in remove) ov.Remove();
            }
            return part;
        }

        private static XElement BuildSheetData(Sheet sheet, IStyleIndexer styles, XDocument sourceSheet, bool date1904)
        {
            Dictionary<int, List<KeyValuePair<int, Cell>>> byRow = new Dictionary<int, List<KeyValuePair<int, Cell>>>();
            foreach (KeyValuePair<long, Cell> kv in sheet.Cells)
            {
                int row = (int)(kv.Key >> 14);
                int col = (int)(kv.Key & 0x3FFF);
                List<KeyValuePair<int, Cell>> list;
                if (!byRow.TryGetValue(row, out list)) { list = new List<KeyValuePair<int, Cell>>(); byRow[row] = list; }
                list.Add(new KeyValuePair<int, Cell>(col, kv.Value));
            }
            foreach (int r in sheet.RowHeights.Keys) if (!byRow.ContainsKey(r)) byRow[r] = new List<KeyValuePair<int, Cell>>();
            foreach (int r in sheet.HiddenRows) if (!byRow.ContainsKey(r)) byRow[r] = new List<KeyValuePair<int, Cell>>();

            Dictionary<int, XElement> sourceRows = new Dictionary<int, XElement>();
            if (sourceSheet != null && sourceSheet.Root != null && !sheet.RowsDirty && !sheet.StructureDirty)
            {
                foreach (XElement rowEl in Descendants(sourceSheet.Root, "row"))
                {
                    int r;
                    if (int.TryParse(Attr(rowEl, "r"), out r) && r > 0) sourceRows[r - 1] = rowEl;
                }
            }

            List<int> rows = new List<int>(byRow.Keys);
            rows.Sort();
            XElement sheetData = new XElement(Ns + "sheetData");
            foreach (int r in rows)
            {
                XElement row = new XElement(Ns + "row", new XAttribute("r", r + 1));
                XElement srcRow;
                if (sourceRows.TryGetValue(r, out srcRow))
                {
                    foreach (XAttribute a in srcRow.Attributes())
                        if (!string.Equals(a.Name.LocalName, "r", StringComparison.OrdinalIgnoreCase)) row.SetAttributeValue(a.Name, a.Value);
                }
                else
                {
                    float ht;
                    if (sheet.RowHeights.TryGetValue(r, out ht))
                    {
                        row.SetAttributeValue("ht", ht.ToString("0.##", CultureInfo.InvariantCulture));
                        row.SetAttributeValue("customHeight", "1");
                    }
                    if (sheet.HiddenRows.Contains(r)) row.SetAttributeValue("hidden", "1");
                }
                List<KeyValuePair<int, Cell>> cells = byRow[r];
                cells.Sort(delegate(KeyValuePair<int, Cell> a, KeyValuePair<int, Cell> b) { return a.Key.CompareTo(b.Key); });
                foreach (KeyValuePair<int, Cell> pair in cells)
                {
                    XElement c = BuildCell(pair.Key, r, pair.Value, styles, date1904);
                    if (c != null) row.Add(c);
                }
                if (row.HasElements || row.Attribute("ht") != null || row.Attribute("hidden") != null || srcRow != null)
                    sheetData.Add(row);
            }
            return sheetData;
        }

        private static bool IsDateStyle(CellStyle st)
        {
            if (st == null) return false;
            if (st.Format == NumberFormatKind.Date || st.Format == NumberFormatKind.Time) return true;
            string code = (st.CustomFormat ?? "").ToLowerInvariant();
            return code.IndexOf("yy", StringComparison.Ordinal) >= 0 ||
                   code.IndexOf("dd", StringComparison.Ordinal) >= 0 ||
                   code.IndexOf("hh", StringComparison.Ordinal) >= 0 ||
                   code.IndexOf("ss", StringComparison.Ordinal) >= 0 ||
                   code.IndexOf("mmm", StringComparison.Ordinal) >= 0 ||
                   code.IndexOf("d/m", StringComparison.Ordinal) >= 0 ||
                   code.IndexOf("m/d", StringComparison.Ordinal) >= 0;
        }

        private static void AdjustDate1904(Workbook wb, double delta)
        {
            foreach (Sheet sheet in wb.Sheets)
            {
                foreach (Cell cell in sheet.Cells.Values)
                {
                    if (cell.IsFormula || cell.Value.Kind != ValueKind.Number) continue;
                    if (!IsDateStyle(cell.EffectiveStyle)) continue;
                    cell.Value = CellValue.FromNumber(cell.Value.Number + delta);
                }
            }
        }

        private static XElement BuildCell(int col, int row, Cell cell, IStyleIndexer styles, bool date1904)
        {
            if (cell == null) return null;
            int styleIndex = styles.IndexFor(cell);
            bool hasValue = cell.IsFormula || cell.Value.Kind != ValueKind.Empty;
            if (!hasValue && styleIndex == 0 && string.IsNullOrEmpty(cell.Hyperlink)) return null;
            XElement c = new XElement(Ns + "c", new XAttribute("r", AddressUtil.CellName(row, col)));
            if (styleIndex > 0) c.SetAttributeValue("s", styleIndex);
            if (cell.IsFormula)
            {
                string f = null;
                if (!cell.FormulaDirty && !string.IsNullOrEmpty(cell.SourceFormulaText)) f = cell.SourceFormulaText;
                if (string.IsNullOrEmpty(f)) f = FormulaExport.ToExcel(cell.Formula);
                if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1);
                XElement fe = new XElement(Ns + "f", f);
                if (!cell.FormulaDirty)
                {
                    if (!string.IsNullOrEmpty(cell.SourceFormulaType))
                        fe.SetAttributeValue("t", cell.SourceFormulaType);
                    if (!string.IsNullOrEmpty(cell.SourceFormulaRef))
                        fe.SetAttributeValue("ref", cell.SourceFormulaRef);
                    if (cell.SourceFormulaAca) fe.SetAttributeValue("aca", "1");
                    if (cell.SourceFormulaBx) fe.SetAttributeValue("bx", "1");
                }
                c.Add(fe);
                AddCachedValue(c, cell.Value);
                return c;
            }
            switch (cell.Value.Kind)
            {
                case ValueKind.Text:
                    c.SetAttributeValue("t", "inlineStr");
                    c.Add(new XElement(Ns + "is", new XElement(Ns + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"), XmlSafe(cell.Value.Text))));
                    break;
                case ValueKind.Number:
                    double number = cell.Value.Number;
                    if (date1904 && IsDateStyle(cell.EffectiveStyle)) number -= 1462;
                    c.Add(new XElement(Ns + "v", number.ToString("G17", CultureInfo.InvariantCulture)));
                    break;
                case ValueKind.Bool:
                    c.SetAttributeValue("t", "b");
                    c.Add(new XElement(Ns + "v", cell.Value.Bool ? "1" : "0"));
                    break;
                case ValueKind.Error:
                    c.SetAttributeValue("t", "e");
                    c.Add(new XElement(Ns + "v", ToEnglishError(cell.Value.Error)));
                    break;
            }
            return c;
        }

        private static XElement BuildCols(Sheet sheet)
        {
            int maxCol = Math.Max(0, sheet.UsedMaxCol);
            foreach (int c in sheet.ColWidths.Keys) if (c > maxCol) maxCol = c;
            foreach (int c in sheet.HiddenCols) if (c > maxCol) maxCol = c;
            foreach (MergeRange m in sheet.Merges) if (m.Col2 > maxCol) maxCol = m.Col2;
            XElement cols = new XElement(Ns + "cols");
            for (int c = 0; c <= maxCol && c < Limits.MaxCols; c++)
            {
                bool hasWidth = sheet.ColWidths.ContainsKey(c);
                bool hidden = sheet.HiddenCols.Contains(c);
                if (!hasWidth && !hidden) continue;
                float px = sheet.GetColWidth(c);
                double width = Math.Max(3.0, Math.Min(100.0, px / 7.5));
                XElement col = new XElement(Ns + "col",
                    new XAttribute("min", c + 1), new XAttribute("max", c + 1),
                    new XAttribute("width", width.ToString("0.##", CultureInfo.InvariantCulture)),
                    new XAttribute("customWidth", "1"));
                if (hidden) col.SetAttributeValue("hidden", "1");
                cols.Add(col);
            }
            return cols.HasElements ? cols : null;
        }

        private static XElement BuildHyperlinks(Sheet sheet, ref XDocument relsDoc, out bool relsChanged)
        {
            relsChanged = false;
            if (relsDoc == null || relsDoc.Root == null)
                relsDoc = new XDocument(new XElement(PkgRel + "Relationships"));
            XElement root = relsDoc.Root;
            Dictionary<string, string> existingByTarget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            List<XElement> remove = new List<XElement>();
            foreach (XElement rel in root.Elements(PkgRel + "Relationship"))
            {
                string type = Attr(rel, "Type");
                if (type != null && type.EndsWith("/hyperlink", StringComparison.OrdinalIgnoreCase))
                {
                    string target = Attr(rel, "Target");
                    string id = Attr(rel, "Id");
                    if (target != null && id != null) existingByTarget[target] = id;
                    remove.Add(rel);
                }
            }
            foreach (XElement rel in remove) { rel.Remove(); relsChanged = true; }

            int maxRid = 0;
            foreach (XElement rel in root.Elements(PkgRel + "Relationship"))
            {
                string id = Attr(rel, "Id");
                if (id != null && id.StartsWith("rId", StringComparison.OrdinalIgnoreCase))
                {
                    int n;
                    if (int.TryParse(id.Substring(3), out n) && n > maxRid) maxRid = n;
                }
            }
            int nextRid = maxRid + 1;
            Dictionary<string, string> used = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            XElement hyperlinks = new XElement(Ns + "hyperlinks");
            foreach (KeyValuePair<long, Cell> kv in sheet.Cells)
            {
                Cell cell = kv.Value;
                if (string.IsNullOrEmpty(cell.Hyperlink)) continue;
                int row = (int)(kv.Key >> 14);
                int col = (int)(kv.Key & 0x3FFF);
                string target = cell.Hyperlink;
                bool internalLink = cell.HyperlinkInternal || target.StartsWith("#", StringComparison.Ordinal);
                if (target.StartsWith("#", StringComparison.Ordinal)) target = target.Substring(1);
                XElement link = new XElement(Ns + "hyperlink", new XAttribute("ref", AddressUtil.CellName(row, col)));
                if (!string.IsNullOrEmpty(cell.HyperlinkTooltip)) link.SetAttributeValue("tooltip", cell.HyperlinkTooltip);
                if (internalLink)
                {
                    link.SetAttributeValue("location", target);
                }
                else
                {
                    string rid;
                    if (!used.TryGetValue(target, out rid))
                    {
                        if (!existingByTarget.TryGetValue(target, out rid)) rid = "rId" + nextRid++;
                        root.Add(new XElement(PkgRel + "Relationship",
                            new XAttribute("Id", rid),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"),
                            new XAttribute("Target", target),
                            new XAttribute("TargetMode", "External")));
                        relsChanged = true;
                        used[target] = rid;
                    }
                    link.SetAttributeValue(Rns + "id", rid);
                }
                hyperlinks.Add(link);
            }
            if (!hyperlinks.HasElements) return null;
            return hyperlinks;
        }

        private static XElement BuildMerges(Sheet sheet)
        {
            List<MergeRange> valid = new List<MergeRange>();
            foreach (MergeRange m in sheet.Merges)
            {
                if (m.Row2 == m.Row1 && m.Col2 == m.Col1) continue;
                if (m.Row2 >= Limits.MaxRows || m.Col2 >= Limits.MaxCols) continue;
                valid.Add(m);
            }
            XElement merges = new XElement(Ns + "mergeCells", new XAttribute("count", valid.Count));
            foreach (MergeRange m in valid)
                merges.Add(new XElement(Ns + "mergeCell", new XAttribute("ref", m.ToRangeString())));
            if (merges.Elements().Count() == 0) return null;
            return merges;
        }

        private static XDocument BuildFreshSheet(Sheet sheet, IStyleIndexer styles, XElement hyperlinks, bool date1904)
        {
            int maxRow = Math.Max(0, sheet.UsedMaxRow);
            int maxCol = Math.Max(0, sheet.UsedMaxCol);
            foreach (MergeRange m in sheet.Merges)
            {
                if (m.Row2 > maxRow) maxRow = m.Row2;
                if (m.Col2 > maxCol) maxCol = m.Col2;
            }
            XElement root = new XElement(Ns + "worksheet",
                new XAttribute(XNamespace.Xmlns + "r", Rns.NamespaceName),
                new XElement(Ns + "dimension", new XAttribute("ref", "A1:" + AddressUtil.CellName(maxRow, maxCol))),
                new XElement(Ns + "sheetViews",
                    BuildSheetView(sheet)),
                new XElement(Ns + "sheetFormatPr", new XAttribute("defaultRowHeight", "15")));
            XElement cols = BuildCols(sheet);
            if (cols != null) root.Add(cols);
            root.Add(BuildSheetData(sheet, styles, null, date1904));
            XElement validations = BuildDataValidations(sheet);
            if (validations != null) root.Add(validations);
            XElement autoFilter = BuildAutoFilter(sheet);
            if (autoFilter != null) root.Add(autoFilter);
            XElement merges = BuildMerges(sheet);
            if (merges != null) root.Add(merges);
            if (hyperlinks != null) root.Add(hyperlinks);
            root.Add(new XElement(Ns + "pageMargins", new XAttribute("left", "0.7"), new XAttribute("right", "0.7"),
                new XAttribute("top", "0.75"), new XAttribute("bottom", "0.75"),
                new XAttribute("header", "0.3"), new XAttribute("footer", "0.3")));
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        }

        private static XElement BuildSheetView(Sheet sheet)
        {
            XElement view = new XElement(Ns + "sheetView", new XAttribute("workbookViewId", "0"));
            if (sheet.FreezeRows > 0 || sheet.FreezeCols > 0)
            {
                XElement pane = new XElement(Ns + "pane",
                    new XAttribute("state", "frozen"),
                    new XAttribute("topLeftCell", AddressUtil.CellName(sheet.FreezeRows, sheet.FreezeCols)),
                    new XAttribute("activePane", "bottomRight"));
                if (sheet.FreezeCols > 0) pane.SetAttributeValue("xSplit", sheet.FreezeCols);
                if (sheet.FreezeRows > 0) pane.SetAttributeValue("ySplit", sheet.FreezeRows);
                view.Add(pane);
            }
            return view;
        }

        private static XElement ChildByLocalName(XElement parent, string localName)
        {
            if (parent == null) return null;
            foreach (XElement e in parent.Elements())
                if (string.Equals(e.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase)) return e;
            return null;
        }

        private static void PatchSheetXml(XDocument doc, Sheet sheet, IStyleIndexer styles, XElement hyperlinks, bool date1904)
        {
            if (doc == null || doc.Root == null) return;
            XElement root = doc.Root;
            XNamespace targetNs = root.Name.Namespace;
            bool strict = targetNs != Ns;
            if (strict && targetNs.NamespaceName.IndexOf("purl.oclc.org/ooxml/spreadsheetml",
                    StringComparison.OrdinalIgnoreCase) < 0)
                throw new FileFormatException("El .xlsx usa un namespace no compatible.");
            XElement oldData = ChildByLocalName(root, "sheetData");
            XElement newData = BuildSheetData(sheet, styles, doc, date1904);
            if (strict) ConvertNamespace(newData, targetNs);
            if (oldData != null) oldData.ReplaceWith(newData);
            else
            {
                XElement fmt = ChildByLocalName(root, "sheetFormatPr");
                if (fmt != null) fmt.AddAfterSelf(newData); else root.AddFirst(newData);
            }
            XElement oldValidations = ChildByLocalName(root, "dataValidations");
            XElement newValidations = BuildDataValidations(sheet);
            if (strict && newValidations != null) ConvertNamespace(newValidations, targetNs);
            if (newValidations != null)
            {
                if (oldValidations != null) oldValidations.ReplaceWith(newValidations);
                else newData.AddAfterSelf(newValidations);
            }
            XElement oldFilter = ChildByLocalName(root, "autoFilter");
            XElement newFilter = BuildAutoFilter(sheet);
            if (strict && newFilter != null) ConvertNamespace(newFilter, targetNs);
            if (newFilter != null)
            {
                if (oldFilter != null) oldFilter.ReplaceWith(newFilter);
                else newData.AddAfterSelf(newFilter);
            }
            else if (oldFilter != null) oldFilter.Remove();

            int maxRow = Math.Max(0, sheet.UsedMaxRow);
            int maxCol = Math.Max(0, sheet.UsedMaxCol);
            foreach (MergeRange m in sheet.Merges)
            {
                if (m.Row2 > maxRow) maxRow = m.Row2;
                if (m.Col2 > maxCol) maxCol = m.Col2;
            }
            XElement dim = ChildByLocalName(root, "dimension");
            string dimRef = "A1:" + AddressUtil.CellName(maxRow, maxCol);
            if (dim == null)
            {
                dim = new XElement(Ns + "dimension", new XAttribute("ref", dimRef));
                if (strict) ConvertNamespace(dim, targetNs);
                root.AddFirst(dim);
            }
            else dim.SetAttributeValue("ref", dimRef);

            if (sheet.WidthsDirty)
            {
                XElement oldCols = ChildByLocalName(root, "cols");
                XElement newCols = BuildCols(sheet);
                if (strict && newCols != null) ConvertNamespace(newCols, targetNs);
                if (oldCols != null)
                {
                    if (newCols != null) oldCols.ReplaceWith(newCols);
                    else oldCols.Remove();
                }
                else if (newCols != null)
                {
                    XElement data = ChildByLocalName(root, "sheetData");
                    if (data != null) data.AddBeforeSelf(newCols); else root.Add(newCols);
                }
            }

            if (sheet.MergesDirty || sheet.StructureDirty)
            {
                XElement oldMerges = ChildByLocalName(root, "mergeCells");
                if (oldMerges != null) oldMerges.Remove();
                XElement merges = BuildMerges(sheet);
                if (strict && merges != null) ConvertNamespace(merges, targetNs);
                if (merges != null)
                {
                    XElement data = ChildByLocalName(root, "sheetData");
                    if (data != null) data.AddAfterSelf(merges); else root.Add(merges);
                }
            }

            XElement oldLinks = ChildByLocalName(root, "hyperlinks");
            if (oldLinks != null) oldLinks.Remove();
            if (hyperlinks != null)
            {
                if (strict) ConvertNamespace(hyperlinks, targetNs);
                string relNs = strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
                    : Rns.NamespaceName;
                root.SetAttributeValue(XNamespace.Xmlns + "r", relNs);
                XElement anchor = null;
                string[] names = new string[] { "printOptions", "pageMargins", "pageSetup", "headerFooter",
                    "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
                    "webPublishItems", "tableParts", "extLst" };
                foreach (XElement el in root.Elements())
                {
                    for (int i = 0; i < names.Length; i++)
                        if (string.Equals(el.Name.LocalName, names[i], StringComparison.OrdinalIgnoreCase)) { anchor = el; break; }
                    if (anchor != null) break;
                }
                if (anchor != null) anchor.AddBeforeSelf(hyperlinks); else root.Add(hyperlinks);
            }
        }

        private static void ConvertNamespace(XElement element, XNamespace target)
        {
            if (element == null) return;
            XNamespace strictRel = "http://purl.oclc.org/ooxml/officeDocument/relationships";
            foreach (XElement e in element.DescendantsAndSelf())
            {
                if (e.Name.Namespace == Ns) e.Name = target + e.Name.LocalName;
                foreach (XAttribute a in e.Attributes().ToList())
                {
                    if (a.Name.Namespace == Rns)
                    {
                        string local = a.Name.LocalName;
                        string value = a.Value;
                        a.Remove();
                        e.SetAttributeValue(strictRel + local, value);
                    }
                }
            }
        }

        private static void MarkSaved(Workbook wb)
        {
            foreach (Sheet sheet in wb.Sheets)
            {
                sheet.WidthsDirty = false;
                sheet.RowsDirty = false;
                sheet.MergesDirty = false;
                sheet.HyperlinksDirty = false;
                sheet.StructureDirty = false;
                sheet.ChartsDirty = false;
                foreach (SheetChart ch in sheet.Charts) ch.Dirty = false;
                foreach (Cell cell in sheet.Cells.Values)
                {
                    if (cell.FormulaDirty)
                    {
                        cell.SourceFormulaText = cell.IsFormula ? FormulaExport.ToExcel(cell.Formula) : null;
                    }
                    cell.Dirty = false;
                    cell.FormulaDirty = false;
                    cell.StyleDirty = false;
                }
            }
        }

        private static void RefreshSourceInfo(Workbook wb, string path)
        {
            try
            {
                FileInfo fi = new FileInfo(path);
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    string wbPart = FindWorkbookPath(zip);
                    XDocument wbDoc = ReadXml(zip, wbPart);
                    string wbDir = PartDirectory(wbPart);
                    XDocument relDoc = ReadXmlIfExists(zip, wbDir + "_rels/" + Path.GetFileName(wbPart) + ".rels");
                    Dictionary<string, string> rels = ParseRelationships(relDoc);
                    XlsxSourceInfo info = new XlsxSourceInfo();
                    info.Path = Path.GetFullPath(path);
                    info.Length = fi.Length;
                    info.LastWriteUtc = fi.LastWriteTimeUtc;
                    info.WorkbookPart = wbPart;
                    info.StylesPart = ResolvePart(wbDir, RelTargetByType(relDoc, "/styles") ?? "styles.xml");
                    info.ThemePart = ResolvePart(wbDir, RelTargetByType(relDoc, "/theme") ?? "theme/theme1.xml");
                    info.SharedStringsPart = ResolvePart(wbDir, RelTargetByType(relDoc, "/sharedStrings") ?? "sharedStrings.xml");
                    info.Date1904 = IsTrue(Attr(FirstDescendant(wbDoc.Root, "workbookPr"), "date1904"));
                    int i = 0;
                    foreach (XElement se in Descendants(wbDoc.Root, "sheet"))
                    {
                        string rid = RelAttr(se, "id");
                        string target;
                        if (i < wb.Sheets.Count && rid != null && rels.TryGetValue(rid, out target))
                            wb.Sheets[i].SourcePart = ResolvePart(wbDir, target);
                        i++;
                    }
                    wb.XlsxSource = info;
                    wb.Date1904 = info.Date1904;
                }
                MarkSaved(wb);
            }
            catch
            {
                // Si no se pudo releer el resultado, no dejamos una fuente vieja
                // que pudiera usarse para una preservación incorrecta.
                wb.XlsxSource = null;
            }
        }

        internal static bool IsSafeRasterImage(byte[] b)
        {
            if (b == null || b.Length < 30) return false;
            int w = 0, h = 0;
            if (b[0] == 0x89 && b[1] == 0x50 && b[2] == 0x4E && b[3] == 0x47)
            {
                w = (b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19];
                h = (b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23];
            }
            else if (b[0] == 'G' && b[1] == 'I' && b[2] == 'F' && b[3] == '8')
            {
                w = b[6] | (b[7] << 8);
                h = b[8] | (b[9] << 8);
            }
            else if (b[0] == 'B' && b[1] == 'M')
            {
                w = b[18] | (b[19] << 8) | (b[20] << 16) | (b[21] << 24);
                h = Math.Abs(b[22] | (b[23] << 8) | (b[24] << 16) | (b[25] << 24));
            }
            else if (b[0] == 0xFF && b[1] == 0xD8)
            {
                int p = 2;
                while (p + 9 < b.Length)
                {
                    if (b[p] != 0xFF) { p++; continue; }
                    int marker = b[p + 1];
                    if (marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7))
                    {
                        p += 2;
                        continue;
                    }
                    int len = (b[p + 2] << 8) | b[p + 3];
                    if (len < 2 || p + 2 + len > b.Length) break;
                    bool sof = (marker >= 0xC0 && marker <= 0xC3) ||
                               (marker >= 0xC5 && marker <= 0xC7) ||
                               (marker >= 0xC9 && marker <= 0xCB) ||
                               (marker >= 0xCD && marker <= 0xCF);
                    if (sof && p + 9 < b.Length)
                    {
                        h = (b[p + 5] << 8) | b[p + 6];
                        w = (b[p + 7] << 8) | b[p + 8];
                        break;
                    }
                    p += 2 + len;
                }
            }
            else return false;
            return w > 0 && h > 0 && w <= MaxImageDimension && h <= MaxImageDimension &&
                (long)w * h <= MaxImagePixels;
        }

        private static byte[] ReadEntryBytes(ZipArchiveEntry entry, long max)
        {
            using (Stream input = entry.Open())
            using (MemoryStream output = new MemoryStream())
            {
                byte[] buffer = new byte[81920];
                long total = 0;
                int n;
                while ((n = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    total += n;
                    if (total > max)
                        throw new FileFormatException("Una parte interna del .xlsx se expande demasiado.");
                    if (_readBudgetRemaining >= 0)
                    {
                        if (n > _readBudgetRemaining)
                            throw new FileFormatException("El paquete .xlsx consume demasiados datos descomprimidos.");
                        _readBudgetRemaining -= n;
                    }
                    output.Write(buffer, 0, n);
                }
                return output.ToArray();
            }
        }

        private static long CopyStreamLimited(Stream input, Stream output, long max)
        {
            byte[] buffer = new byte[81920];
            long total = 0;
            int n;
            while ((n = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                total += n;
                if (total > max)
                    throw new FileFormatException("Una parte del .xlsx se expande demasiado al guardar.");
                output.Write(buffer, 0, n);
            }
            return total;
        }

        private static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[81920];
            int n;
            while ((n = input.Read(buffer, 0, buffer.Length)) > 0) output.Write(buffer, 0, n);
        }

        private static void ReplaceAtomically(string temp, string path)
        {
            IOException last = null;
            for (int attempt = 0; attempt < 6; attempt++)
            {
                try
                {
                    if (File.Exists(path)) File.Replace(temp, path, null);
                    else File.Move(temp, path);
                    return;
                }
                catch (IOException ex)
                {
                    last = ex;
                    System.Threading.Thread.Sleep(80 * (attempt + 1));
                }
            }
            // Algunos antivirus/indexadores mantienen el destino abierto un
            // instante. Hacemos una copia con respaldo y restauramos si falla.
            string backup = path + "." + Guid.NewGuid().ToString("N") + ".bak";
            bool hadDestination = File.Exists(path);
            try
            {
                if (hadDestination) File.Copy(path, backup, true);
                try
                {
                    File.Copy(temp, path, true);
                    File.Delete(temp);
                }
                catch
                {
                    if (hadDestination && File.Exists(backup)) File.Copy(backup, path, true);
                    throw;
                }
            }
            catch
            {
                if (last != null) throw last;
                throw;
            }
            finally
            {
                try { if (File.Exists(backup)) File.Delete(backup); } catch { }
            }
        }

        public static void SaveStrict(Workbook wb, string path)
        {
            string owner = Guid.NewGuid().ToString("N");
            string temp = path + "." + owner + ".transitional.xlsx";
            string strict = path + "." + owner + ".strict.xlsx";
            try
            {
                LastSaveError = null;
                Save(wb, temp);
                ConvertPackageToStrict(temp, strict);
                ReplaceAtomically(strict, path);
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { }
                try { if (File.Exists(strict)) File.Delete(strict); } catch { }
            }
        }

        private static void ConvertPackageToStrict(string source, string destination)
        {
            string dir = Path.GetDirectoryName(Path.GetFullPath(destination));
            string temp = Path.Combine(dir, "~calcrapido_" + Guid.NewGuid().ToString("N") + ".strict.tmp");
            try
            {
                using (ZipArchive input = ZipFile.OpenRead(source))
                using (FileStream output = new FileStream(temp, FileMode.Create, FileAccess.Write))
                using (ZipArchive zip = new ZipArchive(output, ZipArchiveMode.Create))
                {
                    foreach (ZipArchiveEntry entry in input.Entries)
                    {
                        ZipArchiveEntry target = zip.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                        using (Stream si = entry.Open())
                        using (Stream so = target.Open())
                        {
                            string ext = Path.GetExtension(entry.FullName).ToLowerInvariant();
                            if (ext == ".xml" || ext == ".rels")
                            {
                                using (StreamReader reader = new StreamReader(si, Encoding.UTF8, true))
                                {
                                    string text = reader.ReadToEnd();
                                    text = text.Replace(
                                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                                        "http://purl.oclc.org/ooxml/spreadsheetml/main");
                                    text = text.Replace(
                                        "http://schemas.openxmlformats.org/drawingml/2006/main",
                                        "http://purl.oclc.org/ooxml/drawingml/main");
                                    text = text.Replace(
                                        "http://schemas.openxmlformats.org/drawingml/2006/chart",
                                        "http://purl.oclc.org/ooxml/drawingml/chart");
                                    text = text.Replace(
                                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                                        "http://purl.oclc.org/ooxml/officeDocument/relationships");
                                    byte[] data = new UTF8Encoding(false).GetBytes(text);
                                    so.Write(data, 0, data.Length);
                                }
                            }
                            else
                            {
                                si.CopyTo(so);
                            }
                        }
                    }
                }
                if (File.Exists(destination))
                    throw new IOException("El destino estricto ya existe.");
                File.Move(temp, destination);
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { }
            }
        }

        public static void SaveNew(Workbook wb, string path)
        {
            _readBudgetRemaining = -1;
            string dir = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!Directory.Exists(dir)) throw new DirectoryNotFoundException("No existe la carpeta de destino.");
            string temp = Path.Combine(dir, "~calcrapido_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                using (FileStream fs = new FileStream(temp, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    WritePackage(zip, wb);
                }
                ReplaceAtomically(temp, path);
            }
            finally
            {
                if (File.Exists(temp)) { try { File.Delete(temp); } catch { } }
            }
        }

        private static void WritePackage(ZipArchive zip, Workbook wb)
        {
            List<Sheet> sheets = wb.Sheets;
            List<string> names = new List<string>();
            HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Sheet s in sheets)
            {
                string n = SanitizeSheetName(s.Name);
                string baseName = n;
                int k = 2;
                while (used.Contains(n)) { n = baseName + " (" + k + ")"; k++; }
                used.Add(n);
                names.Add(n);
            }

            StyleCollector styles = new StyleCollector(wb);

            int drawingSequence = 0;
            int chartSequence = 0;
            Dictionary<Sheet, int> drawingNumbers = new Dictionary<Sheet, int>();
            Dictionary<Sheet, List<int>> chartNumbers = new Dictionary<Sheet, List<int>>();
            for (int i = 0; i < sheets.Count; i++)
            {
                if (sheets[i].Charts.Count == 0) continue;
                drawingNumbers[sheets[i]] = ++drawingSequence;
                List<int> nums = new List<int>();
                for (int c = 0; c < sheets[i].Charts.Count; c++) nums.Add(++chartSequence);
                chartNumbers[sheets[i]] = nums;
            }

            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            XElement contentTypes = new XElement(ct + "Types",
                new XElement(ct + "Default", new XAttribute("Extension", "rels"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                new XElement(ct + "Override", new XAttribute("PartName", "/xl/workbook.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")),
                new XElement(ct + "Override", new XAttribute("PartName", "/xl/styles.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")));
            for (int i = 0; i < sheets.Count; i++)
                contentTypes.Add(new XElement(ct + "Override", new XAttribute("PartName", "/xl/worksheets/sheet" + (i + 1) + ".xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            for (int i = 1; i <= drawingSequence; i++)
            {
                contentTypes.Add(new XElement(ct + "Override",
                    new XAttribute("PartName", "/xl/drawings/drawing" + i + ".xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")));
            }
            for (int i = 1; i <= chartSequence; i++)
            {
                contentTypes.Add(new XElement(ct + "Override",
                    new XAttribute("PartName", "/xl/charts/chart" + i + ".xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")));
            }
            WriteXml(zip, "[Content_Types].xml", new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), contentTypes));

            XElement rootRels = new XElement(PkgRel + "Relationships",
                new XElement(PkgRel + "Relationship", new XAttribute("Id", "rId1"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                    new XAttribute("Target", "xl/workbook.xml")));
            WriteXml(zip, "_rels/.rels", new XDocument(rootRels));

            XElement wbRels = new XElement(PkgRel + "Relationships");
            XElement sheetElements = new XElement(Ns + "sheets");
            for (int i = 0; i < sheets.Count; i++)
            {
                int ridNum = i + 1;
                wbRels.Add(new XElement(PkgRel + "Relationship", new XAttribute("Id", "rId" + ridNum),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                    new XAttribute("Target", "worksheets/sheet" + ridNum + ".xml")));
                sheetElements.Add(new XElement(Ns + "sheet", new XAttribute("name", names[i]),
                    new XAttribute("sheetId", ridNum), new XAttribute(Rns + "id", "rId" + ridNum)));
            }
            wbRels.Add(new XElement(PkgRel + "Relationship", new XAttribute("Id", "rIdStyles"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"),
                new XAttribute("Target", "styles.xml")));
            WriteXml(zip, "xl/_rels/workbook.xml.rels", new XDocument(wbRels));

            XElement workbookPr = new XElement(Ns + "workbookPr", new XAttribute("defaultThemeVersion", "164011"));
            if (wb.Date1904) workbookPr.SetAttributeValue("date1904", "1");
            XElement definedNames = null;
            if (wb.DefinedNames.Count > 0)
            {
                definedNames = new XElement(Ns + "definedNames");
                foreach (DefinedName d in wb.DefinedNames)
                {
                    XElement dn = new XElement(Ns + "definedName", new XAttribute("name", d.Name));
                    if (d.LocalSheetId >= 0) dn.Add(new XAttribute("localSheetId", d.LocalSheetId));
                    if (d.Hidden) dn.Add(new XAttribute("hidden", "1"));
                    if (!string.IsNullOrEmpty(d.Comment)) dn.Add(new XAttribute("comment", d.Comment));
                    string formula = d.Formula ?? "";
                    if (formula.StartsWith("=", StringComparison.Ordinal)) formula = formula.Substring(1);
                    dn.Value = formula;
                    definedNames.Add(dn);
                }
            }
            XElement workbook = new XElement(Ns + "workbook",
                new XAttribute(XNamespace.Xmlns + "r", Rns.NamespaceName),
                workbookPr,
                new XElement(Ns + "bookViews",
                    new XElement(Ns + "workbookView", new XAttribute("activeTab", Math.Max(0, Math.Min(sheets.Count - 1, wb.ActiveSheetIndex))))),
                sheetElements);
            if (definedNames != null) workbook.Add(definedNames);
            workbook.Add(new XElement(Ns + "calcPr", new XAttribute("calcId", "191029"), new XAttribute("fullCalcOnLoad", "1")));
            WriteXml(zip, "xl/workbook.xml", new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), workbook));

            WriteXml(zip, "xl/styles.xml", new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), styles.BuildStylesXml()));
            for (int i = 0; i < sheets.Count; i++)
            {
                string sheetPart = "xl/worksheets/sheet" + (i + 1) + ".xml";
                XDocument srels = null;
                bool linkChanged;
                XElement links = BuildHyperlinks(sheets[i], ref srels, out linkChanged);
                XDocument sheetDoc = BuildFreshSheet(sheets[i], styles, links, wb.Date1904);
                int drawingNo;
                List<int> chartNos;
                if (drawingNumbers.TryGetValue(sheets[i], out drawingNo) &&
                    chartNumbers.TryGetValue(sheets[i], out chartNos))
                {
                    AddDrawingRelationship(sheetDoc, ref srels,
                        "../drawings/drawing" + drawingNo + ".xml", "rIdDrawing");
                    List<string> chartRels = new List<string>();
                    for (int c = 0; c < chartNos.Count; c++)
                    {
                        chartRels.Add("../charts/chart" + chartNos[c] + ".xml");
                        WriteXml(zip, "xl/charts/chart" + chartNos[c] + ".xml",
                            ChartWriter.ChartXml(sheets[i].Charts[c]));
                    }
                    WriteXml(zip, "xl/drawings/drawing" + drawingNo + ".xml",
                        ChartWriter.DrawingXml(sheets[i].Charts));
                    WriteXml(zip, "xl/drawings/_rels/drawing" + drawingNo + ".xml.rels",
                        ChartWriter.DrawingRels(chartRels));
                }
                WriteXml(zip, sheetPart, sheetDoc);
                if (srels != null && srels.Root != null && srels.Root.HasElements)
                    WriteXml(zip, SheetRelsPath(sheetPart), srels);
            }
        }

        private static void AddDrawingRelationship(XDocument sheetDoc, ref XDocument srels,
                                                   string target, string rid)
        {
            if (sheetDoc == null || sheetDoc.Root == null) return;
            sheetDoc.Root.SetAttributeValue(XNamespace.Xmlns + "r", Rns.NamespaceName);
            XElement drawing = new XElement(Ns + "drawing", new XAttribute(Rns + "id", rid));
            XElement anchor = null;
            foreach (XElement el in sheetDoc.Root.Elements())
            {
                string n = el.Name.LocalName;
                if (string.Equals(n, "legacyDrawing", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(n, "legacyDrawingHF", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(n, "picture", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(n, "oleObjects", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(n, "controls", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(n, "tableParts", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(n, "extLst", StringComparison.OrdinalIgnoreCase))
                {
                    anchor = el; break;
                }
            }
            if (anchor != null) anchor.AddBeforeSelf(drawing); else sheetDoc.Root.Add(drawing);

            if (srels == null || srels.Root == null)
                srels = new XDocument(new XElement(PkgRel + "Relationships"));
            srels.Root.Add(new XElement(PkgRel + "Relationship",
                new XAttribute("Id", rid),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"),
                new XAttribute("Target", target)));
        }

        private static XElement BuildDataValidations(Sheet sheet)
        {
            if (sheet.Validations.Count == 0) return null;
            XElement box = new XElement(Ns + "dataValidations",
                new XAttribute("count", sheet.Validations.Count));
            foreach (DataValidationRule rule in sheet.Validations)
            {
                string sqref = AddressUtil.CellName(rule.Row1, rule.Col1);
                if (rule.Row2 != rule.Row1 || rule.Col2 != rule.Col1)
                    sqref += ":" + AddressUtil.CellName(rule.Row2, rule.Col2);
                string formula = null;
                if (!string.IsNullOrEmpty(rule.SourceRange))
                    formula = rule.SourceRange.TrimStart('=');
                else
                {
                    string list = string.Join(",", rule.Items.ToArray());
                    if (list.Length > 250) list = list.Substring(0, 250);
                    formula = "\"" + list.Replace("\"", "\"\"") + "\"";
                }
                XElement dv = new XElement(Ns + "dataValidation",
                    new XAttribute("type", "list"),
                    new XAttribute("allowBlank", rule.AllowBlank ? "1" : "0"),
                    new XAttribute("showErrorMessage", "1"),
                    new XAttribute("errorTitle", "Mini Excel"),
                    new XAttribute("error", rule.Message ?? "Valor no válido."),
                    new XAttribute("sqref", sqref),
                    new XElement(Ns + "formula1", formula));
                box.Add(dv);
            }
            return box;
        }

        private static XElement BuildAutoFilter(Sheet sheet)
        {
            if (string.IsNullOrEmpty(sheet.AutoFilterRef)) return null;
            XElement filter = new XElement(Ns + "autoFilter",
                new XAttribute("ref", sheet.AutoFilterRef));
            int r1, c1, r2, c2;
            if (AddressUtil.TryParseRange(sheet.AutoFilterRef, out r1, out c1, out r2, out c2))
                AddFilterSelections(filter, sheet, c1);
            return filter;
        }

        private static void ParseFilterSelections(Sheet sheet, XElement autoFilter, int baseCol)
        {
            if (sheet == null || autoFilter == null) return;
            foreach (XElement column in Descendants(autoFilter, "filterColumn"))
            {
                int colId;
                if (!int.TryParse(Attr(column, "colId"), out colId)) continue;
                FilterSelection selection = null;
                foreach (XElement value in Descendants(column, "filter"))
                {
                    string val = Attr(value, "val");
                    if (selection == null)
                    {
                        selection = new FilterSelection { Column = baseCol + colId };
                        sheet.FilterSelections.Add(selection);
                    }
                    selection.Values.Add(val ?? "");
                }
            }
        }

        private static void AddFilterSelections(XElement autoFilter, Sheet sheet, int baseCol)
        {
            if (autoFilter == null || sheet == null) return;
            XNamespace ns = autoFilter.Name.Namespace;
            foreach (FilterSelection selection in sheet.FilterSelections)
            {
                int colId = selection.Column - baseCol;
                if (colId < 0 || selection.Values.Count == 0) continue;
                XElement column = new XElement(ns + "filterColumn", new XAttribute("colId", colId));
                XElement filters = new XElement(ns + "filters");
                foreach (string value in selection.Values)
                    filters.Add(new XElement(ns + "filter", new XAttribute("val", value ?? "")));
                column.Add(filters);
                autoFilter.Add(column);
            }
        }

        private static XElement BuildSheetXml(Sheet sheet, StyleCollector styles)
        {
            int maxRow = Math.Max(0, sheet.UsedMaxRow);
            int maxCol = Math.Max(0, sheet.UsedMaxCol);
            XElement worksheet = new XElement(Ns + "worksheet",
                new XElement(Ns + "dimension", new XAttribute("ref", "A1:" + AddressUtil.CellName(maxRow, maxCol))),
                new XElement(Ns + "sheetViews",
                    new XElement(Ns + "sheetView", new XAttribute("workbookViewId", "0"))),
                new XElement(Ns + "sheetFormatPr", new XAttribute("defaultRowHeight", "15")));

            XElement cols = new XElement(Ns + "cols");
            for (int c = 0; c <= maxCol; c++)
            {
                float px = sheet.GetColWidth(c);
                double width = Math.Max(3.0, Math.Min(100.0, px / 7.0));
                cols.Add(new XElement(Ns + "col", new XAttribute("min", c + 1), new XAttribute("max", c + 1),
                    new XAttribute("width", width.ToString("0.##", CultureInfo.InvariantCulture)),
                    new XAttribute("customWidth", "1")));
            }
            worksheet.Add(cols);

            XElement sheetData = new XElement(Ns + "sheetData");
            List<long> keys = new List<long>(sheet.Cells.Keys);
            keys.Sort();
            int currentRow = -1;
            XElement rowEl = null;
            foreach (long key in keys)
            {
                int row = (int)(key >> 14);
                int col = (int)(key & 0x3FFF);
                if (row != currentRow)
                {
                    rowEl = new XElement(Ns + "row", new XAttribute("r", row + 1));
                    sheetData.Add(rowEl);
                    currentRow = row;
                }
                Cell cell = sheet.Cells[key];
                string reference = AddressUtil.CellName(row, col);
                XElement cEl = new XElement(Ns + "c", new XAttribute("r", reference));
                int styleIndex = styles.IndexFor(cell.EffectiveStyle);
                if (styleIndex > 0) cEl.Add(new XAttribute("s", styleIndex));

                if (cell.IsFormula)
                {
                    string f = FormulaExport.ToExcel(cell.Formula);
                    cEl.Add(new XElement(Ns + "f", f));
                    AddCachedValue(cEl, cell.Value);
                }
                else if (cell.Value.Kind == ValueKind.Text)
                {
                    cEl.Add(new XAttribute("t", "inlineStr"));
                    cEl.Add(new XElement(Ns + "is", new XElement(Ns + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"), cell.Value.Text ?? "")));
                }
                else if (cell.Value.Kind == ValueKind.Number)
                {
                    cEl.Add(new XElement(Ns + "v", cell.Value.Number.ToString("G17", CultureInfo.InvariantCulture)));
                }
                else if (cell.Value.Kind == ValueKind.Bool)
                {
                    cEl.Add(new XAttribute("t", "b"));
                    cEl.Add(new XElement(Ns + "v", cell.Value.Bool ? "1" : "0"));
                }
                else if (cell.Value.Kind == ValueKind.Error)
                {
                    cEl.Add(new XAttribute("t", "e"));
                    cEl.Add(new XElement(Ns + "v", ToEnglishError(cell.Value.Error)));
                }
                else if (cell.Style == null) continue;
                rowEl.Add(cEl);
            }
            worksheet.Add(sheetData);
            XElement validations = BuildDataValidations(sheet);
            if (validations != null) worksheet.Add(validations);
            return worksheet;
        }

        private static string XmlSafe(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            StringBuilder sb = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (XmlConvert.IsXmlChar(c)) { sb.Append(c); continue; }
                if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    sb.Append(c); sb.Append(text[++i]);
                }
                else sb.Append('\uFFFD');
            }
            return sb.ToString();
        }

        private static void AddCachedValue(XElement cEl, CellValue v)
        {
            if (v.Kind == ValueKind.Number)
                cEl.Add(new XElement(Ns + "v", v.Number.ToString("G17", CultureInfo.InvariantCulture)));
            else if (v.Kind == ValueKind.Text)
            {
                cEl.Add(new XAttribute("t", "str"));
                cEl.Add(new XElement(Ns + "v", XmlSafe(v.Text)));
            }
            else if (v.Kind == ValueKind.Bool)
            {
                cEl.Add(new XAttribute("t", "b"));
                cEl.Add(new XElement(Ns + "v", v.Bool ? "1" : "0"));
            }
            else if (v.Kind == ValueKind.Error)
            {
                cEl.Add(new XAttribute("t", "e"));
                cEl.Add(new XElement(Ns + "v", ToEnglishError(v.Error)));
            }
        }

        private static string ToEnglishError(string e)
        {
            if (e == Errors.Div0) return "#DIV/0!";
            if (e == Errors.Value) return "#VALUE!";
            if (e == Errors.Name) return "#NAME?";
            if (e == Errors.Ref) return "#REF!";
            if (e == Errors.NA) return "#N/A";
            if (e == Errors.Num) return "#NUM!";
            return "#VALUE!";
        }

        private interface IStyleIndexer
        {
            int IndexFor(Cell cell);
            void Commit();
        }

        private sealed class SourceStyleIndexer : IStyleIndexer
        {
            public int IndexFor(Cell cell)
            {
                return cell.SourceStyleIndex >= 0 ? cell.SourceStyleIndex : 0;
            }
            public void Commit() { }
        }

        private static TableFormat FindTableFormat(Workbook wb, string part)
        {
            if (wb == null) return null;
            foreach (Sheet sheet in wb.Sheets)
                foreach (TableFormat table in sheet.TableFormats)
                    if (string.Equals(table.SourcePart, part, StringComparison.OrdinalIgnoreCase))
                        return table;
            return null;
        }

        private static void UpdateTablePart(XDocument doc, TableFormat table)
        {
            if (doc == null || doc.Root == null || table == null) return;
            XElement root = doc.Root;
            string range = AddressUtil.RangeName(table.Row1, table.Col1, table.Row2, table.Col2);
            root.SetAttributeValue("ref", range);
            XElement filter = FirstDescendant(root, "autoFilter");
            if (filter != null)
            {
                filter.SetAttributeValue("ref", range);
                filter.Elements().Where(e => string.Equals(e.Name.LocalName, "filterColumn",
                    StringComparison.OrdinalIgnoreCase)).Remove();
                AddFilterSelections(filter, table.Sheet, table.Col1);
            }
            XElement columns = FirstDescendant(root, "tableColumns");
            if (columns != null)
            {
                int width = Math.Max(1, table.Col2 - table.Col1 + 1);
                List<XElement> list = new List<XElement>(Descendants(columns, "tableColumn"));
                while (list.Count > width)
                {
                    list[list.Count - 1].Remove();
                    list.RemoveAt(list.Count - 1);
                }
                int nextId = 1;
                foreach (XElement col in list)
                {
                    int id;
                    if (int.TryParse(Attr(col, "id"), out id) && id >= nextId) nextId = id + 1;
                }
                while (list.Count < width)
                {
                    XElement col = new XElement(columns.Name.Namespace + "tableColumn",
                        new XAttribute("id", nextId++),
                        new XAttribute("name", "Columna" + (list.Count + 1)));
                    columns.Add(col);
                    list.Add(col);
                }
                columns.SetAttributeValue("count", list.Count);
            }
        }

        private static void WriteXml(ZipArchive zip, string path, XDocument doc)
        {
            ZipArchiveEntry entry = zip.CreateEntry(path, CompressionLevel.Optimal);
            using (Stream s = entry.Open())
            using (XmlWriter w = XmlWriter.Create(s, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = false,
                CloseOutput = false
            }))
                doc.Save(w);
        }

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrEmpty(name)) name = "Hoja";
            StringBuilder sb = new StringBuilder();
            foreach (char c in name)
            {
                if (c == '[' || c == ']' || c == ':' || c == '*' || c == '?' || c == '/' || c == '\\')
                    sb.Append('_');
                else sb.Append(c);
            }
            string s = sb.ToString().Trim();
            if (s.Length == 0) s = "Hoja";
            if (s.Length > 31) s = s.Substring(0, 31);
            return s;
        }

        private static string MakeUniqueName(Workbook wb, string name)
        {
            if (string.IsNullOrEmpty(name)) name = "Hoja";
            if (name.Length > 31) name = name.Substring(0, 31);
            string baseName = name;
            int n = 2;
            while (wb.FindSheet(name) != null)
            {
                string suffix = " (" + n + ")";
                int max = 31 - suffix.Length;
                name = (baseName.Length > max ? baseName.Substring(0, max) : baseName) + suffix;
                n++;
            }
            return name;
        }

        private static string SheetRelsPath(string sheetPart)
        {
            return PartDirectory(sheetPart) + "_rels/" + Path.GetFileName(sheetPart) + ".rels";
        }

        private static string PartDirectory(string path)
        {
            int idx = path.LastIndexOf('/');
            return idx < 0 ? "" : path.Substring(0, idx + 1);
        }

        private static string ResolvePart(string baseDir, string target)
        {
            if (string.IsNullOrEmpty(target)) return target;
            target = target.Replace('\\', '/');
            string combined = target.StartsWith("/", StringComparison.Ordinal) ? target.TrimStart('/') : baseDir + target;
            string[] parts = combined.Split('/');
            List<string> stack = new List<string>();
            foreach (string p in parts)
            {
                if (p.Length == 0 || p == ".") continue;
                if (p == "..")
                {
                    if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                }
                else stack.Add(p);
            }
            return string.Join("/", stack.ToArray());
        }

        private static int ParseInt(string s)
        {
            int v;
            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out v) ? v : 0;
        }

        private static string Attr(XElement el, string name)
        {
            if (el == null) return null;
            XAttribute a = el.Attribute(name);
            return a == null ? null : a.Value;
        }

        private static string Attr(XElement el, XName name)
        {
            if (el == null) return null;
            XAttribute a = el.Attribute(name);
            return a == null ? null : a.Value;
        }

        private static string RelAttr(XElement el, string localName)
        {
            if (el == null) return null;
            foreach (XAttribute a in el.Attributes())
                if (string.Equals(a.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase))
                    return a.Value;
            return null;
        }

        private static XElement FirstDescendant(XElement root, string localName)
        {
            if (root == null) return null;
            foreach (XElement e in root.Descendants())
                if (string.Equals(e.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase)) return e;
            return null;
        }

        private static IEnumerable<XElement> Descendants(XElement root, string localName)
        {
            if (root == null) yield break;
            foreach (XElement e in root.Descendants())
                if (string.Equals(e.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase)) yield return e;
        }

        private sealed class StylesAppender : IStyleIndexer
        {
            private readonly XDocument _doc;
            private readonly XElement _fonts;
            private readonly XElement _fills;
            private readonly XElement _borders;
            private readonly XElement _cellXfs;
            private XElement _numFmts;
            private readonly Dictionary<string, int> _numFmtMap = new Dictionary<string, int>(StringComparer.Ordinal);
            private readonly Dictionary<string, int> _styleMap = new Dictionary<string, int>(StringComparer.Ordinal);
            private readonly Dictionary<Cell, int> _resolved = new Dictionary<Cell, int>();
            private int _nextNumFmt = 164;
            public bool Changed;

            public XDocument Document { get { return _doc; } }

            public StylesAppender(XDocument source)
            {
                if (source == null || source.Root == null)
                {
                    _doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                        new XElement(Ns + "styleSheet",
                            new XElement(Ns + "fonts", new XAttribute("count", "1"),
                                new XElement(Ns + "font",
                                    new XElement(Ns + "sz", new XAttribute("val", "10")),
                                    new XElement(Ns + "color", new XAttribute("rgb", "FF1F1F1F")),
                                    new XElement(Ns + "name", new XAttribute("val", "Segoe UI")),
                                    new XElement(Ns + "family", new XAttribute("val", "2")))),
                            new XElement(Ns + "fills", new XAttribute("count", "2"),
                                new XElement(Ns + "fill", new XElement(Ns + "patternFill", new XAttribute("patternType", "none"))),
                                new XElement(Ns + "fill", new XElement(Ns + "patternFill", new XAttribute("patternType", "gray125")))),
                            new XElement(Ns + "borders", new XAttribute("count", "1"),
                                new XElement(Ns + "border", new XElement(Ns + "left"), new XElement(Ns + "right"),
                                    new XElement(Ns + "top"), new XElement(Ns + "bottom"), new XElement(Ns + "diagonal"))),
                            new XElement(Ns + "cellStyleXfs", new XAttribute("count", "1"),
                                new XElement(Ns + "xf", new XAttribute("numFmtId", "0"), new XAttribute("fontId", "0"),
                                    new XAttribute("fillId", "0"), new XAttribute("borderId", "0"))),
                            new XElement(Ns + "cellXfs", new XAttribute("count", "1"),
                                new XElement(Ns + "xf", new XAttribute("numFmtId", "0"), new XAttribute("fontId", "0"),
                                    new XAttribute("fillId", "0"), new XAttribute("borderId", "0"), new XAttribute("xfId", "0"))),
                            new XElement(Ns + "cellStyles", new XAttribute("count", "1"),
                                new XElement(Ns + "cellStyle", new XAttribute("name", "Normal"),
                                    new XAttribute("xfId", "0"), new XAttribute("builtinId", "0")))));
                }
                else _doc = source;
                _fonts = Ensure(_doc.Root, "fonts");
                _fills = Ensure(_doc.Root, "fills");
                _borders = Ensure(_doc.Root, "borders");
                _cellXfs = Ensure(_doc.Root, "cellXfs");
                _numFmts = FirstDescendant(_doc.Root, "numFmts");
                if (_numFmts != null)
                {
                    foreach (XElement nf in Descendants(_numFmts, "numFmt"))
                    {
                        int id;
                        if (!int.TryParse(Attr(nf, "numFmtId"), out id)) continue;
                        string code = Attr(nf, "formatCode") ?? "";
                        _numFmtMap[code] = id;
                        if (id >= _nextNumFmt) _nextNumFmt = id + 1;
                    }
                }
            }

            private static XElement Ensure(XElement root, string name)
            {
                XElement el = ChildByLocalName(root, name);
                if (el == null)
                {
                    el = new XElement(root.Name.Namespace + name);
                    root.Add(el);
                }
                return el;
            }

            public int IndexFor(Cell cell)
            {
                if (!cell.StyleDirty && cell.SourceStyleIndex >= 0) return cell.SourceStyleIndex;
                CellStyle st = cell.EffectiveStyle;
                if (st == null || st.IsDefault())
                {
                    _resolved[cell] = 0;
                    return 0;
                }
                string key = KeyOf(st);
                int idx;
                if (_styleMap.TryGetValue(key, out idx)) { _resolved[cell] = idx; return idx; }
                idx = Append(st);
                _styleMap[key] = idx;
                _resolved[cell] = idx;
                return idx;
            }

            public void Commit()
            {
                foreach (KeyValuePair<Cell, int> kv in _resolved)
                {
                    kv.Key.SourceStyleIndex = kv.Value;
                    kv.Key.StyleDirty = false;
                }
            }

            private int Append(CellStyle st)
            {
                int fontId = _fonts.Elements().Count();
                _fonts.Add(BuildFont(st));
                _fonts.SetAttributeValue("count", _fonts.Elements().Count());

                int fillId = 0;
                if (!st.FillColor.IsEmpty)
                {
                    fillId = _fills.Elements().Count();
                    _fills.Add(new XElement(Ns + "fill", new XElement(Ns + "patternFill",
                        new XAttribute("patternType", "solid"),
                        new XElement(Ns + "fgColor", new XAttribute("rgb", Argbs(st.FillColor))),
                        new XElement(Ns + "bgColor", new XAttribute("indexed", "64")))));
                    _fills.SetAttributeValue("count", _fills.Elements().Count());
                }
                int borderId = 0;
                if (st.Borders != BorderFlags.None)
                {
                    borderId = _borders.Elements().Count();
                    _borders.Add(BuildBorder(st.Borders));
                    _borders.SetAttributeValue("count", _borders.Elements().Count());
                }
                int numFmtId = NumFmtId(st);
                int xfIndex = _cellXfs.Elements().Count();
                XElement xf = new XElement(Ns + "xf",
                    new XAttribute("numFmtId", numFmtId),
                    new XAttribute("fontId", fontId),
                    new XAttribute("fillId", fillId),
                    new XAttribute("borderId", borderId),
                    new XAttribute("xfId", "0"));
                if (numFmtId != 0) xf.Add(new XAttribute("applyNumberFormat", "1"));
                xf.Add(new XAttribute("applyFont", "1"));
                if (fillId != 0) xf.Add(new XAttribute("applyFill", "1"));
                if (borderId != 0) xf.Add(new XAttribute("applyBorder", "1"));
                if (st.Align >= 0 || st.WrapText)
                {
                    xf.Add(new XAttribute("applyAlignment", "1"));
                    XElement alignment = new XElement(Ns + "alignment",
                        new XAttribute("vertical", "center"));
                    if (st.Align >= 0)
                        alignment.SetAttributeValue("horizontal",
                            st.Align == 0 ? "left" : st.Align == 1 ? "center" : "right");
                    if (st.WrapText) alignment.SetAttributeValue("wrapText", "1");
                    xf.Add(alignment);
                }
                _cellXfs.Add(xf);
                _cellXfs.SetAttributeValue("count", _cellXfs.Elements().Count());
                Changed = true;
                return xfIndex;
            }

            private static XElement BuildFont(CellStyle s)
            {
                XElement font = new XElement(Ns + "font",
                    new XElement(Ns + "sz", new XAttribute("val", s.FontSize.ToString("0.##", CultureInfo.InvariantCulture))),
                    new XElement(Ns + "color", new XAttribute("rgb", Argbs(s.FontColor))),
                    new XElement(Ns + "name", new XAttribute("val", s.FontName ?? "Segoe UI")),
                    new XElement(Ns + "family", new XAttribute("val", "2")));
                if (s.Bold) font.Add(new XElement(Ns + "b"));
                if (s.Italic) font.Add(new XElement(Ns + "i"));
                if (s.Underline) font.Add(new XElement(Ns + "u"));
                return font;
            }

            private static XElement BuildBorder(BorderFlags b)
            {
                XElement border = new XElement(Ns + "border");
                border.Add(BorderSideAppend("left", (b & BorderFlags.Left) != 0));
                border.Add(BorderSideAppend("right", (b & BorderFlags.Right) != 0));
                border.Add(BorderSideAppend("top", (b & BorderFlags.Top) != 0));
                border.Add(BorderSideAppend("bottom", (b & BorderFlags.Bottom) != 0));
                border.Add(new XElement(Ns + "diagonal"));
                return border;
            }

            private static XElement BorderSideAppend(string name, bool on)
            {
                if (!on) return new XElement(Ns + name);
                return new XElement(Ns + name, new XAttribute("style", "thin"),
                    new XElement(Ns + "color", new XAttribute("rgb", "FF808080")));
            }

            private static string Argbs(Color c)
            {
                return "FF" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
            }

            private int NumFmtId(CellStyle st)
            {
                if (st.Format == NumberFormatKind.General) return 0;
                if (st.Format == NumberFormatKind.Text) return 49;
                string code = st.CustomFormat;
                if (string.IsNullOrEmpty(code))
                {
                    switch (st.Format)
                    {
                        case NumberFormatKind.Number:
                            code = "#,##0" + (st.Decimals > 0 ? "." + new string('0', st.Decimals) : "");
                            break;
                        case NumberFormatKind.Currency:
                            code = "\"$\"#,##0" + (st.Decimals > 0 ? "." + new string('0', st.Decimals) : "");
                            break;
                        case NumberFormatKind.Percent:
                            code = "0" + (st.Decimals > 0 ? "." + new string('0', st.Decimals) : "") + "%";
                            break;
                        case NumberFormatKind.Date:
                            code = "dd/mm/yyyy";
                            break;
                        case NumberFormatKind.Time:
                            code = "hh:mm:ss";
                            break;
                        default:
                            return 0;
                    }
                }
                int id;
                if (_numFmtMap.TryGetValue(code, out id)) return id;
                id = _nextNumFmt++;
                if (_numFmts == null)
                {
                    _numFmts = new XElement(Ns + "numFmts");
                    _doc.Root.AddFirst(_numFmts);
                }
                _numFmts.Add(new XElement(Ns + "numFmt", new XAttribute("numFmtId", id), new XAttribute("formatCode", code)));
                _numFmts.SetAttributeValue("count", _numFmts.Elements().Count());
                _numFmtMap[code] = id;
                Changed = true;
                return id;
            }

            private static string KeyOf(CellStyle s)
            {
                return (s.Bold ? "1" : "0") + (s.Italic ? "1" : "0") + (s.Underline ? "1" : "0") + "|" +
                       s.Align + "|" + (int)s.Format + "|" + s.Decimals + "|" +
                       s.FontColor.ToArgb() + "|" + (s.FillColor.IsEmpty ? 0 : s.FillColor.ToArgb()) + "|" +
                       (int)s.Borders + "|" + s.FontSize.ToString("0.##", CultureInfo.InvariantCulture) + "|" + s.FontName +
                       "|" + (s.CustomFormat ?? "");
            }
        }

        private sealed class StyleCollector : IStyleIndexer
        {
            private readonly Dictionary<string, int> _map = new Dictionary<string, int>(StringComparer.Ordinal);
            private readonly List<CellStyle> _styles = new List<CellStyle>();
            private readonly List<string> _keys = new List<string>();
            private readonly List<int> _numFmtIds = new List<int>();
            private readonly Dictionary<string, int> _numFmtMap = new Dictionary<string, int>(StringComparer.Ordinal);
            private int _nextNumFmt = 164;

            public StyleCollector(Workbook wb)
            {
                _map[""] = 0;
                _styles.Add(new CellStyle());
                _keys.Add("");
                _numFmtIds.Add(0);
                foreach (Sheet s in wb.Sheets)
                {
                    foreach (Cell c in s.Cells.Values)
                    {
                        CellStyle st = c.EffectiveStyle;
                        if (st.IsDefault()) continue;
                        string key = KeyOf(st);
                        if (_map.ContainsKey(key)) continue;
                        int idx = _styles.Count;
                        _map[key] = idx;
                        _styles.Add(st.Clone());
                        _keys.Add(key);
                        _numFmtIds.Add(NumFmtId(st));
                    }
                }
            }

            public int IndexFor(CellStyle st)
            {
                if (st == null || st.IsDefault()) return 0;
                string key = KeyOf(st);
                int idx;
                return _map.TryGetValue(key, out idx) ? idx : 0;
            }

            public int IndexFor(Cell cell)
            {
                return IndexFor(cell.EffectiveStyle);
            }

            public void Commit() { }

            private static string KeyOf(CellStyle s)
            {
                return (s.Bold ? "1" : "0") + (s.Italic ? "1" : "0") + (s.Underline ? "1" : "0") + "|" +
                       s.Align + "|" + (int)s.Format + "|" + s.Decimals + "|" +
                       s.FontColor.ToArgb() + "|" + (s.FillColor.IsEmpty ? 0 : s.FillColor.ToArgb()) + "|" +
                       (int)s.Borders + "|" + s.FontSize.ToString("0.##", CultureInfo.InvariantCulture) + "|" + s.FontName +
                       "|" + (s.CustomFormat ?? "");
            }

            private int NumFmtId(CellStyle st)
            {
                if (st.Format == NumberFormatKind.General) return 0;
                if (st.Format == NumberFormatKind.Text) return 49;
                string code = null;
                if (!string.IsNullOrEmpty(st.CustomFormat)) code = st.CustomFormat;
                else switch (st.Format)
                {
                    case NumberFormatKind.Number:
                        code = "#,##0" + (st.Decimals > 0 ? "." + new string('0', st.Decimals) : "");
                        break;
                    case NumberFormatKind.Currency:
                        code = "\"$\"#,##0" + (st.Decimals > 0 ? "." + new string('0', st.Decimals) : "");
                        break;
                    case NumberFormatKind.Percent:
                        code = "0" + (st.Decimals > 0 ? "." + new string('0', st.Decimals) : "") + "%";
                        break;
                    case NumberFormatKind.Date:
                        code = "dd/mm/yyyy";
                        break;
                    case NumberFormatKind.Time:
                        code = "hh:mm:ss";
                        break;
                    case NumberFormatKind.Custom:
                        code = st.CustomFormat;
                        break;
                    default:
                        return 0;
                }
                if (string.IsNullOrEmpty(code)) return 0;
                int id;
                if (_numFmtMap.TryGetValue(code, out id)) return id;
                id = _nextNumFmt++;
                _numFmtMap[code] = id;
                return id;
            }

            public XElement BuildStylesXml()
            {
                XElement root = new XElement(Ns + "styleSheet");

                if (_numFmtMap.Count > 0)
                {
                    XElement numFmts = new XElement(Ns + "numFmts", new XAttribute("count", _numFmtMap.Count));
                    foreach (KeyValuePair<string, int> kv in _numFmtMap)
                        numFmts.Add(new XElement(Ns + "numFmt",
                            new XAttribute("numFmtId", kv.Value), new XAttribute("formatCode", kv.Key)));
                    root.Add(numFmts);
                }

                XElement fonts = new XElement(Ns + "fonts", new XAttribute("count", _styles.Count));
                fonts.Add(new XElement(Ns + "font",
                    new XElement(Ns + "sz", new XAttribute("val", "10")),
                    new XElement(Ns + "color", new XAttribute("rgb", "FF1F1F1F")),
                    new XElement(Ns + "name", new XAttribute("val", "Segoe UI")),
                    new XElement(Ns + "family", new XAttribute("val", "2"))));
                for (int i = 1; i < _styles.Count; i++)
                {
                    CellStyle s = _styles[i];
                    XElement font = new XElement(Ns + "font",
                        new XElement(Ns + "sz", new XAttribute("val", s.FontSize.ToString("0.##", CultureInfo.InvariantCulture))),
                        new XElement(Ns + "color", new XAttribute("rgb", Argb(s.FontColor))),
                        new XElement(Ns + "name", new XAttribute("val", s.FontName ?? "Segoe UI")),
                        new XElement(Ns + "family", new XAttribute("val", "2")));
                    if (s.Bold) font.Add(new XElement(Ns + "b"));
                    if (s.Italic) font.Add(new XElement(Ns + "i"));
                    if (s.Underline) font.Add(new XElement(Ns + "u"));
                    fonts.Add(font);
                }
                root.Add(fonts);

                int fillCount = 2;
                for (int i = 1; i < _styles.Count; i++) if (!_styles[i].FillColor.IsEmpty) fillCount++;
                XElement fills = new XElement(Ns + "fills", new XAttribute("count", fillCount));
                fills.Add(new XElement(Ns + "fill", new XElement(Ns + "patternFill", new XAttribute("patternType", "none"))));
                fills.Add(new XElement(Ns + "fill", new XElement(Ns + "patternFill", new XAttribute("patternType", "gray125"))));
                List<int> fillIndex = new List<int>();
                fillIndex.Add(0);
                int nextFill = 2;
                for (int i = 1; i < _styles.Count; i++)
                {
                    CellStyle s = _styles[i];
                    if (s.FillColor.IsEmpty) { fillIndex.Add(0); continue; }
                    fills.Add(new XElement(Ns + "fill", new XElement(Ns + "patternFill",
                        new XAttribute("patternType", "solid"),
                        new XElement(Ns + "fgColor", new XAttribute("rgb", Argb(s.FillColor))),
                        new XElement(Ns + "bgColor", new XAttribute("indexed", "64")))));
                    fillIndex.Add(nextFill++);
                }
                root.Add(fills);

                int borderCount = 1;
                for (int i = 1; i < _styles.Count; i++) if (_styles[i].Borders != BorderFlags.None) borderCount++;
                XElement borders = new XElement(Ns + "borders", new XAttribute("count", borderCount));
                borders.Add(new XElement(Ns + "border",
                    new XElement(Ns + "left"), new XElement(Ns + "right"),
                    new XElement(Ns + "top"), new XElement(Ns + "bottom"), new XElement(Ns + "diagonal")));
                List<int> borderIndex = new List<int>();
                borderIndex.Add(0);
                int nextBorder = 1;
                for (int i = 1; i < _styles.Count; i++)
                {
                    BorderFlags b = _styles[i].Borders;
                    if (b == BorderFlags.None) { borderIndex.Add(0); continue; }
                    XElement border = new XElement(Ns + "border");
                    border.Add(BorderSide("left", (b & BorderFlags.Left) != 0));
                    border.Add(BorderSide("right", (b & BorderFlags.Right) != 0));
                    border.Add(BorderSide("top", (b & BorderFlags.Top) != 0));
                    border.Add(BorderSide("bottom", (b & BorderFlags.Bottom) != 0));
                    border.Add(new XElement(Ns + "diagonal"));
                    borders.Add(border);
                    borderIndex.Add(nextBorder++);
                }
                root.Add(borders);

                root.Add(new XElement(Ns + "cellStyleXfs", new XAttribute("count", "1"),
                    new XElement(Ns + "xf", new XAttribute("numFmtId", "0"), new XAttribute("fontId", "0"),
                        new XAttribute("fillId", "0"), new XAttribute("borderId", "0"))));

                XElement cellXfs = new XElement(Ns + "cellXfs", new XAttribute("count", _styles.Count));
                for (int i = 0; i < _styles.Count; i++)
                {
                    CellStyle s = _styles[i];
                    int fontId = i;
                    int fillId = i < fillIndex.Count ? fillIndex[i] : 0;
                    int borderId = i < borderIndex.Count ? borderIndex[i] : 0;
                    int numId = i < _numFmtIds.Count ? _numFmtIds[i] : 0;
                    XElement xf = new XElement(Ns + "xf",
                        new XAttribute("numFmtId", numId),
                        new XAttribute("fontId", fontId),
                        new XAttribute("fillId", fillId),
                        new XAttribute("borderId", borderId),
                        new XAttribute("xfId", "0"));
                    if (numId != 0) xf.Add(new XAttribute("applyNumberFormat", "1"));
                    if (i > 0) xf.Add(new XAttribute("applyFont", "1"));
                    if (fillId != 0) xf.Add(new XAttribute("applyFill", "1"));
                    if (borderId != 0) xf.Add(new XAttribute("applyBorder", "1"));
                    if (i > 0 && (s.Align >= 0 || s.WrapText))
                    {
                        xf.Add(new XAttribute("applyAlignment", "1"));
                        XElement alignment = new XElement(Ns + "alignment",
                            new XAttribute("vertical", "center"));
                        if (s.Align >= 0)
                            alignment.SetAttributeValue("horizontal",
                                s.Align == 0 ? "left" : s.Align == 1 ? "center" : "right");
                        if (s.WrapText) alignment.SetAttributeValue("wrapText", "1");
                        xf.Add(alignment);
                    }
                    cellXfs.Add(xf);
                }
                root.Add(cellXfs);
                root.Add(new XElement(Ns + "cellStyles", new XAttribute("count", "1"),
                    new XElement(Ns + "cellStyle", new XAttribute("name", "Normal"),
                        new XAttribute("xfId", "0"), new XAttribute("builtinId", "0"))));
                return root;
            }

            private static XElement BorderSide(string name, bool on)
            {
                if (!on) return new XElement(Ns + name);
                return new XElement(Ns + name, new XAttribute("style", "thin"),
                    new XElement(Ns + "color", new XAttribute("rgb", "FF808080")));
            }

            private static string Argb(Color c)
            {
                return "FF" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
            }
        }
    }

    internal static class FormulaExport
    {
        public static string ToExcel(string formula)
        {
            if (string.IsNullOrEmpty(formula)) return "";
            string f = formula;
            if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1);
            try
            {
                FormulaLexer lexer = new FormulaLexer(f);
                List<Token> tokens = lexer.Tokenize();
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < tokens.Count; i++)
                {
                    Token t = tokens[i];
                    if (t.Kind == TokenKind.End) break;
                    if (t.Kind == TokenKind.Ident)
                    {
                        bool isFunction = i + 1 < tokens.Count && tokens[i + 1].Kind == TokenKind.LParen;
                        if (isFunction)
                        {
                            string canonical = Builtins.Canonical(t.Text);
                            sb.Append(canonical ?? t.Text);
                        }
                        else
                        {
                            string u = t.Text.ToUpperInvariant();
                            if (u == "VERDADERO") sb.Append("TRUE");
                            else if (u == "FALSO") sb.Append("FALSE");
                            else sb.Append(t.Text);
                        }
                    }
                    else if (t.Kind == TokenKind.Comma) sb.Append(',');
                    else if (t.Kind == TokenKind.Text) sb.Append('"').Append((t.Text ?? "").Replace("\"", "\"\"")).Append('"');
                    else if (t.Kind == TokenKind.QuotedSheet) sb.Append('\'').Append((t.Text ?? "").Replace("'", "''")).Append('\'');
                    else if (t.Kind == TokenKind.ErrorValue) sb.Append(ToEnglishError(t.Text));
                    else if (t.Kind == TokenKind.Number) sb.Append(t.Text);
                    else sb.Append(t.Text);
                }
                return sb.ToString();
            }
            catch
            {
                return f;
            }
        }

        private static string ToEnglishError(string e)
        {
            if (e == Errors.Div0) return "#DIV/0!";
            if (e == Errors.Value) return "#VALUE!";
            if (e == Errors.Name) return "#NAME?";
            if (e == Errors.Ref) return "#REF!";
            if (e == Errors.NA) return "#N/A";
            if (e == Errors.Num) return "#NUM!";
            return "#VALUE!";
        }
    }
}

