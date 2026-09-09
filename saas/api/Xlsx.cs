using System.IO.Packaging;
using System.Text;
using System.Xml;

namespace ExchangeAuditSaaS;

// Dependency-free ;-delimited CSV -> single-sheet XLSX.
// Bold header, frozen top row, autofilter, auto-width 12-60,
// 1048575 x 16384 limits.
internal static class Xlsx
{
    private const int MaxDataRows = 1048575;
    private const int MaxColumns = 16384;
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static void FromCsv(string csvPath, string xlsxPath, string sheetName)
    {
        sheetName = Sanitize(sheetName);
        string[] headers;
        using (var r = new StreamReader(csvPath, Encoding.UTF8, true))
        {
            string hl = r.ReadLine() ?? throw new InvalidOperationException("CSV is empty");
            headers = Parse(hl).ToArray();
        }
        int cols = Math.Min(headers.Length, MaxColumns);
        if (cols == 0) throw new InvalidOperationException("CSV header has no columns");
        int[] widths = new int[cols];
        for (int c = 0; c < cols; c++) widths[c] = headers[c]?.Length ?? 0;

        int totalRows = 0;
        using (var counter = new StreamReader(csvPath, Encoding.UTF8, true))
        {
            counter.ReadLine();
            string line;
            while ((line = counter.ReadLine()) != null)
            {
                if (totalRows >= MaxDataRows) break;
                totalRows++;
                var cells = Parse(line);
                for (int c = 0; c < cols && c < cells.Count; c++)
                    if (cells[c]?.Length > widths[c]) widths[c] = cells[c].Length;
            }
        }

        using (Package package = Package.Open(xlsxPath, FileMode.Create))
        {
            WritePart(package, "/docProps/core.xml",
                "application/vnd.openxmlformats-package.core-properties+xml", w =>
                {
                    w.WriteStartElement("coreProperties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
                    w.WriteElementString("title", "http://purl.org/dc/elements/1.1/", "Exchange Audit Export");
                    w.WriteEndElement();
                });
            WritePart(package, "/docProps/app.xml",
                "application/vnd.openxmlformats-officedocument.extended-properties+xml", w =>
                {
                    w.WriteStartElement("Properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
                    w.WriteElementString("Application", "Exchange Audit SaaS");
                    w.WriteEndElement();
                });
            WritePart(package, "/xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", w =>
                {
                    w.WriteStartElement("workbook", MainNs);
                    w.WriteAttributeString("xmlns", "r", null, RelNs);
                    w.WriteStartElement("sheets");
                    w.WriteStartElement("sheet");
                    w.WriteAttributeString("name", sheetName);
                    w.WriteAttributeString("sheetId", "1");
                    w.WriteAttributeString("r", "id", RelNs, "rId1");
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                });
            WritePart(package, "/xl/styles.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", w =>
                {
                    w.WriteStartElement("styleSheet", MainNs);
                    w.WriteStartElement("fonts"); w.WriteAttributeString("count", "2");
                    w.WriteStartElement("font");
                    w.WriteStartElement("sz"); w.WriteAttributeString("val", "11"); w.WriteEndElement();
                    w.WriteStartElement("name"); w.WriteAttributeString("val", "Calibri"); w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteStartElement("font");
                    w.WriteStartElement("b"); w.WriteEndElement();
                    w.WriteStartElement("sz"); w.WriteAttributeString("val", "11"); w.WriteEndElement();
                    w.WriteStartElement("name"); w.WriteAttributeString("val", "Calibri"); w.WriteEndElement();
                    w.WriteEndElement(); w.WriteEndElement();
                    w.WriteStartElement("fills"); w.WriteAttributeString("count", "1");
                    w.WriteStartElement("fill");
                    w.WriteStartElement("patternFill"); w.WriteAttributeString("patternType", "none"); w.WriteEndElement();
                    w.WriteEndElement(); w.WriteEndElement();
                    w.WriteStartElement("borders"); w.WriteAttributeString("count", "1");
                    w.WriteStartElement("border");
                    w.WriteStartElement("left"); w.WriteEndElement();
                    w.WriteStartElement("right"); w.WriteEndElement();
                    w.WriteStartElement("top"); w.WriteEndElement();
                    w.WriteStartElement("bottom"); w.WriteEndElement();
                    w.WriteStartElement("diagonal"); w.WriteEndElement();
                    w.WriteEndElement(); w.WriteEndElement();
                    w.WriteStartElement("cellStyleXfs"); w.WriteAttributeString("count", "1");
                    w.WriteStartElement("xf");
                    w.WriteAttributeString("numFmtId", "0"); w.WriteAttributeString("fontId", "0");
                    w.WriteAttributeString("fillId", "0"); w.WriteAttributeString("borderId", "0");
                    w.WriteEndElement(); w.WriteEndElement();
                    w.WriteStartElement("cellXfs"); w.WriteAttributeString("count", "2");
                    w.WriteStartElement("xf");
                    w.WriteAttributeString("numFmtId", "0"); w.WriteAttributeString("fontId", "0");
                    w.WriteAttributeString("fillId", "0"); w.WriteAttributeString("borderId", "0");
                    w.WriteAttributeString("xfId", "0"); w.WriteEndElement();
                    w.WriteStartElement("xf");
                    w.WriteAttributeString("numFmtId", "0"); w.WriteAttributeString("fontId", "1");
                    w.WriteAttributeString("fillId", "0"); w.WriteAttributeString("borderId", "0");
                    w.WriteAttributeString("xfId", "0"); w.WriteAttributeString("applyFont", "1");
                    w.WriteEndElement(); w.WriteEndElement(); w.WriteEndElement();
                });

            string lastCol = ColName(cols - 1);
            PackagePart sheet = package.CreatePart(new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Normal);
            using (Stream s = sheet.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter w = XmlWriter.Create(s, new XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
            {
                w.WriteStartDocument();
                w.WriteStartElement("worksheet", MainNs);
                w.WriteStartElement("dimension"); w.WriteAttributeString("ref", "A1:" + lastCol + (totalRows + 1)); w.WriteEndElement();
                w.WriteStartElement("sheetViews"); w.WriteStartElement("sheetView"); w.WriteAttributeString("workbookViewId", "0");
                w.WriteStartElement("pane"); w.WriteAttributeString("ySplit", "1");
                w.WriteAttributeString("topLeftCell", "A2"); w.WriteAttributeString("activePane", "bottomLeft");
                w.WriteAttributeString("state", "frozen"); w.WriteEndElement();
                w.WriteEndElement(); w.WriteEndElement();
                w.WriteStartElement("cols");
                for (int c = 0; c < cols; c++)
                {
                    int width = Math.Min(60, Math.Max(12, widths[c] + 2));
                    w.WriteStartElement("col");
                    w.WriteAttributeString("min", (c + 1).ToString());
                    w.WriteAttributeString("max", (c + 1).ToString());
                    w.WriteAttributeString("width", width.ToString());
                    w.WriteAttributeString("customWidth", "1");
                    w.WriteEndElement();
                }
                w.WriteEndElement();
                w.WriteStartElement("sheetData");
                w.WriteStartElement("row"); w.WriteAttributeString("r", "1");
                for (int c = 0; c < cols; c++) Cell(w, c, 1, c < headers.Length ? headers[c] : "", true);
                w.WriteEndElement();
                using (var reader = new StreamReader(csvPath, Encoding.UTF8, true))
                {
                    reader.ReadLine();
                    int row = 1; string line;
                    while (row <= totalRows && (line = reader.ReadLine()) != null)
                    {
                        row++;
                        var cells = Parse(line);
                        w.WriteStartElement("row"); w.WriteAttributeString("r", row.ToString());
                        for (int c = 0; c < cols; c++) Cell(w, c, row, c < cells.Count ? cells[c] : "", false);
                        w.WriteEndElement();
                    }
                }
                w.WriteEndElement();
                w.WriteStartElement("autoFilter"); w.WriteAttributeString("ref", "A1:" + lastCol + "1"); w.WriteEndElement();
                w.WriteEndElement();
                w.WriteEndDocument();
            }

            var wb = package.GetPart(new Uri("/xl/workbook.xml", UriKind.Relative));
            wb.CreateRelationship(new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative),
                TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "rId1");
            wb.CreateRelationship(new Uri("/xl/styles.xml", UriKind.Relative),
                TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "rId2");
            package.CreateRelationship(new Uri("/xl/workbook.xml", UriKind.Relative),
                TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId1");
        }
    }

    private static void Cell(XmlWriter w, int col, int row, string v, bool bold)
    {
        w.WriteStartElement("c");
        w.WriteAttributeString("r", ColName(col) + row);
        v ??= "";
        if (v.Length == 0) { w.WriteEndElement(); return; }
        w.WriteAttributeString("t", "inlineStr");
        if (bold) w.WriteAttributeString("s", "1");
        w.WriteStartElement("is"); w.WriteStartElement("t"); w.WriteString(v);
        w.WriteEndElement(); w.WriteEndElement(); w.WriteEndElement();
    }

    private static void WritePart(Package p, string path, string ct, Action<XmlWriter> body)
    {
        var part = p.CreatePart(new Uri(path, UriKind.Relative), ct, CompressionOption.Normal);
        using var s = part.GetStream(FileMode.Create, FileAccess.Write);
        using var w = XmlWriter.Create(s, new XmlWriterSettings { Encoding = new UTF8Encoding(false) });
        w.WriteStartDocument(); body(w); w.WriteEndDocument();
    }

    private static string ColName(int i)
    {
        string s = ""; int n = i + 1;
        while (n > 0) { int m = (n - 1) % 26; s = (char)('A' + m) + s; n = (n - 1) / 26; }
        return s;
    }

    private static string Sanitize(string n)
    {
        if (string.IsNullOrEmpty(n)) return "Audit";
        foreach (char b in new[] { '\\', '/', '*', '?', ':', '[', ']' }) n = n.Replace(b, '-');
        n = n.Trim();
        if (n.Length == 0) n = "Audit";
        return n.Length > 31 ? n.Substring(0, 31) : n;
    }

    private static List<string> Parse(string line)
    {
        var cells = new List<string>();
        var cur = new StringBuilder();
        bool inQ = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (inQ)
            {
                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"') { cur.Append('"'); i++; }
                    else inQ = false;
                }
                else cur.Append(c);
            }
            else
            {
                if (c == '"') inQ = true;
                else if (c == ';') { cells.Add(cur.ToString()); cur.Clear(); }
                else cur.Append(c);
            }
        }
        cells.Add(cur.ToString());
        return cells;
    }
}
