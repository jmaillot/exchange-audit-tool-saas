using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    internal enum AuditScope { ExchangeOnline, OnPremises, Both }

    internal enum GroupMode { MultiCheck, SingleChoice }

    internal sealed class AuditOption
    {
        public string Key;
        public string Label;
        public string Value;
        public bool DefaultChecked;
        public bool DefOnline;
        public bool DefOnPrem;
        public bool Slow;

        public AuditOption(string key, string label, string value, bool defaultChecked)
        {
            Key = key;
            Label = label;
            Value = value;
            DefaultChecked = defaultChecked;
            DefOnline = defaultChecked;
            DefOnPrem = defaultChecked;
        }

        public static AuditOption Prop(string name, bool defaultChecked)
        {
            return new AuditOption(name, name, name, defaultChecked);
        }

        public static AuditOption PropScoped(string name, bool defOnline, bool defOnPrem)
        {
            var o = new AuditOption(name, name, name, defOnline || defOnPrem);
            o.DefOnline = defOnline;
            o.DefOnPrem = defOnPrem;
            return o;
        }

        public AuditOption MarkSlow() { Slow = true; return this; }
    }

    internal sealed class AuditOptionGroup
    {
        public string Key;
        public string Title;
        public string Hint;
        public GroupMode Mode;
        public int Columns;
        public bool Slow;
        public List<AuditOption> Options = new List<AuditOption>();

        public AuditOptionGroup(string key, string title, GroupMode mode)
        {
            Key = key;
            Title = title;
            Mode = mode;
            Columns = 2;
        }

        public AuditOptionGroup Add(AuditOption option) { Options.Add(option); return this; }
        public AuditOptionGroup AddProp(string name, bool def) { Options.Add(AuditOption.Prop(name, def)); return this; }
        public AuditOptionGroup MarkSlow() { Slow = true; return this; }
    }

    internal sealed class AuditSelection
    {
        private readonly Dictionary<string, List<string>> _values = new Dictionary<string, List<string>>();

        public void Set(string groupKey, List<string> values) { _values[groupKey] = values; }

        public List<string> Selected(string groupKey)
        {
            List<string> v;
            return _values.TryGetValue(groupKey, out v) ? v : new List<string>();
        }

        public bool IsSelected(string groupKey, string value) { return Selected(groupKey).Contains(value); }

        public string First(string groupKey, string fallback)
        {
            var v = Selected(groupKey);
            return v.Count > 0 ? v[0] : fallback;
        }

        public string Join(string groupKey, string separator) { return string.Join(separator, Selected(groupKey).ToArray()); }

        public bool Any(string groupKey) { return Selected(groupKey).Count > 0; }
    }

    internal delegate string AuditScriptBuilder(AuditSelection selection, ScriptContext ctx);

    internal sealed class ScriptContext
    {
        public string CsvPath;

        public ScriptContext(string csvPath) { CsvPath = csvPath; }

        public string ExportCsv(string pipelineExpression)
        {
            var sb = new StringBuilder();
            sb.AppendLine(pipelineExpression + " |");
            // CSV delimiter ';' (multi-value cells are joined with ',' so they never clash).
            sb.AppendLine("    Export-Csv -LiteralPath " + PsLiteral(CsvPath) + " -NoTypeInformation -Encoding UTF8 -Delimiter ';'");
            return sb.ToString();
        }

        public static string PsLiteral(string value)
        {
            if (value == null) value = "";
            return "'" + value.Replace("'", "''") + "'";
        }
    }

    internal sealed class AuditSection
    {
        public string Id;
        public string NavTitle;
        public string Title;
        public string Subtitle;
        public string IconKey;
        public AuditScope Scope;
        public string Category;
        public List<AuditOptionGroup> Groups = new List<AuditOptionGroup>();
        public AuditScriptBuilder BuildScript;
        public string DefaultFileName = "audit.csv";
        public bool ScopeAwareDefaults = false;

        public AuditSection(string id, string navTitle, string title, string subtitle, string iconKey, AuditScope scope)
        {
            Id = id;
            NavTitle = navTitle;
            Title = title;
            Subtitle = subtitle;
            IconKey = iconKey;
            Scope = scope;
            Category = "";
        }

        public AuditOptionGroup AddGroup(AuditOptionGroup group) { Groups.Add(group); return group; }
    }

    internal static class AuditRegistry
    {
        public static readonly List<AuditSection> Sections = new List<AuditSection>();

        public static void Register(AuditSection section) { Sections.Add(section); }

        public static void BuildAll()
        {
            Sections.Clear();
            SectionsMailboxes.RegisterMailboxSections();
            SectionsMailboxTypes.Register();
            SectionsGroups.Register();
            SectionsContacts.Register();
            SectionsPublicFolders.Register();
            SectionsDomainsRouting.Register();
            SectionsOrganization.Register();
            SectionsProtection.Register();
        }
    }
}
