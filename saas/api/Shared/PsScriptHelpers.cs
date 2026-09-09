using System.Collections.Generic;
using System.Text;

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("ExchangeAuditTool.Tests")]

namespace ExchangeAuditTool
{
    internal static class PsScriptHelpers
    {
        public static void EmitResolver(StringBuilder sb, string getRecipCmd, bool displayNameFallback)
        {
            sb.AppendLine("$script:recipCache = @{}");
            sb.AppendLine("function Resolve-Recip {");
            sb.AppendLine("    param($values)");
            sb.AppendLine("    $out = New-Object System.Collections.Generic.List[string]");
            sb.AppendLine("    foreach ($v in @($values)) {");
            sb.AppendLine("        $key = [string]$v");
            sb.AppendLine("        if ([string]::IsNullOrEmpty($key)) { continue }");
            sb.AppendLine("        if ($script:recipCache.ContainsKey($key)) { $smtp = $script:recipCache[$key] }");
            sb.AppendLine("        else {");
            sb.AppendLine("            $smtp = $key");
            if (displayNameFallback)
                sb.AppendLine("            try { $r = " + getRecipCmd + " -Identity $key -ErrorAction Stop; if ($r) { if ($r.DisplayName) { $smtp = [string]$r.DisplayName } elseif ($r.PrimarySmtpAddress) { $smtp = [string]$r.PrimarySmtpAddress } } } catch { }");
            else
                sb.AppendLine("            try { $r = " + getRecipCmd + " -Identity $key -ErrorAction Stop; if ($r -and $r.PrimarySmtpAddress) { $smtp = [string]$r.PrimarySmtpAddress } } catch { }");
            sb.AppendLine("            $script:recipCache[$key] = $smtp");
            sb.AppendLine("        }");
            sb.AppendLine("        if (-not $out.Contains($smtp)) { $out.Add($smtp) }");
            sb.AppendLine("    }");
            sb.AppendLine("    ($out -join ',')");
            sb.AppendLine("}");
            sb.AppendLine();
        }

        public static void EmitSizeHelper(StringBuilder sb, string getStatsCmd)
        {
            sb.AppendLine("function Get-SizeMB {");
            sb.AppendLine("    param($identity)");
            sb.AppendLine("    try {");
            sb.AppendLine("        $st = " + getStatsCmd + " -Identity $identity -ErrorAction Stop");
            sb.AppendLine("        if ($st -and $st.TotalItemSize) {");
            sb.AppendLine("            $s = $st.TotalItemSize.ToString()");
            sb.AppendLine("            if ($s -match '\\(([\\d,]+) bytes\\)') { return [math]::Round(([double]($matches[1] -replace ',','')) / 1MB, 2) }");
            sb.AppendLine("        }");
            sb.AppendLine("    } catch { }");
            sb.AppendLine("    return ''");
            sb.AppendLine("}");
            sb.AppendLine();
        }

        public static void EmitCountHelper(StringBuilder sb, string getStatsCmd)
        {
            sb.AppendLine("function Get-ItemCount {");
            sb.AppendLine("    param($identity)");
            sb.AppendLine("    try {");
            sb.AppendLine("        $st = " + getStatsCmd + " -Identity $identity -ErrorAction Stop");
            sb.AppendLine("        if ($st -and $null -ne $st.ItemCount) { return $st.ItemCount }");
            sb.AppendLine("    } catch { }");
            sb.AppendLine("    return ''");
            sb.AppendLine("}");
            sb.AppendLine();
        }

        public static void EmitArchiveMBHelper(StringBuilder sb, string getStatsCmd)
        {
            sb.AppendLine("function Get-ArchiveMB {");
            sb.AppendLine("    param($identity)");
            sb.AppendLine("    try {");
            sb.AppendLine("        $st = " + getStatsCmd + " -Identity $identity -Archive -ErrorAction Stop");
            sb.AppendLine("        if ($st -and $st.TotalItemSize) {");
            sb.AppendLine("            $s = $st.TotalItemSize.ToString()");
            sb.AppendLine("            if ($s -match '\\(([\\d,]+) bytes\\)') { return [math]::Round(([double]($matches[1] -replace ',','')) / 1MB, 2) }");
            sb.AppendLine("        }");
            sb.AppendLine("    } catch { }");
            sb.AppendLine("    return ''");
            sb.AppendLine("}");
            sb.AppendLine();
        }

        public static void AddSizeGroup(AuditSection section)
        {
            var size = new AuditOptionGroup("size", "Result size", GroupMode.SingleChoice); size.Columns = 3;
            size.Add(new AuditOption("unlimited", "Unlimited", "Unlimited", true));
            size.Add(new AuditOption("1000", "First 1000", "1000", false));
            size.Add(new AuditOption("100", "First 100", "100", false));
            section.AddGroup(size);
        }

        public static List<string> Collect(AuditSelection sel, params string[] groupKeys)
        {
            var chosen = new List<string>();
            foreach (string gk in groupKeys)
                foreach (string v in sel.Selected(gk))
                    if (!chosen.Contains(v)) chosen.Add(v);
            return chosen;
        }
    }
}
