using System;
using System.Text;

namespace ExchangeAuditTool
{
    internal enum ConnectionMode
    {
        ExchangeOnlineInteractive,
        ExchangeOnlineApp,
        OnPremisesLocal,
        OnPremisesRemote
    }

    // Authentication used for a remote on-premises PowerShell session.
    internal enum RemoteAuthMode
    {
        Kerberos,   // domain-joined: uses the current identity (or supplied credentials)
        Basic       // off-domain: prompts for credentials (should be used over HTTPS)
    }

    internal static class ConnectionSettings
    {
        public static ConnectionMode Mode = ConnectionMode.ExchangeOnlineInteractive;

        public static string Upn = "";
        public static bool DisableWam = true;

        public static string AppId = "";
        public static string Organization = "";
        public static string CertThumbprint = "";

        // ---- Remote on-premises (implicit remoting) settings ------------------------
        public static string RemoteServer = "";                       // FQDN or hostname of the Exchange server
        public static bool RemoteUseHttps = false;                    // http:// (default) vs https://
        public static RemoteAuthMode RemoteAuth = RemoteAuthMode.Kerberos;
        public static string RemoteUser = "";                         // optional username (default for the credential prompt)

        public static bool IsOnline
        {
            get { return Mode == ConnectionMode.ExchangeOnlineInteractive || Mode == ConnectionMode.ExchangeOnlineApp; }
        }

        // True when establishing a session can prompt (browser sign-in or
        // credential dialog). Such modes share a single audit session so the
        // user signs in exactly once; prompt-free modes run in parallel.
        public static bool RequiresInteractiveAuth
        {
            get
            {
                if (Mode == ConnectionMode.ExchangeOnlineInteractive) return true;
                if (Mode == ConnectionMode.OnPremisesRemote)
                    return RemoteAuth == RemoteAuthMode.Basic || !string.IsNullOrEmpty(RemoteUser);
                return false;
            }
        }

        public static string BuildPrelude()
        {
            var sb = new StringBuilder();
            sb.AppendLine("$ErrorActionPreference = 'Stop'");
            sb.AppendLine("$ProgressPreference = 'SilentlyContinue'");
            sb.AppendLine();

            if (Mode == ConnectionMode.ExchangeOnlineInteractive || Mode == ConnectionMode.ExchangeOnlineApp)
            {
                sb.AppendLine("if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {");
                sb.AppendLine("    throw 'The ExchangeOnlineManagement module is not installed. Use the Connection page to install it.'");
                sb.AppendLine("}");
                sb.AppendLine("Import-Module ExchangeOnlineManagement -ErrorAction Stop");
                sb.AppendLine();
                sb.AppendLine("$connected = $false");
                sb.AppendLine("try { if (Get-ConnectionInformation) { $connected = $true } } catch { $connected = $false }");
                sb.AppendLine("if (-not $connected) {");
                sb.AppendLine("    Write-Host 'Connecting to Exchange Online...'");
                if (Mode == ConnectionMode.ExchangeOnlineInteractive)
                {
                    string upnArg = string.IsNullOrEmpty(Upn) ? "" : " -UserPrincipalName " + ScriptContext.PsLiteral(Upn);
                    if (DisableWam)
                    {
                        sb.AppendLine("    Write-Host 'Your default web browser will open for sign-in...'");
                        sb.AppendLine("    Connect-ExchangeOnline -DisableWAM" + upnArg + " -ShowBanner:$false");
                    }
                    else
                    {
                        sb.AppendLine("    Connect-ExchangeOnline" + upnArg + " -ShowBanner:$false");
                    }
                }
                else
                {
                    sb.AppendLine("    Connect-ExchangeOnline -AppId " + ScriptContext.PsLiteral(AppId) +
                                  " -CertificateThumbprint " + ScriptContext.PsLiteral(CertThumbprint) +
                                  " -Organization " + ScriptContext.PsLiteral(Organization) + " -ShowBanner:$false");
                }
                sb.AppendLine("} else { Write-Host 'Reusing existing Exchange Online session.' }");
            }
            else if (Mode == ConnectionMode.OnPremisesRemote)
            {
                bool basic = RemoteAuth == RemoteAuthMode.Basic;
                bool forcedHttps = false;
                if (basic && !RemoteUseHttps)
                {
                    RemoteUseHttps = true;
                    forcedHttps = true;
                }
                string scheme = RemoteUseHttps ? "https://" : "http://";
                string uri = scheme + RemoteServer.Trim() + "/PowerShell/";
                string authName = basic ? "Basic" : "Kerberos";
                // Credentials are required for Basic; optional for Kerberos (current identity by default).
                bool promptCred = basic || !string.IsNullOrEmpty(RemoteUser);

                sb.AppendLine("if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {");
                sb.AppendLine("    Write-Host 'Opening remote PowerShell session to Exchange (" + authName + ")...'");
                if (forcedHttps)
                {
                    sb.AppendLine("    Write-Host 'Basic authentication requires HTTPS; forcing HTTPS.'");
                }
                sb.AppendLine("    $uri = " + ScriptContext.PsLiteral(uri));

                if (promptCred)
                {
                    // Get-Credential shows the native Windows credential dialog, so the
                    // password never appears in the Activity Log or the log file.
                    string userArg = string.IsNullOrEmpty(RemoteUser) ? "" : " -UserName " + ScriptContext.PsLiteral(RemoteUser);
                    sb.AppendLine("    Write-Host 'A secure credential prompt will appear...'");
                    sb.AppendLine("    $cred = Get-Credential -Message 'Exchange on-premises credentials'" + userArg);
                    sb.AppendLine("    if (-not $cred) { throw 'Credentials are required for this connection.' }");
                }

                sb.Append("    $global:EATRemoteSession = New-PSSession -ConfigurationName Microsoft.Exchange");
                sb.Append(" -ConnectionUri $uri -Authentication " + authName);
                if (promptCred) sb.Append(" -Credential $cred");
                if (basic)
                {
                    // Off-domain servers often use a self-signed certificate; skip the
                    // strict cert checks for Basic so the session can be established.
                    sb.Append(" -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)");
                }
                sb.AppendLine(" -AllowRedirection");
                sb.AppendLine("    Import-PSSession $global:EATRemoteSession -DisableNameChecking -AllowClobber | Out-Null");
                sb.AppendLine("    Write-Host 'Remote Exchange cmdlets imported.'");
                sb.AppendLine("} else { Write-Host 'Exchange cmdlets already available in this session.' }");
            }
            else
            {
                sb.AppendLine("if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {");
                sb.AppendLine("    Write-Host 'Loading local Exchange Management Shell...'");
                sb.AppendLine("    $loaded = $false");
                sb.AppendLine("    if ($env:ExchangeInstallPath -and (Test-Path (Join-Path $env:ExchangeInstallPath 'bin\\RemoteExchange.ps1'))) {");
                sb.AppendLine("        try {");
                sb.AppendLine("            . (Join-Path $env:ExchangeInstallPath 'bin\\RemoteExchange.ps1')");
                sb.AppendLine("            Connect-ExchangeServer -auto -ClientApplication:ManagementShell");
                sb.AppendLine("            $loaded = $true");
                sb.AppendLine("        } catch { $loaded = $false }");
                sb.AppendLine("    }");
                sb.AppendLine("    if (-not $loaded) {");
                sb.AppendLine("        $snap = Get-PSSnapin -Registered -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue");
                sb.AppendLine("        if ($snap) { Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010; $loaded = $true }");
                sb.AppendLine("    }");
                sb.AppendLine("    if (-not $loaded) { throw 'Could not load the local Exchange Management Shell. Run this tool directly on an Exchange server.' }");
                sb.AppendLine("} else { Write-Host 'Local Exchange cmdlets already available.' }");
            }

            sb.AppendLine();
            return sb.ToString();
        }

        public static string BuildDisconnect()
        {
            var sb = new StringBuilder();
            sb.AppendLine("$ErrorActionPreference = 'SilentlyContinue'");
            if (Mode == ConnectionMode.ExchangeOnlineInteractive || Mode == ConnectionMode.ExchangeOnlineApp)
                sb.AppendLine("try { Disconnect-ExchangeOnline -Confirm:$false } catch { }");
            else if (Mode == ConnectionMode.OnPremisesRemote)
                sb.AppendLine("try { if ($global:EATRemoteSession) { Remove-PSSession $global:EATRemoteSession; $global:EATRemoteSession = $null } } catch { }");
            sb.AppendLine("Write-Host 'Disconnected.'");
            return sb.ToString();
        }

        public static string BuildConnectedAsCheck()
        {
            var sb = new StringBuilder();
            if (Mode == ConnectionMode.ExchangeOnlineInteractive || Mode == ConnectionMode.ExchangeOnlineApp)
            {
                sb.AppendLine("try {");
                sb.AppendLine("    $ci = Get-ConnectionInformation | Select-Object -First 1");
                sb.AppendLine("    if ($ci) {");
                sb.AppendLine("        $who = if ($ci.UserPrincipalName) { $ci.UserPrincipalName } else { $ci.Organization }");
                sb.AppendLine("        Write-Host ('CONNECTEDAS ' + $who + ' | ' + $ci.Organization)");
                sb.AppendLine("    } else { Write-Host 'NOTCONNECTED' }");
                sb.AppendLine("} catch { Write-Host 'NOTCONNECTED' }");
            }
            else if (Mode == ConnectionMode.OnPremisesRemote)
            {
                // Prefer the supplied remote username; otherwise fall back to the current identity.
                string who = string.IsNullOrEmpty(RemoteUser)
                    ? "[System.Security.Principal.WindowsIdentity]::GetCurrent().Name"
                    : ScriptContext.PsLiteral(RemoteUser);
                sb.AppendLine("try {");
                sb.AppendLine("    $who = " + who);
                sb.AppendLine("    $org = ''");
                sb.AppendLine("    try { $org = (Get-OrganizationConfig).Name } catch { $org = '' }");
                sb.AppendLine("    if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {");
                sb.AppendLine("        Write-Host ('CONNECTEDAS ' + $who + ' @ ' + " + ScriptContext.PsLiteral(RemoteServer.Trim()) + " + ' | ' + $org)");
                sb.AppendLine("    } else { Write-Host 'NOTCONNECTED' }");
                sb.AppendLine("} catch { Write-Host 'NOTCONNECTED' }");
            }
            else
            {
                sb.AppendLine("try {");
                sb.AppendLine("    $me = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name");
                sb.AppendLine("    $org = ''");
                sb.AppendLine("    try { $org = (Get-OrganizationConfig).Name } catch { $org = '' }");
                sb.AppendLine("    if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {");
                sb.AppendLine("        Write-Host ('CONNECTEDAS ' + $me + ' | ' + $org)");
                sb.AppendLine("    } else { Write-Host 'NOTCONNECTED' }");
                sb.AppendLine("} catch { Write-Host 'NOTCONNECTED' }");
            }
            return sb.ToString();
        }

        public static string BuildModuleCheck()
        {
            var sb = new StringBuilder();
            sb.AppendLine("$m = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1");
            sb.AppendLine("if ($m) { Write-Host ('INSTALLED ' + $m.Version.ToString()) } else { Write-Host 'NOTINSTALLED' }");
            return sb.ToString();
        }

        public static string Summary()
        {
            switch (Mode)
            {
                case ConnectionMode.ExchangeOnlineInteractive:
                    return "Exchange Online (interactive)  -  " + (string.IsNullOrEmpty(Upn) ? "no UPN set" : Upn);
                case ConnectionMode.ExchangeOnlineApp:
                    return "Exchange Online (app-only)  -  " + (string.IsNullOrEmpty(Organization) ? "no tenant set" : Organization);
                case ConnectionMode.OnPremisesRemote:
                    return "On-premises remote (" + (RemoteAuth == RemoteAuthMode.Basic ? "Basic" : "Kerberos") + ")  -  " +
                           (string.IsNullOrEmpty(RemoteServer) ? "no server set" : RemoteServer);
                default:
                    return "On-premises local (Exchange server shell)";
            }
        }
    }
}
