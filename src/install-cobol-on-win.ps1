#Requires -Version 5.1
<#
.SYNOPSIS
    Installs a portable COBOL development environment on Windows:
    VSCodium + SuperBOL extension + GnuCOBOL 3.2 64-bit (bundled MinGW GCC).

.DESCRIPTION
    - No administrator rights required, nothing is compiled: only ready-made archives.
    - Everything goes into a single folder (default C:\cobol-dev):
        VSCodium\         editor in portable mode (settings and extensions in VSCodium\data)
        GnuCOBOL\         cobc compiler + GCC + gdb (MSYS2 build by Arnold Trembley / Chuck Haatvedt)
        projects\         cobol-example-project, ready to build, holding every
                          example (hello\, subroutine\, access\, mariadb\)
        CoBOL-Codium.cmd  sets the GnuCOBOL environment, then opens VSCodium
    - Where does GCC come from? GnuCOBOL translates COBOL into C, then calls a C compiler.
      The GnuCOBOL archive used here ships its own GCC (13.1, MinGW-w64 via MSYS2):
      nothing else needs to be installed.
    - This script lives in src\; its binary sources (VSCodium zip, GnuCOBOL archive,
      7zr.exe, GixSQL archive, embedded under ..\bin\) and its sample project
      (src\example\) are used as-is when found next to the script (no network
      access); otherwise each missing file is downloaded, either from -VSCodiumUrl /
      -GnuCobolUrl / -SevenZipUrl / -GixSqlUrl, or from -ExampleBaseUrl for the
      sample project. This lets the project be hosted at https://www.papit.fr/cow
      and this script still work when only it, alone, is downloaded and run - those
      URLs and -ProjectBinDir can still be overridden on the command line for
      testing or to point at different sources.
    - SQL database access: GixSQL (an EXEC SQL preprocessor for GnuCOBOL) is
      installed alongside GnuCOBOL, ready to use with PostgreSQL, MySQL/MariaDB,
      SQLite and (via ODBC) any database with an ODBC driver, including DB2 (its
      driver is embedded/downloaded the same way as everything else - see
      -Db2CliUrl). Use the "COBOL: build (SQL)" task (Terminal > Run Task...) to
      compile a program containing EXEC SQL statements.
    - The project VSCodium opens holds every example, each in its own folder, with
      hello\hello-world.cob active:
        hello\        DISPLAY, the smallest possible program
        subroutine\   a program CALLing a COBOL library through a shared copy
                      (build greeter-lib.cob with "COBOL: build (module .dll)"
                      once, then F7 on greeter-main.cob)
        access\       EXEC SQL INSERT/SELECT over ODBC, on the Access database
                      sitting right there in the folder (sample-store.mdb,
                      copied from .\bin\example\access\ and reached through the
                      "SampleStore" DSN) - "COBOL: build (SQL)"
        mariadb\      EXEC SQL on a public MariaDB server - "COBOL: build (SQL)"
    - Excel/Access (.xlsx/.accdb) via ODBC: optional, and only installed when this
      script itself runs elevated (as Administrator) - it is the one piece of this
      project that is a real system-wide installer, not a portable unzip. Without
      elevation you get a clear warning instead and everything else still installs.
    - Connector versions pinned by the URLs/archives above (as embedded when this
      was written; the console output at install time always shows what actually
      got installed, which may differ once -VSCodiumUrl etc. point elsewhere):
        VSCodium                1.135.06055
        GnuCOBOL                3.2.0 (bundled GCC 13.1.0, MinGW)
        7-Zip (7zr.exe)         26.04
        GixSQL                  1.0.20b-1
          - PostgreSQL client (libpq)         15.1
          - MySQL/MariaDB client (libmariadb) 3.1.13
          - SQLite                            bundled (~3.38.0)
          - Oracle                            needs Oracle Instant Client (OCI),
                                               not bundled
          - ODBC                              generic; see DB2/Access below for
                                               an actual driver
        DB2 CLI/ODBC driver     IBM DS Driver, resolved via -Db2CliUrl like every
                                other binary above (version printed at install time)
        Access Database Engine  16.0.5044.1000 (2016 redistributable; optional,
                                admin only)
    - Actions (keyboard shortcut in a COBOL file, or menu):
        Ctrl + F5   or  Run > Start Without Debugging'
        F5          or  Run > Start Debugging'
        F6          or  Terminal > Run Build Task... (Ctrl+Shift+B)'
        F7          or  Terminal > Run Task... > COBOL: build and run'
        F8          or  Terminal > Run Task... > COBOL: build (debug)'
        F9          or  Terminal > Run Task... > COBOL: build (SQL)'
        F10         or  Terminal > Run Task... > COBOL: build (module .dll)'
      F9/F10 build only outside a debug session: while one is running they stay
      the debugger's own toggle breakpoint / step over.
.PARAMETER InstallDir
    Installation folder (default c:\cobol-dev). A short path, without spaces or
    accented characters, is recommended for gdb and GnuCOBOL.

.PARAMETER VSCodiumUrl
    VSCodium portable zip to use (default: a pinned win32-x64 release). Used only
    when no file of the same name exists in -ProjectBinDir.

.PARAMETER GnuCobolUrl
    GnuCOBOL .7z archive to use (default: GnuCOBOL 3.2 MSYS2 64-bit). Used only
    when no file of the same name exists in -ProjectBinDir.

.PARAMETER SevenZipUrl
    Standalone 7-Zip extractor (7zr.exe) used to unpack the GnuCOBOL archive. Used
    only when no file of the same name exists in -ProjectBinDir.

.PARAMETER GixSqlUrl
    GixSQL .7z archive to use (default: a pinned win32-x64/MinGW build). GixSQL is
    an EXEC SQL preprocessor for GnuCOBOL with runtime support for PostgreSQL,
    MySQL/MariaDB, SQLite, Oracle and ODBC (so SQL Server, DB2, etc. too). Used
    only when no file of the same name exists in -ProjectBinDir.

.PARAMETER Db2CliUrl
    IBM's DB2 CLI/ODBC driver ("DS Driver", a plain zip - no installer, no
    registration, just PATH) to use. Resolved the same way as every other binary
    in this script: a file of the same name in -ProjectBinDir is used as-is,
    otherwise it is downloaded from this URL (the papit.fr/cow mirror by default,
    since IBM itself has no stable, login-free download URL for it).

.PARAMETER AccessEngineUrl
    Microsoft Access Database Engine redistributable (.exe) to use for Excel/Access
    (.xlsx/.accdb) access via ODBC. This is a real system-wide installer (registry,
    Program Files), unlike everything else in this script: it is only ever run when
    this script itself is running elevated (as Administrator). Without elevation,
    Excel/Access support is skipped and a warning tells you to rerun as admin;
    every other feature (PostgreSQL/MySQL/MariaDB/SQLite/Oracle/DB2) is unaffected.

.PARAMETER ProjectBinDir
    Folder holding the project's embedded binary sources (default: ..\bin next to
    this script). Checked first, by filename, for VSCodiumUrl/GnuCobolUrl/SevenZipUrl
    before falling back to a download.

.PARAMETER ExampleBaseUrl
    Base URL hosting src\example\ (default: this project's site). Used file by file
    for whichever sample-project file isn't found next to this script, in src\example\.

.PARAMETER ExampleBinBaseUrl
    Base URL hosting bin\example\ (default: this project's site), for the example
    data files that are binaries rather than sources: the Access sample database
    bin\example\access\sample-store.mdb, which is copied into the generated
    project. Used only when that file isn't found in -ProjectBinDir.

.PARAMETER PersistUserEnv
    Also stores the GnuCOBOL variables in the Windows user environment, so cobc can be
    used from any terminal (otherwise: use CoBOL-Codium.cmd).

.PARAMETER Force
    Downloads and extracts everything again (VSCodium settings and extensions are kept).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\install-cobol-on-win.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\install-cobol-on-win.ps1 -InstallDir d:\cobol -PersistUserEnv
#>
[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:SystemDrive.ToLower() 'cobol-dev'),
    # Original from https://github.com/VSCodium/vscodium/releases
    [string]$VSCodiumUrl = 'https://www.papit.fr/cow/bin/VSCodium-win32-x64-1.135.06055.zip',
    # Original from $GnuCobolUrl https://www.arnoldtrembley.com/GC32M-BDB-x64.7z
    [string]$GnuCobolUrl = 'https://www.papit.fr/cow/bin/GC32M-BDB-x64.7z',
    # Original from https://www.7-zip.org/a/7zr.exe
    [string]$SevenZipUrl = 'https://www.papit.fr/cow/bin/7zr.exe',
    # Original from https://github.com/mridoni/gixsql/releases
    [string]$GixSqlUrl = 'https://www.papit.fr/cow/bin/gixsql-binaries-windows-x64-mingw-1.0.20b-1.7z',
    # Original: IBM's DB2 CLI/ODBC driver ("DS Driver") has no stable, login-free
    # download URL of its own, so the papit.fr/cow mirror is the only default here.
    [string]$Db2CliUrl = 'https://www.papit.fr/cow/bin/ntx64_odbc_cli.zip',
    # Original from Microsoft: see .PARAMETER AccessEngineUrl below. Only installed
    # when this script itself is running elevated (Administrator) - see section 2c.
    [string]$AccessEngineUrl = 'https://www.papit.fr/cow/bin/accessdatabaseengine_X64.exe',
    # Embedded binary sources for this project (see CLAUDE.md: src/ + bin/).
    [string]$ProjectBinDir,
    # Hosted copy of src/example/, used the same way as -ProjectBinDir when this
    # script runs standalone (see CLAUDE.md).
    [string]$ExampleBaseUrl = 'https://www.papit.fr/cow/src/example',
    # Same, for the example data files that are binaries and so live in
    # bin/example/ instead: the Access sample database.
    [string]$ExampleBinBaseUrl = 'https://www.papit.fr/cow/bin/example',
    [switch]$PersistUserEnv,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # makes Invoke-WebRequest much faster on PowerShell 5.1
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$UserAgent = @{ 'User-Agent' = 'install-cobol-on-win.ps1' }

# $PSScriptRoot is empty while the param block's defaults are evaluated (seen with
# "powershell -File"), so the project's src\/bin\ folders are located here instead.
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $ProjectBinDir) { $ProjectBinDir = Join-Path $ScriptDir '..\bin' }
$ExampleDir = Join-Path $ScriptDir 'example'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-FileProductVersion([string]$Path) {
    if (-not (Test-Path $Path)) { return $null }
    $v = (Get-Item $Path).VersionInfo.ProductVersion
    if ($v) { return $v.Trim() }
    return $null
}

# SQLite is statically linked into libgixsql-sqlite.dll (no separate version
# resource): its own version string is embedded as plain text by the amalgamation
# build, so it is recovered by scanning the file's readable text instead.
function Get-BundledSqliteVersion([string]$DllPath) {
    if (-not (Test-Path $DllPath)) { return $null }
    try {
        $text = [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($DllPath))
        $m = [regex]::Match($text, '3\.\d{2}\.\d{1,2}')
        if ($m.Success) { return $m.Value }
    }
    catch {}
    return $null
}

# Confirmed GixSQL 1.0.20b bug: when gixpp emits the lazy "ESQL CURSOR
# DECLARATIONS" trailer for EXEC SQL DECLARE CURSOR, the GIXSQLCursorDeclare
# CALL's closing END-CALL is missing its terminating period, so GnuCOBOL 3.2
# reads the next line (the GIX-SKIP-CRSR-INIT paragraph name) as part of the
# same unfinished sentence instead of a new paragraph - a syntax error that
# breaks every cursor-based program as gixpp emits it. Patched here with a
# narrow, idempotent regex rather than in gixpp.exe itself (no toolchain
# available to rebuild its own preprocessor from source in this project).
# The same patch is applied inside the IDE by src\example\.vscode\build-sql.ps1,
# which backs the "COBOL: build (SQL)" task: keep the two in step.
function Repair-GixsqlCursorTrailer([string]$CbsqlPath) {
    if (-not (Test-Path $CbsqlPath)) { return }
    $text = [IO.File]::ReadAllText($CbsqlPath)
    $pattern = '(?m)^(?<call>GIXSQL[ \t]*END-CALL)[ \t]*(?<nl>\r?\n)(?=GIXSQL[ \t]+[\w-]+\.[ \t]*(\r?\n|$))'
    $fixed = [regex]::Replace($text, $pattern, '${call}.${nl}')
    if ($fixed -ne $text) {
        [IO.File]::WriteAllText($CbsqlPath, $fixed)
    }
}

# cmd.exe reads "%" as the start of a variable reference, so a literal one in
# a value has to be doubled. Used for both the per-example launchers and
# CoBOL-Codium.cmd.
function ConvertTo-CmdValue([string]$Value) { $Value -replace '%', '%%' }

# Writes a .cmd next to a compiled example that sets the GnuCOBOL/GixSQL
# runtime environment and then runs the .exe. Without it the .exe cannot be
# started from Explorer or a plain terminal at all: it needs libcob-4.dll
# (and, for the EXEC SQL examples, libgixsql + its driver and client libs) on
# PATH, and dies with 0xC0000135 STATUS_DLL_NOT_FOUND otherwise. The
# alternative - copying a dozen DLLs into each example folder - would
# duplicate the whole runtime four times over.
function New-ExampleLauncher([string]$ExePath) {
    $cmdFile = [IO.Path]::ChangeExtension($ExePath, '.cmd')
    $body = New-Object Text.StringBuilder
    [void]$body.AppendLine('@echo off')
    [void]$body.AppendLine('rem Generated by install-cobol-on-win.ps1: runtime environment, then the example')
    foreach ($key in $CobEnv.Keys) {
        [void]$body.AppendLine("set `"$key=$(ConvertTo-CmdValue $CobEnv[$key])`"")
    }
    [void]$body.AppendLine("set `"PATH=$(ConvertTo-CmdValue ($CobPathAdd -join ';'));%PATH%`"")
    [void]$body.AppendLine('cd /d "%~dp0"')
    [void]$body.AppendLine("`"%~dp0$(Split-Path $ExePath -Leaf)`" %*")
    # Holds the window open so a double-click from Explorer is readable. This
    # .cmd is the entry point for a person; anything automated should set the
    # same environment once and run the .exe directly, the way this script's
    # own test steps below do, instead of waiting on a keypress.
    [void]$body.AppendLine('pause')

    $oemEnc = [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
    [IO.File]::WriteAllText($cmdFile, $body.ToString(), $oemEnc)
    return $cmdFile
}

# Compiles one example from src\example\<group>\ into bin\example\<group>\, so
# what ships in bin\ is ready to run with no toolchain knowledge. Throws on
# failure (callers decide whether that is fatal) after writing the compiler
# output to a log: a failed GixSQL link can dump hundreds of "undefined
# reference" lines, which would flood the console for no actionable benefit.
#   -Sql     : preprocess EXEC SQL with gixpp first, and link libgixsql
#   -Module  : build a dynamically CALLable .dll (-m) instead of an .exe (-x)
# Both -I paths matter: the example's own folder, plus src\example\ itself for
# the copybooks shared by every example (sqlca.cpy). gixpp resolves
# "EXEC SQL INCLUDE", cobc resolves plain "COPY".
function New-CobolExampleBinary {
    param(
        [Parameter(Mandatory)][string]$Source,
        [switch]$Sql,
        [switch]$Module
    )

    $srcDir = Split-Path $Source -Parent
    $group = Split-Path $srcDir -Leaf
    $name = [IO.Path]::GetFileNameWithoutExtension($Source)
    # GetFullPath collapses the "src\..\bin" that $ProjectBinDir carries, so
    # the paths reported to the customer read as real locations.
    $outDir = [IO.Path]::GetFullPath((Join-Path (Join-Path $ProjectBinDir 'example') $group))
    $outFile = Join-Path $outDir ($name + $(if ($Module) { '.dll' } else { '.exe' }))
    $log = Join-Path $Downloads "$name-build.log"

    New-Item -ItemType Directory -Force -Path $outDir | Out-Null

    $toCompile = $Source
    $cbsql = $null
    try {
        if ($Sql) {
            $cbsql = Join-Path $outDir "$name.cbsql"
            $ppOut = & $gixpp.FullName -e -S -P varchar -I $srcDir -I $ExampleDir -i $Source -o $cbsql 2>&1
            if ($LASTEXITCODE -ne 0) {
                $ppOut | Out-File $log -Encoding utf8
                throw "gixpp failed on $name.cob (exit code $LASTEXITCODE) - see $log"
            }
            Repair-GixsqlCursorTrailer $cbsql
            $toCompile = $cbsql
        }

        $cobcArgs = @($(if ($Module) { '-m' } else { '-x' }), '-I', $srcDir, '-I', $ExampleDir, '-o', $outFile)
        if ($Sql) { $cobcArgs += @('-L', $GixSqlLib, '-lgixsql') }
        $cobcOut = & $cobc.FullName @cobcArgs $toCompile 2>&1
        if ($LASTEXITCODE -ne 0) {
            $cobcOut | Out-File $log -Encoding utf8
            throw "cobc failed on $name (exit code $LASTEXITCODE) - see $log"
        }
    }
    finally {
        # The .cbsql is a gixpp intermediate, not a deliverable: keep bin\ to
        # just what the customer runs.
        if ($cbsql) { Remove-Item $cbsql -ErrorAction SilentlyContinue }
    }

    Remove-Item $log -ErrorAction SilentlyContinue
    return $outFile
}

function Save-Utf8NoBom([string]$Path, [string]$Content) {
    # UTF-8 without BOM: a BOM at the start of a COBOL source confuses cobc.
    $dir = Split-Path $Path -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [IO.File]::WriteAllText($Path, $Content, (New-Object Text.UTF8Encoding $false))
}

function Save-Json([string]$Path, $Object) {
    Save-Utf8NoBom $Path ($Object | ConvertTo-Json -Depth 10)
}

function Get-RemoteFile([string]$Url, [string]$OutFile) {
    if ((Test-Path $OutFile) -and -not $Force) {
        Write-Host "   already downloaded: $(Split-Path $OutFile -Leaf)"
        return
    }
    Write-Host "   downloading: $Url"
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -Headers $UserAgent
}

# Prefers a local embedded copy (under $LocalDir, at $RelativePath) over the network,
# so the project still works when only this script is downloaded: its embedded
# sources/binaries double as the hosted fallback at papit.fr/cow (see CLAUDE.md).
function Get-EmbeddedOrRemoteFile([string]$RelativePath, [string]$Url, [string]$OutFile, [string]$LocalDir = $ProjectBinDir) {
    if ((Test-Path $OutFile) -and -not $Force) {
        Write-Host "   already downloaded: $(Split-Path $OutFile -Leaf)"
        return
    }
    # The sample project is mirrored folder by folder (cobol-example-project\
    # access\, subroutine\, .vscode\...), so the destination tree often does
    # not exist yet: created here once, for the embedded copy and the download
    # alike (Invoke-WebRequest -OutFile does not create it).
    $destDir = Split-Path $OutFile -Parent
    if ($destDir -and -not (Test-Path $destDir)) { New-Item -ItemType Directory -Force -Path $destDir | Out-Null }
    $embedded = Join-Path $LocalDir $RelativePath
    if (Test-Path $embedded) {
        Write-Host "   using embedded file: $embedded"
        Copy-Item -LiteralPath $embedded -Destination $OutFile -Force
        return
    }
    Get-RemoteFile $Url $OutFile
}

# ---------------------------------------------------------------------------
# Folders
# ---------------------------------------------------------------------------
$Downloads = Join-Path $InstallDir '_downloads'
$CodiumDir = Join-Path $InstallDir 'VSCodium'
$CobolDir = Join-Path $InstallDir 'GnuCOBOL'
$GixSqlDir = Join-Path $InstallDir 'GixSQL'
$ProjectDir = Join-Path $InstallDir 'projects\cobol-example-project'
New-Item -ItemType Directory -Force -Path $InstallDir, $Downloads | Out-Null

Write-Host "Installing the COBOL environment into: $InstallDir" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 1. Portable VSCodium
# ---------------------------------------------------------------------------
Write-Step 'VSCodium (portable zip)'

$codiumZip = Join-Path $Downloads (Split-Path $VSCodiumUrl -Leaf)
Get-EmbeddedOrRemoteFile (Split-Path $VSCodiumUrl -Leaf) $VSCodiumUrl $codiumZip

$codiumExe = Join-Path $CodiumDir 'VSCodium.exe'
if ($Force -or -not (Test-Path $codiumExe)) {
    Write-Host "   extracting $(Split-Path $codiumZip -Leaf)..."
    Expand-Archive -Path $codiumZip -DestinationPath $CodiumDir -Force
}
# A "data" folder enables portable mode: nothing is written to %APPDATA%.
New-Item -ItemType Directory -Force -Path (Join-Path $CodiumDir 'data') | Out-Null
Write-Host "   VSCodium ready ($(Split-Path $codiumZip -Leaf))."

# ---------------------------------------------------------------------------
# 2. GnuCOBOL 3.2 64-bit (with bundled GCC and gdb)
# ---------------------------------------------------------------------------
Write-Step 'GnuCOBOL 3.2 64-bit (bundled MinGW GCC)'

# 7zr.exe: official standalone 7-Zip extractor (no installation)
$sevenZr = Join-Path $Downloads '7zr.exe'
Get-EmbeddedOrRemoteFile (Split-Path $SevenZipUrl -Leaf) $SevenZipUrl $sevenZr

$cobolArchive = Join-Path $Downloads (Split-Path $GnuCobolUrl -Leaf)
Get-EmbeddedOrRemoteFile (Split-Path $GnuCobolUrl -Leaf) $GnuCobolUrl $cobolArchive

$cobc = Get-ChildItem $CobolDir -Recurse -Filter 'cobc.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($Force -or -not $cobc) {
    Write-Host '   extracting (may take a minute)...'
    & $sevenZr x $cobolArchive "-o$CobolDir" -y | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "GnuCOBOL extraction failed (exit code $LASTEXITCODE)." }
    $cobc = Get-ChildItem $CobolDir -Recurse -Filter 'cobc.exe' | Select-Object -First 1
}
if (-not $cobc) { throw "cobc.exe not found in $CobolDir." }

$CobBin = $cobc.DirectoryName
$CobRoot = Split-Path $CobBin -Parent
Write-Host "   cobc found: $($cobc.FullName)"

# ---------------------------------------------------------------------------
# 2b. GixSQL (EXEC SQL preprocessor: PostgreSQL, MySQL/MariaDB, SQLite, Oracle,
#     ODBC - so DB2 too, see section 2d below)
# ---------------------------------------------------------------------------
Write-Step 'GixSQL ESQL preprocessor'

$gixSqlArchive = Join-Path $Downloads (Split-Path $GixSqlUrl -Leaf)
Get-EmbeddedOrRemoteFile (Split-Path $GixSqlUrl -Leaf) $GixSqlUrl $gixSqlArchive

$gixpp = Get-ChildItem $GixSqlDir -Recurse -Filter 'gixpp.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($Force -or -not $gixpp) {
    Write-Host '   extracting...'
    & $sevenZr x $gixSqlArchive "-o$GixSqlDir" -y | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "GixSQL extraction failed (exit code $LASTEXITCODE)." }
    $gixpp = Get-ChildItem $GixSqlDir -Recurse -Filter 'gixpp.exe' | Select-Object -First 1
}
if (-not $gixpp) { throw "gixpp.exe not found in $GixSqlDir." }

$GixSqlBin = $gixpp.DirectoryName
$GixSqlRoot = Split-Path $GixSqlBin -Parent
$GixSqlLib = Get-ChildItem (Join-Path $GixSqlRoot 'lib') -Recurse -Filter 'libgixsql-sqlite.dll' -ErrorAction SilentlyContinue |
Select-Object -First 1 | ForEach-Object { $_.DirectoryName }
if (-not $GixSqlLib) { throw "GixSQL runtime libraries not found under $GixSqlRoot\lib." }
$gixppVerOut = & $gixpp.FullName --version 2>&1
$gixppVerMatch = $gixppVerOut | Select-String 'Version:\s*(\S+)' | Select-Object -First 1
$gixppVer = if ($gixppVerMatch) { $gixppVerMatch.Matches[0].Groups[1].Value } else { $null }
$pgVer = Get-FileProductVersion (Join-Path $GixSqlLib 'libpq.dll')
$mariaVer = Get-FileProductVersion (Join-Path $GixSqlLib 'libmariadb.dll')
$sqliteVer = Get-BundledSqliteVersion (Join-Path $GixSqlLib 'libgixsql-sqlite.dll')

Write-Host "   gixpp found: $($gixpp.FullName) (GixSQL $gixppVer)"
Write-Host "   runtime libraries: $GixSqlLib"
Write-Host "     - PostgreSQL client (libpq)        : $(if ($pgVer) { $pgVer } else { 'version not found' })"
Write-Host "     - MySQL/MariaDB client (libmariadb) : $(if ($mariaVer) { $mariaVer } else { 'version not found' })"
Write-Host "     - SQLite                            : bundled$(if ($sqliteVer) { ", $sqliteVer" })"
Write-Host '     - Oracle                            : needs Oracle Instant Client (OCI), not bundled'
Write-Host '     - ODBC                               : generic; see DB2/Access below for an actual driver'

# ---------------------------------------------------------------------------
# 2c. Excel/Access connector (optional, admin only)
#    The Microsoft Access Database Engine is a real system-wide installer
#    (registry, Program Files) - the one piece of this project that isn't a
#    portable unzip. It is only run when this script itself is elevated.
# ---------------------------------------------------------------------------
Write-Step 'Excel/Access connector (optional)'

if (Test-IsAdmin) {
    $accessEngineExe = Join-Path $Downloads (Split-Path $AccessEngineUrl -Leaf)
    Get-EmbeddedOrRemoteFile (Split-Path $AccessEngineUrl -Leaf) $AccessEngineUrl $accessEngineExe
    $aceVer = Get-FileProductVersion $accessEngineExe

    Write-Host "   installing Microsoft Access Database Engine $aceVer (admin, system-wide, silent)..."
    $aceProc = Start-Process -FilePath $accessEngineExe -ArgumentList '/quiet', '/norestart' -Wait -PassThru
    if ($aceProc.ExitCode -eq 0) {
        Write-Host '   Excel (.xlsx) / Access (.accdb) ready via ODBC.' -ForegroundColor Green
    }
    else {
        Write-Host "   installer exited with code $($aceProc.ExitCode) (may already be installed)." -ForegroundColor Yellow
    }
}
else {
    Write-Host ''
    Write-Host '   ############################################################'  -ForegroundColor Red
    Write-Host '   #  EXCEL / ACCESS NOT INSTALLED                             #' -ForegroundColor Red
    Write-Host '   #  This script is not running as Administrator.            #' -ForegroundColor Red
    Write-Host '   #  Run the script as admin to install the Microsoft Access #' -ForegroundColor Red
    Write-Host '   #  Database Engine (ODBC driver for .xlsx / .accdb files). #' -ForegroundColor Red
    Write-Host '   #  Right-click PowerShell > Run as administrator, then     #' -ForegroundColor Red
    Write-Host '   #  rerun this script.                                      #' -ForegroundColor Red
    Write-Host '   ############################################################'  -ForegroundColor Red
    Write-Host ''
}

# ---------------------------------------------------------------------------
# 2d. DB2 ODBC/CLI driver - embedded in bin\, or downloaded from -Db2CliUrl,
#    exactly like every other binary above. IBM's "DS Driver" is a plain zip
#    by design: no installer, no registration needed, just PATH.
# ---------------------------------------------------------------------------
Write-Step 'DB2 ODBC/CLI driver'

function Get-Db2DriverVersion([string]$DriverDir) {
    $db2level = Join-Path $DriverDir 'db2level.exe'
    if (-not (Test-Path $db2level)) { return $null }
    $out = & $db2level 2>&1
    $m = ($out | Select-String 'DB2 v[\d.]+') | Select-Object -First 1
    if ($m) { return $m.Matches[0].Value }
    return $null
}

$db2Zip = Join-Path $Downloads (Split-Path $Db2CliUrl -Leaf)
Get-EmbeddedOrRemoteFile (Split-Path $Db2CliUrl -Leaf) $Db2CliUrl $db2Zip

$Db2Dir = Join-Path $InstallDir 'DB2CLI'
$db2cli = Get-ChildItem $Db2Dir -Recurse -Filter 'db2cli64.dll' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($Force -or -not $db2cli) {
    Write-Host '   extracting...'
    Expand-Archive -Path $db2Zip -DestinationPath $Db2Dir -Force
    $db2cli = Get-ChildItem $Db2Dir -Recurse -Filter 'db2cli64.dll' | Select-Object -First 1
}
if (-not $db2cli) { throw "db2cli64.dll not found in $Db2Dir." }

$Db2DriverDir = $db2cli.DirectoryName
$db2Ver = Get-Db2DriverVersion $Db2DriverDir
Write-Host "   DB2 ODBC/CLI driver ready: $Db2DriverDir$(if ($db2Ver) { " ($db2Ver)" })"

# ---------------------------------------------------------------------------
# 3. GnuCOBOL environment variables
#    Run the set_env.cmd shipped with the package and capture what it defines:
#    more reliable than guessing paths (COB_CFLAGS, COB_LDFLAGS, etc.).
# ---------------------------------------------------------------------------
Write-Step 'GnuCOBOL environment variables'

$CobEnv = [ordered]@{}
$CobPathAdd = @()
$setEnvCmd = Get-ChildItem $CobRoot -Filter 'set_env.cmd' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

if ($setEnvCmd) {
    $oldPath = $env:PATH -split ';' | Where-Object { $_ }
    $lines = cmd.exe /d /c "cd /d `"$($setEnvCmd.DirectoryName)`" && call set_env.cmd >nul 2>&1 <nul && set"
    foreach ($line in $lines) {
        if ($line -match '^([^=]+)=(.*)$') {
            $name, $value = $Matches[1], $Matches[2]
            if ($name -like 'COB*') {
                $CobEnv[$name] = $value
            }
            elseif ($name -eq 'Path' -or $name -eq 'PATH') {
                # @(...) forces an array even when exactly one entry matches - otherwise
                # PowerShell unwraps a single match to a bare string, and every later
                # "$CobPathAdd + ..." below silently becomes string concatenation instead
                # of array append, corrupting PATH.
                $CobPathAdd = @($value -split ';' | Where-Object { $_ -and ($oldPath -notcontains $_) })
            }
        }
    }
    Write-Host "   $($CobEnv.Count) COB_* variable(s) read from $($setEnvCmd.Name)."
}

if ($CobEnv.Count -eq 0) {
    Write-Host '   set_env.cmd missing or silent: values derived from the folder layout.' -ForegroundColor Yellow
    $CobEnv['COB_MAIN_DIR'] = "$CobRoot\"
    foreach ($pair in @(@('COB_CONFIG_DIR', 'config'), @('COB_COPY_DIR', 'copy'))) {
        $dir = Join-Path $CobRoot $pair[1]
        if (Test-Path $dir) { $CobEnv[$pair[0]] = $dir }
    }
    $inc = Join-Path $CobRoot 'include'
    $lib = Join-Path $CobRoot 'lib'
    if (Test-Path $inc) { $CobEnv['COB_CFLAGS'] = "-I`"$inc`"" }
    if (Test-Path $lib) { $CobEnv['COB_LDFLAGS'] = "-L`"$lib`"" }
}

# GixSQL always substitutes host variables as "$n" (its own internal
# convention) regardless of target DBMS, then is supposed to let each
# driver's fixup_parameters pass translate that into whatever placeholder
# style the real driver needs ("?" for ODBC/MySQL/SQLite, "$n" - a no-op -
# for PostgreSQL). It defaults OFF, and even when turned on, GixSQL
# 1.0.20b's ODBC and MySQL drivers only apply it on the PREPARE/EXECUTE
# path, not on a plain "EXEC SQL INSERT/UPDATE/DELETE ... :host-var": that
# gap is independently patched into this project's embedded ODBC driver
# (see 7c below), but the flag is still needed to engage it, and still
# matters as-is for any other driver/path. Set globally here so it always
# applies, not just for this script's own test steps.
$CobEnv['GIXSQL_FIXUP_PARAMS'] = '1'

# libgixsql otherwise drops an (empty) gixsql.log in the working directory of
# every program that touches SQL, which would litter bin\example\. "off" is
# the only value that suppresses the file; "none"/"err" still create it. Raise
# it to trace, with GIXSQL_LOG_FILE, when debugging an EXEC SQL program.
$CobEnv['GIXSQL_LOG_LEVEL'] = 'off'

if ($CobPathAdd.Count -eq 0) { $CobPathAdd = @($CobBin) }

# cobc shells out to "gcc" by name (not by full path): its folder must be on PATH.
# set_env.cmd is supposed to add it, but isn't always reliable, so it's added
# explicitly here too (harmless if already present via set_env.cmd).
$gccExe = Get-ChildItem $CobRoot -Recurse -Filter 'gcc.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($gccExe -and ($CobPathAdd -notcontains $gccExe.DirectoryName)) {
    $CobPathAdd = @($CobPathAdd + $gccExe.DirectoryName)
}

# GixSQL's preprocessor deps and runtime driver libraries, plus the DB2 driver:
# added to the same PATH as GnuCOBOL so they flow through to the
# terminal/launcher/persisted-env steps below unchanged.
$CobPathAdd = @($CobPathAdd + $GixSqlBin + $GixSqlLib + $Db2DriverDir)

# Apply to the current session (for the test below)
foreach ($key in $CobEnv.Keys) { Set-Item -Path "env:$key" -Value $CobEnv[$key] }
$env:PATH = (($CobPathAdd -join ';') + ';' + $env:PATH)

if ($gccExe) { Write-Host "   bundled GCC: $(& $gccExe.FullName -dumpfullversion 2>$null) ($($gccExe.FullName))" }
Write-Host "   $((& $cobc.FullName --version 2>&1 | Select-Object -First 1))"

# ---------------------------------------------------------------------------
# 4. SuperBOL extension (from Open VSX, the registry used by VSCodium)
# ---------------------------------------------------------------------------
Write-Step 'SuperBOL Studio OSS extension'

$codiumCli = Join-Path $CodiumDir 'bin\codium.cmd'
& $codiumCli --install-extension 'OCamlPro.SuperBOL' --force
if ($LASTEXITCODE -ne 0) { throw "SuperBOL extension installation failed (exit code $LASTEXITCODE)." }

# ---------------------------------------------------------------------------
# 4b. Windows fix for the SuperBOL debugger
#   When gdb stops on a breakpoint, the debugger checks that the C line maps to a
#   COBOL line by comparing paths verbatim (e.fileC === fileC). On Windows its table
#   holds the path lowercased with "\", while gdb returns "C:/..." or "C:\...": the
#   comparison fails, the debugger thinks it is outside COBOL code and resumes.
#   => gdb hits the breakpoints but execution never stops in the editor.
#   The fix normalizes paths (lowercase, "\") before comparing, and passes the source
#   directories to gdb with "/" (otherwise gdb gets "C:cobol-devprojects...").
# ---------------------------------------------------------------------------
Write-Step 'Windows fix for the SuperBOL debugger'

function Repair-SuperbolDebugger([string]$ExtensionsDir) {
    # The debugger runs "inline" in the extension host: its code is loaded from the
    # main bundle, whose name depends on the version (superbol-vscode-bundle.js in
    # 1.0.0, superbol-vscode-oss-bundle.js later). superbol-vscode-gdb.js is only the
    # standalone entry point. So every .js in _dist containing that code is patched.
    $bundles = Get-ChildItem $ExtensionsDir -Directory -Filter 'ocamlpro.superbol-*' -ErrorAction SilentlyContinue |
    ForEach-Object { Get-ChildItem (Join-Path $_.FullName '_dist') -Filter '*.js' -ErrorAction SilentlyContinue } |
    Where-Object { Select-String -LiteralPath $_.FullName -SimpleMatch 'gdb-set directories' -Quiet } |
    ForEach-Object { $_.FullName }
    if (-not $bundles) {
        Write-Host '   SuperBOL debugger not found: fix not applied.' -ForegroundColor Yellow
        return
    }
    foreach ($bundle in $bundles) {
        $js = [IO.File]::ReadAllText($bundle)
        if ($js.Contains('__sbN(')) { Write-Host "   already patched: $bundle"; continue }

        $count = 0
        foreach ($pair in @(
                @('e.fileC===fileC', '__sbN(e.fileC)===__sbN(fileC)'),
                @('e.fileCobol===fileCobol', '__sbN(e.fileCobol)===__sbN(fileCobol)'),
                @("sourceDirs.join('`" `"')", "sourceDirs.map(function(d){return String(d).replace(/\\/g,'/')}).join('`" `"')")
            )) {
            $count += ([regex]::Matches($js, [regex]::Escape($pair[0]))).Count
            $js = $js.Replace($pair[0], $pair[1])
        }
        if ($count -eq 0) {
            Write-Host "   patterns not found (different SuperBOL version): $bundle" -ForegroundColor Yellow
            continue
        }

        $helper = 'var __sbN=function(p){return typeof p=="string"&&process.platform==="win32"?p.replace(/\//g,"\\").toLowerCase():p};'
        if ($js.StartsWith('"use strict";')) {
            $js = '"use strict";' + $helper + $js.Substring('"use strict";'.Length)
        }
        else {
            $js = $helper + $js
        }
        [IO.File]::WriteAllText($bundle, $js, (New-Object Text.UTF8Encoding $false))
        Write-Host "   $count fix(es) applied to $bundle"
    }
    Write-Host '   (restart VSCodium to load the patched debugger)'
}
Repair-SuperbolDebugger (Join-Path $CodiumDir 'data\extensions')

# ---------------------------------------------------------------------------
# 5. VSCodium user settings and keyboard shortcuts
# ---------------------------------------------------------------------------
Write-Step 'VSCodium settings and shortcuts'

$terminalEnv = [ordered]@{}
foreach ($key in $CobEnv.Keys) { $terminalEnv[$key] = $CobEnv[$key] }
$terminalEnv['PATH'] = ($CobPathAdd -join ';') + ';${env:PATH}'

$userSettings = [ordered]@{
    'superbol.cobcPath'                            = $cobc.FullName
    # Read by the "COBOL: build (SQL)" task (src\example\.vscode\tasks.json).
    'gixsql.gixppPath'                             = $gixpp.FullName
    'gixsql.libDir'                                = $GixSqlLib
    'terminal.integrated.env.windows'              = $terminalEnv
    '[cobol]'                                      = [ordered]@{ 'editor.rulers' = @(6, 7, 11, 72) }
    # No "Restricted Mode" and no "Do you trust the authors?" prompt
    'security.workspace.trust.enabled'             = $false
    # Secondary side bar (on the right) hidden on startup
    'workbench.secondarySideBar.defaultVisibility' = 'hidden'
}
$gdbExe = Get-ChildItem $CobRoot -Recurse -Filter 'gdb.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($gdbExe) { $userSettings['superbol.debugger.gdbPath'] = $gdbExe.FullName }

$userDir = Join-Path $CodiumDir 'data\user-data\User'
$userSettingsFile = Join-Path $userDir 'settings.json'
if (Test-Path $userSettingsFile) {
    Copy-Item $userSettingsFile "$userSettingsFile.bak" -Force
    Write-Host "   previous settings.json saved as settings.json.bak"
}
Save-Json $userSettingsFile $userSettings
Write-Host "   $userSettingsFile"

# One function key per task (only in COBOL files); F5 = debug is built in.
# F9 and F10 are the editor's own debugger keys (toggle breakpoint, step over),
# so those two are bound with "!inDebugMode": while a debug session is running
# they keep their debugger meaning, and they build when it is not.
Save-Utf8NoBom (Join-Path $userDir 'keybindings.json') @'
[
    { "key": "f6",  "command": "workbench.action.tasks.runTask", "args": "COBOL: build",              "when": "editorLangId == cobol" },
    { "key": "f7",  "command": "workbench.action.tasks.runTask", "args": "COBOL: build and run",      "when": "editorLangId == cobol" },
    { "key": "f8",  "command": "workbench.action.tasks.runTask", "args": "COBOL: build (debug)",      "when": "editorLangId == cobol" },
    { "key": "f9",  "command": "workbench.action.tasks.runTask", "args": "COBOL: build (SQL)",        "when": "editorLangId == cobol && !inDebugMode" },
    { "key": "f10", "command": "workbench.action.tasks.runTask", "args": "COBOL: build (module .dll)", "when": "editorLangId == cobol && !inDebugMode" }
]
'@
Write-Host "   $(Join-Path $userDir 'keybindings.json')"

# ---------------------------------------------------------------------------
# 6. Sample project: every example, embedded in src\example\ next to this
#    script; downloaded from -ExampleBaseUrl file by file otherwise - see
#    CLAUDE.md
# ---------------------------------------------------------------------------
Write-Step 'Sample project cobol-example-project (every example)'

# The project was called hello-cobol when hello-world.cob was all it held. An
# existing one is renamed rather than left behind, so a reinstall does not end
# up with two projects and anything edited in the old one is kept. VSCodium
# locks the folder it has open, so the rename is best-effort: the project is
# recreated under its new name below either way, and only the leftover folder
# needs a hand.
$previousProjectDir = Join-Path $InstallDir 'projects\hello-cobol'
if (Test-Path $previousProjectDir) {
    if (Test-Path $ProjectDir) {
        # Both folders exist: the rename was not possible on an earlier run and
        # the project has been recreated under its new name since. Said on every
        # run, so the leftover is noticed rather than forgotten.
        Write-Host "   note: the old hello-cobol project folder is still there ($previousProjectDir) - delete it when you no longer need it." -ForegroundColor Yellow
    }
    else {
        try {
            Move-Item -LiteralPath $previousProjectDir -Destination $ProjectDir -ErrorAction Stop
            Write-Host "   renamed the previous hello-cobol project folder to $(Split-Path $ProjectDir -Leaf)"
        }
        catch {
            Write-Host "   could not rename the previous hello-cobol project folder: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "   close VSCodium and rerun to have it renamed, or delete $previousProjectDir by hand." -ForegroundColor Yellow
        }
    }
}

# src\example\ is mirrored one for one into the project - same relative path
# under src\example\, under -ExampleBaseUrl when downloading, and inside the
# project - so each example keeps its own folder and the copybooks resolve
# there exactly as they do in the sources: the .vscode tasks compile with
# -I"${fileDirname}" (the example's own folder: orders.cpy, exchange.cpy) and
# -I"${fileDirname}\.." (the project root: the shared sqlca.cpy).
# The Access database the access\ examples read and write comes along too, from
# bin\example\access\ - it is a binary, so that is where it lives rather than in
# src\example\ (see CLAUDE.md). The "SampleStore" ODBC DSN (section 7c) is
# pointed at this copy, inside the project the customer has open, which leaves
# the one in bin\example\access\ as the pristine template.
foreach ($rel in @(
        'hello\hello-world.cob',
        'subroutine\exchange.cpy',
        'subroutine\greeter-lib.cob',
        'subroutine\greeter-main.cob',
        'access\orders.cpy',
        'access\insert-order.cob',
        'access\select-orders.cob',
        'mariadb\select-tables.cob',
        'sqlca.cpy',
        '.vscode\settings.json',
        '.vscode\tasks.json',
        '.vscode\launch.json',
        '.vscode\build-sql.ps1',
        '.gitignore'
    )) {
    $url = "$ExampleBaseUrl/$($rel -replace '\\', '/')"
    Get-EmbeddedOrRemoteFile $rel $url (Join-Path $ProjectDir $rel) $ExampleDir
}

$AccessDbRel = 'access\sample-store.mdb'
$AccessDb = Join-Path $ProjectDir $AccessDbRel
Get-EmbeddedOrRemoteFile "example\$AccessDbRel" "$ExampleBinBaseUrl/access/sample-store.mdb" $AccessDb $ProjectBinDir

New-Item -ItemType Directory -Force -Path (Join-Path $ProjectDir 'copybooks') | Out-Null

# Earlier versions of this script put hello-world.cob at the root of a flat
# project; it now lives in hello\ with the other examples. Dropping the old
# copy, and whatever a build of it left next to it (cobc -save-temps writes
# .c/.c.h/.i/.o beside the source), keeps a reinstall from showing the same
# program twice.
$flatHello = @(Get-ChildItem $ProjectDir -Filter 'hello-world.*' -File -ErrorAction SilentlyContinue)
if ($flatHello.Count -gt 0) {
    $flatHello | Remove-Item -Force
    Write-Host "   removed $($flatHello.Count) file(s) of the previous flat hello-world (now in hello\)"
}

# Same story for the editor configuration: it was once written to the project
# root, where VSCodium never reads it. Only the .vscode\ copy counts, so a
# root-level leftover is deleted once its .vscode\ counterpart is in place.
foreach ($stale in @('settings.json', 'tasks.json', 'launch.json')) {
    $rootCopy = Join-Path $ProjectDir $stale
    if ((Test-Path $rootCopy) -and (Test-Path (Join-Path $ProjectDir ".vscode\$stale"))) {
        Remove-Item $rootCopy -Force
        Write-Host "   removed the stale $stale from the project root (kept .vscode\$stale)"
    }
}

# ---------------------------------------------------------------------------
# 7. Compile every example from src\example\<group>\ into bin\example\<group>\
#    so the customer can just run what ships in bin\, with no toolchain
#    knowledge and no build step of their own. Each example is independent:
#    one that fails to build is reported and the rest still get built.
#    greeter-lib is built with -m (a dynamically CALLable .dll, which is what
#    greeter-main's CALL "greeter-lib" loads at run time), everything else
#    with -x.
# ---------------------------------------------------------------------------
Write-Step 'Compiling the examples into bin\example\'

$gccOnPath = Get-Command 'gcc.exe' -ErrorAction SilentlyContinue
if (-not $gccOnPath) {
    Write-Host '   DIAGNOSTIC: gcc.exe does not resolve via PATH in this session.' -ForegroundColor Yellow
    Write-Host "   gcc.exe found on disk at: $($gccExe.FullName)"
    Write-Host "   CobPathAdd: $($CobPathAdd -join ' | ')"
    Write-Host "   env:PATH  : $env:PATH"
}
else {
    Write-Host "   gcc.exe resolves to: $($gccOnPath.Source)"
}

$Examples = [ordered]@{}
foreach ($ex in @(
        @{ Rel = 'hello\hello-world.cob' },
        @{ Rel = 'subroutine\greeter-lib.cob'; Module = $true },
        @{ Rel = 'subroutine\greeter-main.cob' },
        @{ Rel = 'access\insert-order.cob'; Sql = $true },
        @{ Rel = 'access\select-orders.cob'; Sql = $true },
        @{ Rel = 'mariadb\select-tables.cob'; Sql = $true }
    )) {
    $source = Join-Path $ExampleDir $ex.Rel
    if (-not (Test-Path $source)) {
        Write-Host "   skipped (not found): $($ex.Rel)" -ForegroundColor Yellow
        continue
    }
    try {
        $outFile = New-CobolExampleBinary -Source $source -Sql:([bool]$ex.Sql) -Module:([bool]$ex.Module)
        $Examples[$ex.Rel] = $outFile
        if ($ex.Module) {
            # A CALLed module is loaded by its caller, not started on its own.
            Write-Host "   $outFile" -ForegroundColor Green
        }
        else {
            $launcher = New-ExampleLauncher $outFile
            Write-Host "   $outFile (+ $(Split-Path $launcher -Leaf))" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "   FAILED $($ex.Rel): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
# 7b. Smoke-run the two examples that need nothing but the compiler.
#     greeter-main is run from its own folder so the dynamic CALL finds
#     greeter-lib.dll next to it.
# ---------------------------------------------------------------------------
Write-Step 'Test: running the hello-world and subroutine examples'

foreach ($rel in @('hello\hello-world.cob', 'subroutine\greeter-main.cob')) {
    $exe = $Examples[$rel]
    if (-not $exe) { continue }
    Push-Location (Split-Path $exe -Parent)
    try {
        $output = & $exe 2>&1
        if ($LASTEXITCODE -ne 0) { throw "exited with code ${LASTEXITCODE}: $output" }
        Write-Host "   $(Split-Path $exe -Leaf) output: $output" -ForegroundColor Green
    }
    catch {
        Write-Host "   $(Split-Path $exe -Leaf) failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    finally { Pop-Location }
}

# ---------------------------------------------------------------------------
# 7b. End-to-end test: build and run select-tables.cob (GixSQL + MariaDB)
#    Connects to the public "financial" database at relational.fel.cvut.cz
#    (CTU Prague Relational Learning Repository, guest/ctu-relational) and
#    lists its tables - a live smoke test of the GixSQL MySQL/MariaDB client
#    path. Best-effort only: a firewalled network or a moved/changed public
#    server must not abort the rest of the install, so failures here are a
#    warning, not a throw.
#    The upstream gixsql-binaries-windows-x64-mingw-1.0.20b-1 package has a
#    packaging bug (CI copies the x64 build's libgixsql.dll into lib\x86\gcc
#    instead of lib\x64\gcc, leaving x64\gcc with only a static libgixsql.a
#    that is missing its fmt/spdlog object code and cannot be linked). The
#    embedded copy in .\bin has been patched in place: the misplaced x64 DLL
#    was recovered, a matching import library was regenerated with dlltool,
#    and libfmt.dll (its one external runtime dependency; spdlog is linked
#    in statically) was added alongside - see bin\gixsql-binaries-windows-
#    x64-mingw-1.0.20b-1.7z\lib\x64\gcc. A papit.fr-hosted copy still needs
#    to be re-uploaded separately to pick up this fix for the download path.
#    libgixsql-mysql.dll in that same embedded archive is also a rebuild:
#    _mysql_exec_params() never applied connection_opts->fixup_parameters
#    (only prepare() did), so a parameterized statement - this example's
#    "WHERE table_schema = :ws-schema", or any INSERT/UPDATE/DELETE with
#    host variables - reached the server with GixSQL's internal "$n"
#    markers and failed with "Unknown column '$1' in 'WHERE'". Verified
#    against the live server before and after. The ODBC driver had the
#    identical bug plus a hardcoded ColumnSize; both are fixed in 7c.
# ---------------------------------------------------------------------------
$mariaDbExe = $Examples['mariadb\select-tables.cob']
if ($mariaDbExe) {
    Write-Step 'Test: running select-tables.cob (GixSQL + MariaDB)'
    try {
        $output = & $mariaDbExe 2>&1
        if ($LASTEXITCODE -ne 0) { throw "program exited with code ${LASTEXITCODE}: $output" }
        Write-Host "   program output: $output" -ForegroundColor Green
    }
    catch {
        Write-Host "   SQL test skipped: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host '   (needs internet access to relational.fel.cvut.cz)' -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
# 7c. End-to-end test: build and run insert-order.cob + select-orders.cob
#    (GixSQL + ODBC + Access, against the sample-store.mdb that step 6 copied
#    into the project: that one copy is what the DSN names, so the examples run
#    from bin\example\ and the ones opened in the IDE share a single database.)
#    Needs the Access Database Engine (installed only when this script runs
#    elevated - see section 2c) and a "SampleStore" ODBC DSN pointing at the
#    .mdb, which is created here (a User DSN: no elevation needed once the
#    driver itself is installed). Best-effort only, like 7b: skipped with an
#    informational message - not a warning banner, since this is optional and
#    its own section already explains how to enable it - when the driver
#    isn't present this run.
# ---------------------------------------------------------------------------
Write-Step 'Test: running insert-order.cob + select-orders.cob (GixSQL + ODBC + Access)'

$accessDriver = Get-OdbcDriver -Platform '64-bit' -ErrorAction SilentlyContinue |
Where-Object { $_.Name -match 'Microsoft Access Driver' } | Select-Object -First 1

if (-not $accessDriver) {
    Write-Host '   skipped: Access ODBC driver not installed (rerun as Administrator to enable this test).'
}
else {
    $accessExample = $AccessDb
    if (-not (Test-Path $accessExample)) {
        Write-Host "   skipped: $accessExample not found."
    }
    else {
        # The DSN is written straight to the registry rather than with the
        # Add/Set/Remove-OdbcDsn cmdlets: those go through a CIM provider
        # that fails here with "The remote procedure call failed", and
        # Set-OdbcDsn/Remote-OdbcDsn then leave a stale DSN behind. A User
        # DSN *is* just these two HKCU keys, so writing them needs no
        # elevation and is idempotent. Guarded like the builds below - a
        # DSN that cannot be written must not abort the install.
        $dsnName = 'SampleStore'
        $dsnOk = $false
        try {
            $dsnKey = "HKCU:\Software\ODBC\ODBC.INI\$dsnName"
            $dsnIndex = 'HKCU:\Software\ODBC\ODBC.INI\ODBC Data Sources'
            $previousDbq = (Get-ItemProperty $dsnKey -Name DBQ -ErrorAction SilentlyContinue).DBQ

            New-Item -Path $dsnKey -Force | Out-Null
            New-Item -Path $dsnIndex -Force | Out-Null
            Set-ItemProperty $dsnKey -Name 'Driver' -Value $accessDriver.Attribute['Driver']
            Set-ItemProperty $dsnKey -Name 'DBQ' -Value $accessExample
            Set-ItemProperty $dsnKey -Name 'DriverId' -Value 25 -Type DWord
            Set-ItemProperty $dsnKey -Name 'SafeTransactions' -Value 0 -Type DWord
            Set-ItemProperty $dsnKey -Name 'UID' -Value ''
            Set-ItemProperty $dsnIndex -Name $dsnName -Value $accessDriver.Name

            if ($previousDbq -eq $accessExample) {
                Write-Host "   ODBC DSN '$dsnName' already points at $accessExample"
            }
            else {
                Write-Host "   ODBC DSN '$dsnName' -> $accessExample"
            }
            $dsnOk = $true
        }
        catch {
            Write-Host "   Access test skipped: could not configure the '$dsnName' ODBC DSN: $($_.Exception.Message)" -ForegroundColor Yellow
        }

        # insert-order first, then select-orders: the row the first one adds
        # shows up at the bottom of the second one's listing.
        foreach ($rel in @('access\insert-order.cob', 'access\select-orders.cob')) {
            if (-not $dsnOk) { break }
            $exe2 = $Examples[$rel]
            if (-not $exe2) { continue }
            $name2 = [IO.Path]::GetFileNameWithoutExtension($exe2)
            try {
                $output2 = & $exe2 2>&1
                if ($LASTEXITCODE -ne 0) { throw "exited with code ${LASTEXITCODE}: $output2" }
                Write-Host "   $name2 output: $output2" -ForegroundColor Green
            }
            catch {
                Write-Host "   $name2 test skipped: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}

# ---------------------------------------------------------------------------
# 8. Launcher (and optionally a persistent user environment)
# ---------------------------------------------------------------------------
Write-Step 'Launcher CoBOL-Codium.cmd'

$cmd = New-Object Text.StringBuilder
[void]$cmd.AppendLine('@echo off')
[void]$cmd.AppendLine('rem Generated by install-cobol-on-win.ps1: GnuCOBOL environment, then VSCodium')
foreach ($key in $CobEnv.Keys) {
    [void]$cmd.AppendLine("set `"$key=$(ConvertTo-CmdValue $CobEnv[$key])`"")
}
[void]$cmd.AppendLine("set `"PATH=$(ConvertTo-CmdValue ($CobPathAdd -join ';'));%PATH%`"")
[void]$cmd.AppendLine('start "" "%~dp0VSCodium\VSCodium.exe" "%~dp0projects\cobol-example-project" "%~dp0projects\cobol-example-project\hello\hello-world.cob"')

# cmd.exe reads .cmd files in the OEM code page: write them in that encoding.
$oem = [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
$launcher = Join-Path $InstallDir 'CoBOL-Codium.cmd'
[IO.File]::WriteAllText($launcher, $cmd.ToString(), $oem)
Write-Host "   $launcher"

# Desktop shortcut
try {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'COBOL (VSCodium).lnk'))
    $shortcut.TargetPath = $launcher
    $shortcut.WorkingDirectory = $InstallDir
    $shortcut.IconLocation = "$codiumExe,0"
    $shortcut.WindowStyle = 7   # minimized cmd window
    $shortcut.Save()
    Write-Host '   "COBOL (VSCodium)" shortcut created on the Desktop.'
}
catch { Write-Host "   shortcut not created: $($_.Exception.Message)" -ForegroundColor Yellow }

if ($PersistUserEnv) {
    Write-Step 'Storing the variables in the user environment'
    foreach ($key in $CobEnv.Keys) {
        [Environment]::SetEnvironmentVariable($key, $CobEnv[$key], 'User')
    }
    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    $entries = @($userPath -split ';' | Where-Object { $_ })
    $toAdd = @($CobPathAdd | Where-Object { $entries -notcontains $_ })
    if ($toAdd.Count -gt 0) {
        [Environment]::SetEnvironmentVariable('Path', (($toAdd + $entries) -join ';'), 'User')
    }
    Write-Host '   done: open a new terminal to use it.'
}

# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Installation complete.' -ForegroundColor Green
Write-Host "  Start the IDE         : Type CoBOL in the Start Menu or $launcher (or the Desktop shortcut)"
Write-Host '  Run                   : Ctrl + F5   or  Run > Start Without Debugging'
Write-Host '  Debug                 : F5          or  Run > Start Debugging'
Write-Host '  Build                 : F6          or  Terminal > Run Build Task... (Ctrl+Shift+B)'
Write-Host '  Build and run         : F7          or  Terminal > Run Task... > COBOL: build and run'
Write-Host '  Build for debugging   : F8          or  Terminal > Run Task... > COBOL: build (debug)'
Write-Host '  Build with EXEC SQL   : F9          or  Terminal > Run Task... > COBOL: build (SQL)'
Write-Host '  Build a CALLable .dll : F10         or  Terminal > Run Task... > COBOL: build (module .dll)'
Write-Host '                          (F9/F10 build outside a debug session; during one they stay'
Write-Host "                           the debugger's toggle breakpoint / step over)"
Write-Host "  Examples in the IDE   : $ProjectDir"
Write-Host '                          hello\ (DISPLAY), subroutine\ (CALL a COBOL library),'
Write-Host '                          access\ (EXEC SQL + ODBC on sample-store.mdb), mariadb\ (EXEC SQL)'
Write-Host "  SQL preprocessor      : GixSQL $gixppVer"
Write-Host "                          PostgreSQL (libpq $pgVer), MySQL/MariaDB (libmariadb $mariaVer),"
Write-Host "                          SQLite (bundled$(if ($sqliteVer) { ", $sqliteVer" })), Oracle (needs OCI, not bundled), ODBC"
Write-Host "                          DB2 via ODBC, using the driver in $Db2DriverDir$(if ($db2Ver) { " ($db2Ver)" })"
if (Test-IsAdmin) {
    Write-Host "                          Excel/Access via ODBC (installed - ran elevated, Access Database Engine $aceVer)" -ForegroundColor Green
}
else {
    Write-Host '                          Excel/Access NOT installed - rerun as Administrator for it' -ForegroundColor Yellow
}
