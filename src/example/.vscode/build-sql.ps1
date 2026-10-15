#Requires -Version 5.1
<#
.SYNOPSIS
    Builds one COBOL program that contains EXEC SQL statements: gixpp, then cobc.

.DESCRIPTION
    Called by the "COBOL: build (SQL)" task (.vscode\tasks.json). It exists
    because the task needs to do one thing a plain command line cannot: patch
    what gixpp emits before cobc reads it.

    Confirmed GixSQL 1.0.20b bug: in the lazy "ESQL CURSOR DECLARATIONS"
    trailer gixpp writes for EXEC SQL DECLARE CURSOR, the GIXSQLCursorDeclare
    CALL's closing END-CALL is missing its terminating period, so GnuCOBOL 3.2
    reads the next line (a paragraph name) as part of the same unfinished
    sentence and every cursor-based program fails to compile - access\
    select-orders.cob, for one. The same narrow, idempotent patch lives in
    install-cobol-on-win.ps1 (Repair-GixsqlCursorTrailer), which is what
    compiles the examples into bin\example\: keep the two in step.

    The .cbsql gixpp produces is left next to the source on purpose - it is the
    COBOL that cobc actually compiles, so it is worth reading when an EXEC SQL
    statement misbehaves. .gitignore keeps it out of the repository.

.PARAMETER Source
    COBOL source to build. The task passes ${file}.

.PARAMETER GixppPath
    gixpp.exe, GixSQL's EXEC SQL preprocessor. The task passes the
    gixsql.gixppPath setting written by install-cobol-on-win.ps1.

.PARAMETER CobcPath
    cobc.exe, the GnuCOBOL compiler. The task passes the superbol.cobcPath
    setting.

.PARAMETER GixSqlLibDir
    Folder holding libgixsql and its drivers, linked with -lgixsql. The task
    passes the gixsql.libDir setting.

.EXAMPLE
    .\build-sql.ps1 -Source .\access\select-orders.cob -GixppPath c:\cobol-dev\GixSQL\gixsql-binaries-windows-x64-mingw\bin\gixpp.exe -CobcPath c:\cobol-dev\GnuCOBOL\bin\cobc.exe -GixSqlLibDir c:\cobol-dev\GixSQL\gixsql-binaries-windows-x64-mingw\lib\x64\gcc
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Source,
    [Parameter(Mandatory)][string]$GixppPath,
    [Parameter(Mandatory)][string]$CobcPath,
    [Parameter(Mandatory)][string]$GixSqlLibDir
)

$sourceDir = Split-Path $Source -Parent
$name = [IO.Path]::GetFileNameWithoutExtension($Source)
$cbsql = Join-Path $sourceDir "$name.cbsql"
$exe = Join-Path $sourceDir "$name.exe"

# Both -I paths matter, and they are the ones the installer uses too: the
# program's own folder (orders.cpy, exchange.cpy) and the project root (the
# shared sqlca.cpy). gixpp resolves EXEC SQL INCLUDE, cobc resolves plain COPY.
$includeDirs = @('-I', $sourceDir, '-I', (Join-Path $sourceDir '..'))

& $GixppPath -e -S -P varchar @includeDirs -i $Source -o $cbsql
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$text = [IO.File]::ReadAllText($cbsql)
$pattern = '(?m)^(?<call>GIXSQL[ \t]*END-CALL)[ \t]*(?<nl>\r?\n)(?=GIXSQL[ \t]+[\w-]+\.[ \t]*(\r?\n|$))'
$fixed = [regex]::Replace($text, $pattern, '${call}.${nl}')
if ($fixed -ne $text) { [IO.File]::WriteAllText($cbsql, $fixed) }

& $CobcPath -x @includeDirs -L $GixSqlLibDir -lgixsql -o $exe $cbsql
exit $LASTEXITCODE
