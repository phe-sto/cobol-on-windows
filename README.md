# COBOL on Windows

A COBOL development environment that installs itself on a plain Windows PC or Windows Server: one PowerShell
script, no clicks, no administrator rights, no proprietary licence to buy. Modern editor,
working debugger, SQL access to real databases, and examples that compile and run **preconfigured** at the end
of the install.

```powershell
powershell -ExecutionPolicy Bypass -File .\src\install-cobol-on-win.ps1
```

Project home: <https://github.com/phe-sto/cobol-on-windows>

## Why

COBOL still runs payroll, banking and insurance, but the tooling around it is often the part
nobody wants to touch: an editor from another decade, a compiler you can only use on a machine
somebody else owns, a licence renewal per seat, and a debugger that means reading a dump. New
developers bounce off that long before they bounce off the language, which is a shame, because
COBOL itself is not the hard part.

Everything needed to do the job well already exists as open source. It is simply scattered: a
compiler here, an editor extension there, an EXEC SQL preprocessor somewhere else, each with
its own install dance, its own PATH variables and its own Windows quirks. Wiring it together is
a day of work the first time, and a day you repeat on every new machine.

This project is that day, done once, as a script:

- **A real IDE.** VSCodium with the SuperBOL extension: syntax highlighting, navigation,
  completion, diagnostics, and breakpoints that actually stop in your COBOL source — on
  Windows, where that normally does not work (the script patches the debugger's path handling
  for you). The **compiler** and **debugger** already configured in the IDE.
- **Real databases.** EXEC SQL through GixSQL, over ODBC, so DB2, MariaDB/MySQL, PostgreSQL,
  SQLite, SQL Server, Access and Excel are all reachable from COBOL. The DB2 driver is included.
- **Examples that run.** A DISPLAY program, a program calling a COBOL subroutine through a
  shared copybook, an INSERT and a SELECT against an Access database over ODBC, and a query
  against a public MariaDB server. The install compiles all of them and runs them in front of
  you, so you know the toolchain works before you write a line.
- **Offline-capable and self-contained.** The sources and every binary are embedded in the
  project (`src\`, `bin\`). Downloaded the script alone? It fetches the same files from this
  repository. Nothing is pulled from a registry that may change under you.
- **No clicks.** One command. No installer wizard, no "next, next, finish", nothing to choose.
- **No administrator rights.** Everything lands in a single folder (`c:\cobol-dev` by default)
  and can be deleted by deleting it. The one exception is the optional Excel/Access ODBC
  connector, which is a real system-wide Microsoft installer and is only run when you choose to
  start the script elevated.

Built and maintained by Christophe Brun (PapIT) — <christophe.brun@papit.fr>. If your team is
sitting on COBOL and wondering how to modernise it without a rewrite, that is exactly the work
I do freelance.

## Install

### Requirements

- Windows 10 or 11 or Server, 64-bit
- Windows PowerShell 5.1 (shipped with Windows) or newer
- About 2 GB of free disk space
- No administrator rights, except for the optional Excel/Access connector

### Run it

From a clone or a copy of this project — uses the embedded binaries, no network needed:

```powershell
powershell -ExecutionPolicy Bypass -File .\src\install-cobol-on-win.ps1
```

Or without cloning anything at all — two commands, the script then fetches what it needs from
this repository, file by file:

```powershell
Invoke-WebRequest -UseBasicParsing -OutFile install-cobol-on-win.ps1 https://raw.githubusercontent.com/phe-sto/cobol-on-windows/main/src/install-cobol-on-win.ps1
powershell -ExecutionPolicy Bypass -File .\install-cobol-on-win.ps1
```

Useful options:

| Option                                                                                                                              | What it does                                                                                                                   |
| ----------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `-InstallDir d:\cobol`                                                                                                              | Installs somewhere else (default `c:\cobol-dev`). Prefer a short path without spaces or accents: gdb and GnuCOBOL are happier. |
| `-PersistUserEnv`                                                                                                                   | Also stores the `COB_*` variables and PATH entries in your Windows user environment, so `cobc` works in any terminal.          |
| `-Force`                                                                                                                            | Downloads and extracts everything again, and refreshes the example project. Close VSCodium first.                              |
| `-VSCodiumUrl`, `-GnuCobolUrl`, `-GixSqlUrl`, `-Db2CliUrl`, `-AccessEngineUrl`, `-SevenZipUrl`, `-ExampleBaseUrl`, `-ProjectBinDir` | Point any component at a different archive, mirror or local folder.                                                            |

Run the same command from an **elevated** PowerShell to also install the Microsoft Access
Database Engine, which is what gives ODBC access to `.accdb`/`.mdb` and `.xlsx` files. Without
elevation that one step is skipped with a warning and everything else still installs.

### What you end up with

```
c:\cobol-dev\
  CoBOL-Codium.cmd            sets the environment, then opens the IDE  (+ Desktop shortcut)
  VSCodium\                   the editor, in portable mode
  GnuCOBOL\                   cobc, GCC, gdb
  GixSQL\                     the EXEC SQL preprocessor and its runtime drivers
  DB2CLI\                     IBM's DB2 ODBC/CLI driver
  projects\cobol-example-project\
      hello\                  hello-world.cob
      subroutine\             greeter-main.cob + greeter-lib.cob + exchange.cpy
      access\                 insert-order.cob + select-orders.cob + orders.cpy  (EXEC SQL, ODBC)
                              + sample-store.mdb, the Access database they read and write
      mariadb\                select-tables.cob                                  (EXEC SQL)
      sqlca.cpy  .vscode\     shared copybook, tasks, keybindings, SQL build helper
```

Start from `CoBOL-Codium.cmd` (or the Desktop shortcut) rather than the Start Menu: it sets the
GnuCOBOL and GixSQL environment before launching the editor, which is what lets both the
integrated terminal and the debugger find the runtime DLLs.

Keys, inside a COBOL file:

| Key       | Action                                                  |
| --------- | ------------------------------------------------------- |
| Ctrl + F5 | Run without debugging                                   |
| F5        | Start debugging (gdb, breakpoints in your COBOL source) |
| F6        | Build                                                   |
| F7        | Build and run                                           |
| F8        | Build for debugging                                     |
| F9        | Build with EXEC SQL (GixSQL, then cobc)                 |
| F10       | Build a CALLable `.dll`                                 |

F9 and F10 build only outside a debug session; while one is running they stay the debugger's
own toggle breakpoint and step over.

The install also compiles every example into `bin\example\`, each with a `.cmd` launcher that
sets the runtime environment for you — so they are ready to run, from Explorer or a plain
terminal, without knowing anything about the toolchain. Those binaries are build output and are
not committed here: only `bin\example\access\sample-store.mdb`, the Access database the examples
query, is, since the script downloads it when it runs standalone.

## Licences

This project — the script, the examples, the copybooks — is released under the
[WTFPL](https://www.wtfpl.net/), version 2. See [LICENSE](LICENSE). Do what the fuck you want
to.

What it installs is not all under one licence, so here is the full list, with the version
embedded in `bin\` at the time of writing. The console output at install time always prints
what actually got installed.

### Open source

| Component                                                                                                                                  | Version     | Licence                                                                                                                                |
| ------------------------------------------------------------------------------------------------------------------------------------------ | ----------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| [VSCodium](https://github.com/VSCodium/vscodium) — the editor, Microsoft-free build of VS Code                                             | 1.135.06055 | MIT (third-party notices in `VSCodium\LICENSES.chromium.html`)                                                                         |
| [SuperBOL Studio OSS](https://github.com/OCamlPro/superbol-studio-oss) — COBOL extension and LSP                                           | 1.0.0       | MIT for the VS Code extension, AGPL-3.0 for the LSP server, ISC for vendored parts                                                     |
| [GnuCOBOL](https://gnucobol.sourceforge.io/) — the compiler                                                                                | 3.2.0       | GPL-3.0 for `cobc` (`COPYING.txt`), LGPL-3.0 for the `libcob` runtime (`COPYING_LESSER.txt`), GFDL for the manuals (`COPYING_DOC.txt`) |
| GCC, bundled with the GnuCOBOL build (GnuCOBOL compiles COBOL to C)                                                                        | 13.1.0      | GPL-3.0-or-later, with the GCC Runtime Library Exception                                                                               |
| GDB, bundled with the GnuCOBOL build                                                                                                       | 13.2        | GPL-3.0-or-later (embeds CPython 3.10, PSF licence)                                                                                    |
| [Berkeley DB](https://www.oracle.com/database/technologies/related/berkeleydb.html) — indexed (ISAM) file support in this GnuCOBOL build   | 18.1        | AGPL-3.0 (Oracle). Worth knowing if you ship indexed-file programs; GnuCOBOL can also be built without it.                             |
| [GixSQL](https://github.com/mridoni/gixsql) — EXEC SQL preprocessor, now also maintained by [OCamlPro](https://github.com/OCamlPro/gixsql) | 1.0.20b-1   | GPL-3.0 as shipped (`doc\LICENSE` in the archive) — check upstream before linking it into something you distribute                     |
| libpq — PostgreSQL client, via GixSQL                                                                                                      | 15.1        | PostgreSQL Licence (BSD-style)                                                                                                         |
| libmariadb — MariaDB/MySQL client, via GixSQL                                                                                              | 3.1.13      | LGPL-2.1                                                                                                                               |
| [SQLite](https://www.sqlite.org/) — statically linked into the GixSQL SQLite driver                                                        | ~3.38.0     | Public domain                                                                                                                          |
| [7-Zip](https://www.7-zip.org/) (`7zr.exe`, used to unpack the GnuCOBOL and GixSQL archives)                                               | 26.04       | LGPL-2.1-or-later                                                                                                                      |

### Not open source, included for connectivity

| Component                                                                                                             | Version                 | Terms                                                                                                                              |
| --------------------------------------------------------------------------------------------------------------------- | ----------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| IBM DB2 ODBC/CLI driver ("DS Driver")                                                                                 | printed at install time | IBM proprietary. Licence and redistribution terms ship with it: `DB2CLI\clidriver\license\` (`odbc_LI_en.rtf`, `odbc_REDIST.txt`). |
| Microsoft Access Database Engine redistributable — optional, admin only, gives ODBC access to `.accdb`/`.mdb`/`.xlsx` | 16.0.5044.1000          | Microsoft proprietary (its own EULA, accepted when the installer runs).                                                            |

Those two are committed in `bin\` (Git LFS), which is also what the download fallback serves, so
the install works without hunting for a download page. They remain the property of IBM and
Microsoft and are used under
their own terms — if you fork or re-host this project, read `odbc_REDIST.txt` and Microsoft's
redistribution terms first. Oracle's Instant Client (OCI), for GixSQL's Oracle driver, is *not*
included for exactly that reason.

The `mariadb\select-tables.cob` example queries the public `financial` database of the
[CTU Prague Relational Learning Repository](https://relational.fel.cvut.cz/) (guest access,
courtesy of the Czech Technical University) — thank you for keeping it open.

## Thanks

This project is a script. The compiler, the editor, the debugger and the SQL preprocessor it
installs are decades of other people's work, and all of it is given away for free. Thank you to
everyone who has kept COBOL alive and open, everyone still maintaining COBOL in production: the language is not the legacy, the
tooling was but hat part is fixable.
