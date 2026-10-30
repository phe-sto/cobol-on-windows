# Context

An easy way to develop CoBOL on Windows platform. Get out of old proprietary environments to a
one with modern IDE with syntax highlighting, Debugger, examples.

Business aim is to sell my freelance expertise in development.

# Git

I, the human developer is the only one to commit.

# Requirements

* No clicks, just a script our our installer.
* Compatible with relational databases, DB2, MariaDB, *etc*.
* Add all versions of the connector to the comment of `.\src\install-cobol-on-win.ps1` and CLI output.
* No admin rights required, except for few options:
    * Excel and Access ODBC connectors

# IDE

Codium for being Open Source and vendor free.

# Code and binary sources

Are to be embedded in the project. The project is hosted in the GitHub repo
https://github.com/phe-sto/cobol-on-windows.
It all serve as a fallback if only `.\src\install-cobol-on-win.ps1` is downloaded and ran.
Binaries in `.\bin\` are Git LFS objects, so the fallback URLs use
`https://media.githubusercontent.com/media/phe-sto/cobol-on-windows/main/...` for those
(`raw.githubusercontent.com` returns the LFS pointer instead of the file) and
`https://raw.githubusercontent.com/phe-sto/cobol-on-windows/main/...` for everything else.
Two folders, `.\src\` for sources and bin for binaries.
Cobol example files are in `.\src\example\`

# PowerShell

The MS ones from https://learn.microsoft.com/en-us/powershell/scripting/community/contributing/powershell-style-guide?view=powershell-7.6 site.+

# CoBOL

Keywords and instructions are upper case (`PROCEDURE DIVISION.`, `DISPLAY`, *etc*) our declarations of variables, copies, paragraphs are kebab-case. 

# SQL

## Convention

Keywords and instructions are upper case (`SELECT`, `WHERE`, *etc*) our declarations of variables, are snake_case. 

## SQL Preprocessor

Gix SQL with ODBC. The one from OCaml PRO still actively maintained https://github.com/OCamlPro/gixsql.

## SQL Client

ODBC, so compatible with everything...

## SQLCA

Source is a copy of https://raw.githubusercontent.com/mridoni/gixsql/refs/heads/main/copy/SQLCA.cpy/

# File naming

All files in `bin\example`, `src\example` and `\.` must be lower case hyphen separated (kebab-case).

# Test steps:

- Compile cobol source `.\src\example\hello\hello-world.cob`

- CoBOL subroutine, a CoBOL program calling a CoBOL library, via copy exchange

- A CoBOL program insert an order into `.\bin\example\access\sample-store.mdb`.
It shall use the `.\src\example\access\orders.cpy` copy.
The request is:

```
INSERT INTO orders (customer_id, order_date, shipped_date, status)
VALUES (4, DATE(), NULL, 'Pending');
```

- A CoBOL program outputting all rows from orders table in .\bin\example\access\sample-store.mdb.
It shall use the `.\src\example\access\orders.cpy` copy.
The request is:
```
SELECT *
FROM orders
ORDER BY order_date
```

# Monorepo with the website

`.\wwww\` contains the GitHub website, the marketing entry point of the project.

# License

See WTFPL site https://www.wtfpl.net/.