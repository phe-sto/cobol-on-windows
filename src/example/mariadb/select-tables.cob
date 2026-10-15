      *> Lists every table of the public "financial" MariaDB database
      *> hosted at relational.fel.cvut.cz, via GixSQL's native MySQL/
      *> MariaDB driver (no ODBC DSN needed).
      *>   SELECT GROUP_CONCAT(TABLE_NAME ORDER BY TABLE_NAME
      *>          SEPARATOR ', ') FROM information_schema.tables
      *>   WHERE table_schema = :ws-schema
      *> Uses a singleton SELECT INTO rather than a cursor: a confirmed
      *> GixSQL bug (the generated lazy CURSOR DECLARE paragraph is
      *> missing its terminating period) makes GnuCOBOL reject any
      *> EXEC SQL DECLARE CURSOR in this GixSQL 1.0.20b/1.1.19a build.
      *> The schema name is passed as a bound host variable rather than a
      *> literal, so this also covers the parameterized exec_params path
      *> that two other confirmed GixSQL bugs used to break (see the
      *> comments in .\src\install-cobol-on-win.ps1, step 7b). It needs
      *> gixpp's "-P varchar": as "char" (the default) the PIC X(20) is
      *> sent space-padded and matches no schema.
       IDENTIFICATION DIVISION.
       PROGRAM-ID. mariadb-select.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01  ws-table-list          PIC X(500).
       01  ws-schema              PIC X(20) VALUE "financial".
       01  ws-datasrc             PIC X(50)
           VALUE "mysql://relational.fel.cvut.cz:3306/financial".
       01  ws-user                PIC X(30)
           VALUE "guest.ctu-relational".
       COPY SQLCA.

       PROCEDURE DIVISION.
       main-para.
           EXEC SQL CONNECT TO :ws-datasrc USER :ws-user END-EXEC
           IF SQLCODE NOT = 0
               DISPLAY "Connect failed, SQLCODE=" SQLCODE
               STOP RUN
           END-IF

           EXEC SQL
               SELECT GROUP_CONCAT(TABLE_NAME ORDER BY TABLE_NAME
                                    SEPARATOR ', ')
               INTO :ws-table-list
               FROM information_schema.tables
               WHERE table_schema = :ws-schema
           END-EXEC
           IF SQLCODE NOT = 0
               DISPLAY "Select failed, SQLCODE=" SQLCODE
           ELSE
               DISPLAY "Tables in financial: " ws-table-list
           END-IF

           EXEC SQL DISCONNECT END-EXEC
           STOP RUN.
