      *> Outputs every row of the orders table of sample-store.mdb (shipped in
      *> .\bin\example\access\, copied next to this program in the installed
      *> project, reached through the "SampleStore" ODBC DSN) via GixSQL/ODBC,
      *> using the shared .\src\example\access\orders.cpy copy.
      *>   SELECT * FROM orders ORDER BY order_date
      *> Works around a confirmed GixSQL bug (createCobolData() writes an
      *> unconditional NULL-indicator value even for host variables that
      *> have none, corrupting unrelated COBOL literals) by giving every
      *> fetched column its own explicit null indicator.
       IDENTIFICATION DIVISION.
       PROGRAM-ID. select-orders.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       EXEC SQL INCLUDE orders END-EXEC.
       01  ws-order-id-null       PIC S9(4) COMP-5.
       01  ws-customer-id-null    PIC S9(4) COMP-5.
       01  ws-order-date-null     PIC S9(4) COMP-5.
       01  ws-status-null         PIC S9(4) COMP-5.
       01  ws-datasrc             PIC X(40) VALUE "odbc://SampleStore".
       01  ws-user                PIC X(10) VALUE "x.x".
       01  ws-shipped-display     PIC X(19).
       01  ws-row-count           PIC 9(9) VALUE 0.
       EXEC SQL INCLUDE SQLCA END-EXEC.

       PROCEDURE DIVISION.
       main-para.
           EXEC SQL DECLARE ordcur CURSOR FOR
               SELECT * FROM orders ORDER BY order_date
           END-EXEC

           EXEC SQL CONNECT TO :ws-datasrc USER :ws-user END-EXEC
           IF SQLCODE NOT = 0
               DISPLAY "Connect failed, SQLCODE=" SQLCODE
               STOP RUN
           END-IF

           EXEC SQL OPEN ordcur END-EXEC
           IF SQLCODE NOT = 0
               DISPLAY "Open failed, SQLCODE=" SQLCODE
               STOP RUN
           END-IF

           DISPLAY "order_id   cust_id    order_date           "
               "shipped_date         status"

           PERFORM fetch-next-row UNTIL SQLCODE NOT = 0

           DISPLAY ws-row-count " row(s)."
           EXEC SQL CLOSE ordcur END-EXEC
           EXEC SQL DISCONNECT END-EXEC
           STOP RUN.

       fetch-next-row.
           EXEC SQL
               FETCH ordcur
               INTO :ws-order-id:ws-order-id-null,
                    :ws-customer-id:ws-customer-id-null,
                    :ws-order-date:ws-order-date-null,
                    :ws-shipped-date:ws-shipped-date-null,
                    :ws-status:ws-status-null
           END-EXEC
           IF SQLCODE = 0
               ADD 1 TO ws-row-count
               IF ws-shipped-date-null < 0
                   MOVE SPACES TO ws-shipped-display
               ELSE
                   MOVE ws-shipped-date TO ws-shipped-display
               END-IF
               DISPLAY ws-order-id "  " ws-customer-id "  "
                       ws-order-date "  " ws-shipped-display "  "
                       ws-status
           END-IF.

       end-of-program.
           CONTINUE.
