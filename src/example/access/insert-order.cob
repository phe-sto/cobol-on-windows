      *> Inserts one order into the orders table of sample-store.mdb (shipped
      *> in .\bin\example\access\, copied next to this program in the installed
      *> project, reached through the "SampleStore" ODBC DSN) via GixSQL/ODBC,
      *> using the shared .\src\example\access\orders.cpy copy.
      *>   INSERT INTO orders (customer_id, order_date, shipped_date,
      *>                       status)
      *>   VALUES (4, DATE(), NULL, 'Pending');
       IDENTIFICATION DIVISION.
       PROGRAM-ID. insert-order.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       EXEC SQL INCLUDE orders END-EXEC.
       01  ws-datasrc             PIC X(40) VALUE "odbc://SampleStore".
       01  ws-user                PIC X(10) VALUE "x.x".
       EXEC SQL INCLUDE SQLCA END-EXEC.
       PROCEDURE DIVISION.
           MOVE 4         TO ws-customer-id
           MOVE "Pending" TO ws-status

           EXEC SQL CONNECT TO :ws-datasrc USER :ws-user END-EXEC
           IF SQLCODE NOT = 0
               DISPLAY "Connect failed, SQLCODE=" SQLCODE
               STOP RUN
           END-IF

           EXEC SQL
               INSERT INTO orders
                   (customer_id, order_date, shipped_date, status)
               VALUES
                   (:ws-customer-id, DATE(), NULL, :ws-status)
           END-EXEC

           IF SQLCODE NOT = 0
               DISPLAY "Insert failed, SQLCODE=" SQLCODE
           ELSE
               EXEC SQL COMMIT END-EXEC
               DISPLAY "Order inserted for customer_id " ws-customer-id
           END-IF

           EXEC SQL DISCONNECT END-EXEC
           STOP RUN.
