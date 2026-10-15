      *> Host-variable record for the orders table of sample-store.mdb, the
      *> Access database shipped in .\bin\example\access\ (customer_id,
      *> order_date, shipped_date, status), shared by the INSERT/SELECT
      *> examples.
       01  ws-order.
           05  ws-order-id           PIC 9(9).
           05  ws-customer-id        PIC 9(9).
           05  ws-order-date         PIC X(19).
           05  ws-shipped-date       PIC X(19).
           05  ws-shipped-date-null  PIC S9(4) COMP-5.
           05  ws-status             PIC X(20).
