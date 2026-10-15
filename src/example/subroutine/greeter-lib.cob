      *> Subroutine ("library"): builds a greeting from the name it is given.
       IDENTIFICATION DIVISION.
       PROGRAM-ID. greeter-lib.
       DATA DIVISION.
       LINKAGE SECTION.
       COPY "exchange.cpy".
       PROCEDURE DIVISION USING ws-greeting-rec.
           MOVE SPACES TO ws-message
           STRING "Hello, " DELIMITED BY SIZE
                  FUNCTION TRIM(ws-name) DELIMITED BY SIZE
                  "!" DELIMITED BY SIZE
             INTO ws-message
           GOBACK.
