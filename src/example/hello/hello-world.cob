      *> Example COBOL program for GnuCOBOL on Windows.
       IDENTIFICATION DIVISION.
       PROGRAM-ID. hello-world.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01  ws-message   PIC X(40)
           VALUE "Hello from GnuCOBOL on Windows!".
       PROCEDURE DIVISION.
           DISPLAY ws-message
           STOP RUN.
