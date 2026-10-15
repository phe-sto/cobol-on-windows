      *> Main program: calls the greeter-lib subroutine, exchanging data
      *> through the shared exchange.cpy copy.
       IDENTIFICATION DIVISION.
       PROGRAM-ID. greeter-main.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       COPY "exchange.cpy".
       PROCEDURE DIVISION.
           MOVE "World" TO ws-name
           CALL "greeter-lib" USING ws-greeting-rec
           DISPLAY ws-message
           STOP RUN.
