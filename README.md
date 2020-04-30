# utl-using-proc-odstext-to-add-documentation-tabs-to-your-excel-workbook
Using proc odstext to add documentation tabs to your excel workbook
    SAS Forum: Using proc odstext to add documentation tabs to your excel workbook

    github
    https://tinyurl.com/y9xr5dbv
    https://github.com/rogerjdeangelis/utl-using-proc-odstext-to-add-documentation-tabs-to-your-excel-workbook

    SAS Forum
    https://tinyurl.com/yd49zdw7
    https://communities.sas.com/t5/ODS-and-Base-Reporting/How-to-Get-PROC-ODS-TEXT-on-a-new-Tab-in-ODS-EXCEL-output/m-p/644309

    Reeza
    https://communities.sas.com/t5/user/viewprofilepage/user-id/13879

    *_                   _
    (_)_ __  _ __  _   _| |_
    | | '_ \| '_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    ;

    Three inputs

       sashelp.class  -> Tab [CLASS]
       text for tab   -> Tab [AGENDA]
       havSexAge datasets (sashelp.class(kee=sex age)) -> Tab [DEXAGE]

    SASHELP.CLASS total obs=19

    Obs    NAME       SEX    AGE    HEIGHT    WEIGHT

      Joyce       F      11     51.3       50.5
      Louise      F      12     56.3       77.0
      Alice       F      13     56.5       84.0
      James       M      12     57.3       83.0
    ....


    .p ANALYSIS OF SASHELP.CLASS


    data havSexAge;
      set sashelp.class(keep=sex age obs=5);
    run;quit;

    WORK.HAVSEXAGE total obs=5

     SEX    AGE

      F      11
      F      12
      F      13
      M      12
      M      11


    *            _               _
      ___  _   _| |_ _ __  _   _| |_ ___
     / _ \| | | | __| '_ \| | | | __/ __|
    | (_) | |_| | |_| |_) | |_| | |_\__ \
     \___/ \__,_|\__| .__/ \__,_|\__|___/
                    |_|
    ;

    One workbook with these three tabs

    d:/xls/want.xlsx

          +--------------------------------------------------------------+
          |     A      |    B       |     C      |    D       |    E     |
          +--------------------------------------------------------------+
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT  |
          +------------+------------+------------+------------+----------+
       2  | Alice      |    f      |    12       |   88       |   56     |
          +------------+------------+------------+------------+----------+
       3  | Mary       |    Mf      |    16      |    11      |   67     |
          +------------+------------+------------+------------+----------+
       4  | Tom        |    M       |    15      |    99      |   33     |
          +------------+------------+------------+------------+----------+
       5  | John       |    M       |    14      |    77      |   90     |
          +------------+------------+------------+------------+----------+
       6  | Jane       |    F       |    13      |    44      |   67     |
          +------------+------------+------------+------------+----------+

         [CLASS]


          +------------------------------------+
          |                A                   |
          +------------------------------------+
       1  |      ANALYSIS OF SASHELP.CLASS     |
          +------------------------------------+

         [AGENDA]


          +-------------------------+
          |     A      |    B       |
          +-------------------------+
       1  | NAME       |   SEX      |
          +------------+------------+
       2  | Alice      |    f       |
          +------------+------------+
       3  | Mary       |    Mf      |
          +------------+------------+
       4  | Tom        |    M       |
          +------------+------------+
       5  | John       |    M       |
          +------------+------------+
       6  | Jane       |    F       |
          +------------+------------+

         [SEXAGE]


    run;quit;

    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| '_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|

    ;


    ods excel file="d:/xls/want.xlsx"
       options(sheet_name="CLASS" sheet_interval="proc");

    proc print data=sashelp.class;
    run;

    ods excel options(sheet_name="AGENDA" sheet_interval="proc");

    proc odstext ;
    p 'ANALYSIS OF SASHELP.CLASS';
    run;

    ods excel options(sheet_name="AGESEX" sheet_interval="proc");

    proc print data=havSexAge;
    run;quit;

    ods excel close;


