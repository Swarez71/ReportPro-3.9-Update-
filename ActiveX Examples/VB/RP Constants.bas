Attribute VB_Name = "Module1"
Const CONNECTION_JASMINE = 2
Const CONNECTION_ODBC = 1
Const EB_ENFORCE_DATE = 4
Const EB_ENFORCE_LOGIC = 1
Const EB_ENFORCE_NONE = 0
Const EB_ENFORCE_NUMERIC = 2
Const EB_ENFORCE_STRING = 8
Const JASQUERY_ATTR_CHILD_TABLES = 1012
Const JASQUERY_ATTR_DATABASE = 4000
Const JASQUERY_ATTR_ENV_FILE = 4001
Const JASQUERY_ATTR_POSTQUERY_ODQL = 4002
Const JASQUERY_ATTR_PREQUERY_ODQL = 4003
Const JASQUERY_ATTR_PWD = 4004
Const JASQUERY_ATTR_UID = 4005
Const JASQUERY_ATTR_WHERE = 4006
Const RPT_ATTR_EPOCH = 10010
Const RPT_ATTR_ANSI = 10006
Const RPT_ATTR_CENTURY = 10007
Const RPT_ATTR_DELETED = 10008
Const RPT_ATTR_COLLATION = 10009
Const RPT_ATTR_ALLPAGES = 11000
Const RPT_ATTR_COLLATE_COPIES = 11001
Const RPT_ATTR_CONNECTED = 10000
Const RPT_ATTR_COPIES = 11002
Const RPT_ATTR_DEFAULT_PAPER_SOURCE = 11003
Const RPT_ATTR_DISABLE_PRINT2FILE = 11004
Const RPT_ATTR_DUPLEX = 11005
Const RPT_ATTR_EXPORT_CAPTION = 11006
Const RPT_ATTR_EXPORT_FILE_NAME = 11007
Const RPT_ATTR_EXPORT_MESSAGE = 11008
Const RPT_ATTR_FROM_PAGE = 11009
Const RPT_ATTR_HIDE_PRINT2FILE = 11010
Const DMPAPER_10X11 = 45                            ' 10 x 11 in
Const DMPAPER_10X14 = 16                            ' 10x14 in
Const DMPAPER_11X17 = 17                            ' 11x17 in
Const DMPAPER_15X11 = 46                            ' 15 x 11 in
Const DMPAPER_9X11 = 44                             ' 9 x 11 in
Const DMPAPER_A_PLUS = 57                           ' SuperA/SuperA/A4 227 x 356 mm
Const DMPAPER_A2 = 66                               ' A2 420 x 594 mm
Const DMPAPER_A3 = 8                                ' A3 297 x 420 mm
Const DMPAPER_A3_EXTRA = 63                         ' A3 Extra 322 x 445 mm
Const DMPAPER_A3_EXTRA_TRANSVERSE = 68              ' A3 Extra Transverse 322 x 445 mm
Const DMPAPER_A3_TRANSVERSE = 67                    ' A3 Transverse 297 x 420 mm
Const DMPAPER_A4 = 9                                ' A4 210 x 297 mm
Const DMPAPER_A4_EXTRA = 53                         ' A4 Extra 9.27 x 12.69 in
Const DMPAPER_A4_PLUS = 60                          ' A4 Plus 210 x 330 mm
Const DMPAPER_A4_TRANSVERSE = 55                    ' A4 Transverse 210 x 297 mm
Const DMPAPER_A4SMALL = 10                          ' A4 Small 210 x 297 mm
Const DMPAPER_A5 = 11                               ' A5 148 x 210 mm
Const DMPAPER_A5_EXTRA = 64                         ' A5 Extra 174 x 235 mm
Const DMPAPER_A5_TRANSVERSE = 61                    ' A5 Transverse 148 x 210 mm
Const DMPAPER_B_PLUS = 58                           ' SuperB/SuperB/A3 305 x 487 mm
Const DMPAPER_B4 = 12                               ' B4 250 x 354
Const DMPAPER_B5 = 13                               ' B5 182 x 257 mm
Const DMPAPER_B5_EXTRA = 65                         ' B5 (ISO) Extra 201 x 276 mm
Const DMPAPER_B5_TRANSVERSE = 62                    ' B5 (JIS) Transverse 182 x 257 mm
Const DMPAPER_CSHEET = 24                           ' C size sheet
Const DMPAPER_DSHEET = 25                           ' D size sheet
Const DMPAPER_ENV_10 = 20                           ' Envelope #10 4 1/8 x 9 1/2
Const DMPAPER_ENV_11 = 21                           ' Envelope #11 4 1/2 x 10 3/8
Const DMPAPER_ENV_12 = 22                           ' Envelope #12 4 \276 x 11
Const DMPAPER_ENV_14 = 23                           ' Envelope #14 5 x 11 1/2
Const DMPAPER_ENV_9 = 19                            ' Envelope #9 3 7/8 x 8 7/8
Const DMPAPER_ENV_B4 = 33                           ' Envelope B4  250 x 353 mm
Const DMPAPER_ENV_B5 = 34                           ' Envelope B5  176 x 250 mm
Const DMPAPER_ENV_B6 = 35                           ' Envelope B6  176 x 125 mm
Const DMPAPER_ENV_C3 = 29                           ' Envelope C3  324 x 458 mm
Const DMPAPER_ENV_C4 = 30                           ' Envelope C4  229 x 324 mm
Const DMPAPER_ENV_C5 = 28                           ' Envelope C5 162 x 229 mm
Const DMPAPER_ENV_C6 = 31                           ' Envelope C6  114 x 162 mm
Const DMPAPER_ENV_C65 = 32                          ' Envelope C65 114 x 229 mm
Const DMPAPER_ENV_DL = 27                           ' Envelope DL 110 x 220mm
Const DMPAPER_ENV_INVITE = 47                       ' Envelope Invite 220 x 220 mm
Const DMPAPER_ENV_ITALY = 36                        ' Envelope 110 x 230 mm
Const DMPAPER_ENV_MONARCH = 37                      ' Envelope Monarch 3.875 x 7.5 in
Const DMPAPER_ENV_PERSONAL = 38                     ' 6 3/4 Envelope 3 5/8 x 6 1/2 in
Const DMPAPER_ESHEET = 26                           ' E size sheet
Const DMPAPER_EXECUTIVE = 7                         ' Executive 7 1/4 x 10 1/2 in
Const DMPAPER_FANFOLD_LGL_GERMAN = 41               ' German Legal Fanfold 8 1/2 x 13 in
Const DMPAPER_FANFOLD_STD_GERMAN = 40               ' German Std Fanfold 8 1/2 x 12 in
Const DMPAPER_FANFOLD_US = 39                       ' US Std Fanfold 14 7/8 x 11 in
Const DMPAPER_FIRST = DMPAPER_LETTER
Const DMPAPER_FOLIO = 14                            ' Folio 8 1/2 x 13 in
Const DMPAPER_ISO_B4 = 42                           ' B4 (ISO) 250 x 353 mm
Const DMPAPER_JAPANESE_POSTCARD = 43                ' Japanese Postcard 100 x 148 mm
Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN     ' DMPAPER_A3_EXTRA_TRANSVERSE
Const DMPAPER_LEDGER = 4                            ' Ledger 17 x 11 in
Const DMPAPER_LEGAL = 5                             ' Legal 8 1/2 x 14 in
Const DMPAPER_LEGAL_EXTRA = 51                      ' Legal Extra 9 \275 x 15 in
Const DMPAPER_LETTER = 1                            ' Letter 8 1/2 x 11 in
Const DMPAPER_LETTER_EXTRA = 50                     ' Letter Extra 9 \275 x 12 in
Const DMPAPER_LETTER_EXTRA_TRANSVERSE = 56          ' Letter Extra Transverse 9\275 x 12 in
Const DMPAPER_LETTER_PLUS = 59                      ' Letter Plus 8.5 x 12.69 in
Const DMPAPER_LETTER_TRANSVERSE = 54                ' Letter Transverse 8 \275 x 11 in
Const DMPAPER_LETTERSMALL = 2                       ' Letter Small 8 1/2 x 11 in
Const DMPAPER_NOTE = 18                             ' Note 8 1/2 x 11 in
Const DMPAPER_QUARTO = 15                           ' Quarto 215 x 275 mm
Const DMPAPER_RESERVED_48 = 48                      ' RESERVED--DO NOT USE
Const DMPAPER_RESERVED_49 = 49                      ' RESERVED--DO NOT USE
Const DMPAPER_STATEMENT = 6                         ' Statement 5 1/2 x 8 1/2 in
Const DMPAPER_TABLOID = 3                           ' Tabloid 11 x 17 in
Const DMPAPER_TABLOID_EXTRA = 52                    ' Tabloid Extra 11.69 x 18 in
Const DMPAPER_USER = 256
Const ZOOM_100 = 2
Const ZOOM_PAGE_WIDTH = 1
Const ZOOM_WHOLE_PAGE = 0
Const RPT_ATTR_MAX_PAGE = 11011
Const RPT_ATTR_MIN_PAGE = 11012
Const RPT_ATTR_NO_PAGENUMS = 11013
Const RPT_ATTR_PAGE_NUMS = 11014
Const RPT_ATTR_PREVIEW_CAPTION = 11015
Const RPT_ATTR_PREVIEW_MODAL = 11016
Const RPT_ATTR_PREVIEW_NOZOOM = 11017
Const RPT_ATTR_PREVIEW_PAGECOUNT = 11018
Const RPT_ATTR_PREVIEW_SHOW_MODE = 11037
Const RPT_ATTR_PREVIEW_ZOOM_MODE = 11019
Const RPT_ATTR_PRINT_CAPTION = 11020
Const RPT_ATTR_PRINT_JOB_TITLE = 11021
Const RPT_ATTR_PRINT_MESSAGE1 = 11022
Const RPT_ATTR_PRINT_MESSAGE2 = 11023
Const RPT_ATTR_PRINT2FILE = 11024
Const RPT_ATTR_PRINT2FILE_NAME = 11025
Const RPT_ATTR_PRINTER_DRIVER = 11026
Const RPT_ATTR_PRINTER_NAME = 11027
Const RPT_ATTR_PRINTER_PORT = 11028
Const RPT_ATTR_PROMPT_FOR_FILE = 11029
Const RPT_ATTR_PROMPT_FOR_PRINTDLG = 11030
Const RPT_ATTR_REPORT_DESCRIPTION = 10001
Const RPT_ATTR_REPORT_FILE = 10002
Const RPT_ATTR_REPORT_TITLE = 10003
Const RPT_ATTR_SECTION_COUNT = 10004
Const RPT_ATTR_SUPPORT_1_OF_N = 10005
Const RPT_ATTR_TO_PAGE = 11031
Const RPT_ATTR_YIELDAFTER = 11032
Const SECTION_ATTR_BOTTOM_MARGIN = 12000
Const SECTION_ATTR_FILTER_EXP = 12001
Const SECTION_ATTR_LANDSCAPE = 12002
Const SECTION_ATTR_LEFT_MARGIN = 12003
Const SECTION_ATTR_PAPER_BIN = 12004
Const SECTION_ATTR_PAPER_LENGTH = 12005
Const SECTION_ATTR_PAPER_SIZE = 12006
Const SECTION_ATTR_PAPER_WIDTH = 12007
Const SECTION_ATTR_PRIMARY_TABLE = 12014
Const SECTION_ATTR_RIGHT_MARGIN = 12008
Const SECTION_ATTR_SORT_ORDER_TEXT = 12009
Const SECTION_ATTR_SORT_ORDER_UNIQUE = 12010
Const SECTION_ATTR_TOP_MARGIN = 12011
Const SECTION_ATTR_VARIABLE_COUNT = 12012
Const SQLQUERY_ATTR_CHILD_SQLTABLES = 2014
Const SQLQUERY_ATTR_CHILD_TABLES = 1012
Const SQLQUERY_ATTR_ODBC_SOURCE = 2000
Const SQLQUERY_ATTR_PWD = 2001
Const SQLQUERY_ATTR_SQL_COL_DELIM = 2002
Const SQLQUERY_ATTR_SQL_DISTINCT = 2003
Const SQLQUERY_ATTR_SQL_FILTER_WHERE = 2004
Const SQLQUERY_ATTR_SQL_FROM = 2005
Const SQLQUERY_ATTR_SQL_GROUP_BY = 2006
Const SQLQUERY_ATTR_SQL_HAVING = 2007
Const SQLQUERY_ATTR_SQL_ORDERBY = 2008
Const SQLQUERY_ATTR_SQL_NULL_AS_DEFAULT = 2009
Const SQLQUERY_ATTR_SQL_TABLE_WHERE = 2010
Const SQLQUERY_ATTR_SQL_UNION = 2011
Const SQLQUERY_ATTR_SQL_USER_COLS = 2012
Const SQLQUERY_ATTR_UID = 2013
Const SQLTABLE_ATTR_CHILD_TABLES = 1012
Const SQLTABLE_ATTR_TABLE = 3000
Const TABLE_ATTR_CHILD_TABLES = 1012
Const TABLE_ATTR_CLASSNAME = 1011
Const TABLE_ATTR_DRIVER = 1001
Const TABLE_ATTR_FILTER_EXPRESSION = 1002
Const TABLE_ATTR_INDEX_FILE = 1003
Const TABLE_ATTR_INDEX_TAG = 1004
Const TABLE_ATTR_SEEK_EXPRESSION = 1005
Const TABLE_ATTR_START_REC = 1006
Const TABLE_ATTR_TABLE = 1007
Const TABLE_ATTR_WHILE_EXPRESSION = 1008
Const TABLE_ATTR_WORK_AREA = 1009
Const UPDATE_PAPERINFO_ALWAYS = 1
Const UPDATE_PAPERINFO_NEVER = 0
Const UPDATE_PAPERINFO_PROMPTUSER = 2
Const VARIABLE_ATTR_INIT_EXPRESSION = 13001
Const VARIABLE_ATTR_NAME = 13002
Const VARIABLE_ATTR_RESET_LEVEL = 13003
Const VARIABLE_ATTR_UPDATE_EXPRESSION = 13004
Const VARIABLE_ATTR_UPDATE_LEVEL = 13005
