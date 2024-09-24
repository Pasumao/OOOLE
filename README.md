# OOOLE
객체 지향 OLE 문법의 간단한 래핑
## [CODE](http://github.com/Pasumao/OOOLE)
## DEMO
![image](https://github.com/user-attachments/assets/7ea148c2-8be6-4b82-a3cd-10daba524a9f)
![image](https://github.com/user-attachments/assets/28e7c1ca-117f-4a9c-8caf-17bbdad46881)
![image](https://github.com/user-attachments/assets/ee5eafa9-7cea-43e0-8770-bd84c952734e)
## API
### CL_OLE
<a id="cl_ole"></a>
#### Overview
Extends: NONE
#### DATA
``` ABAP
OBJ      TYPE OLE2_OBJECT
CHILDREN TYPE TABLE OF REF TO CL_OLE
PARENT   TYPE REF TO CL_OLE
```
#### METHOD
##### FREE_ALL
FREE ALL OBJECT。
##### FREE
FREE OBJECT（don't use）。
##### CONSTRUCTOR
``` ABAP
IMPORTING
  E_PARENT TYPE REF TO CL_OLE OPTIONAL
```
##### SET_PROPERTY
SET_PROPERTY。
``` ABAP
IMPORTING
  E_PROPERTY TYPE CHAR32
  E_VALUE    TYPE ANY
```
##### GET_PROPERTY
GET_PROPERTY。
``` ABAP
IMPORTING
  E_PROPERTY TYPE CHAR32
EXPORTING
  I_VALUE    TYPE ANY
```
##### CALL_METHOD
CALL_METHOD。
``` ABAP
IMPORTING
  E_METHOD TYPE C
  E_ARG1   TYPE ANY OPTIONAL
  E_ARG2   TYPE ANY OPTIONAL
  E_ARG3   TYPE ANY OPTIONAL
  E_ARG4   TYPE ANY OPTIONAL
  E_ARG5   TYPE ANY OPTIONAL
CHANGING
  C_RETURN TYPE ANY OPTIONAL
```
##### CALL_METHOD_OF
CALL_METHOD OF RETURN OBJECT。
``` ABAP
IMPORTING
  E_METHOD TYPE C
  E_ARG1   TYPE ANY OPTIONAL
  E_ARG2   TYPE ANY OPTIONAL
  E_ARG3   TYPE ANY OPTIONAL
  E_ARG4   TYPE ANY OPTIONAL
  E_ARG5   TYPE ANY OPTIONAL
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE
```
##### CREATE_OBJECT
CREATE_OBJECT。
``` ABAP
IMPORTING
  E_OBJ TYPE C
```

### CL_OLE_EXCEL
#### Overview
Extends: [CL_OLE](#cl_ole)
#### METHOD
##### WORKBOOKS
workbooks 객체를 반환합니다。
``` ABAP
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_WORKBOOKS
```
##### ACTIVESHEET
ACTIVESHEET객체를 반환합니다。
``` ABAP
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_SHEET
```
##### QUIT
QUIT AND FREE ALL NODE。
##### CELLS
CELLS객체를 반환합니다。
``` ABAP
IMPORTING
  E_ROW       TYPE I
  E_COL       TYPE I
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_CELL
```
##### WORKSHEETS
WORKSHEETS객체를 반환합니다。
``` ABAP
IMPORTING
  E_INDEX      TYPE I
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_SHEET
```
##### SAVEAS
SAVEAS
``` ABAP
IMPORTING
  E_PATH TYPE CHAR1024
```
##### SELECTION
SELECED RANGE를 반환합니다。
``` ABAP
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_RANGE
```
##### RUN
MACRO RUN
``` ABAP
IMPORTING
  E_MACRO TYPE CHAR1024
```
##### SETPROPERTY
SET PROPERTY
``` ABAP
IMPORTING
  E_TITLE               TYPE CHAR1024 OPTIONAL
  E_VISIBLE             TYPE I OPTIONAL
  E_SHEETSINNEWWORKBOOK TYPE I OPTIONAL
```
### CL_OLE_WORKBOOK
#### Overview
Extends: [CL_OLE](#cl_ole)
### CL_OLE_WORKBOOKS
#### Overview
Extends: [CL_OLE](#cl_ole)
#### METHOD
##### ADD
ADD WORKBOOK
``` ABAP
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_WORKBOOK
```
##### OPEN
OPEN LOCAL EXCEL FILE
``` ABAP
IMPORTING
  E_PATH TYPE CHAR1024
```
### CL_OLE_FONT
#### Overview
Extends: [CL_OLE](#cl_ole)
#### METHOD
##### SERPROPERTY
``` ABAP
IMPORTING 
  E_NAME         TYPE CHAR32 OPTIONAL
  E_BOLD         TYPE I OPTIONAL
  E_SIZE         TYPE I OPTIONAL
  E_COLOR        TYPE INT8 OPTIONAL
  E_TINTANDSHADE TYPE I OPTIONAL
  E_ITALIC       TYPE I OPTIONAL
  E_UNDERLINE    TYPE I OPTIONAL
```
### CL_OLE_INTERIOR
#### Overview
Extends: [CL_OLE](#cl_ole)
#### METHOD
##### SERPROPERTY
``` ABAP
IMPORTING
  E_COLOR TYPE INT8 OPTIONAL
```
### CL_OLE_BORDERS
#### Overview
Extends: [CL_OLE](#cl_ole)
#### METHOD
##### SERPROPERTY
``` ABAP
IMPORTING
  E_LINESTYLE TYPE C OPTIONAL
  E_WEIGHT TYPE I OPTIONAL
```
### CL_OLE_RANGE  
#### Overview <a id="CL_OLE_RANGE"></a>  
Extends: [CL_OLE](#cl_ole)  
#### METHOD
##### INSERT
INSERT ONE LINE.
##### DELETE
DELETE ONE LINE.
##### FONT
GET FONT
``` ABAP
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_FONT
```
##### SELECT
SET SELECT
##### INTERIOR
GET INTERIOR
``` ABAP
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_INTERIOR
```
##### BORDERS
GET BORDERS
``` ABAP
IMPORTING
  E_BORDER TYPE C
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_BORDERS
```
###### SETPROPERTY
``` ABAP
IMPORTING
  E_NUMBERFORMAT TYPE CHAR1024 OPTIONAL
```
### CL_OLE_CELL
#### Overview
Extends: [CL_OLE_RANGE](#CL_OLE_RANGE)
#### METHOD
##### VALUE
SET VALUE
``` ABAP
IMPORTING
  E_VALUE TYPE CHAR1024
```
### CL_OLE_ROW
#### Overview
Extends: [CL_OLE_RANGE](#CL_OLE_RANGE)
### CL_OLE_COLUMN
#### Overview
Extends: [CL_OLE_RANGE](#CL_OLE_RANGE)
### CL_OLE_SHEET
#### Overview
Extends: [CL_OLE](#cl_ole)
#### METHOD
##### CELLS
GET CELLS
``` ABAP
IMPORTING
  E_ROW       TYPE I
  E_COL       TYPE I
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_CELL
```
##### ACTIVATE
ACTIVE SHEET
##### ROWS
GET ROWS
``` ABAP
IMPORTING
  E_INDEX     TYPE I
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_ROW
```
##### COLUMNS
GET COLUMNS
``` ABAP
IMPORTING
  E_INDEX     TYPE I
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_COLUMN
```
##### SETPORPERTY
``` ABAP
IMPORTING
  E_NAME TYPE CHAR1024 OPTIONAL
```
##### RANGE
GET RANGE SEND(E_CELL1 AND E_CELL2) OR R_RANGE EX."A1:B2"
``` ABAP
IMPORTING
  E_CELL1 TYPE REF TO CL_OLE_CELL
  E_CELL2 TYPE REF TO CL_OLE_CELL
  E_RANGE TYPE CHAR32
RETURNING VALUE(R_OBJ) TYPE REF TO CL_OLE_RANGE
```
