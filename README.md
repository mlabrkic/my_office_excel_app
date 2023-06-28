# my_office_excel_app

When you're in the Visual Basic Editor, you can see that a small portion of the macros are located in these "Microsoft Excel objects":
1. ThisWorkbook ( ThisWorkbook.cls )
2. Sheet1 (RADNA) ( Sheet1.cls )

Most macros are located in "Modules": <br>
&nbsp;&nbsp;&nbsp;&nbsp;3. Module1 ( Module1.bas )


## Module1.bas

#### Sub No_01_wwms_W() <br>

' CTRL-w

1. Open your task in the work orders application <br>
After that: (Select All: CTRL-a), (Copy: CTRL-c)

2. Activate the "RADNA" sheet (click anywhere in "RADNA") <br>
3. Run this macro (No_01_wwms_W):  CTRL-w

==> <br>
The macro will paste into "WWMS" Excel sheet, and after processing, paste the data we need into "RADNA" sheet.


#### Sub No_02_Copy_uredjaj_L()

' CTRL + l (lowercase letter L) <br>
The macro copies the line we need from the "MPLS" sheet (list of devices) to the "RADNA" sheet.


#### Sub No_03_A_Otvori_Excel_tabl_portova_T()

' CTRL + t <br>
The macro opens the correct sheet (according to the device name) in the Excel file on the network drive.


#### Sub No_03_B_Rezerviraj_port_Excel_tablica_S_T()

' CTRL + SHIFT + t <br>
The macro copies some data from the "RADNA" sheet to an already open Excel file on the network drive.


#### Sub No_04_A_INV_PORT_P()

' CTRL + p <br>
  sFile = "Pokreni_inventory_01_port.bat" <br>
  x1 = Shell(sPath + sFile + " " + sUR_ID + " " + sSlot + " " + Chr(34) & sDescription & Chr(34), vbNormalFocus)

The macro runs my Java utility [Inventory_01_port](https://github.com/mlabrkic/Inventory_01_port) using "Pokreni_inventory_01_port.bat". <br>

Pokreni_inventory_01_port.bat:
``` vba
set EQUIPMENT=%1
set SLOT=%2
set DESCRIPTION=%3

:: Change Current Directory to the location of this batch file
:: http://ss64.com/nt/cd.html
CD /d "%~dp0"

CMD /c %JAVA_HOME%bin\java -cp Inventory_01_port-1.0-SNAPSHOT.jar;dependency204/* com.mxb.inventory.port.Inventory_01_port %EQUIPMENT% %SLOT% %DESCRIPTION%
```
<br>

To make it easier to work with the Ericsson Adaptive Inventory (formerly Ericsson’s Granite Inventory) application, I use my 3 Java utilities below every day. <br>
My utilities [Inventory_01_port](https://github.com/mlabrkic/Inventory_01_port),
 [Inventory_02_ui](https://github.com/mlabrkic/Inventory_02_ui), and
 [Inventory_03_access](https://github.com/mlabrkic/Inventory_03_access)
use the SikuliX Java library. <br>
Open this file and read more about SikuliX:
[Inventory_01_port.java](https://github.com/mlabrkic/Inventory_01_port/blob/main/src/main/java/com/mxb/inventory/port/Inventory_01_port.java)

https://github.com/RaiMan/SikuliX1 <br>
It uses image recognition powered by OpenCV to identify GUI components and can act on them with mouse and keyboard actions.

For use in Java Maven projects the dependency coordinates are:

``` Java Maven
<dependency>
  <groupId>com.sikulix</groupId>
  <artifactId>sikulixapi</artifactId>
  <version>2.0.5</version>
</dependency>
```

I use the SikuliX library because the Inventory application I use runs through a Citrix server
(Citrix server -there is no easy access to a GUI's internals).


#### Sub No_05_Multiple_find_and_replace_D()

CTRL + d (docx)

Here vba [ScriptingDictionary](https://github.com/mlabrkic/vba/tree/main/ScriptingDictionary) is used.

1. First you need to open your Word template.
2. This Excel macro in the Word template replaces keywords ( #words ) with data from "RADNA"
(They are loaded into "ScriptingDictionary". ).


#### Sub No_06_Kopiraj_u_MP_REPORT_R()

' CTRL + r <br>
The macro opens an Excel file on the network drive, then searches for the last row in the table, and copies the data from "MP_REPORT" to the next, empty, row.


#### Sub No_07_SendEmail_M()

' CTRL + m <br>
Macro sends INFO mail. <br>
The data is in "RADNA" and "POPISI".


#### Sub No_08_Prebaci_u_novi_folder_N()

' CTRL + n <br>
The macro creates new folders, and copies some Excel and Word files into them.

