# Excel Helper

Allows to easily access data from Excel sheet and also cache it in the SV Context for faster access.

Cache is refreshed if Excel timestamp changes.

## How to Use

Compile and put resulting dlls to the Extension folder of the SV Designer/Server. 

E.g. `c:\Program Files\Micro Focus\Service Virtualization Designer\Designer\Extensions\`

## How to Compile

Create `SV_BIN_FOLDER` environment variable and point it to the SV bin folder. 

E.g. `c:\Program Files\Micro Focus\Service Virtualization Designer\Designer\bin\`

## Usage Example

```
// creates and a sheet which supports headers and fast lookup on INDEX column and stores it in the Service context
ExcelSheet sheet = sv.Contexts.Service.GetOrCreateCachedExcelSheet(@"C:\data\excel.xlsx", "Sheet1", true, "INDEX");

// gets the cell value from the row in "INDEX" column identified by "mykey" and retrieves value from the "DATA" column
string value = sheet.LookupCellValue("mykey", "DATA");
```

