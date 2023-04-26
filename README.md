# ADO Recordset Type Provider

The simple F# type provider for an ADO (ActiveX Data Objects) recordset stored in a file.

## Prerequisites

- [.NET Framework 4.8](https://dotnet.microsoft.com/download/visual-studio-sdks);
- [Paket](https://fsprojects.github.io/Paket/) to manage dependencies.

## Usage

*See [sample/Script.fsx](https://github.com/HenryKovalevsky/FSharp.ADODB.RecordsetProvider/blob/master/sample/Script.fsx) for usage example.*

_______________________________________________________

Currently supported data types (others are not tested):

```ocaml
ADODB.DataTypeEnum.adBoolean         
ADODB.DataTypeEnum.adUnsignedTinyInt 
ADODB.DataTypeEnum.adChar            
ADODB.DataTypeEnum.adDBTimeStamp     
ADODB.DataTypeEnum.adDate            
ADODB.DataTypeEnum.adCurrency        
ADODB.DataTypeEnum.adDouble          
ADODB.DataTypeEnum.adSmallInt        
ADODB.DataTypeEnum.adInteger         
ADODB.DataTypeEnum.adBigInt          
ADODB.DataTypeEnum.adTinyInt         
ADODB.DataTypeEnum.adSingle          
ADODB.DataTypeEnum.adUnsignedSmallInt
ADODB.DataTypeEnum.adUnsignedInt     
ADODB.DataTypeEnum.adUnsignedBigInt  
ADODB.DataTypeEnum.adLongVarWChar    
ADODB.DataTypeEnum.adVarChar         
ADODB.DataTypeEnum.adVarWChar        
ADODB.DataTypeEnum.adLongVarBinary   
ADODB.DataTypeEnum.adGUID            
```