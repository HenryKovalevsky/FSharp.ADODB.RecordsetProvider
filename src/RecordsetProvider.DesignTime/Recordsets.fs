module Recordsets

let private mapAdoTypeToNetType = function
  | ADODB.DataTypeEnum.adBoolean          ->   typeof<System.Boolean>  
  | ADODB.DataTypeEnum.adUnsignedTinyInt  ->   typeof<System.Byte>     
  | ADODB.DataTypeEnum.adChar             ->   typeof<System.Char>     
  | ADODB.DataTypeEnum.adDBTimeStamp      ->   typeof<System.DateTime> 
  | ADODB.DataTypeEnum.adDate             ->   typeof<System.DateTime> 
  | ADODB.DataTypeEnum.adCurrency         ->   typeof<System.Decimal>  
  | ADODB.DataTypeEnum.adDouble           ->   typeof<System.Double>   
  | ADODB.DataTypeEnum.adSmallInt         ->   typeof<System.Int16>    
  | ADODB.DataTypeEnum.adInteger          ->   typeof<System.Int32>    
  | ADODB.DataTypeEnum.adBigInt           ->   typeof<System.Int64>    
  | ADODB.DataTypeEnum.adTinyInt          ->   typeof<System.SByte>    
  | ADODB.DataTypeEnum.adSingle           ->   typeof<System.Single>   
  | ADODB.DataTypeEnum.adUnsignedSmallInt ->   typeof<System.UInt16>   
  | ADODB.DataTypeEnum.adUnsignedInt      ->   typeof<System.UInt32>   
  | ADODB.DataTypeEnum.adUnsignedBigInt   ->   typeof<System.UInt64>   
  | ADODB.DataTypeEnum.adLongVarWChar     ->   typeof<System.Object>   
  | ADODB.DataTypeEnum.adVarChar          ->   typeof<System.String>   
  | ADODB.DataTypeEnum.adVarWChar         ->   typeof<System.String>   
  | ADODB.DataTypeEnum.adLongVarBinary    ->   typeof<System.Byte[]>   
  | ADODB.DataTypeEnum.adGUID             ->   typeof<System.Guid>     
  // todo: extend mapping.
  | _                                     ->   typeof<System.Object>   

type Field =
  { Name: string
    Type: System.Type
    /// ADODB.DataType
    DataType: string }

type IRecord = 
  abstract member Item : string -> obj with get, set
  abstract member GetField : string -> Field 

type Recordset(fileName) =
  let recordset : ADODB.Recordset = upcast new ADODB.RecordsetClass()
  
  do recordset.Open(
       fileName,
       "Provider=MSPersist",
       ADODB.CursorTypeEnum.adOpenDynamic,
       ADODB.LockTypeEnum.adLockOptimistic
     )

  /// Move cursor position.
  let move index = 
    let pos = int recordset.AbsolutePosition
    recordset.Move(index - pos)

  static let _lock = obj();

  member _.Fields : Map<string, Field> =
    recordset.Fields 
    |> Seq.cast<ADODB.Field> 
    |> Seq.map (fun field -> 
        { Name = field.Name
          Type = mapAdoTypeToNetType field.Type
          DataType = string field.Type })
    |> Seq.map (fun field -> field.Name, field)
    |> Map

  member this.Records =
    seq { 
      for index in 1..recordset.RecordCount do
        { new IRecord with
            member _.Item 
              with get key  = 
                lock _lock (fun _ -> move index; recordset.Fields.[key].Value)
              and set key value = 
                lock _lock (fun _ -> move index; recordset.Fields.[key].Value <- value)
            member _.GetField key  = this.Fields.[key] }
    }

  member _.RecordCount =
    recordset.RecordCount

  member _.Save() = 
    recordset.Save(fileName, ADODB.PersistFormatEnum.adPersistADTG)

  member _.Close() =
    recordset.CancelUpdate()
    recordset.Close()

  interface System.IDisposable with
    member this.Dispose() = this.Close()