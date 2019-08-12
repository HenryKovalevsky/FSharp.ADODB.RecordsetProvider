﻿module Recordsets

open System.Collections.Generic

let openRecordset fn = 
  let rs = upcast new ADODB.RecordsetClass() : ADODB.Recordset
  
  rs.Open(
      fn,
      "Provider=MSPersist",
      ADODB.CursorTypeEnum.adOpenDynamic,
      ADODB.LockTypeEnum.adLockOptimistic
      )
  
  rs

let readFields (rs: ADODB.Recordset) =
  let dict = Dictionary<string, ADODB.Field>()

  rs.Fields 
      |> Seq.cast<ADODB.Field> 
      |> Seq.iter (fun f -> dict.[f.Name] <- f)

  dict

let saveRecordset (rs: ADODB.Recordset) (fn: string) =
  rs.Save(fn, ADODB.PersistFormatEnum.adPersistADTG)
  
let getRows (rs: ADODB.Recordset) =
  // if recordset is empty
  if (rs.BOF && rs.EOF) then Seq.empty<ADODB.Recordset> else 
  
  // move to very beginning of the recordset
  rs.MoveFirst()
  rs.MovePrevious()
  
  Seq.unfold (fun (r: ADODB.Recordset) -> 
                r.MoveNext()
                if r.EOF then None else Some(r, r)) rs

let mapAdoType = function
  | ADODB.DataTypeEnum.adBoolean -> Some typeof<System.Boolean>
  | ADODB.DataTypeEnum.adUnsignedTinyInt -> Some typeof<System.Byte>
  | ADODB.DataTypeEnum.adChar -> Some typeof<System.Char>
  | ADODB.DataTypeEnum.adDBTimeStamp -> Some typeof<System.DateTime>
  | ADODB.DataTypeEnum.adDate -> Some typeof<System.DateTime>
  | ADODB.DataTypeEnum.adCurrency -> Some typeof<System.Decimal>
  | ADODB.DataTypeEnum.adDouble -> Some typeof<System.Double>
  | ADODB.DataTypeEnum.adSmallInt -> Some typeof<System.Int16>
  | ADODB.DataTypeEnum.adInteger -> Some typeof<System.Int32>
  | ADODB.DataTypeEnum.adBigInt -> Some typeof<System.Int64>
  | ADODB.DataTypeEnum.adTinyInt -> Some typeof<System.SByte>
  | ADODB.DataTypeEnum.adSingle -> Some typeof<System.Single>
  | ADODB.DataTypeEnum.adUnsignedSmallInt -> Some typeof<System.UInt16>
  | ADODB.DataTypeEnum.adUnsignedInt -> Some typeof<System.UInt32>
  | ADODB.DataTypeEnum.adUnsignedBigInt -> Some typeof<System.UInt64>
  | ADODB.DataTypeEnum.adLongVarWChar -> Some typeof<System.Object>
  | ADODB.DataTypeEnum.adVarChar -> Some typeof<System.String>
  | ADODB.DataTypeEnum.adVarWChar -> Some typeof<System.String>
  | ADODB.DataTypeEnum.adLongVarBinary -> Some typeof<System.Byte[]>
  | _ -> None

