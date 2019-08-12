namespace FSharp.ADODB.RecordsetTypeProvider

open System
open System.Reflection

open Microsoft.FSharp.Core.CompilerServices
open Microsoft.FSharp.Quotations

open ProviderImplementation.ProvidedTypes

// Disable incomplete matches warning
// Incomplete matches are used extensively within this file
// to simplify the code
#nowarn "0025"

[<TypeProvider>]
type RecordsetProvider() as this =
  inherit TypeProviderForNamespaces()

  let buildGetterExpr key =
    fun [rs] -> <@@ (((%%rs: obj) :?> ADODB.Recordset).Fields.[key].Value) @@>

  let buildSetterExpr key =
    fun [rs; newval] -> <@@ (((%%rs:obj) :?> ADODB.Recordset).Fields.[key].Value <- %%(Expr.Coerce(newval, typeof<obj>))) @@>

  let buildGuidGetterExpr key =
    fun [rs] -> <@@ (((%%rs: obj) :?> ADODB.Recordset).Fields.[key].Value :?> System.String |> System.Guid.Parse) @@>

  let buildGuidSetterExpr key =
    fun [rs; newval] -> <@@ (((%%rs: obj) :?> ADODB.Recordset).Fields.[key].Value <- (%%(Expr.Coerce(newval, typeof<System.Guid>)): System.Guid).ToString("B")) @@>

  let ns = "FSharp.ADODB.RecordsetProvider"
  let asm = Assembly.GetExecutingAssembly()

  let recordsetType = ProvidedTypeDefinition(asm, ns, "RecordsetProvider", None)

  do
    recordsetType.DefineStaticParameters(
      [ ProvidedStaticParameter("fileName", typeof<string>) ],
      instantiationFunction = (
        fun typeName [| :? string as fileName |] -> 
          let ty = ProvidedTypeDefinition(asm, ns, typeName, None)

          makeProvidedConstructor
              []
              (fun [] -> <@@ fileName |> Recordsets.openRecordset @@>)
          |> addDelayedXmlComment "Open recordset from a file."
          |> ty.AddMember

          let rowTy = ProvidedTypeDefinition("Row", Some(typeof<obj>))
          
          fileName 
          |> Recordsets.openRecordset
          |> Recordsets.readFields
          |> Seq.map (fun i -> match i.Value.Type with
                                | ADODB.DataTypeEnum.adGUID ->
                                        i.Key
                                        |> makeProvidedProperty typeof<Guid> (buildGuidGetterExpr i.Key) (buildGuidSetterExpr i.Key)
                                        |> addDelayedXmlComment (string ADODB.DataTypeEnum.adGUID)
                                | adoType -> match Recordsets.mapAdoType adoType with
                                              | Some dataType ->
                                                      i.Key
                                                      |> makeProvidedProperty dataType (buildGetterExpr i.Key) (buildSetterExpr i.Key)
                                                      |> addDelayedXmlComment (string adoType)
                                              | _ -> 
                                                      i.Key
                                                      |> makeProvidedProperty typeof<System.Object> (buildGetterExpr i.Key) (buildSetterExpr i.Key)
                                                      |> addDelayedXmlComment "Unknown type." 
                                )
          |> Seq.toList
          |> rowTy.AddMembers
         
          "Fields"
          |> makeReadOnlyProvidedProperty (typedefof<seq<_>>.MakeGenericType(typeof<ADODB.Field>)) (fun [rs] -> <@@ Seq.cast<ADODB.Field> ((%%rs: obj) :?> ADODB.Recordset).Fields @@>)
          |> addDelayedXmlComment "Sequence of recordset fields."
          |> ty.AddMember

          "Rows"
          |> makeReadOnlyProvidedProperty (typedefof<seq<_>>.MakeGenericType(rowTy)) (fun [rs] -> <@@ Recordsets.getRows ((%%rs: obj) :?> ADODB.Recordset) @@>)
          |> addDelayedXmlComment "Sequence of data rows."
          |> ty.AddMember

          "Save"
          |> makeProvidedMethod<unit>
               []
               (fun [rs] -> <@@ Recordsets.saveRecordset ((%%rs:obj) :?> ADODB.Recordset) fileName @@>)
          |> addDelayedXmlComment "Save recordset."
          |> ty.AddMember

          rowTy
          |> ty.AddMember

          ty
      ))

  do
    this.AddNamespace(ns, [ recordsetType ])

[<assembly:TypeProviderAssembly>]
do ()