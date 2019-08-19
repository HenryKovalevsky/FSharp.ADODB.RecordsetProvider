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

  let getRecordsetExpr rs =
    <@@ ((%%rs: obj) :?> (unit->ADODB.Recordset))() @@>

  let buildGetterExpr key =
    fun [rsa] -> <@@ Recordsets.getValue %%(getRecordsetExpr rsa) key @@>

  let buildSetterExpr key =
    fun [rsa; newval] -> <@@ Recordsets.setValue %%(getRecordsetExpr rsa) key %%(Expr.Coerce(newval, typeof<obj>)) @@>

  let buildGuidGetterExpr key =
    fun [rsa] -> <@@ Recordsets.getValue %%(getRecordsetExpr rsa) key :?> string |> Guid.Parse @@>

  let buildGuidSetterExpr key =
    fun [rsa; newval] -> <@@ Recordsets.setValue %%(getRecordsetExpr rsa) key ((%%newval: Guid).ToString("B")) @@>

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

          let recordTy = ProvidedTypeDefinition("Record", Some(typeof<obj>))
          
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
                                                      |> makeProvidedProperty typeof<obj> (buildGetterExpr i.Key) (buildSetterExpr i.Key)
                                                      |> addDelayedXmlComment "Unknown type." 
                                )
          |> Seq.toList
          |> recordTy.AddMembers
         
          "Fields"
          |> makeReadOnlyProvidedProperty (typedefof<seq<_>>.MakeGenericType(typeof<ADODB.Field>)) (fun [rs] -> <@@ Seq.cast<ADODB.Field> ((%%rs: obj) :?> ADODB.Recordset).Fields @@>)
          |> addDelayedXmlComment "Sequence of recordset fields."
          |> ty.AddMember

          "Records"
          |> makeReadOnlyProvidedProperty (typedefof<seq<_>>.MakeGenericType(recordTy)) (fun [rs] -> <@@ Recordsets.initRecordsetAccessor ((%%rs: obj) :?> ADODB.Recordset) @@>)
          |> addDelayedXmlComment "Sequence of data records."
          |> ty.AddMember

          "Save"
          |> makeProvidedMethod<unit>
               []
               (fun [rs] -> <@@ Recordsets.saveRecordset ((%%rs: obj) :?> ADODB.Recordset) fileName @@>)
          |> addDelayedXmlComment "Save recordset."
          |> ty.AddMember

          recordTy
          |> ty.AddMember

          ty
      ))

  do
    this.AddNamespace(ns, [ recordsetType ])

[<assembly:TypeProviderAssembly>]
do ()