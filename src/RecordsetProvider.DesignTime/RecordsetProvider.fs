namespace FSharp.ADODB.RecordsetTypeProvider

open System
open System.Reflection

open Microsoft.FSharp.Core.CompilerServices
open Microsoft.FSharp.Quotations

open ProviderImplementation.ProvidedTypes

open Recordsets

/// A series of helper functions intended to make OO ProvidedTypes interface more functional
[<AutoOpen>]
module ProvidedTypesHelpers =
  let inline makeProvidedConstructor parameters invokeCode =
    ProvidedConstructor(parameters, invokeCode)

  let inline makeReadOnlyProvidedProperty propType getterCode propName =
    ProvidedProperty(propName, propType, getterCode)

  let inline makeProvidedProperty propType getterCode setterCode propName =
    ProvidedProperty(propName, propType, getterCode, setterCode)

  let inline makeProvidedMethod<'T> parameters invokeCode methodName =
    ProvidedMethod(methodName, parameters, typeof<'T>, invokeCode)

  let inline makeProvidedParameter< ^T> paramName =
    ProvidedParameter(paramName, typeof< ^T>)

  let inline addDelayedXmlComment comment providedMember =
    (^a : (member AddXmlDocDelayed : (unit -> string) -> unit) providedMember, (fun () -> comment))
    providedMember

// Disable incomplete matches warning
// Incomplete matches are used extensively within this file
// to simplify the code
#nowarn "0025"

[<assembly:TypeProviderAssembly>]
do ()

[<TypeProvider>]
type RecordsetProvider(config : TypeProviderConfig) as this =
  inherit TypeProviderForNamespaces(config)

  let ns = "FSharp.ADODB.RecordsetProvider"
  let asm = Assembly.GetExecutingAssembly()

  let recordsetType = ProvidedTypeDefinition(asm, ns, "RecordsetProvider", None)

  do
    recordsetType.DefineStaticParameters(
      [ ProvidedStaticParameter("fileName", typeof<string>) ], 
      instantiationFunction = (
        fun typeName [| :? string as fileName |] -> 
          let ty = ProvidedTypeDefinition(asm, ns, typeName, Some typeof<Recordset>)

          makeProvidedConstructor [] (fun [] -> <@@ new Recordset(fileName) @@>)
          |> addDelayedXmlComment "Open recordset from a file."
          |> ty.AddMember

          let recordTy = ProvidedTypeDefinition("Record", Some typeof<IRecord>)

          let buildGetterExpr key = fun [record] -> 
            <@@
              let record = (%%record : IRecord)

              match record.GetField(key).Type with
              | t when t.Equals typeof<Guid> -> record.[key] :?> string |> Guid.Parse :> obj
              | _ -> record.[key]
            @@>

          let buildSetterExpr key = fun [record; value] -> 
            <@@ 
              let record = (%%record : IRecord)

              record.[key] <-
                match record.GetField(key).Type with
                | t when t.Equals typeof<Guid> -> ((%%value: Guid).ToString("B")) :> obj
                | _ -> %%(Expr.Coerce(value, typeof<obj>)) 
            @@>

          use recordset = new Recordset(fileName)

          recordset.Fields
          |> Seq.map (fun field -> 
                field.Key
                |> makeProvidedProperty field.Value.Type (buildGetterExpr field.Key) (buildSetterExpr field.Key)
                |> addDelayedXmlComment field.Value.DataType)
          |> Seq.toList
          |> recordTy.AddMembers

          "Records"
          |> makeReadOnlyProvidedProperty 
              (typedefof<seq<_>>.MakeGenericType(recordTy))
              (fun [recordset] -> <@@ (%%recordset : Recordset).Records @@>)
          |> addDelayedXmlComment "Sequence of data records."
          |> ty.AddMember

          "Save"
          |> makeProvidedMethod<unit>
               []
               (fun [recordset] -> <@@ (%%recordset : Recordset).Save() @@>)
          |> addDelayedXmlComment "Save recordset."
          |> ty.AddMember

          recordTy
          |> ty.AddMember

          ty
      ))

  do 
    this.AddNamespace(ns, [ recordsetType ])