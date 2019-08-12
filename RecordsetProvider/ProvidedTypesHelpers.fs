/// A series of helper functions intended to make OO ProvidedTypes interface more functional

[<AutoOpen>]
module ProviderImplementation.ProvidedTypes.ProvidedTypesHelpers

open ProviderImplementation.ProvidedTypes

let inline makeProvidedConstructor parameters invokeCode =
  ProvidedConstructor(parameters, InvokeCode = invokeCode)

let inline makeReadOnlyProvidedProperty propType getterCode propName =
  ProvidedProperty(propName, propType, GetterCode = getterCode)

let inline makeProvidedProperty propType getterCode setterCode propName =
  ProvidedProperty(propName, propType, GetterCode = getterCode, SetterCode = setterCode)

let inline makeProvidedMethod< 'T> parameters invokeCode methodName =
  ProvidedMethod(methodName, parameters, typeof< 'T>, InvokeCode = invokeCode)

let inline makeProvidedParameter< ^T> paramName =
  ProvidedParameter(paramName, typeof< ^T>)

let inline addDelayedXmlComment comment providedMember =
  (^a : (member AddXmlDocDelayed : (unit -> string) -> unit) providedMember, (fun () -> comment))
  providedMember