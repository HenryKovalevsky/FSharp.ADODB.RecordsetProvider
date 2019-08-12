#r @"bin\Debug\FSharp.ADODB.RecordsetTypeProvider.dll"
#r "ADODB"

let [<Literal>] fileName = __SOURCE_DIRECTORY__ + @"\..\Data\recordset.rst" 

open FSharp.ADODB.RecordsetProvider

let rs = new RecordsetProvider< fileName >()

for r in rs.Rows do
  printfn "%A"  r.ServerContextID

for r in rs.Rows do
  r.ServerContextID <- System.Guid.NewGuid()
  printfn "%A"  r.ServerContextID

rs.Save()
