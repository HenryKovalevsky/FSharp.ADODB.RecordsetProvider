#r @"..\src\RecordsetProvider.DesignTime\bin\Debug\FSharp.ADODB.RecordsetTypeProvider.dll"
#r "ADODB"

let [<Literal>] FileName = __SOURCE_DIRECTORY__ + @"\recordset.rst" 

open FSharp.ADODB.RecordsetProvider

let rs = new RecordsetProvider<FileName>()

for r in rs.Records do
  printfn "%A" r.SrcSentence

let fr = Seq.head rs.Records
let lr = Seq.last rs.Records

printfn "%A" fr.SrcSentence
printfn "%A" lr.SrcSentence

for r in rs.Records do
  printfn "%A"  r.ServerContextID
  r.ServerContextID <- System.Guid.NewGuid()
  
rs.Save()