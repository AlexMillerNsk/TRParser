
open FSharp.Interop.Excel
open FSharp.Json
open Types


type DataTypesTest = ExcelFile<"Grafik.xlsx",ForceString=true>
type Row = DataTypesTest.Row




//Make convenience methods

let file = new DataTypesTest()



let notIsNull (row:Row) = row.``№ п/п`` |> isNull |> not

let validate x = 
    let w = string x
    match w with
    | "0"       -> None
    | ""        -> None
    | "Рзаб"    -> None
    | "тех огр" -> None
    | _         -> printf $"{x}"
                   Some x 
             
let validateDebit (x:obj) = 
    let w = x|>string
    match w with
    | "0"    -> None
    | "12"   -> printfn "12"
                Some 12
    | "24"   -> printfn "24"
                Some 24
    | _    -> None

 
let parseProduction (x:int) (r:Row) = validateDebit (r.GetValue(x)) 

//let myFun() =
//    for index in {13..23} do parseProduction index (r:Row)
        
let createProcess (x:Row) =
    for index in {13..43} do 
    let duration = (parseProduction index x).Value
    let date index = 
        let day = index - 11
        if day <10 then $"0{day}" else  string day
    let day = date index
    let itog = WellInfoConstructor.Create x.скв x.куст x.процесс duration $"{day}.07.2023"
    let json = Json.serialize itog
    printfn "%s" json



let parseProcess (r:Row) = 
    match r.процесс with
    | "Подъем флюида"              -> let processName = "FluidLifting"
                                      printf $"{processName}"                                     
    | "Замер дебита"               -> let processName = "MeasuremebtDebit"
                                      printf $"{processName}"
                                      createProcess r
    | "замер забойного"            -> let processName = "BottomHolePressure"
                                      printf $"{processName}"
    | "Фрезерование Этанол"        -> let processName = "Milling"
                                      printf $"{processName}"
    | "фрезерование Кредит-Альянс" -> let processName = "Milling"
                                      printf $"{processName}"
    | "фрезерование ОТК"           -> let processName = "Milling"
                                      printf $"{processName}"
    | "метанольная обработка"      -> let processName = "Methanol"
                                      printf $"{processName}"
    | "тепловая обработка"         -> let processName = "Treasurement"
                                      printf $"{processName}"
    | "Отбор проб"                 -> let processName = "zabil kak "
                                      printf $"{processName}"
    | _                            -> let processName = "something wrong"
                                      printf $"{processName}"

//let row = file.Data |> Seq.filter notIsNull |> List.ofSeq|>List.head|> List.map (fun x -> parseProduction x)

let row = file.Data |> Seq.filter notIsNull |> List.ofSeq

let row1 = file.Data |> Seq.filter notIsNull |> List.ofSeq|>List.map (fun x -> parseProcess x)





//myFun r0
//myFun r1
//myFun r2
//myFun r3
