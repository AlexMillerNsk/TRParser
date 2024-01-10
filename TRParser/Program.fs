
open FSharp.Interop.Excel
open FSharp.Json
open System.IO
open Types


type DataTypesTest = ExcelFile<"Grafik.xlsx",ForceString=true>
type Row = DataTypesTest.Row

let file = new DataTypesTest()

let notIsNull (row:Row) = row.``№ п/п`` |> isNull |> not
             
let validateDebit (x:obj) = 
    let w = x|>string
    match w with
    | "0"    -> None
    | "12"   -> Some 12              
    | "24"   -> Some 24               
    | _      -> None

let validatePressure (x:obj) = 
    let w = x|>string
    match w with
    | "Рзаб" -> Some 12             
    | _      -> None

let validateMilling (x:obj) = 
    let w = x|>string
    match w with
    | "1"    -> Some 12             
    | _      -> None 

let validateMethanol (x:obj) = 
    let w = x|>string
    match w with
    | "X"    -> Some 12             
    | _      -> None 

let validateFluidSampling (x:obj) = 
    let w = x|>string
    match w with
    | "%"    -> Some 12
    | "П"    -> Some 12
    | _      -> None 

let parseProduction (validateFunc: obj -> int option) (x:int) (r:Row) = validateFunc (r.GetValue(x)) 

let path = Path.Combine(Path.GetTempPath(), "Sampl546.txt") 
       
let createProcess func (x:Row) processName =
    for index in {13..43} do 
    if (parseProduction func index x).IsSome then
        let duration = (parseProduction func index x).Value
        let date index = 
            let day = index - 10
            if day <10 then $"0{day}" else  string day
        let day = date index
        let wellInfo = WellInfoConstructor.Create x.скв x.куст processName $"2023-07-{day} 00:00:00.000 +0700" duration
        let json = Json.serialize wellInfo    
        use sw = new StreamWriter(new FileStream(path, FileMode.Append, FileAccess.Write))
        sw.WriteLineAsync(json)|>ignore
    else ()


let parseProcess (r:Row) = 
    match r.процесс with
    | "Подъем флюида"              -> let processName = "FluidLifting"
                                      printf "0"                                     
    | "Замер дебита"               -> let processName = "MeasurementDebit"                                      
                                      createProcess validateDebit r processName
    | "замер забойного"            -> let processName = "BottomHolePressure"
                                      createProcess validatePressure r processName
    | "Фрезерование Этанол"        -> let processName = "Milling"
                                      createProcess validateMilling r processName
    | "фрезерование Этанол"        -> let processName = "Milling"
                                      createProcess validateMilling r processName
    | "фрезерование Кредит-Альянс" -> let processName = "Milling"
                                      createProcess validateMilling r processName
    | "фрезерование ОТК"           -> let processName = "Milling"
                                      createProcess validateMilling r processName
    | "метанольная обработка"      -> let processName = "Methanol"
                                      createProcess validateMethanol r processName
    | "тепловая обработка"         -> let processName = "Treasurement"
                                      printf "1"
    | "Отбор проб"                 -> let processName = "FluidSampling"
                                      createProcess validateFluidSampling r processName
    | _ as xt                      -> printf $"{xt}"
                                      
let row1 = file.Data |> Seq.filter notIsNull |> List.ofSeq|>List.map (fun x -> parseProcess x)






