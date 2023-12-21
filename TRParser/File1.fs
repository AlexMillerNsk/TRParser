module File1

open Microsoft.Office.Interop.Excel

//Open Excel
let xl = new ApplicationClass(Visible = true)




//Make convenience methods
let get (x: int) (y:int) (sheet: _Worksheet) =
    (sheet.Cells.Item(y,x) :?> Range).Value2

let get_string x y sheet =
    get x y sheet :?> string

let get_int x y sheet = 
    get x y sheet :?> int




//Print out the titles needeing retry in a given worksheet
let get_titles_for_retest sheet = 
    let device_name = (get_string 1 1 sheet)

    printfn "%s" (device_name.Substring(37, device_name.Length-37))
    
    for y in 18 .. 55 do
        let status = get_string 4 y sheet

        if status <> null then
            if status.Contains("Fail") || status.Contains("fail") then
                let title = get_string 2 y sheet
                printfn "%s" title
            
            //printfn "%s" status
    printfn "----------"


//Main function in application
//Loop through the files

let wb = xl.Workbooks.Open("Grafik.xlsx")


    //get the worksheet you'll be working with
let sheets = wb.Sheets
    
for i in 3 .. sheets.Count do
    let sheet = (sheets.[box i] :?> _Worksheet)
        
    //Print out the titles needing restest
    get_titles_for_retest sheet

printfn "============"
