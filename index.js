const reader = require('xlsx');
const file = reader.readFile('./test.xlsm');
const fs = require("fs");

let read = (cell) => {
    let value;
    if(cell) {
        if(cell.v != undefined) {
            if(typeof cell.v == "number") { value = Math.round(cell.v) }
            else { value = cell.v }
        } else {
            value = cell
        }
    } else {
        value = null;
    }
    
    return value;
}

  
let sheetdata = [];
  
let sheets = []
file.SheetNames.forEach(sn => {
    if(sn != "TAKEOFF" && sn != "CONTRACT" && sn != "RECAP") { sheets.push(sn) }
})

console.log(sheets);
// for each sheet > get the relevant values
// k15 - k31

for(let i = 0; i < sheets.length; i++) {
    let s = file.Sheets[sheets[i]];
    sheetdata.push({
        MATERIAL: read(s["K15"]),
        FAB_LABOR: read(s["K16"]),
        FAB_FRINGE: read(s["K17"]),
        FAB_OTHER: read(s["K18"]),
        INSTALL_LABOR: read(s["K19"]),
        INSTALL_FRINGE: read(s["K20"]),
        INSTALL_OTHER: read(s["K21"]),
        FUEL: read(s["K22"]),
        EQUIPMENT: read(s["K23"]),
        SMALL_TOOLS: read(s["K24"]),
        MEALS: read(s["K25"]),
        LODGING: read(s["K26"]),
        TRAVEL_TIME: read(s["K27"]),
        PM: read(s["K28"]),
        FEE: read(s["K30"]),
        TOTAL: read(s["K31"]),
        // TS_DATA: {
        //     PROJECT: read(s["L2"]),
        //     MATERIAL: read(s["L3"]),
        //     WASTE: read(s["D6"]),
        //     SLAB_WIDTH: read(s["E6"]),
        //     SLAB_HEIGHT: read(s["F6"]),
        //     MATERIAL_SQFT_COST: read(s["H6"]),
        //     SQFT: read(s["O6"]),
        //     LABOR: read(s["Q6"]),
        // }
    })
}

let formatCSV = (data) => {
    let rows = [];
    rows.push(`${Object.keys(data[0]).join(",")}`);
    data.forEach(row => {
        rows.push(Object.keys(row).map(k => row[k]).join(","));
    })
    rows = rows.join("\n");
    return rows;
}

fs.writeFile("./test.csv", formatCSV(sheetdata), (error, res) => {
    if(error) {
        console.log(error);
    } else {
        console.log("file saved");
    }
});
// console.log(sheetdata)