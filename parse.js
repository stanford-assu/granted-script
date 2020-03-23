
//STEP 1: Read File

var fs = require('fs');
const jsdom = require("jsdom");
const { JSDOM } = jsdom;

var contents = fs.readFileSync('funding_detail_report.xls', 'utf8');

const dom = new JSDOM(contents);

var table = dom.window.document.querySelector("table");

//STEP 2: Parse into JS Array
keys = [];

var row = table.rows[1];
for (j in row.cells) {
    var cell = row.cells[j];
    try {
        keys.push(cell.textContent.trim());
    } catch (err) {
        //dont add to array
    }
}

data = [];

for (let i = 2; i < table.rows.length; i++) {
    var row = table.rows[i];
    var text = [];
    for (j in row.cells) {
        var cell = row.cells[j];
        try {
            text.push(cell.textContent.trim());
        } catch (err) {
            //dont add to array
        }
    }
    data.push(text);
}

error = false;
for (i in data) {
    if(data[i].length != keys.length){
        console.log("Parsing error at line: ",i+3);
        error = true;
    }
}
if(!error){
    console.log("Successfully parsed", data.length, "lines of data!");
}

//STEP 3: Reduce into JSON of application
ids = new Set();

var apps = {};

appwide = [0,1,2,3,4,5,6,7,8,9] // these columns describe the application
linewide = [10,11,12,13,14] // these columns describe the 

for(i in data){
    var id = data[i][2];
    //this only runs for the first line of each app
    if(!(id in apps)){
        app = {}
        for(j in appwide){
            app[keys[appwide[j]]] = data[i][appwide[j]];
        }
        app["lines"] = [];
        apps[id] = app;
    }
    //this is added for every line
    line = {};
    for(j in linewide){
        line[keys[linewide[j]]] = data[i][linewide[j]];
    }
    apps[id]["lines"].push(line);
}

console.log(JSON.stringify(apps, null, 4))