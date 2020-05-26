const FS = require("fs");
const XLSX = require("node-xlsx");
const ChildProcess = require("child_process");
const Readline = require("readline");

let tool= {
    readInput(tips){
        tips = tips || "> ";
        return new Promise((resolve, reject) => {
            let readline = Readline.createInterface({
                input: process.stdin,
                output: process.stdout
            });
            readline.question(tips, (answer) => {
                readline.close();
                resolve(answer.trim());
            });
        })
    },

    makeEmptyDir(path){
        if (FS.existsSync(path)){
            ChildProcess.execSync(`rm -f -d -R ${path}`);
        }
        FS.mkdirSync(path);
    },
};

let settings = {
    workSpace:  "",
    jsonPath:  "",
    tsPath:  "",
    jsPath:  "",
    isJS:  false,
    isTS:  false,
    isJSON:  false,
}

let maker = {
    start(){
        this.preprocess();
    },

    preprocess(){
        tool.readInput("Input a work space which include xlsx files: ")
        .then(input => {
            
            settings.workSpace = input;
            console.log(settings.workSpace);
            settings.jsonPath = `${settings.workSpace}/json`;
            settings.tsPath = `${settings.workSpace}/ts`;
            settings.jsPath = `${settings.workSpace}/js`;

            if (!FS.existsSync(settings.workSpace))
            {
                console.error(`No directory: ${settings.workSpace}`);
                return;
            }
            
            tool.readInput("JSON file need?(y or n)").then(input => {
                settings.isJSON = input === "y" ? true : false;
                tool.readInput("TS file need?(y or n)").then(input => {
                    settings.isTS = input === "y" ? true : false;
                    tool.readInput("JS file need?(y or n)").then(input => {
                        settings.isJS = input === "y" ? true : false;

                        if (settings.isJSON) tool.makeEmptyDir(settings.jsonPath);
                        if (settings.isTS) tool.makeEmptyDir(settings.tsPath);
                        if (settings.isJS) tool.makeEmptyDir(settings.jsPath);

                    });
                });
            });
        });
    },

};

maker.start();

// let tool = {

//     fliterValue(type, value){
//         if (type === "int"){
//             return value || 0;
//         }
//         else if (type === "float"){
//             return value || 0.00;
//         }
//         else if (type === "string"){
//             return value || "";
//         }
//         else if (type === "array"){
//             if (value){
//                 if (value.indexOf === undefined || value.indexOf(",") === -1){
//                     return [value];
//                 }
//                 else {
//                     return value.split(",");
//                 }
//             }
//             else {
//                 return [];
//             }
//         }
//         else {
//         }

//         return "";
//     },

//     clear(){
//         if (FS.existsSync("json")){
//             ChildProcess.execSync("rm -f -d -R json");
//         }

//         if (FS.existsSync("ts")){
//             ChildProcess.execSync("rm -f -d -R ts");
//         }

//         ChildProcess.execSync("mkdir json");
//         ChildProcess.execSync("mkdir ts");
//     },

//     getJSONPath(excelDataItem){
//         return `json/${excelDataItem.name.toLowerCase()}.json`;
//     },

//     getTSPath(excelDataItem){
//         return `ts/${excelDataItem.name.toLowerCase()}.ts`;
//     },

//     getJSONData(excelDataItem){
//         let sourceData = excelDataItem.data;
//         let json = {};
//         let comments = sourceData[0];
//         let types = sourceData[1];
//         let keys = sourceData[2];

//         for (let index = 3; index < sourceData.length; ++index){
//             let dataItem = sourceData[index];
//             json[`${dataItem[0]}`] = {};
//             for (let i in keys){
//                 json[`${dataItem[0]}`][keys[i]] = this.fliterValue(types[i], dataItem[i]);
//             }
//         }

//         let commentData = "";

//         for (let i in keys){
//             commentData += `${keys[i]}: ${this.fliterValue("string", comments[i])}\n`;
//         }

//         return `/*\n${commentData}*/\n${JSON.stringify(json, null, 4)}`;
//     },

//     getTSData(excelDataItem){
//         let sourceData = excelDataItem.data;
//         let json = {};
//         let comments = sourceData[0];
//         let types = sourceData[1];
//         let keys = sourceData[2];

//         for (let index = 3; index < sourceData.length; ++index){
//             let dataItem = sourceData[index];
//             json[`${dataItem[0]}`] = {};
//             for (let i in keys){
//                 json[`${dataItem[0]}`][keys[i]] = this.fliterValue(types[i], dataItem[i]);
//             }
//         }

//         let commentData = "";

//         for (let i in keys){
//             commentData += `${keys[i]}: ${this.fliterValue("string", comments[i])}\n`;
//         }

//         return `/*\n${commentData}*/\nexport default\n${JSON.stringify(json, null, 4)}`;
//     },

//     makeJSON(excelPath){
//         let workSheetsFromFile = XLSX.parse(excelPath);
//         for (let index = 1; index < workSheetsFromFile.length; ++index){
//             let jsonData = this.getJSONData(workSheetsFromFile[index]);
//             let jsonPath = this.getJSONPath(workSheetsFromFile[index]);
//             if (FS.existsSync(jsonPath)){
//                 FS.unlinkSync(jsonPath);
//             }
//             FS.writeFileSync(jsonPath, jsonData);
//         }
//     },

//     makeTS(excelPath){
//         let workSheetsFromFile = XLSX.parse(excelPath);
//         for (let index = 1; index < workSheetsFromFile.length; ++index){
//             let tsData = this.getTSData(workSheetsFromFile[index]);
//             let tsPath = this.getTSPath(workSheetsFromFile[index]);
//             if (FS.existsSync(tsPath)){
//                 FS.unlinkSync(tsPath);
//             }
//             FS.writeFileSync(tsPath, tsData);
//         }
//     },
// }

// tool.clear();

// let files = FS.readdirSync("excel");

// for (let i = 0; i < files.length; ++i){
//     let filename = files[i];
//     if (filename.indexOf(".DS_Store") !== -1){
//         continue;
//     }

//     console.log("********************************");
//     console.log(`start export ${filename}`);
//     tool.makeJSON(`excel/${filename}`);
//     tool.makeTS(`excel/${filename}`);
//     console.log(`finish export ${filename}`);
//     console.log("********************************");
// }





