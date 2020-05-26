const FS = require("fs");
const XLSX = require("node-xlsx");
const ChildProcess = require("child_process");
const Readline = require("readline");

String.prototype.isInclude = function(string){
    return this.indexOf(string) !== -1;
}

String.prototype.duplicate = function(times){
    return (new Array(times + 1)).join(this);
}

String.prototype.replaceAll = function(string, replace){
    let regExp = new RegExp(string, "g");
    return this.replace(string, replace);
}

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

    fliterValue(type, value){
        if (type === "int"){
            return value || 0;
        }
        else if (type === "float"){
            return value || 0.00;
        }
        else if (type === "string"){
            return `"${value}"` || `""`;
        }
        else if (type === "array"){
            return value || "[]";
        }
        else {
            console.error(`Got unknowen type: ${type}`);
        }

        return "";
    },

    wirteFile(path, data){
        if (FS.existsSync(path)){
            FS.unlinkSync(path);
        }
        FS.writeFileSync(path, data);
    },

    printCutlineStart(title){
        console.log(`---------------------------------------------`);
        console.log(`${title} Start ------------------------------`);
    },

    printCurlineDone(title){
        console.log(`${title} Done  ---------------------------------`);
        console.log(`---------------------------------------------`);
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
    xlsxList: [],
}

let maker = {
    start(){
        this._preprocess();
    },

    _preprocess(){
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

                        this._make();

                    });
                });
            });
        });
    },

    _make(){
        let files = FS.readdirSync(settings.workSpace);
        for (let i = 0; i < files.length; ++i){
            if (files[i].indexOf(".xlsx") !== -1){
                settings.xlsxList.push(`${settings.workSpace}/${files[i]}`);
            }
        }

        for (let i = 0; i < settings.xlsxList.length; ++i){
            if (settings.isJSON) this._makeJSON(settings.xlsxList[i]);
            if (settings.isTS) this._makeTS(settings.xlsxList[i]);
            if (settings.isJS) this._makeJS(settings.xlsxList[i]);
        }
    },
    
    _makeXLSXData(excelData){
        let mainName = excelData[1].name;
        let xlsxData = {};
        for (let i = 1; i < excelData.length; ++i){
            xlsxData[excelData[i].name] = excelData[i];
        }
        return xlsxData;
    },

    _makeJSON(xlsxPath){
        tool.printCutlineStart(`Make JSON ${xlsxPath}`);
        let excelData = XLSX.parse(xlsxPath);
        let mainName = excelData[1].name;
        let xlsxData = this._makeXLSXData(excelData);
        let data = this._makeXLSXChunk(xlsxData, mainName);
        tool.wirteFile(`${settings.jsonPath}/${mainName}.json`, data);
        tool.printCurlineDone(`Make JSON`);
    },

    _makeTS(xlsxPath){
        tool.printCutlineStart(`Make TS ${xlsxPath}`);
        let excelData = XLSX.parse(xlsxPath);
        let mainName = excelData[1].name;
        let xlsxData = this._makeXLSXData(excelData);
        let data = this._makeXLSXChunk(xlsxData, mainName);
        tool.wirteFile(`${settings.tsPath}/${mainName}.ts`, `export default\n${data}`);
        tool.printCurlineDone(`Make TS`);
    },

    _makeJS(xlsxPath){
        tool.printCutlineStart(`Make JS ${xlsxPath}`);
        let excelData = XLSX.parse(xlsxPath);
        let mainName = excelData[1].name;
        let xlsxData = this._makeXLSXData(excelData);
        let data = this._makeXLSXChunk(xlsxData, mainName);
        tool.wirteFile(`${settings.jsPath}/${mainName}.js`, `module.exports=${data}`);
        tool.printCurlineDone(`Make JS`);
    },

    _makeXLSXChunk(xlsxData, name){
        let sourceData = xlsxData[name].data;
        let types = sourceData[1];
        let data = "";
        for (let i = 3; i < sourceData.length; ++i){
            data += `\t"${this._makeCellChunk(xlsxData, name, types[0], sourceData[i][0], 0)}":${this._makeRowChunk(xlsxData, name, i, 1)}`
            if (i < sourceData.length - 1){
                data += ",\n";
            }
            else{
                data += "\n";
            }
        }
        return `{\n${data}\n}`;
    },

    _makeSheetChunk(xlsxData, name, indentedCount){
        let sourceData = xlsxData[name].data;
        let indentedString = "\t".duplicate(indentedCount);
        let data = "";
        for (let i = 3; i < sourceData.length; ++i){
            data += `${indentedString}\t${this._makeRowChunk(xlsxData, name, i, indentedCount + 1)}`
            if (i < sourceData.length - 1){
                data += ",\n";
            }
            else{
                data += "\n";
            }
        }
        return `${indentedString}[\n${data}\n${indentedString}]`;
    },

    _makeRowChunk(xlsxData, name, row, indentedCount){
        let sheetData = xlsxData[name].data;
        let types = sheetData[1];
        let keys = sheetData[2];
        let rowData = sheetData[row];
        let indentedString = "\t".duplicate(indentedCount);

        let data = "";
        for (let i = 0; i < types.length; ++i){
            data += `${indentedString}\t"${keys[i]}": ${this._makeCellChunk(xlsxData, name, types[i], rowData[i], indentedCount + 1)}`;
            if (i < types.length - 1){
                data += ",\n";
            }
            else{
                data += "\n";
            }
        }

        return `\n${indentedString}{\n${data}${indentedString}}`;
    },

    _makeCellChunk(xlsxData, name, type, value, indentedCount){
        let indentedString = "\t".duplicate(indentedCount);
        if (type === "array"){
            if (this._isSimpleArray(value)){
                return tool.fliterValue(type, value);
            }

            if (value.isInclude("[") || value.isInclude(",")){
                value = value.substr(0, 1) === "[" ? value.substr(1, value.length - 1) : value;
                value = value.substr(value.length - 1, 1) === "]" ? value.substr(0, value.length - 1) : value;
                let listData = value.split(",");
                let data = "";
                for (let i = 0; i < listData.length; ++i){
                    let sheetName = this._extractSheetName(listData[i]);
                    if (!xlsxData[sheetName]){
                        console.error(`Array data has no sheet1: ${sheetName}`);
                        return "[]";
                    }
                    let row = this._extractSheetRow(listData[i]);
                    data += `${indentedString}\t${this._makeRowChunk(xlsxData, sheetName, row, indentedCount + 1)}`;
                    if (i < listData.length - 1){
                        data += ",";
                    }
                }
                return `\n${indentedString}[${data}\n${indentedString}]`;
            }
            else {
                if (!xlsxData[value]){
                    console.error(`Array data has no sheet2: ${value}`);
                    return "[]";
                }
                return `${indentedString}${this._makeSheetChunk(xlsxData, value, indentedCount + 1)}`;
            }
        }
        else if (type === "object"){
            if (value === null || value === undefined || value === "" || value === {}){
                return "{}";
            }
            if (value.isInclude("(")){
                let sheetName = this._extractSheetName(value);
                if (!xlsxData[sheetName]){
                    console.error(`Obejct data has no sheet: ${sheetName}`);
                    return "{}";
                }
                let row = this._extractSheetRow(value);
                return `${indentedString}${this._makeRowChunk(xlsxData, sheetName, row, indentedCount + 1)}`;
            }
            else {
                if (!xlsxData[value]){
                    console.error(`Obejct data has no sheet: ${value}`);
                    return "{}";
                }
                let data = "";
                for (let i = 3; i < xlsxData[value].data.length; ++i){
                    let rowData = this._makeRowChunk(xlsxData, value, i, indentedCount + 1);
                    data += `${indentedString}\t"${this._makeCellChunk(xlsxData, value, xlsxData[value].data[1][0], xlsxData[value].data[i][0], indentedCount + 1)}":${rowData}`;
                    if (i < xlsxData[value].data.length - 1){
                        data += ",\n";
                    }
                    else {
                        data += "\n";
                    }
                }

                return `${indentedString}{\n${data}\n${indentedString}}`;
            }
        }
        else {
            return tool.fliterValue(type, value);
        }
    },

    _isSimpleArray(cellData){
        if (cellData === null || cellData === undefined || cellData === "" || cellData === "[]"){
            return true;
        }
        let list = cellData.split(",");
        let firstItem = list[0].substr(0, 1) === "[" ? list[0].substr(1, list[0].length - 1) : list[0];
        return (firstItem.indexOf("(") === -1);
    },

    _extractSheetName(data){
        return data.substr(0, data.indexOf("(")).replaceAll(" ", "");
    },

    _extractSheetRow(data){
        let leftIndex = data.indexOf("(");
        let rightIndex = data.indexOf(")");
        if (leftIndex === -1 || rightIndex === -1){
            console.error(`Row Error ${data}`);
            return -1;
        }
        return parseInt(data.substr(leftIndex + 1, rightIndex - leftIndex - 1)) - 1;
    },
};

maker.start();