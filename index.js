// npm install js-yaml
//npm install exceljs


const fs = require('fs');
const path = require('path');
const readline = require('readline');
const yaml = require('js-yaml');
const json2xls = require('json2xls');
const ExcelJS = require('exceljs');

function readDir(dirPath,yaml_keys,result) {
  const files = fs.readdirSync(dirPath);
  files.forEach((file) => {
    const filePath = path.join(dirPath, file);
    if (fs.statSync(filePath).isDirectory()) {
        
      readDir(filePath,yaml_keys,result);
    } 
    else {
        checkYamlFile(filePath,yaml_keys,result);
    }
  });
}

function hasOneKey (data,yaml_keys){
    var result = false;
    yaml_keys.forEach((key) => {
        if(data.hasOwnProperty(key)){
            result = true;
        }
    });
    // // data must have this 3 properties to be a valid tempalte
    // if (data.hasOwnProperty('id') && data.hasOwnProperty('name')&& data.hasOwnProperty('description')){
    //     return false;
    // }

    return result;
}


function readYamlFile(filePath,yaml_keys, result) {
    try{
        const fileContent = fs.readFileSync(filePath, 'utf8');
        const data = yaml.load(fileContent);
       
        if (data != null && hasOneKey(data,yaml_keys)){
            fileNameToken = 'FileName'
            if (fileNameToken in result){
                result[fileNameToken].push(filePath)
            }
            else{
                result[fileNameToken] = [filePath];
            }
    
            yaml_keys.forEach((key) => {
    
                if (data.hasOwnProperty(key) && key in result){

                    result[key].push(data[key])
                }
                else{
                    if(data.hasOwnProperty(key)){
                        result[key] = [data[key]];
                    }
                    
                }
            });
        }
    }
    catch(error){
        console.log("Could not process - "+ filePath);
    }
  }
  
  function checkYamlFile(filePath,yaml_keys, result) {
    const extname = path.extname(filePath);
    if (extname === '.yaml' || extname === '.yml') {
        readYamlFile(filePath,yaml_keys, result);
    } 
    else {
    }
  }
  
function writeExcelFile(resultJson){
    workbook = new ExcelJS.Workbook();
    worksheet = workbook.addWorksheet('Sheet1');

    let rowNumber = 1;
    let colNumber = 1;
    for (const key in resultJson) {
        if (resultJson.hasOwnProperty(key)) {
            values = resultJson[key];
            worksheet.getCell(`${String.fromCharCode(64 + colNumber)}${rowNumber}`).value = key;
            rowNumber++;
            values.forEach((value, index) => {
                worksheet.getCell(`${String.fromCharCode(64 + colNumber)}${rowNumber}`).value = value;
                rowNumber++;
            });
            rowNumber =1;
        }
        colNumber++;
    }
    workbook.xlsx.writeFile('output.xlsx');
}

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

rl.question('Enter directory name: ', (dirName) => {
    var yaml_keys;
    var result={};
    console.log(`Directory name is ${dirName}`);
    rl.question('Enter input separated by commas: ', (input) => {
        yaml_keys = input.split(',');
        console.log(yaml_keys);
        readDir(dirName,yaml_keys,result);
        writeExcelFile(result)
        rl.close();
    });
   
});
