import * as reader from "xlsx";
const fs = require('fs');
import specialCharacters from "./files/SpecialCharacters.json";


 function addRows(data){
   console.log(JSON.stringify(data));
   const workbook = reader.readFile('./test.xlsx');
   const testCasesSheet = workbook.Sheets['Casos de prueba'];
   reader.utils.sheet_add_json(testCasesSheet, data, {skipHeader: true, origin: 'A2'});
   reader.writeFile(workbook, './test-filled.xlsx');
 }

function getTests(jsonData) {
  const testcases = [];

  if (jsonData.suites && Array.isArray(jsonData?.suites)) {
    for (const suite of jsonData.suites) {
      testcases.push(...getTests(suite));
    }
  } else if (jsonData.suites) {
    const keys = Object.keys(jsonData.suites);
    for (const key of keys) {
      testcases.push(...getTests(jsonData.suites[key]));
    }
  } if (jsonData?.cases.length > 0) {
    for (const testcase of jsonData.cases) {
      const parsedSteps = parseSteps(testcase);
      testcases.push({
        suite: solveSpecialCharacters(jsonData.title),
        title: solveSpecialCharacters(testcase.title),
        description: solveSpecialCharacters(testcase.description),
        precondition: solveSpecialCharacters(testcase.preconditions),
        steps: parsedSteps[0],
        data: parsedSteps[2],
        expected_results: parsedSteps[1]
      });
    }
  }
  else{
    console.log('This suite passed as argument has not tests nor suites.');
  }
  return testcases;
}

function parseSteps(test){
  let steps = '';
  let expectedResults = '';
  let data = '';

  for(const step of test.steps){
    steps = steps.concat(solveSpecialCharacters(`${step.position}. ${step.action}\n`));
    expectedResults = expectedResults.concat(solveSpecialCharacters(`${step.position}. ${step.expected_result}\n`));
    data = data.concat(solveSpecialCharacters(`${step.data}\n`));
  }
  return [steps, expectedResults, data];
}

function solveSpecialCharacters(text: string){
   let solvedText = text;
   if(solvedText){
     for(const key of Object.keys(specialCharacters)){
       solvedText = solvedText.replace(new RegExp(key, "g"), specialCharacters[key]);
     }
   }
   return solvedText;
}

let rawdata = fs.readFileSync('./tests.json');
let tests = JSON.parse(rawdata);
const testData = getTests(tests);
addRows(testData);