import * as reader from "xlsx";
import tests from "../tests.json";
import specialCharacters from "./files/SpecialCharacters.json";


 function addRows(data){
   const workbook = reader.readFile('./test.xlsx');
   const testCasesSheet = workbook.Sheets['Casos de prueba'];
   reader.utils.sheet_add_json(testCasesSheet, data, {skipHeader: true, origin: 'A2'});
   reader.writeFile(workbook, './test-filled.xlsx');
 }

function getTests(jsonData) {
  const testcases = [];
  if(jsonData?.cases.length > 0){
    for(const testcase of jsonData.cases){
      const parsedSteps = parseSteps(testcase);
      testcases.push({
        suite: solveSpecialCharacters(jsonData.title),
        title: solveSpecialCharacters(testcase.title),
        description: solveSpecialCharacters(testcase.description),
        precondition: solveSpecialCharacters(testcase.preconditions),
        steps: parsedSteps[0],
        expected_results: parsedSteps[1]
      });
    }
  }
  else if(jsonData?.suites.length > 0){
    for(const suite of jsonData.suites){
      console.log(suite.title);
      testcases.push(...getTests(suite));
    }
  }
  else{
    throw 'Object passed as argument has not tests nor suites.'
  }
  return testcases;
}

function parseSteps(test){
  let steps = '';
  let expectedResults = '';

  for(const step of test.steps){
    steps = steps.concat(solveSpecialCharacters(`${step.position}. ${step.action}\n`));
    expectedResults = expectedResults.concat(solveSpecialCharacters(`${step.position}. ${step.expected_result}\n`));
  }
  return [steps, expectedResults];
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

const testData = getTests(tests);
addRows(testData);