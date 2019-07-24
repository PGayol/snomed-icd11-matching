//Do not use - testing purposes only

let request = require('request');
if (typeof require !== 'undefined') XLSX = require('xlsx');


// Complete neccesary fields:
//Oauth2 
const token_endpoint = 'https://icdaccessmanagement.who.int/connect/token';
const client_id = '' ; // Your client ID Here
const client_secret = '' ; // Your client secret here
const scope = 'icdapi_access';
const grant_type = 'client_credentials';
let access_token;


//XLS source details
var sourceWorkbook = XLSX.readFile('mini-snomed-list.xls');
var sourceWorksheet = sourceWorkbook.Sheets['Hoja1'];

//Options object to request token
let options = {
  url: token_endpoint,
  form: {
    client_id: client_id,
    client_secret: client_secret,
    scope: scope,
    grant_type: grant_type
  }
};

request.post(options, (error, res, body) => {
  if (error) {
    console.error(error)
    return
  }
  if(`${res.statusCode}` == 200){
    console.log("Successfully connected to server");
  }
  console.log(`statusCode: ${res.statusCode}`);
  var info = JSON.parse(body)
  access_token = info.access_token;

  //Options object to request search
  searchOptions = {
    uri: 'https://id.who.int/icd/entity/search?q={}',
    headers: {
      'Authorization': 'Bearer ' + access_token,
      'Accept': 'application/json',
      'Accept-Language': 'en'
    }
  }

  let searchIndex = 2; //if data matrix starts in row 2
  let finalIndex = 6194 + 2; // #ConceptsToSearch + 2
  sourceWorksheet['A1'] = sourceWorksheet['A1'] ? sourceWorksheet['A1'] : {};
  sourceWorksheet['A1'].t = 's';
  sourceWorksheet['A1'].v = 'Concept ID'
  sourceWorksheet['B1'] = sourceWorksheet['B1'] ? sourceWorksheet['B1'] : {};
  sourceWorksheet['B1'].t = 's';
  sourceWorksheet['B1'].v = 'Name'
  sourceWorksheet['C1'] = sourceWorksheet['C1'] ? sourceWorksheet['C1'] : {};
  sourceWorksheet['C1'].t = 's';
  sourceWorksheet['C1'].v = 'Effective Time'
  sourceWorksheet['D1'] = sourceWorksheet['D1'] ? sourceWorksheet['D1'] : {};
  sourceWorksheet['D1'].t = 's';
  sourceWorksheet['D1'].v = 'ICD11 match'
  sourceWorksheet['E1'] = sourceWorksheet['E1'] ? sourceWorksheet['E1'] : {};
  sourceWorksheet['E1'].t = 's';
  sourceWorksheet['E1'].v = 'Score'

  while (searchIndex < finalIndex) {
    let sourceCell = 'B' + searchIndex;
    let resultCell = 'D' + searchIndex;
    let scoreCell = 'E' + searchIndex;
    const range = {
      s: {
        c: 0,
        r: 0
      },
      e: {
        c: 100,
        r: searchIndex + 1
      }
    };
    sourceWorksheet['!ref'] = XLSX.utils.encode_range(range);
    searchOptions.uri = 'https://id.who.int/icd/entity/search?q={' + sourceWorksheet[sourceCell].v + '}';
    sourceWorksheet[resultCell] = sourceWorksheet[resultCell] ? sourceWorksheet[resultCell] : {};
    sourceWorksheet[resultCell].t = 's';
    sourceWorksheet[scoreCell] = sourceWorksheet[scoreCell] ? sourceWorksheet[scoreCell] : {};
    sourceWorksheet[scoreCell].t = 'n';
    request.get(searchOptions, (error, res, body) => {
      if (error) {
        console.error(error);
        sourceWorksheet[resultCell].v = 'Bad request';
        return;
      }
      console.log('Cell' + sourceCell + ' request returned ' + `statusCode: ${res.statusCode}`);
      let info = JSON.parse(body);
      info.DestinationEntities.sort(function (a, b) {
        return b.Score - a.Score;
      });
      if (info.DestinationEntities.length == 0) {
        sourceWorksheet[resultCell].v = 'No match';
      } else {
        sourceWorksheet[resultCell].v = info.DestinationEntities[0].Title;
        sourceWorksheet[scoreCell].v = info.DestinationEntities[0].Score;
      }
      XLSX.writeFile(sourceWorkbook, 'mini-snomed-list-results.xls');

      // console.log("Best match for SNOMEDCT term:'" + worksheet[cell].v + "' is:" + info.DestinationEntities[0].Title);
      // let i = 0;
      // while (i < info.DestinationEntities.length) {
      //   console.log("Best match:" + info.DestinationEntities[i].Title + ' with a score of: ' + info.DestinationEntities[i].Score);
      //   i++;
      // }
    });
    searchIndex++;
  }
});