let request = require('request');
if (typeof require !== 'undefined') XLSX = require('xlsx');
let striptags = require('striptags');

// Complete neccesary fields:
//Oauth2 
const token_endpoint = 'https://icdaccessmanagement.who.int/connect/token';
const client_id = '***REMOVED***';
const client_secret = '***REMOVED***';
const scope = 'icdapi_access';
const grant_type = 'client_credentials';
let access_token;


//XLS source details
let sourceWorkbook = XLSX.readFile('mini-snomed-list.xls');
let sourceWorkbook2 = XLSX.readFile('mini-snomed-list.xls');
let sourceWorksheet = sourceWorkbook.Sheets['Hoja1'];
let sourceWorksheet2 = sourceWorkbook2.Sheets['Hoja1'];
let searchIndex = 2; //if data matrix starts in row 2
let finalIndex = 6194 + 2; // #ConceptsToSearch + 2

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

let cells_letters = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

let main = function () {

  request.post(options, (error, res, body) => {
    if (error) {
      console.error(error)
      return
    }
    if (`${res.statusCode}` == 200) {
      console.log("Successfully connected to server");
    }
    console.log(`statusCode: ${res.statusCode}`);
    let info = JSON.parse(body)
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

    let getRequest = function (searchIndex) {
      if (searchIndex == finalIndex) {
        return;
      } else {
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
        sourceWorksheet2['!ref'] = XLSX.utils.encode_range(range);
        searchOptions.uri = 'https://id.who.int/icd/entity/search?q={' + sourceWorksheet[sourceCell].v + '}';
        sourceWorksheet[resultCell] = sourceWorksheet[resultCell] ? sourceWorksheet[resultCell] : {};
        sourceWorksheet[resultCell].t = 's';
        sourceWorksheet[scoreCell] = sourceWorksheet[scoreCell] ? sourceWorksheet[scoreCell] : {};
        sourceWorksheet[scoreCell].t = 'n';
        sourceWorksheet2[resultCell] = sourceWorksheet2[resultCell] ? sourceWorksheet2[resultCell] : {};
        sourceWorksheet2[resultCell].t = 's';
        sourceWorksheet2[scoreCell] = sourceWorksheet2[scoreCell] ? sourceWorksheet2[scoreCell] : {};
        sourceWorksheet2[scoreCell].t = 'n';
        request.get(searchOptions, (error, res, body) => {
          if (res.statusCode == 401 || res.statusCode == '401') {
            main();
          } else {
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
              sourceWorksheet2[resultCell].v = 'No match';
              XLSX.writeFile(sourceWorkbook, 'mini-snomed-list-results-small.xls');
              XLSX.writeFile(sourceWorkbook2, 'mini-snomed-list-results-large.xls');
            } else {
              sourceWorksheet[resultCell].v = striptags(info.DestinationEntities[0].Title);
              sourceWorksheet2[resultCell].v = striptags(info.DestinationEntities[0].Title);
              sourceWorksheet[scoreCell].v = info.DestinationEntities[0].Score;
              sourceWorksheet2[scoreCell].v = info.DestinationEntities[0].Score;             
              XLSX.writeFile(sourceWorkbook, 'mini-snomed-list-results-small.xls');
              XLSX.writeFile(sourceWorkbook2, 'mini-snomed-list-results-large.xls');
              if (info.DestinationEntities.length > 1) {
                let i = 1;
                let letter = 0;
                while (i < info.DestinationEntities.length && i < 9) {
                  //extra text cell
                  sourceWorksheet2[cells_letters[letter] + searchIndex] = sourceWorksheet2[cells_letters[letter] + searchIndex] ? sourceWorksheet2[cells_letters[letter] + searchIndex] : {};
                  sourceWorksheet2[cells_letters[letter] + searchIndex].t = 's';
                  sourceWorksheet2[cells_letters[letter] + searchIndex].v = striptags(info.DestinationEntities[i].Title);
                  letter++;
                  //extra score cell
                  sourceWorksheet2[cells_letters[letter] + searchIndex] = sourceWorksheet2[cells_letters[letter] + searchIndex] ? sourceWorksheet2[cells_letters[letter] + searchIndex] : {};
                  sourceWorksheet2[cells_letters[letter] + searchIndex].t = 'n';
                  sourceWorksheet2[cells_letters[letter] + searchIndex].v = info.DestinationEntities[i].Score;               
                  letter++;
                  i++;
                }
                XLSX.writeFile(sourceWorkbook, 'mini-snomed-list-results-large.xls');
              }
            }
            searchIndex++;
            getRequest(searchIndex);
          }
        });
      };
    }
    getRequest(searchIndex);
  });
  return;
}
main();