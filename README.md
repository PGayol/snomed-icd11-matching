# snomed-icd11-matching


Tool for matching SNOMEDCT concepts with ICD11, using public ICD API (https://icd.who.int/icdapi)

- Prerequisistes:
  - Node v8.12.0
  - NPM v6.4.1
  - SNOMEDCT concepts must be in an .xls/.xlsx file following the format of the "mini-snomed-list.xls" file. File cannot contain any       special character, otherwise API request may fail.
  - You must complete your OAuth credentials inside index-asnyc.js file, as well as some other properties of your .xls/.xlsx input file.        Look the code in index-asnyc.js for details.

- Run:
  - Clone the project
  - $ cd ./snomed-icd11-matching
  - $ npm install
  - $ node index-asnyc.js

You'll get two .xls result files. One containing the best match for each concept (if there is one) with it's correspondent
 match score, and the other containing the 10 first best matches.
