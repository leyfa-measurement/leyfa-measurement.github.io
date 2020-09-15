var MAIN_SHEET_NAME="Étude";
var FIRST_ROW_MAIN_SHEET=4;
var SC_START_ROW=14;
var SC_H_COL=5;
var SC_L_COL=4;
var SC_NOM_COL=0;

let COLUMNS=["Bornes", "Flèche Intiale", "Flèche proposée", "Ripage", "Dévers initial", "Dévers proposé",
"PK", "ENT G", "ENT D", "CAT1", "CAT2", "CAT3", "CAT MA"];//, "Annotations"];

var COLUMNS_REGEX=[{
  "name": "PK",
  "value_r1": /PK/,
  "value_r2": /(?:)/,
},
{
  "name": "Bornes",
  "value_r1":/Marquage/,
  "value_r2": /LEYFA/
},
{
  "name": "Flèche Intiale",
  "value_r1": /Flèche/,
  "value_r2": /Actuelle/
},
{
  "name": "Flèche proposée",
  "value_r1": /Flèche/,
  "value_r2": /Future/
},
{
  "name": "Ripage",
  "value_r1": /Ripage/,
  "value_r2": /mm/
},
{
  "name": "Dévers initial",
  "value_r1": /Dévers/,
  "value_r2": /Actuel/
},
{
  "name": "Dévers proposé",
  "value_r1":/Dévers/,
  "value_r2": /Futur/
},    
{
  "name": "ENT G",
  "value_r1": /Entraxe final/,
  "value_r2": /gauche/
},
{
  "name": "ENT D",
  "value_r1": /Entraxe final/,
  "value_r2": /droite/
},
{
  "name": "CAT1",
  "value_r1": /V catégorie I/,
  "value_r2": /autre matériel/
},
{
  "name": "CAT2",
  "value_r1": /V catégorie II/,
  "value_r2": /(?:)/
},
{
  "name": "CAT3",
  "value_r1": /V catégorie III/,
  "value_r2": /(?:)/
},
{
  "name": "CAT MA",
  "value_r1": /V catégorie 1/,
  "value_r2": /MA 100/
},
{
  "name": "Elements tracé",
  "value_r1": /Éléments de tracé/,
  "value_r2": /(?:)/
},
{
  "name": "Points Fixes",
  "value_r1": /Points fixes/,
  "value_r2": /(?:)/
}];



var file = document.getElementById('file');
var infos_p = document.getElementById('infos');
var submit_file = document.getElementById('convert_file');
var get_info_file = document.getElementById('get_info_file');


submit_file.addEventListener('click', importFile);
get_info_file.addEventListener('click', showInfoFile);

function changeTextInfos(new_text){
  infos_p.innerHTML=new_text;
}

function addTextInfos(new_text){
  infos_p.innerHTML+=new_text+"<br/>";
}

function showInfos(){
  infos_p.style.display="block";
}

function hideInfos(){
  infos_p.style.display="none";
}

function hasFile(){
  var f = file.files[0];
  return Boolean(f);
}

function showInfoFile(evt){
  changeTextInfos("");
  showInfos();
  if(hasFile()){
    var f = file.files[0];
    addTextInfos(`Calcul des infos pour le fichier ${f.name}`);
    getWorkbookFromFile(f);
    
    
  }else{
    addTextInfos(`<span class="error_info">Pas de fichier détecté.</span>`);
  }
}


function getWorkbookFromFile(f){
  var r = new FileReader();
  r.onload = e => {
    var workbook = getWorkbookFromData(e.target.result);
    addTextInfos(`<b>Sheets dans le fichier Excel</b>: ${workbook.SheetNames}`);
    var contents=to_json(workbook, [MAIN_SHEET_NAME]);
    var contents_main_sheet=removeTrailingEmptyRows(contents[MAIN_SHEET_NAME]);

    var cut_sheet=parseInfoFromMainSheet(contents_main_sheet);
    var csv_string=convertToCSV(cut_sheet);
    console.log(csv_string)
    download("fichier.csv", csv_string);
    console.log(cut_sheet);
  }

  r.readAsBinaryString(f);
}


function convertToCSV(obj) {
  var csv="";
  COLUMNS.forEach(col => csv+=col+"; ");
  csv+="\n";

  // CAREFUL, WILL BUG IF NOT SAME LENGTH FOR EACH ARRAY
  for(i=0; i<obj[COLUMNS[0]].length; i++){
    COLUMNS.forEach(function(col){
      if(obj[col][i] != null){
        csv+=obj[col][i]+"; ";
      }else{
        csv+="; ";
      }
    });
    csv+="\n";
  }

  return csv;
}

function getWorkbookFromData(data) {
  var workbook = XLSX.read(data, {
      type: 'binary'
      });
  //var firstSheet = workbook.SheetNames[0];
  //var data = to_json(workbook);
  return workbook
};


function to_json(workbook, sheet_names) {
  var result = {};
  sheet_names.forEach(function(sheetName) {
    var roa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1
    });
    if (roa.length) result[sheetName] = roa;
  });
  //return JSON.stringify(result, 2, 2);
  return result;
};

function importFile(evt) {
  var f = file.files[0];

  if (f) {
    var r = new FileReader();
    r.onload = e => {
      var contents = processExcel(e.target.result);
      console.log(contents)
    }

    r.readAsBinaryString(f);
  } else {
  }
};


function parseInfoFromMainSheet(sheet){
  var cut_sheet=sheet.slice(FIRST_ROW_MAIN_SHEET);
  cut_sheet[0]=Array.from(cut_sheet[0], item => typeof item === 'undefined' ? "" : item);
  cut_sheet[1]=Array.from(cut_sheet[1], item => typeof item === 'undefined' ? "" : item);
  console.log(cut_sheet);
  //COLUMNS_REGEX
  var data={};
  console.log("Starting to parse the columns")
  for(i=0; i<COLUMNS_REGEX.length; i++){
    var col=COLUMNS_REGEX[i];
    console.log(col.name)

    var j=findColNum(cut_sheet, col.value_r1, col.value_r2);
  
    if(j<0){
      addTextInfos(`<span class="error_info">Colonne ${col.name} non trouvée </span>`);
      data[col.name]=[]
    }else{
      data[col.name]=getCol(cut_sheet, j).slice(2);
    }
    console.log(col.name)
  }
  

  return data;
}

function findColNum(arr, value_r1, value_r2){
  var j = 0;
  while(j<arr[0].length && !(value_r1.test(arr[0][j]) && (value_r2.test(arr[1][j]) || arr[1][j]==null))){
    j++;
  }

  if(j>=arr[0].length){
    return -1;
  }else{
    return j;
  }
}


function getCol(matrix, j){
  var column = [];
  for(var i=0; i<matrix.length; i++){
     column.push(matrix[i][j]);
  }
  return column;
}



function download(filename, text) {
  var element = document.createElement('a');
  element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
  element.setAttribute('download', filename);

  element.style.display = 'none';
  document.body.appendChild(element);

  element.click();

  document.body.removeChild(element);
}


function removeTrailingEmptyRows(arr){
  var i=arr.length-1;
  while(i>0 && arr[i].length==0){
    i--;
  }

  return arr.slice(0, i);
}