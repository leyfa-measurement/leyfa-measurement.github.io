
var MAIN_SHEET_NAME="Étude";
var FIRST_ROW_MAIN_SHEET=4;
var SC_START_ROW=14;
var SC_H_COL=5;
var SC_L_COL=4;
var SC_NOM_COL=0;

let ANNOTATION_SEP=" , ";


// Informations on the Eléments tracé

let r_orp_th=/orp th à b?[+-]?[0-9]+\s?\+?\s?[0-9]*[,.]?[0-9]*\s?m?/gi;
let r_orp_th_dist=/orp th à b?[+-]?[0-9]+\s?\+?\s?([0-9]*[,.]?[0-9]*)\s?m?/gi;
let r_orp_th_txt="ORP TH";

let r_frp_th=/frp th à b?[+-]?[0-9]+\s?\+?\s?[0-9]*[,.]?[0-9]*\s?m?/gi;
let r_frp_th_dist=/frp th à b?[+-]?[0-9]+\s?\+?\s?([0-9]*[,.]?[0-9]*)\s?m?/gi;
let r_frp_th_txt="FRP TH";

let r_orpf=/orp\s?f? à b?[+-]?[0-9]+\s?\+?\s?[0-9]*[,.]?[0-9]*\s?m?/gi;
let r_orpf_dist=/orp\s?f? à b?[+-]?[0-9]+\s?\+?\s?([0-9]*[,.]?[0-9]*)\s?m?/gi;
let orpf_txt="ORP TH";

let r_frpf=/frp\s?f? à b?[+-]?[0-9]+\s?\+?\s?[0-9]*[,.]?[0-9]*\s?m?/gi;
let r_frpf_dist=/frp\s?f? à b?[+-]?[0-9]+\s?\+?\s?([0-9]*[,.]?[0-9]*)\s?m?/gi;
let frpf_txt="FRP TH";

let r_orp_pr=/orp\s?pr?\s?à?\s?b[-]?[0-9]+\s?\+?\s?[0-9]*[,.]?[0-9]*\s?m?/gi;
let r_orp_pr_dist=/orp\s?pr?\s?à?\s?b[-]?[0-9]+\s?\+?\s?([0-9]*[,.]?[0-9]*)\s?m?/gi;
let r_orp_pr_txt="ORP PR";

let r_frp_pr=/frp\s?pr?\s?à?\s?b[-]?[0-9]+\s?\+?\s?[0-9]*[,.]?[0-9]*\s?m?/gi;
let r_frp_pr_dist=/frp\s?pr?\s?à?\s?b[-]?[0-9]+\s?\+?\s?([0-9]*[,.]?[0-9]*)\s?m?/gi;
let frp_pr_txt="FRP PR";

let r_pi=/\bpi\b à b?[+-]?[0-9]+\s?\+?\s?[0-9]*[,.]?[0-9]*\s?m?/gi;
let r_pi_dist=/\bpi\b à b?[+-]?[0-9]+\s?\+?\s?([0-9]*[,.]?[0-9]*)\s?m?/gi;
let pi_txt="PI";

let ELT_TRACE_REGEX=[
  {
    "g_reg":r_orp_th,
    "dist_reg":r_orp_th_dist,
    "txt":r_orp_th_txt
  },
  {
    "g_reg":r_frp_th,
    "dist_reg":r_frp_th_dist,
    "txt":r_frp_th_txt
  },
  {
    "g_reg":r_orpf,
    "dist_reg":r_orpf_dist,
    "txt":orpf_txt
  },
  {
    "g_reg":r_frpf,
    "dist_reg":r_frpf_dist,
    "txt":frpf_txt
  },
  {
    "g_reg":r_orp_pr,
    "dist_reg":r_orp_pr_dist,
    "txt":r_orp_pr_txt
  },
  {
    "g_reg":r_frp_pr,
    "dist_reg":r_frp_pr_dist,
    "txt":frp_pr_txt
  },
  {
    "g_reg":r_pi,
    "dist_reg":r_pi_dist,
    "txt":pi_txt
  }
]

//Informations on the Points fixes
let r_pt_fixes=/\+ [0-9]*[.]?[0-9]+\s?m [^+]*/g;
let r_pt_fixes_dist=/\+ ([0-9]*[.]?[0-9]+)\s?m [^+]*/g;
let r_pt_fixes_txt=/\+ [0-9]*[.]?[0-9]+\s?m ([^+]*)/g;







let COLUMNS=["Bornes", "Flèche Intiale", "Flèche proposée", "Ripage", "Dévers initial", "Dévers proposé",
"PK", "ENT G", "ENT D", "CAT1", "CAT2", "CAT3", "CAT MA", "Annotations"];

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


submit_file.addEventListener('click', convertFile);
get_info_file.addEventListener('click', showInfoFile);


function returnEltsTrace(sheet_data){
  var elt_trace_raw=sheet_data["Elements tracé"];
  var elt_trace_parsed=Array(elt_trace_raw.length).fill("");

  

  for(i=0; i<elt_trace_raw.length; i++){
    var s="";
    ELT_TRACE_REGEX.forEach(reg =>{
      var matches=elt_trace_raw[i].match(reg.g_reg);
      if(matches != null){
        matches.forEach(m =>{

          reg.dist_reg.lastIndex=0;
          var dist_res=reg.dist_reg.exec(m);
          var dist=null;
          if(dist_res != null){
            dist=dist_res[1];
          }
          if(dist==null || dist==""){
            dist="0";
          }
  
          dist=dist.replace(",", ".");
  
          s+=reg.txt +" -- "+ dist + " m " + ANNOTATION_SEP +" ";
        });
      }
    });
    elt_trace_parsed[i]=s;

  }
  return elt_trace_parsed;
}

function returnPtsFixes(sheet_data){
  var pts_fixes_raw=sheet_data["Points Fixes"];
  var pts_fixes_parsed=Array(pts_fixes_raw.length).fill("");
  

  for(i=0; i<pts_fixes_raw.length; i++){
    var s="";

    var pt_fixes=pts_fixes_raw[i].match(r_pt_fixes);

    if(pt_fixes != null){
      pt_fixes.forEach(p =>{
        r_pt_fixes_dist.lastIndex = 0;
        var dist_res=r_pt_fixes_dist.exec(p);
        var dist=null;
        if(dist_res != null){
          dist=dist_res[1];
        }
        if(dist==null || dist==""){
          dist="0";
        }
        
        //dist=dist.replace(",", "."); //Useless as of now

        r_pt_fixes_txt.lastIndex=0
        var txt_res=r_pt_fixes_txt.exec(p);
        var txt=null;
        if(txt_res != null){
          txt=txt_res[1];
        }
        if(dist!=null){
          s+=txt +" -- "+ dist + " m " + ANNOTATION_SEP +" ";
        }
      })
    }

    pts_fixes_parsed[i]=s;
  }
  return pts_fixes_parsed;
}

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
    parseEpureFromFile(f, false);
    
    
  }else{
    addTextInfos(`<span class="error_info">Pas de fichier détecté.</span>`);
  }
}


function convertFile(evt){
  changeTextInfos("");
  showInfos();
  if(hasFile()){
    var f = file.files[0];
    addTextInfos(`Calcul des infos pour le fichier ${f.name}`);
    parseEpureFromFile(f, true);
    
    
  }else{
    addTextInfos(`<span class="error_info">Pas de fichier détecté.</span>`);
  }
}

function parseEpureFromFile(f, auto_download){
  var r = new FileReader();
  r.onload = e => {
    var workbook = getWorkbookFromData(e.target.result);
    addTextInfos(`<b>Sheets dans le fichier Excel</b>: ${workbook.SheetNames}`);
    var contents=to_json(workbook, [MAIN_SHEET_NAME]);
    var contents_main_sheet=removeTrailingEmptyRows(contents[MAIN_SHEET_NAME]);

    var sheet_data=parseInfoFromMainSheet(contents_main_sheet);

    var elt_trace_parsed=returnEltsTrace(sheet_data);
    var pts_fixes_parsed=returnPtsFixes(sheet_data);

    var l_annot=Math.max(elt_trace_parsed.length, pts_fixes_parsed.length);

    var annotations=Array(l_annot).fill("");

    for(i=0; i<l_annot; i++){
      var elt_trace="";
      var pt_fixe="";

      if(i<elt_trace_parsed.length && elt_trace_parsed[i] != null){
        elt_trace=elt_trace_parsed[i];
      }

      if(i<pts_fixes_parsed.length && pts_fixes_parsed[i] != null){
        pt_fixe=pts_fixes_parsed[i];
      }

      var s=elt_trace+pt_fixe;
      annotations[i]=s;
    }

    sheet_data["Annotations"]=annotations;
    console.log(sheet_data)

    
    if(auto_download){
      var csv_string=convertToCSV(sheet_data);
      download(`${f.name.replace(/\.[^/.]+$/, ".csv")}`, csv_string)
    }
    
    //console.log(csv_string)
    //download("fichier.csv", csv_string);
    //console.log(cut_sheet);
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

    var j=findColNum(cut_sheet, col.value_r1, col.value_r2);
  
    if(j<0){
      addTextInfos(`<span class="error_info">Colonne ${col.name} non trouvée </span>`);
      data[col.name]=[]
    }else{
      data[col.name]=getCol(cut_sheet, j).slice(2);
    }
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