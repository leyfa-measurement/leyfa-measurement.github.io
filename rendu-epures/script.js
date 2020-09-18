
var MAIN_SHEET_NAME="Étude";
var IMPLANTATION_SHEET_NAME="Implantation";
var FIRST_ROW_MAIN_SHEET=4;
var SC_HEADER_ROW=0;
var SC_H_COL=5;
var SC_L_COL=4;
var SC_COTE_COL=2;
var SC_NOM_COL=0;
var INCLUDE_CAT_MA=false;

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

let r_fin_quais=/\+ [0-9]*[.]?[0-9]+\s?m [^+]*FIN QUAI[^+]*/i;
let r_debut_quais=/\+ [0-9]*[.]?[0-9]+\s?m [^+]*DEBUT QUAI[^+]*/i;

let r_sc=/\+\s?[0-9]*[.]?[0-9]+\s?m\s+SC\s?[0-9]+\/[a-zA-Z0-9]+/i;
let r_sc_pk=/\+\s?[0-9]*[.]?[0-9]+\s?m\s+SC\s?([0-9]+)\/[a-zA-Z0-9]+/i;
let r_sc_num=/\+\s?[0-9]*[.]?[0-9]+\s?m\s+SC\s?[0-9]+\/([a-zA-Z0-9]+)/i;

let r_only_sc=/SC\s?[0-9]+\/[a-zA-Z0-9]+/i;
let r_only_sc_pk=/SC\s?([0-9]+)\/[a-zA-Z0-9]+/i;
let r_only_sc_num=/SC\s?[0-9]+\/([a-zA-Z0-9]+)/i;



let COLUMNS=["Bornes", "Flèche Intiale", "Flèche proposée", "Ripage", "Dévers initial", "Dévers proposé",
"PK", "ENT G", "ENT D", "CAT1", "CAT2", "CAT3", "CAT MA", "Annotations"];


let COLUMNS_NUMERIC=["Flèche Intiale", "Flèche proposée", "Ripage", "Dévers initial", "Dévers proposé",
"PK", "ENT G", "ENT D", "CAT1", "CAT2", "CAT3", "CAT MA"];

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
},
{
  "name": "Distance quai gauche",
  "value_r1": /Distance finale quai/,
  "value_r2": /gauche/
},
{
  "name": "Distance quai droit",
  "value_r1": /Distance finale quai/,
  "value_r2": /droite/
},
{
  "name": "Hauteur quai gauche",
  "value_r1": /Hauteur finale quai/,
  "value_r2": /gauche/
},
{
  "name": "Hauteur quai droit",
  "value_r1": /Hauteur finale quai/,
  "value_r2": /droite/
}];





var el_MAIN_SHEET_NAME;
var el_IMPLANTATION_SHEET_NAME;
var el_FIRST_ROW_MAIN_SHEET;
var el_SC_HEADER_ROW;
var el_SC_H_COL;
var el_SC_L_COL;
var el_SC_COTE_COL;
var el_SC_NOM_COL;
var el_INCLUDE_CAT_MA;

var file;
var infos_p;
var submit_file;
var get_info_file;

window.onload = function (){
  file = document.getElementById('file');
  infos_p = document.getElementById('infos');
  submit_file = document.getElementById('convert_file');
  //get_info_file = document.getElementById('get_info_file');

  el_show_options = document.getElementById('show_options');
  
  el_MAIN_SHEET_NAME=document.getElementById("MAIN_SHEET_NAME");
  el_IMPLANTATION_SHEET_NAME=document.getElementById("IMPLANTATION_SHEET_NAME");
  el_FIRST_ROW_MAIN_SHEET=document.getElementById("FIRST_ROW_MAIN_SHEET");
  el_SC_HEADER_ROW=document.getElementById("SC_HEADER_ROW");
  el_SC_H_COL=document.getElementById("SC_H_COL");
  el_SC_L_COL=document.getElementById("SC_L_COL");
  el_SC_COTE_COL=document.getElementById("SC_COTE_COL");
  el_SC_NOM_COL=document.getElementById("SC_NOM_COL");
  el_INCLUDE_CAT_MA=document.getElementById("INCLUDE_CAT_MA");

  el_MAIN_SHEET_NAME.value="Étude";
  el_IMPLANTATION_SHEET_NAME.value="Implantation";
  el_FIRST_ROW_MAIN_SHEET.value="4";
  el_SC_HEADER_ROW.value="0";
  el_SC_H_COL.value="5";
  el_SC_L_COL.value="4";
  el_SC_COTE_COL.value="2";
  el_SC_NOM_COL.value="0";
  el_INCLUDE_CAT_MA.checked=INCLUDE_CAT_MA;

  el_show_options.checked=false;

  
  submit_file.addEventListener('click', convertFile);
  //get_info_file.addEventListener('click', showInfoFile);

}










function parseInfosFromForm(){
  MAIN_SHEET_NAME=el_MAIN_SHEET_NAME.value;
  IMPLANTATION_SHEET_NAME=el_IMPLANTATION_SHEET_NAME.value;

  FIRST_ROW_MAIN_SHEET=Math.trunc(parseFloat(String(el_FIRST_ROW_MAIN_SHEET.value).replace(",", ".")));
  SC_HEADER_ROW=Math.trunc(parseFloat(String(el_SC_HEADER_ROW.value).replace(",", ".")));
  SC_H_COL=Math.trunc(parseFloat(String(el_SC_H_COL.value).replace(",", ".")));
  SC_L_COL=Math.trunc(parseFloat(String(el_SC_L_COL.value).replace(",", ".")));
  SC_COTE_COL=Math.trunc(parseFloat(String(el_SC_COTE_COL.value).replace(",", ".")));
  SC_NOM_COL=Math.trunc(parseFloat(String(el_SC_NOM_COL.value).replace(",", ".")));

  INCLUDE_CAT_MA=el_INCLUDE_CAT_MA.checked;
}


function returnEltsTrace(sheet_data){
  var elt_trace_raw=sheet_data["Elements tracé"];
  var elt_trace_parsed=Array(elt_trace_raw.length).fill("");

  

  for(i=0; i<elt_trace_raw.length; i++){
    var s="";
    ELT_TRACE_REGEX.forEach(reg =>{
      var matches=null;
      if(elt_trace_raw[i]!=null){
        matches=elt_trace_raw[i].match(reg.g_reg);
      }
      
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

function returnPtsFixes(sheet_data, contents_implantation_sheet){
  var pts_fixes_raw=sheet_data["Points Fixes"];
  var pts_fixes_parsed=Array(pts_fixes_raw.length).fill("");

  var debut_quais_ind=[];
  var fin_quais_ind=[];

  var sc=[];
  

  for(i=0; i<pts_fixes_raw.length; i++){
    var s="";
    var pt_fixes=null;
    if(pts_fixes_raw[i] != null){
      pt_fixes=pts_fixes_raw[i].match(r_pt_fixes);
    }
    

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
        
        var is_fin_quais=r_fin_quais.test(p);
        var is_debut_quais=r_debut_quais.test(p);
        var is_sc=r_sc.test(p);

        if(is_debut_quais){
          debut_quais_ind.push({"ind":i, "dist":dist});
        } else if(is_fin_quais){
          fin_quais_ind.push({"ind":i, "dist":dist});
        }else if(is_sc){
          var pk=r_sc_pk.exec(p)[1];
          var num=r_sc_num.exec(p)[1];
          sc.push({"ind":i,
          "dist":dist,
          "pk":pk,
          "num":num
          });

        }else{
          r_pt_fixes_txt.lastIndex=0
          var txt_res=r_pt_fixes_txt.exec(p);
          var txt=null;
          if(txt_res != null){
            txt=txt_res[1];
          }
          if(txt!=null){
            s+=txt +" -- "+ dist + " m " + ANNOTATION_SEP +" ";
          }
        }

        
      })
    }

    pts_fixes_parsed[i]=s;
  }


  //Now we handle the quais

  if(debut_quais_ind.length!=fin_quais_ind.length){
    addTextInfos(`<span class="error_info">Il n'y a pas le même nombre de début de quais et de fin de quais. Les quais seront ignorés. #Début de quais: ${debut_quais_ind.length}, #Fin de quais: ${fin_quais_ind.length}</span>`);
  }else{
    var distance_quai_gauche = sheet_data["Distance quai gauche"];
    var distance_quai_droit = sheet_data["Distance quai droit"];
    var hauteur_quai_gauche = sheet_data["Hauteur quai gauche"];
    var hauteur_quai_droit = sheet_data["Hauteur quai droit"];

    addTextInfos(`${debut_quais_ind.length} quai(s) repéré(s).`);
    for(i=0; i<debut_quais_ind.length;i++){
      if(debut_quais_ind[i].ind<=fin_quais_ind[i].ind){
        var deb=debut_quais_ind[i];
        var fin=fin_quais_ind[i];
      }else{
        var deb=fin_quais_ind[i];
        var fin=debut_quais_ind[i];
      }

      

      //Let's determine if the quais are right or left
      var n_info_droite=0;
      var n_info_gauche=0;
      var n_non_null_droite=0;
      var n_non_null_gauche=0;
      var quais_cote="DROIT";
       
      for(j=deb.ind; j<=fin.ind; j++){
        var d_q_g=distance_quai_gauche[j];
        var d_q_d=distance_quai_droit[j];
        var h_q_g=hauteur_quai_gauche[j];
        var h_q_d=hauteur_quai_droit[j];
        [d_q_g, h_q_g].forEach(e =>{
          if(e!="" && e != null && e!=0 && e!="0"){
            n_info_gauche++;
          }

          if(e!="" && e != null){
            n_non_null_gauche++;
          }
        });

        [d_q_d, h_q_d].forEach(e =>{
          if(e!="" && e != null && e!=0 && e!="0"){
            n_info_droite++;
          }
          if(e!="" && e != null){
            n_non_null_droite++;
          }
        });
      }

      

      if(n_info_droite==n_info_gauche){
        if(n_non_null_droite==n_non_null_gauche){
          addTextInfos(`<span class="error_info">L'algorithme n'a pas réussi à déterminer le côté du quai, par défaut, ce dernier sera mis à droite.</span>`);
        }else if(n_non_null_droite>n_non_null_gauche){
          quais_cote="DROIT";
        }else{
          quais_cote="GAUCHE";
        }
        
      }else if(n_info_droite>n_info_gauche){
        quais_cote="DROIT";
      }else{
        quais_cote="GAUCHE";
      }


      var h_deb="0";
      var d_deb="0";
      var h_fin="0";
      var d_fin="0";

      if(quais_cote=="DROIT"){
        h_deb=hauteur_quai_droit[deb.ind];
        d_deb=distance_quai_droit[deb.ind];
        h_fin=hauteur_quai_droit[fin.ind];
        d_fin=distance_quai_droit[fin.ind];
      }else{
        h_deb=hauteur_quai_gauche[deb.ind];
        d_deb=distance_quai_gauche[deb.ind];
        h_fin=hauteur_quai_gauche[fin.ind];
        d_fin=distance_quai_gauche[fin.ind];
      }

      if(h_deb=="" || h_deb==null){
        h_deb="0";
      }

      if(d_deb=="" || d_deb==null){
        d_deb="0";
      }

      if(h_fin=="" || h_fin==null){
        h_fin="0";
      }

      if(d_fin=="" || d_fin==null){
        d_fin="0";
      }


      
    
      h_deb=Math.round(parseFloat(String(h_deb).replace(",", ".")));
      d_deb=Math.round(parseFloat(String(d_deb).replace(",", ".")));
      h_fin=Math.round(parseFloat(String(h_fin).replace(",", ".")));
      d_fin=Math.round(parseFloat(String(d_fin).replace(",", ".")));
      pts_fixes_parsed[deb.ind]+=`DEBUT QUAIS -- ${deb.dist} -- ${d_deb} -- ${h_deb} -- ${quais_cote} ${ANNOTATION_SEP} `;
      pts_fixes_parsed[fin.ind]+=`FIN QUAIS -- ${fin.dist} -- ${d_fin} -- ${h_fin} -- ${quais_cote} ${ANNOTATION_SEP} `;

      for(j=deb.ind+1; j<fin.ind; j++){
        if(quais_cote=="DROIT"){
          var h=hauteur_quai_droit[j];
          var d=distance_quai_droit[j];
        }else{
          var h=hauteur_quai_gauche[j];
          var d=distance_quai_gauche[j];
        }

        if(h=="" || h==null){
          h="0";
        }
  
        if(d=="" || d==null){
          d="0";
        }

        h=Math.round(parseFloat(String(h).replace(",", ".")));
        d=Math.round(parseFloat(String(d).replace(",", ".")));

        pts_fixes_parsed[j]+=`QUAIS -- ${0} -- ${d} -- ${h} -- ${quais_cote} ${ANNOTATION_SEP} `;
      }
    }
    
  }


  //Now we handle the SC
  addTextInfos(`${sc.length} support(s) caténaire repéré(s).`);

  if(contents_implantation_sheet!=null){
    sc_data=extractSCImplantationData(contents_implantation_sheet);

    sc.forEach(e=>{
      sc_info=getInfoForSC(e.num, e.pk, sc_data);
      if(sc_info==null){
        addTextInfos(`<span class="error_info">Implantation non trouvée pour le SC${e.pk}/${e.num}.</span>`);
      }else{
        var l=Math.round(parseFloat(String(sc_info.l).replace(",", ".")));
        var h=Math.round(parseFloat(String(sc_info.h).replace(",", ".")));
        pts_fixes_parsed[e.ind]+=`Cote Implantation SC${e.pk}/${e.num} -- ${e.dist} -- ${l} -- ${h} -- ${sc_info.cote} ${ANNOTATION_SEP} `;

      }
    })

    console.log("Données de support caténaires:")
    console.log(sc_data);
  }



  return pts_fixes_parsed;
}

function getInfoForSC(num, pk, sc_data){
  for(i=0; i<sc_data.length;i++){
    if(String(sc_data[i].num)==String(num) && String(sc_data[i].pk)==String(pk)){
      return sc_data[i];
    }
  }
  return null;
}

function extractSCImplantationData(contents_implantation_sheet){
  var implantation_data={};
  implantation_data["nom"]=getCol(contents_implantation_sheet, SC_NOM_COL).slice(SC_HEADER_ROW+1);
  implantation_data["h"]=getCol(contents_implantation_sheet, SC_H_COL).slice(SC_HEADER_ROW+1);
  implantation_data["l"]=getCol(contents_implantation_sheet, SC_L_COL).slice(SC_HEADER_ROW+1);
  implantation_data["cote"]=getCol(contents_implantation_sheet, SC_COTE_COL).slice(SC_HEADER_ROW+1);

  var sc_data=[];

  for(i=0; i<implantation_data.nom.length; i++){
    if(r_only_sc.test(implantation_data.nom[i])){
      var pk=r_only_sc_pk.exec(implantation_data.nom[i])[1];
      var num=r_only_sc_num.exec(implantation_data.nom[i])[1];
      var h=implantation_data.h[i];
      var l=implantation_data.l[i];
      var cote=implantation_data.cote[i];

      if(/gauche/i.test(cote)){
        cote="GAUCHE";
      }else{
        cote="DROIT";
      }

      sc_data.push({"pk":pk,
      "num":num,
      "h":h,
      "l":l,
      "cote":cote});
    }


  }

  return sc_data;

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
  parseInfosFromForm();
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
  var has_implantation_sheet=true;
  r.onload = e => {
    var workbook = getWorkbookFromData(e.target.result);
    addTextInfos(`<b>Sheets dans le fichier Excel</b>: ${workbook.SheetNames}`);

    if(!workbook.SheetNames.includes(MAIN_SHEET_NAME)){
      addTextInfos(`<span class="error_info">Feuille principale "${MAIN_SHEET_NAME}" non trouvée.</span>`);
      return null
    }

    if(!workbook.SheetNames.includes(IMPLANTATION_SHEET_NAME)){
      addTextInfos(`<span class="error_info">Feuille d'implantation "${IMPLANTATION_SHEET_NAME}" non trouvée.</span>`);
      has_implantation_sheet=false;
    }

    var sheets_to_parse=[MAIN_SHEET_NAME];
    if(has_implantation_sheet){
      sheets_to_parse.push(IMPLANTATION_SHEET_NAME);
    }

    var contents=to_json(workbook, [MAIN_SHEET_NAME, IMPLANTATION_SHEET_NAME]);

    var contents_main_sheet=removeTrailingEmptyRows(contents[MAIN_SHEET_NAME]);

    var contents_implantation_sheet;
    
    if(has_implantation_sheet){
      contents_implantation_sheet=removeTrailingEmptyRows(contents[IMPLANTATION_SHEET_NAME]);
    }
    

    console.log("Informations de la feuille d'implantation:")
    console.log(contents_implantation_sheet);
    var sheet_data=parseInfoFromMainSheet(contents_main_sheet);

    console.log("Informations de la feuille principale:")
    console.log(sheet_data);

    var elt_trace_parsed=returnEltsTrace(sheet_data);
    var pts_fixes_parsed=returnPtsFixes(sheet_data, contents_implantation_sheet);

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

    //We replace all "" with 0
    cleanFinalData(sheet_data);

    console.log("Informations de la feuille principale:")
    console.log(sheet_data)

    
    if(auto_download){
      var csv_string=convertToCSV(sheet_data);
      csv_string="R2D1.v1 ;;;;;;;;;;;;;\ndu PK048+668 au PK049+878;;;;;;;;;;;;;\n"+csv_string;
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
    }

    r.readAsBinaryString(f);
  } else {
  }
};


function parseInfoFromMainSheet(sheet){
  var cut_sheet=sheet.slice(FIRST_ROW_MAIN_SHEET);
  cut_sheet[0]=Array.from(cut_sheet[0], item => typeof item === 'undefined' ? "" : item);
  cut_sheet[1]=Array.from(cut_sheet[1], item => typeof item === 'undefined' ? "" : item);
  console.log("Informations de la feuille principale:")
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
  while(i>=0 && arr[i].length==0){
    i--;
  }

  return arr.slice(0, i+1);
}

function cleanFinalData(sheet_data){

  //We fill the unfilled columns
  var max_length=0;
  COLUMNS.forEach(col=>{
    max_length=Math.max(max_length, sheet_data[col].length);
  })

  COLUMNS.forEach(col=>{
    for(i=sheet_data[col].length; i<max_length; i++){
      sheet_data[col].push("");
    }
  })


  //We remove all non numeric chars in numeric fields and replace empty strings with 0
  COLUMNS_NUMERIC.forEach(col =>{
    for(i=0; i<sheet_data[col].length; i++){
      sheet_data[col][i]=String(sheet_data[col][i]).replace(",", ".").replace(/[^0-9.]/g, "");
      if(sheet_data[col][i]=="" || sheet_data[col][i]=="undefined" ||  sheet_data[col][i]==null){
        sheet_data[col][i]="0";
      }
    }
    
  });

  //We remove all lines that don't have a correct Bornes

  var to_remove=[];
  

  for(i=0; i<sheet_data["Bornes"].length; i++){
    if((!/B?[0-9]+/i.test(String(sheet_data["Bornes"][i]))) || sheet_data["Bornes"][i]==null){
      to_remove.push(i);
    }
  }
  if(to_remove.length==1){
    addTextInfos(`<span class="error_info">La ligne ${to_remove} n'a pas de borne valide et sera donc ignorée. </span>`);
  }else if(to_remove.length>1){
    addTextInfos(`<span class="error_info">Les lignes ${to_remove} n'ont pas de borne valide et seront donc ignorées. </span>`);
  }
  

  for(i=to_remove.length-1; i>=0; i--){
    Object.keys(sheet_data).forEach(col =>{
      sheet_data[col].splice(to_remove[i], 1);
    })
  }

  if(!INCLUDE_CAT_MA){
    sheet_data["CAT MA"].fill("0");
  }
}