/**
 * GarantieController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

const Garantie = require('../models/Garantie');
module.exports = {
  //routes vers la page de recherche du fichier
    accueilGarantie : async function(req,res)
    {
      return res.view('Garantie/accueilreportingGarantiesansdouble');
    },
    //INSERTION CHEMIN SANS DOUBLON
    insertChemingarantiesansdouble : function(req,res)
    {
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
      var date = annee+mois+jour;
      var date_indus = jour+'.'+mois+'.'+annee;
      var datej_1 = annee+mois+jour -1;
      console.log(datej_1);
      var nomtable = [];
      var numligne = [];
      var numfeuille = [];
      var nomcolonne = [];
      var nomcolonne2 = [];
      var nomcolonne3 = [];
      console.log(date);
      var cheminp = [];
      var MotCle= [];
      var Sup= [];
      var nomBase = "chemingarantiesansdouble";
      workbook.xlsx.readFile('Garantie.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil1');
            var nomColonne3 = newworksheet.getColumn(3);
            var numFeuille = newworksheet.getColumn(4);
            var nomColonne = newworksheet.getColumn(5);
            var nomTable = newworksheet.getColumn(6);
            var nomColonne2 = newworksheet.getColumn(7);
            var numLigne = newworksheet.getColumn(8);
            var cheminparticulier = newworksheet.getColumn(9);
            var motcle = newworksheet.getColumn(10);
            var suppleant = newworksheet.getColumn(11);
           /************************************************************/
           var newworksheet_1 = workbook.getWorksheet('Feuil2');
           var nomColonne3_1 = newworksheet_1.getColumn(3);
           var numFeuille_1 = newworksheet_1.getColumn(4);
           var nomColonne_1 = newworksheet_1.getColumn(5);
           var nomTable_1 = newworksheet_1.getColumn(6);
           var nomColonne2_1 = newworksheet_1.getColumn(7);
           var numLigne_1 = newworksheet_1.getColumn(8);
           var cheminparticulier_1 = newworksheet_1.getColumn(9);
           var motcle_1 = newworksheet_1.getColumn(10);

          /************************************************************/
          numFeuille.eachCell(function(cell, rowNumber) {
            numfeuille.push(cell.value);
          });
          nomColonne.eachCell(function(cell, rowNumber) {
            nomcolonne.push(cell.value);
          });
          nomColonne2.eachCell(function(cell, rowNumber) {
            nomcolonne2.push(cell.value);
          });
          nomColonne3.eachCell(function(cell, rowNumber) {
            nomcolonne3.push(cell.value);
          });
          nomTable.eachCell(function(cell, rowNumber) {
            nomtable.push(cell.value);
          });
          numLigne.eachCell(function(cell, rowNumber) {
            numligne.push(cell.value);
          });
            cheminparticulier.eachCell(function(cell, rowNumber) {
              cheminp.push(cell.value);
          });
            motcle.eachCell(function(cell, rowNumber) {
              MotCle.push(cell.value);
          });
          suppleant.eachCell(function(cell, rowNumber) {
            Sup.push(cell.value);
          });


          /************************************************************/
            numFeuille_1.eachCell(function(cell, rowNumber) {
              numfeuille.push(cell.value);
            });
            nomColonne_1.eachCell(function(cell, rowNumber) {
              nomcolonne.push(cell.value);
            });
            nomColonne2_1.eachCell(function(cell, rowNumber) {
              nomcolonne2.push(cell.value);
            });
            nomColonne3_1.eachCell(function(cell, rowNumber) {
              nomcolonne3.push(cell.value);
            });
            nomTable_1.eachCell(function(cell, rowNumber) {
              nomtable.push(cell.value);
            });
            numLigne_1.eachCell(function(cell, rowNumber) {
              numligne.push(cell.value);
            });
              cheminparticulier_1.eachCell(function(cell, rowNumber) {
                cheminp.push(cell.value);
            });
              motcle_1.eachCell(function(cell, rowNumber) {
                MotCle.push(cell.value);
            });
            /************************************************************/
            
              console.log(cheminp[0]);
              console.log(MotCle[0]);
              async.series([  
                  function(cb){
                      Garantie.deleteFromChemin(table,cb);
                    },
                    function(cb){
                      Garantie.importEssaidemat(table,cheminp,date,MotCle,0,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],nomcolonne2[0],nomcolonne3[0],cb);
                    },
                    // function(cb){
                    //   Garantie.importEssaidemat1(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaiavis(table,cheminp,date,MotCle,2,nomtable[2],numligne[2],numfeuille[2],nomcolonne[2],nomcolonne2[2],nomcolonne3[2],Sup[2],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaiavis(table,cheminp,date,MotCle,3,nomtable[3],numligne[3],numfeuille[3],nomcolonne[3],nomcolonne2[3],nomcolonne3[3],Sup[3],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,4,nomtable[4],numligne[4],numfeuille[4],nomcolonne[4],nomcolonne2[4],nomcolonne3[4],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,5,nomtable[5],numligne[5],numfeuille[5],nomcolonne[5],nomcolonne2[5],nomcolonne3[5],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,6,nomtable[6],numligne[6],numfeuille[6],nomcolonne[6],nomcolonne2[6],nomcolonne3[6],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,7,nomtable[7],numligne[7],numfeuille[7],nomcolonne[7],nomcolonne2[7],nomcolonne3[7],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,8,nomtable[8],numligne[8],numfeuille[8],nomcolonne[8],nomcolonne2[8],nomcolonne3[8],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,9,nomtable[9],numligne[9],numfeuille[9],nomcolonne[9],nomcolonne2[9],nomcolonne3[9],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,10,nomtable[10],numligne[10],numfeuille[10],nomcolonne[10],nomcolonne2[10],nomcolonne3[10],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,11,nomtable[11],numligne[11],numfeuille[11],nomcolonne[11],nomcolonne2[11],nomcolonne3[11],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,12,nomtable[12],numligne[12],numfeuille[12],nomcolonne[12],nomcolonne2[12],nomcolonne3[12],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,13,nomtable[13],numligne[13],numfeuille[13],nomcolonne[13],nomcolonne2[13],nomcolonne3[13],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,14,nomtable[14],numligne[14],numfeuille[14],nomcolonne[14],nomcolonne2[14],nomcolonne3[14],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,15,nomtable[15],numligne[15],numfeuille[15],nomcolonne[15],nomcolonne2[15],nomcolonne3[15],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,16,nomtable[16],numligne[16],numfeuille[16],nomcolonne[16],nomcolonne2[16],nomcolonne3[16],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,17,nomtable[17],numligne[17],numfeuille[17],nomcolonne[17],nomcolonne2[17],nomcolonne3[17],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,18,nomtable[18],numligne[18],numfeuille[18],nomcolonne[18],nomcolonne2[18],nomcolonne3[18],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaiencindus(table,cheminp,date,MotCle,19,nomtable[19],numligne[19],numfeuille[19],nomcolonne[19],nomcolonne2[19],nomcolonne3[19],Sup[19],date_indus,cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaiencindus(table,cheminp,date,MotCle,20,nomtable[20],numligne[20],numfeuille[20],nomcolonne[20],nomcolonne2[20],nomcolonne3[20],Sup[20],date_indus,cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,21,nomtable[21],numligne[21],numfeuille[21],nomcolonne[21],nomcolonne2[21],nomcolonne3[21],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,22,nomtable[22],numligne[22],numfeuille[22],nomcolonne[22],nomcolonne2[22],nomcolonne3[22],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,23,nomtable[23],numligne[23],numfeuille[23],nomcolonne[23],nomcolonne2[23],nomcolonne3[23],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,24,nomtable[24],numligne[24],numfeuille[24],nomcolonne[24],nomcolonne2[24],nomcolonne3[24],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,25,nomtable[25],numligne[25],numfeuille[25],nomcolonne[25],nomcolonne2[25],nomcolonne3[25],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematretention(table,cheminp,date,MotCle,26,nomtable[26],numligne[26],numfeuille[26],nomcolonne[26],nomcolonne2[26],nomcolonne3[26],datej_1,cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematretention(table,cheminp,date,MotCle,27,nomtable[27],numligne[27],numfeuille[27],nomcolonne[27],nomcolonne2[27],nomcolonne3[27],datej_1,cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematretention(table,cheminp,date,MotCle,28,nomtable[28],numligne[28],numfeuille[28],nomcolonne[28],nomcolonne2[28],nomcolonne3[28],datej_1,cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematretention(table,cheminp,date,MotCle,29,nomtable[29],numligne[29],numfeuille[29],nomcolonne[29],nomcolonne2[29],nomcolonne3[29],datej_1,cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,30,nomtable[30],numligne[30],numfeuille[30],nomcolonne[30],nomcolonne2[30],nomcolonne3[30],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,31,nomtable[31],numligne[31],numfeuille[31],nomcolonne[31],nomcolonne2[31],nomcolonne3[31],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,32,nomtable[32],numligne[32],numfeuille[32],nomcolonne[32],nomcolonne2[32],nomcolonne3[32],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,33,nomtable[33],numligne[33],numfeuille[33],nomcolonne[33],nomcolonne2[33],nomcolonne3[33],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,34,nomtable[34],numligne[34],numfeuille[34],nomcolonne[34],nomcolonne2[34],nomcolonne3[34],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,35,nomtable[35],numligne[35],numfeuille[35],nomcolonne[35],nomcolonne2[35],nomcolonne3[35],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,36,nomtable[36],numligne[36],numfeuille[36],nomcolonne[36],nomcolonne2[36],nomcolonne3[36],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,37,nomtable[37],numligne[37],numfeuille[37],nomcolonne[37],nomcolonne2[37],nomcolonne3[37],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,38,nomtable[38],numligne[38],numfeuille[38],nomcolonne[38],nomcolonne2[38],nomcolonne3[38],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,39,nomtable[39],numligne[39],numfeuille[39],nomcolonne[39],nomcolonne2[39],nomcolonne3[39],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,40,nomtable[40],numligne[40],numfeuille[40],nomcolonne[40],nomcolonne2[40],nomcolonne3[40],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,41,nomtable[41],numligne[41],numfeuille[41],nomcolonne[41],nomcolonne2[41],nomcolonne3[41],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,42,nomtable[42],numligne[42],numfeuille[42],nomcolonne[42],nomcolonne2[42],nomcolonne3[42],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,43,nomtable[43],numligne[43],numfeuille[43],nomcolonne[43],nomcolonne2[43],nomcolonne3[43],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,44,nomtable[44],numligne[44],numfeuille[44],nomcolonne[44],nomcolonne2[44],nomcolonne3[44],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,45,nomtable[45],numligne[45],numfeuille[45],nomcolonne[45],nomcolonne2[45],nomcolonne3[45],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,46,nomtable[46],numligne[46],numfeuille[46],nomcolonne[46],nomcolonne2[46],nomcolonne3[46],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,47,nomtable[47],numligne[47],numfeuille[47],nomcolonne[47],nomcolonne2[47],nomcolonne3[47],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,48,nomtable[48],numligne[48],numfeuille[48],nomcolonne[48],nomcolonne2[48],nomcolonne3[48],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,49,nomtable[49],numligne[49],numfeuille[49],nomcolonne[49],nomcolonne2[49],nomcolonne3[49],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,50,nomtable[50],numligne[50],numfeuille[50],nomcolonne[50],nomcolonne2[50],nomcolonne3[50],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,51,nomtable[51],numligne[51],numfeuille[51],nomcolonne[51],nomcolonne2[51],nomcolonne3[51],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematretention(table,cheminp,date,MotCle,52,nomtable[52],numligne[52],numfeuille[52],nomcolonne[52],nomcolonne2[52],nomcolonne3[52],datej_1,cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,53,nomtable[53],numligne[53],numfeuille[53],nomcolonne[53],nomcolonne2[53],nomcolonne3[53],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,54,nomtable[54],numligne[54],numfeuille[54],nomcolonne[54],nomcolonne2[54],nomcolonne3[54],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,55,nomtable[55],numligne[55],numfeuille[55],nomcolonne[55],nomcolonne2[55],nomcolonne3[55],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,56,nomtable[56],numligne[56],numfeuille[56],nomcolonne[56],nomcolonne2[56],nomcolonne3[56],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematconvention(table,cheminp,date,MotCle,57,nomtable[57],numligne[57],numfeuille[57],nomcolonne[57],nomcolonne2[57],nomcolonne3[57],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematbpo(table_1,cheminp,date,MotCle,58,nomtable[58],numligne[58],numfeuille[58],nomcolonne[58],nomcolonne2[58],nomcolonne3[58],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematbpo(table_1,cheminp,date,MotCle,59,nomtable[59],numligne[59],numfeuille[59],nomcolonne[59],nomcolonne2[59],nomcolonne3[59],cb);
                    // },
                    // function(cb){
                    //   Garantie.importEssaidematbpo(table_1,cheminp,date,MotCle,60,nomtable[60],numligne[60],numfeuille[60],nomcolonne[60],nomcolonne2[60],nomcolonne3[60],cb);
                    // },
                    
                    
              ],
              function(err, resultat){
                var sql4= "select count(chemin) as ok from "+nomBase+" ";
                        console.log(sql4);
                        Reportinghtp.getDatastore().sendNativeQuery(sql4 ,function(err, nc) {
                           nc = nc.rows;
                           console.log(nc);
                           console.log('nc'+nc[0].ok);
                           var f = parseInt(nc[0].ok);
                           console.log(f);
                              if (err){
                                return res.view('Inovcom/erreur');
                              }
                            if(f==0)
                              {
                                return res.view('Inovcom/erreur');
                              }
                              else
                              {
                                return res.view('Garantie/accueilreportingGarantieligne', {date : datetest});
                                                              
                              };
                          });
            });
          });
    },
/*****************************************************************************************************/
  //INSERTION CHEMIN NOMBRE DE LIGNE
  insertChemingarantieligne : function(req,res)
  {
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
    // var table = ['/dev/pro/Retour_Easytech_'];
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var date = annee+mois+jour;
    var date_indus = jour+'.'+mois+'.'+annee;
    var datej_1 = annee+mois+jour -1;
    console.log(datej_1);
    var nomtable = [];
    var numligne = [];
    var numfeuille = [];
    var nomcolonne = [];
    var nomcolonne2 = [];
    var nomcolonne3 = [];
    console.log(date);
    var cheminp = [];
    var MotCle= [];
    var Sup= [];
    var nomBase = "chemingarantieligne";
    workbook.xlsx.readFile('Garantie.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil1');
          var nomColonne3 = newworksheet.getColumn(3);
          var numFeuille = newworksheet.getColumn(4);
          var nomColonne = newworksheet.getColumn(5);
          var nomTable = newworksheet.getColumn(6);
          var nomColonne2 = newworksheet.getColumn(7);
          var numLigne = newworksheet.getColumn(8);
          var cheminparticulier = newworksheet.getColumn(9);
          var motcle = newworksheet.getColumn(10);
          var suppleant = newworksheet.getColumn(11);
       
        numFeuille.eachCell(function(cell, rowNumber) {
          numfeuille.push(cell.value);
        });
        nomColonne.eachCell(function(cell, rowNumber) {
          nomcolonne.push(cell.value);
        });
        nomColonne2.eachCell(function(cell, rowNumber) {
          nomcolonne2.push(cell.value);
        });
        nomColonne3.eachCell(function(cell, rowNumber) {
          nomcolonne3.push(cell.value);
        });
        nomTable.eachCell(function(cell, rowNumber) {
          nomtable.push(cell.value);
        });
        numLigne.eachCell(function(cell, rowNumber) {
          numligne.push(cell.value);
        });
          cheminparticulier.eachCell(function(cell, rowNumber) {
            cheminp.push(cell.value);
        });
          motcle.eachCell(function(cell, rowNumber) {
            MotCle.push(cell.value);
        });
        suppleant.eachCell(function(cell, rowNumber) {
          Sup.push(cell.value);
        });


          
            console.log(cheminp[0]);
            console.log(MotCle[0]);
            async.series([  
                function(cb){
                    Garantie.deleteFromChemin_1(table,cb);
                  },
                  function(cb){
                    Garantie.importEssaidemat1(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],cb);
                  },
                                   
            ],
            function(err, resultat){
              var sql4= "select count(chemin) as ok from "+nomBase+" ";
                      console.log(sql4);
                      Reportinghtp.getDatastore().sendNativeQuery(sql4 ,function(err, nc) {
                         nc = nc.rows;
                         console.log(nc);
                         console.log('nc'+nc[0].ok);
                         var f = parseInt(nc[0].ok);
                         console.log(f);
                            if (err){
                              return res.view('Inovcom/erreur');
                            }
                          if(f==0)
                            {
                              return res.view('Inovcom/erreur');
                            }
                            else
                            {
                              return res.view('Garantie/accueilreportingGarantiebpo1', {date : datetest});
                              
                            };
                        });
          });
        });
  },
/****************************************************************************************************/
insertChemingarantiebpo1 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  var table = ['\\\\10.128.1.2\\bpo_almerys'];
  // var table = ['/dev/prod/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "chemingarantiebpo";
  workbook.xlsx.readFile('Garantie.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil2');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

      numFeuille.eachCell(function(cell, rowNumber) {
        numfeuille.push(cell.value);
      });
      nomColonne.eachCell(function(cell, rowNumber) {
        nomcolonne.push(cell.value);
      });
      nomColonne2.eachCell(function(cell, rowNumber) {
        nomcolonne2.push(cell.value);
      });
      nomColonne3.eachCell(function(cell, rowNumber) {
        nomcolonne3.push(cell.value);
      });
      nomTable.eachCell(function(cell, rowNumber) {
        nomtable.push(cell.value);
      });
      numLigne.eachCell(function(cell, rowNumber) {
        numligne.push(cell.value);
      });
        cheminparticulier.eachCell(function(cell, rowNumber) {
          cheminp.push(cell.value);
      });
        motcle.eachCell(function(cell, rowNumber) {
          MotCle.push(cell.value);
      });
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          async.series([  
              function(cb){
                  Garantie.deleteFromChemin_bpo1(table,cb);
                },
           function(cb){
                  Garantie.importEssaidematbpo(table,cheminp,date,MotCle,0,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],nomcolonne2[0],nomcolonne3[0],cb);
                },
                
          ],
          function(err, resultat){
            var sql4= "select count(chemin) as ok from "+nomBase+" ";
                    console.log(sql4);
                    Reportinghtp.getDatastore().sendNativeQuery(sql4 ,function(err, nc) {
                       nc = nc.rows;
                       console.log(nc);
                       console.log('nc'+nc[0].ok);
                       var f = parseInt(nc[0].ok);
                       console.log(f);
                          if (err){
                            return res.view('Inovcom/erreur');
                          }
                        if(f==0)
                          {
                            return res.view('Inovcom/erreur');
                          }
                          else
                          {
                            return res.view('Garantie/importGarantiesansdouble', {date : datetest});
                                                          
                          };
                      });
        });
      });
},
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importGarantiesansdouble : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from chemingarantiesansdouble';
      Reportinghtp.getDatastore().sendNativeQuery(sql1,function(err, nc1) {
        if (err){
          console.log(err);
          return next(err);
        }
        else
        {
          nc1 = nc1.rows;
          var nbs = nc1[0].nb;
          var x = parseInt(nbs);
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from chemingarantiesansdouble limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
            nc = nc.rows;
            sails.log(nc[0].chemin);
            var Excel = require('exceljs');
            var workbook = new Excel.Workbook();
            var cheminc = [];
            var cheminp = [];
            var dernierl = [];
            var feuil = [];
            var cellule = [];
            var cellule2 = [];
            var table = [];
            var trameflux = [];
            var numligne = [];
            var datetest = req.param("date",0);
            var annee = datetest.substr(0, 4);
            var mois = datetest.substr(5, 2);
            var jour = datetest.substr(8, 2);
            var date = annee+mois+jour;
            var dateexport = jour + '/' + mois + '/' +annee;
            var nb = x;
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
                      feuil.push(a);                      
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a =nc[i].colonnecible;
                      cellule.push(a);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a =nc[i].nomtable;
                      table.push(a);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a =nc[i].colonnecible2;
                      cellule2.push(a);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a =nc[i].numligne;
                      numligne.push(a);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a =nc[i].colonnecible3;
                      dernierl.push(a);
                    };
                    console.log(trameflux);
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                        function(cb){
                          Garantie.deleteHtp(table,nb,cb);
                       },
                        function(cb){
                          Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      return res.view('Garantie/importGarantieligne', {date : datetest});
                  });
               
              }
          })
        };
      });
  },
/***************************************************************************************************************/
importGarantieligne : function(req,res)
{
  var datetest = req.param("date",0);
  var sql1= 'select count(*) as nb from chemingarantieligne';
    Reportinghtp.getDatastore().sendNativeQuery(sql1,function(err, nc1) {
      if (err){
        console.log(err);
        return next(err);
      }
      else
      {
        nc1 = nc1.rows;
        var nbs = nc1[0].nb;
        var x = parseInt(nbs);
        //var sql='select * from cheminretourvrai limit' + " " + x ;
        var sql= 'select * from chemingarantieligne limit'  + " " + x;
        Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
          console.log(nc);
          if (err){
            console.log('ato am if erreur');
            console.log(err);
            return next(err);
          }
          else
          {
            console.log('ato amm else');              
          nc = nc.rows;
          sails.log(nc[0].chemin);
          var Excel = require('exceljs');
          var workbook = new Excel.Workbook();
          var cheminc = [];
          var cheminp = [];
          var dernierl = [];
          var feuil = [];
          var cellule = [];
          var cellule2 = [];
          var table = [];
          var trameflux = [];
          var numligne = [];
          var datetest = req.param("date",0);
          var annee = datetest.substr(0, 4);
          var mois = datetest.substr(5, 2);
          var jour = datetest.substr(8, 2);
          var date = annee+mois+jour;
          var dateexport = jour + '/' + mois + '/' +annee;
          var nb = x;
          var nbre = [];
                  for(var i=0;i<nb;i++)
                  {
                    var a = nc[i].chemin;
                    trameflux.push(a);
                    nbre.push(i);
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a = nc[i].numfeuille;
                    feuil.push(a);                      
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].colonnecible;
                    cellule.push(a);
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].nomtable;
                    table.push(a);
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].colonnecible2;
                    cellule2.push(a);
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].numligne;
                    numligne.push(a);
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].colonnecible3;
                    dernierl.push(a);
                  };
                  console.log(trameflux);
                  async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                    async.series([
                      function(cb){
                        Garantie.deleteHtp(table,nb,cb);
                     },
                      function(cb){
                        Garantie.importTrameDemat_1(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,cb);
                      },
                    ],function(erroned, lotValues){
                      if(erroned) return res.badRequest(erroned);
                      return callback_reporting_suivant();
                    });
                  },
                  // async.series([
                  //   function(cb){
                  //     Garantie.deleteHtp(table,nb,cb);
                  //   }, 
                  //   function(cb){
                  //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                  //   }, 
                  //   // function(cb){
                  //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                  //   // }, 
                  // ],
                  function(err, resultat){
                    if (err) { return res.view('Inovcom/erreur'); }
                    return res.view('Garantie/importGarantiebpo1', {date : datetest});
                });
             
              }
          })
        };
    });
},
/*****************************************************************************************************************/
importGarantiebpo1 : function(req,res)
{
  var datetest = req.param("date",0);
  var sql1= 'select count(*) as nb from chemingarantiebpo';
    Reportinghtp.getDatastore().sendNativeQuery(sql1,function(err, nc1) {
      if (err){
        console.log(err);
        return next(err);
      }
      else
      {
        nc1 = nc1.rows;
        var nbs = nc1[0].nb;
        var x = parseInt(nbs);
        //var sql='select * from cheminretourvrai limit' + " " + x ;
        var sql= 'select * from chemingarantiebpo limit'  + " " + x;
        Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
          console.log(nc);
          if (err){
            console.log('ato am if erreur');
            console.log(err);
            return next(err);
          }
          else
          {
            console.log('ato amm else');              
          nc = nc.rows;
          sails.log(nc[0].chemin);
          var Excel = require('exceljs');
          var workbook = new Excel.Workbook();
          var cheminc = [];
          var cheminp = [];
          var dernierl = [];
          var feuil = [];
          var cellule = [];
          var cellule2 = [];
          var table = [];
          var trameflux = [];
          var numligne = [];
          var datetest = req.param("date",0);
          var annee = datetest.substr(0, 4);
          var mois = datetest.substr(5, 2);
          var jour = datetest.substr(8, 2);
          var date = annee+mois+jour;
          var dateexport = jour + '/' + mois + '/' +annee;
          var nb = x;
          var nbre = [];
                  for(var i=0;i<nb;i++)
                  {
                    var a = nc[i].chemin;
                    trameflux.push(a);
                    nbre.push(i);
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a = nc[i].numfeuille;
                    feuil.push(a);                      
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].colonnecible;
                    cellule.push(a);
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].nomtable;
                    table.push(a);
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].colonnecible2;
                    cellule2.push(a);
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].numligne;
                    numligne.push(a);
                  };
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].colonnecible3;
                    dernierl.push(a);
                  };
                  console.log(trameflux);
                  //BOUCLE
                  async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                    async.series([
                      function(cb){
                        Garantie.deleteHtp(table,nb,cb);
                     },
                      function(cb){
                        Garantie.importTrameDematbpo1(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,cb);
                      },
                    ],function(erroned, lotValues){
                      if(erroned) return res.badRequest(erroned);
                      return callback_reporting_suivant();
                    });
                  },
                  // async.series([
                  //   function(cb){
                  //     Garantie.deleteHtp(table,nb,cb);
                  //   }, 
                  //   function(cb){
                  //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                  //   }, 
                  //   // function(cb){
                  //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                  //   // }, 
                  // ],
                  function(err, resultat){
                    if (err) { return res.view('Inovcom/erreur'); }
                    return res.view('Garantie/exportGarantie', {date : datetest});//ROUTE EXPORT
                });
             
              }
          })
        };
    });
},
   /***********************************************************************************/  
   /*
   *
   *
   *                              EXPORT GARANTIE
   * 
   * 
   *  
  /***********************************************************************************/
  //CONFIGURATION ROUTES PAS ENCORE PRET
exportgarantiefinprod : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var feuille = annee+mois+'_EASY';
  console.log(feuille);
  var mois1 = 'Janvier' ;
  if(mois==01)
  {
    mois1= 'Janvier';
  };
  if(mois==02)
  {
    mois1= 'Fevrier';
  };
  if(mois==03)
  {
    mois1= 'Mars';
  };
  if(mois==04)
  {
    mois1= 'Avril';
  };
  if(mois==05)
  {
    mois1= 'Mai';
  };
  if(mois==06)
  {
    mois1= 'Juin';
  };
  if(mois==07)
  {
    mois1= 'Juillet';
  };
  if(mois==08)
  {
    mois1= 'Aout';
  };
  if(mois==09)
  {
    mois1= 'Septembre';
  };
  if(mois==10)
  {
    mois1= 'octobre';
  };
  if(mois==11)
  {
    mois1= 'Novembre';
  };
  if(mois==12)
  {
    mois1= 'Decembre';
  };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      Garantie.recupdata("garantiedematfinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantieavisfinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantieconvfinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantietpmepfinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantieindusfinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantiecontrfinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantieretenfinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantiercforcefinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantiefraudesfinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantiecodelisfinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantieremisefinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantiearnofinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantiedefraifinp",callback);
    },
    function (callback) {
      Garantie.recupdata("garantiecurefinp",callback);
    },
    function (callback) {
      Garantie.recupdatasum("garantierechfinp",callback);
    },
    function (callback) {
      Garantie.recupdatasum("garantiecmucfinp",callback);
    },
    
   
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Garantie.ecrituredata16tri(result[0],"garantiedematfinp",date_export,feuille,callback);
      },
      function (callback) {
        Garantie.ecrituredata16facM(result[1],"garantieavisfinp",date_export,feuille,callback);
      },
      function (callback) {
        Garantie.ecrituredata16devi(result[2],"garantieconvfinp",date_export,feuille,callback);
      },
      function (callback) {
        Garantie.ecrituredata16sales(result[3],"garantietpmepfinp",date_export,feuille,callback);
      },
      function (callback) {
        Garantie.ecrituredata16flux(result[4],"garantieindusfinp",date_export,feuille,callback);
      },
      function (callback) {
          Garantie.ecrituredata16rejet(result[5],"garantiecontrfinp",date_export,feuille,callback);
        },
      function (callback) {
          Garantie.ecrituredata16cotlamie(result[6],"garantieretenfinp",date_export,feuille,callback);
      },
      function (callback) {
          Garantie.ecrituredatafinptri(result[7],"garantiercforcefinp",date_export,feuille,callback);
        },
        function (callback) {
          Garantie.ecrituredatafinpfacM(result[8],"garantiefraudesfinp",date_export,feuille,callback);
        },
        function (callback) {
          Garantie.ecrituredatafinpdevi(result[9],"garantiecodelisfinp",date_export,feuille,callback);
        },
        function (callback) {
          Garantie.ecrituredatafinpsales(result[10],"garantieremisefinp",date_export,feuille,callback);
        },
        function (callback) {
          Garantie.ecrituredatafinpflux(result[11],"garantiearnofinp",date_export,feuille,callback);
        },
        function (callback) {
            Garantie.ecrituredatafinprejet(result[12],"garantiedefraifinp",date_export,feuille,callback);
          },
        function (callback) {
        Garantie.ecrituredatafinpcotlamie(result[13],"garantiecurefinp",date_export,feuille,callback);
        },
        function (callback) {
          Garantie.ecrituredata16cotite(result[14],"garantierechfinp",date_export,feuille,callback);
        },
        function (callback) {
          Garantie.ecrituredatafinpcotite(result[15],"garantiecmucfinp",date_export,feuille,callback);
        },
        
    
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Retour/erera');
        }
        if(resultExcel[0]=='OK')
        {
          res.view('HTPengagement/exportHTPengagementsuivant_1', {date : datetest});
          // res.view('Retour/succes');
        }

    })
  })
},

/*************************************************************************************/
exportgarantiej2 : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var feuille = annee+mois+'_EASY';
  console.log(feuille);
  var mois1 = 'Janvier' ;
  if(mois==01)
  {
    mois1= 'Janvier';
  };
  if(mois==02)
  {
    mois1= 'Fevrier';
  };
  if(mois==03)
  {
    mois1= 'Mars';
  };
  if(mois==04)
  {
    mois1= 'Avril';
  };
  if(mois==05)
  {
    mois1= 'Mai';
  };
  if(mois==06)
  {
    mois1= 'Juin';
  };
  if(mois==07)
  {
    mois1= 'Juillet';
  };
  if(mois==08)
  {
    mois1= 'Aout';
  };
  if(mois==09)
  {
    mois1= 'Septembre';
  };
  if(mois==10)
  {
    mois1= 'octobre';
  };
  if(mois==11)
  {
    mois1= 'Novembre';
  };
  if(mois==12)
  {
    mois1= 'Decembre';
  };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      Engagementhtp.recupdata("garantiedematj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieavisj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieconvj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantietpmepj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieindusj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiecontrj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieretenj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiercforcej2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiefraudesj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiecodelisj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiermisej2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiearnoj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiedefraij2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiecurej2",callback);
    },
    
   
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredata16tri(result[0],"garantiedematj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16facM(result[1],"garantieavisj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16devi(result[2],"garantieconvj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16sales(result[3],"garantietpmepj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16flux(result[4],"garantieindusj2",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredata16rejet(result[5],"garantiecontrj2",date_export,feuille,callback);
        },
      function (callback) {
      Engagementhtp.ecrituredata16cotlamie(result[6],"garantieretenj2",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatafinptri(result[7],"garantiercforcej2",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpfacM(result[8],"garantiefraudesj2",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpdevi(result[9],"garantiecodelisj2",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpsales(result[10],"garantiermisej2",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpflux(result[11],"garantiearnoj2",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredatafinprejet(result[12],"garantiedefraij2",date_export,feuille,callback);
          },
        function (callback) {
        Engagementhtp.ecrituredatafinpcotlamie(result[13],"garantiecurej2",date_export,feuille,callback);
        },
       
    
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Retour/erera');
        }
        if(resultExcel[0]=='OK')
        {
          res.view('HTPengagement/exportHTPengagementsuivant_1', {date : datetest});
          // res.view('Retour/succes');
        }

    })
  })
},
/*******************************************************************************************/
exportgarantiej5 : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var feuille = annee+mois+'_EASY';
  console.log(feuille);
  var mois1 = 'Janvier' ;
  if(mois==01)
  {
    mois1= 'Janvier';
  };
  if(mois==02)
  {
    mois1= 'Fevrier';
  };
  if(mois==03)
  {
    mois1= 'Mars';
  };
  if(mois==04)
  {
    mois1= 'Avril';
  };
  if(mois==05)
  {
    mois1= 'Mai';
  };
  if(mois==06)
  {
    mois1= 'Juin';
  };
  if(mois==07)
  {
    mois1= 'Juillet';
  };
  if(mois==08)
  {
    mois1= 'Aout';
  };
  if(mois==09)
  {
    mois1= 'Septembre';
  };
  if(mois==10)
  {
    mois1= 'octobre';
  };
  if(mois==11)
  {
    mois1= 'Novembre';
  };
  if(mois==12)
  {
    mois1= 'Decembre';
  };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      Engagementhtp.recupdata("garantiedematj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieavisj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieconvj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantietpmepj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieindusj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiecontrj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieretenj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiercforcej5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiefraudesj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiecodelisj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieremisej5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiearnoj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiedefraij5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiecurej5",callback);
    },
    
   
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredata16tri(result[0],"garantiedematj5",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16facM(result[1],"garantieavisj5",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16devi(result[2],"garantieconvj5",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16sales(result[3],"garantietpmepj5",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16flux(result[4],"garantieindusj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredata16rejet(result[5],"garantiecontrj5",date_export,feuille,callback);
        },
      function (callback) {
      Engagementhtp.ecrituredata16cotlamie(result[6],"garantieretenj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatafinptri(result[7],"garantiercforcej5",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpfacM(result[8],"garantiefraudesj5",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpdevi(result[9],"garantiecodelisj5",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpsales(result[10],"garantieremisej5",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpflux(result[11],"garantiearnoj5",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredatafinprejet(result[12],"garantiedefraij5",date_export,feuille,callback);
          },
        function (callback) {
        Engagementhtp.ecrituredatafinpcotlamie(result[13],"garantiecurej5",date_export,feuille,callback);
        },
        
    
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Retour/erera');
        }
        if(resultExcel[0]=='OK')
        {
          res.view('HTPengagement/exportHTPengagementsuivant_1', {date : datetest});
          // res.view('Retour/succes');
        }

    })
  })
},

/********************************************************************************************/
exportgarantieetpentrants : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var feuille = annee+mois+'_EASY';
  console.log(feuille);
  var mois1 = 'Janvier' ;
  if(mois==01)
  {
    mois1= 'Janvier';
  };
  if(mois==02)
  {
    mois1= 'Fevrier';
  };
  if(mois==03)
  {
    mois1= 'Mars';
  };
  if(mois==04)
  {
    mois1= 'Avril';
  };
  if(mois==05)
  {
    mois1= 'Mai';
  };
  if(mois==06)
  {
    mois1= 'Juin';
  };
  if(mois==07)
  {
    mois1= 'Juillet';
  };
  if(mois==08)
  {
    mois1= 'Aout';
  };
  if(mois==09)
  {
    mois1= 'Septembre';
  };
  if(mois==10)
  {
    mois1= 'octobre';
  };
  if(mois==11)
  {
    mois1= 'Novembre';
  };
  if(mois==12)
  {
    mois1= 'Decembre';
  };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      Engagementhtp.recupdata("garantiedematetp",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieavisetp",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiedematentrants",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieavisentrants",callback);
    },
    
   
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredata16tri(result[0],"garantiedematetp",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16facM(result[1],"garantieavisetp",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16devi(result[2],"garantiedematentrants",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16sales(result[3],"garantieavisentrants",date_export,feuille,callback);
      },
      
    
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Retour/erera');
        }
        if(resultExcel[0]=='OK')
        {
          res.view('HTPengagement/exportHTPengagementsuivant_1', {date : datetest});
          // res.view('Retour/succes');
        }

    })
  })
},
/*********************************************************************************************/
exportgarantietachenont : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var feuille = annee+mois+'_EASY';
  console.log(feuille);
  var mois1 = 'Janvier' ;
  if(mois==01)
  {
    mois1= 'Janvier';
  };
  if(mois==02)
  {
    mois1= 'Fevrier';
  };
  if(mois==03)
  {
    mois1= 'Mars';
  };
  if(mois==04)
  {
    mois1= 'Avril';
  };
  if(mois==05)
  {
    mois1= 'Mai';
  };
  if(mois==06)
  {
    mois1= 'Juin';
  };
  if(mois==07)
  {
    mois1= 'Juillet';
  };
  if(mois==08)
  {
    mois1= 'Aout';
  };
  if(mois==09)
  {
    mois1= 'Septembre';
  };
  if(mois==10)
  {
    mois1= 'octobre';
  };
  if(mois==11)
  {
    mois1= 'Novembre';
  };
  if(mois==12)
  {
    mois1= 'Decembre';
  };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      Engagementhtp.recupdata("garantieconvtnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantietpmeptnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieindustnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiecontrtnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("grantieretentnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiefraudestnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiecodelistnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantieremisetnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiearnotnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiedefraitnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiecuretnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantierechtnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiecmuctnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiercforcetnt",callback);
    },
    function (callback) {
      Engagementhtp.recupdatasum("garantiercforcetntj1",callback);
    },
    function (callback) {
      Engagementhtp.recupdatasum("garantiercforcetntj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("garantiercforcetntj5",callback);
    },
   
   
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredata16tri(result[0],"garantieconvtnt",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16facM(result[1],"garantietpmeptnt",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16devi(result[2],"garantieindustnt",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16sales(result[3],"garantiecontrtnt",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredata16flux(result[4],"grantieretentnt",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredata16rejet(result[5],"garantiefraudestnt",date_export,feuille,callback);
        },
      function (callback) {
      Engagementhtp.ecrituredata16cotlamie(result[6],"garantiecodelistnt",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatafinptri(result[7],"garantieremisetnt",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpfacM(result[8],"garantiearnotnt",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpdevi(result[9],"garantiedefraitnt",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpsales(result[10],"garantiecuretnt",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpflux(result[11],"garantierechtnt",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredatafinprejet(result[12],"garantiecmuctnt",date_export,feuille,callback);
          },
        function (callback) {
        Engagementhtp.ecrituredatafinpcotlamie(result[13],"garantiercforcetnt",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredata16cotite(result[14],"garantiercforcetntj1",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatafinpcotite(result[15],"garantiercforcetntj2",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredata16faclamie(result[16],"garantiercforcetntj5",date_export,feuille,callback);
        },
      
    
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Retour/erera');
        }
        if(resultExcel[0]=='OK')
        {
          res.view('HTPengagement/exportHTPengagementsuivant_1', {date : datetest});
          // res.view('Retour/succes');
        }

    })
  })
},








};

