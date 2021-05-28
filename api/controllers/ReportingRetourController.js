/**
 * ReportingRetourController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

const ReportingRetour = require('../models/ReportingRetour');
module.exports = {
    accueil1 : function(req,res)
    {
      return res.view('Retour/accueil1');
    },
    Essaii : function(req,res)
    {
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
      var date = annee+mois+jour;
      var nomtable = [];
      var numligne = [];
      var numfeuille = [];
      var nomcolonne = [];
      console.log(date);
      var cheminp = [];
      var MotCle= [];
      workbook.xlsx.readFile('ReportingRetour.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil2');
            var numFeuille = newworksheet.getColumn(4);
            var nomColonne = newworksheet.getColumn(5);
            var nomTable = newworksheet.getColumn(6);
            var numLigne = newworksheet.getColumn(8);
            var cheminparticulier = newworksheet.getColumn(9);
            var motcle = newworksheet.getColumn(10);
              numFeuille.eachCell(function(cell, rowNumber) {
                numfeuille.push(cell.value);
              });
              nomColonne.eachCell(function(cell, rowNumber) {
                nomcolonne.push(cell.value);
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
              console.log(cheminp[0]);
              console.log(MotCle[0]);
              async.series([  
                  function(cb){
                      ReportingRetour.deleteFromChemin(table,cb);
                    },
                 function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,0,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],cb);
                    },
                 function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],cb);
                    },
                  function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,2,nomtable[2],numligne[2],numfeuille[2],nomcolonne[2],cb);
                    },
                  function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,3,nomtable[3],numligne[3],numfeuille[3],nomcolonne[3],cb);
                    },
                  function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,4,nomtable[4],numligne[4],numfeuille[4],nomcolonne[4],cb);
                    },
                  function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,5,nomtable[5],numligne[5],numfeuille[5],nomcolonne[5],cb);
                    },
                  function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,6,nomtable[6],numligne[6],numfeuille[6],nomcolonne[6],cb);
                    },
                  function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,7,nomtable[7],numligne[7],numfeuille[7],nomcolonne[7],cb);
                    },
                    function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,8,nomtable[8],numligne[8],numfeuille[8],nomcolonne[8],cb);
                    },
                    function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,9,nomtable[9],numligne[9],numfeuille[9],nomcolonne[9],cb);
                    },
                 function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,10,nomtable[10],numligne[10],numfeuille[10],nomcolonne[10],cb);
                    },
                  function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,11,nomtable[11],numligne[11],numfeuille[11],nomcolonne[11],cb);
                    },
                  function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,12,nomtable[12],numligne[12],numfeuille[12],nomcolonne[12],cb);
                    },
                    function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,13,nomtable[13],numligne[13],numfeuille[13],nomcolonne[13],cb);
                    },
                    function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,14,nomtable[14],numligne[14],numfeuille[14],nomcolonne[14],cb);
                    },
                  function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,15,nomtable[15],numligne[15],numfeuille[15],nomcolonne[15],cb);
                    },
                  function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,16,nomtable[16],numligne[16],numfeuille[16],nomcolonne[16],cb);
                    },
                    function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,17,nomtable[17],numligne[17],numfeuille[17],nomcolonne[17],cb);
                    },
                    function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,18,nomtable[18],numligne[18],numfeuille[18],nomcolonne[18],cb);
                    },
                    function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,19,nomtable[19],numligne[19],numfeuille[19],nomcolonne[19],cb);
                    },
                    function(cb){
                      ReportingRetour.importEssai(table,cheminp,date,MotCle,20,nomtable[20],numligne[20],numfeuille[20],nomcolonne[20],cb);
                    },
              ],
              function(err, resultat){
                if (err) { return res.view('Retour/erreur'); }
                else
                {
                  return res.view('Retour/accueil', {date : datetest});
                }
            });
          });
    },
    accueil : function(req,res)
    {
      return res.view('Retour/accueil');
    },
    EssaiExcel : function(req,res)
    {
      var sql1= 'select count(*) as nb from cheminretourvrai;';
      Reportinghtp.query(sql1,function(err, nc1) {
        if (err){
          console.log(err);
          return next(err);
        }
        else
        {
          nc1 = nc1.rows;
          var nbs = nc1[0].nb;
          var x = parseInt(nbs);
          var sql='select * from cheminretourvrai limit' + " " + x ;
          Reportinghtp.query(sql,function(err, nc) {
            if (err){
              console.log(err);
              return next(err);
            }
            else
            {
            nc = nc.rows;
            sails.log(nc[0].chemin);
            var feuil = [];
            var cellule = [];
            var cellule2 = [];
            var table = [];
            var trameflux = [];
            var numligne = [];
            var nb = x;
            for(var i=0;i<nb;i++)
            {
              var a = nc[i].chemin;
              trameflux.push(a);
            };
            for(var i=0;i<nb;i++)
            {
              var a = nc[i].numfeuile;
              feuil.push(a);
            };
            for(var i=0;i<nb;i++)
            {
              var a = nc[i].numligne;
              numligne.push(a);
            };
            for(var i=0;i<nb;i++)
            {
              var a = nc[i].colonnecible;
              cellule.push(a);
            };
            for(var i=0;i<nb;i++)
            {
              var a = nc[i].colonnecible;
              cellule2.push(a);
            };
            for(var i=0;i<nb;i++)
            {
              var a = nc[i].nomtable;
              table.push(a);
            };
                    console.log(table);
                    async.series([
                      function(cb){
                        ReportingRetour.deleteHtp(table,nb,cb);
                      }, 
                      function(cb){
                        ReportingRetour.importTrameFlux929type2(trameflux,feuil,cellule,table,cellule2,nb,numligne,cb);
                      }, 
                    ],
                    function(err, resultat){
                      if (err) { return res.view('Retour/erreur'); }
                      return res.view('Retour/exportExcel');
                  })
             }
             })
        }
    });
  },

};

