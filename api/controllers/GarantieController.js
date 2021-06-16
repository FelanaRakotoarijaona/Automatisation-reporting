/**
 * GarantieController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

module.exports = {
    accueilGarantie : async function(req,res)
    {
      return res.view('Garantie/accueilreportingGarantie');
    },
    essaiGarantie : function(req,res)
    {
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
      //var table = ['/dev/prod/Retour_Easytech_'];
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
      var date = annee+mois+jour;
      console.log(date);
      var cheminp = [];
      var MotCle= [];
      var nomtable = [];
      var numligne = [];
      var numfeuille = [];
      var nomcolonne = [];
      var colonnecible2 = [];
      var essai = 'essai';
      workbook.xlsx.readFile('htp.xlsx')
      //workbook.xlsx.readFile('ex.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil1');
            var numFeuille = newworksheet.getColumn(4);
            var nomColonne = newworksheet.getColumn(5);
            var nomTable = newworksheet.getColumn(6);
            var cible2 = newworksheet.getColumn(7);
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
            cible2.eachCell(function(cell, rowNumber) {
              colonnecible2.push(cell.value);
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
                      Reportinghtp.deleteFromChemin(table,cb);
                    },
                  function(cb){
                      Reportinghtp.deleteFromChemin2(table,cb);
                    },
                 function(cb){
                      Reportinghtp.importEssai(table,cheminp,date,MotCle,0,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],colonnecible2[0],cb);
                    },
                    function(cb){
                      Reportinghtp.importEssai(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],colonnecible2[1],cb);
                    },
                    function(cb){
                      Reportinghtp.importEssai(table,cheminp,date,MotCle,2,nomtable[2],numligne[2],numfeuille[2],nomcolonne[2],colonnecible2[2],cb);
                    },
                    function(cb){
                      Reportinghtp.importEssai(table,cheminp,date,MotCle,3,nomtable[3],numligne[3],numfeuille[3],nomcolonne[3],colonnecible2[3],cb);
                    },
                    function(cb){
                      Reportinghtp.importEssaitype2(table,cheminp,date,MotCle,4,nomtable[4],numligne[4],numfeuille[4],nomcolonne[4],colonnecible2[4],cb);
                    },
                    function(cb){
                      Reportinghtp.existenceRoute(essai,cb);
                      },
                    function(cb){
                      Reportinghtp.existenceRoute2(essai,cb);
                      },
              ],
              function(err, resultat){
                let val = resultat[8].rows;
                let val2 = resultat[7].rows;
  
                var f = parseInt(val[0].ok) + parseInt(val2[0].ok);
                console.log(val[0].ok);
                if (err) { return res.view('reporting/erreur'); }
                if(f==0)
                {
                  return res.view('reporting/erreur');
                }
                else
                {
                  return res.view('reporting/accueil', {date : datetest});
                }
            });
  
  
          });
    },
};

