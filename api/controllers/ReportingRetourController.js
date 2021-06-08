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
     /*// var a = "~Trame_Tableau retours PEC Optique.xlsx";
      var a = "~Trame_Tableau retours PEC Optique.xls";
      //var m = '.xlsx|.xls|.xlsm|.xlsb$';
      var m = '^[^~]';
      const regex = new RegExp(m);

            if(regex.test(a))
            {
              console.log('ok');
            }
            else
            {
              console.log('ko');
            };*/
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
      var nomBase = "cheminretourvrai";
      var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20];
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
                    ReportingInovcom.deleteFromChemin(nomBase,cb);
                    },
              ],
              function(err, resultat){
                if (err) { return res.view('Retour/erreur'); }
                else
                {
                  async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                    async.series([
                      function(cb){
                        ReportingInovcom.delete(nomtable,lot,cb);
                      },
                      function(cb){
                        ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomBase,cb);
                        //ReportingRetour.importEssai(table,cheminp,date,MotCle,lot,nomtable,numligne[17],numfeuille[17],nomcolonne[17],cb);
                        //ReportingInovcom.importEssai(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomBase,cb);
                      },
                    ],function(erroned, lotValues){
                      if(erroned) return res.badRequest(erroned);
                      return callback_reporting_suivant();
                    });
                  },
                    function(err)
                    {
                      var sql4= "select count(chemin) as ok from "+nomBase+" ";
                      console.log(sql4);
                      Reportinghtp.getDatastore().sendNativeQuery(sql4 ,function(err, nc) {
                         nc = nc.rows;
                         console.log('nc'+nc[0].ok);
                         var f = parseInt(nc[0].ok);
                            if (err){
                              return res.view('reporting/erreur');
                            }
                           if(f==0)
                            {
                              return res.view('reporting/erreur');
                            }
                            else
                            {
                              return res.view('Inovcom/accueil', {date : datetest});
                              
                            };
                        });
                      /*console.log('vofafa ddol');
                      return res.view('Retour/accueil', {date : datetest});*/
                      //return res.view('Inovcom/accueil', {date : datetest});
                    });
                 
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
            var nbre = [];
            for(var i=0;i<nb;i++)
            {
              var a = nc[i].nomtable;
              table.push(a);
              nbre.push(i);
            };
                    console.log(table);
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                        function(cb){
                          ReportingRetour.importTrameFlux929type2(trameflux,feuil,cellule,table,cellule2,lot,numligne,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                      function(err)
                      {
                        console.log('vofafa ddol');
                        return res.view('Retour/exportExcel');
                      }); 
             }
             })
        }
    });
  },

};

