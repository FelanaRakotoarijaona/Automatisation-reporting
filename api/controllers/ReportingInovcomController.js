/**
 * ReportingInovcomController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

const ReportingInovcom = require('../models/ReportingInovcom');

module.exports = {
  
    accueil1 : function(req,res)
    {
      return res.view('Inovcom/accueil1');
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
      console.log(date);
      var cheminp = [];
      var MotCle= [];
      workbook.xlsx.readFile('Inovcom.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil1');
            var cheminparticulier = newworksheet.getColumn(9);
            var motcle = newworksheet.getColumn(10);
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
                      ReportingInovcom.deleteFromChemin(table,cb);
                    },
                 function(cb){
                      ReportingInovcom.importEssai(table,cheminp,date,MotCle,0,cb);
                    },
                 function(cb){
                      ReportingInovcom.importEssai(table,cheminp,date,MotCle,1,cb);
                    },
                 function(cb){
                      ReportingInovcom.importEssai(table,cheminp,date,MotCle,2,cb);
                    },
                  /*function(cb){
                      Reportinghtp.importEssai(table,cheminp,date,MotCle,3,cb);
                    },
                    function(cb){
                      Reportinghtp.importEssai(table,cheminp,date,MotCle,4,cb);
                    },*/
              ],
              function(err, resultat){
                if (err) { return res.view('Inovcom/erreur'); }
                else
                {
                  return res.view('Inovcom/accueil', {date : datetest});
                }
            });
          });
    },
    accueil : function(req,res)
    {
      return res.view('Inovcom/accueil');
    },
    EssaiExcel : function(req,res)
    {
      var sql= 'select * from chemininovcom limit 3;';
      Reportinghtp.query(sql,function(err, nc) {
        if (err){
          console.log(err);
          return next(err);
        }
        else
        {
            nc = nc.rows;
            sails.log(nc[0].typologiedelademande);
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
            var nb = 3;
            workbook.xlsx.readFile('Inovcom.xlsx')
                .then(function() {
                  var newworksheet = workbook.getWorksheet('Feuil1');
                  var chemincommun = newworksheet.getColumn(1);
                  var cheminparticulier = newworksheet.getColumn(2);
                  var dernierligne = newworksheet.getColumn(3);
                  var feuille = newworksheet.getColumn(4);
                  var cel = newworksheet.getColumn(5);
                  var tab = newworksheet.getColumn(6);
                  var cel2 = newworksheet.getColumn(7);
                  var numeroligne = newworksheet.getColumn(8);
                    chemincommun.eachCell(function(cell, rowNumber) {
                      cheminc.push(cell.value);
                    });
                    cheminparticulier.eachCell(function(cell, rowNumber) {
                      cheminp.push(cell.value);
                    });
                    dernierligne.eachCell(function(cell, rowNumber) {
                      dernierl.push(cell.value);
                    });
                    feuille.eachCell(function(cell, rowNumber) {
                      feuil.push(cell.value);
                    });
                    cel.eachCell(function(cell, rowNumber) {
                      cellule.push(cell.value);
                    });
                    cel2.eachCell(function(cell, rowNumber) {
                      cellule2.push(cell.value);
                    });
                    tab.eachCell(function(cell, rowNumber) {
                      table.push(cell.value);
                    });
                    numeroligne.eachCell(function(cell, rowNumber) {
                        numligne.push(cell.value);
                      });
                    for(var i=0;i<nb;i++)
                    {
                      var a = cheminc[i]+date+cheminp[i]+nc[i].typologiedelademande;
                      trameflux.push(a);
                    };
                    console.log(trameflux);
                    async.series([
                      function(cb){
                        ReportingInovcom.deleteHtp(table,nb,cb);
                      }, 
                      function(cb){
                        ReportingInovcom.importTrameFlux929(trameflux,feuil,cellule,table,cellule2,nb,numligne,cb);
                      }, 
                      /*function(cb){
                            ReportingInovcom.deleteTout(table,nb,cb);
                          }, 
                       function(cb){
                          ReportingInovcom.deleteHtp(table,nb,cb);
                        }, 
                     function(cb){
                          ReportingInovcom.importInovcom(trameflux,feuil,cellule,table,cellule2,numligne,nb,cb);
                          },
                     function(cb){
                        ReportingInovcom.importTout(trameflux,table,nb,cb);
                        }, */
                    ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                  })
                });
        }
    })
    },
};

