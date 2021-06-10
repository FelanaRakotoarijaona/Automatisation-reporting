/**
 * ReportingInduController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

const ReportingIndu = require('../models/ReportingIndu');

module.exports = {
    accueil1 : function(req,res)
    {
      return res.view('Indu/accueil1');
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
      var nomcolonne2 = [];
      //console.log(date);
      var cheminp = [];
      var MotCle= [];
      var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19];
      //var r = [0];
      var nomBase = "cheminindu";
      workbook.xlsx.readFile('ReportingIndu.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil2');
            var numFeuille = newworksheet.getColumn(4);
            var nomColonne = newworksheet.getColumn(5);
            var nomTable = newworksheet.getColumn(6);
            var nomColonne2 = newworksheet.getColumn(7);
            var numLigne = newworksheet.getColumn(8);
            var cheminparticulier = newworksheet.getColumn(9);
            var motcle = newworksheet.getColumn(10);
            numFeuille.eachCell(function(cell, rowNumber) {
              numfeuille.push(cell.value);
            });
            nomColonne.eachCell(function(cell, rowNumber) {
              nomcolonne.push(cell.value);
            });
            nomColonne2.eachCell(function(cell, rowNumber) {
              nomcolonne2.push(cell.value);
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
             /* console.log(cheminp[0]);
              console.log(MotCle[0]);*/
              console.log(nomtable);
              async.series([ 
                   function(cb){
                      ReportingInovcom.deleteFromChemin(nomBase,cb);
                    },
              ],
              function(err, resultat){
                if (err) { return res.view('Indu/erreur'); }
                else
                {
                  async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                    async.series([
                      function(cb){
                        ReportingInovcom.delete(nomtable,lot,cb);
                      },
                      function(cb){
                        //console.log(lot + 'numero');
                        ReportingInovcom.importEssai(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomBase,cb);
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
                              return res.view('Indu/erreur');
                            }
                           if(f==0)
                            {
                              return res.view('Indu/erreur');
                            }
                            else
                            {
                              return res.view('Indu/accueil', {date : datetest});
                              
                            };
                        });
                    });
                }
            });
          });
    },
    accueil : function(req,res)
    {
      return res.view('Indu/accueil');
    },
    EssaiExcel : function(req,res)
    {
      var dateFormat = require("dateformat");
      datetest = req.param("date",0);
      var today = new Date(datetest);
      var tomorrow = new Date(today);
      var f = tomorrow.setDate(today.getDate()- 1);
      var date2=dateFormat(f,"shortDate");
      console.log(date2);
      var sql1= 'select count(*) as nb from cheminindu;';
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
      var sql= 'select * from cheminindu limit' + " " + x ;
      Reportinghtp.query(sql,function(err, nc) {
        if (err){
          console.log(err);
          return next(err);
        }
        else
        {
            nc = nc.rows;
            var cheminc = [];
            var cheminp = [];
            var dernierl = [];
            var feuil = [];
            var cellule = [];
            var cellule2 = [];
            var table = [];
            var trameflux = [];
            var numligne = [];
            /*var datetest = req.param("date",0);
            var annee = datetest.substr(0, 4);
            var mois = datetest.substr(5, 2);
            var jour = datetest.substr(8, 2);
            var date = annee+mois+jour;
            var dateexport = jour + '/' + mois + '/' +annee;*/
            var nb = x;
            var nbre = [];
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
              var a = nc[i].colonnecible2;
              cellule2.push(a);
            };
            for(var i=0;i<nb;i++)
            {
              var a = nc[i].nomtable;
              table.push(a);
            };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    console.log(trameflux);
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                        function(cb){
                          ReportingIndu.importTrameFlux929(trameflux,feuil,cellule,table,cellule2,lot,numligne,date2,cb);
                        }, 
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    function(err)
                      {
                        console.log('vofafa ddol');
                        return res.view('Indu/exportExcelIndu');
                        //return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                      });
        };
    })
    };
  });
},

// type 2

   accueiltype2 : function(req,res)
    {
      return res.view('Indu/accueiltype2');
    },
    Essaii2 : function(req,res)
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
      var nomcolonne2 = [];
      console.log('ato tsika');
      //console.log(date);
      var cheminp = [];
      var MotCle= [];
      //var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13];
      //var r = [0,1,2,3,4,5];//,5,6,7,8,9,10,11,12];
      var r = [0];
      var nomBase = "cheminindu2";
      workbook.xlsx.readFile('ReportingIndu.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil4');
            var numFeuille = newworksheet.getColumn(4);
            var nomColonne = newworksheet.getColumn(5);
            var nomTable = newworksheet.getColumn(6);
            var nomColonne2 = newworksheet.getColumn(7);
            var numLigne = newworksheet.getColumn(8);
            var cheminparticulier = newworksheet.getColumn(9);
            var motcle = newworksheet.getColumn(10);
            numFeuille.eachCell(function(cell, rowNumber) {
              numfeuille.push(cell.value);
            });
            nomColonne.eachCell(function(cell, rowNumber) {
              nomcolonne.push(cell.value);
            });
            nomColonne2.eachCell(function(cell, rowNumber) {
              nomcolonne2.push(cell.value);
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
             /* console.log(cheminp[0]);
              console.log(MotCle[0]);*/
              var nomtables = ['indurelevedecomptealmerys','indurelevedecomptecbtp'];
              console.log(nomtable);
              async.series([ 
                   function(cb){
                      ReportingInovcom.deleteFromChemin(nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.delete(nomtables,0,cb);
                    },
                    function(cb){
                      ReportingInovcom.delete(nomtables,1,cb);
                    },
              ],
              function(err, resultat){
                if (err) { return res.view('Indu/erreur'); }
                else
                {
                  async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                    async.series([
                      function(cb){
                        ReportingIndu.importEssaitype7(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,cb);
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
                              return res.view('Indu/erreur');
                            }
                           if(f==0)
                            {
                              return res.view('Indu/erreur');
                            }
                            else
                            {
                              return res.view('Indu/accueil2', {date : datetest});
                              
                            };
                        });
                    });
                }
            });
          });
    },

    accueil2 : function(req,res)
    {
      return res.view('Indu/accueil2');
    },
    EssaiExcel2 : function(req,res)
    {
      var dateFormat = require("dateformat");
      datetest = req.param("date",0);
      var today = new Date(datetest);
      var tomorrow = new Date(today);
      var f = tomorrow.setDate(today.getDate()- 1);
      var date2=dateFormat(f,"shortDate");
      console.log(date2);
      var sql1= 'select count(*) as nb from cheminindu2;';
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
      var sql= 'select * from cheminindu2 limit' + " " + x ;
      Reportinghtp.query(sql,function(err, nc) {
        if (err){
          console.log(err);
          return next(err);
        }
        else
        {
            nc = nc.rows;
            var cheminc = [];
            var cheminp = [];
            var dernierl = [];
            var feuil = [];
            var cellule = [];
            var cellule2 = [];
            var table = [];
            var trameflux = [];
            var numligne = [];
            /*var datetest = req.param("date",0);
            var annee = datetest.substr(0, 4);
            var mois = datetest.substr(5, 2);
            var jour = datetest.substr(8, 2);
            var date = annee+mois+jour;
            var dateexport = jour + '/' + mois + '/' +annee;*/
            var nb = x;
            var nbre = [];
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
              var a = nc[i].colonnecible2;
              cellule2.push(a);
            };
            for(var i=0;i<nb;i++)
            {
              var a = nc[i].nomtable;
              table.push(a);
            };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    console.log(trameflux);
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                        function(cb){
                          ReportingIndu.importTrameFlux9292(trameflux,feuil,cellule,table,cellule2,lot,numligne,cb);
                        }, 
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    function(err)
                      {
                        console.log('vofafa ddol');
                        return res.view('Indu/exportExcelIndu2');
                        //return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                      });
        };
    })
    };
  });
},

};

