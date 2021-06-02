/**
 * ReportingInovcomController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */
module.exports = {
    accueil1 : async function(req,res)
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
                    ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                  })
                });
        }
    })
    },
//Type 2
    accueil1type2 : function(req,res)
    {
      return res.view('Inovcom/accueil1type2');
    },
    Essaiitype2 : function(req,res)
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
      var nomBase = "chemininovcomtype2";
      workbook.xlsx.readFile('Inovcom.xlsx')
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
                      ReportingInovcom.deleteFromChemin2(table,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,0,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,2,nomtable[2],numligne[2],numfeuille[2],nomcolonne[2],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,3,nomtable[3],numligne[3],numfeuille[3],nomcolonne[3],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,4,nomtable[4],numligne[4],numfeuille[4],nomcolonne[4],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,5,nomtable[5],numligne[5],numfeuille[5],nomcolonne[5],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,6,nomtable[6],numligne[6],numfeuille[6],nomcolonne[6],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,7,nomtable[7],numligne[7],numfeuille[7],nomcolonne[7],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,8,nomtable[8],numligne[8],numfeuille[8],nomcolonne[8],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,9,nomtable[9],numligne[9],numfeuille[9],nomcolonne[9],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,10,nomtable[10],numligne[10],numfeuille[10],nomcolonne[10],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,11,nomtable[11],numligne[11],numfeuille[11],nomcolonne[11],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,12,nomtable[12],numligne[12],numfeuille[12],nomcolonne[12],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,13,nomtable[13],numligne[13],numfeuille[13],nomcolonne[13],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,14,nomtable[14],numligne[14],numfeuille[14],nomcolonne[14],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,15,nomtable[15],numligne[15],numfeuille[15],nomcolonne[15],nomBase,cb);
                    },
                    function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,16,nomtable[16],numligne[16],numfeuille[16],nomcolonne[16],nomBase,cb);
                    },
              ],
              function(err, resultat){
                if (err) { return res.view('Inovcom/erreur'); }
                else
                {
                  return res.view('Inovcom/accueiltype2', {date : datetest});
                }
            });
          });
    },
    accueiltype2 : function(req,res)
    {
      return res.view('Inovcom/accueiltype2');
    },
    EssaiExceltype2 : function(req,res)
    {
      var sql1= 'select count(*) as nb from chemininovcomtype2;';
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
          var sql='select * from chemininovcomtype2 limit' + " " + x ;
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
            var nb = x;
            console.log(nb);
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
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                    };
                    console.log(trameflux);
                    async.series([
                      function(cb){
                        ReportingInovcom.deleteHtp(table,nb,cb);
                      }, 
                      function(cb){
                        ReportingInovcom.importTrameFlux929type2(trameflux,feuil,cellule,table,cellule2,nb,numligne,cb);
                      }, 
                    ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                  });
               
        };
    });
  };
});
    },
    //Type 3
    accueil1type3 : function(req,res)
    {
      return res.view('Inovcom/accueil1type3');
    },
    Essaiitype3 : function(req,res)
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
      var nomBase = "chemininovcomtype3";
      workbook.xlsx.readFile('Inovcom.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil3');
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
                      ReportingInovcom.deleteFromChemin3(table,cb);
                    },
                  function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,0,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],nomBase,cb);
                    },
                  function(cb){
                    ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomBase,cb);
                  },
                  function(cb){
                    ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,2,nomtable[2],numligne[2],numfeuille[2],nomcolonne[2],nomBase,cb);
                  },
              ],
              function(err, resultat){
                if (err) { return res.view('Inovcom/erreur'); }
                else
                {
                  return res.view('Inovcom/accueiltype3', {date : datetest});
                }
            });
          });
    },

    accueiltype3 : function(req,res)
    {
      return res.view('Inovcom/accueiltype3');
    },
    EssaiExceltype3 : function(req,res)
    {
      var sql1= 'select count(*) as nb from chemininovcomtype3;';
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
      var sql= 'select * from chemininovcomtype3 limit' + " " + x;
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
            var date2 = jour + '-' + mois + '-' + annee;
            var dateexport = jour + '/' + mois + '/' +annee;
            var nb = x;
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
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      
                      trameflux.push(a);
                    };
                    console.log(trameflux);
                    async.series([
                      function(cb){
                        ReportingInovcom.deleteHtp(table,nb,cb);
                      }, 
                      function(cb){
                        ReportingInovcom.importTrameFlux929type3(trameflux,feuil,cellule,table,cellule2,nb,numligne,cb);
                      }, 
                    ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                  })
        };
    });
  };
});
    },

    //Type 4
    accueil1type4 : function(req,res)
    {
      return res.view('Inovcom/accueil1type4');
    },
    Essaiitype4 : function(req,res)
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
      var nomBase = "chemininovcomtype4";
      workbook.xlsx.readFile('Inovcom.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil4');
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
                      ReportingInovcom.deleteFromChemin4(table,cb);
                    },
                  function(cb){
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,0,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],nomBase,cb);
                    },
                  function(cb){
                    ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomBase,cb);
                  },
                  function(cb){
                    ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,2,nomtable[2],numligne[2],numfeuille[2],nomcolonne[2],nomBase,cb);
                  },
                  function(cb){
                    ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,3,nomtable[3],numligne[3],numfeuille[3],nomcolonne[3],nomBase,cb);
                  },
                  function(cb){
                    ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,4,nomtable[4],numligne[4],numfeuille[4],nomcolonne[4],nomBase,cb);
                  },
                  function(cb){
                    ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,5,nomtable[5],numligne[5],numfeuille[5],nomcolonne[5],nomBase,cb);
                  },
              ],
              function(err, resultat){
                if (err) { return res.view('Inovcom/erreur'); }
                else
                {
                  return res.view('Inovcom/accueiltype4', {date : datetest});
                }
            });
          });
    },
    accueiltype4 : function(req,res)
    {
      return res.view('Inovcom/accueiltype4');
    },
    EssaiExceltype4 : function(req,res)
    {
      var sql1= 'select count(*) as nb from chemininovcomtype4;';
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
          var sql='select * from chemininovcomtype4 limit' + " " + x ;
      Reportinghtp.query(sql,function(err, nc) {
        if (err){
          console.log(err);
          return next(err);
        }
        else
        {
            nc = nc.rows;
            sails.log('ko'+nc[0].chemin);
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
              
                    for(var i=0;i<nb;i++)
                    {
                     
                      var a = nc[i].chemin;
                      trameflux.push(a);
              
                    };
                    console.log("trameflux"+trameflux);
                    async.series([
                      function(cb){
                        ReportingInovcom.deleteHtp(table,nb,cb);
                      }, 
                      function(cb){
                        ReportingInovcom.importTrameFlux929type4(trameflux,feuil,cellule,table,cellule2,nb,numligne,cb);
                      }, 
                    ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                  })
               
        }
    })
  }
});
    },
    //Type 5
    accueil1type5 : function(req,res)
    {
      return res.view('Inovcom/accueil1type5');
    },
    Essaiitype5 : function(req,res)
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
            var newworksheet = workbook.getWorksheet('Feuil5');
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
                      ReportingInovcom.deleteFromChemin5(table,cb);
                    },
                  function(cb){
                      ReportingInovcom.importEssaitype5(table,cheminp,date,MotCle,0,cb);
                    },
              ],                                                                                                                                                                                   
              function(err, resultat){
                if (err) { return res.view('Inovcom/erreur'); }
                else
                {
                  return res.view('Inovcom/accueiltype5', {date : datetest});
                }
            });
          });
    },
    accueiltype5 : function(req,res)
    {
      return res.view('Inovcom/accueiltype5');
    },
    EssaiExceltype5 : function(req,res)
    {
      var sql1= 'select nb from nbinovcomtype5;';
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
          var sql='select * from chemininovcomtype5 limit 1'
      Reportinghtp.query(sql,function(err, nc) {
        if (err){
          console.log(err);
          return next(err);
        }
        else
        {
            nc = nc.rows;
            sails.log('ko'+nc[0].typologiedelademande);
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
            workbook.xlsx.readFile('Inovcom.xlsx')
                .then(function() {
                  var newworksheet = workbook.getWorksheet('Feuil5');
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
                      var a = cheminc[0]+date+cheminp[0]+nc[0].typologiedelademande;
                      trameflux.push(a);
                    console.log("trameflux"+trameflux);
                    async.series([
                      function(cb){
                        ReportingInovcom.deleteHtp(table,nb,cb);
                      }, 
                      function(cb){
                        ReportingInovcom.importTrameFlux929type5(trameflux,feuil,cellule,table,cellule2,nb,numligne,cb);
                      }, 
                    ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                  })
                });
        }
    })
  }
});
    },
  
  //Type 6
  accueil1type6 : function(req,res)
  {
    return res.view('Inovcom/accueil1type6');
  },
  Essaiitype6 : function(req,res)
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
          var newworksheet = workbook.getWorksheet('Feuil6');
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
                    ReportingInovcom.deleteFromChemin6(table,cb);
                  },
                function(cb){
                    ReportingInovcom.importEssaitype6(table,cheminp,date,MotCle,0,cb);
                  },
            ],
            function(err, resultat){
              if (err) { return res.view('Inovcom/erreur'); }
              else
              {
                return res.view('Inovcom/accueiltype6', {date : datetest});
              }
          });
        });
  },
  accueiltype6 : function(req,res)
  {
    return res.view('Inovcom/accueiltype6');
  },
  EssaiExceltype6 : function(req,res)
  {
    var sql= 'select * from chemininovcomtype6 limit 1;';
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
            var nb = 1;
            workbook.xlsx.readFile('Inovcom.xlsx')
                .then(function() {
                  var newworksheet = workbook.getWorksheet('Feuil6');
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
                        ReportingInovcom.importTrameFlux929type6(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                      }, 
                    ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                  })
                });
        }
    })
  },

  //Type 7
  accueil1type7 : function(req,res)
  {
    return res.view('Inovcom/accueil1type7');
  },
  Essaiitype7 : function(req,res)
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
    var nomcolonne3 = [];
    console.log(date);
    var cheminp = [];
    var MotCle= [];
    workbook.xlsx.readFile('Inovcom.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil7');
          var nomColonne3 = newworksheet.getColumn(3);
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
            console.log(cheminp[0]);
            console.log(MotCle[0]);
            async.series([  
                function(cb){
                    ReportingInovcom.deleteFromChemin7(table,cb);
                  },
                function(cb){
                    ReportingInovcom.importEssaitype7(table,cheminp,date,MotCle,0,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],nomcolonne2[0],nomcolonne3[0],cb);
                  },
            ],
            function(err, resultat){
              if (err) { return res.view('Inovcom/erreur'); }
              else
              {
                return res.view('Inovcom/accueiltype7', {date : datetest});
              }
          });
        });
  },
  accueiltype7 : function(req,res)
  {
    return res.view('Inovcom/accueiltype6');
  },
  EssaiExceltype7 : function(req,res)
  {
    var sql1= 'select count(*) as nb from chemininovcomtype7;';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from chemininovcomtype7 limit'  + " " + x;
          Reportinghtp.query(sql,function(err, nc) {
            if (err){
              console.log(err);
              return next(err);
            }
            else
            {
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
                    async.series([
                      function(cb){
                        ReportingInovcom.deleteHtp(table,nb,cb);
                      }, 
                      function(cb){
                        ReportingInovcom.importTrameFlux929type7(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                      }, 
                    ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                  });
               
        }
    })
  };
});
  },

   //Type 8
   accueil1type8 : function(req,res)
   {
     return res.view('Inovcom/accueil1type8');
   },
   Essaiitype8 : function(req,res)
   {
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
     var datetest = req.param("date",0);
     var annee = datetest.substr(0, 4);
     var mois = datetest.substr(5, 2);
     var jour = datetest.substr(8, 2);
     var date = annee+mois+jour;
     var type = [];
     var type2 = [];
     console.log(date);
     var cheminp = [];
     var MotCle= [];
     var nomtable = [];
     var numligne = [];
     var numfeuille = [];
     var nomcolonne = [];
     workbook.xlsx.readFile('Inovcom.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil8');
           var numFeuille = newworksheet.getColumn(4);
           var nomColonne = newworksheet.getColumn(5);
           var nomTable = newworksheet.getColumn(6);
           var numLigne = newworksheet.getColumn(8);
           var cheminparticulier = newworksheet.getColumn(9);
           var motcle = newworksheet.getColumn(10);
           var tipe = newworksheet.getColumn(3);
           var tipe2 = newworksheet.getColumn(7);
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
           tipe.eachCell(function(cell, rowNumber) {
              type.push(cell.value);
            });
           tipe2.eachCell(function(cell, rowNumber) {
              type2.push(cell.value);
            });
             console.log(cheminp[0]);
             console.log(MotCle[0]);
             async.series([  
                 function(cb){
                     ReportingInovcom.deleteFromChemin8(table,cb);
                   },
                 function(cb){
                     ReportingInovcom.importEssaitype8(table,cheminp,date,MotCle,0,type,type2,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],cb);
                   },
                 function(cb){
                    ReportingInovcom.importEssaitype8(table,cheminp,date,MotCle,1,type,type2,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],cb);
                  },
             ],
             function(err, resultat){
               if (err) { return res.view('Inovcom/erreur'); }
               else
               {
                 return res.view('Inovcom/accueiltype8', {date : datetest});
               }
           });
         });
   },
   accueiltype8 : function(req,res)
   {
     return res.view('Inovcom/accueiltype8');
   },
   EssaiExceltype8 : function(req,res)
   {
    var sql1= 'select count(*) as nb from chemininovcomtype8;';
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
       var sql= 'select * from chemininovcomtype8 limit' + " " + x;
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
                  ReportingInovcom.deleteHtp(table,nb,cb);
                }, 
               function(cb){
                  ReportingInovcom.importTrameFlux929type8(trameflux,feuil,cellule,table,cellule2,nb,numligne,cb);
                }, 
                     ],
                     function(err, resultat){
                       if (err) { return res.view('Inovcom/erreur'); }
                       return res.view('Retour/exportExcel');
                   });
                
         };
         
     });
    };
  });

   },

    //Type 9
    accueil1type9 : function(req,res)
    {
      return res.view('Inovcom/accueil1type9');
    },
    Essaiitype9 : function(req,res)
    {
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
      var date = annee+mois+jour;
      var type = [];
      var type2 = [];
      console.log(date);
      var cheminp = [];
      var MotCle= [];
      workbook.xlsx.readFile('Inovcom.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil9');
            var cheminparticulier = newworksheet.getColumn(9);
            var motcle = newworksheet.getColumn(10);
            var tipe = newworksheet.getColumn(3);
            var tipe2 = newworksheet.getColumn(7);
              cheminparticulier.eachCell(function(cell, rowNumber) {
                cheminp.push(cell.value);
              });
              motcle.eachCell(function(cell, rowNumber) {
                MotCle.push(cell.value);
              });
              tipe.eachCell(function(cell, rowNumber) {
               type.push(cell.value);
             });
             tipe2.eachCell(function(cell, rowNumber) {
               type2.push(cell.value);

             });
              console.log(cheminp[0]);
              var tab= 'recherchefactureinteriale';
              async.series([  
                  function(cb){
                      ReportingInovcom.deletetype9(tab,cb);
                    },
                  function(cb){
                      ReportingInovcom.importEssaitype9(table,cheminp,date,MotCle,0,cb);
                    },
                  
              ],
              function(err, resultat){
                if (err) { return res.view('Inovcom/erreur'); }
                else
                {
                  return res.view('Inovcom/accueiltype9', {date : datetest});
                }
            });
          });
    },
};

