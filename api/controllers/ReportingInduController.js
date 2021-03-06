/**
 * ReportingInduController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

const ReportingInovcom = require('../models/ReportingInovcom');
module.exports = {
  //type 3
  accueiltype3 : function(req,res)
  {
    return res.view('Indu/accueiltype3');
  },
  Essaii3 : function(req,res)
  {
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    //var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
    var table = ['/dev/pro/Retour_Easytech_'];
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
    var chem2 = [];
    var option2 = [];
    var cheminp = [];
    var MotCle= [];
    var r = [0,1,2,3,4,5,6,7];
    var nomBase = "cheminindu3";
    //workbook.xlsx.readFile('ReportingIndu.xlsx')
    workbook.xlsx.readFile('ReportingInduserveur.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil1');
          var numFeuille = newworksheet.getColumn(4);
          var nomColonne = newworksheet.getColumn(5);
          var nomTable = newworksheet.getColumn(6);
          var nomColonne2 = newworksheet.getColumn(7);
          var numLigne = newworksheet.getColumn(8);
          var cheminparticulier = newworksheet.getColumn(9);
          var motcle = newworksheet.getColumn(10);
          var chemin2 = newworksheet.getColumn(11);
          var opt2 = newworksheet.getColumn(12);
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
            chemin2.eachCell(function(cell, rowNumber) {
              chem2.push(cell.value);
            });
            opt2.eachCell(function(cell, rowNumber) {
              option2.push(cell.value);
            });
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
                      ReportingIndu.importEssaitype3(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,cb);
                    },
                  ],function(erroned, lotValues){
                    if(erroned) return res.badRequest(erroned);
                    return callback_reporting_suivant();
                  });
                },
                  function(err)
                  {
                    if (err){
                      return res.view('Contentieux/erreur');
                    }
                    else
                    {
                    var sql4= "select count(*) as ok from "+nomBase+" ";
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
                            return res.view('Indu/accueil3', {date : datetest});
                            
                          };
                      });
                    }
                  });
              }
          });
        });
  },
  accueil3 : function(req,res)
  {
    return res.view('Indu/accueil3');
  },
  essaiExcel3 : function(req,res)
  {
    var dateFormat = require("dateformat");
    var datetest = req.param("date",0);
    var today = new Date(datetest);
    var tomorrow = new Date(today);
    var f = tomorrow.setDate(today.getDate()- 1);
    var date2=dateFormat(f,"shortDate");
    console.log(date2);
    var sql1= 'select count(*) as nb from cheminindu3;';
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
    var sql= 'select * from cheminindu3 limit' + " " + x ;
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
                        ReportingIndu.importindulignerouge(trameflux,feuil,cellule,table,cellule2,lot,numligne,cb);
                      }, 
                    ],function(erroned, lotValues){
                      if(erroned) return res.badRequest(erroned);
                      return callback_reporting_suivant();
                    });
                  },
                  function(err)
                    {
                      console.log('vofafa ddol');
                      return res.view('Indu/exportExcelIndu3',{date : datetest});//exportExcelIndu3
                      //return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                    });
      };
  })
  };
});
},

  // type 1
    accueil1 : function(req,res)
    {
      return res.view('Indu/accueil1');
    },
    //CIBLAGE DU FICHIER DANS LE SERVEUR
    Essaii : function(req,res)
    {
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      //var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
      var table = ['/dev/pro/Retour_Easytech_'];
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
      var chem2 = [];
      var option2 = [];
      var cheminp = [];
      var MotCle= [];
      var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21];
     // var r = [0,1,2,3,4,5];
      var nomBase = "cheminindu";
      //workbook.xlsx.readFile('ReportingIndu.xlsx')
      workbook.xlsx.readFile('ReportingInduserveur.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil2');
            var numFeuille = newworksheet.getColumn(4);
            var nomColonne = newworksheet.getColumn(5);
            var nomTable = newworksheet.getColumn(6);
            var nomColonne2 = newworksheet.getColumn(7);
            var numLigne = newworksheet.getColumn(8);
            var cheminparticulier = newworksheet.getColumn(9);
            var motcle = newworksheet.getColumn(10);
            var chemin2 = newworksheet.getColumn(11);
            var opt2 = newworksheet.getColumn(12);
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
              chemin2.eachCell(function(cell, rowNumber) {
                chem2.push(cell.value);
              });
              opt2.eachCell(function(cell, rowNumber) {
                option2.push(cell.value);
              });
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
                        ReportingIndu.importEssai(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomBase,chem2,option2,cb);
                      },
                    ],function(erroned, lotValues){
                      if(erroned) return res.badRequest(erroned);
                      return callback_reporting_suivant();
                    });
                  },
                    function(err)
                    {
                      if (err){
                        return res.view('Contentieux/erreur');
                      }
                      else
                      {
                      var sql4= "select count(*) as ok from "+nomBase+" ";
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
                      }
                    });
                }
            });
          });
    },
    accueil : function(req,res)
    {
      return res.view('Indu/accueil');
    },

    //REQUETE BASE DE DONNEE (donn??ee des chemins du serveur)
    EssaiExcel : function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var today = new Date(datetest);
      var tomorrow = new Date(today);
      var f = tomorrow.setDate(today.getDate()- 10);
      var date2=dateFormat(f,"shortDate");
      var date3 =dateFormat(today,"shortDate");
      console.log(date3 + 'dattte');
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
                          ReportingIndu.importTrameFlux929(trameflux,feuil,cellule,table,cellule2,lot,numligne,date2,date3,cb);
                        }, 
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    function(err)
                      {
                        console.log('vofafa ddol');
                        return res.view('Indu/exportExcelIndu',{date : datetest});
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
    //REQUETE BASE DE DONNEE (donn??ee des chemins du serveur)
    Essaii2 : function(req,res)
    {
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      //var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
      var table = ['/dev/pro/Retour_Easytech_'];
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
      //workbook.xlsx.readFile('ReportingIndu.xlsx')
      workbook.xlsx.readFile('ReportingInduserveur.xlsx')
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
                      if (err){
                        return res.view('Contentieux/erreur');
                      }
                      else
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
                      }
                    });
                }
            });
          });
    },
    //REDIRECTION VERS ACCEUILL2
    accueil2 : function(req,res)
    {
      return res.view('Indu/accueil2');
    },
    //REQUETE BASE DE DONNEE (donn??ee des chemins du serveur)
    EssaiExcel2 : function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
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
                        return res.view('Indu/exportExcelIndu2',{date : datetest});
                        //return res.redirect('/exportInovcom/'+dateexport +'/'+'<h1><h1>');
                      });
        };
    })
    };
  });
},

//AJOUT FONCTION RECHERCHECOLONNE POUR RETOUR
rechercheColonne : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  // var jour = req.param("jour");
  // var mois = req.param("mois");
  // var annee = req.param("annee");
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
    mois1= 'Octobre';
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
      ReportingIndu.countOkKoDoubleSum("induse",callback);
    },
   function (callback) {
      ReportingIndu.countOkKoDoubleSum("induhospi",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoDoubleSum("indusansnotif",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoDoubleSum("indutiers",callback);
    },
   function (callback) {
      ReportingIndu.countOkKoSum("indufraudelmg",callback);
    }, 
    function (callback) {
      ReportingIndu.countOkKoSum("induinterialepre",callback);
    },  
    function (callback) {
      ReportingIndu.countOkKoSum("induinterialepost",callback);
    }, 
    function (callback) {
      ReportingIndu.countOkKoSum("inducodelisftp",callback);
    }, 
    function (callback) {
      ReportingIndu.countOkKoSum("inducodelismail",callback);
    }, 
    function (callback) {
      ReportingIndu.countOkKoSum("inducodelisappel",callback);
    }, 
    function (callback) {
      ReportingIndu.countOkKoDoubleSum("inducheque",callback);
    }, 
    function (callback) {
      ReportingIndu.countOkKoSum("indupecrefus",callback);
    }, 
    function (callback) {
      ReportingIndu.countOkKoSum("induinterialeaudio",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoSumko("induentrain",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoSumko("induaudio",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoDoubleSum("indusansnotifcbtp",callback);
    },


  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
    console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
    console.log("Count OK 2 ==> " + result[2].ok + " / " + result[2].ko);
    console.log("Count OK 3 ==> " + result[3].ok + " / " + result[3].ko);
    console.log("Count OK 4 ==> " + result[4].ok + " / " + result[4].ko);
    console.log("Count OK 5 ==> " + result[5].ok + " / " + result[5].ko);
    console.log("Count OK 6 ==> " + result[6].ok + " / " + result[6].ko);
    console.log("Count OK 7 ==> " + result[7].ok + " / " + result[7].ko);
    console.log("Count OK 8 ==> " + result[8].ok + " / " + result[8].ko);
    console.log("Count OK 9 ==> " + result[9].ok + " / " + result[9].ko);
    console.log("Count OK 10 ==> " + result[10].ok + " / " + result[10].ko);
    console.log("Count OK 11 ==> " + result[11].ok + " / " + result[11].ko);
    console.log("Count OK 12 ==> " + result[12].ok + " / " + result[12].ko);
    console.log("Count OK 13 ==> " + result[13].ok + " / " + result[13].ko);
    console.log("Count OK 14 ==> " + result[14].ok + " / " + result[14].ko);
    console.log("Count OK 15 ==> " + result[15].ok + " / " + result[15].ko);
   
    async.series([
      function (callback) {
        ReportingIndu.ecritureOkKoDouble(result[0],"induse",date_export,mois1,callback);
      },
     function (callback) {
        ReportingIndu.ecritureOkKoDouble(result[1],"induhospi",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoDouble(result[2],"indusansnotif",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoDouble(result[3],"indutiers",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKo(result[4],"indufraudelmg",date_export,mois1,callback);
      },
     function (callback) {
        ReportingIndu.ecritureOkKo(result[5],"induinterialepre",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKo(result[6],"induinterialepost",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKo(result[7],"inducodelisftp",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKo(result[8],"inducodelismail",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKo(result[9],"inducodelisappel",date_export,mois1,callback);
      },
     function (callback) {
        ReportingIndu.ecritureOkKoDouble(result[10],"inducheque",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKo(result[11],"indupecrefus",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKo(result[12],"induinterialeaudio",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoDoubleInduentrain1(result[13],"induentrain",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoalm(result[14],"induaudio",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoDoublecbtp(result[15],"indusansnotifcbtp",date_export,mois1,callback);
      },

      
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('reporting/erera');
        }
        if(resultExcel[0]=='OK' || resultExcel[1]== 'OK' || resultExcel[2]== 'OK'  || resultExcel[3]== 'OK' || resultExcel[4]== 'OK' || resultExcel[5]== 'OK' || resultExcel[6]== 'OK' || resultExcel[7]== 'OK' || resultExcel[8]== 'OK' || resultExcel[9]== 'OK' || resultExcel[10]== 'OK' || resultExcel[11]== 'OK' || resultExcel[12]== 'OK' || resultExcel[13]== 'OK'   || resultExcel[14]== 'OK' || resultExcel[15]== 'OK')
        {
          // res.redirect('/exportRetour/'+date_export+'/x')
          // res.view('reporting/succes');
          return res.view('Indu/exportExcelIndusuivant',{date : datetest});
        }

        
      
      
    })
  });  
},
/*************************************************************************************/
rechercheColonneindusuivant : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  // var jour = req.param("jour");
  // var mois = req.param("mois");
  // var annee = req.param("annee");
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
    mois1= 'Octobre';
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
      ReportingIndu.countOkKoContest("inducontestation",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoSum("indufactstc",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoSumko("indufactstc",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoSumko("indufraudelmg",callback);
    }, 
    function (callback) {
      ReportingIndu.countOkKoDoubleSumcbtp("indutiers",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoDoubleSumcbtp("induse",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoSum("indufactstc",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoSumko("indufactstc",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoSum("induentrain",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
    console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
    console.log("Count OK 2 ==> " + result[2].ok + " / " + result[2].ko);    
    console.log("Count OK 3 ==> " + result[3].ok + " / " + result[3].ko);
    console.log("Count OK 4 ==> " + result[4].ok + " / " + result[4].ko);
    console.log("Count OK 5 ==> " + result[5].ok );
    console.log("Count OK 6 ==> " + result[6].ok );
    console.log("Count OK 7 ==> " + result[7].ok );
    async.series([
      function (callback) {
        ReportingIndu.ecritureOkKoContest(result[0],"inducontestation",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoSante(result[1],"indufactstc",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoSante(result[2],"indufactstcdy",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoko(result[3],"indufraudelmgdent",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoDoublecbtp(result[4],"indutiers",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoDoublecbtp(result[5],"induse",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoSantesaisis(result[6],"indufactstcobs",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoSantesaisis(result[7],"indufactstcazur",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoDoubleInduentrain1(result[8],"induentrainavant",date_export,mois1,callback);
      },
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('reporting/erera');
        }
        if(resultExcel[0]=='OK' || resultExcel[1]== 'OK' || resultExcel[2]== 'OK'  || resultExcel[3]== 'OK' || resultExcel[4]== 'OK' || resultExcel[5]== 'OK' || resultExcel[6]== 'OK' || resultExcel[7]== 'OK'  || resultExcel[8]== 'OK')
        {
          // res.redirect('/exportRetour/'+date_export+'/x')
          res.view('reporting/succes');
        }

        
      
      
    })
  });  
},

/****************************************************************************************/
//RELEVE DE COMPTE
rechercheColonne2 : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  // var jour = req.param("jour");
  // var mois = req.param("mois");
  // var annee = req.param("annee");
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
    mois1= 'Octobre';
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
      ReportingIndu.countOkKoIndu2("indurelevedecomptealmerys",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoIndu2("indurelevedecomptecbtp",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoIndu2("induvalidation",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0==> " + result[0].ok + " / " + result[0].ko);
    console.log("Count OK 1==> " + result[1].ok + " / " + result[1].ko);
    async.series([
      function (callback) {
        ReportingIndu.ecritureOkKoIndu2(result[0],"indurelevedecomptealmerys",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoIndu2cbtp(result[1],"indurelevedecomptecbtp",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoIndu2(result[2],"induvalidation",date_export,mois1,callback);
      },
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Contentieux/erera');
        }
        if(resultExcel[0]=='OK' || resultExcel[1]=='OK')
        {
          // res.redirect('/exportRetour/'+date_export+'/x')
          res.view('Contentieux/succes');
        }

        
      
      
    })
  })
},
/****************************************************************************************/
//INDU LIGNE ROUGE
rechercheColonne3 : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  // var jour = req.param("jour");
  // var mois = req.param("mois");
  // var annee = req.param("annee");
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
    mois1= 'Octobre';
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
      ReportingIndu.countOkKoIndu3("indufraudeinterialej",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoIndu3("indufraudeinterialej1",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoIndu3("indufraudeinteriale12mois",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoIndu3("indufraudeinteriale15j",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoIndu3("indufraudeeolej",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoIndu3("indufraudeinterialej1",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoIndu3("indufraudeeole12mois",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoIndu3("indufraudeeole15j",callback);
    },
 
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0==> " + result[0].ok);
    console.log("Count OK 1==> " + result[1].ok);
    console.log("Count OK 2==> " + result[2].ok);
    console.log("Count OK 3==> " + result[3].ok);
    console.log("Count OK 4==> " + result[4].ok);
    console.log("Count OK 5==> " + result[5].ok);
    console.log("Count OK 6==> " + result[6].ok);
    console.log("Count OK 7==> " + result[7].ok);
    async.series([
      function (callback) {
        ReportingIndu.ecritureOkKoIndu3(result[0],"indufraudeinterialej",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoIndu3(result[1],"indufraudeinterialej1",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoIndu3(result[2],"indufraudeinteriale12mois",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoIndu3(result[3],"indufraudeinteriale15j",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoIndu3(result[4],"indufraudeeolej",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoIndu3(result[5],"indufraudeeolej1",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoIndu3(result[6],"indufraudeeole12mois",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoIndu3(result[7],"indufraudeeole15j",date_export,mois1,callback);
      },
  
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Contentieux/erera');
        }
        if(resultExcel[0]=='OK' || resultExcel[1]== 'OK' || resultExcel[2]== 'OK'  || resultExcel[3]== 'OK' || resultExcel[4]== 'OK' || resultExcel[5]== 'OK' || resultExcel[6]== 'OK' || resultExcel[7]== 'OK'  )
        {
          // res.redirect('/exportRetour/'+date_export+'/x')
          res.view('Contentieux/succes');
        }

    })
  })
},

};

