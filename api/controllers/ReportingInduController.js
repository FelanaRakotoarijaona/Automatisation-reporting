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
      var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19];
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
                        ReportingInovcom.importEssai(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomBase,chem2,option2,cb);
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
    EssaiExcel : function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
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

    accueil2 : function(req,res)
    {
      return res.view('Indu/accueil2');
    },
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
      ReportingIndu.countOkKoContest("inducontestation",callback);
    },
    function (callback) {
      ReportingIndu.countOkKoDoubleSum("indufactstc",callback);
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
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK InduseAlmerys ==> " + result[0].ok + " / " + result[0].ko);
    console.log("Count OK 2 ==> " + result[1].ok + " / " + result[1].ko);
    console.log("Count OK 3 ==> " + result[2].ok + " / " + result[2].ko);
    console.log("Count OK IndutiersAlmerys ==> " + result[3].ok + " / " + result[3].ko);
    console.log("Count OK indufraudelmg ==> " + result[4].ok + " / " + result[4].ko);
    console.log("Count OK INDUCONTESTSATION ==> " + result[13].ok + " / " + result[13].ko);
    console.log("Count OK SANTE ==> " + result[14].ok + " / " + result[14].ko);
    console.log("Count OK Indutierscbtp ==> " + result[16].ok + " / " + result[16].ko);
    console.log("Count OK Indusecbtp ==> " + result[17].ok + " / " + result[17].ko);
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
        ReportingIndu.ecritureOkKoContest(result[13],"inducontestation",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoSante(result[14],"indufactstc",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoko(result[15],"indufraudelmgdent",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoDoublecbtp(result[16],"indutiers",date_export,mois1,callback);
      },
      function (callback) {
        ReportingIndu.ecritureOkKoDoublecbtp(result[17],"induse",date_export,mois1,callback);
      },
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('reporting/erera');
        }
        if(resultExcel[0]=='OK')
        {
          // res.redirect('/exportRetour/'+date_export+'/x')
          res.view('reporting/succes');
        }


        /*console.log("Traitement terminé ===> "+ resultExcel[0]);
        console.log("Traitement terminé ===> "+ resultExcel[1]);
        console.log("Traitement terminé ===> "+ resultExcel[2]);
        console.log("Traitement terminé ===> "+ resultExcel[3]);
        console.log("Traitement terminé ===> "+ resultExcel[4]);
        var html = "Echec d'enregistrement";
        return res.redirect('/accueil');*/
        
      
      
    })
  });  
},
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
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Contentieux/erera');
        }
        if(resultExcel[0]=='OK')
        {
          // res.redirect('/exportRetour/'+date_export+'/x')
          res.view('Contentieux/succes');
        }


        /*console.log("Traitement terminé ===> "+ resultExcel[0]);
        console.log("Traitement terminé ===> "+ resultExcel[1]);
        console.log("Traitement terminé ===> "+ resultExcel[2]);
        console.log("Traitement terminé ===> "+ resultExcel[3]);
        console.log("Traitement terminé ===> "+ resultExcel[4]);
        var html = "Echec d'enregistrement";
        return res.redirect('/accueil');*/
        
      
      
    })
  })
},
};

