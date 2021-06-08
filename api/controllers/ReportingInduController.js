/**
 * ReportingInduController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

module.exports = {
    accueilI : function(req,res)
    {
      return res.view('Indu/exportExcelIndu');
    },
    accueilI2 : function(req,res)
    {
      return res.view('Indu/exportExcelIndu2');
    },
    accueil1 : function(req,res)
    {
      return res.view('Indu/accueil1');
    },
    Essaii : function(req,res)
    {
      console.log('indu ve?');
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
      workbook.xlsx.readFile('Indu.xlsx')
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
                      ReportingIndu.deleteFromChemin(table,cb);
                    },
                 function(cb){
                      ReportingIndu.importEssai(table,cheminp,date,MotCle,0,cb);
                    },
                 /*function(cb){
                      ReportingIndu.importEssai(table,cheminp,date,MotCle,1,cb);
                    },
                    /*function(cb){
                      Reportinghtp.importEssai(table,cheminp,date,MotCle,2,cb);
                    },
                  function(cb){
                      Reportinghtp.importEssai(table,cheminp,date,MotCle,3,cb);
                    },
                    function(cb){
                      Reportinghtp.importEssai(table,cheminp,date,MotCle,4,cb);
                    },*/
              ],
              function(err, resultat){
                if (err) { return res.view('Indu/erreur'); }
                else
                {
                  return res.view('Indu/accueil', {date : datetest});
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
      var sql= 'select * from cheminindu limit 1;';
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
            workbook.xlsx.readFile('Indu.xlsx')
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
                        ReportingIndu.deleteHtp(table,nb,cb);
                      }, 
                      function(cb){
                        ReportingIndu.importTrameFlux929(trameflux,feuil,cellule,table,cellule2,nb,numligne,cb);
                      }, 
                      /*function(cb){
                            ReportingIndu.deleteTout(table,nb,cb);
                          }, 
                       function(cb){
                          ReportingIndu.deleteHtp(table,nb,cb);
                        }, 
                     function(cb){
                          ReportingIndu.importInovcom(trameflux,feuil,cellule,table,cellule2,numligne,nb,cb);
                          },
                     function(cb){
                        ReportingIndu.importTout(trameflux,table,nb,cb);
                        }, */
                    ],
                    function(err, resultat){
                      if (err) { return res.view('Indu/erreur'); }
                      // return res.redirect('/exportIndu/'+dateexport +'/'+'<h1><h1>');
                      return res.view('Indu/exportExcelIndu');
                  })
                });
        }
    })
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
    console.log(mois1);
    var date_export = jour + '/' + mois + '/' +annee;
    console.log("RECHERCHE COLONNE");
    async.series([
      function (callback) {
        ReportingIndu.countOkKoDouble("induse",callback);
      },
     function (callback) {
        ReportingIndu.countOkKoDouble("induhospi",callback);
      },
      function (callback) {
        ReportingIndu.countOkKoDoubleSum("indusansnotif",callback);
      },
      function (callback) {
        ReportingIndu.countOkKo("indutiers",callback);
      },
     function (callback) {
        ReportingIndu.countOkKoDoubleSum("indufraudelmg",callback);
      }, 
      function (callback) {
        ReportingIndu.countOkKoSum("induinterialepre",callback);
      },  
      function (callback) {
        ReportingIndu.countOkKo("induinterialepost",callback);
      }, 
      function (callback) {
        ReportingIndu.countOkKo("inducodelisftp",callback);
      }, 
      function (callback) {
        ReportingIndu.countOkKoSum("inducodelismail",callback);
      }, 
      function (callback) {
        ReportingIndu.countOkKo("inducodelisappel",callback);
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
    ],function(err,result){
      if(err) return res.badRequest(err);
      console.log("Count OK 1==> " + result[0].ok + " / " + result[0].ko);
      console.log("Count OK 2 ==> " + result[1].ok + " / " + result[1].ko);
      console.log("Count OK 3 ==> " + result[2].ok + " / " + result[2].ko);
      console.log("Count OK 4 ==> " + result[3].ok + " / " + result[3].ko);
      // console.log("Count OK tramelamiestock ==> " + result[4].ok + " / " + result[4].ko);
      // console.log("Count OK tramelamiestockResiliation ==> " + result[5].ok + " / " + result[5].ko);
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
          ReportingIndu.ecritureOkKo(result[3],"indutiers",date_export,mois1,callback);
        },
        function (callback) {
          ReportingIndu.ecritureOkKoDouble(result[4],"indufraudelmg",date_export,mois1,callback);
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
    })
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
    })
  },
};

