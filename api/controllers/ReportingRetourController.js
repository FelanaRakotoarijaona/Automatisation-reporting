/**
 * ReportingRetourController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */
module.exports = {
  // Accueil rechercher fichier
  accueilRetour : function(req,res)
  {
    return res.view('Retour/accueilRecherchefichier');
  },
  rechercheFichier : function(req,res)
  {
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
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
    console.log(date);
    var cheminp = [];
    var MotCle= [];
    var chem2 = [];
    var option2 = [];
    var nomBase = "cheminretour2";
    var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15];
    workbook.xlsx.readFile('ReportingRetourServeur.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil1');
          var numFeuille = newworksheet.getColumn(4);
          var nomColonne = newworksheet.getColumn(5);
          var nomTable = newworksheet.getColumn(6);
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
                      ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomBase,chem2,option2,cb);
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
                            return res.view('Retour/erreur');
                          }
                         if(f==0)
                          {
                            return res.view('Retour/erreur');
                          }
                          else
                          {
                            return res.view('Retour/accueilImport', {date : datetest});
                            
                          };
                      });
                  });
               
              }
          });
        });
  },
  // Accueil import
  accueilImportRetour : function(req,res)
  {
    return res.view('Retour/accueilImport');
  },
  ImportRetour : function(req,res)
  {
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var sql1= 'select count(*) as nb from cheminretour2;';
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
        var sql='select * from cheminretour2 limit' + " " + x ;
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
                        ReportingRetour.importRetour(trameflux,feuil,cellule,table,cellule2,lot,numligne,cb);
                      },
                    ],function(erroned, lotValues){
                      if(erroned) return res.badRequest(erroned);
                      return callback_reporting_suivant();
                    });
                  },
                    function(err)
                    {
                      console.log('vofafa ddol');
                      return res.view('Retour/exportExcel_2', {date : datetest});
                    }); 
           }
           })
      }
  });
},
   /* ********** */ 
    // Accueeil import
    accueil : function(req,res)
    {
      return res.view('Retour/accueil');
    },
    EssaiExcel : function(req,res)
    {
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
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
                          ReportingRetour.importRetour(trameflux,feuil,cellule,table,cellule2,lot,numligne,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                      function(err)
                      {
                        console.log('vofafa ddol');
                        return res.view('Retour/exportExcel', {date : datetest});
                      }); 
             }
             })
        }
    });
  },
  /* */
    accueilR : function(req,res)
    {
      return res.view('Retour/exportExcel');
    },
    accueil1 : function(req,res)
    {
      return res.view('Retour/accueil1');
    },
    Essaii : function(req,res)
    {
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      var table = ['/dev/pro/Retour_Easytech_'];
      //var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
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
      var chem2 = [];
      var option2 = [];
      var nomBase = "cheminretourvrai";
      var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14];
      workbook.xlsx.readFile('ReportingRetourServeur.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil2');
            var numFeuille = newworksheet.getColumn(4);
            var nomColonne = newworksheet.getColumn(5);
            var nomTable = newworksheet.getColumn(6);
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
                        ReportingInovcom.importEssaitype4(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomBase,chem2,option2,cb);
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
                              return res.view('Retour/erreur');
                            }
                           if(f==0)
                            {
                              return res.view('Retour/erreur');
                            }
                            else
                            {
                              return res.view('Retour/accueil', {date : datetest});
                              
                            };
                        });
                    });
                 
                }
            });
          });
    },
  
  /* EXPORT */

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
        Retour.countOkKoSum("trhospimulti",callback);
      },
   function (callback) {
        Retour.countOkKoSum("trstcdentaire",callback);
      },
      function (callback) {
        Retour.countOkKoSum("trstcoptique",callback);
      },
      function (callback) {
        Retour.countOkKoSum("trstcaudio",callback);
      },
     function (callback) {
        Retour.countOkKoSum("trretourfacttiers",callback);
      },
      // function (callback) {
      //   Retour.countOkKoSum("retouralmgto2",callback);
      // },
      function (callback) {
        Retour.countOkKoSum("trffacturehospi",callback);
      },
      function (callback) {
        Retour.countOkKoSum("trffacturedentaire",callback);
      },
      function (callback) {
        Retour.countOkKoSum("trfactureoptique",callback);
      },
      // function (callback) {
      //   Retour.countOkKoSum("retouralmgto6",callback);
      // },
      function (callback) {
        Retour.countOkKoSum("trhospi",callback);
      },
      function (callback) {
        Retour.countOkKoSum("trpecdentaire",callback);
      },
      function (callback) {
        Retour.countOkKoSum("trpecoptique",callback);
      },
      function (callback) {
        Retour.countOkKoSum("trpecaudio",callback);
      },
      // function (callback) {
      //   Retour.countOkKoSum("trldralmerys",callback);
      // },
     function (callback) {
        Retour.countOkKoSum("traaotd",callback);
      },
      // function (callback) {
      //   Retour.countOkKoSum("trretourotdn2",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("trtre",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retouralmftp4",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retouralmpackspe1",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retouralmpackspe2",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retouralmpackspe3",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("trindunoehtp",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retouralmcbtpgto",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retouretat1",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("trcentredesoin",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retouretat3",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retouretat4",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retouretat5",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retourpublipostage1",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retourpublipostage2",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("trhospimulti",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retourpublipostage4",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retourpublipostage5",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retourpublipostage6",callback);
      // },
     
    ],function(err,result){
      if(err) return res.badRequest(err);
      console.log("Count OK 1==> " + result[0].ok);
      console.log("Count OK 2 ==> " + result[1].ok );
      console.log("Count OK 3 ==> " + result[2].ok );
      console.log("Count OK 4 ==> " + result[3].ok );
      console.log("Count OK 1==> " + result[4].ok);
      console.log("Count OK 2 ==> " + result[5].ok );
      console.log("Count OK 3 ==> " + result[6].ok );
      console.log("Count OK 4 ==> " + result[7].ok );
      console.log("Count OK 1==> " + result[8].ok);
      console.log("Count OK 2 ==> " + result[9].ok );
      console.log("Count OK 3 ==> " + result[10].ok );

      async.series([
        function (callback) {
          Retour.ecritureOkKo8(result[0],"trhospimulti",date_export,mois1,callback);
        },
       function (callback) {
          Retour.ecritureOkKo2(result[1],"trstcdentaire",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo2(result[2],"trstcoptique",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo2(result[3],"trstcaudio",date_export,mois1,callback);
        },
      function (callback) {
          Retour.ecritureOkKo10(result[4],"trretourfacttiers",date_export,mois1,callback);
        },
        // function (callback) {
        //    Retour.ecritureOkKo3(result[5],"retouralmgto2",date_export,mois1,callback);
        // },
        function (callback) {
          Retour.ecritureOkKo3(result[5],"trffacturehospi",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo3(result[6],"trffacturedentaire",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo3(result[7],"trfactureoptique",date_export,mois1,callback);
        },
        // function (callback) {
        //   Retour.ecritureOkKo3(result[9],"retouralmgto6",date_export,mois1,callback);
        // },
        function (callback) {
          Retour.ecritureOkKo3(result[8],"trhospi",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo3(result[9],"trpecdentaire",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo3(result[10],"trpecoptique",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo3(result[11],"trpecaudio",date_export,mois1,callback);
        },
        // function (callback) {
        //   Retour.ecritureOkKo3(result[11],"trldralmerys",date_export,mois1,callback);
        // },
       function (callback) {
          Retour.ecritureOkKo4(result[12],"traaotd",date_export,mois1,callback);
        },
        // function (callback) {
        //   Retour.ecritureOkKo4(result[12],"trretourotdn2",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo4(result[13],"trtre",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo4(result[15],"retouralmftp4",date_export,mois1,callback);
        // },
        // function (callback) {
        //  Retour.ecritureOkKo5(result[19],"retouralmpackspe1",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo5(result[20],"retouralmpackspe2",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo5(result[21],"retouralmpackspe3",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo5(result[14],"trindunoehtp",date_export,mois1,callback);
        // },
      //  function (callback) {
      //     Retour.ecritureOkKo6(result[23],"retouralmcbtpgto",date_export,mois1,callback);
      //   },
      //  function (callback) {
      //     Retour.ecritureOkKo7(result[24],"retouretat1",date_export,mois1,callback);
      //   },
        // function (callback) {
        //   Retour.ecritureOkKo7(result[15],"trcentredesoin",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo7(result[26],"retouretat3",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo7(result[27],"retouretat4",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo7(result[28],"retouretat5",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[29],"retourpublipostage1",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[30],"retourpublipostage2",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[16],"trhospimulti",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[32],"retourpublipostage4",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[33],"retourpublipostage5",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[34],"retourpublipostage6",date_export,mois1,callback);
        // },
      
      ],function(err,resultExcel){
     console.log(resultExcel[0]);
     if(err)
     {
       console.log("Une erreur s'est prod");
       res.view('Retour/erera');
     }
     else if(resultExcel[0]=='OK' || resultExcel[1]== 'OK' || resultExcel[2]== 'OK'  || resultExcel[3]== 'OK' || resultExcel[4]== 'OK' || resultExcel[5]== 'OK' || resultExcel[6]== 'OK' || resultExcel[7]== 'OK' || resultExcel[8]== 'OK' || resultExcel[9]== 'OK' || resultExcel[10]=='OK')
     {
       // res.redirect('/exportRetour/'+date_export+'/x')
       // res.view('Retour/succes');
       res.view('Retour/exportretoursuivant', {date : datetest});
     }
     else
     {
      res.view('Retour/erera');
     }

          
        
        
      })
    })
  },
/********************************************/
rechercheColonne_suivant : function (req, res) {
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
      Retour.countOkKoSum("trldralmerys",callback);
    },
    function (callback) {
      Retour.countOkKoSum("trretourotdn2",callback);
    },
    function (callback) {
      Retour.countOkKoSum("trtre",callback);
    },
    function (callback) {
      Retour.countOkKoSum("trindunoehtp",callback);
    // },
    // function (callback) {
    //   Retour.countOkKoSum("trcentredesoin",callback);
    // },
    function (callback) {
      Retour.countOkKoSum("trhospimulti",callback);
    },
    function (callback) {
      Retour.countOkKoSum("trldrcbtp",callback);
    },
    function (callback) {
      Retour.countOkKoSum("trfactoptique",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0==> " + result[0].ok);
    console.log("Count OK 1==> " + result[1].ok );
    console.log("Count OK 2==> " + result[2].ok );
    console.log("Count OK 3==> " + result[3].ok );
    console.log("Count OK 4==> " + result[4].ok);
    console.log("Count OK 5==> " + result[5].ok );
    console.log("Count OK 6==> " + result[6].ok );
    console.log("Count OK 6==> " + result[7].ok );

    async.series([

      function (callback) {
        Retour.ecritureOkKo3(result[0],"trldralmerys",date_export,mois1,callback);
      },
      function (callback) {
        Retour.ecritureOkKo4(result[1],"trretourotdn2",date_export,mois1,callback);
      },
      function (callback) {
        Retour.ecritureOkKo4(result[2],"trtre",date_export,mois1,callback);
      },
      function (callback) {
        Retour.ecritureOkKo5(result[3],"trindunoehtp",date_export,mois1,callback);
      },
      // function (callback) {
      //   Retour.ecritureOkKo7(result[4],"trcentredesoin",date_export,mois1,callback);
      // },
      function (callback) {
        Retour.ecritureOkKo8(result[4],"trhospimulti",date_export,mois1,callback);
      },
      function (callback) {
        Retour.ecritureOkKo6(result[5],"trldrcbtp",date_export,mois1,callback);
      },
      function (callback) {
        Retour.ecritureOkKoFactOptique(result[6],"trfactoptique",date_export,mois1,callback);
      },
    
    ],function(err,resultExcel){
   console.log(resultExcel[0]);
   if(err)
   {
     console.log("Une erreur s'est prod");
     res.view('Retour/erera');
   }
   else if(resultExcel[0]=='OK' || resultExcel[1]== 'OK' || resultExcel[2]== 'OK'  || resultExcel[3]== 'OK' || resultExcel[4]== 'OK' || resultExcel[5]== 'OK' || resultExcel[6]== 'OK' )
   {
     // res.redirect('/exportRetour/'+date_export+'/x')
     res.view('Retour/accueilRecherchefichier');
   }
   else
   {
    res.view('Retour/erera');
   }
    })
  })
},


/********************************************/
rechercheColonnetest : function (req, res) {
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
      Retour.countOkKoSum("trhospimulti",callback);
    },
   
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok);
    async.series([
      function (callback) {
        Retour.ecritureOkKotest(result[0],"trhospimulti",date_export,mois1,callback);
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
          // res.redirect('/exportRetour/'+date_export+'/x')
          res.view('Retour/succes');
        }
        
      
      
    })
  })
},
 //AJOUT FONCTION RECHERCHECOLONNE POUR RETOUR
 rechercheColonne_2 : function (req, res) {
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
      Retour.countOkKoSum("trse",callback);
    },
 function (callback) {
      Retour.countOkKoSum("trretourotdn2cbtp",callback);
    },
    function (callback) {
      Retour.countOkKoSum("trautoqueryannulation",callback);
    },
    function (callback) {
      Retour.countOkKoSum("trautoqueryradiation",callback);
    },
   function (callback) {
      Retour.countOkKoSum("trautoquerydoublon",callback);
    },
   
  
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    

    async.series([
      function (callback) {
        Retour.ecritureOkKo8(result[0],"trse",date_export,mois1,callback);
      },
     function (callback) {
        Retour.ecritureOkKo9(result[1],"trretourotdn2cbtp",date_export,mois1,callback);
      },
      function (callback) {
        Retour.ecritureOkKo5(result[2],"trautoqueryannulation",date_export,mois1,callback);
      },
      function (callback) {
        Retour.ecritureOkKo5(result[3],"trautoqueryradiation",date_export,mois1,callback);
      },
    function (callback) {
        Retour.ecritureOkKo5(result[4],"trautoquerydoublon",date_export,mois1,callback);
      },
  
    
    ],function(err,resultExcel){
      if(err)
      {
        console.log("Une erreur s'est prod");
        res.view('Retour/erera');
      }
      else if(resultExcel[0]=='OK' || resultExcel[1]== 'OK' || resultExcel[2]== 'OK'  || resultExcel[3]== 'OK' || resultExcel[4]== 'OK' )
      {
        // res.redirect('/exportRetour/'+date_export+'/x')
        res.view('Retour/succes');
        // res.view('Retour/exportretoursuivant', {date : datetest});
      }
      else
      {
        res.view('Retour/erera');
      }
    })
  })
},

};

