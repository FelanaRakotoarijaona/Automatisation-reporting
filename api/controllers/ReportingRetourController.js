/**
 * ReportingRetourController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

// const ReportingRetour = require('../models/ReportingRetour');
module.exports = {

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
      var chem2 = [];
      var option2 = [];
      var nomBase = "cheminretourvrai";
      var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19];
      workbook.xlsx.readFile('ReportingRetour.xlsx')
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
                        return res.view('Retour/exportExcel', {date : datetest});
                      }); 
             }
             })
        }
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
        Retour.countOkKoSum("trtramepecdentaire",callback);
      },
      // function (callback) {
      //   Retour.countOkKoSum("retouralmgto9",callback);
      // },
      function (callback) {
        Retour.countOkKoSum("trpecaudio",callback);
      },
      function (callback) {
        Retour.countOkKoSum("trldralmerys",callback);
      },
    //  function (callback) {
    //     Retour.countOkKoSum("retouralmftp1",callback);
    //   },
      function (callback) {
        Retour.countOkKoSum("trretourotdn2",callback);
      },
      function (callback) {
        Retour.countOkKoSum("trtre",callback);
      },
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
      function (callback) {
        Retour.countOkKoSum("trindunoehtp",callback);
      },
      // function (callback) {
      //   Retour.countOkKoSum("retouralmcbtpgto",callback);
      // },
      // function (callback) {
      //   Retour.countOkKoSum("retouretat1",callback);
      // },
      function (callback) {
        Retour.countOkKoSum("trcentredesoin",callback);
      },
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
      function (callback) {
        Retour.countOkKoSum("trhospimulti",callback);
      },
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
      console.log("Count OK 4 ==> " + result[11].ok );
      console.log("Count OK 1==> " + result[12].ok);
      console.log("Count OK 2 ==> " + result[13].ok );
      console.log("Count OK 3 ==> " + result[14].ok );
      console.log("Count OK 4 ==> " + result[15].ok );
      console.log("Count OK 4 ==> " + result[16].ok );
      // console.log("Count OK tramelamiestock ==> " + result[4].ok + " / " + result[4].ko);
      // console.log("Count OK tramelamiestockResiliation ==> " + result[5].ok + " / " + result[5].ko);
      async.series([
        function (callback) {
          Retour.ecritureOkKo(result[0],"trhospimulti",date_export,mois1,callback);
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
          Retour.ecritureOkKo3(result[4],"trretourfacttiers",date_export,mois1,callback);
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
          Retour.ecritureOkKo3(result[9],"trtramepecdentaire",date_export,mois1,callback);
        },
        // function (callback) {
        //   Retour.ecritureOkKo3(result[12],"retouralmgto9",date_export,mois1,callback);
        // },
        function (callback) {
          Retour.ecritureOkKo3(result[10],"trpecaudio",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo3(result[11],"trldralmerys",date_export,mois1,callback);
        },
      //  function (callback) {
      //     Retour.ecritureOkKo4(result[12],"retouralmftp1",date_export,mois1,callback);
      //   },
        function (callback) {
          Retour.ecritureOkKo4(result[12],"trretourotdn2",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo4(result[13],"trtre",date_export,mois1,callback);
        },
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
        function (callback) {
          Retour.ecritureOkKo5(result[14],"trindunoehtp",date_export,mois1,callback);
        },
      //  function (callback) {
      //     Retour.ecritureOkKo6(result[23],"retouralmcbtpgto",date_export,mois1,callback);
      //   },
      //  function (callback) {
      //     Retour.ecritureOkKo7(result[24],"retouretat1",date_export,mois1,callback);
      //   },
        function (callback) {
          Retour.ecritureOkKo7(result[15],"trcentredesoin",date_export,mois1,callback);
        },
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
        function (callback) {
          Retour.ecritureOkKo8(result[16],"trhospimulti",date_export,mois1,callback);
        },
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

