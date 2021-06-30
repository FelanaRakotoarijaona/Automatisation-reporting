/**
 * ReportingContetieuxController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

module.exports = {
    accueil1 : function(req,res)
    {
      return res.view('Contentieux/accueil1');
    },

    accueilCont : function(req,res)
      {
        return res.view('Contentieux/exportExcelContentieux');
      },
  
   
      accueil1 : function(req,res)
      {
        return res.view('Contentieux/accueil1');
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
      console.log(date);
      var cheminp = [];
      var MotCle= [];
      var chem2 = [];
      var option2 = [];
      var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22];
      //var r = [0,1];
      //workbook.xlsx.readFile('ReportingContetieux.xlsx')
      workbook.xlsx.readFile('ReportingContetieuxserveur.xlsx')
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
              var nomBase = "chemincontetieux";
              console.log(cheminp[0]);
              console.log(MotCle[0]);
              async.series([  
                function(cb){
                  ReportingInovcom.deleteFromChemin(nomBase,cb);
                  },
              ],
              function(err, resultat){
                if (err) { return res.view('Contentieux/erreur'); }
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
                              return res.view('Contentieux/erreur');
                            }
                           if(f==0)
                            {
                              return res.view('Contentieux/erreur');
                            }
                            else
                            {
                              return res.view('Contentieux/accueil', {date : datetest});
                              
                            };
                        });  
                      };
                      
                    });
                 
                }
            });
          });
    },
      
    /* Import */
      accueil : function(req,res)
      {
        // return res.view('Retour/accueil');
        return res.view('Contentieux/accueil');
      },
      EssaiExcel : function(req,res)
      {
        var datetest = req.param("date",0);
        var sql1= 'select count(*) as nb from chemincontetieux;';
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
            var sql='select * from chemincontetieux limit' + " " + x ;
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
                          return res.view('Contentieux/exportExcelContentieux', {date : datetest});
                        }); 
               }
               })
          }
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
          ReportingContetieux.countOkKoSum("coaaotdalmerys",callback);
        },
        function (callback) {
          ReportingContetieux.countOkKo("coldralmeryspublic",callback);
        },
        function (callback) {
          ReportingContetieux.countOkKoSum("cootdalmerys",callback);
        },
        function (callback) {
          ReportingContetieux.countOkKo("cosdralmerys",callback);
        },
       function (callback) {
          ReportingContetieux.countOkKoSum("cootdclient",callback);
        },
        function (callback) {
          ReportingContetieux.countOkKo("coadraphpalmerys",callback);
        },
        function (callback) {
           ReportingContetieux.countOkKoSum("coadrclassiquealmerys",callback);
        },
        function (callback) {
           ReportingContetieux.countOkKoSum("coimputationalmerys",callback);
        },
        function (callback) {
            ReportingContetieux.countOkKo("coaaotdcbtp",callback);
        },
        function (callback) {
        ReportingContetieux.countOkKoDoubleSum("coldrcbtppublic",callback);
        },
        function (callback) {
        ReportingContetieux.countOkKo("cootdcbtp",callback);
        },
        function (callback) {
        ReportingContetieux.countOkKo("cosdrcbtp",callback);
        },
        function (callback) {
        ReportingContetieux.countOkKo("coadraphpcbtp",callback);
        },
        function (callback) {
            ReportingContetieux.countOkKoSum("coadrclassiquecbtp",callback);
        },
        function (callback) {
            ReportingContetieux.countOkKoSum("coimputationcbtp",callback);
        },

       
      ],function(err,result){
        if(err) return res.badRequest(err);
        console.log("Count OK 1==> " + result[0].ok + " / " + result[0].ko);
        console.log("Count OK 2 ==> " + result[1].ok + " / " + result[1].ko);
        console.log("Count OK 3 ==> " + result[2].ok + " / " + result[2].ko);
        console.log("Count OK 4 ==> " + result[3].ok + " / " + result[3].ko);
        console.log("Count OK 9 ==> " + result[9].ok + " / " + result[9].ko);
        // console.log("Count OK tramelamiestock ==> " + result[4].ok + " / " + result[4].ko);
        // console.log("Count OK tramelamiestockResiliation ==> " + result[5].ok + " / " + result[5].ko);
        async.series([
        function (callback) {
            ReportingContetieux.ecritureOkKo(result[0],"coaaotdalmerys",date_export,mois1,callback);
          },
        function (callback) {
            ReportingContetieux.ecritureOkKo(result[1],"coldralmeryspublic",date_export,mois1,callback);
          },
        function (callback) {
            ReportingContetieux.ecritureOkKo(result[2],"cootdalmerys",date_export,mois1,callback);
          },
        function (callback) {
            ReportingContetieux.ecritureOkKo(result[3],"cosdralmerys",date_export,mois1,callback);
          },
        function (callback) {
            ReportingContetieux.ecritureOkKo(result[4],"cootdclient",date_export,mois1,callback);
          },
        function (callback) {
            ReportingContetieux.ecritureOkKo(result[5],"coadraphpalmerys",date_export,mois1,callback);
          },
        function (callback) {
            ReportingContetieux.ecritureOkKo(result[6],"coadrclassiquealmerys",date_export,mois1,callback);
          },
        function (callback) {
            ReportingContetieux.ecritureOkKo(result[7],"coimputationalmerys",date_export,mois1,callback);
         },
        function (callback) {
            ReportingContetieux.ecritureOkKo2(result[8],"coaaotdcbtp",date_export,mois1,callback);
        },
        function (callback) {
            ReportingContetieux.ecritureOkKoDouble(result[9],"coldrcbtppublic",date_export,mois1,callback);
        },
        function (callback) {
            ReportingContetieux.ecritureOkKo2(result[10],"cootdcbtp",date_export,mois1,callback);
        },
        function (callback) {
            ReportingContetieux.ecritureOkKo2(result[11],"cosdrcbtp",date_export,mois1,callback);
        },
        function (callback) {
            ReportingContetieux.ecritureOkKo2(result[12],"coadraphpcbtp",date_export,mois1,callback);
        },
        function (callback) {
            ReportingContetieux.ecritureOkKo2(result[13],"coadrclassiquecbtp",date_export,mois1,callback);
        },
        function (callback) {
            ReportingContetieux.ecritureOkKo2(result[14],"coimputationcbtp",date_export,mois1,callback);
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


