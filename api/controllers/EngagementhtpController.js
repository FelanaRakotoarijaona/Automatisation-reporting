/**
 * EngagementhtpController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

module.exports = {
  //fonction n'est pas encore en service
  exporthtpengagement1: function(req, res){
      console.log('commencer');
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
      var feuille = annee+mois+'_EASY';
      console.log(feuille);

      var date_export = jour + '/' + mois + '/' +annee;
      console.log("RECHERCHE FEUILLE EXCEL");

      async.series([
          function (callback) {
              Engagementhtp.recupdata("trhospimulti",callback);
          },
         
        ],function(err,result){
          if(err) return res.badRequest(err);
          console.log("Count OK 0 ==> " + result[0].ok);
          async.series([
            function (callback) {
              Engagementhtp.ecriture(result[0],"trhospimulti",date_export,feuille,callback);
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
  /*********************************/
  accueilengagementhtp : async function(req,res)
  {
    return res.view('HTPengagement/accueilreportingengagementhtp');
  },
  /***********************************************************************************/
  //FONCTION POUR L'IMPORT DU CHEMIN UTILISER (22-07-2021)
insertcheminengagementhtp_1 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
  // var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtp";
  var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19];
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil1');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

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
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          console.log(nomtable[0]);
          console.log(nomtable[1]);
          async.series([  
            function(cb){
                  Engagementhtp.deleteFromChemin(nomBase,cb);
                },
                                  
          ],
          function(err, resultat){
            if(err){
              return res.view('Inovcom/erreur');
            }
            else{
              async.forEachSeries(r, function(lot, callback_reporting_suivant){
                async.series([
                  function(cb){
                    Engagementhtp.delete(nomtable,lot,cb);
                  },
                  function(cb){
                    Engagementhtp.importcheminhtp(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                  },
                  // function(cb){
                  //   Engagementhtp.importcheminhtp(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                  // },
                ],
                function(erroned, lotValues){
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
                        return res.view('Inovcom/erreur');
                      }
                     if(f==0)
                      {
                        return res.view('Inovcom/erreur');
                      }
                      else
                      {
                        // return res.view('HTPengagement/accueilreportingengagementhtpsuivant', {date : datetest});
                        // return res.view('HTPengagement/importHTPengagement_1', {date : datetest});
                        return res.view('HTPengagement/accueilreportingengagementhtpsuivant2', {date : datetest});
                      };
                  });
                }
              });

            }
           
        });
      });
},

  /********************************************************************************/
  insertcheminengagementhtpsuivant_2 : function(req,res)
  {
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
    // var table = ['/dev/pro/Retour_Easytech_'];
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var date = annee+mois+jour;
    var date_indus = jour+'.'+mois+'.'+annee;
    var datej_1 = annee+mois+jour -1;
    console.log(datej_1);
    var nomtable = [];
    var numligne = [];
    var numfeuille = [];
    var nomcolonne = [];
    var nomcolonne2 = [];
    var nomcolonne3 = [];
    console.log(date);
    var cheminp = [];
    var MotCle= [];
    var Sup= [];
    var nomBase = "cheminengagementhtpligne";
    var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13];
    workbook.xlsx.readFile('engagementhtp.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil2');
          var nomColonne3 = newworksheet.getColumn(3);
          var numFeuille = newworksheet.getColumn(4);
          var nomColonne = newworksheet.getColumn(5);
          var nomTable = newworksheet.getColumn(6);
          var nomColonne2 = newworksheet.getColumn(7);
          var numLigne = newworksheet.getColumn(8);
          var cheminparticulier = newworksheet.getColumn(9);
          var motcle = newworksheet.getColumn(10);
          var suppleant = newworksheet.getColumn(11);
  
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
        suppleant.eachCell(function(cell, rowNumber) {
          Sup.push(cell.value);
        });
  
          
            console.log(cheminp[0]);
            console.log(MotCle[0]);
            console.log(nomtable[0]);
            console.log(nomtable[1]);
            async.series([  
              function(cb){
                    Engagementhtp.deleteFromChemin(nomBase,cb);
                  },
                                    
            ],
            function(err, resultat){
              if(err){
                return res.view('Inovcom/erreur');
              }
              else{
                async.forEachSeries(r, function(lot, callback_reporting_suivant){
                  async.series([
                    function(cb){
                      Engagementhtp.delete(nomtable,lot,cb);
                    },
                    function(cb){
                      Engagementhtp.importcheminhtpligne(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                    },
                    // function(cb){
                    //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                    // },
                  ],
                  function(erroned, lotValues){
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
                          return res.view('Inovcom/erreur');
                        }
                       if(f==0)
                        {
                          return res.view('Inovcom/erreur');
                        }
                        else
                        {
                          // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                          return res.view('HTPengagement/accueilreportingengagementhtpsuivant3', {date : datetest});
                          // return res.view('HTPengagement/importHTPengagement_1', {date : datetest});
                        };
                    });
                  }
                });
  
              }
             
          });
        });
  },

     /********************************************************************************/
     insertcheminengagementhtpsuivant_3 : function(req,res)
     {
       var Excel = require('exceljs');
       var workbook = new Excel.Workbook();
       var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
      //  var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
       var datetest = req.param("date",0);
       var annee = datetest.substr(0, 4);
       var mois = datetest.substr(5, 2);
       var jour = datetest.substr(8, 2);
       var date = annee+mois+jour;
       var date_indus = jour+'.'+mois+'.'+annee;
       var datej_1 = annee+mois+jour -1;
       console.log(datej_1);
       var nomtable = [];
       var numligne = [];
       var numfeuille = [];
       var nomcolonne = [];
       var nomcolonne2 = [];
       var nomcolonne3 = [];
       console.log(date);
       var cheminp = [];
       var MotCle= [];
       var Sup= [];
       var nomBase = "cheminengagementhtpsales";
       var r = [0,1];
       workbook.xlsx.readFile('engagementhtp.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil11');
             var nomColonne3 = newworksheet.getColumn(3);
             var numFeuille = newworksheet.getColumn(4);
             var nomColonne = newworksheet.getColumn(5);
             var nomTable = newworksheet.getColumn(6);
             var nomColonne2 = newworksheet.getColumn(7);
             var numLigne = newworksheet.getColumn(8);
             var cheminparticulier = newworksheet.getColumn(9);
             var motcle = newworksheet.getColumn(10);
             var suppleant = newworksheet.getColumn(11);
     
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
           suppleant.eachCell(function(cell, rowNumber) {
             Sup.push(cell.value);
           });
     
             
               console.log(cheminp[0]);
               console.log(MotCle[0]);
               console.log(nomtable[0]);
               console.log(nomtable[1]);
               async.series([  
                 function(cb){
                       Engagementhtp.deleteFromChemin(nomBase,cb);
                     },
                                       
               ],
               function(err, resultat){
                 if(err){
                   return res.view('Inovcom/erreur');
                 }
                 else{
                   async.forEachSeries(r, function(lot, callback_reporting_suivant){
                     async.series([
                       function(cb){
                         Engagementhtp.delete(nomtable,lot,cb);
                       },
                       function(cb){
                         Engagementhtp.importcheminhtpsales(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                       },
                       // function(cb){
                       //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                       // },
                     ],
                     function(erroned, lotValues){
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
                             return res.view('Inovcom/erreur');
                           }
                          if(f==0)
                           {
                             return res.view('Inovcom/erreur');
                           }
                           else
                           {
                            //  return res.view('HTPengagement/accueilreportingengagementhtpsuivant4', {date : datetest});
                             // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                             return res.view('HTPengagement/importHTPengagement_1', {date : datetest});
                           };
                       });
                     }
                   });
     
                 }
                
             });
           });
     },

 
/************************************************************************************/
/*
*
*                 ANCIEN INSERTION CHEMIN
*
*
*/
/***********************************************************************************/
  //FONCTION POUR L'IMPORT DU CHEMIN UTILISER (TCD FACTURE MGEFI ET DEVIS MGEFI)
  insertcheminengagementhtpsuivant1 : function(req,res)
  {
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
    var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var date = annee+mois+jour;
    var date_indus = jour+'.'+mois+'.'+annee;
    var datej_1 = annee+mois+jour -1;
    console.log(datej_1);
    var nomtable = [];
    var numligne = [];
    var numfeuille = [];
    var nomcolonne = [];
    var nomcolonne2 = [];
    var nomcolonne3 = [];
    console.log(date);
    var cheminp = [];
    var MotCle= [];
    var Sup= [];
    var nomBase = "cheminengagementhtpfacture";
    var r = [0,1];
    workbook.xlsx.readFile('engagementhtp.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil2');
          var nomColonne3 = newworksheet.getColumn(3);
          var numFeuille = newworksheet.getColumn(4);
          var nomColonne = newworksheet.getColumn(5);
          var nomTable = newworksheet.getColumn(6);
          var nomColonne2 = newworksheet.getColumn(7);
          var numLigne = newworksheet.getColumn(8);
          var cheminparticulier = newworksheet.getColumn(9);
          var motcle = newworksheet.getColumn(10);
          var suppleant = newworksheet.getColumn(11);
  
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
        suppleant.eachCell(function(cell, rowNumber) {
          Sup.push(cell.value);
        });
  
          
            console.log(cheminp[0]);
            console.log(MotCle[0]);
            console.log(nomtable[0]);
            console.log(nomtable[1]);
            async.series([  
              function(cb){
                    Engagementhtp.deleteFromCheminfacture(nomBase,cb);
                  },
                                    
            ],
            function(err, resultat){
              if(err){
                return res.view('Inovcom/erreur');
              }
              else{
                async.forEachSeries(r, function(lot, callback_reporting_suivant){
                  async.series([
                    function(cb){
                      Engagementhtp.delete(nomtable,lot,cb);
                    },
                    function(cb){
                      Engagementhtp.importcheminhtpfacture(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                    },
                    // function(cb){
                    //   Engagementhtp.importcheminhtpfacture(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                    // },
                  ],
                  function(erroned, lotValues){
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
                          return res.view('Inovcom/erreur');
                        }
                       if(f==0)
                        {
                          return res.view('Inovcom/erreur');
                        }
                        else
                        {
                          return res.view('HTPengagement/accueilreportingengagementhtpsuivant1', {date : datetest});
                          
                        };
                    });
                  }
                });
  
              }
             
          });  
        });
  },
 /***********************************************************************************/
  //FONCTION POUR L'IMPORT DU CHEMIN UTILISER (TCD FACTURE MGEFI ET DEVIS MGEFI)
  insertcheminengagementhtpsuivant2 : function(req,res)
  {
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
    var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var date = annee+mois+jour;
    var date_indus = jour+'.'+mois+'.'+annee;
    var datej_1 = annee+mois+jour -1;
    console.log(datej_1);
    var nomtable = [];
    var numligne = [];
    var numfeuille = [];
    var nomcolonne = [];
    var nomcolonne2 = [];
    var nomcolonne3 = [];
    console.log(date);
    var cheminp = [];
    var MotCle= [];
    var Sup= [];
    var nomBase = "cheminengagementhtpdevis";
    var r = [0,1];
    workbook.xlsx.readFile('engagementhtp.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil3');
          var nomColonne3 = newworksheet.getColumn(3);
          var numFeuille = newworksheet.getColumn(4);
          var nomColonne = newworksheet.getColumn(5);
          var nomTable = newworksheet.getColumn(6);
          var nomColonne2 = newworksheet.getColumn(7);
          var numLigne = newworksheet.getColumn(8);
          var cheminparticulier = newworksheet.getColumn(9);
          var motcle = newworksheet.getColumn(10);
          var suppleant = newworksheet.getColumn(11);
  
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
        suppleant.eachCell(function(cell, rowNumber) {
          Sup.push(cell.value);
        });
  
          
            console.log(cheminp[0]);
            console.log(MotCle[0]);
            console.log(nomtable[0]);
            console.log(nomtable[1]);
            async.series([  
              function(cb){
                    Engagementhtp.deleteFromChemindevis(nomBase,cb);
                  },
                                    
            ],
            function(err, resultat){
              if(err){
                return res.view('Inovcom/erreur');
              }
              else{
                async.forEachSeries(r, function(lot, callback_reporting_suivant){
                  async.series([
                    function(cb){
                      Engagementhtp.delete(nomtable,lot,cb);
                    },
                    function(cb){
                      Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                    },
                    // function(cb){
                    //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                    // },
                  ],
                  function(erroned, lotValues){
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
                          return res.view('Inovcom/erreur');
                        }
                       if(f==0)
                        {
                          return res.view('Inovcom/erreur');
                        }
                        else
                        {
                          // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                          return res.view('HTPengagement/accueilreportingengagementhtpsuivant2', {date : datetest});
                        };
                    });
                  }
                });
  
              }
             
          });
        });
  },

   /********************************************************************************/
   insertcheminengagementhtpsuivant4 : function(req,res)
   {
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
    //  var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
    var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
     var datetest = req.param("date",0);
     var annee = datetest.substr(0, 4);
     var mois = datetest.substr(5, 2);
     var jour = datetest.substr(8, 2);
     var date = annee+mois+jour;
     var date_indus = jour+'.'+mois+'.'+annee;
     var datej_1 = annee+mois+jour -1;
     console.log(datej_1);
     var nomtable = [];
     var numligne = [];
     var numfeuille = [];
     var nomcolonne = [];
     var nomcolonne2 = [];
     var nomcolonne3 = [];
     console.log(date);
     var cheminp = [];
     var MotCle= [];
     var Sup= [];
     var nomBase = "cheminengagementhtpdevisj2";
     var r = [0];
     workbook.xlsx.readFile('engagementhtp.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil5');
           var nomColonne3 = newworksheet.getColumn(3);
           var numFeuille = newworksheet.getColumn(4);
           var nomColonne = newworksheet.getColumn(5);
           var nomTable = newworksheet.getColumn(6);
           var nomColonne2 = newworksheet.getColumn(7);
           var numLigne = newworksheet.getColumn(8);
           var cheminparticulier = newworksheet.getColumn(9);
           var motcle = newworksheet.getColumn(10);
           var suppleant = newworksheet.getColumn(11);
   
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
         suppleant.eachCell(function(cell, rowNumber) {
           Sup.push(cell.value);
         });
   
           
             console.log(cheminp[0]);
             console.log(MotCle[0]);
             console.log(nomtable[0]);
             console.log(nomtable[1]);
             async.series([  
               function(cb){
                     Engagementhtp.deleteFromChemin(nomBase,cb);
                   },
                                     
             ],
             function(err, resultat){
               if(err){
                 return res.view('Inovcom/erreur');
               }
               else{
                 async.forEachSeries(r, function(lot, callback_reporting_suivant){
                   async.series([
                     function(cb){
                       Engagementhtp.delete(nomtable,lot,cb);
                     },
                     function(cb){
                       Engagementhtp.importcheminhtpligne(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                     },
                     // function(cb){
                     //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                     // },
                   ],
                   function(erroned, lotValues){
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
                           return res.view('Inovcom/erreur');
                         }
                        if(f==0)
                         {
                           return res.view('Inovcom/erreur');
                         }
                         else
                         {
                           return res.view('HTPengagement/accueilreportingengagementhtpsuivant4', {date : datetest});
                         };
                     });
                   }
                 });
   
               }
              
           });
         });
   },
    /********************************************************************************/
  insertcheminengagementhtpsuivant5 : function(req,res)
  {
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
    var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var date = annee+mois+jour;
    var date_indus = jour+'.'+mois+'.'+annee;
    var datej_1 = annee+mois+jour -1;
    console.log(datej_1);
    var nomtable = [];
    var numligne = [];
    var numfeuille = [];
    var nomcolonne = [];
    var nomcolonne2 = [];
    var nomcolonne3 = [];
    console.log(date);
    var cheminp = [];
    var MotCle= [];
    var Sup= [];
    var nomBase = "cheminengagementhtpdevisj5";
    var r = [0];
    workbook.xlsx.readFile('engagementhtp.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil6');
          var nomColonne3 = newworksheet.getColumn(3);
          var numFeuille = newworksheet.getColumn(4);
          var nomColonne = newworksheet.getColumn(5);
          var nomTable = newworksheet.getColumn(6);
          var nomColonne2 = newworksheet.getColumn(7);
          var numLigne = newworksheet.getColumn(8);
          var cheminparticulier = newworksheet.getColumn(9);
          var motcle = newworksheet.getColumn(10);
          var suppleant = newworksheet.getColumn(11);
  
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
        suppleant.eachCell(function(cell, rowNumber) {
          Sup.push(cell.value);
        });
  
          
            console.log(cheminp[0]);
            console.log(MotCle[0]);
            console.log(nomtable[0]);
            console.log(nomtable[1]);
            async.series([  
              function(cb){
                    Engagementhtp.deleteFromChemin(nomBase,cb);
                  },
                                    
            ],
            function(err, resultat){
              if(err){
                return res.view('Inovcom/erreur');
              }
              else{
                async.forEachSeries(r, function(lot, callback_reporting_suivant){
                  async.series([
                    function(cb){
                      Engagementhtp.delete(nomtable,lot,cb);
                    },
                    function(cb){
                      Engagementhtp.importcheminhtpligne(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                    },
                    // function(cb){
                    //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                    // },
                  ],
                  function(erroned, lotValues){
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
                          return res.view('Inovcom/erreur');
                        }
                       if(f==0)
                        {
                          return res.view('Inovcom/erreur');
                        }
                        else
                        {
                          // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                          return res.view('HTPengagement/accueilreportingengagementhtpsuivant5', {date : datetest});
                        };
                    });
                  }
                });
  
              }
             
          });
        });
  },
   /********************************************************************************/
   insertcheminengagementhtpsuivant6 : function(req,res)
   {
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
    //  var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
    var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
     var datetest = req.param("date",0);
     var annee = datetest.substr(0, 4);
     var mois = datetest.substr(5, 2);
     var jour = datetest.substr(8, 2);
     var date = annee+mois+jour;
     var date_indus = jour+'.'+mois+'.'+annee;
     var datej_1 = annee+mois+jour -1;
     console.log(datej_1);
     var nomtable = [];
     var numligne = [];
     var numfeuille = [];
     var nomcolonne = [];
     var nomcolonne2 = [];
     var nomcolonne3 = [];
     console.log(date);
     var cheminp = [];
     var MotCle= [];
     var Sup= [];
     var nomBase = "cheminengagementhtpfacmgj2";
     var r = [0];
     workbook.xlsx.readFile('engagementhtp.xlsx')
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
           var suppleant = newworksheet.getColumn(11);
   
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
         suppleant.eachCell(function(cell, rowNumber) {
           Sup.push(cell.value);
         });
   
           
             console.log(cheminp[0]);
             console.log(MotCle[0]);
             console.log(nomtable[0]);
             console.log(nomtable[1]);
             async.series([  
               function(cb){
                     Engagementhtp.deleteFromChemin(nomBase,cb);
                   },
                                     
             ],
             function(err, resultat){
               if(err){
                 return res.view('Inovcom/erreur');
               }
               else{
                 async.forEachSeries(r, function(lot, callback_reporting_suivant){
                   async.series([
                     function(cb){
                       Engagementhtp.delete(nomtable,lot,cb);
                     },
                     function(cb){
                       Engagementhtp.importcheminhtpligne(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                     },
                     // function(cb){
                     //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                     // },
                   ],
                   function(erroned, lotValues){
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
                           return res.view('Inovcom/erreur');
                         }
                        if(f==0)
                         {
                           return res.view('Inovcom/erreur');
                         }
                         else
                         {
                          //  return res.view('HTPengagement/importHTPengagement', {date : datetest});
                           return res.view('HTPengagement/accueilreportingengagementhtpsuivant6', {date : datetest});
                         };
                     });
                   }
                 });
   
               }
              
           });
         });
   },
    /********************************************************************************/
  insertcheminengagementhtpsuivant7 : function(req,res)
  {
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
    var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var date = annee+mois+jour;
    var date_indus = jour+'.'+mois+'.'+annee;
    var datej_1 = annee+mois+jour -1;
    console.log(datej_1);
    var nomtable = [];
    var numligne = [];
    var numfeuille = [];
    var nomcolonne = [];
    var nomcolonne2 = [];
    var nomcolonne3 = [];
    console.log(date);
    var cheminp = [];
    var MotCle= [];
    var Sup= [];
    var nomBase = "cheminengagementhtpfacmgj5";
    var r = [0];
    workbook.xlsx.readFile('engagementhtp.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil8');
          var nomColonne3 = newworksheet.getColumn(3);
          var numFeuille = newworksheet.getColumn(4);
          var nomColonne = newworksheet.getColumn(5);
          var nomTable = newworksheet.getColumn(6);
          var nomColonne2 = newworksheet.getColumn(7);
          var numLigne = newworksheet.getColumn(8);
          var cheminparticulier = newworksheet.getColumn(9);
          var motcle = newworksheet.getColumn(10);
          var suppleant = newworksheet.getColumn(11);
  
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
        suppleant.eachCell(function(cell, rowNumber) {
          Sup.push(cell.value);
        });
  
          
            console.log(cheminp[0]);
            console.log(MotCle[0]);
            console.log(nomtable[0]);
            console.log(nomtable[1]);
            async.series([  
              function(cb){
                    Engagementhtp.deleteFromChemin(nomBase,cb);
                  },
                                    
            ],
            function(err, resultat){
              if(err){
                return res.view('Inovcom/erreur');
              }
              else{
                async.forEachSeries(r, function(lot, callback_reporting_suivant){
                  async.series([
                    function(cb){
                      Engagementhtp.delete(nomtable,lot,cb);
                    },
                    function(cb){
                      Engagementhtp.importcheminhtpligne(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                    },
                    // function(cb){
                    //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                    // },
                  ],
                  function(erroned, lotValues){
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
                          return res.view('Inovcom/erreur');
                        }
                       if(f==0)
                        {
                          return res.view('Inovcom/erreur');
                        }
                        else
                        {
                          return res.view('HTPengagement/accueilreportingengagementhtpsuivant7', {date : datetest});
                        };
                    });
                  }
                });
  
              }
             
          });
        });
  },
 /********************************************************************************/
 insertcheminengagementhtpsuivant8 : function(req,res)
 {
   var Excel = require('exceljs');
   var workbook = new Excel.Workbook();
   // var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
   var table = ['/dev/pro/Retour_Easytech_'];
   var datetest = req.param("date",0);
   var annee = datetest.substr(0, 4);
   var mois = datetest.substr(5, 2);
   var jour = datetest.substr(8, 2);
   var date = annee+mois+jour;
   var date_indus = jour+'.'+mois+'.'+annee;
   var datej_1 = annee+mois+jour -1;
   console.log(datej_1);
   var nomtable = [];
   var numligne = [];
   var numfeuille = [];
   var nomcolonne = [];
   var nomcolonne2 = [];
   var nomcolonne3 = [];
   console.log(date);
   var cheminp = [];
   var MotCle= [];
   var Sup= [];
   var nomBase = "cheminengagementhtpcotlamiej2";
   var r = [0];
   workbook.xlsx.readFile('engagementhtp.xlsx')
       .then(function() {
         var newworksheet = workbook.getWorksheet('Feuil9');
         var nomColonne3 = newworksheet.getColumn(3);
         var numFeuille = newworksheet.getColumn(4);
         var nomColonne = newworksheet.getColumn(5);
         var nomTable = newworksheet.getColumn(6);
         var nomColonne2 = newworksheet.getColumn(7);
         var numLigne = newworksheet.getColumn(8);
         var cheminparticulier = newworksheet.getColumn(9);
         var motcle = newworksheet.getColumn(10);
         var suppleant = newworksheet.getColumn(11);
 
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
       suppleant.eachCell(function(cell, rowNumber) {
         Sup.push(cell.value);
       });
 
         
           console.log(cheminp[0]);
           console.log(MotCle[0]);
           console.log(nomtable[0]);
           console.log(nomtable[1]);
           async.series([  
             function(cb){
                   Engagementhtp.deleteFromChemin(nomBase,cb);
                 },
                                   
           ],
           function(err, resultat){
             if(err){
               return res.view('Inovcom/erreur');
             }
             else{
               async.forEachSeries(r, function(lot, callback_reporting_suivant){
                 async.series([
                   function(cb){
                     Engagementhtp.delete(nomtable,lot,cb);
                   },
                   function(cb){
                     Engagementhtp.importcheminhtpligne(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                   },
                   // function(cb){
                   //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                   // },
                 ],
                 function(erroned, lotValues){
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
                         return res.view('Inovcom/erreur');
                       }
                      if(f==0)
                       {
                         return res.view('Inovcom/erreur');
                       }
                       else
                       {
                         return res.view('HTPengagement/accueilreportingengagementhtpsuivant8', {date : datetest});
                       };
                   });
                 }
               });
 
             }
            
         });
       });
 },
 /********************************************************************************/
 insertcheminengagementhtpsuivant9 : function(req,res)
 {
   var Excel = require('exceljs');
   var workbook = new Excel.Workbook();
   // var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
   var table = ['/dev/pro/Retour_Easytech_'];
   var datetest = req.param("date",0);
   var annee = datetest.substr(0, 4);
   var mois = datetest.substr(5, 2);
   var jour = datetest.substr(8, 2);
   var date = annee+mois+jour;
   var date_indus = jour+'.'+mois+'.'+annee;
   var datej_1 = annee+mois+jour -1;
   console.log(datej_1);
   var nomtable = [];
   var numligne = [];
   var numfeuille = [];
   var nomcolonne = [];
   var nomcolonne2 = [];
   var nomcolonne3 = [];
   console.log(date);
   var cheminp = [];
   var MotCle= [];
   var Sup= [];
   var nomBase = "cheminengagementhtpcotlamiej5";
   var r = [0];
   workbook.xlsx.readFile('engagementhtp.xlsx')
       .then(function() {
         var newworksheet = workbook.getWorksheet('Feuil10');
         var nomColonne3 = newworksheet.getColumn(3);
         var numFeuille = newworksheet.getColumn(4);
         var nomColonne = newworksheet.getColumn(5);
         var nomTable = newworksheet.getColumn(6);
         var nomColonne2 = newworksheet.getColumn(7);
         var numLigne = newworksheet.getColumn(8);
         var cheminparticulier = newworksheet.getColumn(9);
         var motcle = newworksheet.getColumn(10);
         var suppleant = newworksheet.getColumn(11);
 
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
       suppleant.eachCell(function(cell, rowNumber) {
         Sup.push(cell.value);
       });
 
         
           console.log(cheminp[0]);
           console.log(MotCle[0]);
           console.log(nomtable[0]);
           console.log(nomtable[1]);
           async.series([  
             function(cb){
                   Engagementhtp.deleteFromChemin(nomBase,cb);
                 },
                                   
           ],
           function(err, resultat){
             if(err){
               return res.view('Inovcom/erreur');
             }
             else{
               async.forEachSeries(r, function(lot, callback_reporting_suivant){
                 async.series([
                   function(cb){
                     Engagementhtp.delete(nomtable,lot,cb);
                   },
                   function(cb){
                     Engagementhtp.importcheminhtpligne(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                   },
                   // function(cb){
                   //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                   // },
                 ],
                 function(erroned, lotValues){
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
                         return res.view('Inovcom/erreur');
                       }
                      if(f==0)
                       {
                         return res.view('Inovcom/erreur');
                       }
                       else
                       {
                         // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                         return res.view('HTPengagement/accueilreportingengagementhtpsuivant9', {date : datetest});
                       };
                   });
                 }
               });
 
             }
            
         });
       });
 },

  
       /********************************************************************************/
       insertcheminengagementhtpsuivant11 : function(req,res)
       {
         var Excel = require('exceljs');
         var workbook = new Excel.Workbook();
         // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
         var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
         var datetest = req.param("date",0);
         var annee = datetest.substr(0, 4);
         var mois = datetest.substr(5, 2);
         var jour = datetest.substr(8, 2);
         var date = annee+mois+jour;
         var date_indus = jour+'.'+mois+'.'+annee;
         var datej_1 = annee+mois+jour -1;
         console.log(datej_1);
         var nomtable = [];
         var numligne = [];
         var numfeuille = [];
         var nomcolonne = [];
         var nomcolonne2 = [];
         var nomcolonne3 = [];
         console.log(date);
         var cheminp = [];
         var MotCle= [];
         var Sup= [];
         var nomBase = "cheminengagementhtpsalesstock";
         var r = [0];
         workbook.xlsx.readFile('engagementhtp.xlsx')
             .then(function() {
               var newworksheet = workbook.getWorksheet('Feuil12');
               var nomColonne3 = newworksheet.getColumn(3);
               var numFeuille = newworksheet.getColumn(4);
               var nomColonne = newworksheet.getColumn(5);
               var nomTable = newworksheet.getColumn(6);
               var nomColonne2 = newworksheet.getColumn(7);
               var numLigne = newworksheet.getColumn(8);
               var cheminparticulier = newworksheet.getColumn(9);
               var motcle = newworksheet.getColumn(10);
               var suppleant = newworksheet.getColumn(11);
       
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
             suppleant.eachCell(function(cell, rowNumber) {
               Sup.push(cell.value);
             });
       
               
                 console.log(cheminp[0]);
                 console.log(MotCle[0]);
                 console.log(nomtable[0]);
                 console.log(nomtable[1]);
                 async.series([  
                   function(cb){
                         Engagementhtp.deleteFromChemin(nomBase,cb);
                       },
                                         
                 ],
                 function(err, resultat){
                   if(err){
                     return res.view('Inovcom/erreur');
                   }
                   else{
                     async.forEachSeries(r, function(lot, callback_reporting_suivant){
                       async.series([
                         function(cb){
                           Engagementhtp.delete(nomtable,lot,cb);
                         },
                         function(cb){
                           Engagementhtp.importcheminhtpsales(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                         },
                         // function(cb){
                         //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                         // },
                       ],
                       function(erroned, lotValues){
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
                               return res.view('Inovcom/erreur');
                             }
                            if(f==0)
                             {
                               return res.view('Inovcom/erreur');
                             }
                             else
                             {
                              //  return res.view('HTPengagement/accueilreportingengagementhtpsuivant11', {date : datetest});
                               return res.view('HTPengagement/importHTPengagement', {date : datetest});
                             };
                         });
                       }
                     });
       
                   }
                  
               });
             });
       },


 /********************************************************************************/
      insertcheminengagementhtpsuivant12 : function(req,res)
      {
        var Excel = require('exceljs');
        var workbook = new Excel.Workbook();
        // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
        var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
        var datetest = req.param("date",0);
        var annee = datetest.substr(0, 4);
        var mois = datetest.substr(5, 2);
        var jour = datetest.substr(8, 2);
        var date = annee+mois+jour;
        var date_indus = jour+'.'+mois+'.'+annee;
        var datej_1 = annee+mois+jour -1;
        console.log(datej_1);
        var nomtable = [];
        var numligne = [];
        var numfeuille = [];
        var nomcolonne = [];
        var nomcolonne2 = [];
        var nomcolonne3 = [];
        console.log(date);
        var cheminp = [];
        var MotCle= [];
        var Sup= [];
        var nomBase = "cheminengagementhtpstockfacmg";
        var r = [0];
        workbook.xlsx.readFile('engagementhtp.xlsx')
            .then(function() {
              var newworksheet = workbook.getWorksheet('Feuil13');
              var nomColonne3 = newworksheet.getColumn(3);
              var numFeuille = newworksheet.getColumn(4);
              var nomColonne = newworksheet.getColumn(5);
              var nomTable = newworksheet.getColumn(6);
              var nomColonne2 = newworksheet.getColumn(7);
              var numLigne = newworksheet.getColumn(8);
              var cheminparticulier = newworksheet.getColumn(9);
              var motcle = newworksheet.getColumn(10);
              var suppleant = newworksheet.getColumn(11);
      
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
            suppleant.eachCell(function(cell, rowNumber) {
              Sup.push(cell.value);
            });
      
              
                console.log(cheminp[0]);
                console.log(MotCle[0]);
                console.log(nomtable[0]);
                console.log(nomtable[1]);
                async.series([  
                  function(cb){
                        Engagementhtp.deleteFromChemin(nomBase,cb);
                      },
                                        
                ],
                function(err, resultat){
                  if(err){
                    return res.view('Inovcom/erreur');
                  }
                  else{
                    async.forEachSeries(r, function(lot, callback_reporting_suivant){
                      async.series([
                        function(cb){
                          Engagementhtp.delete(nomtable,lot,cb);
                        },
                        function(cb){
                          Engagementhtp.importcheminhtpstockfacdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                        },
                        // function(cb){
                        //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                        // },
                      ],
                      function(erroned, lotValues){
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
                              return res.view('Inovcom/erreur');
                            }
                           if(f==0)
                            {
                              return res.view('Inovcom/erreur');
                            }
                            else
                            {
                              return res.view('HTPengagement/accueilreportingengagementhtpsuivant12', {date : datetest});
                              // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                            };
                        });
                      }
                    });
      
                  }
                 
              });
            });
      },
/********************************************************************************/
insertcheminengagementhtpsuivant13 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
  var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtpstockdevis";
  var r = [0];
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil14');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

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
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          console.log(nomtable[0]);
          console.log(nomtable[1]);
          async.series([  
            function(cb){
                  Engagementhtp.deleteFromChemin(nomBase,cb);
                },
                                  
          ],
          function(err, resultat){
            if(err){
              return res.view('Inovcom/erreur');
            }
            else{
              async.forEachSeries(r, function(lot, callback_reporting_suivant){
                async.series([
                  function(cb){
                    Engagementhtp.delete(nomtable,lot,cb);
                  },
                  function(cb){
                    Engagementhtp.importcheminhtpstockfacdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                  },
                  // function(cb){
                  //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                  // },
                ],
                function(erroned, lotValues){
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
                        return res.view('Inovcom/erreur');
                      }
                      if(f==0)
                      {
                        return res.view('Inovcom/erreur');
                      }
                      else
                      {
                        return res.view('HTPengagement/accueilreportingengagementhtpsuivant13', {date : datetest});
                        // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                      };
                  });
                }
              });

            }
            
        });
      });
},
/********************************************************************************/
insertcheminengagementhtpsuivant14 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
  var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtpfacmgtnontj";
  var r = [0];
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil15');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

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
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          console.log(nomtable[0]);
          console.log(nomtable[1]);
          async.series([  
            function(cb){
                  Engagementhtp.deleteFromChemin(nomBase,cb);
                },
                                  
          ],
          function(err, resultat){
            if(err){
              return res.view('Inovcom/erreur');
            }
            else{
              async.forEachSeries(r, function(lot, callback_reporting_suivant){
                async.series([
                  function(cb){
                    Engagementhtp.delete(nomtable,lot,cb);
                  },
                  function(cb){
                    Engagementhtp.importcheminhtpstockfacdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                  },
                  // function(cb){
                  //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                  // },
                ],
                function(erroned, lotValues){
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
                        return res.view('Inovcom/erreur');
                      }
                      if(f==0)
                      {
                        return res.view('Inovcom/erreur');
                      }
                      else
                      {
                        return res.view('HTPengagement/accueilreportingengagementhtpsuivant14', {date : datetest});
                        // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                      };
                  });
                }
              });

            }
            
        });
      });
},
/*********************************************************************************************/
insertcheminengagementhtpsuivant15 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
  var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtpfacmgtnontj1";
  var r = [0];
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil16');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

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
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          console.log(nomtable[0]);
          console.log(nomtable[1]);
          async.series([  
            function(cb){
                  Engagementhtp.deleteFromChemin(nomBase,cb);
                },
                                  
          ],
          function(err, resultat){
            if(err){
              return res.view('Inovcom/erreur');
            }
            else{
              async.forEachSeries(r, function(lot, callback_reporting_suivant){
                async.series([
                  function(cb){
                    Engagementhtp.delete(nomtable,lot,cb);
                  },
                  function(cb){
                    Engagementhtp.importcheminhtpstockfacdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                  },
                  // function(cb){
                  //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                  // },
                ],
                function(erroned, lotValues){
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
                        return res.view('Inovcom/erreur');
                      }
                      if(f==0)
                      {
                        return res.view('Inovcom/erreur');
                      }
                      else
                      {
                        return res.view('HTPengagement/accueilreportingengagementhtpsuivant15', {date : datetest});
                        // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                      };
                  });
                }
              });

            }
            
        });
      });
},
/*****************************************************************************/
insertcheminengagementhtpsuivant16 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
  var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtpfacmgtnontj2";
  var r = [0];
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil17');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

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
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          console.log(nomtable[0]);
          console.log(nomtable[1]);
          async.series([  
            function(cb){
                  Engagementhtp.deleteFromChemin(nomBase,cb);
                },
                                  
          ],
          function(err, resultat){
            if(err){
              return res.view('Inovcom/erreur');
            }
            else{
              async.forEachSeries(r, function(lot, callback_reporting_suivant){
                async.series([
                  function(cb){
                    Engagementhtp.delete(nomtable,lot,cb);
                  },
                  function(cb){
                    Engagementhtp.importcheminhtpstockfacdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                  },
                  // function(cb){
                  //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                  // },
                ],
                function(erroned, lotValues){
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
                        return res.view('Inovcom/erreur');
                      }
                      if(f==0)
                      {
                        return res.view('Inovcom/erreur');
                      }
                      else
                      {
                        return res.view('HTPengagement/accueilreportingengagementhtpsuivant16', {date : datetest});
                        // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                      };
                  });
                }
              });

            }
            
        });
      });
},
/**************************************************************************************/
insertcheminengagementhtpsuivant17 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
  var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtpfacmgtnontj5";
  var r = [0];
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil18');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

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
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          console.log(nomtable[0]);
          console.log(nomtable[1]);
          async.series([  
            function(cb){
                  Engagementhtp.deleteFromChemin(nomBase,cb);
                },
                                  
          ],
          function(err, resultat){
            if(err){
              return res.view('Inovcom/erreur');
            }
            else{
              async.forEachSeries(r, function(lot, callback_reporting_suivant){
                async.series([
                  function(cb){
                    Engagementhtp.delete(nomtable,lot,cb);
                  },
                  function(cb){
                    Engagementhtp.importcheminhtpstockfacdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                  },
                  // function(cb){
                  //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                  // },
                ],
                function(erroned, lotValues){
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
                        return res.view('Inovcom/erreur');
                      }
                      if(f==0)
                      {
                        return res.view('Inovcom/erreur');
                      }
                      else
                      {
                        return res.view('HTPengagement/accueilreportingengagementhtpsuivant17', {date : datetest});
                        // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                      };
                  });
                }
              });

            }
            
        });
      });
},
/**************************************************************************************/
insertcheminengagementhtpsuivant18 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
  var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtpdevitnontj";
  var r = [0];
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil19');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

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
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          console.log(nomtable[0]);
          console.log(nomtable[1]);
          async.series([  
            function(cb){
                  Engagementhtp.deleteFromChemin(nomBase,cb);
                },
                                  
          ],
          function(err, resultat){
            if(err){
              return res.view('Inovcom/erreur');
            }
            else{
              async.forEachSeries(r, function(lot, callback_reporting_suivant){
                async.series([
                  function(cb){
                    Engagementhtp.delete(nomtable,lot,cb);
                  },
                  function(cb){
                    Engagementhtp.importcheminhtpstockfacdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                  },
                  // function(cb){
                  //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                  // },
                ],
                function(erroned, lotValues){
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
                        return res.view('Inovcom/erreur');
                      }
                      if(f==0)
                      {
                        return res.view('Inovcom/erreur');
                      }
                      else
                      {
                        return res.view('HTPengagement/accueilreportingengagementhtpsuivant18', {date : datetest});
                        // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                      };
                  });
                }
              });

            }
            
        });
      });
},
/**************************************************************************************/
insertcheminengagementhtpsuivant19 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
  var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtpdevitnontj1";
  var r = [0];
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil20');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

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
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          console.log(nomtable[0]);
          console.log(nomtable[1]);
          async.series([  
            function(cb){
                  Engagementhtp.deleteFromChemin(nomBase,cb);
                },
                                  
          ],
          function(err, resultat){
            if(err){
              return res.view('Inovcom/erreur');
            }
            else{
              async.forEachSeries(r, function(lot, callback_reporting_suivant){
                async.series([
                  function(cb){
                    Engagementhtp.delete(nomtable,lot,cb);
                  },
                  function(cb){
                    Engagementhtp.importcheminhtpstockfacdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                  },
                  // function(cb){
                  //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                  // },
                ],
                function(erroned, lotValues){
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
                        return res.view('Inovcom/erreur');
                      }
                      if(f==0)
                      {
                        return res.view('Inovcom/erreur');
                      }
                      else
                      {
                        return res.view('HTPengagement/accueilreportingengagementhtpsuivant19', {date : datetest});
                        // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                      };
                  });
                }
              });

            }
            
        });
      });
},
/**************************************************************************************/
insertcheminengagementhtpsuivant20 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
  var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtpdevitnontj2";
  var r = [0];
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil21');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

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
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          console.log(nomtable[0]);
          console.log(nomtable[1]);
          async.series([  
            function(cb){
                  Engagementhtp.deleteFromChemin(nomBase,cb);
                },
                                  
          ],
          function(err, resultat){
            if(err){
              return res.view('Inovcom/erreur');
            }
            else{
              async.forEachSeries(r, function(lot, callback_reporting_suivant){
                async.series([
                  function(cb){
                    Engagementhtp.delete(nomtable,lot,cb);
                  },
                  function(cb){
                    Engagementhtp.importcheminhtpstockfacdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                  },
                  // function(cb){
                  //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                  // },
                ],
                function(erroned, lotValues){
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
                        return res.view('Inovcom/erreur');
                      }
                      if(f==0)
                      {
                        return res.view('Inovcom/erreur');
                      }
                      else
                      {
                        return res.view('HTPengagement/accueilreportingengagementhtpsuivant20', {date : datetest});
                        // return res.view('HTPengagement/importHTPengagement', {date : datetest});
                      };
                  });
                }
              });

            }
            
        });
      });
},
/**************************************************************************************/
insertcheminengagementhtpsuivant21 : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  // var table = ['\\\\10.128.1.2\\bpo_almerys\\00-TOUS\\06-DOSSIER POLE\\01-HTP\\05- REPORTING\\03-HTP\\DOC_HTP\\'];
  var table = ['/dev/prod/00-TOUS/06-DOSSIER POLE/01-HTP/05- REPORTING/03-HTP/DOC_HTP/'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtpdevitnontj5";
  var r = [0];
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil22');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

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
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          console.log(nomtable[0]);
          console.log(nomtable[1]);
          async.series([  
            function(cb){
                  Engagementhtp.deleteFromChemin(nomBase,cb);
                },
                                  
          ],
          function(err, resultat){
            if(err){
              return res.view('Inovcom/erreur');
            }
            else{
              async.forEachSeries(r, function(lot, callback_reporting_suivant){
                async.series([
                  function(cb){
                    Engagementhtp.delete(nomtable,lot,cb);
                  },
                  function(cb){
                    Engagementhtp.importcheminhtpstockfacdevis(table,cheminp,date,MotCle,lot,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,nomBase,cb);
                  },
                  // function(cb){
                  //   Engagementhtp.importcheminhtpdevis(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],nomcolonne2[1],nomcolonne3[1],nomBase,cb);
                  // },
                ],
                function(erroned, lotValues){
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
                        return res.view('Inovcom/erreur');
                      }
                      if(f==0)
                      {
                        return res.view('Inovcom/erreur');
                      }
                      else
                      {
                        // return res.view('HTPengagement/accueilreportingengagementhtpsuivant7', {date : datetest});
                        return res.view('HTPengagement/importHTPengagement', {date : datetest});
                      };
                  });
                }
              });

            }
            
        });
      });
},
/***********************************************************************************/  
/*
*
*
*                              EXPORT HTP
* 
* 
*  
/***********************************************************************************/
  //FONCTION POUR L'EXPORT TACHES TRAITEES
  exporthtpengagement : function (req, res) {
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
      var feuille = annee+mois+'_EASY';
      console.log(feuille);
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
          Engagementhtp.recupdata("htptri16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpfacmg16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpdevi16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpsales16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpflux16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htprejet16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpcotlamie16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htptrifin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpfacmgfin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpdevifin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpsalesfin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpfluxfin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htprejetfin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpcotlamiefin",callback);
        },
        function (callback) {
          Engagementhtp.recupdatasum("htpcotite16",callback);
        },
        function (callback) {
          Engagementhtp.recupdatasum("htpcotitefin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpfaclamie16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpacs16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpfaclamiefin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpacsfin",callback);
        },
       
      ],function(err,result){
        if(err) return res.badRequest(err);
        console.log("Count OK 1==> " + result[0].ok);
        async.series([
          function (callback) {
            Engagementhtp.ecrituredata16tri(result[0],"htptri16",date_export,feuille,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16facM(result[1],"htpfacmg16",date_export,feuille,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16devi(result[2],"htpdevi16",date_export,feuille,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16sales(result[3],"htpsales16",date_export,feuille,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16flux(result[4],"htpflux16",date_export,feuille,callback);
          },
          function (callback) {
              Engagementhtp.ecrituredata16rejet(result[5],"htprejet16",date_export,feuille,callback);
            },
          function (callback) {
          Engagementhtp.ecrituredata16cotlamie(result[6],"htpcotlamie16",date_export,feuille,callback);
          },
          function (callback) {
              Engagementhtp.ecrituredatafinptri(result[7],"htptrifin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpfacM(result[8],"htpfacmgfin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpdevi(result[9],"htpdevifin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpsales(result[10],"htpsalesfin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpflux(result[11],"htpfluxfin",date_export,feuille,callback);
            },
            function (callback) {
                Engagementhtp.ecrituredatafinprejet(result[12],"htprejetfin",date_export,feuille,callback);
              },
            function (callback) {
            Engagementhtp.ecrituredatafinpcotlamie(result[13],"htpcotlamiefin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16cotite(result[14],"htpcotite16",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpcotite(result[15],"htpcotitefin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16faclamie(result[16],"htpfaclamie16",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16acs(result[17],"htpacs16",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpfaclamie(result[18],"htpfaclamiefin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpacs(result[19],"htpacsfin",date_export,feuille,callback);
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
              res.view('HTPengagement/exportHTPengagementsuivant_1', {date : datetest});
              // res.view('Retour/succes');
            }

        })
      })
    },
/**********************************************************************************************************/ 
//FONCTION EXPORT SUIVANT 1 (taches traitees)
exporthtpengagementsuivant_1 : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var feuille = annee+mois+'_EASY';
  console.log(feuille);
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
    // function (callback) {
    //   Engagementhtp.recupdata("htptrij2",callback);
    // },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevij2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpsales16",callback);//htpsalesj2
    },
    function (callback) {
      Engagementhtp.recupdata("htpflux16",callback);//htpfluxj2
    },
    function (callback) {
      Engagementhtp.recupdata("htprejet16",callback);//htprejetj2
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamiej2",callback);
    },
    // function (callback) {
    //   Engagementhtp.recupdata("htptrij5",callback);
    // },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevij5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpsales16",callback);//htpsalesj5
    },
    function (callback) {
      Engagementhtp.recupdata("htpflux16",callback);//htpfluxj5
    },
    function (callback) {
      Engagementhtp.recupdata("htprejet16",callback);//htprejetj5
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamiej5",callback);
    },
    function (callback) {
      Engagementhtp.recupdatasum("htpcotite16",callback);//htpcotitej2
    },
    function (callback) {
      Engagementhtp.recupdatasum("htpcotite16",callback);//htpcotitej5
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamiej2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacsj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamiej5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacsj5",callback);
    },

  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      // function (callback) {
      //   Engagementhtp.ecrituredataj2tri(result[0],"htptrij2",date_export,feuille,callback);
      // },
      function (callback) {
        Engagementhtp.ecrituredataj2facM(result[0],"htpfacmgj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataj2devi(result[1],"htpdevij2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataj2sales(result[2],"htpsales16",date_export,feuille,callback);//htpsalesj2
      },
      function (callback) {
        Engagementhtp.ecrituredataj2flux(result[3],"htpflux16",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj2rejet(result[4],"htprejet16",date_export,feuille,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredataj2cotlamie(result[5],"htpcotlamiej2",date_export,feuille,callback);
      },
      // function (callback) {
      //     Engagementhtp.ecrituredataj5tri(result[7],"htptrij5",date_export,feuille,callback);
      // },
      function (callback) {
          Engagementhtp.ecrituredataj5facM(result[6],"htpfacmgj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5devi(result[7],"htpdevij5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5sales(result[8],"htpsales16",date_export,feuille,callback);//htpsalesj5
      },
      function (callback) {
          Engagementhtp.ecrituredataj5flux(result[9],"htpflux16",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5rejet(result[10],"htprejet16",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5cotlamie(result[11],"htpcotlamiej5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj2cotite(result[12],"htpcotite16",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5cotite(result[13],"htpcotite16",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataj2faclamie(result[14],"htpfaclamiej2",date_export,feuille,callback);
        },
      function (callback) {
        Engagementhtp.ecrituredataj2acs(result[15],"htpacsj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataj5faclamie(result[16],"htpfaclamiej5",date_export,feuille,callback);
      },
    function (callback) {
        Engagementhtp.ecrituredataj5acs(result[17],"htpacsj5",date_export,feuille,callback);
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
          res.view('HTPengagement/exportHTPengagementsuivant_2', {date : datetest});
          // res.view('Retour/succes');
        }

    })
  })
},

/**********************************************************************************************************/ 
//FONCTION EXPORT STOCKS
exporthtpengagementsuivant_2 : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var feuille = annee+mois+'_EASY';
  console.log(feuille);
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
      Engagementhtp.recupdata("htptristock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgstocktot",callback);//htpfacmgstock
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevistocktot",callback);//htpdevistock
    },
    function (callback) {
      Engagementhtp.recupdata("htpsalesstocktot",callback);//htpsalesstock
    },
    function (callback) {
      Engagementhtp.recupdata("htpfluxstock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htprejetstock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamiestock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotitestock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamiestock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacsstock",callback);
    },
    /**************************/
    function (callback) {
      Engagementhtp.recupdata("htptristocktot",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgstocktot",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevistocktot",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpsalesstocktot",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfluxstocktot",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htprejetstocktot",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamiestocktot",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotitestocktot",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamiestocktot",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacsstocktot",callback);
    },

  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredatastock16tri(result[0],"htptristock",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastock16facM(result[1],"htpfacmgstocktot",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastock16devi(result[2],"htpdevistocktot",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastock16sales(result[3],"htpsalesstocktot",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastock16flux(result[4],"htpfluxstock",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastock16rejet(result[5],"htprejetstock",date_export,feuille,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastock16cotlamie(result[6],"htpcotlamiestock",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastock16cotite(result[7],"htpcotitestock",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastock16faclamie(result[8],"htpfaclamiestock",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastock16acs(result[9],"htpacsstock",date_export,feuille,callback);
      },
      /*********************************************************/
      function (callback) {
        Engagementhtp.ecrituredatastocktottri(result[10],"htptristocktot",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocktotfacM(result[11],"htpfacmgstocktot",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocktotdevi(result[12],"htpdevistocktot",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocktotsales(result[13],"htpsalesstocktot",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocktotflux(result[14],"htpfluxstocktot",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocktotrejet(result[15],"htprejetstocktot",date_export,feuille,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastocktotcotlamie(result[16],"htpcotlamiestocktot",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocktotcotite(result[17],"htpcotitestocktot",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocktotfaclamie(result[18],"htpfaclamiestocktot",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocktotacs(result[19],"htpacsstocktot",date_export,feuille,callback);
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
          res.view('HTPengagement/exportHTPengagementsuivant_3', {date : datetest});
          // res.view('Retour/succes');
        }

    })
  })
},

/**********************************************************************************************************/ 
//FONCTION EXPORT ETP
exporthtpengagementsuivant_3 : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var feuille = annee+mois+'_EASY';
  console.log(feuille);
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
  var dateduree = annee+''+mois+''+jour;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      Engagementhtp.recupdata("htptrietp",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgetp",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevietp",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamieetp",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamieetp",callback);
    },
    
  

  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredataetptri(result[0],"htptrietp",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataetpfacM(result[1],"htpfacmgetp",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataetpdevi(result[2],"htpdevietp",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataetpfaclamie(result[3],"htpfaclamieetp",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataetpcotlamie(result[4],"htpcotlamieetp",date_export,feuille,callback);
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
          res.view('HTPengagement/exportHTPengagementsuivant_4', {date : datetest});
          // res.view('Retour/succes');
        }

    })
  })
},
/**************************************************************************/
/**********************************************************************************************************/ 
//FONCTION EXPORT TACHES NON TRAITEES
exporthtpengagementsuivant_4 : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var feuille = annee+mois+'_EASY';
  console.log(feuille);
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
      Engagementhtp.recupdata("htptritnontj",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgtnontj",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevitnontj",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpsalestnontj",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfluxtnontj",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htprejettnontj",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamietnontj",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotitetnontj",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamietnontj",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacstnontj",callback);
    },
    /**************************/
    function (callback) {
      Engagementhtp.recupdata("htptritnontj1",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgtnontj1",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevitnontj1",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpsalestnontj1",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfluxtnontj1",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htprejettnontj1",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamietnontj1",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotitetnontj1",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamietnontj1",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacstnontj1",callback);
    },

  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredatastocknontjtri(result[0],"htptritnontj",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontjfacM(result[1],"htpfacmgtnontj",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontjdevi(result[2],"htpdevitnontj",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontjsales(result[3],"htpsalestnontj",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontjflux(result[4],"htpfluxtnontj",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontjrejet(result[5],"htprejettnontj",date_export,feuille,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastocknontjcotlamie(result[6],"htpcotlamietnontj",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontjcotite(result[7],"htpcotitetnontj",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontjfaclamie(result[8],"htpfaclamietnontj",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontjacs(result[9],"htpacstnontj",date_export,feuille,callback);
      },
      /*********************************************************/
      function (callback) {
        Engagementhtp.ecrituredatastocknontj1tri(result[10],"htptritnontj1",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj1facM(result[11],"htpfacmgtnontj1",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj1devi(result[12],"htpdevitnontj1",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj1sales(result[13],"htpsalestnontj1",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj1flux(result[14],"htpfluxtnontj1",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj1rejet(result[15],"htprejettnontj1",date_export,feuille,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj1cotlamie(result[16],"htpcotlamietnontj1",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj1cotite(result[17],"htpcotitetnontj1",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj1faclamie(result[18],"htpfaclamietnontj1",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj1acs(result[19],"htpacstnontj1",date_export,feuille,callback);
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
          res.view('HTPengagement/exportHTPengagementsuivant_5', {date : datetest});
          // res.view('Retour/succes');
        }

    })
  })
},
/*****************************************************************************************/
/**********************************************************************************************************/ 
//FONCTION EXPORT TACHES NON TRAITEES
exporthtpengagementsuivant_5 : function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var feuille = annee+mois+'_EASY';
  console.log(feuille);
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
      Engagementhtp.recupdata("htptritnontj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgtnontj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevitnontj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpsalestnontj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfluxtnontj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htprejettnontj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamietnontj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotitetnontj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamietnontj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacstnontj2",callback);
    },
    /**************************/
    function (callback) {
      Engagementhtp.recupdata("htptritnontj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgtnontj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevitnontj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpsalestnontj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfluxtnontj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htprejettnontj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamietnontj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotitetnontj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamietnontj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacstnontj5",callback);
    },

  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredatastocknontj2tri(result[0],"htptritnontj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj2facM(result[1],"htpfacmgtnontj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj2devi(result[2],"htpdevitnontj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj2sales(result[3],"htpsalestnontj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj2flux(result[4],"htpfluxtnontj2",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj2rejet(result[5],"htprejettnontj2",date_export,feuille,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj2cotlamie(result[6],"htpcotlamietnontj2",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj2cotite(result[7],"htpcotitetnontj2",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj2faclamie(result[8],"htpfaclamietnontj2",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj2acs(result[9],"htpacstnontj2",date_export,feuille,callback);
      },
      /*********************************************************/
      function (callback) {
        Engagementhtp.ecrituredatastocknontj5tri(result[10],"htptritnontj5",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj5facM(result[11],"htpfacmgtnontj5",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj5devi(result[12],"htpdevitnontj5",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj5sales(result[13],"htpsalestnontj5",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj5flux(result[14],"htpfluxtnontj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj5rejet(result[15],"htprejettnontj5",date_export,feuille,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj5cotlamie(result[16],"htpcotlamietnontj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj5cotite(result[17],"htpcotitetnontj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj5faclamie(result[18],"htpfaclamietnontj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj5acs(result[19],"htpacstnontj5",date_export,feuille,callback);
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
          res.view('Retour/succes');
        }

    })
  })
},
/*******************************************************************************************************/
/******************************************************************************************************/
/*****************************************************************************************************/
  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtp_1 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtp';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtp limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtp_1(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      // return res.view('HTPengagement/importHTPengagementsuivant_1', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_3', {date : datetest});
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantdevis : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpdevis';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpdevis limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpdevis(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_3', {date : datetest});
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantfacture : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpfacture';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpfacture limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpfacture(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      return res.view('HTPengagement/importHTPengagementsuivant_2', {date : datetest});
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantligne : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpligne';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpligne limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    console.log('ddddddddddddddddddddddd lllllllllllllllll');
                    console.log(dernierl);
                    console.log(numligne);
                    console.log(trameflux);
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpligne(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }
                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_4', {date : datetest})
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantdevisj2 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpdevisj2';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpdevisj2 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpdevisj2(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_5', {date : datetest})
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantdevisj5 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpdevisj5';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpdevisj5 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpdevisj5(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_6', {date : datetest})
                  });
               
              }
          })
        };
      });
  },

/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantfacmgj2 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpfacmgj2';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpfacmgj2 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpfacmgj2(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_7', {date : datetest})
                  });
               
              }
          })
        };
      });
  },

/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantfacmgj5 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpfacmgj5';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpfacmgj5 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpfacmgj5(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_8', {date : datetest})
                  });
               
              }
          })
        };
      });
  },

/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantcotlamiej2 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpcotlamiej2';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpcotlamiej2 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpcotlamiej2(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      return res.view('HTPengagement/importHTPengagementsuivant_9', {date : datetest})
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantcotlamiej5 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpcotlamiej5';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpcotlamiej5 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpcotlamiej5(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_10', {date : datetest})
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantsales : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpsales';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpsales limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpsales(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,dateexport,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_11', {date : datetest})
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantsalesstock : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpsalesstock';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpsalesstock limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpsalesstock(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,dateexport,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_12', {date : datetest})
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantfacmgstock : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpstockfacmg';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpstockfacmg limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log('99999999999999999999999'+nc);
            if (err){
              console.log('ato am if erreur FACMGSTOCK');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpfacmgstock(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,dateexport,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_13', {date : datetest})
                  });
               
              }
          })
        };
      });
  },

/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantdevistock : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpstockdevis';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpstockdevis limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpdevistock(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,dateexport,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_14', {date : datetest})
                  });
               
              }
          })
        };
      });
  },

/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantfacmgtnontj : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpfacmgtnontj';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpfacmgtnontj limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpfacmgtnontj(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_15', {date : datetest})
                  });
               
              }
          })
        };
      });
  },

/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantfacmgtnontj1 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpfacmgtnontj1';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpfacmgtnontj1 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpfacmgtnontj1(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_16', {date : datetest})
                  });
               
              }
          })
        };
      });
  },

/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantfacmgtnontj2 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpfacmgtnontj2';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpfacmgtnontj2 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpfacmgtnontj2(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_17', {date : datetest})
                  });
               
              }
          })
        };
      });
  },

/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantfacmgtnontj5 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpfacmgtnontj5';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpfacmgtnontj5 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpfacmgtnontj5(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_18', {date : datetest})
                  });
               
              }
          })
        };
      });
  },

/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantdevitnontj : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpdevitnontj';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpdevitnontj limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpdevitnontj(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_19', {date : datetest})
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantdevitnontj1 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpdevitnontj1';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpdevitnontj1 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpdevitnontj1(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_20', {date : datetest})
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantdevitnontj2 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpdevitnontj2';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpdevitnontj2 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpdevitnontj2(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      // return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      return res.view('HTPengagement/importHTPengagementsuivant_21', {date : datetest})
                  });
               
              }
          })
        };
      });
  },
/*****************************************************************************************************/

  //IMPORTATION DES DONNEES SUR EXCEL DANS LA BASE
  importengagementhtpsuivantdevitnontj5 : function(req,res)
  {
    var datetest = req.param("date",0);
    var sql1= 'select count(*) as nb from cheminengagementhtpdevitnontj5';
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
          //var sql='select * from cheminretourvrai limit' + " " + x ;
          var sql= 'select * from cheminengagementhtpdevitnontj5 limit'  + " " + x;
          Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
            console.log(nc);
            if (err){
              console.log('ato am if erreur');
              console.log(err);
              return next(err);
            }
            else
            {
              console.log('ato amm else');              
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
            var nbre = [];
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].chemin;
                      trameflux.push(a);
                      nbre.push(i);
                    };
                    for(var i=0;i<nb;i++)
                    {
                      var a = nc[i].numfeuille;
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
                    async.forEachSeries(nbre, function(lot, callback_reporting_suivant) {
                      async.series([
                      //   function(cb){
                      //     Engagementhtp.deleteHtp(table,nb,cb);
                      //  },
                        function(cb){
                          Engagementhtp.importengagementhtpdevitnontj5(trameflux,feuil,cellule,table,cellule2,lot,numligne,dernierl,date,cb);
                        },
                      ],function(erroned, lotValues){
                        if(erroned) return res.badRequest(erroned);
                        return callback_reporting_suivant();
                      });
                    },
                    // async.series([
                    //   function(cb){
                    //     Garantie.deleteHtp(table,nb,cb);
                    //   }, 
                    //   function(cb){
                    //     Garantie.importTrameDemat(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   }, 
                    //   // function(cb){
                    //   //   Garantie.importTrameRcindeterminable(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,cb);
                    //   // }, 
                    // ],
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      // return res.view('HTPengagement/importHTPengagementsuivant_16', {date : datetest})
                  });
               
              }
          })
        };
      });
  },



  
};

