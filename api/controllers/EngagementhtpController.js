/**
 * Odilon 01421
 * EngagementhtpController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

module.exports = {
  
  /*********************************/
  accueilengagementhtp : async function(req,res)
  {
    return res.view('HTPengagement/accueilreportingengagementhtp');
  },
/***********************************************************************************/  
/*
*
*
*                              INSERTION CHEMIN HTP
* 
* 
*  
/***********************************************************************************/
  //FONCTION POUR L'IMPORT DU CHEMIN UTILISER 
insertcheminengagementhtp_1 : function(req,res)
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
                          return res.view('HTPengagement/accueilreportingengagementhtpsuivant3', {date : datetest});
                      
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
       var nomBase = "cheminengagementhtpsales";
       var r = [0,1,2];
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
                             return res.view('HTPengagement/importHTPengagement_1', {date : datetest});
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
*                              IMPORT HTP
* 
* 
*  
/************************************************************************************/

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
                      return res.view('HTPengagement/importHTPengagementsuivant_3', {date : datetest})
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

                      return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                      // return res.view('HTPengagement/importHTPengagementsuivant_4', {date : datetest})
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
                    function(err, resultat){
                      if (err) { return res.view('Inovcom/erreur'); }

                      return res.view('HTPengagement/exportHTPengagement', {date : datetest});
                  });
               
              }
          })
        };
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
          Engagementhtp.recupdatasum("htpfaclamie16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpacs16",callback);
        },
        function (callback) {
          Engagementhtp.recupdatasum("htpfaclamiefin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpacsfin",callback);
        },
       /****************************************************************************************/
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
            Engagementhtp.recupdatasum("htpcotlamiej2",callback);
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
          /****************************************************************/
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
          /*************************************************************************/
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



      ],function(err,result){
        if(err) return res.badRequest(err);
        console.log("Count OK 1==> " + result[0].ok);
        async.series([
          function (callback) {
            Engagementhtp.ecrituredata16tri(result[0],"htptri16",date_export,mois1,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16facM(result[1],"htpfacmg16",date_export,mois1,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16devi(result[2],"htpdevi16",date_export,mois1,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16sales(result[3],"htpsales16",date_export,mois1,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16flux(result[4],"htpflux16",date_export,mois1,callback);
          },
          function (callback) {
              Engagementhtp.ecrituredata16rejet(result[5],"htprejet16",date_export,mois1,callback);
            },
          function (callback) {
          Engagementhtp.ecrituredata16cotlamie(result[6],"htpcotlamie16",date_export,mois1,callback);
          },
          function (callback) {
              Engagementhtp.ecrituredatafinptri(result[7],"htptrifin",date_export,mois1,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpfacM(result[8],"htpfacmgfin",date_export,mois1,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpdevi(result[9],"htpdevifin",date_export,mois1,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpsales(result[10],"htpsalesfin",date_export,mois1,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpflux(result[11],"htpfluxfin",date_export,mois1,callback);
            },
            function (callback) {
                Engagementhtp.ecrituredatafinprejet(result[12],"htprejetfin",date_export,mois1,callback);
              },
            function (callback) {
            Engagementhtp.ecrituredatafinpcotlamie(result[13],"htpcotlamiefin",date_export,mois1,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16cotite(result[14],"htpcotite16",date_export,mois1,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpcotite(result[15],"htpcotitefin",date_export,mois1,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16faclamie(result[16],"htpfaclamie16",date_export,mois1,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16acs(result[17],"htpacs16",date_export,mois1,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpfaclamie(result[18],"htpfaclamiefin",date_export,mois1,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpacs(result[19],"htpacsfin",date_export,mois1,callback);
            },
        /**********************************************************************************************/
         // function (callback) {
                  //   Engagementhtp.ecrituredataj2tri(result[0],"htptrij2",date_export,mois1,callback);
                  // },
                  function (callback) {
                    Engagementhtp.ecrituredataj2facM(result[20],"htpfacmgj2",date_export,mois1,callback);
                  },
                  function (callback) {
                    Engagementhtp.ecrituredataj2devi(result[21],"htpdevij2",date_export,mois1,callback);
                  },
                  function (callback) {
                    Engagementhtp.ecrituredataj2sales(result[22],"htpsales16",date_export,mois1,callback);//htpsalesj2
                  },
                  function (callback) {
                    Engagementhtp.ecrituredataj2flux(result[23],"htpflux16",date_export,mois1,callback);
                  },
                  function (callback) {
                      Engagementhtp.ecrituredataj2rejet(result[24],"htprejet16",date_export,mois1,callback);
                    },
                  function (callback) {
                      Engagementhtp.ecrituredataj2cotlamie(result[25],"htpcotlamiej2",date_export,mois1,callback);
                  },
                  // function (callback) {
                  //     Engagementhtp.ecrituredataj5tri(result[7],"htptrij5",date_export,mois1,callback);
                  // },
                  function (callback) {
                      Engagementhtp.ecrituredataj5facM(result[26],"htpfacmgj5",date_export,mois1,callback);
                  },
                  function (callback) {
                      Engagementhtp.ecrituredataj5devi(result[27],"htpdevij5",date_export,mois1,callback);
                  },
                  function (callback) {
                      Engagementhtp.ecrituredataj5sales(result[28],"htpsales16",date_export,mois1,callback);//htpsalesj5
                  },
                  function (callback) {
                      Engagementhtp.ecrituredataj5flux(result[29],"htpflux16",date_export,mois1,callback);
                  },
                  function (callback) {
                      Engagementhtp.ecrituredataj5rejet(result[30],"htprejet16",date_export,mois1,callback);
                  },
                  function (callback) {
                      Engagementhtp.ecrituredataj5cotlamie(result[31],"htpcotlamiej5",date_export,mois1,callback);
                  },
                  function (callback) {
                      Engagementhtp.ecrituredataj2cotite(result[32],"htpcotite16",date_export,mois1,callback);
                  },
                  function (callback) {
                      Engagementhtp.ecrituredataj5cotite(result[33],"htpcotite16",date_export,mois1,callback);
                    },
                    function (callback) {
                      Engagementhtp.ecrituredataj2faclamie(result[34],"htpfaclamiej2",date_export,mois1,callback);
                    },
                  function (callback) {
                    Engagementhtp.ecrituredataj2acs(result[35],"htpacsj2",date_export,mois1,callback);
                  },
                  function (callback) {
                    Engagementhtp.ecrituredataj5faclamie(result[36],"htpfaclamiej5",date_export,mois1,callback);
                  },
                function (callback) {
                    Engagementhtp.ecrituredataj5acs(result[37],"htpacsj5",date_export,mois1,callback);
                  },
                /**************************************************************************************************/
                function (callback) {
                  Engagementhtp.ecrituredataetptri(result[38],"htptrietp",date_export,mois1,callback);
                },
                function (callback) {
                  Engagementhtp.ecrituredataetpfacM(result[39],"htpfacmgetp",date_export,mois1,callback);
                },
                function (callback) {
                  Engagementhtp.ecrituredataetpdevi(result[40],"htpdevietp",date_export,mois1,callback);
                },
                function (callback) {
                  Engagementhtp.ecrituredataetpfaclamie(result[41],"htpfaclamieetp",date_export,mois1,callback);
                },
                function (callback) {
                  Engagementhtp.ecrituredataetpcotlamie(result[42],"htpcotlamieetp",date_export,mois1,callback);
                },
                /*********************************************************************************************/

                function (callback) {
                  Engagementhtp.ecrituredatastock16tri(result[43],"htptristock",date_export,mois1,callback);
                },
                function (callback) {
                  Engagementhtp.ecrituredatastock16facM(result[44],"htpfacmgstocktot",date_export,mois1,callback);
                },
                function (callback) {
                  Engagementhtp.ecrituredatastock16devi(result[45],"htpdevistocktot",date_export,mois1,callback);
                },
                function (callback) {
                  Engagementhtp.ecrituredatastock16sales(result[46],"htpsalesstocktot",date_export,mois1,callback);
                },
                function (callback) {
                  Engagementhtp.ecrituredatastock16flux(result[47],"htpfluxstock",date_export,mois1,callback);
                },
                function (callback) {
                    Engagementhtp.ecrituredatastock16rejet(result[48],"htprejetstock",date_export,mois1,callback);
                  },
                function (callback) {
                    Engagementhtp.ecrituredatastock16cotlamie(result[49],"htpcotlamiestock",date_export,mois1,callback);
                },
                function (callback) {
                    Engagementhtp.ecrituredatastock16cotite(result[50],"htpcotitestock",date_export,mois1,callback);
                },
                function (callback) {
                    Engagementhtp.ecrituredatastock16faclamie(result[51],"htpfaclamiestock",date_export,mois1,callback);
                },
                function (callback) {
                    Engagementhtp.ecrituredatastock16acs(result[52],"htpacsstock",date_export,mois1,callback);
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
//FONCTION EXPORT TACHES NON TRAITEES
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
/*************************************************************************************/
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
        /**********************************************************************/
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
        Engagementhtp.ecrituredatastocknontjtri(result[0],"htptritnontj",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontjfacM(result[1],"htpfacmgtnontj",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontjdevi(result[2],"htpdevitnontj",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontjsales(result[3],"htpsalestnontj",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontjflux(result[4],"htpfluxtnontj",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontjrejet(result[5],"htprejettnontj",date_export,mois1,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastocknontjcotlamie(result[6],"htpcotlamietnontj",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontjcotite(result[7],"htpcotitetnontj",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontjfaclamie(result[8],"htpfaclamietnontj",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontjacs(result[9],"htpacstnontj",date_export,mois1,callback);
      },
      /*********************************************************/
      function (callback) {
        Engagementhtp.ecrituredatastocknontj1tri(result[10],"htptritnontj1",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj1facM(result[11],"htpfacmgtnontj1",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj1devi(result[12],"htpdevitnontj1",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj1sales(result[13],"htpsalestnontj1",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj1flux(result[14],"htpfluxtnontj1",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj1rejet(result[15],"htprejettnontj1",date_export,mois1,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj1cotlamie(result[16],"htpcotlamietnontj1",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj1cotite(result[17],"htpcotitetnontj1",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj1faclamie(result[18],"htpfaclamietnontj1",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj1acs(result[19],"htpacstnontj1",date_export,mois1,callback);
      },
      /*******************************************************************************************************/
      function (callback) {
        Engagementhtp.ecrituredatastocknontj2tri(result[20],"htptritnontj2",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj2facM(result[21],"htpfacmgtnontj2",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj2devi(result[22],"htpdevitnontj2",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj2sales(result[23],"htpsalestnontj2",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj2flux(result[24],"htpfluxtnontj2",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj2rejet(result[25],"htprejettnontj2",date_export,mois1,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj2cotlamie(result[26],"htpcotlamietnontj2",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj2cotite(result[27],"htpcotitetnontj2",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj2faclamie(result[28],"htpfaclamietnontj2",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj2acs(result[29],"htpacstnontj2",date_export,mois1,callback);
      },
      /*********************************************************/
      function (callback) {
        Engagementhtp.ecrituredatastocknontj5tri(result[30],"htptritnontj5",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj5facM(result[31],"htpfacmgtnontj5",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj5devi(result[32],"htpdevitnontj5",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj5sales(result[33],"htpsalestnontj5",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocknontj5flux(result[34],"htpfluxtnontj5",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj5rejet(result[35],"htprejettnontj5",date_export,mois1,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj5cotlamie(result[36],"htpcotlamietnontj5",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj5cotite(result[37],"htpcotitetnontj5",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj5faclamie(result[38],"htpfaclamietnontj5",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocknontj5acs(result[39],"htpacstnontj5",date_export,mois1,callback);
      },
      /********************************************************************************************************/

      function (callback) {
        Engagementhtp.ecrituredatastocktottri(result[40],"htptristocktot",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocktotfacM(result[41],"htpfacmgstocktot",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocktotdevi(result[42],"htpdevistocktot",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocktotsales(result[43],"htpsalesstocktot",date_export,mois1,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastocktotflux(result[44],"htpfluxstocktot",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocktotrejet(result[45],"htprejetstocktot",date_export,mois1,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastocktotcotlamie(result[46],"htpcotlamiestocktot",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocktotcotite(result[47],"htpcotitestocktot",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocktotfaclamie(result[48],"htpfaclamiestocktot",date_export,mois1,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastocktotacs(result[49],"htpacsstocktot",date_export,mois1,callback);
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
          // res.view('HTPengagement/exportHTPengagementsuivant_5', {date : datetest});
          res.view('Retour/succes');
        }

    })
  })
},


  
};

