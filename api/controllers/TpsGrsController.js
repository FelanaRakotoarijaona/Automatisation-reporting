/**
 * TpsGrsController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */
module.exports = {
  ecritureetpgrs: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      var r = [0,1,2,3,4,5,6,7];
      var table = [];
      var motcle = [];
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil9');
          var motcle1 = newworksheet.getColumn(2);
          var tablem = newworksheet.getColumn(1);
            motcle1.eachCell(function(cell, rowNumber) {
              motcle.push(cell.value);
            });
            tablem.eachCell(function(cell, rowNumber) {
              table.push(cell.value);
            });
                  async.series([
                    function(cb){
                      TpsGrs.countEtp("sdmnument",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("pechospi",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("factse",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("santeclair",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("pecdentaire",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("factoptique",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("facttiers",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("factaudio",cb);
                    },
                   /* function(cb){
                      TpsGrs.countEtp("pecoptique",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("pecaudio",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("tritp",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("trinument",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("sdpnument",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("factdentaire",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("facthospi",cb);
                    },*/
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('Contentieux/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecritureEtp(tab,date,motcle,lot,cb);
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
                                        return res.view('TpsGrs/ecritureetp2',{date:datetest});
                                      };
                              });
                          };
                  });
                });
            
    },
  accueilecritureetpgrs : function(req,res)
   {
     return res.view('TpsGrs/ecritureetp2');
   },
   ecritureetpgrs2: function(req,res)
   {
     var dateFormat = require("dateformat");
     var datetest = req.param("date",0);
     var j = dateFormat(datetest, "dd");
     var m = dateFormat(datetest, "mm");
     var an = dateFormat(datetest, "yyyy");
     var date = j + '/' + m +'/' + an ;
     var r = [0,1,2,3,4,5,6];
     var table = [];
     var motcle = [];
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     workbook.xlsx.readFile('grs16h.xlsx')
       .then(function() {
         var newworksheet = workbook.getWorksheet('Feuil10');
         var motcle1 = newworksheet.getColumn(2);
         var tablem = newworksheet.getColumn(1);
           motcle1.eachCell(function(cell, rowNumber) {
             motcle.push(cell.value);
           });
           tablem.eachCell(function(cell, rowNumber) {
             table.push(cell.value);
           });
                 async.series([
                   /*function(cb){
                     TpsGrs.countEtp("sdmnument",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("pechospi",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("factse",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("santeclair",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("pecdentaire",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("factoptique",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("facttiers",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("factaudio",cb);
                   },*/
                   function(cb){
                     TpsGrs.countEtp("pecoptique",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("pecaudio",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("tritp",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("trinument",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("sdpnument",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("factdentaire",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("facthospi",cb);
                   },
                 ],function(err,result)
                 {
                         if (err){
                           return res.view('Contentieux/erreur');
                         }
                         else
                         {
                           console.log('ok');
 
                           async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                            var tab = result[lot];
                             async.series([
                               function(cb){
                                 TpsGrs.ecritureEtp(tab,date,motcle,lot,cb);
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
                                       return res.view('Contentieux/succes');
                                     };
                             });
                         };
                 });
               });
           
   },
 accueilecritureetpgrs2 : function(req,res)
  {
    return res.view('TpsGrs/ecritureetp');
  },
  traitementstockj1et2et5suivant : function(req,res)
  {
     var sql1= 'select chemin from chemingrsstock16h;';
     TpsGrs.getDatastore().sendNativeQuery(sql1,function(err, nc1) {
       if (err){
         console.log('erreur');
         console.log(err);
       }
       else
       {
        nc1 = nc1.rows;  
        console.log('nc1'+nc1[0].chemin);
        console.log('nc1'+nc1[1].chemin);
        var chemintpssuiviprod23h =nc1[1].chemin;
        var dateFormat = require("dateformat");
        var datetest = req.param("date",0);
        var jour = dateFormat(datetest, "dddd");
        var j = dateFormat(datetest, "dd");
        var m = dateFormat(datetest, "mm");
        var an = dateFormat(datetest, "yyyy");
        var date = an+m+j;
        console.log(jour + date);
        console.log(typeof(date));
        var Excel = require('exceljs');
        var workbook = new Excel.Workbook();
        var trait = [];
        var astt = [];
        var mcle1 = [];
        var mcle2 = [];
        var mcle3 = [];
        var mcle4 = [];
        var table = [];
        var r = [0,1,2,3,4,5,6];
        workbook.xlsx.readFile('grs16h.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil6');
            var traitement = newworksheet.getColumn(2);
            var ast = newworksheet.getColumn(1);
            var motcle1 = newworksheet.getColumn(3);
            var motcle2 = newworksheet.getColumn(4);
            var motcle3 = newworksheet.getColumn(5);
            var motcle4 = newworksheet.getColumn(6);
            var tab = newworksheet.getColumn(7);
              traitement.eachCell(function(cell, rowNumber) {
                trait.push(cell.value);
              });
              ast.eachCell(function(cell, rowNumber) {
                astt.push(cell.value);
              });
              motcle1.eachCell(function(cell, rowNumber) {
                mcle1.push(cell.value);
              });
              motcle2.eachCell(function(cell, rowNumber) {
                mcle2.push(cell.value);
              });
              motcle3.eachCell(function(cell, rowNumber) {
                mcle3.push(cell.value);
              });
              motcle4.eachCell(function(cell, rowNumber) {
                mcle4.push(cell.value);
              });
              tab.eachCell(function(cell, rowNumber) {
                table.push(cell.value);
              });
              async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                async.series([
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                          return res.view('TpsGrs/ecriture');
                        };
                });
            });
          }
      });
  },
  accueilstockj1et2et5suivant : function(req,res)
  {
    return res.view('TpsGrs/suivantstockj1et2et5');
  },
  traitementstockj1et2et5 : function(req,res)
  {
     var sql1= 'select chemin from chemingrsstock16h;';
     TpsGrs.getDatastore().sendNativeQuery(sql1,function(err, nc1) {
       if (err){
         console.log('erreur');
         console.log(err);
       }
       else
       {
        nc1 = nc1.rows;  
        console.log('nc1'+nc1[0].chemin);
        console.log('nc1'+nc1[1].chemin);
        var chemintpssuiviprod23h =nc1[1].chemin;
        //var chemintpssuiviprod23h ='D:/Copie de STT Stock GRS (0023).xls';
        var dateFormat = require("dateformat");
        var datetest = req.param("date",0);
        var jour = dateFormat(datetest, "dddd");
        var j = dateFormat(datetest, "dd");
        var m = dateFormat(datetest, "mm");
        var an = dateFormat(datetest, "yyyy");
        var date = an+m+j;
        console.log(jour + date);
        console.log(typeof(date));
        var Excel = require('exceljs');
        var workbook = new Excel.Workbook();
        var trait = [];
        var astt = [];
        var mcle1 = [];
        var mcle2 = [];
        var mcle3 = [];
        var mcle4 = [];
        var table = [];
        var r = [0,1,2,3,4,5,6,7];
        workbook.xlsx.readFile('grs16h.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil2');
            var traitement = newworksheet.getColumn(2);
            var ast = newworksheet.getColumn(1);
            var motcle1 = newworksheet.getColumn(3);
            var motcle2 = newworksheet.getColumn(4);
            var motcle3 = newworksheet.getColumn(5);
            var motcle4 = newworksheet.getColumn(6);
            var tab = newworksheet.getColumn(7);
              traitement.eachCell(function(cell, rowNumber) {
                trait.push(cell.value);
              });
              ast.eachCell(function(cell, rowNumber) {
                astt.push(cell.value);
              });
              motcle1.eachCell(function(cell, rowNumber) {
                mcle1.push(cell.value);
              });
              motcle2.eachCell(function(cell, rowNumber) {
                mcle2.push(cell.value);
              });
              motcle3.eachCell(function(cell, rowNumber) {
                mcle3.push(cell.value);
              });
              motcle4.eachCell(function(cell, rowNumber) {
                mcle4.push(cell.value);
              });
              tab.eachCell(function(cell, rowNumber) {
                table.push(cell.value);
              });
              async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                async.series([
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                          return res.view('TpsGrs/suivantstockj1et2et5', {date : datetest});
                        };
                });
            });
          }
      });
  },
  accueilstockj1et2et5 : function(req,res)
  {
    return res.view('TpsGrs/bonj1et2et5');
  },
  accueilstock16h: function(req,res)
    {
      return res.view('TpsGrs/stock');
    },
  traitementgrsstock16h: function(req,res)
    {
      var sql1= 'select chemin from chemingrsstock16h;';
      TpsGrs.getDatastore().sendNativeQuery(sql1,function(err, nc1) {
        if (err){
          console.log('erreur');
          console.log(err);
        }
        else
        {
          nc1 = nc1.rows;  
          console.log('nc1'+nc1[0].chemin);
          console.log('nc1'+nc1[1].chemin);
          var chemintpssuiviprod16h =nc1[0].chemin;
          var chemintpssuiviprod23h =nc1[1].chemin;
          /*var chemintpssuiviprod16h ='D:/Copie de STT Stock GRS (002).xls';
          var chemintpssuiviprod23h ='D:/Copie de STT Stock GRS (0023).xls';*/
       var dateFormat = require("dateformat");
       var datetest = req.param("date",0);
       var jour = dateFormat(datetest, "dddd");
       //var jour = 'hafa';
       var j = dateFormat(datetest, "dd");
       var m = dateFormat(datetest, "mm");
       var an = dateFormat(datetest, "yyyy");
       var date = an+m+j;
       console.log(jour + date);
       console.log(typeof(date));
       var Excel = require('exceljs');
       var workbook = new Excel.Workbook();
       var trait = [];
       var astt = [];
       var mcle1 = [];
       var mcle2 = [];
       var mcle3 = [];
       var mcle4 = [];
       var table = [];
       var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14];
       workbook.xlsx.readFile('grs16h.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil2');
           var traitement = newworksheet.getColumn(2);
           var ast = newworksheet.getColumn(1);
           var motcle1 = newworksheet.getColumn(3);
           var motcle2 = newworksheet.getColumn(4);
           var motcle3 = newworksheet.getColumn(5);
           var motcle4 = newworksheet.getColumn(6);
           var tab = newworksheet.getColumn(7);
             traitement.eachCell(function(cell, rowNumber) {
               trait.push(cell.value);
             });
             ast.eachCell(function(cell, rowNumber) {
               astt.push(cell.value);
             });
             motcle1.eachCell(function(cell, rowNumber) {
               mcle1.push(cell.value);
             });
             motcle2.eachCell(function(cell, rowNumber) {
               mcle2.push(cell.value);
             });
             motcle3.eachCell(function(cell, rowNumber) {
               mcle3.push(cell.value);
             });
             motcle4.eachCell(function(cell, rowNumber) {
               mcle4.push(cell.value);
             });
             tab.eachCell(function(cell, rowNumber) {
               table.push(cell.value);
             });
             async.forEachSeries(r, function(lot, callback_reporting_suivant) {
               async.series([
                 function(cb){
                   TpsGrs.traitementInsertionstock16h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
                 },
                 function(cb){
                   TpsGrs.traitementInsertionstockbonJ(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                         return res.view('Tpstpc/bonj1', {date : datetest});
                       };
               });
           });
          }
      });
    },

    accueilrecherchefichier : function(req,res)
    {
      return res.view('TpsGrs/accueilrecherchefichier');
    },
    accueilEtp : function(req,res)
    {
     return res.view('TpsGrs/copieetp');
    },
    copieEtp : function(req,res){
        var dateFormat = require("dateformat");
        var datetest = req.param("date",0);
        var date = dateFormat(datetest, "shortDate");
        console.log('daty'+ date);
        var trameflux= 'D:/Reporting Engagement/GRS.xlsb';
        //var trameflux= '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/GRS_Reporting-Traitement-J-SLA.xlsb';
        var nomColonne = [
            'tritp',
            'trinument',
            'sdpnument',
            'sdmnument',
            'factse',
            'facttiers',
            'factoptique',
            'factaudio',
            'factdentaire',
            'facthospi',
            'santeclair',
            'pecoptique',
            'pecaudio',
            'pecdentaire',
            'pechospi'
        ];
        var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14];
        async.series([  
            function(cb){
                TpsGrs.delete("tpsgrsetp",cb);
              },
        ],
        function(err, resultat){
          if (err) { return res.view('Inovcom/erreur'); }
          else
          {
        async.forEachSeries(r, function(lot, callback_reporting_suivant) {
            async.series([
              function(cb){
                TpsGrs.copieEtp(date,lot,trameflux,nomColonne,cb);
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
                 return res.view('TpsGrs/accueil');
              };
            });
        }
        });
    },
    accueil : function(req,res)
    {
      return res.view('TpsGrs/accueil');
    },
    traitementTacheTraite : function(req,res)
    {
       var sql1= 'select chemin from chemingrssuiviprod16h;';
       ReportingInovcom.getDatastore().sendNativeQuery(sql1,function(err, nc1) {
         if (err){
           console.log('erreur');
           console.log(err);
         }
         else
         {
           nc1 = nc1.rows;  
           console.log('nc1'+nc1[0].chemin);
           console.log('nc1'+nc1[1].chemin);
           var chemintpssuiviprod16h =nc1[0].chemin;
           var chemintpssuiviprod23h =nc1[1].chemin;
        
        var dateFormat = require("dateformat");
        var datetest = req.param("date",0);
        var jour = dateFormat(datetest, "dddd");
        var j = dateFormat(datetest, "dd");
        var m = dateFormat(datetest, "mm");
        var an = dateFormat(datetest, "yyyy");
        var date = an+m+j;
        console.log(jour + date);
        console.log(typeof(date));
        var Excel = require('exceljs');
        var workbook = new Excel.Workbook();
        var trait = [];
        var astt = [];
        var mcle1 = [];
        var mcle2 = [];
        var mcle3 = [];
        var mcle4 = [];
        var table = [];
        var r = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14];
        //var r = [0];
        workbook.xlsx.readFile('grs16h.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil1');
            var traitement = newworksheet.getColumn(2);
            var ast = newworksheet.getColumn(1);
            var motcle1 = newworksheet.getColumn(3);
            var motcle2 = newworksheet.getColumn(4);
            var motcle3 = newworksheet.getColumn(5);
            var motcle4 = newworksheet.getColumn(6);
            var tab = newworksheet.getColumn(7);
              traitement.eachCell(function(cell, rowNumber) {
                trait.push(cell.value);
              });
              ast.eachCell(function(cell, rowNumber) {
                astt.push(cell.value);
              });
              motcle1.eachCell(function(cell, rowNumber) {
                mcle1.push(cell.value);
              });
              motcle2.eachCell(function(cell, rowNumber) {
                mcle2.push(cell.value);
              });
              motcle3.eachCell(function(cell, rowNumber) {
                mcle3.push(cell.value);
              });
              motcle4.eachCell(function(cell, rowNumber) {
                mcle4.push(cell.value);
              });
              tab.eachCell(function(cell, rowNumber) {
                table.push(cell.value);
              });
              async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                async.series([
                  function(cb){
                    ReportingInovcom.delete(table,lot,cb);
                  },
                  function(cb){
                    TpsGrs.traitementInsertion(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
                  },
                 function(cb){
                    Tpstpc.traitementInsertion23h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    Tpstpc.traitementInsertionJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    Tpstpc.traitementInsertionJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                          return res.view('TpsGrs/stock',{date : datetest});
                        };
                });
            });
           }
       });
    },
     
      accueil1 : function(req,res)
      {
        return res.view('Tpstpc/accueil');
      },

    recherchefichier: function(req,res)
    {
          var Excel = require('exceljs');
          var workbook = new Excel.Workbook();
          var dateFormat = require("dateformat");
          var datetest = req.param("date",0);
          var j = dateFormat(datetest, "dd");
          var m = dateFormat(datetest, "mm");
          var an = dateFormat(datetest, "yyyy");
          var date = j  + m + an ;
          var chemin = '//10.128.1.2/bpo_almerys/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/';
          //var chemin= '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/';
          var cheminTotal = chemin + date + '/' ;
          var r = [0,1,2,3];
          var cheminpart = [];
          var motcle = [];
          var table = [];
          var table2= []; // pour la suppression
          var cheminfinal = [];
          workbook.xlsx.readFile('cheminTpsGrs.xlsx')
          .then(function() {
            var newworksheet = workbook.getWorksheet('Feuil1');
            var cheminparticulier = newworksheet.getColumn(1);
            var motcle1 = newworksheet.getColumn(2);
            var nomTable = newworksheet.getColumn(3);
            var nomTable2 = newworksheet.getColumn(4);
            cheminparticulier.eachCell(function(cell, rowNumber) {
                 cheminpart.push(cell.value);
              });
            motcle1.eachCell(function(cell, rowNumber) {
                 motcle.push(cell.value);
              });
            nomTable.eachCell(function(cell, rowNumber) {
                 table.push(cell.value);
              });
            nomTable2.eachCell(function(cell, rowNumber) {
               table2.push(cell.value);
            });
             for (var i=0;i<table.length;i++)
             {
              var chem = cheminTotal  + cheminpart[i];
              cheminfinal.push(chem);
             };
             console.log(cheminfinal);
                  async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                    async.series([
                      function(cb){
                        ReportingInovcom.delete(table2,lot,cb);
                      },
                      function(cb){
                        Tpstpc.importfichier(cheminfinal,motcle,table,lot,cb);
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
                         return res.view('TpsGrs/copieetp');
                      };
                    });
          });
          
    },
    accueilecriture : function(req,res)
    {
      return res.view('TpsGrs/ecriture');
    },
    ecrituregrs1: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      var r = [0,1,2,3,4];
      var table = [];
      var motcle = [];
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil4');
          var motcle1 = newworksheet.getColumn(8);
          var tablem = newworksheet.getColumn(7);
            motcle1.eachCell(function(cell, rowNumber) {
              motcle.push(cell.value);
            });
            tablem.eachCell(function(cell, rowNumber) {
              table.push(cell.value);
            });
                  async.series([
                    function(cb){
                      Tpstpc.countOkKo(table,0,cb);
                    },
                   function(cb){
                      Tpstpc.countOkKo(table,1,cb);
                    },
                    function(cb){
                      Tpstpc.countOkKo(table,2,cb);
                    },
                    function(cb){
                      Tpstpc.countOkKo(table,3,cb);
                    },
                    function(cb){
                      Tpstpc.countOkKo(table,4,cb);
                    },
  
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('Contentieux/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecriture(tab,date,motcle,lot,cb);
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
                                        return res.view('TpsGrs/ecrituresuivant',{date : datetest});
                                      };
                              });
                          };
                  });
                });
            
    },
    accueilecriture3 : function(req,res)
    {
      return res.view('TpsGrs/ecrituresuivant2');
    },
    ecrituregrs3: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      var r = [0,1,2,3,4];
      var table = [];
      var motcle = [];
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil8');
          var motcle1 = newworksheet.getColumn(2);
          var tablem = newworksheet.getColumn(1);
            motcle1.eachCell(function(cell, rowNumber) {
              motcle.push(cell.value);
            });
            tablem.eachCell(function(cell, rowNumber) {
              table.push(cell.value);
            });
                  async.series([
                    function(cb){
                      Tpstpc.countOkKo(table,0,cb);
                    },
                   function(cb){
                      Tpstpc.countOkKo(table,1,cb);
                    },
                    function(cb){
                      Tpstpc.countOkKo(table,2,cb);
                    },
                    function(cb){
                      Tpstpc.countOkKo(table,3,cb);
                    },
                    function(cb){
                      Tpstpc.countOkKo(table,4,cb);
                    },
  
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('Contentieux/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecriture(tab,date,motcle,lot,cb);
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
                                        return res.view('TpsGrs/ecritureetp',{date : datetest});
                                      };
                              });
                          };
                  });
                });
            
    },
    accueilecriture2 : function(req,res)
    {
      return res.view('TpsGrs/ecrituresuivant');
    },
    ecrituregrs2: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      var r = [0,1,2,3,4];
      var table = [];
      var motcle = [];
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil7');
          var motcle1 = newworksheet.getColumn(2);
          var tablem = newworksheet.getColumn(1);
            motcle1.eachCell(function(cell, rowNumber) {
              motcle.push(cell.value);
            });
            tablem.eachCell(function(cell, rowNumber) {
              table.push(cell.value);
            });
                  async.series([
                    function(cb){
                      Tpstpc.countOkKo(table,0,cb);
                    },
                   function(cb){
                      Tpstpc.countOkKo(table,1,cb);
                    },
                    function(cb){
                      Tpstpc.countOkKo(table,2,cb);
                    },
                    function(cb){
                      Tpstpc.countOkKo(table,3,cb);
                    },
                    function(cb){
                      Tpstpc.countOkKo(table,4,cb);
                    },
  
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('Contentieux/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecriture(tab,date,motcle,lot,cb);
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
                                        return res.view('TpsGrs/ecrituresuivant2',{date : datetest});
                                      };
                              });
                          };
                  });
                });
            
    },


    
};

