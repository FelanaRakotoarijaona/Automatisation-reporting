/**
 * TpstpcController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */
 module.exports = {

  // Recherche fichier
  accueilrecherchefichier : function(req,res)
  {
    return res.view('Tpstpc/accueilrecherchefichier');
  },
  recherchefichiertpstpc: function(req,res)
   {
         var Excel = require('exceljs');
         var workbook = new Excel.Workbook();
         var dateFormat = require("dateformat");
         var datetest = req.param("date",0);
         var j = dateFormat(datetest, "dd");
         var m = dateFormat(datetest, "mm");
         var an = dateFormat(datetest, "yyyy");
         var date = j  + m + an ;
         //var chemin = '//10.128.1.2/bpo_almerys/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/';
         var chemin= '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/';
         var cheminTotal = chemin + date + '/' ;
         var r = [0,1,2,3,4,5,6];
         var cheminpart = [];
         var motcle = [];
         var table = [];
         var table2= []; // pour la suppression
         var cheminfinal = [];
         workbook.xlsx.readFile('cheminTpstpc.xlsx')
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
                    var sql4= "select count(chemin) as ok from chemintpsstock16h ";
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
                            return res.view('Tpstpc/accueil');
                            
                          };
                      });
                   });
         });
         
   },

   // Tache traité 16h
   accueil1 : function(req,res)
     {
       return res.view('Tpstpc/accueil');
     },
   traitementTacheTraite16h : function(req,res)
     {
        var sql1= 'select chemin from chemintpssuiviprod16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
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
                     Tpstpc.traitementInsertion(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
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
                           return res.view('Tpstpc/accueilTacheTraite16h2',{date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementTacheTraite16h2 : function(req,res)
     {
        var sql1= 'select chemin from chemintpssuiviprod16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil7');
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
                     Tpstpc.traitementInsertion(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
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
                           return res.view('Tpstpc/accueilTacheTraite16h3',{date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementTacheTraite16h3 : function(req,res)
     {
        var sql1= 'select chemin from chemintpssuiviprod16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil8');
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
                     Tpstpc.traitementInsertion(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
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
                           return res.view('Tpstpc/accueilTacheTraite16h4',{date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementTacheTraite16h4 : function(req,res)
     {
        var sql1= 'select chemin from chemintpssuiviprod16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil9');
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
                     Tpstpc.traitementInsertion(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
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
                           return res.view('Tpstpc/accueilTacheTraite23h',{date : datetest});
                         };
                 });
             });
            }
        });
     },
  // Tache traités 23 h
  accueilTacheTraite23h : function(req,res)
  {
    return res.view('Tpstpc/accueilTacheTraite23h ');
  },
  traitementTacheTraite23h : function(req,res)
  {
 
     var sql1= 'select chemin from chemintpssuiviprod16h;';
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
      var r = [0,1,2];
      workbook.xlsx.readFile('tps16h.xlsx')
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
                  Tpstpc.traitementInsertion23h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                        return res.view('Tpstpc/accueilTacheTraite23h2',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite23h2 : function(req,res)
  {
 
     var sql1= 'select chemin from chemintpssuiviprod16h;';
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
      var r = [0,1,2];
      workbook.xlsx.readFile('tps16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil7');
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
                  Tpstpc.traitementInsertion23h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                        return res.view('Tpstpc/accueilTacheTraite23h3',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite23h3 : function(req,res)
  {
 
     var sql1= 'select chemin from chemintpssuiviprod16h;';
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
      var r = [0,1,2];
      workbook.xlsx.readFile('tps16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil8');
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
                  Tpstpc.traitementInsertion23h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                        return res.view('Tpstpc/accueilTacheTraite23h4',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite23h4 : function(req,res)
  {
 
     var sql1= 'select chemin from chemintpssuiviprod16h;';
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
      var r = [0,1,2];
      workbook.xlsx.readFile('tps16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil9');
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
                  Tpstpc.traitementInsertion23h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                        return res.view('Tpstpc/accueilTacheTraiteJ2',{date : datetest});
                      };
              });
          });
         }
     });
  },
  // Tache traité J2
  accueilTacheTraiteJ2 : function(req,res)
  {
    return res.view('Tpstpc/accueilTacheTraiteJ2');
  },
  traitementTacheTraiteJ2 : function(req,res)
  {
 
     var sql1= 'select chemin from chemintpssuiviprod16h;';
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
      var r = [0,1,2];
      workbook.xlsx.readFile('tps16h.xlsx')
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
                  Tpstpc.traitementInsertionJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                        return res.view('Tpstpc/accueilTacheTraiteJ22',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraiteJ22 : function(req,res)
  {
 
     var sql1= 'select chemin from chemintpssuiviprod16h;';
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
      var r = [0,1,2];
      workbook.xlsx.readFile('tps16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil7');
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
                  Tpstpc.traitementInsertionJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                        return res.view('Tpstpc/accueilTacheTraiteJ23',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraiteJ23 : function(req,res)
  {
 
     var sql1= 'select chemin from chemintpssuiviprod16h;';
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
      var r = [0,1,2];
      workbook.xlsx.readFile('tps16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil8');
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
                  Tpstpc.traitementInsertionJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                        return res.view('Tpstpc/accueilTacheTraiteJ24',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraiteJ24 : function(req,res)
  {
 
     var sql1= 'select chemin from chemintpssuiviprod16h;';
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
      var r = [0,1,2];
      workbook.xlsx.readFile('tps16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil9');
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
                  Tpstpc.traitementInsertionJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                        return res.view('Tpstpc/accueilTacheTraiteJ5',{date : datetest});
                      };
              });
          });
         }
     });
  },
   // Tache traité J5
   accueilTacheTraiteJ5 : function(req,res)
   {
     return res.view('Tpstpc/accueilTacheTraiteJ5');
   },
   traitementTacheTraiteJ5 : function(req,res)
     {
    
        var sql1= 'select chemin from chemintpssuiviprod16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
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
                           return res.view('Tpstpc/accueilTacheTraiteJ52',{date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementTacheTraiteJ52 : function(req,res)
     {
    
        var sql1= 'select chemin from chemintpssuiviprod16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil7');
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
                           return res.view('Tpstpc/accueilTacheTraiteJ53',{date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementTacheTraiteJ53 : function(req,res)
     {
    
        var sql1= 'select chemin from chemintpssuiviprod16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil8');
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
                           return res.view('Tpstpc/accueilTacheTraiteJ54',{date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementTacheTraiteJ54 : function(req,res)
     {
    
        var sql1= 'select chemin from chemintpssuiviprod16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil9');
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
                           return res.view('Tpstpc/accueilstock16h',{date : datetest});
                         };
                 });
             });
            }
        });
     },
    // Traitement stock 16h
    accueilstock16h : function(req,res)
    {
     return res.view('Tpstpc/accueilstock16h');
    },
    traitementStock16h: function(req,res)
     {
      
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
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
                     Tpstpc.traitementInsertionstock16h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
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
                           return res.view('Tpstpc/accueilstock16h2', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementStock16h2: function(req,res)
     {
      
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil10');
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
                     Tpstpc.traitementInsertionstock16h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
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
                           return res.view('Tpstpc/accueilstock16h3', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementStock16h3: function(req,res)
     {
      
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil11');
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
                     Tpstpc.traitementInsertionstock16h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
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
                           return res.view('Tpstpc/accueilstock16h4', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementStock16h4: function(req,res)
     {
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil12');
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
                     Tpstpc.traitementInsertionstock16h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
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
                           return res.view('Tpstpc/accueilstockbonj', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     // Traitement stock bon J
    accueilstockbonj : function(req,res)
    {
     return res.view('Tpstpc/accueilstockbonj');
    },
    traitementstockbonj: function(req,res)
     {
      
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
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
                   /*function(cb){
                     Tpstpc.traitementInsertionstock16h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
                   },*/
                   function(cb){
                     Tpstpc.traitementInsertionstockbonJ(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/accueilstockbonj2', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementstockbonj2: function(req,res)
     {
      
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil10');
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
                   /*function(cb){
                     Tpstpc.traitementInsertionstock16h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
                   },*/
                   function(cb){
                     Tpstpc.traitementInsertionstockbonJ(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/accueilstockbonj3', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementstockbonj3: function(req,res)
     {
      
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil11');
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
                   /*function(cb){
                     Tpstpc.traitementInsertionstock16h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
                   },*/
                   function(cb){
                     Tpstpc.traitementInsertionstockbonJ(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/accueilstockbonj4', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementstockbonj4: function(req,res)
     {
      
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil12');
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
                   /*function(cb){
                     Tpstpc.traitementInsertionstock16h(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod16h,cb);
                   },*/
                   function(cb){
                     Tpstpc.traitementInsertionstockbonJ(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
     // Traitement stock bon J1
     accueilstockJ1 : function(req,res)
     {
      return res.view('Tpstpc/bonj1');
     },
     traitementBonJ1 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
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
                     Tpstpc.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/bonj12', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementBonJ12 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil10');
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
                     Tpstpc.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/bonj13', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementBonJ13 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil11');
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
                     Tpstpc.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/bonj14', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementBonJ14 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil12');
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
                     Tpstpc.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/bonj2', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     // Traitement stock bon J2
     accueilstockJ2 : function(req,res)
     {
      return res.view('Tpstpc/bonj2');
     },
     traitementBonJ2 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
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
                   /*function(cb){
                     Tpstpc.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },*/
                   function(cb){
                     Tpstpc.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },
                   /*function(cb){
                     Tpstpc.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },*/
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
                           return res.view('Tpstpc/bonj22', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementBonJ22 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil10');
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
                   /*function(cb){
                     Tpstpc.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },*/
                   function(cb){
                     Tpstpc.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },
                   /*function(cb){
                     Tpstpc.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },*/
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
                           return res.view('Tpstpc/bonj23', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementBonJ23 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil11');
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
                   /*function(cb){
                     Tpstpc.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },*/
                   function(cb){
                     Tpstpc.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },
                   /*function(cb){
                     Tpstpc.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },*/
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
                           return res.view('Tpstpc/bonj24', {date : datetest});
                         };
                 });
             });
            }
        });
     },

     traitementBonJ24 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil12');
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
                   /*function(cb){
                     Tpstpc.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },*/
                   function(cb){
                     Tpstpc.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },
                   /*function(cb){
                     Tpstpc.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },*/
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
                           return res.view('Tpstpc/bonj5', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     // Traitement stock bon J5
     accueilstockJ5 : function(req,res)
     {
      return res.view('Tpstpc/bonj5');
     },
     traitementBonJ5 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
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
                  /* function(cb){
                     Tpstpc.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },
                   /*function(cb){
                     Tpstpc.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },*/
                   function(cb){
                     Tpstpc.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/bonj52', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementBonJ52 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil10');
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
                  /* function(cb){
                     Tpstpc.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },
                   /*function(cb){
                     Tpstpc.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                   },*/
                   function(cb){
                     Tpstpc.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/bonj53', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementBonJ53 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1,2];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil11');
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
                     Tpstpc.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/bonj54', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     traitementBonJ54 : function(req,res)
     {
        
        var sql1= 'select chemin from chemintpsstock16h;';
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
         var r = [0,1];
         //var r = [0,1,2];
         workbook.xlsx.readFile('tps16h.xlsx')
           .then(function() {
             var newworksheet = workbook.getWorksheet('Feuil12');
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
                     Tpstpc.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                           return res.view('Tpstpc/santeclair', {date : datetest});
                         };
                 });
             });
            }
        });
     },
     // Traitement stock Santeclair
     accueil4 : function(req,res)
     {
      return res.view('Tpstpc/santeclair');
     },
     traitementSanteclair : function(req,res)
     {
      var sql1= 'select chemin from cheminsanteclairstock16h;';
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
          var chemintpsstockprod16h =nc1[0].chemin;
          var chemintpsstockprod23h =nc1[1].chemin;
  
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
           var r = [0];
           workbook.xlsx.readFile('tps16h.xlsx')
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
                       Tpstpc.traitementInsertionstocksanteclair(lot,jour,date,table,chemintpsstockprod16h,cb);
                     },
                     function(cb){
                       Tpstpc.traitementInsertionstocksanteclairJ(lot,jour,date,table,chemintpsstockprod23h,cb);
                     },
                     function(cb){
                       Tpstpc.traitementInsertionstocksanteclairJ1(lot,jour,date,table,chemintpsstockprod23h,cb);
                     },
                     function(cb){
                       Tpstpc.traitementInsertionstocksanteclairJ2(lot,jour,date,table,chemintpsstockprod23h,cb);
                     },
   
                     function(cb){
                       Tpstpc.traitementInsertionstocksanteclairJ5(lot,jour,date,table,chemintpsstockprod23h,cb);
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
                             return res.view('Tpstpc/erreureasy',{date : datetest});
                           };
                   });
               });
              }
          });
     },
     
    /* */
   ecriture3: function(req,res)
   {
     var dateFormat = require("dateformat");
     var datetest = req.param("date",0);
     var j = dateFormat(datetest, "dd");
     var m = dateFormat(datetest, "mm");
     var an = dateFormat(datetest, "yyyy");
     var date = j + '/' + m +'/' + an ;
     var r = [0,1,2,3];
     var table = [];
     var motcle = [];
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     workbook.xlsx.readFile('tps16h.xlsx')
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
                                 Tpstpc.ecriture(tab,date,motcle,lot,cb);
                               },
                             ],function(erroned, lotValues){
                               if(erroned) return res.badRequest(erroned);
                               return callback_reporting_suivant();
                             });
                           },
                             function(err,result)
                             {
                                     if (err){
                                       console.log('res0' + result[0]+result[1]);
                                       if(result[0]=='KO' || result[1]=='KO')
                                       return res.view('Contentieux/erreur');
                                     }
                                     else
                                     {
                                      console.log('res0' + result[0]);
                                      return res.view('Tpstpc/ecritureetp', {date : datetest});
                                     };
                             });
                         };
                 });
               });
           
   },
   selection: function(req,res)
   {
     var dateFormat = require("dateformat");
     var datetest = req.param("date",0);
     var j = dateFormat(datetest, "dd");
     var m = dateFormat(datetest, "mm");
     var an = dateFormat(datetest, "yyyy");
     var date = j + '/' + m +'/' + an ;
     //var r = [0,1,2,3,4,5,6];
     var r = [0,1,2,3];
     var table = [];
     var motcle = [];
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     workbook.xlsx.readFile('tps16h.xlsx')
       .then(function() {
         var newworksheet = workbook.getWorksheet('Feuil1');
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
                                 Tpstpc.ecriture(tab,date,motcle,lot,cb);
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
                                       return res.view('Tpstpc/ecrituresuivant', {date : datetest});
                                     };
                             });
                         };
                 });
               });
           
   },
   ecritureExcel: function(req,res)
   {
    var dateFormat = require("dateformat");
     var datetest = req.param("date",0);
     var j = dateFormat(datetest, "dd");
     var m = dateFormat(datetest, "mm");
     var an = dateFormat(datetest, "yyyy");
     var date = j + '/' + m +'/' + an ;
     //var r = [0,1,2,3,4,5,6];
     var r = [0,1,2,3];
     var table = [];
     var motcle = [];
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     workbook.xlsx.readFile('tps16h.xlsx')
       .then(function() {
         var newworksheet = workbook.getWorksheet('Feuil3');
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
                                 Tpstpc.ecriture(tab,date,motcle,lot,cb);
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
                                       return res.view('Tpstpc/ecrituresuivant2',{date : datetest});
                                     };
                             });
                         };
                 });
               });
           
   },
   ecritureEtp: function(req,res)
   {
    var dateFormat = require("dateformat");
     var datetest = req.param("date",0);
     var j = dateFormat(datetest, "dd");
     var m = dateFormat(datetest, "mm");
     var an = dateFormat(datetest, "yyyy");
     var date = j + '/' + m +'/' + an ;
     var datepouretp = an + m +j;
     var r = [0,1,2,3];
     var table = [];
     var motcle = [];
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     workbook.xlsx.readFile('tps16h.xlsx')
       .then(function() {
         var newworksheet = workbook.getWorksheet('Feuil1');
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
                     Tpstpc.selectionSanteclair(datepouretp,cb);
                   },
                   function(cb){
                    Tpstpc.selection(36139,936,1222,datepouretp,cb);
                   },
                    function(cb){
                      Tpstpc.selectionFactOpt(datepouretp,cb);
                    },
                    function(cb){
                      Tpstpc.selectionFactTiers(datepouretp,cb);
                    },
                  /*function(cb){
                    Tpstpc.selectionSE(datepouretp,cb);
                  },
                  function(cb){
                    Tpstpc.selection(36138,931,1205,datepouretp,cb);
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
                            var tab = parseFloat(result[lot]) / 7.5;
                            console.log(tab);
                             async.series([
                               function(cb){
                                 Tpstpc.ecritureEtp(tab,date,motcle,lot,cb);
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
                                       return res.view('Tpstpc/ecrituresuivantetp0',{date : datetest});
                                     };
                             });
                         };
                 });
               });
           
   },
   ecritureEtp3: function(req,res)
   {
    var dateFormat = require("dateformat");
     var datetest = req.param("date",0);
     var j = dateFormat(datetest, "dd");
     var m = dateFormat(datetest, "mm");
     var an = dateFormat(datetest, "yyyy");
     var date = j + '/' + m +'/' + an ;
     var datepouretp = an + m +j;
     var r = [0,1,2,3];
     var table = [];
     var motcle = [];
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     workbook.xlsx.readFile('tps16h.xlsx')
       .then(function() {
         var newworksheet = workbook.getWorksheet('Feuil6');
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
                    Tpstpc.selectionSE(datepouretp,cb);
                  },
                  function(cb){
                    Tpstpc.selection(36138,931,1205,datepouretp,cb);
                  },
                  function(cb){
                    Tpstpc.selectionPecOptique(datepouretp,cb);
                  },
                  function(cb){
                    Tpstpc.selectionPecAudio(datepouretp,cb);
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
                            var tab = parseFloat(result[lot]) / 7.5;
                            console.log(tab);
                             async.series([
                               function(cb){
                                 Tpstpc.ecritureEtp(tab,date,motcle,lot,cb);
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
                                       return res.view('Tpstpc/ecrituresuivantetp',{date : datetest});
                                     };
                             });
                         };
                 });
               });
           
   },
   ecritureEtp2: function(req,res)
   {
    var dateFormat = require("dateformat");
     var datetest = req.param("date",0);
     var j = dateFormat(datetest, "dd");
     var m = dateFormat(datetest, "mm");
     var an = dateFormat(datetest, "yyyy");
     var date = j + '/' + m +'/' + an ;
     var datepouretp = an + m +j;
     //var r = [0,1,2,3,4,5,6];
     var r = [0,1,2,3];
     var table = [];
     var motcle = [];
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     workbook.xlsx.readFile('tps16h.xlsx')
       .then(function() {
         var newworksheet = workbook.getWorksheet('Feuil5');
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
                    Tpstpc.selectionFactDentaire(datepouretp,cb);
                  },
                  function(cb){
                    Tpstpc.selectionFactHospi(datepouretp,cb);
                  },
                  function(cb){
                    Tpstpc.selectionNument(datepouretp,cb);
                  },
                  function(cb){
                    Tpstpc.selectionPecHospi(datepouretp,cb);
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
                            var tab = parseFloat(result[lot]) / 7.5;
                            console.log(tab);
                             async.series([
                               function(cb){
                                 Tpstpc.ecritureEtp(tab,date,motcle,lot,cb);
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
                                       return res.view('Tpstpc/ecritureerreur',{date : datetest});
                                     };
                             });
                         };
                 });
               });
           
   },
   accueilEcritureDate : function(req,res)
   {
    return res.view('Tpstpc/ecrituredate');
   },
   ecritureDate: function(req,res)
   {
    var dateFormat = require("dateformat");
     var datetest = req.param("date",0);
     var j = dateFormat(datetest, "dd");
     var m = dateFormat(datetest, "mm");
     var an = dateFormat(datetest, "yyyy");
     var date = j + '/' + m +'/' + an ;
     //var date = '14/06/2021';
     //var r = [0,1,2,3,4,5,6];
     var r = [0];
     var table = [];
     var motcle = [];
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     workbook.xlsx.readFile('tps16h.xlsx')
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
                    Tpstpc.countErreur(table,4,cb);
                  },

                 ],function(err,result)
                 {
                         if (err){
                           return res.view('Contentieux/erreur');
                         }
                         else
                         {
                           async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                            var tab = parseFloat(result[lot]);
                            console.log(tab);
                             async.series([
                               function(cb){
                                 Tpstpc.ecritureDate(tab,date,cb);
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


   ecritureErreur: function(req,res)
   {
    var dateFormat = require("dateformat");
     var datetest = req.param("date",0);
     var j = dateFormat(datetest, "dd");
     var m = dateFormat(datetest, "mm");
     var an = dateFormat(datetest, "yyyy");
     var date = j + '/' + m +'/' + an ;
     var datepouretp = an + m +j;
     //var r = [0,1,2,3,4,5,6];
     var r = [0];
     var table = [];
     var motcle = [];
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     workbook.xlsx.readFile('tps16h.xlsx')
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
                    Tpstpc.countErreur(table,4,cb);
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
                            var tab = parseFloat(result[lot]);
                            console.log(tab);
                             async.series([
                               function(cb){
                                 Tpstpc.ecritureEtp2(tab,date,motcle,lot,cb);
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
                                       return res.view('Tpstpc/ecrituredate',{date : datetest});
                                     };
                             });
                         };
                 });
               });
           
   },
   traitementErreurEasy: function(req,res)
   {
    var sql1= 'select chemin from chemintpssuiviproderreur;';
    TpsGrs.getDatastore().sendNativeQuery(sql1,function(err, nc1) {
      if (err){
        console.log('erreur');
        console.log(err);
      }
      else
      {
        nc1 = nc1.rows;  
        console.log('nc1'+nc1[0].chemin);
        var chemintpserreur =nc1[0].chemin;

         var table = "tpserreur";
         var r = [0];
               async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                 async.series([
                   function(cb){
                     Tpstpc.traitementInsertionErreur(table,chemintpserreur,cb);
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
                           return res.view('Tpstpc/ecriture');
                         };
                 });
                }
            });
           
   },
  
   accueiletp : function(req,res)
   {
     return res.view('Tpstpc/ecritureetp');
   },
     accueil : function(req,res)
     {
       return res.view('Tpstpc/etp');
     },
     
    
     
     accueil5 : function(req,res)
     {
       return res.view('Tpstpc/erreureasy');
     },
     accueil6 : function(req,res)
     {
       return res.view('Tpstpc/ecriture');
     },
     
     
     
     
    
 };
 