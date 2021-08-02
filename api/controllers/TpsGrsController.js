/**
 * TpsGrsController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */
module.exports = {
  //Ecriture Etp
  accueilecritureetpgrs : function(req,res)
  {
    return res.view('TpsGrs/ecritureetp');
  },
  ecritureetpgrs: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      var row = req.param("row",0);
      var r = [0,1,2];
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
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('TpsGrs/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecritureEtp(tab,row,date,motcle,lot,cb);
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
                                        return res.view('TpsGrs/ecritureetp4',{date:datetest,row:row});
                                      };
                              });
                          };
                  });
                });
            
    },
    ecritureetpgrs4: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      var row = req.param("row",0);
      var r = [0,1];
      var table = [];
      var motcle = [];
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil24');
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
                      TpsGrs.countEtp("santeclair",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("pecdentaire",cb);
                    },
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('TpsGrs/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecritureEtp(tab,row,date,motcle,lot,cb);
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
                                        return res.view('TpsGrs/ecritureetp3',{date:datetest,row:row});
                                      };
                              });
                          };
                  });
                });
            
    },
    accueilecritureetpgrs3 : function(req,res)
    {
      return res.view('TpsGrs/ecritureetp3');
    },
    ecritureetpgrs3: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      var row = req.param("row",0);
      var r = [0,1,2];
      var table = [];
      var motcle = [];
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil12');
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
                      TpsGrs.countEtp("factoptique",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("facttiers",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("factaudio",cb);
                    },
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('TpsGrs/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecritureEtp(tab,row,date,motcle,lot,cb);
                                },
                              ],function(erroned, lotValues){
                                if(erroned) return res.badRequest(erroned);
                                return callback_reporting_suivant();
                              });
                            },
                              function(err)
                              {
                                      if (err){
                                        return res.view('TpsGrs/erreur');
                                      }
                                      else
                                      {
                                        return res.view('TpsGrs/ecritureetp5',{date:datetest,row:row});
                                      };
                              });
                          };
                  });
                });
            
    },
    ecritureetpgrs5: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      var row = req.param("row",0);
      var r = [0,1];
      var table = [];
      var motcle = [];
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil25');
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
                      TpsGrs.countEtp("factdentaire",cb);
                    },
                    function(cb){
                      TpsGrs.countEtp("facthospi",cb);
                    },
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('TpsGrs/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecritureEtp(tab,row,date,motcle,lot,cb);
                                },
                              ],function(erroned, lotValues){
                                if(erroned) return res.badRequest(erroned);
                                return callback_reporting_suivant();
                              });
                            },
                              function(err)
                              {
                                      if (err){
                                        return res.view('TpsGrs/erreur');
                                      }
                                      else
                                      {
                                        return res.view('TpsGrs/ecritureetp2',{date:datetest,row:row});
                                      };
                              });
                          };
                  });
                });
            
    },
    accueilecritureetpgrs2 : function(req,res)
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
     var row= req.param("row",0);
     var date = j + '/' + m +'/' + an ;
     var r = [0,1,2];
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
                   function(cb){
                     TpsGrs.countEtp("pecoptique",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("pecaudio",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("tritp",cb);
                   },
                 ],function(err,result)
                 {
                         if (err){
                           return res.view('TpsGrs/erreur');
                         }
                         else
                         {
                           console.log('ok');
 
                           async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                            var tab = result[lot];
                             async.series([
                               function(cb){
                                 TpsGrs.ecritureEtp(tab,row,date,motcle,lot,cb);
                               },
                             ],function(erroned, lotValues){
                               if(erroned) return res.badRequest(erroned);
                               return callback_reporting_suivant();
                             });
                           },
                             function(err)
                             {
                                     if (err){
                                       return res.view('TpsGrs/erreur');
                                     }
                                     else
                                     {
                                       return res.view('TpsGrs/ecritureetp6');
                                     };
                             });
                         };
                 });
               });
           
   },
   ecritureetpgrs6: function(req,res)
   {
     var dateFormat = require("dateformat");
     var datetest = req.param("date",0);
     var j = dateFormat(datetest, "dd");
     var m = dateFormat(datetest, "mm");
     var an = dateFormat(datetest, "yyyy");
     var row= req.param("row",0);
     var date = j + '/' + m +'/' + an ;
     var r = [0,1];
     var table = [];
     var motcle = [];
     var Excel = require('exceljs');
     var workbook = new Excel.Workbook();
     workbook.xlsx.readFile('grs16h.xlsx')
       .then(function() {
         var newworksheet = workbook.getWorksheet('Feuil26');
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
                     TpsGrs.countEtp("trinument",cb);
                   },
                   function(cb){
                     TpsGrs.countEtp("sdpnument",cb);
                   }
                 ],function(err,result)
                 {
                         if (err){
                           return res.view('TpsGrs/erreur');
                         }
                         else
                         {
                           console.log('ok');
 
                           async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                            var tab = result[lot];
                             async.series([
                               function(cb){
                                 TpsGrs.ecritureEtp(tab,row,date,motcle,lot,cb);
                               },
                             ],function(erroned, lotValues){
                               if(erroned) return res.badRequest(erroned);
                               return callback_reporting_suivant();
                             });
                           },
                             function(err)
                             {
                                     if (err){
                                       return res.view('TpsGrs/erreur');
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
  // Recherche fichier 
  accueilrecherchefichier : function(req,res)
  {
    return res.view('TpsGrs/accueilrecherchefichier');
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
        var chemin= '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/';
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
                      var sql4= "select count(chemin) as ok from chemingrsstock16h ";
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
                              return res.view('TpsGrs/copieetp');
                              
                            };
                        });
                    };
                  });
        });
        
  },
  // Copie ETP
  accueilEtp : function(req,res)
  {
   return res.view('TpsGrs/copieetp');
  },
  copieEtp : function(req,res){
    var dateFormat = require("dateformat");
    var datetest = req.param("date",0);
    var date = dateFormat(datetest, "shortDate");
    console.log('daty'+ date);
    var trameflux= '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/GRS_Reporting-Traitement-J-SLA.xlsb';
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
    var r = [0,1,2];
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
             return res.view('TpsGrs/accueilEtp3',{date:datetest});
          };
        });
    }
    });
},
copieEtp3 : function(req,res){
  var dateFormat = require("dateformat");
  var datetest = req.param("date",0);
  var date = dateFormat(datetest, "shortDate");
  console.log('daty'+ date);
  var trameflux= '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/GRS_Reporting-Traitement-J-SLA.xlsb';
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
  var r = [3,4,5];
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
           return res.view('TpsGrs/accueilEtp4',{date:datetest});
        };
      });
},
copieEtp4 : function(req,res){
  var dateFormat = require("dateformat");
  var datetest = req.param("date",0);
  var date = dateFormat(datetest, "shortDate");
  console.log('daty'+ date);
  var trameflux= '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/GRS_Reporting-Traitement-J-SLA.xlsb';
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
  var r = [6,7,8];
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
           return res.view('TpsGrs/accueilEtp5',{date:datetest});
        };
      });
},
copieEtp5 : function(req,res){
  var dateFormat = require("dateformat");
  var datetest = req.param("date",0);
  var date = dateFormat(datetest, "shortDate");
  console.log('daty'+ date);
  var trameflux= '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/GRS_Reporting-Traitement-J-SLA.xlsb';
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
  var r = [9,10,11];
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
           return res.view('TpsGrs/accueilEtp2',{date:datetest});
        };
      });
},
  copieEtp2 : function(req,res){
    var dateFormat = require("dateformat");
    var datetest = req.param("date",0);
    var date = dateFormat(datetest, "shortDate");
    console.log('daty'+ date);
    var trameflux= '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/GRS_Reporting-Traitement-J-SLA.xlsb';
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
    var r = [12,13,14];
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
},
  //tache traités 16h
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
                        return res.view('TpsGrs/accueiltachetraite16h1',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite16h1 : function(req,res)
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
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil13');
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
                        return res.view('TpsGrs/accueiltachetraite16h2',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite16h2 : function(req,res)
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
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil14');
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
                        return res.view('TpsGrs/accueiltachetraite16h3',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite16h3: function(req,res)
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
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil15');
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
                        return res.view('TpsGrs/accueiltachetraite16h4',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite16h4 : function(req,res)
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
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil16');
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
                        return res.view('TpsGrs/accueiltachetraite23h',{date : datetest});
                      };
              });
          });
         }
     });
  },
 
  //tache traités 23h
  accueiltachetraite23h : function(req,res)
  {
    return res.view('TpsGrs/accueiltachetraite23h');
  },
  traitementTacheTraite23h : function(req,res)
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
      var r = [0,1,2];
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
                        return res.view('TpsGrs/accueiltachetraite23h1',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite23h1 : function(req,res)
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
      var r = [0,1,2];
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil13');
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
                        return res.view('TpsGrs/accueiltachetraite23h2',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite23h2 : function(req,res)
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
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil14');
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
                        return res.view('TpsGrs/accueiltachetraite23h3',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite23h3 : function(req,res)
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
      var r = [0,1,2];
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil15');
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
                        return res.view('TpsGrs/accueiltachetraite23h4',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraite23h4 : function(req,res)
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
      var r = [0,1,2];
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil16');
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
                        return res.view('TpsGrs/accueiltachetraiteJ2',{date : datetest});
                      };
              });
          });
         }
     });
  },
  //tache traités J2
  accueiltachetraiteJ2 : function(req,res)
  {
    return res.view('TpsGrs/accueiltachetraiteJ2');
  },
  traitementTacheTraiteJ2 : function(req,res)
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
      var r = [0,1,2];
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
                        return res.view('TpsGrs/accueiltachetraiteJ21',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraiteJ21 : function(req,res)
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
      var r = [0,1,2];
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil13');
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
                        return res.view('TpsGrs/accueiltachetraiteJ22',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraiteJ22 : function(req,res)
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
      var r = [0,1,2];
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil14');
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
                        return res.view('TpsGrs/accueiltachetraiteJ23',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraiteJ23 : function(req,res)
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
      var r = [0,1,2];
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil15');
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
                        return res.view('TpsGrs/accueiltachetraiteJ24',{date : datetest});
                      };
              });
          });
         }
     });
  },
  traitementTacheTraiteJ24 : function(req,res)
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
      var r = [0,1,2];
      //var r = [0];
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil16');
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
                        return res.view('TpsGrs/accueiltachetraiteJ5',{date : datetest});
                      };
              });
          });
         }
     });
  },
   //tache traités J5
   accueiltachetraiteJ5 : function(req,res)
   {
     return res.view('TpsGrs/accueiltachetraiteJ5');
   },
   traitementTacheTraiteJ5 : function(req,res)
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
       var r = [0,1,2];
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
                         return res.view('TpsGrs/accueiltachetraiteJ51',{date : datetest});
                       };
               });
           });
          }
      });
   },
   traitementTacheTraiteJ51 : function(req,res)
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
       var r = [0,1,2];
       //var r = [0];
       workbook.xlsx.readFile('grs16h.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil13');
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
                         return res.view('TpsGrs/accueiltachetraiteJ52',{date : datetest});
                       };
               });
           });
          }
      });
   },
   traitementTacheTraiteJ52 : function(req,res)
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
       var r = [0,1,2];
       //var r = [0];
       workbook.xlsx.readFile('grs16h.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil14');
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
                         return res.view('TpsGrs/accueiltachetraiteJ53',{date : datetest});
                       };
               });
           });
          }
      });
   },
   traitementTacheTraiteJ53 : function(req,res)
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
       var r = [0,1,2];
       //var r = [0];
       workbook.xlsx.readFile('grs16h.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil15');
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
                         return res.view('TpsGrs/accueiltachetraiteJ54',{date : datetest});
                       };
               });
           });
          }
      });
   },
   traitementTacheTraiteJ54 : function(req,res)
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
       var r = [0,1,2];
       //var r = [0];
       workbook.xlsx.readFile('grs16h.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil16');
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
                         return res.view('TpsGrs/stock',{date : datetest});
                       };
               });
           });
          }
      });
   },
   // stock 16h
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
                         return res.view('TpsGrs/stock1', {date : datetest});
                       };
               });
           });
          }
      });
    },
    traitementgrsstock16h1: function(req,res)
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
       workbook.xlsx.readFile('grs16h.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil17');
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
                         return res.view('TpsGrs/stock2', {date : datetest});
                       };
               });
           });
          }
      });
    },
    traitementgrsstock16h2: function(req,res)
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
       workbook.xlsx.readFile('grs16h.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil18');
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
                         return res.view('TpsGrs/stock3', {date : datetest});
                       };
               });
           });
          }
      });
    },
    traitementgrsstock16h3: function(req,res)
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
       workbook.xlsx.readFile('grs16h.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil19');
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
                         return res.view('TpsGrs/stock4', {date : datetest});
                       };
               });
           });
          }
      });
    },
    traitementgrsstock16h4: function(req,res)
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
       workbook.xlsx.readFile('grs16h.xlsx')
         .then(function() {
           var newworksheet = workbook.getWorksheet('Feuil20');
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
                         return res.view('TpsGrs/bonj', {date : datetest});
                       };
               });
           });
          }
      });
    },
// stock bon J
accueilstockbonj: function(req,res)
{
  return res.view('TpsGrs/bonj');
},
traitementgrsstockbonj: function(req,res)
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
                     return res.view('TpsGrs/bonj01', {date : datetest});
                   };
           });
       });
      }
  });
},
traitementgrsstockbonj1: function(req,res)
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
   workbook.xlsx.readFile('grs16h.xlsx')
     .then(function() {
       var newworksheet = workbook.getWorksheet('Feuil17');
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
                     return res.view('TpsGrs/bonj02', {date : datetest});
                   };
           });
       });
      }
  });
},
traitementgrsstockbonj2: function(req,res)
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
   workbook.xlsx.readFile('grs16h.xlsx')
     .then(function() {
       var newworksheet = workbook.getWorksheet('Feuil18');
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
                     return res.view('TpsGrs/bonj03', {date : datetest});
                   };
           });
       });
      }
  });
},
traitementgrsstockbonj3: function(req,res)
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
   workbook.xlsx.readFile('grs16h.xlsx')
     .then(function() {
       var newworksheet = workbook.getWorksheet('Feuil19');
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
                     return res.view('TpsGrs/bonj04', {date : datetest});
                   };
           });
       });
      }
  });
},
traitementgrsstockbonj4: function(req,res)
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
                     return res.view('TpsGrs/bonj1', {date : datetest});
                   };
           });
       });
      }
  });
},
// stock bon J1
  /* */
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
        var r = [0,1,2,3,4];
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
                 /* function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                          return res.view('TpsGrs/traitementstock23h1', {date : datetest});
                        };
                });
            });
          }
      });
  },
  traitementstock23h1 : function(req,res)
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
        var r = [0,1,2,3,4];
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
                 /* function(cb){
                    TpsGrs.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },*/
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                 /* function(cb){
                    TpsGrs.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                          return res.view('TpsGrs/traitementstock23h2', {date : datetest});
                        };
                });
            });
          }
      });
  },
  traitementstock23h2 : function(req,res)
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
        var r = [0,1,2,3,4];
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
                 /* function(cb){
                    TpsGrs.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },*/
                  /*function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },*/
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
        var r = [0,1,2,3,4];
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
                 /* function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                          return res.view('TpsGrs/traitementstock23h3',{date:datetest});
                        };
                });
            });
          }
      });
  },
  traitementstock23h3 : function(req,res)
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
        var r = [0,1,2,3,4];
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
                 /* function(cb){
                    TpsGrs.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },*/
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  /*function(cb){
                    TpsGrs.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                          return res.view('TpsGrs/traitementstock23h4',{date:datetest});
                        };
                });
            });
          }
      });
  },
  traitementstock23h4 : function(req,res)
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
        var r = [0,1,2,3,4];
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
                 /* function(cb){
                    TpsGrs.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },*/
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
                          return res.view('TpsGrs/bonj5',{date:datetest});
                        };
                });
            });
          }
      });
  },
  accueilstockj1et2et5final : function(req,res)
  {
    return res.view('TpsGrs/bonj5');
  },

  traitementstockj1et2et5final : function(req,res)
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
        var r = [0,1,2,3,4];
        workbook.xlsx.readFile('grs16h.xlsx')
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
                    TpsGrs.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  /*function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                          return res.view('TpsGrs/traitementstock23h5',{date:datetest});
                        };
                });
            });
          }
      });
  },
  traitementstock23h5 : function(req,res)
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
        var r = [0,1,2,3,4];
        workbook.xlsx.readFile('grs16h.xlsx')
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
                    TpsGrs.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },*/
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                 /* function(cb){
                    TpsGrs.traitementInsertionstockbonJ5(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
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
                          return res.view('TpsGrs/traitementstock23h6',{date:datetest});
                        };
                });
            });
          }
      });
  },
  traitementstock23h6 : function(req,res)
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
        var r = [0,1,2,3,4];
        workbook.xlsx.readFile('grs16h.xlsx')
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
                    TpsGrs.traitementInsertionstockbonJ1(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },
                  function(cb){
                    TpsGrs.traitementInsertionstockbonJ2(astt,trait,mcle1,mcle2,mcle3,mcle4,lot,jour,date,table,chemintpssuiviprod23h,cb);
                  },*/
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
                          return res.view('TpsGrs/rechercheDate');
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
  accueilstockj1et2et5 : function(req,res)
  {
    return res.view('TpsGrs/bonj1');
  },
      accueil1 : function(req,res)
      {
        return res.view('TpsGrs/accueil');
      },
  accueilrechercheDate : function(req,res)
      {
        return res.view('TpsGrs/rechercheDate');
      },
    rechercheDate : async function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      console.log('date'+date);
      const Excel = require('exceljs');
      const newWorkbook = new Excel.Workbook();
      const path_reporting = '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/TestReporting/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
      //var path_reporting = 'D:/Reporting Engagement/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
      await newWorkbook.xlsx.readFile(path_reporting);
      const newworksheet = newWorkbook.getWorksheet('202106_Easy');
      var colonneDate = newworksheet.getColumn('A');
      var ligneDate1;
      console.log(date);
      var ligne = 0;
      var row = 0;
      colonneDate.eachCell(function(cell, rowNumber) {
        var dateExcel = ReportingInovcomExport.convertDate(cell.text);
        if(dateExcel==date)
        {
          ligne = rowNumber;
          if(ligne>row)
          {
            row = ligne
          };
        };
      });
      console.log('row'+ row);
      return res.view('TpsGrs/ecritureDate',{row:row,date:datetest});
    },
    // ecritureee Date
    ecrituredateGrs : function(req,res)
    {
    var dateFormat = require("dateformat");
    var datetest = req.param("date",0);
    var row = req.param("row",0);
    var j = dateFormat(datetest, "dd");
    var m = dateFormat(datetest, "mm");
    var an = dateFormat(datetest, "yyyy");
    var date = j + '/' + m +'/' + an ;
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
                  TpsGrs.ecritureDate('tab',date,row,cb);
                 },

                ],function(err,result)
                {
                        if (err){
                          return res.view('TpsGrs/erreur');
                        }
                        else
                        {
                          return res.view('TpsGrs/ecrituredate2',{date:datetest,row:row});
                        };
                });
              });

    },
    ecrituredateGrs2 : function(req,res)
    {
    var dateFormat = require("dateformat");
    var datetest = req.param("date",0);
    var row = req.param("row",0);
    var j = dateFormat(datetest, "dd");
    var m = dateFormat(datetest, "mm");
    var an = dateFormat(datetest, "yyyy");
    var date = j + '/' + m +'/' + an ;
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
                  TpsGrs.ecritureDate1('tab',date,row,cb);
                 },

                ],function(err,result)
                {
                        if (err){
                          return res.view('TpsGrs/erreur');
                        }
                        else
                        {
                          return res.view('TpsGrs/ecriture',{date:datetest,row:row});
                        };
                });
              });

    },
    // Ecriture  
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
      var row = req.param("row",0);
      var r = [0,1,2];
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
                   
  
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('TpsGrs/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecriture(tab,row,date,motcle,lot,cb);
                                },
                              ],function(erroned, lotValues){
                                if(erroned) return res.badRequest(erroned);
                                return callback_reporting_suivant();
                              });
                            },
                              function(err)
                              {
                                      if (err){
                                        return res.view('TpsGrs/erreur');
                                      }
                                      else
                                      {
                                        return res.view('TpsGrs/ecrituregrs12',{date : datetest,row:row});
                                      };
                              });
                          };
                  });
                });
            
    },
    ecrituregrs12: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      var row = req.param("row",0);
      //var ligne = req.param("row",0);
      var r = [0,1];
      var table = [];
      var motcle = [];
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil21');
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
  
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('TpsGrs/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecriture(tab,row,date,motcle,lot,cb);
                                },
                              ],function(erroned, lotValues){
                                if(erroned) return res.badRequest(erroned);
                                return callback_reporting_suivant();
                              });
                            },
                              function(err)
                              {
                                      if (err){
                                        return res.view('TpsGrs/erreur');
                                      }
                                      else
                                      {
                                        return res.view('TpsGrs/ecrituresuivant',{date : datetest,row:row});
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
      var row = req.param("row",0);
      var r = [0,1,2];
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
                    
  
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('TpsGrs/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecriture(tab,row,date,motcle,lot,cb);
                                },
                              ],function(erroned, lotValues){
                                if(erroned) return res.badRequest(erroned);
                                return callback_reporting_suivant();
                              });
                            },
                              function(err)
                              {
                                      if (err){
                                        return res.view('TpsGrs/erreur');
                                      }
                                      else
                                      {
                                        return res.view('TpsGrs/ecrituregrs13',{date : datetest,row:row});
                                      };
                              });
                          };
                  });
                });
            
    },
    ecrituregrs13: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var date = j + '/' + m +'/' + an ;
      var row = req.param("row",0);
      var r = [0,1];
      var table = [];
      var motcle = [];
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil22');
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
                 
  
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('TpsGrs/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecriture(tab,row,date,motcle,lot,cb);
                                },
                              ],function(erroned, lotValues){
                                if(erroned) return res.badRequest(erroned);
                                return callback_reporting_suivant();
                              });
                            },
                              function(err)
                              {
                                      if (err){
                                        return res.view('TpsGrs/erreur');
                                      }
                                      else
                                      {
                                        return res.view('TpsGrs/ecrituresuivant2',{date : datetest,row:row});
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
      var row = req.param("row",0);
      var date = j + '/' + m +'/' + an ;
      var r = [0,1,2];
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
                   
  
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('TpsGrs/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecriture3(tab,row,date,motcle,lot,cb);
                                },
                              ],function(erroned, lotValues){
                                if(erroned) return res.badRequest(erroned);
                                return callback_reporting_suivant();
                              });
                            },
                              function(err)
                              {
                                      if (err){
                                        return res.view('TpsGrs/erreur');
                                      }
                                      else
                                      {
                                        return res.view('TpsGrs/ecrituregrs14',{date : datetest,row:row});
                                      };
                              });
                          };
                  });
                });
            
    },
   
    ecrituregrs14: function(req,res)
    {
      var dateFormat = require("dateformat");
      var datetest = req.param("date",0);
      var j = dateFormat(datetest, "dd");
      var m = dateFormat(datetest, "mm");
      var an = dateFormat(datetest, "yyyy");
      var row = req.param("row",0);
      var date = j + '/' + m +'/' + an ;
      var r = [0,1];
      var table = [];
      var motcle = [];
      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('grs16h.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil23');
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
                  ],function(err,result)
                  {
                          if (err){
                            return res.view('TpsGrs/erreur');
                          }
                          else
                          {
                            console.log('ok');
  
                            async.forEachSeries(r, function(lot, callback_reporting_suivant) {
                             var tab = result[lot];
                              async.series([
                                function(cb){
                                  TpsGrs.ecriture3(tab,row,date,motcle,lot,cb);
                                },
                              ],function(erroned, lotValues){
                                if(erroned) return res.badRequest(erroned);
                                return callback_reporting_suivant();
                              });
                            },
                              function(err)
                              {
                                      if (err){
                                        return res.view('TpsGrs/erreur');
                                      }
                                      else
                                      {
                                        return res.view('TpsGrs/ecritureetp',{date : datetest,row:row});
                                      };
                              });
                          };
                  });
                });
            
    },
   
    
};

