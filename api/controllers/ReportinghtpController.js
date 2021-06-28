/**
 * ReportinghtpController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */
const { table } = require("console");
const { setMaxListeners } = require("process");
module.exports = {
  accueil1 : async function(req,res)
    {
      return res.view('reporting/accueilreporting');
    },
  EssaiExcel2 : function(req,res)
  {
    var sql1= 'select count(*) as nb from cheminhtp2;';
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
        var sql='select * from cheminhtp2 limit' + " " + x ;
    Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
      if (err){
        console.log(err);
        return next(err);
      }
      else
      {
          nc = nc.rows;
          var feuil = [];
          var cellule = [];
          var cellule2 = [];
          var table = [];
          var numligne = [];
          var trameflux = [];
          var datetest = req.param("date",0);
          /*var annee = datetest.substr(0, 4);
          var mois = datetest.substr(5, 2);
          var jour = datetest.substr(8, 2);
          var date = annee+mois+jour;
          var dateexport = jour + '/' + mois + '/' +annee;*/
          var nb = x;
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
          for(var i=0;i<nb;i++)
          {
            var a = nc[i].nomtable;
            table.push(a);
          };
                  console.log(table);
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].chemin;
                    trameflux.push(a);
                  };
                  console.log(trameflux);
                  console.log(table);
                  console.log(nb);
                  async.series([  
                    function(cb){
                        Reportinghtp.importTrameFlux929type2(trameflux,feuil,cellule,table,cellule2,0,numligne,cb);
                      }, 
                  ],
                  function(err, resultat){
                    if (err) { return res.view('reporting/erreur'); }
                    return res.view('reporting/exportExcel', {date : datetest});
                });
      };
      
  });
};
      
});
  },
  EssaiExcel : function(req,res)
  {
    var sql1= 'select count(*) as nb from cheminhtp;';
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
        var sql='select * from cheminhtp limit' + " " + x ;
    Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
      if (err){
        console.log(err);
        return next(err);
      }
      else
      {
          nc = nc.rows;
          var feuil = [];
          var cellule = [];
          var cellule2 = [];
          var table = [];
          var numligne = [];
          var trameflux = [];
          var datetest = req.param("date",0);
          var annee = datetest.substr(2, 2);
          var mois = datetest.substr(5, 2);
          var jour = datetest.substr(8, 2);
          if(mois=='01')
          {
            mois= '1';
          };
          if(mois=='02')
          {
            mois= '2';
          };
          if(mois=='03')
          {
            mois= '3';
          };
          if(mois=='04')
          {
            mois= '4';
          };
          if(mois=='05')
          {
            mois= '5';
          };
          if(mois=='06')
          {
            mois= '6';
          };
          if(mois=='07')
          {
            mois= '7';
          };
          if(mois=='08')
          {
            mois= '8';
          };
          if(mois=='09')
          {
            mois= '9';
          };
          if(jour=='01')
          {
            jour= '1';
          };
          if(jour=='02')
          {
            jour= '2';
          };
          if(jour=='03')
          {
            jour= '3';
          };
          if(jour=='04')
          {
            jour= '4';
          };
          if(jour=='05')
          {
            jour= '5';
          };
          if(jour=='06')
          {
            jour= '6';
          };
          if(jour=='07')
          {
            jour= '7';
          };
          if(jour=='08')
          {
            jour= '8';
          };
          if(jour=='09')
          {
            jour= '9';
          };
          console.log('jou'+jour);
          var date = annee+mois+jour;
          var dateexport = mois + '/' + jour + '/' +annee;
          var nb = x;
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
          for(var i=0;i<nb;i++)
          {
            var a = nc[i].nomtable;
            table.push(a);
          };
                  console.log(table);
                  for(var i=0;i<nb;i++)
                  {
                    var a =nc[i].chemin;
                    trameflux.push(a);
                  };
                  console.log(trameflux);
                  console.log(table);
                  console.log(nb);
                  tabletout = ['trameflux','suivisaisieprodite','suivisaisielmde','suivisaisiemgas','tramelamiestock'];
                  async.series([ 
                    function(cb){
                        Reportinghtp.deleteReportingHtp(tabletout,0,cb);
                      },
                    function(cb){
                      Reportinghtp.deleteReportingHtp(tabletout,1,cb);
                    },
                    function(cb){
                      Reportinghtp.deleteReportingHtp(tabletout,2,cb);
                    },
                    function(cb){
                      Reportinghtp.deleteReportingHtp(tabletout,3,cb);
                    },
                    function(cb){
                      Reportinghtp.deleteReportingHtp(tabletout,4,cb);
                    },
                    function(cb){
                        Reportinghtp.importTrameFlux929type4(trameflux,feuil,cellule,table,cellule2,0,numligne,dateexport,cb);
                      }, 
                    function(cb){
                      Reportinghtp.importTrameFlux929type4(trameflux,feuil,cellule,table,cellule2,1,numligne,dateexport,cb);
                    },
                    function(cb){
                      Reportinghtp.importTrameFlux929type4(trameflux,feuil,cellule,table,cellule2,2,numligne,dateexport,cb);
                    },
                    function(cb){
                      Reportinghtp.importTrameFlux929type4(trameflux,feuil,cellule,table,cellule2,3,numligne,dateexport,cb);
                    },
                  ],
                  function(err, resultat){
                    if (err) { return res.view('reporting/erreur'); }
                    // return res.view('reporting/exportExcelHTP');
                    //return res.redirect('/export/'+dateexport +'/'+'<h1><h1>');
                    return res.view('reporting/accueil2', {date : datetest});
                })
              
      };
      
  });
};
      
});
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
    console.log(date);
    var cheminp = [];
    var MotCle= [];
    var nomtable = [];
    var numligne = [];
    var numfeuille = [];
    var nomcolonne = [];
    var colonnecible2 = [];
    var essai = 'essai';
    //workbook.xlsx.readFile('htp.xlsx')
    workbook.xlsx.readFile('ex.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil1');
          var numFeuille = newworksheet.getColumn(4);
          var nomColonne = newworksheet.getColumn(5);
          var nomTable = newworksheet.getColumn(6);
          var cible2 = newworksheet.getColumn(7);
          var numLigne = newworksheet.getColumn(8);
          var cheminparticulier = newworksheet.getColumn(9);
          var motcle = newworksheet.getColumn(10);
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
          cible2.eachCell(function(cell, rowNumber) {
            colonnecible2.push(cell.value);
          });
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
                    Reportinghtp.deleteFromChemin(table,cb);
                  },
                function(cb){
                    Reportinghtp.deleteFromChemin2(table,cb);
                  },
               function(cb){
                    Reportinghtp.importEssai(table,cheminp,date,MotCle,0,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],colonnecible2[0],cb);
                  },
                  function(cb){
                    Reportinghtp.importEssai(table,cheminp,date,MotCle,1,nomtable[1],numligne[1],numfeuille[1],nomcolonne[1],colonnecible2[1],cb);
                  },
                  function(cb){
                    Reportinghtp.importEssai(table,cheminp,date,MotCle,2,nomtable[2],numligne[2],numfeuille[2],nomcolonne[2],colonnecible2[2],cb);
                  },
                  function(cb){
                    Reportinghtp.importEssai(table,cheminp,date,MotCle,3,nomtable[3],numligne[3],numfeuille[3],nomcolonne[3],colonnecible2[3],cb);
                  },
                  function(cb){
                    Reportinghtp.importEssaitype2(table,cheminp,date,MotCle,4,nomtable[4],numligne[4],numfeuille[4],nomcolonne[4],colonnecible2[4],cb);
                  },
                  function(cb){
                    Reportinghtp.existenceRoute(essai,cb);
                    },
                  function(cb){
                    Reportinghtp.existenceRoute2(essai,cb);
                    },
            ],
            function(err, resultat){
              let val = resultat[8].rows;
              let val2 = resultat[7].rows;

              var f = parseInt(val[0].ok) + parseInt(val2[0].ok);
              console.log(val[0].ok);
              if (err) { return res.view('reporting/erreur'); }
              if(f==0)
              {
                return res.view('reporting/erreur');
              }
              else
              {
                return res.view('reporting/accueil', {date : datetest});
              }
          });


        });
  },
 
CompterExcel : function (req, res) {
      var calendrier = req.param('calendrier');
      const Excels = require('exceljs');
      // var sql = "select count(col_4) as ok from excel where col_4='ok'";
      var sql_flux = "select nbok as ok from trameflux";
      var nb_flux;

      sails.sendNativeQuery(sql_flux, function(err, res){
          if (err) { return res.badRequest(err); }
          else {

          // var pdo = "select count(col_4) as ok from excel where col_4='ko'"; 
          var sql_flux_1 = "select nbko as ko from trameflux";   
          var nb_flux_1;
              sails.sendNativeQuery(sql_flux_1, function(err, res){
                  if(err){
                      res.send(err);
                  }else{
                      
                      return nb_flux_1 = res.rows[0].ok;
                      
                  }
              });

              //REQUETE TABLE SUIVISAISIEPRODITE
              var sql_ITE = "select nbok as ok from suivisaisieprodite ";            
              var nb_ITE;
                  sails.sendNativeQuery(sql_ITE, function(err, res){
                      if(err){
                          res.send(err);
                      }else{
                          return nb_ITE = res.rows[0].ok;
                      }
                  });
                  var sql_ITE_1 = "select count(okko) as ok from suivisaisieprodite where okko='KO'";            
                  var nb_ITE_1;
                      sails.sendNativeQuery(sql_ITE_1, function(err, res){
                          if(err){
                              res.send(err);
                          }else{
                              return nb_ITE_1 = res.rows[0].ok;
                          }
                  });

                  //REQUETE TABLE SUIVISAISIEMGAS
              var sql_MGAS = "select nbok as ok from suivisaisiemgas ";            
              var nb_MGAS;
                  sails.sendNativeQuery(sql_MGAS, function(err, res){
                      if(err){
                          res.send(err);
                      }else{
                          return nb_MGAS = res.rows[0].ok;
                      }
                  });
                  var sql_MGAS_1 = "select count(okko) as ok from suivisaisiemgas where okko='KO'";            
                  var nb_MGAS_1;
                      sails.sendNativeQuery(sql_MGAS_1, function(err, res){
                          if(err){
                              res.send(err);
                          }else{
                              return nb_MGAS_1 = res.rows[0].ok;
                          }
                  });

                  //REQUETE TABLE SUIVISAISIELMDE
              var sql_LMDE = "select nbok as ok from suivisaisielmde ";            
              var nb_LMDE;
                  sails.sendNativeQuery(sql_LMDE, function(err, res){
                      if(err){
                          res.send(err);
                      }else{
                          return nb_LMDE = res.rows[0].ok;
                      }
                  });
                  var sql_LMDE_1 = "select count(okko) as ok from suivisaisielmde where okko='KO'";            
                  var nb_LMDE_1;
                      sails.sendNativeQuery(sql_LMDE_1, function(err, res){
                          if(err){
                              res.send(err);
                          }else{
                              return nb_LMDE_1 = res.rows[0].ok;
                          }
                  });

                   //REQUETE TABLE SUIVISAISIELMDE tramelamiestock
              var sql = "select count(okko) as ok from tramelamiestock where okko='OK'";            
              var nb;
                  sails.sendNativeQuery(sql, function(err, res){
                      if(err){
                          res.send(err);
                      }else{
                          return nb = res.rows[0].ok;
                      }
                  });
                  var sql_1 = "select count(okko) as ok from tramelamiestock where okko='KO'";            
                  var nb_1;
                      sails.sendNativeQuery(sql_1, function(err, res){
                          if(err){
                              res.send(err);
                          }else{
                              return nb_1 = res.rows[0].ok;
                          }
                  });

            async function exTest()
            {
              
              nb_flux= res.rows[0].ok;
              
              const newWorkbook = new Excels.Workbook();
              
              await newWorkbook.xlsx.readFile('D:/Reporting/Reporting/REPORTING HTP  Type.xlsx');
              var newworksheet;
              // console.log(newworksheet);
              var donnee = new Date(calendrier);
              var feuille = donnee.getMonth()+1;
              console.log(feuille);
              
              switch(feuille){
                  case 1:
                      newworksheet = newWorkbook.getWorksheet('Janvier');
                      break;
                  case 2:
                      newworksheet = newWorkbook.getWorksheet('Fevrier');
                      break;
                  case 3:
                      newworksheet = newWorkbook.getWorksheet('Mars');
                      break;
                  case 4:
                      newworksheet = newWorkbook.getWorksheet('Avril');
                      break;
                  case 5:
                      newworksheet = newWorkbook.getWorksheet('Mai');
                      break;
                  case 6:
                      newworksheet = newWorkbook.getWorksheet('Juin');
                      break;
                  case 7:
                      newworksheet = newWorkbook.getWorksheet('Juillet');
                      break;
                  case 8:
                      newworksheet = newWorkbook.getWorksheet('Aout');
                      break;
                  case 9:
                      newworksheet = newWorkbook.getWorksheet('Septembre');
                      break;
                  case 10:
                      newworksheet = newWorkbook.getWorksheet('Octobre');
                      break;
                  case 11:
                      newworksheet = newWorkbook.getWorksheet('Novembre');
                      break;
                  case 12:
                      newworksheet = newWorkbook.getWorksheet('Decembre');
                      break;
              }
                            
              //TRAITEMENT LIGNE
             var row = newworksheet.getRow(1);
              // console.log(row);
              var saisis = [];
              var retours = [];
              var retourLigne = [];
              var a,a1,a2,a3,a4,a5,a6;
              var b,b1,b2,b3,b4,b5,b6;
              var c,c1,c2,c3,c4,c5,c6;
              row.eachCell(function(cell, colNumber) {
                if(cell.value == 'DOCUMENTS SAISIS')
                {
                  a = parseInt(colNumber);
                  // console.log(a);
                  saisis.push(a);                     
                  // console.log(saisis[3]);
                  
                  var row1 = newworksheet.getRow(3);
                          row1.eachCell(function(cell, colNumber){
                          if(cell.value == 'report 929' && colNumber == saisis[0]){
                              a1 = parseInt(colNumber);
                              // console.log(a1);
                              }
                          });
                          row1.eachCell(function(cell, colNumber){
                          if(cell.value == 'ITE' && colNumber == saisis[13]){
                              a2 = parseInt(colNumber);
                              
                              }
                          });
                          row1.eachCell(function(cell, colNumber){
                          if(cell.value == 'MGAS' && colNumber == saisis[14]){
                              a3 = parseInt(colNumber);
                          
                              }
                          });
                          row1.eachCell(function(cell, colNumber){
                          if(cell.value == 'LAMIE' && colNumber == saisis[15]){
                              a4 = parseInt(colNumber);
                          
                              }
                          });
                          row1.eachCell(function(cell, colNumber){
                          if(cell.value == 'resiliations' && colNumber == saisis[18]){
                              a5 = parseInt(colNumber);
                             
                          }
                          });
                          row1.eachCell(function(cell, colNumber){
                          if(cell.value == 'LMDE' && colNumber == saisis[19]){
                              a6 = parseInt(colNumber);
                              
                              }
                          });
                       

                  }
              });

              row.eachCell(function(cell, colNumber) {
                  if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
                  {
                  c = parseInt(colNumber);
                  retours.push(c);                     
                  // console.log(retours);
                  
                  var row2 = newworksheet.getRow(3);                                                        
                          row2.eachCell(function(cell, colNumber){
                          if(cell.value == 'report 929' && colNumber == retours[0]){
                              c1 = parseInt(colNumber);
                              
                              }
                          });
                          row2.eachCell(function(cell, colNumber){
                          if(cell.value == 'ITE' && colNumber == retours[13]){
                              c2 = parseInt(colNumber);
                              
                              }
                          });
                          row2.eachCell(function(cell, colNumber){
                          if(cell.value == 'MGAS' && colNumber == retours[14]){
                              c3 = parseInt(colNumber);
                              
                              }
                          });
                          row2.eachCell(function(cell, colNumber){
                          if(cell.value == 'LAMIE' && colNumber == retours[15]){
                              c4 = parseInt(colNumber);
                              
                              }
                          });
                          row2.eachCell(function(cell, colNumber){
                          if(cell.value == 'resiliations' && colNumber == retours[18]){
                              c5 = parseInt(colNumber);
                              
                              }
                          });
                          row2.eachCell(function(cell, colNumber){
                          if(cell.value == 'LMDE' && colNumber == retours[19]){
                              c6 = parseInt(colNumber);
                              
                              }
                          });                                           


                  }
              });

              //TRAITEMENT COLONNE
              var col = newworksheet.getColumn(1);                 
              col.eachCell(function(cell, rowNumber) {     
                  //console.log(cell.text);  
                   //var date = new Date('2021-04-04');
                  var date = new Date(calendrier);
                  if(cell.text == date){ 
                  b = parseInt(rowNumber);
                  // console.log(b);
                  retourLigne.push(b);
                  // console.log(retourLigne);
                  // break;
                  var col2 = newworksheet.getColumn(3);
                  col2.eachCell(function(cell, rowNumber){
                      if(cell.value == "Pack normal" && rowNumber == retourLigne[2]){
                          b4 = parseInt(rowNumber);
                          // console.log(b4);
                          
                          }
                      });
                  }
              });

              
              var rowVrai = newworksheet.getRow(b4);
              // console.log(rowVrai);
              rowVrai.getCell(a1).value = nb_flux;
              rowVrai.getCell(a2).value = nb_ITE;
              rowVrai.getCell(a3).value = nb_MGAS;
              rowVrai.getCell(a4).value = nb;
              rowVrai.getCell(a5).value = nb;
              rowVrai.getCell(a6).value = nb_LMDE;
              rowVrai.getCell(c1).value = nb_flux_1;
              rowVrai.getCell(c2).value = nb_ITE_1;
              rowVrai.getCell(c3).value = nb_MGAS_1;
              rowVrai.getCell(c4).value = nb_1;
              rowVrai.getCell(c5).value = nb_1;
              rowVrai.getCell(c6).value = nb_LMDE_1;
              
              // console.log(nb);
              // console.log(nb1);
              await newWorkbook.xlsx.writeFile('D:/Reporting/Reporting/REPORTING HTP  Type.xlsx')
              // sails.log(a + "b="+ b);  
            }

            exTest();  
          }
      });

      
    },
   
};

