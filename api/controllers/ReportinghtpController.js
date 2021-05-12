/**
 * ReportinghtpController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */
const { table } = require("console");
const { setMaxListeners } = require("process");
module.exports = {
  EssaiExcel : function(req,res)
  {
    var sql= 'select * from chemin limit 5;';
    Reportinghtp.getDatastore().sendNativeQuery(sql,function(err, nc) {
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
          var numligne = [];
          var trameflux = [];
          var datetest = req.param("date",0);
          var annee = datetest.substr(0, 4);
          var mois = datetest.substr(5, 2);
          var jour = datetest.substr(8, 2);
          var date = annee+mois+jour;
          var dateexport = jour + '/' + mois + '/' +annee;
          var nb = 5;
          workbook.xlsx.readFile('ex.xlsx')
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
                  console.log(table);
                  console.log(nb);
                  async.series([
                    function(cb){
                          Reportinghtp.deleteTout(table,nb,cb);
                      }, 
                    function(cb){
                        Reportinghtp.deleteHtp(table,nb,cb);
                      }, 
                    function(cb){
                        Reportinghtp.importTrameFlux929(trameflux,feuil,cellule,table,cellule2,nb,cb);
                      }, 
                   /*function(cb){
                        Reportinghtp.importInovcom(trameflux,feuil,cellule,table,cellule2,numligne,nb,cb);
                        },
                  function(cb){
                    Reportinghtp.importFinal(table,0,cb);
                      },
                  function(cb){
                    Reportinghtp.importFinal(table,1,cb);
                      },
                  function(cb){
                    Reportinghtp.importFinal(table,2,cb);
                      },
                  function(cb){
                      Reportinghtp.importFinal(table,3,cb);
                      },
                  function(cb){
                      Reportinghtp.importFinal(table,4,cb);
                        },*/
                      
                  ],
                  function(err, resultat){
                    if (err) { return res.view('reporting/erreur'); }
                    return res.redirect('/export/'+dateexport +'/'+'<h1><h1>');
                })
              });
      }
      
  })
    
  },
  Essaii : function(req,res)
  {
    /*var Sails = require('sails').constructor;
    var sailsApp = new Sails();
    
    sailsApp.lift({
      log: { level: 'warn' }
    }, function (err) {
      if (err) {
        console.log('Error occurred lifting Sails app:', err);
        return;
      }
    
      // --•
      console.log('Sails app lifted successfully!');
    
    });*/
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    //var table = ['\\\\10.128.1.2\\almerys-out\\Retour_Easytech_'];
    //var table = ['//10.128.1.2/almerys-out/Retour_Easytech_'];
    var table = ['//10.128.1.2'];
    var table1 = ['\\\\10.128.1.2'];
    var table2 = ['10.128.1.2'];
    var table3 = ['\\\\share.easytech.mg'];
    var table4 = ['D:\\Reporting\\Reporting'];
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var date = annee+mois+jour;
    console.log(date);
    var cheminp = [];
    var MotCle= [];
    workbook.xlsx.readFile('ex.xlsx')
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
                    Reportinghtp.deleteFromChemin(table,cb);
                  },
               function(cb){
                    Reportinghtp.importEssai(table,cheminp,date,MotCle,0,cb);
                  },
               function(cb){
                    Reportinghtp.importEssai(table1,cheminp,date,MotCle,1,cb);
                  },
                  function(cb){
                    Reportinghtp.importEssai(table2,cheminp,date,MotCle,2,cb);
                  },
                function(cb){
                    Reportinghtp.importEssai(table3,cheminp,date,MotCle,3,cb);
                  },
                  function(cb){
                    Reportinghtp.importEssai(table4,cheminp,date,MotCle,4,cb);
                  },
            ],
            function(err, resultat){
              if (err) { return res.view('reporting/erreur'); }
              else
              {
                return res.view('reporting/accueil', {date : datetest});
              }
          });


        });
  },
  ReportingInovcom : function(req,res)
  {
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
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var date = annee+mois+jour;
    var nb = 5;
    workbook.xlsx.readFile('Inovcom.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil1');
          var chemincommun = newworksheet.getColumn(1);
          var cheminparticulier = newworksheet.getColumn(2);
          var dernierligne = newworksheet.getColumn(3);
          var feuille = newworksheet.getColumn(4);
          var cel = newworksheet.getColumn(5);
          var tab = newworksheet.getColumn(6);
          var cel2 = newworksheet.getColumn(7);
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
            for(var i=0;i<1;i++)
            {
              var a = cheminc[i]+date+cheminp[i]+date+dernierl[i];
              trameflux.push(a);
            };
            //console.log(trameflux);
            async.series([  
                function(cb){
                    Reportinghtp.deleteTout(table,nb,cb);
                  },  
              function(cb){
                  Reportinghtp.deleteHtp(table,nb,cb);
                },  
              function(cb){
                  Reportinghtp.importInovcom(trameflux,feuil,cellule,table,cellule2,cb);
                  },  
              /*function(cb){
                Reportinghtp.importTout(trameflux,table,cb);
                }, */
            ],
            function(err, resultat){
              if (err) { return res.view('reporting/erreur'); }
              return res.view('reporting/export');
          })
        });
  },
  ReportingInovcomType2 : function(req,res)
  {
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
    var numeroligne = [];
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var date1 = annee+mois+jour;
    var date2 = jour+'-'+mois+annee;
    var nb = 5;
    workbook.xlsx.readFile('InovcomType2.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil1');
          var chemincommun = newworksheet.getColumn(1);
          var cheminparticulier = newworksheet.getColumn(2);
          var dernierligne = newworksheet.getColumn(3);
          var feuille = newworksheet.getColumn(4);
          var cel = newworksheet.getColumn(5);
          var tab = newworksheet.getColumn(6);
          var cel2 = newworksheet.getColumn(7);
          var numerolign = newworksheet.getColumn(8);
            numerolign.eachCell(function(cell, rowNumber) {
                numeroligne.push(cell.value);
            });
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
            for(var i=0;i<3;i++)
            {
              var a = cheminc[i]+date1+cheminp[i]+dernierl[i];
              trameflux.push(a);
            };
            console.log(trameflux);
            async.series([  
                function(cb){
                    Reportinghtp.deleteTout(table,nb,cb);
                  },  
              function(cb){
                  Reportinghtp.deleteHtp(table,nb,cb);
                },  
              function(cb){
                  Reportinghtp.importInovcom(trameflux,feuil,cellule,table,cellule2,numeroligne,cb);
                  },  
              function(cb){
                Reportinghtp.importTout(trameflux,table,cb);
                }, 
            ],
            function(err, resultat){
                if (err) { return res.view('reporting/erreur'); }
                return res.redirect('/export/'+dateexport +'/'+'<h1><h1>');
          })
        });
  },
  accueil1 : function(req,res)
  {
    /*var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    var cheminp = [];
    var MotCle= [];
    workbook.xlsx.readFile('ex.xlsx')
        .then(function() {
          var newworksheet = workbook.getWorksheet('Feuil1');
          var cheminparticulier = newworksheet.getColumn(9);
          var motcle = newworksheet.getColumn(10);
            cheminparticulier.eachCell(function(cell, rowNumber) {
              //console.log(cell.value);
              cheminp.push(cell.value);
            });
            motcle.eachCell(function(cell, rowNumber) {
              MotCle.push(cell.value);
            });
        });*/
    return res.view('reporting/accueil1');
  },
  accueil : function(req,res)
  {
    /*async.series([  
        function(cb){
            Reportinghtp.importEssai(cb);
          },   
    ],
    function(err, resultat){
        if (err) { return res.view('reporting/erreur'); }
        return res.view('reporting/erera');
  })*/

    

    return res.view('reporting/accueil');
    /*XLSX = require('xlsx');
    var workbook = XLSX.readFile('D:/Reporting/Reporting/REPORTING HTP  Type.xlsx');
    //var first_sheet_name = workbook.SheetNames;
    //var data = XLSX.utils.sheet_to_json(workbook.Sheets[first_sheet_name[0]]);
    //console.log('longueur data' + data.length);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    var ref = sheet['!ref'];
    //var range = XLSX.utils.decode_range(sheet['!ref']);
   // console.log('Nombre de colonne' + range.e.c);
    console.log('refr'+ ref);
    console.log('refra'+ sheet.v);
    //console.log('ref'+sheet[!ref]);
    //console.log('Nombre de ligne' + range.e.r);
    /*read(filename) {
    const wb = XLSX.readFile(filename);
    for (let i = 0, l = wb.SheetNames.length; i < l; i += 1) {
    this.processSheet(wb.Sheets[wb.SheetNames[i]]);
    }
    }*/
                    /*row.eachCell(function(cell, colNumber) {
                    if(cell.text==cellule[nb])
                    {
                      a = parseInt(colNumber);
                    }
                  });
                  var col = newworksheet.getColumn(a);
                  var tab = [];
                  col.eachCell(function(cell, rowNumber) {
                    tab.push(cell.text);
                  });
                  var b;
                  row.eachCell(function(cell, colNumber) {
                    if(cell.text==cellule2[nb])
                    {
                      b = parseInt(colNumber);
                    }
                  });
                  var col2 = newworksheet.getColumn(b);
                  var tabl = [];
                  col2.eachCell(function(cell, rowNumber) {
                    var m = cell.text;
                    m=m.replace("'", "''");
                    console.log(m);
                    tabl.push(m);
                  });
                  var i;
                  var j = parseInt(i);
                  var cell;
                  var cell2;
                  var myJsonString = JSON.stringify(tabl);
                  //console.log(myJsonString.length);
                  var f = JSON.parse(myJsonString);
                  //console.log(f);
                 for(j=1;j<tab.length;j++)
                  {
                    cell = tab[j];
                    cell2 = tabl[j];
                    var sql = "insert into "+table[nb]+" (typologiedelademande,okko) values ('"+cell2+"','"+cell+"') ";
                    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                      if(err) return console.log(err);
                      else return callback(null, true);        
                                          })
                  };
              });*/
      
    //var html= req.param("html");
    //return res.view('reporting/accueil');
    /*const fs = require('fs');

    let lyrics = 'But still I\'m having memories of high speeds when the cops crashed\n' + 
                'As I laugh, pushin the gas while my Glocks blast\n';
    var tab = ['ok','ok','ko'];
    var tab1 = ['oko','oko','oko'];
    var com = [];
    for(var i=0;i<tab.length;i++)
    {
        com.push(tab[i]+';'+tab1[i]+'\n');
    };
    var f = 'ok' ;
    /*for(var i=0;i<5;i++)
    {
        tab.push(f+'\n');
    }
    /*var myJsonString = JSON.stringify(tab);
    console.log('myJson'+myJsonString);
    var f = JSON.parse(myJsonString);

    //m=m.replace("'", "''");
    console.log(f);*/
        /*fs.writeFile('4pac.txt', com, (err) => {
            // throws an error, you could also catch it here
            if (err) throw err;
    
            // success case, the file was saved
            console.log('Lyric saved!');
    
    });*/

    // write to a new file named 2pac.txt
    
    
  },
CompterExcel : function (req, res) {
      var calendrier = req.param('calendrier');
      const Excels = require('exceljs');
      // var sql = "select count(col_4) as ok from excel where col_4='ok'";
      var sql_flux = "select count(okko) as ok from trameflux where okko='OK'";
      var nb_flux;

      sails.sendNativeQuery(sql_flux, function(err, res){
          if (err) { return res.badRequest(err); }
          else {

          // var pdo = "select count(col_4) as ok from excel where col_4='ko'"; 
          var sql_flux_1 = "select count(okko) as ok from trameflux where okko='KO'";   
          var nb_flux_1;
              sails.sendNativeQuery(sql_flux_1, function(err, res){
                  if(err){
                      res.send(err);
                  }else{
                      
                      return nb_flux_1 = res.rows[0].ok;
                      
                  }
              });

              //REQUETE TABLE SUIVISAISIEPRODITE
              var sql_ITE = "select count(okko) as ok from suivisaisieprodite where okko='OK'";            
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
              var sql_MGAS = "select count(okko) as ok from suivisaisiemgas where okko='OK'";            
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
              var sql_LMDE = "select count(okko) as ok from suivisaisielmde where okko='OK'";            
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

