/**
 * ReportingIndu.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
const path_reporting = 'D:/LDR8_1421_nouv/PROJET_FELANA/REPORTING INDU Type.xlsx';
module.exports = {

  attributes: {
  },
  lectureEtInsertion2:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = numligne[nb];
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var col ;
      console.log('Nombre de colonne' + range.e.c);
      console.log('Nombre de ligne' + range.e.r);
      for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(desired_value==cellule[nb])
        {
          col=ra;
        }
      };
      console.log('colonne cible' +col);
      var tab = [];
      var tabl = [];
      var col2;
      for(var ra=0;ra<=range.e.c;ra++)
        {
          var address_of_cell = {c:ra, r:numeroligne};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          if(desired_value==cellule2[nb])
          {
            col2=ra;
          }
        };
      console.log('colonne cible2' +col2);
      if(col!=undefined && col2!=undefined)
      {
        for(var a=0;a<=range.e.r;a++)
          {
            var address_of_cell = {c:col, r:a};
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            var desired_value1 = (desired_cell ? desired_cell.v : undefined);
            var address_of_cell2 = {c:col2, r:a};
            var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
            var desired_cell2 = sheet[cell_ref2];
            var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
            var sql = "insert into "+table[nb]+" (typologiedelademande,okko) values ('"+desired_value2+"','"+desired_value1+"') ";
                      ReportingIndu.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err) return console.log(err);
                        else return callback(null, true);        
                                            })
          }
      }
      else
      {
        console.log('Colonne non trouvé');
      }
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
  },
  importTrameFlux929 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
    var tab = [];
    tab = ReportingIndu.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      ReportingIndu.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log(j);
      ReportingIndu.lectureEtInsertion2( trameflux,feuil,cellule,table,cellule2,j,numligne,callback)
    }
    };
  },
deleteFromChemin : function (table,callback) {
      var sql = "delete from chemininovcom ";
      ReportingIndu.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    existenceFichier : function (pathparam) {
      const fs = require('fs');
  
        var existe ='vrai';
        try{
          fs.accessSync(pathparam, fs.constants.F_OK);
        
        }catch(e){
          console.log(e);
          existe = 'faux';
        }
        return existe;
    },
    importEssai: function (table,table2,date,option,nb,callback) {
       const fs = require('fs');
      var re  = 'a';
      var tab = [];
      var a = table[0]+date+table2[nb];
      var b = option[nb];
      var c = ReportingIndu.existenceFichier(a);
      console.log(c);
      if(c=='vrai')
      {
        fs.readdir(a, (err, files) => {
          console.log(a);
              files.forEach(file => {
                const regex = new RegExp(b+'*');
                if(regex.test(file))
                {
                   re = file;
                   console.log(re);  
                } 
            });
            var sql = "insert into cheminindu (typologiedelademande) values ('"+re+"') ";
                    ReportingIndu.getDatastore().sendNativeQuery(sql, function(err,res){
                      if(err) return console.log(err);
                      else return callback(null, true);        
                                          })  
            console.log('ato anatiny'+re);
          });
      }
      else
      {
        var sql = "insert into cheminindu (typologiedelademande) values ('k') ";
        ReportingIndu.getDatastore().sendNativeQuery(sql, function(err,res){
          if(err) return console.log(err);
          else return callback(null, true);        
                              })   
      }   
    },
    deleteToutHtp : function (table,nb,callback) {
      var sql = "delete from "+table+" ";
      ReportingIndu.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    totalFichierExistant : function (trameflux,nb,callback) {
      var tab = [];
      var j;
      var i = parseInt(j);
      for(i=0;i<nb;i++)
      {
        var a = ReportingIndu.existenceFichier(trameflux[i]);
        if(a=='vrai')
        {
          tab.push(i);
        }
        else
        {
          console.log('faux');
        }
      };
      return tab ;
  
    },
    deleteTout: function (table,nb,callback) {
      for(var i=0;i<nb;i++){
        ReportingIndu.deleteFichier(table,i,callback);
      };
    },
    deleteFichier: function (table,nb,callback) {
      var tab= '';
      console.log(tab);
      const fs = require('fs');
      fs.writeFile(table[nb]+'.txt', tab, (err) => {
        var sql = "insert into trame (typologiedelademande) values ('k') ";
        ReportingIndu.getDatastore().sendNativeQuery(sql, function(err,res){
          if(err) return console.log(err);
          else return callback(null, true);        
                              })      
      });
    },
    
    deleteReportingHtp : function (table,nb,callback) {
      var sql = "delete from "+table[nb]+" ";
      ReportingIndu.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return console.log(err); }
        return callback(null, true);
        });
    },
    deleteHtp : function (table,nb,callback) {
      var j;
      var i = parseInt(j);
      for(i=0;i<5;i++)
      {
        ReportingIndu.deleteReportingHtp(table,i,callback);
      };
    },


/**********************************************************************************/
 // Récuperer nombre OK ou KO
 countOkKoDouble : function (table, callback) {
  const Excel = require('exceljs');
  var sqlOk ="select nbok from "+table; 
  var sqlKo ="select nbko from "+table;
 
  console.log(sqlOk);
  console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sqlOk, function(err, res){
        if (err) return res.badRequest(err);
        callback(null, res.rows[0].nbok);
      });
    },
    function (callback) {
      ReportingIndu.query(sqlKo, function(err, resKo){
        if (err) return res.badRequest(err);
        callback(null, resKo.rows[0].nbko);
      });
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK ==> " + result[0]);
    console.log("Count KO ==> " + result[1]);
    var okko = {};
    okko.ok = result[0];
    okko.ko = result[1];      
    return callback(null, okko);
  })
},
/******************************************************************************************/
countOkKoDoubleSum : function (table, callback) {
  const Excel = require('exceljs');
  var sqlOk ="select sum(nbok::integer) from "+table; 
  var sqlKo ="select sum(nbko::integer) from  "+table;
 
  console.log(sqlOk);
  console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sqlOk, function(err, res){
        if (err) return res.badRequest(err);
        callback(null, res.rows[0].sum);
      });
    },
    function (callback) {
      ReportingIndu.query(sqlKo, function(err, resKo){
        if (err) return res.badRequest(err);
        callback(null, resKo.rows[0].sum);
      });
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK ==> " + result[0]);
    console.log("Count KO ==> " + result[1]);
    var okko = {};
    okko.ok = result[0];
    okko.ko = result[1];      
    return callback(null, okko);
  })
},
/**********************************************************************************************/
countOkKo : function (table, callback) {
  const Excel = require('exceljs');
  var sqlOk ="select nbok from "+table; //trameFlux
  // var sqlKo ="select nbko from "+table;
  console.log(sqlOk);
  // console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingInovcomExport.query(sqlOk, function(err, res){
        if (err) return res.badRequest(err);
        callback(null, res.rows[0].nbok);
      });
     },
    // function (callback) {
    //   ReportingInovcomExport.query(sqlKo, function(err, resKo){
    //     if (err) return res.badRequest(err);
    //     callback(null, resKo.rows[0].nbko);
    //   });
    // },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK ==> " + result[0]);
    // console.log("Count KO ==> " + result[1]);
    var okko = {};
    okko.ok = result[0];
    // okko.ko = result[1];
    return callback(null, okko);
  })
},
/********************************************************************************************/
countOkKoSum : function (table, callback) {
  const Excel = require('exceljs');
  // var sqlOk ="select count(okko) as ok from "+table+" where okko='OK'"; //trameFlux
  // var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'";
  var sql ="select sum(nbok::integer) from "+table; 
 
  console.log(sql);
  // console.log(sqlOk);
  // console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sql, function(err, res){
        if (err) return res.badRequest(err);
        // callback(null, res.rows[0].ok);
        console.log(res.rows[0].sum);
        if(res.rows[0].sum != undefined){
          callback(null, res.rows[0].sum);
        }
        else{
          return res.rows[0].sum = 0;
        }
        
      });
    },
    // function (callback) {
    //   Retour.query(sqlKo, function(err, resKo){
    //     if (err) return res.badRequest(err);
    //     callback(null, resKo.rows[0].ko);
    //   });
    // },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK somme ==> " + result[0]);
    // console.log("Count KO ==> " + result[1]);
    var okko = {};
    okko.ok = result[0];
    // okko.ko = result[1];      
    return callback(null, okko);
    // var intro = {};
    // intro.nb = result[0];
    // return callback(null, intro);
  })
},
countOkKoSumko : function (table, callback) {
  const Excel = require('exceljs');
  // var sqlOk ="select count(okko) as ok from "+table+" where okko='OK'"; //trameFlux
  // var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'";
  var sql ="select sum(nbko::integer) from "+table; 
 
  console.log(sql);
  // console.log(sqlOk);
  // console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sql, function(err, res){
        if (err) return res.badRequest(err);
        // callback(null, res.rows[0].ok);
        console.log(res.rows[0].sum);
        if(res.rows[0].sum != undefined){
          callback(null, res.rows[0].sum);
        }
        else{
          return res.rows[0].sum = 0;
        }
        
      });
    },
    // function (callback) {
    //   Retour.query(sqlKo, function(err, resKo){
    //     if (err) return res.badRequest(err);
    //     callback(null, resKo.rows[0].ko);
    //   });
    // },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK ==> " + result[0]);
    // console.log("Count KO ==> " + result[1]);
    var okko = {};
    okko.ok = result[0];
    // okko.ko = result[1];      
    return callback(null, okko);
    // var intro = {};
    // intro.nb = result[0];
    // return callback(null, intro);
  })
},
/***************************************************************************/
countOkKoIndu2 : function (table, callback) {
  const Excel = require('exceljs');
  var sqlOk ="select nbok from "+table; 
  var sqlKo ="select nbko from "+table;
 
  console.log(sqlOk);
  console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sqlOk, function(err, res){
        if (err) return res.badRequest(err);
        callback(null, res.rows[0].nbok);
      });
    },
    function (callback) {
      ReportingIndu.query(sqlKo, function(err, resKo){
        if (err) return res.badRequest(err);
        callback(null, resKo.rows[0].nbko);
      });
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK ==> " + result[0]);
    console.log("Count KO ==> " + result[1]);
    var okko = {};
    okko.ok = result[0];
    okko.ko = result[1];      
    return callback(null, okko);
  })
},
  // Convert date
  convertDate : function (dateExcel){
    var date = new Date(dateExcel);
    var year = date.getFullYear();
    var month = date.getMonth()+1;
    var dt = date.getDate();
    if (dt < 10) {
      dt = '0' + dt;
    }
    if (month < 10) {
      month = '0' + month;
    }
    return dt +"/"+ month +"/"+year;
  },
/************************************************************************* */
ecritureOkKo : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
   
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(mois1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = ReportingIndu.convertDate(cell.text);
    // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
    // var valiny = Retour.convertDate(andro);
    // console.log(valiny);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(3).value;
      // console.log(f);
      if(f == "almerys")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  var iniValue = ReportingIndu.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
 

  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'DOCUMENTS SAISIS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      if(f == iniValue.ok)
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = nombre_ok_ko.ok;
  
  await newWorkbook.xlsx.writeFile(path_reporting);
  sails.log("Ecriture OK KO terminé"); 
  return callback(null, "OK");

  }
  catch
  {
    console.log("Une erreur s'est produite");
    Reportinghtp.deleteToutHtp(table,3,callback);
  }
  },
  /***************************************************************/
  ecritureOkKoDouble : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
     
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingIndu.convertDate(cell.text);
      // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
      // var valiny = Retour.convertDate(andro);
      // console.log(valiny);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        // console.log(f);
        if(f == "almerys")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingIndu.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
    var colonnne;
    var colDate1;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS SAISIS')
      {
        colDate1 = parseInt(colNumber);
        //var col = newworksheet.getColumn(colDate1);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate1).value;
        if(f == iniValue.ok)
        {
          colonnne = parseInt(colNumber);
        }
        }
    });
    console.log(" Colnumber"+colonnne);
  
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        if(f == iniValue.ok)
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(colonnne).value = nombre_ok_ko.ok;
    numeroLigne.getCell(collonne).value = nombre_ok_ko.ko;
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO terminé"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
    },
    /***************************************************************/
  ecritureOkKoIndu2 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingIndu.convertDate(cell.text);
      // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
      // var valiny = Retour.convertDate(andro);
      // console.log(dateExcel);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        // console.log(f);
        if(f == "almerys")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingIndu.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
    var colonnne;
    var colDate1;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS SAISIS')
      {
        colDate1 = parseInt(colNumber);
        //var col = newworksheet.getColumn(colDate1);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate1).value;
        if(f == iniValue.ok)
        {
          colonnne = parseInt(colNumber);
        }
        }
    });
    console.log(" Colnumber"+colonnne);
  
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        if(f == iniValue.ok)
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(colonnne).value = nombre_ok_ko.ok;
    numeroLigne.getCell(collonne).value = nombre_ok_ko.ko;
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO terminé"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
    },
    /***************************************************************/
    ecritureOkKoIndu2cbtp : async function (nombre_ok_ko, table,date_export,mois1,callback) {
      const Excel = require('exceljs');
      const cmd=require('node-cmd');
      const newWorkbook = new Excel.Workbook();
      
      try{
      
        await newWorkbook.xlsx.readFile(path_reporting);
      const newworksheet = newWorkbook.getWorksheet(mois1);
      var colonneDate = newworksheet.getColumn('A');
      var ligneDate1;
      var ligneDate;
      colonneDate.eachCell(function(cell, rowNumber) {
        var dateExcel = ReportingIndu.convertDate(cell.text);
        // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
        // var valiny = Retour.convertDate(andro);
        // console.log(dateExcel);
        if(dateExcel==date_export)
        {
          ligneDate1 = parseInt(rowNumber);
          var line = newworksheet.getRow(ligneDate1);
          var f = line.getCell(3).value;
          // console.log(f);
          if(f == "cbtp")
          {
            ligneDate = parseInt(rowNumber);
          }
        }
      });
      console.log("LIGNE DATE ===> "+ ligneDate);
      var rowDate = newworksheet.getRow(ligneDate);
      var numeroLigne = rowDate;
      var iniValue = ReportingIndu.getIniValue(table);
      
      var a5;
    
      var rowm = newworksheet.getRow(1);
      var colonnne;
      var colDate1;
      rowm.eachCell(function(cell, colNumber) {
        if(cell.value == 'DOCUMENTS SAISIS')
        {
          colDate1 = parseInt(colNumber);
          //var col = newworksheet.getColumn(colDate1);
          var man = newworksheet.getRow(3);
          var f = man.getCell(colDate1).value;
          if(f == iniValue.ok)
          {
            colonnne = parseInt(colNumber);
          }
          }
      });
      console.log(" Colnumber"+colonnne);
    
      var collonne;
      var colDate2;
      rowm.eachCell(function(cell, colNumber) {
        if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
        {
          colDate2 = parseInt(colNumber);
          var man = newworksheet.getRow(3);
          var f = man.getCell(colDate2).value;
          if(f == iniValue.ok)
          {
            collonne = parseInt(colNumber);
          }
        }
      });
      console.log(" Colnumber2"+collonne);
      numeroLigne.getCell(colonnne).value = nombre_ok_ko.ok;
      numeroLigne.getCell(collonne).value = nombre_ok_ko.ko;
      await newWorkbook.xlsx.writeFile(path_reporting);
      sails.log("Ecriture OK KO terminé"); 
      return callback(null, "OK");
    
      }
      catch
      {
        console.log("Une erreur s'est produite");
        Reportinghtp.deleteToutHtp(table,3,callback);
      }
      },
      /***************************************************************/
  getConfigIni : function() {
    const fs = require('fs');
    const ini = require('ini');
    const config = ini.parse(fs.readFileSync('./config_excelIndu.ini', 'utf-8'));
    console.log(config);
    return config;
  },

  getIniValue : function(table) {
    var iniValue = ReportingIndu.getConfigIni();
    var numeroColonneOk,numeroColonneKo;
    
    if(table == "indurelevedecomptealmerys"){
      numeroColonneOk = iniValue.indurelevedecomptealmerys.ok;
      numeroColonneKo = iniValue.indurelevedecomptealmerys.ko;
    }
    if(table == "indurelevedecomptecbtp"){
      numeroColonneOk = iniValue.indurelevedecomptecbtp.ok;
      numeroColonneKo = iniValue.indurelevedecomptecbtp.ko;
    }
    if(table == "induse"){
      numeroColonneOk = iniValue.induse.ok;
      numeroColonneKo = iniValue.induse.ko;
    }
    if(table == "induhospi"){
      numeroColonneOk = iniValue.induhospi.ok;
      numeroColonneKo = iniValue.induhospi.ko;
    }
    if(table == "indusansnotif"){
      numeroColonneOk = iniValue.indusansnotif.ok;
      numeroColonneKo = iniValue.indusansnotif.ko;
    }
   if(table == "indutiers"){
      numeroColonneOk = iniValue.indutiers.ok;
      numeroColonneKo = iniValue.indutiers.ko;
    }
    if(table == "indufraudelmg"){
      numeroColonneOk = iniValue.indufraudelmg.ok;
      numeroColonneKo = iniValue.indufraudelmg.ko;
    }
    if(table == "indufraudelmg"){
      numeroColonneOk = iniValue.indufraudelmg.ok;
      numeroColonneKo = iniValue.indufraudelmg.ko;
    }
    if(table == "induinterialepre"){
      numeroColonneOk = iniValue.induinterialepre.ok;
      numeroColonneKo = iniValue.induinterialepre.ko;
    }
    if(table == "induinterialepost"){
      numeroColonneOk = iniValue.induinterialepost.ok;
      numeroColonneKo = iniValue.induinterialepost.ko;
    }
    if(table == "inducodelisftp"){
      numeroColonneOk = iniValue.inducodelisftp.ok;
      numeroColonneKo = iniValue.inducodelisftp.ko;
    }
    if(table == "inducodelismail"){
      numeroColonneOk = iniValue.inducodelismail.ok;
      numeroColonneKo = iniValue.inducodelismail.ko;
    }
    if(table == "inducodelisappel"){
      numeroColonneOk = iniValue.inducodelisappel.ok;
      numeroColonneKo = iniValue.inducodelisappel.ko;
    }
    if(table == "inducheque"){
      numeroColonneOk = iniValue.inducheque.ok;
      numeroColonneKo = iniValue.inducheque.ko;
    }
    if(table == "indupecrefus"){
      numeroColonneOk = iniValue.indupecrefus.ok;
      numeroColonneKo = iniValue.indupecrefus.ko;
    }
    if(table == "induinterialeaudio"){
      numeroColonneOk = iniValue.induinterialeaudio.ok;
      numeroColonneKo = iniValue.induinterialeaudio.ko;
    }
    // if(table == ""){
    //   numeroColonneOk = iniValue..ok;
    //   numeroColonneKo = iniValue..ko;
    // }
   
    var ok_ko = {};
    ok_ko.ok = numeroColonneOk;
    ok_ko.ko = numeroColonneKo;

    console.log("INI OK = "+ok_ko.ok);
    console.log("INI KO = "+ok_ko.ko);
    return ok_ko;
  },


};

