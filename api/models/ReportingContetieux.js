/**
 * ReportingContetieux.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
const path_reporting = 'D:/LDR8_1421_nouv/PROJET_FELANA/REPORTING CONTENTIEUX type.xlsx';

module.exports = {
  attributes: {
  },
  // Récuperer nombre OK ou KO
  countOkKo : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlOk ="select count(okko) as ok from "+table+" where okko='OK'"; //trameFlux
    // var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'";
    var sql ="select * from "+table; 
   
    console.log(sql);
    // console.log(sqlOk);
    // console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingContetieux.query(sql, function(err, res){
          if (err) return res.badRequest(err);
          // callback(null, res.rows[0].ok);
          console.log(res.rows[0].nb);
          if(res.rows[0].nb != undefined){
            callback(null, res.rows[0].nb);
          }
          else{
            return res.rows[0].nb = 0;
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
  countOkKoSum : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlOk ="select count(okko) as ok from "+table+" where okko='OK'"; //trameFlux
    // var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'";
    var sql ="select sum(nb::integer) from "+table; 
   
    console.log(sql);
    // console.log(sqlOk);
    // console.log(sqlKo);
    async.series([
      function (callback) {
        Retour.query(sql, function(err, res){
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
      var dateExcel = ReportingContetieux.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        //console.log();
        if(f == "ALMERYS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingContetieux.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
    var colonnne;
    var colDate1;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
      {
        colDate1 = parseInt(colNumber);
        
        //var col = newworksheet.getColumn(colDate1);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate1).value;
        var getko_ini = man.getCell(colDate1).address;
        if(getko_ini == iniValue.ko+3 && f == iniValue.ok)
        {
          colonnne = parseInt(colNumber);
        }
        }
    });
    console.log(" Colnumber"+colonnne);
   
    numeroLigne.getCell(colonnne).value = nombre_ok_ko.ok;
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
    /****************************************************************/
    ecritureOkKo2 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
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
        var dateExcel = ReportingContetieux.convertDate(cell.text);
        if(dateExcel==date_export)
        {
          ligneDate1 = parseInt(rowNumber);
          var line = newworksheet.getRow(ligneDate1);
          var f = line.getCell(3).value;
          //console.log();
          if(f == "CBTP")
          {
            ligneDate = parseInt(rowNumber);
          }
        }
      });
      console.log("LIGNE DATE ===> "+ ligneDate);
      var rowDate = newworksheet.getRow(ligneDate);
      var numeroLigne = rowDate;
      var iniValue = ReportingContetieux.getIniValue(table);
      
      var a5;
    
      var rowm = newworksheet.getRow(1);
      var colonnne;
      var colDate1;
      rowm.eachCell(function(cell, colNumber) {
        if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
        {
          colDate1 = parseInt(colNumber);
          
          //var col = newworksheet.getColumn(colDate1);
          var man = newworksheet.getRow(3);
          var f = man.getCell(colDate1).value;
          var getko_ini = man.getCell(colDate1).address;
          if(getko_ini == iniValue.ko+3 && f == iniValue.ok)
          {
            colonnne = parseInt(colNumber);
          }
          }
      });
      console.log(" Colnumber"+colonnne);
     
      numeroLigne.getCell(colonnne).value = nombre_ok_ko.ok;
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
    const config = ini.parse(fs.readFileSync('./config_excelContentieux.ini', 'utf-8'));
    console.log(config);
    return config;
  },

  getIniValue : function(table) {
    var iniValue = ReportingContetieux.getConfigIni();
    var numeroColonneOk,numeroColonneKo;
    if(table == "coaaotdalmerys"){
      numeroColonneOk = iniValue.coaaotdalmerys.ok;
      numeroColonneKo = iniValue.coaaotdalmerys.ko;
    }
    if(table == "coldralmeryspublic"){
      numeroColonneOk = iniValue.coldralmeryspublic.ok;
      numeroColonneKo = iniValue.coldralmeryspublic.ko;
    }
    if(table == "cootdalmerys"){
      numeroColonneOk = iniValue.cootdalmerys.ok;
      numeroColonneKo = iniValue.cootdalmerys.ko;
    }
    if(table == "cosdralmerys"){
      numeroColonneOk = iniValue.cosdralmerys.ok;
      numeroColonneKo = iniValue.cosdralmerys.ko;
    }
   if(table == "cootdclient"){
      numeroColonneOk = iniValue.cootdclient.ok;
      numeroColonneKo = iniValue.cootdclient.ko;
    }
    if(table == "coadraphpalmerys"){
      numeroColonneOk = iniValue.coadraphpalmerys.ok;
      numeroColonneKo = iniValue.coadraphpalmerys.ko;
    }
    if(table == "coadrclassiquealmerys"){
      numeroColonneOk = iniValue.coadrclassiquealmerys.ok;
      numeroColonneKo = iniValue.coadrclassiquealmerys.ko;
    }
    if(table == "coimputationalmerys"){
      numeroColonneOk = iniValue.coimputationalmerys.ok;
      numeroColonneKo = iniValue.coimputationalmerys.ko;
    }
    if(table == "coaaotdcbtp"){
      numeroColonneOk = iniValue.coaaotdcbtp.ok;
      numeroColonneKo = iniValue.coaaotdcbtp.ko;
    }
    if(table == "coldrcbtppublic"){
      numeroColonneOk = iniValue.coldrcbtppublic.ok;
      numeroColonneKo = iniValue.coldrcbtppublic.ko;
    }
    if(table == "cootdcbtp"){
      numeroColonneOk = iniValue.cootdcbtp.ok;
      numeroColonneKo = iniValue.cootdcbtp.ko;
    }
    if(table == "cosdrcbtp"){
      numeroColonneOk = iniValue.cosdrcbtp.ok;
      numeroColonneKo = iniValue.cosdrcbtp.ko;
    }
   if(table == "coadraphpcbtp"){
      numeroColonneOk = iniValue.coadraphpcbtp.ok;
      numeroColonneKo = iniValue.coadraphpcbtp.ko;
    }
    if(table == "coadrclassiquecbtp"){
      numeroColonneOk = iniValue.coadrclassiquecbtp.ok;
      numeroColonneKo = iniValue.coadrclassiquecbtp.ko;
    }
    if(table == "coimputationcbtp"){
      numeroColonneOk = iniValue.coimputationcbtp.ok;
      numeroColonneKo = iniValue.coimputationcbtp.ko;
    }
   
    var ok_ko = {};
    ok_ko.ok = numeroColonneOk;
    ok_ko.ko = numeroColonneKo;

    console.log("INI OK = "+ok_ko.ok);
    console.log("INI KO = "+ok_ko.ko);
    return ok_ko;
  },



};
