/**
 * ReportingExcel.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
const path_reporting = '/dev/prod/00-TOUS/TestReporting/REPORTING HTP  Type.xlsx';
//const path_reporting = 'D:/Reporting/Reporting/Nouveau dossier/REPORTING HTP.xlsx';
//const path_reporting = 'D:/LDR8_1421_nouv/PROJET_FELANA/REPORTING HTP Type.xlsx';

module.exports = {
  attributes: {
  },
  // Récuperer nombre OK ou KO
  countOkKo : function (table, callback) {
    const Excel = require('exceljs');
    var sqlOk ="select sum(nbok) as ok from "+table+" "; //trameFlux
    var sqlKo ="select sum(nbko) as ko from "+table+" ";
    /*console.log(sqlOk);
    console.log(sqlKo);*/
    async.series([
      function (callback) {
        ReportingExcel.getDatastore().sendNativeQuery(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].ok);
        });
      },
      function (callback) {
        ReportingExcel.getDatastore().sendNativeQuery(sqlKo, function(err, resKo){
          if (err) return res.badRequest(err);
          callback(null, resKo.rows[0].ko);
        });
      },
    ],function(err,result){
      if(err) return res.badRequest(err);
     // console.log("Count OK ==> " + result[0]);
      console.log("Count KO ==> " + result[1]);
      var okko = {};
      okko.ok = result[0];
      okko.ko = result[1];
      return callback(null, okko);
    })
  },
  countOkKoTrameLamieResiliation : function (table, callback) {
    const Excel = require('exceljs');
    var sqlOk ="select sum(nbokresiliation) as ok from "+table+" "; //trameFlux
    var sqlKo ="select sum(nbkoresiliation) as ko from "+table+" ";
    /*console.log(sqlOk);
    console.log(sqlKo);*/
    async.series([
      function (callback) {
        ReportingExcel.getDatastore().sendNativeQuery(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].ok);
        });
      },
      function (callback) {
        ReportingExcel.getDatastore().sendNativeQuery(sqlKo, function(err, resKo){
          if (err) return res.badRequest(err);
          callback(null, resKo.rows[0].ko);
        });
      },
    ],function(err,result){
      if(err) return res.badRequest(err);
      //console.log("Count OK ==> " + result[0]);
      console.log("Count KO ==> " + result[1]);
      var okko = {};
      okko.ok = result[0];
      okko.ko = result[1];
      return callback(null, okko);
    })
  },
  countOkKoTrameLamie : function (table, callback) {
    const Excel = require('exceljs');
    var sqlOk ="select sum(nbok) as ok from "+table+" "; //trameFlux
    var sqlKo ="select sum(nbko) as ko from "+table+" ";
    console.log(sqlOk);
    console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingExcel.getDatastore().sendNativeQuery(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].ok);
        });
      },
      function (callback) {
        ReportingExcel.getDatastore().sendNativeQuery(sqlKo, function(err, resKo){
          if (err) return res.badRequest(err);
          callback(null, resKo.rows[0].ko);
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
  countOkKoTrameLamie2 : function (table, callback) {
    const Excel = require('exceljs');
    var sqlOk ="select count(okko) as ok from "+table+" where okko='OK' AND typologiedelademande='Résiliation' "; //trameFlux
    var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'  AND typologiedelademande='Résiliation' ";
    /*console.log(sqlOk);
    console.log(sqlKo);*/
    async.series([
      function (callback) {
        ReportingExcel.getDatastore().sendNativeQuery(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].ok);
        });
      },
      function (callback) {
        ReportingExcel.getDatastore().sendNativeQuery(sqlKo, function(err, resKo){
          if (err) return res.badRequest(err);
          callback(null, resKo.rows[0].ko);
        });
      },
    ],function(err,result){
      if(err) return res.badRequest(err);
      //console.log("Count OK ==> " + result[0]);
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
  // Ecriture nb OK et nb KO
  /*ecritureOkKo : function (nombre_ok_ko, table,date_export,callback) {
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    newWorkbook.xlsx.readFile(path_reporting)
    .then(function(){

  
    const newworksheet = newWorkbook.getWorksheet('Janvier');
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingExcel.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        //console.log();
        if(f == "Pack normal")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingExcel.getIniValue(table);
    
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
        //console.log();
        //console.log(iniValue.ok);
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
    newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO terminé"); 
    return callback(null, "OK");
    })
    console.log("Une erreur s'est produite");
    Reportinghtp.deleteToutHtp(table,3,callback);

    },*/
 ecritureOkKo : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingExcel.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        /*console.log(dateExcel);
        console.log(date_export);*/
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        //console.log();
        if(f == "Pack normal")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    //console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingExcel.getIniValue(table);
    
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
        //console.log();
        //console.log(iniValue.ok);
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
    //console.log(" Colnumber2"+collonne);

    var ok = parseInt(nombre_ok_ko.ok);
    var ko = parseInt(nombre_ok_ko.ko);
    numeroLigne.getCell(colonnne).value = ok;
    numeroLigne.getCell(collonne).value = ko;
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
  getConfigIni : function() {
    const fs = require('fs');
    const ini = require('ini');
    const config = ini.parse(fs.readFileSync('./config_excel.ini', 'utf-8'));
    //console.log(config);
    return config;
  },

  getIniValue : function(table) {
    var iniValue = ReportingExcel.getConfigIni();
    var numeroColonneOk,numeroColonneKo;

    if(table == "trameflux"){
      numeroColonneOk = iniValue.trame_flux.ok;
      numeroColonneKo = iniValue.trame_flux.ko;
    }
    if(table == "suivisaisielmde"){
      numeroColonneOk = iniValue.suivi_saisie_lmde.ok;
      numeroColonneKo = iniValue.suivi_saisie_lmde.ko;
    }
    if(table == "suivisaisiemgas"){
      numeroColonneOk = iniValue.suivi_saisie_mgas.ok;
      numeroColonneKo = iniValue.suivi_saisie_mgas.ko;
    }
    if(table == "suivisaisieprodite"){
      numeroColonneOk = iniValue.suivi_saisie_ite.ok;
      numeroColonneKo = iniValue.suivi_saisie_ite.ko;
    }
    if(table == "tramelamiestock"){
      numeroColonneOk = iniValue.trame_lamie_stock.ok;
      numeroColonneKo = iniValue.trame_lamie_stock.ko;
    }
    if(table == "tramelamiestocknr"){
      numeroColonneOk = iniValue.trame_lamie_stock_nr.ok;
      numeroColonneKo = iniValue.trame_lamie_stock_nr.ko;
    }
    var ok_ko = {};
    ok_ko.ok = numeroColonneOk;
    ok_ko.ko = numeroColonneKo;

    console.log("INI OK = "+ok_ko.ok);
    console.log("INI KO = "+ok_ko.ko);
    return ok_ko;
  },

  //________________________________________________________________________________________________
  //________________________________________________________________________________________________
  //________________________________________________________________________________________________
  /*rechercheLigneColonne : function (nbChemin, callback) {
    async function exTest()
    {
      const Excel = require('exceljs');
      var nb= nbChemin.nb;
      var chemin= nbChemin.chemin;
      const newWorkbook = new Excel.Workbook();

      await newWorkbook.xlsx.readFile('D:/Erica/ALMERYS REPORTING/NEED Almerys_reporting/REPORTING HTP  Type.xlsx');
      const newworksheet = newWorkbook.getWorksheet('Mois');
      var row = newworksheet.getRow(6);
      var colonne,b;
      row.eachCell(function(cell, colNumber) {
        var cellText = cell.text.toString();
        var cheminText = chemin.toString();
        console.log("[ "+cellText + "=" + cheminText +" ]");
        console.log("---------------------------------------");
        if(cellText == cheminText)
        {
          console.log("Chemin OK");
          console.log(cell.text);
          colonne = parseInt(colNumber);
        }
      });
      console.log("COLONNE ===> "+colonne);
      var col = newworksheet.getColumn(3);
      col.eachCell(function(cell, rowNumber) {
        if(cell.value=='LAMIE1')
        {
          b = parseInt(rowNumber);
        }
      });
      var rowVrai = newworksheet.getRow(b);
      rowVrai.getCell(colonne).value = nb;
      await newWorkbook.xlsx.writeFile('D:/Erica/ALMERYS REPORTING/NEED Almerys_reporting/REPORTING HTP  Type.xlsx')
      sails.log(colonne + "b="+ b);  
    }
    exTest();  
    return callback(null, "OK");
  },

  rechercheColonne : function (nb, callback) {
    async function exTest()
    {
      const Excel = require('exceljs');
      //nb= 10;
      const newWorkbook = new Excel.Workbook();

      await newWorkbook.xlsx.readFile('D:/Erica/ALMERYS REPORTING/NEED Almerys_reporting/REPORTING HTP  Type.xlsx');
      const newworksheet = newWorkbook.getWorksheet('Janvier');
      var row = newworksheet.getRow(3);
      var a,b;
      row.eachCell(function(cell, colNumber) {
        if(cell.value=='report 929')
        {
          a = parseInt(colNumber);
        }
      });
      var col = newworksheet.getColumn(3);
      col.eachCell(function(cell, rowNumber) {
        if(cell.value=='LAMIE1')
        {
          b = parseInt(rowNumber);
        }
      });
      var rowVrai = newworksheet.getRow(b);
      rowVrai.getCell(a).value = nb;
      await newWorkbook.xlsx.writeFile('D:/Erica/ALMERYS REPORTING/NEED Almerys_reporting/REPORTING HTP  Type.xlsx')
      sails.log(a + "b="+ b);  
    }
    exTest();  
    return callback(null, "OK");
  }*/

};

