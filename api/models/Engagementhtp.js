const { Console } = require('console');

/**
 * Engagementhtp.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
// const path_reporting = 'D:/LDR8_1421_nouv/PROJET_FELANA/HTP/Test.xlsx';
// const path_reporting = 'D:/LDR8_1421_nouv/PROJET_FELANA/HTP/REPORTING_RETOUR.xlsx';
const path_reporting = '/dev/prod/00-TOUS/TestReporting/Test.xlsx';
module.exports = {

  attributes: {

  },
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
  /******************************/
  //CONVERTION DATE EXCEL
  convertionexceldate : function (serial){
    var utc_days  = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;                                        
    var date_info = new Date(utc_value * 1000);
    var fractional_day = serial - Math.floor(serial) + 0.0000001;
    var total_seconds = Math.floor(86400 * fractional_day);
    var seconds = total_seconds % 60;
 
    total_seconds -= seconds;
 
    var hours = Math.floor(total_seconds / (60 * 60));
    var minutes = Math.floor(total_seconds / 60) % 60;
 
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
  },
  //RECUPERATION VALEUR DANS LA BASE
  recupdata : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlOk ="select count(okko) as ok from "+table+" where okko='OK'"; //trameFlux
    // var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'";
    var sql ="select * from "+table ; 
   
    console.log(sql);
    // console.log(sqlOk);
    // console.log(sqlKo);
    async.series([
      function (callback) {
        Retour.query(sql, function(err, res){          
          if (err) {
            console.log(err);
            //return null;
          }
          else
          {
            if(res.rows[0])
            {
              console.log('ok');
              callback(null, res.rows[0].nb);
            }
            else
            {
              console.log("null");
              callback(null, 0);
            }
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
  /*******************************************************************************/
  recupdatasum : function (table, callback) {
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
          if (err) {
            console.log(err);
            //return null;
          }
          else
          {
            if(res.rows[0])
            {
              console.log('ok');
              callback(null, res.rows[0].sum);
            }
            else
            {
              console.log("null");
              callback(null, 0);
            }
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
  
  /***************************************************************/
  //fonction n'est pas encore en service
  ecrituredata16tri1 : function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    var workbook = new Excel.Workbook(); 
try{

   workbook.xlsx.readFile(path_reporting)
        .then(function() {
            var worksheet = workbook.getWorksheet(1);
            var colonneDate = worksheet.getColumn('A');
            var ligneDate1;
            var ligneDate;
            colonneDate.eachCell(function(cell, rowNumber) {
              var dateExcel = Retour.convertDate(cell.text);
              if(dateExcel==date_export)
              {
                ligneDate1 = parseInt(rowNumber);
                var line = worksheet.getRow(ligneDate1);
                var f = line.getCell(4).value;
                // console.log(f);
                if(f == "Tri MGEFI")
                {
                  ligneDate = parseInt(rowNumber);
                }
              }
            });
            console.log("LIGNE DATE ===> "+ ligneDate);

            var rowDate = worksheet.getRow(ligneDate);
            
            // var iniValue = Retour.getIniValue(table);

            var rowm = worksheet.getRow(1);

            var colonne;
            var colcible;
            rowm.eachCell(function(cell, colNumber) {
              if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
              {
                colcible = parseInt(colNumber);
                var man = worksheet.getRow(3);
                var f = man.getCell(colcible).value;
                // var getko_ini = man.getCell(colcible).address;
                // // console.log(getko_ini);
                // if(getko_ini == iniValue.ko+3 && f == iniValue.ok)
                if(f == "16H")
                {
                  colonne = parseInt(colNumber);
                }
              }
            });
            console.log(" Numero colonne: "+colonne);
            
            var numeroLigne = rowDate;
            numeroLigne.getCell(colonne).value = nombre_ok_ko;
            workbook.xlsx.writeFile(path_reporting);
            sails.log("Ecriture OK KO terminé"); 
            
        });  
        
        return callback(null, "OK");
        
      }

catch
      {
        console.log("Une erreur s'est produite");
        Reportinghtp.deleteToutHtp(table,3,callback);
      }
    
    },
  
 /************************************************************************************/ 
 //ECRITURE ET REMPLISSAGE FICHIER
 ecrituredata16tri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Tri MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "16H")
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

/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredata16facM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "16H")
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
 /************************************************************************************/ 
 //ECRITURE ET REMPLISSAGE FICHIER
 ecrituredata16devi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "16H")
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

   /************************************************************************************/ 
   //ECRITURE ET REMPLISSAGE FICHIER
   ecrituredata16sales : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Salesforce MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "16H")
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
 /************************************************************************************/ 
 //ECRITURE ET REMPLISSAGE FICHIER
 ecrituredata16flux : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Flux Noe MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "16H")
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
 /************************************************************************************/ 
 //ECRITURE ET REMPLISSAGE FICHIER
 ecrituredata16rejet : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Rejet MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "16H")
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
 /************************************************************************************/ 
 //ECRITURE ET REMPLISSAGE FICHIER
 ecrituredata16cotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "16H")
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
  /***********************************************************************************/
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredata16cotite : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "Contrat Cot ITE et MGAS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "16H")
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
  /***********************************************************************************/
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredata16faclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "Facture LAMIE")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "16H")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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
  /***********************************************************************************/
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredata16acs : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "ACS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "16H")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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

 /************************************************************************************/ 
 //ECRITURE ET REMPLISSAGE FICHIER
 ecrituredatafinptri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Tri MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "Fin  Prod")
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

/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredatafinpfacM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "Fin  Prod")
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
 /************************************************************************************/ 
 //ECRITURE ET REMPLISSAGE FICHIER
 ecrituredatafinpdevi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "Fin  Prod")
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

   /************************************************************************************/ 
   //ECRITURE ET REMPLISSAGE FICHIER
   ecrituredatafinpsales : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Salesforce MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "Fin  Prod")
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
 /************************************************************************************/ 
 //ECRITURE ET REMPLISSAGE FICHIER
 ecrituredatafinpflux : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Flux Noe MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "Fin  Prod")
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
 /************************************************************************************/ 
 ecrituredatafinprejet : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Rejet MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "Fin  Prod")
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
 /************************************************************************************/ 
 ecrituredatafinpcotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "Fin  Prod")
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
  /************************************************************************************/ 
 ecrituredatafinpcotite : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot ITE et MGAS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "Fin  Prod")
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
    /************************************************************************************/ 
  ecrituredatafinpfaclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Facture LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "Fin  Prod")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
    /************************************************************************************/ 
    ecrituredatafinpacs : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "ACS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "Fin  Prod")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
//EXPORT TACHES TRAITEES SUIVANT
/************************************************************************************/ 
ecrituredataj2tri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Tri MGEFI" || f == "Tri MGEFI ")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+2")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 1421;
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

/************************************************************************************/ 
ecrituredataj2facM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+2")
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
 /************************************************************************************/ 
 ecrituredataj2devi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+2")
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

   /************************************************************************************/ 
   ecrituredataj2sales : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Salesforce MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+2")
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
 /************************************************************************************/ 
 ecrituredataj2flux : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Flux Noe MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+2")
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
 /************************************************************************************/ 
 ecrituredataj2rejet : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Rejet MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+2")
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
 /************************************************************************************/ 
 ecrituredataj2cotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+2")
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
 /************************************************************************************/ 
 ecrituredataj2cotite : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot ITE et MGAS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+2")
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
  /************************************************************************************/ 
 ecrituredataj2faclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Facture LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+2")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
 ecrituredataj2acs : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "ACS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+2")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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

//COLONNE J5
/************************************************************************************/ 
ecrituredataj5tri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Tri MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+5")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 1421;
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

/************************************************************************************/ 
ecrituredataj5facM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+5")
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
 /************************************************************************************/ 
 ecrituredataj5devi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+5")
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

   /************************************************************************************/ 
   ecrituredataj5sales : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Salesforce MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+5")
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
 /************************************************************************************/ 
 ecrituredataj5flux : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Flux Noe MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+5")
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
 /************************************************************************************/ 
 ecrituredataj5rejet : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Rejet MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+5")
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
 /************************************************************************************/ 
 ecrituredataj5cotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+5")
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
 /************************************************************************************/ 
 ecrituredataj5cotite : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot ITE et MGAS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+5")
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
   /************************************************************************************/ 
 ecrituredataj5faclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Facture LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+5")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
   /************************************************************************************/ 
 ecrituredataj5acs : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "ACS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES TRAITEES ' || cell.value == 'TACHES TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≤ J+5")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
//REMPLISSAGE STOCKS
/************************************************************************************/ 
ecrituredatastock16tri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Tri MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK traitable à 16h00")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastock16facM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK traitable à 16h00")
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
/************************************************************************************/ 
ecrituredatastock16devi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK traitable à 16h00")
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
/************************************************************************************/ 
ecrituredatastock16sales : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Salesforce MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK traitable à 16h00")
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
/************************************************************************************/ 
ecrituredatastock16flux : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Flux Noe MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK traitable à 16h00")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastock16rejet : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Rejet MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK traitable à 16h00")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredatastock16cotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK traitable à 16h00")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 1421;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastock16cotite : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot ITE et MGAS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK traitable à 16h00")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastock16faclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "Facture LAMIE")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'STOCKS')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "STOCK traitable à 16h00")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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
     /************************************************************************************/ 
     //ECRITURE ET REMPLISSAGE FICHIER
     ecrituredatastock16acs : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "ACS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'STOCKS')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "STOCK traitable à 16h00")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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
    
   /************************************************************************************/ 
ecrituredatastocktottri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Tri MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK total à 16h00")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastocktotfacM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK total à 16h00")
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
/************************************************************************************/ 
ecrituredatastocktotdevi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK total à 16h00")
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
/************************************************************************************/ 
ecrituredatastocktotsales : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Salesforce MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK total à 16h00")
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
/************************************************************************************/ 
ecrituredatastocktotflux : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Flux Noe MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK total à 16h00")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastocktotrejet : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Rejet MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK total à 16h00")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredatastocktotcotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK total à 16h00")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 1421;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastocktotcotite : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot ITE et MGAS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'STOCKS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "STOCK total à 16h00")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastocktotfaclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "Facture LAMIE")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'STOCKS')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "STOCK total à 16h00")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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
     /************************************************************************************/ 
     //ECRITURE ET REMPLISSAGE FICHIER
     ecrituredatastocktotacs : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "ACS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'STOCKS')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "STOCK total à 16h00")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredataetptri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "Tri MGEFI")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'ETP')
      {
        collonne = parseInt(colNumber);
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 1;
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

/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredataetpfacM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'ETP')
    {
      collonne = parseInt(colNumber);
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 7;
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

/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredataetpdevi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'ETP')
    {
      collonne = parseInt(colNumber);
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 4;
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
/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredataetpfaclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Facture LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'ETP')
    {
      collonne = parseInt(colNumber);
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredataetpcotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'ETP')
    {
      collonne = parseInt(colNumber);
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 2;
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

//REMPLISSAGE TACHES NON TRAITEES
/************************************************************************************/ 
ecrituredatastocknontjtri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Tri MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "J")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastocknontjfacM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "J")
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
/************************************************************************************/ 
ecrituredatastocknontjdevi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "J")
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
/************************************************************************************/ 
ecrituredatastocknontjsales : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Salesforce MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "J")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 1421;
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
/************************************************************************************/ 
ecrituredatastocknontjflux : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Flux Noe MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "J")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastocknontjrejet : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Rejet MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "J")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredatastocknontjcotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "J")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastocknontjcotite : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot ITE et MGAS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "J")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastocknontjfaclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "Facture LAMIE")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES NON TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "J")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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
     /************************************************************************************/ 
     //ECRITURE ET REMPLISSAGE FICHIER
     ecrituredatastocknontjacs : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "ACS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES NON TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "J")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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
    
/************************************************************************************/ 
ecrituredatastocknontj1tri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Tri MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "= J+1")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastocknontj1facM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "= J+1")
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
/************************************************************************************/ 
ecrituredatastocknontj1devi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "= J+1")
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
/************************************************************************************/ 
ecrituredatastocknontj1sales : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Salesforce MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "= J+1")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 1421;
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
/************************************************************************************/ 
ecrituredatastocknontj1flux : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Flux Noe MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "= J+1")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastocknontj1rejet : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Rejet MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "= J+1")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredatastocknontj1cotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "= J+1")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastocknontj1cotite : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot ITE et MGAS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "= J+1")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastocknontj1faclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "Facture LAMIE")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES NON TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "= J+1")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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
     /************************************************************************************/ 
     //ECRITURE ET REMPLISSAGE FICHIER
     ecrituredatastocknontj1acs : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "ACS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES NON TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "= J+1")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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

/************************************************************************************/ 
ecrituredatastocknontj2tri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Tri MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+2")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastocknontj2facM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+2")
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
/************************************************************************************/ 
ecrituredatastocknontj2devi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+2")
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
/************************************************************************************/ 
ecrituredatastocknontj2sales : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Salesforce MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+2")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 1421;
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
/************************************************************************************/ 
ecrituredatastocknontj2flux : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Flux Noe MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+2")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastocknontj2rejet : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Rejet MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+2")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredatastocknontj2cotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+2")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastocknontj2cotite : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot ITE et MGAS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+2")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastocknontj2faclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "Facture LAMIE")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES NON TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "≥ J+2")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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
     /************************************************************************************/ 
     //ECRITURE ET REMPLISSAGE FICHIER
     ecrituredatastocknontj2acs : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "ACS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES NON TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "≥ J+2")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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

/************************************************************************************/ 
ecrituredatastocknontj5tri : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Tri MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+5")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastocknontj5facM : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Factures MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+5")
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
/************************************************************************************/ 
ecrituredatastocknontj5devi : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Devis MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+5")
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
/************************************************************************************/ 
ecrituredatastocknontj5sales : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Salesforce MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+5")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 1421;
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
/************************************************************************************/ 
ecrituredatastocknontj5flux : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Flux Noe MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+5")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
ecrituredatastocknontj5rejet : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Rejet MGEFI")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+5")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
/************************************************************************************/ 
//ECRITURE ET REMPLISSAGE FICHIER
ecrituredatastocknontj5cotlamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot LAMIE")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+5")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastocknontj5cotite : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  const Excel = require('exceljs');
  const cmd=require('node-cmd');
  const newWorkbook = new Excel.Workbook();
  
  try{
  
          
    await newWorkbook.xlsx.readFile(path_reporting);
  const newworksheet = newWorkbook.getWorksheet(1);
  var colonneDate = newworksheet.getColumn('A');
  var ligneDate1;
  var ligneDate;
  colonneDate.eachCell(function(cell, rowNumber) {
    var dateExcel = Retour.convertDate(cell.text);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(4).value;
      // console.log(f);
      if(f == "Contrat Cot ITE et MGAS")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  // var iniValue = Retour.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
    
  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'TACHES NON TRAITEES')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      // var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(f == "≥ J+5")
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = 0;
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
  /************************************************************************************/ 
  //ECRITURE ET REMPLISSAGE FICHIER
  ecrituredatastocknontj5faclamie : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "Facture LAMIE")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES NON TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "≥ J+5")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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
     /************************************************************************************/ 
     //ECRITURE ET REMPLISSAGE FICHIER
     ecrituredatastocknontj5acs : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
            
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = Retour.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        // console.log(f);
        if(f == "ACS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    // var iniValue = Retour.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
      
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'TACHES NON TRAITEES')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        // var getko_ini = man.getCell(colDate2).address;
        // console.log(getko_ini);
        if(f == "≥ J+5")
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = 0;
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

/************************************************************************/
deleteFromChemin : function (nomBase,callback) {
  var sql = "delete from "+nomBase+" ";
  Garantie.getDatastore().sendNativeQuery(sql, function(err, res){
    if (err) { return callback(err); }
    return callback(null, true);
    });
},
/************************************************************************/
deleteFromCheminfacture : function (nomBase,callback) {
  var sql = "delete from "+nomBase+" ";
  Garantie.getDatastore().sendNativeQuery(sql, function(err, res){
    if (err) { return callback(err); }
    return callback(null, true);
    });
},
/************************************************************************/
deleteFromChemindevis : function (nomBase,callback) {
  var sql = "delete from "+nomBase+" ";
  Garantie.getDatastore().sendNativeQuery(sql, function(err, res){
    if (err) { return callback(err); }
    return callback(null, true);
    });
},
 /*******************************************************/
  //INSERTION DU CHEMIN DANS LA BASE DE DONNEE
  importcheminhtp: function (table_1,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,nomBase,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    // var a = table[0]+date+table2[nb];
    var a = table_1[0]+date+table2[nb];
    console.log('*****************************');
    console.log('chemin de a : '+a);
    //var a ='\\\\10.128.1.2\\almerys-out\\Retour_Easytech_20210512\\TRAITEMENT_RETOUR_OTD_N2\\' ;
    var b = option[nb];
    //var b = 'OTD_ALMERYS SATD';
    //var c = 'vrai';
    //console.log(a);
    var nomTable = nomtable;
    var numLigne= numligne;
    var numFeuille = numfeuille;
    var nomColonne = nomcolonne;
    var c = Garantie.existenceFichier(a);
    console.log('ccccccccccccccccccccccc: '+c);
    if(c=='vrai')
    {
      fs.readdir(a, (err, files) => {
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              console.log('**********************************************************************');
              console.log(b);
              console.log(file);
              console.log(regex.test(file));
              console.log('***************************************************************************');
              if(regex.test(file))
              {
                 //re = a+'\\'+file;
                 re = a+''+file;
                 console.log(re);

                 var sql = "insert into "+nomBase+" (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable[nb]+"','"+numLigne[nb]+"','"+numFeuille[nb]+"','"+nomColonne[nb]+"','"+colonnecible2[nb]+"','"+colonnecible3[nb]+"') ";
                 console.log(sql);
                 Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                  if (err) { 
                    console.log("Une erreur ve? import 1");
                    //return callback(err);
                   }
                  else
                  {
                    console.log("eto le requete alefany io : "+sql);
                    return callback(null, true);
                  };
                   
                });
             }
              else
              {
               var sql = "insert into chemintsisy (typologiedelademande) values ('"+re+"') ";
               Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log(sql);
                  return callback(null, true);
                };
                 
            });
              }
             
             
          });
          
         
        });
    }
    else
    {
      var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
      Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log("Une erreur ve? import 1");
          //return callback(err);
         }
        else
        {
          console.log('eto njay iz le ts mety an : '+sql);
          return callback(null, true);
        };
         
    });
    }   
  },
/*******************************************************/
  //INSERTION DU CHEMIN DANS LA BASE DE DONNEE
  importcheminhtpfacture: function (table_1,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,nomBase,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    // var a = table[0]+date+table2[nb];
    var a = table_1[0]+date+table2[nb];
    console.log('*****************************');
    console.log('chemin de a : '+a);
    //var a ='\\\\10.128.1.2\\almerys-out\\Retour_Easytech_20210512\\TRAITEMENT_RETOUR_OTD_N2\\' ;
    var b = option[nb];
    //var b = 'OTD_ALMERYS SATD';
    //var c = 'vrai';
    //console.log(a);
    var nomTable = nomtable;
    var numLigne= numligne;
    var numFeuille = numfeuille;
    var nomColonne = nomcolonne;
    var c = Garantie.existenceFichier(a);
    console.log('ccccccccccccccccccccccc: '+c);
    if(c=='vrai')
    {
      fs.readdir(a, (err, files) => {
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              console.log('**********************************************************************');
              console.log(b);
              console.log(file);
              console.log(regex.test(file));
              console.log('***************************************************************************');
              if(regex.test(file))
              {
                 //re = a+'\\'+file;
                 re = a+''+file;
                 console.log(re);

                 var sql = "insert into "+nomBase+" (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable[nb]+"','"+numLigne[nb]+"','"+numFeuille[nb]+"','"+nomColonne[nb]+"','"+colonnecible2[nb]+"','"+colonnecible3[nb]+"') ";
                 console.log(sql);
                 Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                  if (err) { 
                    console.log("Une erreur ve? import 1");
                    //return callback(err);
                   }
                  else
                  {
                    console.log("eto le requete alefany io : "+sql);
                    return callback(null, true);
                  };
                   
                });
             }
              else
              {
               var sql = "insert into chemintsisy (typologiedelademande) values ('"+re+"') ";
               Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log(sql);
                  return callback(null, true);
                };
                 
            });
              }
             
             
          });
          
         
        });
    }
    else
    {
      var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
      Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log("Une erreur ve? import 1");
          //return callback(err);
         }
        else
        {
          console.log('eto njay iz le ts mety an : '+sql);
          return callback(null, true);
        };
         
    });
    }   
  },
  /*******************************************************/
  //INSERTION DU CHEMIN DANS LA BASE DE DONNEE
  importcheminhtpdevis: function (table_1,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,nomBase,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    // var a = table[0]+date+table2[nb];
    var a = table_1[0]+date+table2[nb];
    console.log('*****************************');
    console.log('chemin de a : '+a);
    //var a ='\\\\10.128.1.2\\almerys-out\\Retour_Easytech_20210512\\TRAITEMENT_RETOUR_OTD_N2\\' ;
    var b = option[nb];
    //var b = 'OTD_ALMERYS SATD';
    //var c = 'vrai';
    //console.log(a);
    var nomTable = nomtable;
    var numLigne= numligne;
    var numFeuille = numfeuille;
    var nomColonne = nomcolonne;
    var c = Garantie.existenceFichier(a);
    console.log('ccccccccccccccccccccccc: '+c);
    if(c=='vrai')
    {
      fs.readdir(a, (err, files) => {
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              console.log('**********************************************************************');
              console.log(b);
              console.log(file);
              console.log(regex.test(file));
              console.log('***************************************************************************');
              if(regex.test(file))
              {
                 //re = a+'\\'+file;
                 re = a+''+file;
                 console.log(re);

                 var sql = "insert into "+nomBase+" (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable[nb]+"','"+numLigne[nb]+"','"+numFeuille[nb]+"','"+nomColonne[nb]+"','"+colonnecible2[nb]+"','"+colonnecible3[nb]+"') ";
                 console.log(sql);
                 Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                  if (err) { 
                    console.log("Une erreur ve? import 1");
                    //return callback(err);
                   }
                  else
                  {
                    console.log("eto le requete alefany io : "+sql);
                    return callback(null, true);
                  };
                   
                });
             }
              else
              {
               var sql = "insert into chemintsisy (typologiedelademande) values ('"+re+"') ";
               Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log(sql);
                  return callback(null, true);
                };
                 
            });
              }
             
             
          });
          
         
        });
    }
    else
    {
      var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
      Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log("Une erreur ve? import 1");
          //return callback(err);
         }
        else
        {
          console.log('eto njay iz le ts mety an : '+sql);
          return callback(null, true);
        };
         
    });
    }   
  },
  /*******************************************************/
  //INSERTION DU CHEMIN DANS LA BASE DE DONNEE
  importcheminhtpligne: function (table_1,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,nomBase,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    // var a = table[0]+date+table2[nb];
    var a = table_1[0]+date+table2[nb];
    console.log('*****************************');
    console.log('chemin de a : '+a);
    //var a ='\\\\10.128.1.2\\almerys-out\\Retour_Easytech_20210512\\TRAITEMENT_RETOUR_OTD_N2\\' ;
    var b = option[nb];
    //var b = 'OTD_ALMERYS SATD';
    //var c = 'vrai';
    //console.log(a);
    var nomTable = nomtable;
    var numLigne= numligne;
    var numFeuille = numfeuille;
    var nomColonne = nomcolonne;
    var c = Garantie.existenceFichier(a);
    console.log('ccccccccccccccccccccccc: '+c);
    if(c=='vrai')
    {
      fs.readdir(a, (err, files) => {
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              console.log('**********************************************************************');
              console.log(b);
              console.log(file);
              console.log(regex.test(file));
              console.log('***************************************************************************');
              if(regex.test(file))
              {
                 //re = a+'\\'+file;
                 re = a+''+file;
                 console.log(re);

                 var sql = "insert into "+nomBase+" (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable[nb]+"','"+numLigne[nb]+"','"+numFeuille[nb]+"','"+nomColonne[nb]+"','"+colonnecible2[nb]+"','"+colonnecible3[nb]+"') ";
                 console.log(sql);
                 Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                  if (err) { 
                    console.log("Une erreur ve? import 1");
                    //return callback(err);
                   }
                  else
                  {
                    console.log("eto le requete alefany io : "+sql);
                    return callback(null, true);
                  };
                   
                });
             }
              else
              {
               var sql = "insert into chemintsisy (typologiedelademande) values ('"+re+"') ";
               Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log(sql);
                  return callback(null, true);
                };
                 
            });
              }
             
             
          });
          
         
        });
    }
    else
    {
      var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
      Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log("Une erreur ve? import 1");
          //return callback(err);
         }
        else
        {
          console.log('eto njay iz le ts mety an : '+sql);
          return callback(null, true);
        };
         
    });
    }   
  },
   /*******************************************************/
  //INSERTION DU CHEMIN DANS LA BASE DE DONNEE
  importcheminhtpsales: function (table_1,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,nomBase,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    // var a = table[0]+date+table2[nb];
    var a = table_1[0]+table2[nb];
    console.log('*****************************');
    console.log('chemin de a : '+a);
    //var a ='\\\\10.128.1.2\\almerys-out\\Retour_Easytech_20210512\\TRAITEMENT_RETOUR_OTD_N2\\' ;
    var b = option[nb];
    //var b = 'OTD_ALMERYS SATD';
    //var c = 'vrai';
    //console.log(a);
    var nomTable = nomtable;
    var numLigne= numligne;
    var numFeuille = numfeuille;
    var nomColonne = nomcolonne;
    var c = Garantie.existenceFichier(a);
    console.log('ccccccccccccccccccccccc: '+c);
    if(c=='vrai')
    {
      fs.readdir(a, (err, files) => {
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              console.log('**********************************************************************');
              console.log(b);
              console.log(file);
              console.log(regex.test(file));
              console.log('***************************************************************************');
              if(regex.test(file))
              {
                 //re = a+'\\'+file;
                 re = a+''+file;
                 console.log(re);

                 var sql = "insert into "+nomBase+" (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable[nb]+"','"+numLigne[nb]+"','"+numFeuille[nb]+"','"+nomColonne[nb]+"','"+colonnecible2[nb]+"','"+colonnecible3[nb]+"') ";
                 console.log(sql);
                 Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                  if (err) { 
                    console.log("Une erreur ve? import 1");
                    //return callback(err);
                   }
                  else
                  {
                    console.log("eto le requete alefany io : "+sql);
                    return callback(null, true);
                  };
                   
                });
             }
              else
              {
               var sql = "insert into chemintsisy (typologiedelademande) values ('"+re+"') ";
               Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log(sql);
                  return callback(null, true);
                };
                 
            });
              }
             
             
          });
          
         
        });
    }
    else
    {
      var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
      Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log("Une erreur ve? import 1");
          //return callback(err);
         }
        else
        {
          console.log('eto njay iz le ts mety an : '+sql);
          return callback(null, true);
        };
         
    });
    }   
  },
   /*******************************************************/
  //INSERTION DU CHEMIN DANS LA BASE DE DONNEE
  importcheminhtpstockfacdevis: function (table_1,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,nomBase,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    // var a = table[0]+date+table2[nb];
    var a = table_1[0]+date;
    console.log('*****************************');
    console.log('chemin de a : '+a);
    //var a ='\\\\10.128.1.2\\almerys-out\\Retour_Easytech_20210512\\TRAITEMENT_RETOUR_OTD_N2\\' ;
    var b = option[nb];
    //var b = 'OTD_ALMERYS SATD';
    //var c = 'vrai';
    //console.log(a);
    var nomTable = nomtable;
    var numLigne= numligne;
    var numFeuille = numfeuille;
    var nomColonne = nomcolonne;
    var c = Garantie.existenceFichier(a);
    console.log('ccccccccccccccccccccccc: '+c);
    if(c=='vrai')
    {
      fs.readdir(a, (err, files) => {
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              console.log('**********************************************************************');
              console.log(b);
              console.log(file);
              console.log(regex.test(file));
              console.log('***************************************************************************');
              if(regex.test(file))
              {
                 //re = a+'\\'+file;
                 re = a+file;
                 console.log(re);

                 var sql = "insert into "+nomBase+" (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable[nb]+"','"+numLigne[nb]+"','"+numFeuille[nb]+"','"+nomColonne[nb]+"','"+colonnecible2[nb]+"','"+colonnecible3[nb]+"') ";
                 console.log(sql);
                 Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                  if (err) { 
                    console.log("Une erreur ve? import 1");
                    //return callback(err);
                   }
                  else
                  {
                    console.log("eto le requete alefany io : "+sql);
                    return callback(null, true);
                  };
                   
                });
             }
              else
              {
               var sql = "insert into chemintsisy (typologiedelademande) values ('"+re+"') ";
               Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log(sql);
                  return callback(null, true);
                };
                 
            });
              }
             
             
          });
          
         
        });
    }
    else
    {
      var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
      Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log("Une erreur ve? import 1");
          //return callback(err);
         }
        else
        {
          console.log('eto njay iz le ts mety an : '+sql);
          return callback(null, true);
        };
         
    });
    }   
  },
  /*********************************************************************/
  deleteReportingHtp : function (table,nb,callback) {
    var sql = "delete from "+table[nb]+" ";
    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return console.log(err); }
      return callback(null, true);
      });
  },
  deleteHtp : function (table,nb,callback) {
    var j;
    var i = parseInt(j);
    for(i=0;i<nb;i++)
    {
      Engagementhtp.deleteReportingHtp(table,i,callback);
    };
  },
  delete : function (table,nb,callback) {
    var nbr = parseInt(nb);
    var sql = "delete from "+table[nbr]+" ";
    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log("Une erreur supprooo?");
        console.log(err);
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
      });
  },

/****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE
importengagementhtptri : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtptri( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
  /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtptri:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col ;
          var col1;
          var col2;
          console.log('Nombre de colonne' + range.e.c);
          console.log('Nombre de ligne' + range.e.r);
          console.log(numeroligne + 'numlign');
          console.log(numerofeuille + 'numfeuille');
          console.log(cellule2[1] + 'c2');
          console.log(cellule[1] + 'c1');
          console.log(dernierl[1] + 'c3');
          console.log(table[1] + 'table');
          for(var ra=0;ra<=range.e.c;ra++)
          {
            var address_of_cell = {c:ra, r:numeroligne};
            // console.log(address_of_cell);//c:5 r:0
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            // console.log(desired_cell);
            var desired_value = (desired_cell ? desired_cell.v : undefined);
            // console.log(desired_value);//No Facture
            if(desired_value==dernierl[0])
            {
              col=ra;
            }
          };
          
          console.log('colonne cible : ' +col);
          if(col!=undefined)
          {
            var nbr = 0;
            for(var a=0;a<=range.e.r;a++)
              {
                var address_of_cell = {c:col, r:a};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value1 = (desired_cell ? desired_cell.v : undefined);
                var bi = 'MGEFI HTP - Tri Courrier  ';
                const regex = new RegExp(bi+'*');
                if(regex.test(desired_value1))
                {
                  nbr=nbr + 1;
                };
              };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
              console.log("nombreeeeebr__tri__"+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha");
        }
   

  },

/****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpdevis : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpdevis( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpdevis:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col ;
          var col1;
          var col2;
          console.log('Nombre de colonne' + range.e.c);
          console.log('Nombre de ligne' + range.e.r);
          console.log(numeroligne + 'numlign');
          console.log(numerofeuille + 'numfeuille');
          console.log(cellule2[1] + 'c2');
          console.log(cellule[1] + 'c1');
          console.log(dernierl[1] + 'c3');
          console.log(table[1] + 'table');
          for(var ra=0;ra<=range.e.c;ra++)
          {
            var address_of_cell = {c:ra, r:numeroligne};
            // console.log(address_of_cell);//c:5 r:0
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            // console.log(desired_cell);
            var desired_value = (desired_cell ? desired_cell.v : undefined);
            // console.log(desired_value);//No Facture
            if(desired_value==dernierl[0])
            {
              col=ra;
            }
          };
          
          console.log('colonne cible : ' +col);
          if(col!=undefined)
          {
            var nbr = 0;
            for(var a=0;a<=range.e.r;a++)
              {
                var address_of_cell = {c:col, r:a};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value1 = (desired_cell ? desired_cell.v : undefined);
                var bi = 'MGEFI HTP - Saisir Devis ';
                const regex = new RegExp(bi+'*');
                if(regex.test(desired_value1))
                {
                  nbr=nbr + 1;
                };
              };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
              console.log("nombreeeeebr_facture_MGEFI__"+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha");
        }
   

  },
/****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE FACTURE
importengagementhtpfacture : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpfacture( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpfacture:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col ;
          var col1;
          var col2;
          console.log('Nombre de colonne' + range.e.c);
          console.log('Nombre de ligne' + range.e.r);
          console.log(numeroligne + 'numlign');
          console.log(numerofeuille + 'numfeuille');
          console.log(cellule2[1] + 'c2');
          console.log(cellule[1] + 'c1');
          console.log(dernierl[1] + 'c3');
          console.log(table[1] + 'table');
          for(var ra=0;ra<=range.e.c;ra++)
          {
            var address_of_cell = {c:ra, r:numeroligne};
            // console.log(address_of_cell);//c:5 r:0
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            // console.log(desired_cell);
            var desired_value = (desired_cell ? desired_cell.v : undefined);
            // console.log(desired_value);//No Facture
            if(desired_value==dernierl[0])
            {
              col=ra;
            }
          };
          
          console.log('colonne cible : ' +col);
          if(col!=undefined)
          {
            var nbr = 0;
            for(var a=0;a<=range.e.r;a++)
              {
                var address_of_cell = {c:col, r:a};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value1 = (desired_cell ? desired_cell.v : undefined);
                var bi = 'MGEFI HTP - Saisir Facture ';
                const regex = new RegExp(bi+'*');
                if(regex.test(desired_value1))
                {
                  nbr=nbr + 1;
                };
              };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
              console.log("nombreeeeebr_facture_MGEFI__"+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha");
        }
   

  },

/****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE FACTURE
importengagementhtpligne : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpligne( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpligne:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          var nbr = 0;
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col=0;
          var nbe = parseInt(nb);
          if(col!=undefined)
          {
            var debutligne = numeroligne + 1;
            for(var a=debutligne;a<=range.e.r;a++)
              {
                var address_of_cell = {c:col, r:a};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value1 = (desired_cell ? desired_cell.v : undefined);
                if(desired_value1!=undefined)
                {
                  nbr=nbr + 1;
                }
              }; 
          }
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
              console.log("nombreeeeebr"+ nbr);
              return tab;
        }
        catch
        {
          console.log("erreur absolu haaha");
        }       
   

  },
/****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpdevisj2 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpdevisj2( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpdevisj2:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col ;
          var col1;
          var col2;
          console.log('Nombre de colonne' + range.e.c);
          console.log('Nombre de ligne' + range.e.r);
          console.log(numeroligne + 'numlign');
          console.log(numerofeuille + 'numfeuille');
          console.log(cellule2[1] + 'c2');
          console.log(cellule[1] + 'c1');
          console.log(dernierl[1] + 'c3');
          console.log(table[1] + 'table');
          for(var ra=0;ra<=range.e.c;ra++)
          {
            var address_of_cell = {c:ra, r:numeroligne};
            // console.log(address_of_cell);//c:5 r:0
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            // console.log(desired_cell);
            var desired_value = (desired_cell ? desired_cell.v : undefined);
            // console.log(desired_value);//No Facture
            if(desired_value==dernierl[0])
            {
              col=ra;
            }
          };
          
          console.log('colonne cible : ' +col);
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Devis ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:6, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j2 = conv - 2;
              if(regex.test(desired_value) && desired_value1<=j2)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
              console.log("nombreeeeebr_devis_j2__"+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   

  },
/****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpdevisj5 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpdevisj5( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpdevisj5:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col ;
          var col1;
          var col2;
          console.log('Nombre de colonne' + range.e.c);
          console.log('Nombre de ligne' + range.e.r);
          console.log(numeroligne + 'numlign');
          console.log(numerofeuille + 'numfeuille');
          console.log(cellule2[1] + 'c2');
          console.log(cellule[1] + 'c1');
          console.log(dernierl[1] + 'c3');
          console.log(table[1] + 'table');
          for(var ra=0;ra<=range.e.c;ra++)
          {
            var address_of_cell = {c:ra, r:numeroligne};
            // console.log(address_of_cell);//c:5 r:0
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            // console.log(desired_cell);
            var desired_value = (desired_cell ? desired_cell.v : undefined);
            // console.log(desired_value);//No Facture
            if(desired_value==dernierl[0])
            {
              col=ra;
            }
          };
          
          console.log('colonne cible : ' +col);
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Devis ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:6, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j5 = conv - 5;
              if(regex.test(desired_value) && desired_value1<=j5)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
              console.log("nombreeeeebr_devis_j5"+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   

  },

/****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpfacmgj2 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpfacmgj2( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpfacmgj2:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col ;
          var col1;
          var col2;
          console.log('Nombre de colonne' + range.e.c);
          console.log('Nombre de ligne' + range.e.r);
          console.log(numeroligne + 'numlign');
          console.log(numerofeuille + 'numfeuille');
          console.log(cellule2[1] + 'c2');
          console.log(cellule[1] + 'c1');
          console.log(dernierl[1] + 'c3');
          console.log(table[1] + 'table');
          for(var ra=0;ra<=range.e.c;ra++)
          {
            var address_of_cell = {c:ra, r:numeroligne};
            // console.log(address_of_cell);//c:5 r:0
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            // console.log(desired_cell);
            var desired_value = (desired_cell ? desired_cell.v : undefined);
            // console.log(desired_value);//No Facture
            if(desired_value==dernierl[0])
            {
              col=ra;
            }
          };
          
          console.log('colonne cible : ' +col);
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Facture ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:6, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j2 = conv - 2;
              if(regex.test(desired_value) && desired_value1<=j2)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
          console.log("nombreeeeebr_Facture_MGEFI_J2"+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   

  },

/****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpfacmgj5 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpfacmgj5( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpfacmgj5:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col ;
          var col1;
          var col2;
          console.log('Nombre de colonne' + range.e.c);
          console.log('Nombre de ligne' + range.e.r);
          console.log(numeroligne + 'numlign');
          console.log(numerofeuille + 'numfeuille');
          console.log(cellule2[1] + 'c2');
          console.log(cellule[1] + 'c1');
          console.log(dernierl[1] + 'c3');
          console.log(table[1] + 'table');
          for(var ra=0;ra<=range.e.c;ra++)
          {
            var address_of_cell = {c:ra, r:numeroligne};
            // console.log(address_of_cell);//c:5 r:0
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            // console.log(desired_cell);
            var desired_value = (desired_cell ? desired_cell.v : undefined);
            // console.log(desired_value);//No Facture
            if(desired_value==dernierl[0])
            {
              col=ra;
            }
          };
          
          console.log('colonne cible : ' +col);
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Facture ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:6, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j5 = conv - 5;
              if(regex.test(desired_value) && desired_value1<=j5)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
              console.log("nombreeeeebr_Facture_MGEFI_J5"+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   

  },
/****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpcotlamiej2 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpcotlamiej2( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpcotlamiej2:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col;
          var col1;
          var col2;
          console.log('Nombre de colonne' + range.e.c);
          console.log('Nombre de ligne' + range.e.r);
          console.log(numeroligne + 'numlign');
          console.log(numerofeuille + 'numfeuille');
          console.log(cellule2[1] + 'c2');
          console.log(cellule[1] + 'c1');
          console.log(dernierl[1] + 'c3');
          console.log(table[1] + 'table');
          for(var ra=0;ra<=range.e.c;ra++)
          {
            var address_of_cell = {c:ra, r:numeroligne};
            // console.log(address_of_cell);//c:5 r:0
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            // console.log(desired_cell);
            var desired_value = (desired_cell ? desired_cell.v : undefined);
            // console.log(desired_value);//No Facture
            if(desired_value==dernierl[0])
            {
              col=ra;
            }
          };
          
          console.log('colonne cible : ' +col);

          if(col!=undefined)
          {
            var nbr = 0;
            var debutligne = numeroligne + 1;
            for(var a=debutligne;a<=range.e.r;a++)
              {
               
                var address_of_cell1 = {c:col, r:a};
                var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
                var desired_cell1 = sheet[cell_ref1];
                var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
          
                var test = Engagementhtp.convertionexceldate(desired_value1);
                var y = test.getFullYear();
                var m = test.getMonth()+1;
                var d = test.getDate();

                if (d < 10) {
                  d = '0' + d;
                }
                if (m < 10) {
                  m = '0' + m;
                }

                var datetime = y+''+m+''+d;
                // console.log(datetime);

                var conv = parseInt(date);
                var j2 = conv - 2;
                
                  if(datetime<=j2)
                 {
                    nbr=nbr + 1;  
                 }
              }; 
          }
   
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
              console.log("nombreeeeebr_cot_LAMIE_J2"+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   

  },
  /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpcotlamiej5 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpcotlamiej5( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpcotlamiej5:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col;
          var col1;
          var col2;
          console.log('Nombre de colonne' + range.e.c);
          console.log('Nombre de ligne' + range.e.r);
          console.log(numeroligne + 'numlign');
          console.log(numerofeuille + 'numfeuille');
          console.log(cellule2[1] + 'c2');
          console.log(cellule[1] + 'c1');
          console.log(dernierl[1] + 'c3');
          console.log(table[1] + 'table');
          for(var ra=0;ra<=range.e.c;ra++)
          {
            var address_of_cell = {c:ra, r:numeroligne};
            // console.log(address_of_cell);//c:5 r:0
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            // console.log(desired_cell);
            var desired_value = (desired_cell ? desired_cell.v : undefined);
            // console.log(desired_value);//No Facture
            if(desired_value==dernierl[0])
            {
              col=ra;
            }
          };
          
          console.log('colonne cible : ' +col);

          if(col!=undefined)
          {
            var nbr = 0;
            var debutligne = numeroligne + 1;
            for(var a=debutligne;a<=range.e.r;a++)
              {
               
                var address_of_cell1 = {c:col, r:a};
                var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
                var desired_cell1 = sheet[cell_ref1];
                var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
          
                var test = Engagementhtp.convertionexceldate(desired_value1);
                var y = test.getFullYear();
                var m = test.getMonth()+1;
                var d = test.getDate();
                
                if (d < 10) {
                  d = '0' + d;
                }
                if (m < 10) {
                  m = '0' + m;
                }

                var datetime = y+''+m+''+d;
                // console.log(datetime);

                var conv = parseInt(date);
                var j5 = conv - 5;
                
                  if(datetime<=j5)
                 {
                    nbr=nbr + 1;  
                 }
              }; 
          }
   
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
              console.log("nombreeeeebr_cot_LAMIE_J5"+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   

  },
  /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpsales : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpsales( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpsales:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);

        try{
          console.log('miditra ato am try v iz?');
          var data;
          var donnee;
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          if(sheet){
            data = XLSX.utils.sheet_to_json(sheet, {raw:true, dateNf: 'dd-mm-yyyy', header:[0, 1, 2, 3], cellDate:true});
            console.log(data);

            

            
            for(var i=1; i<data.length; i++){
              var test = Engagementhtp.convertionexceldate(data[i][0]);
              var y = test.getFullYear();
              var m = test.getMonth()+1;
              var d = test.getDate();

              if (d < 10) {
                d = '0' + d;
              }
              if (m < 10) {
                m = '0' + m;
              }

              var datetime = d+'/'+m+'/'+y;

              if(datetime == dateexport){
                donnee = data[i];
              }
            }
          }
          else
          {
            console.log('Colonne non trouvé');
          }
        

          var tab = [donnee[2]];
              console.log("valeur_SALESFORCE: "+ tab);
              return tab; 



        }
      
        catch
        {
          console.log("erreur absolu haaha salesssssss");
        }
   

  },
 /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpsalesstock : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpsalesstock( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpsalesstock:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);

        try{
          console.log('miditra ato am try v iz?');
          var data;
          var donnee;
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          if(sheet){
            data = XLSX.utils.sheet_to_json(sheet, {raw:true, dateNf: 'dd-mm-yyyy', header:[0, 1, 2, 3], cellDate:true});
            console.log(data);

            

            
            for(var i=1; i<data.length; i++){
              var test = Engagementhtp.convertionexceldate(data[i][0]);
              var y = test.getFullYear();
              var m = test.getMonth()+1;
              var d = test.getDate();

              if (d < 10) {
                d = '0' + d;
              }
              if (m < 10) {
                m = '0' + m;
              }

              var datetime = d+'/'+m+'/'+y;

              if(datetime == dateexport){
                donnee = data[i];
              }
            }
          }
          else
          {
            console.log('Colonne non trouvé');
          }
        

          var tab = [donnee[3]];
              console.log("valeur_SALESFORCE: "+ tab);
              return tab; 



        }
      
        catch
        {
          console.log("erreur absolu haaha salesssssss");
        }
   

  },
/****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpfacmgstock : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('tonga ato v lets ben aaaaaaaaaa');
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpfacmgstock( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpfacmgstock:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback){
    XLSX = require('xlsx');
    
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = parseInt(numligne[nb]);
    console.log('lign ato am lecture et insertion : ' +numeroligne);
    try{
      console.log('miditra ato am try v iz?');
      console.log(numerofeuille);
      const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var col = 2;

      if(col!=undefined)
      {
        var nbr = 0;
        for(var a=0;a<=range.e.r;a++)
          {
            var address_of_cell = {c:col, r:a};
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            var desired_value1 = (desired_cell ? desired_cell.v : undefined);
            var bi = 'MGEFI HTP - Saisir Facture ';
            const regex = new RegExp(bi+'*');
            if(regex.test(desired_value1))
            {
              nbr=nbr + 1;
            };
          };
      }


      
      else
      {
        console.log('Colonne non trouvé');
      }
      var tab = [nbr];
          console.log("nombreeeeebr_FACTURE_MGEFI_STOCK "+ nbr);
          return tab; 

      
    }
  
    catch
    {
      console.log("erreur absolu haaha");
    }
   

  },
  /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpdevistock : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpdevistock( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpdevistock:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,dateexport,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col = 2;
          if(col!=undefined)
          {
            var nbr = 0;
            for(var a=0;a<=range.e.r;a++)
              {
                var address_of_cell = {c:col, r:a};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value1 = (desired_cell ? desired_cell.v : undefined);
                var bi = 'MGEFI HTP - Saisir Devis ';
                const regex = new RegExp(bi+'*');
                if(regex.test(desired_value1))
                {
                  nbr=nbr + 1;
                };
              };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
              console.log("nombreeeeebr_MGEFI_DEVIS_STOCK "+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha");
        }
   

  },

 /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpfacmgtnontj : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpfacmgtnontj( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpfacmgtnontj:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col = 2;
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Facture ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:5, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j = conv;
              if(regex.test(desired_value) && desired_value1==j)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
          console.log("nombreeeeebr_NON_TRAITEES_FACTURE_J "+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   
    

  },

   /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpfacmgtnontj1 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpfacmgtnontj1( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpfacmgtnontj1:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col = 2;
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Facture ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:5, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j1 = conv - 1;
              if(regex.test(desired_value) && desired_value1==j1)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
          console.log("nombreeeeebr_NON_TRAITEES_FACTURE_J1 "+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   
    

  },
   /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpfacmgtnontj2 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpfacmgtnontj2( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpfacmgtnontj2:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col = 2;
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Facture ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:5, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j2 = conv - 2;
              if(regex.test(desired_value) && desired_value1>=j2)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
          console.log("nombreeeeebr_NON_TRAITEES_FACTURE_J2 "+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   
    

  },


   /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpfacmgtnontj5 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpfacmgtnontj5( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpfacmgtnontj5:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col = 2;
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Facture ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:5, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j5 = conv - 5;
              if(regex.test(desired_value) && desired_value1>=j5)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
          console.log("nombreeeeebr_NON_TRAITEES_FACTURE_J5 "+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   
    

  },
   /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpdevitnontj : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpdevitnontj( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpdevitnontj:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col = 2;
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Devis ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:5, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j = conv;
              if(regex.test(desired_value) && desired_value1==j)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
          console.log("nombreeeeebr_NON_TRAITEES_DEVIS_J "+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   
    

  },
   /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpdevitnontj1 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpdevitnontj1( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpdevitnontj1:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col = 2;
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Devis ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:5, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j1 = conv - 1;
              if(regex.test(desired_value) && desired_value1==j1)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
          console.log("nombreeeeebr_NON_TRAITEES_DEVIS_J1 "+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   
    

  },

   /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpdevitnontj2 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpdevitnontj2( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpdevitnontj2:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col = 2;
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Devis ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:5, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j2 = conv - 2;
              if(regex.test(desired_value) && desired_value1>=j2)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
          console.log("nombreeeeebr_NON_TRAITEES_DEVIS_J2 "+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   
    

  },

   /****************************************************************************************/
//IMPORT DES DONNES EXCELS VERS LA BASE DE DONNEE DEVIS
importengagementhtpdevitnontj5 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback) {
  console.log('****************');
  console.log(nb);
  console.log(trameflux[nb]);
  console.log('****************');
  if(trameflux[nb]==undefined)
  {
    console.log('trame undefined');
    var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
    });
  }
  else{

    var tab = [];
    tab = Engagementhtp.lectureEtInsertionengagementhtpdevitnontj5( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
   

    Engagementhtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve ok?");
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          });
  };

},
 /****************************************************************************/
  //LECTURE DU CHEMIN ET INSERTION DANS LA BASE
  lectureEtInsertionengagementhtpdevitnontj5:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,date,callback){
    XLSX = require('xlsx');
    
        var workbook = XLSX.readFile(trameflux[nb]);
        var numerofeuille = feuil[nb];
        var numeroligne = parseInt(numligne[nb]);
        console.log('lign ato am lecture et insertion : ' +numeroligne);
        try{
          console.log('miditra ato am try v iz?');
          console.log(numerofeuille);
          const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
          var range = XLSX.utils.decode_range(sheet['!ref']);
          var col = 2;
          if(col!=undefined)
          {
            var nbr = 0;
            for(var ra=2;ra<=range.e.r;ra++)
            {
             // console.log(ra);
              var address_of_cell = {c:col, r:ra};
              var cell_ref = XLSX.utils.encode_cell(address_of_cell);
              var desired_cell = sheet[cell_ref];
              var desired_value = (desired_cell ? desired_cell.v : undefined);
              var bi = 'MGEFI HTP - Saisir Devis ';
              const regex = new RegExp(bi+'*');
    
              var address_of_cell1 = {c:5, r:ra};
              var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
              var desired_cell1 = sheet[cell_ref1];
              var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);
    
              var conv = date;
              var j5 = conv - 5;
              if(regex.test(desired_value) && desired_value1>=j5)
              {
                nbr=nbr + 1;  
              }
              
            };
          }
    
    
          
          else
          {
            console.log('Colonne non trouvé');
          }
          var tab = [nbr];
          console.log("nombreeeeebr_NON_TRAITEES_DEVIS_J5 "+ nbr);
              return tab; 
    
          
        }
      
        catch
        {
          console.log("erreur absolu haaha iiiiiiiii");
        }
   
    

  },



};

