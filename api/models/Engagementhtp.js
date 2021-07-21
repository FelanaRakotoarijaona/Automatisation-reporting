/**
 * Engagementhtp.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
const path_reporting = 'D:/LDR8_1421_nouv/PROJET_FELANA/HTP/Test.xlsx';
// const path_reporting = 'D:/LDR8_1421_nouv/PROJET_FELANA/HTP/REPORTING_RETOUR.xlsx';
module.exports = {

  attributes: {

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
            numeroLigne.getCell(colonne).value = 1421;
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
  numeroLigne.getCell(collonne).value = 1422;
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
  numeroLigne.getCell(collonne).value = 1423;
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
  numeroLigne.getCell(collonne).value = 1424;
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
  numeroLigne.getCell(collonne).value = 1425;
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
  numeroLigne.getCell(collonne).value = 1426;
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
  numeroLigne.getCell(collonne).value = 1427;
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
    numeroLigne.getCell(collonne).value = 1427;
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
    numeroLigne.getCell(collonne).value = 1500;
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
    numeroLigne.getCell(collonne).value = 1500;
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
  numeroLigne.getCell(collonne).value = 1422;
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
  numeroLigne.getCell(collonne).value = 1423;
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
  numeroLigne.getCell(collonne).value = 1424;
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
  numeroLigne.getCell(collonne).value = 1425;
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
  numeroLigne.getCell(collonne).value = 1426;
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
  numeroLigne.getCell(collonne).value = 1427;
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
  numeroLigne.getCell(collonne).value = 1427;
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
  numeroLigne.getCell(collonne).value = 1500;
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
  numeroLigne.getCell(collonne).value = 1500;
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
  numeroLigne.getCell(collonne).value = 1422;
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
  numeroLigne.getCell(collonne).value = 1423;
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
  numeroLigne.getCell(collonne).value = 1424;
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
  numeroLigne.getCell(collonne).value = 1425;
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
  numeroLigne.getCell(collonne).value = 1426;
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
  numeroLigne.getCell(collonne).value = 1427;
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
  numeroLigne.getCell(collonne).value = 1427;
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
  numeroLigne.getCell(collonne).value = 1500;
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
  numeroLigne.getCell(collonne).value = 1500;
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
  numeroLigne.getCell(collonne).value = 1422;
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
  numeroLigne.getCell(collonne).value = 1423;
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
  numeroLigne.getCell(collonne).value = 1424;
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
  numeroLigne.getCell(collonne).value = 1425;
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
  numeroLigne.getCell(collonne).value = 1426;
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
  numeroLigne.getCell(collonne).value = 1427;
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
  numeroLigne.getCell(collonne).value = 1427;
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
  numeroLigne.getCell(collonne).value = 1427;
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
  numeroLigne.getCell(collonne).value = 1427;
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

};

