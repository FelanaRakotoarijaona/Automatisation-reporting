/*
 * ReportingInovcomExport.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
//CHEMIN DU FICHIER EXCEL DANS LE SERVEUR
//const path_reporting = 'D:/Reporting/Reporting/REPORTING INOVCOM Type.xlsx';
const path_reporting = '/dev/prod/00-TOUS/TestReporting/REPORTING INOVCOM Type.xlsx';
// const path_reporting = 'D:/LDR8_1421_nouv/PROJET_FELANA/REPORTING INOVCOM Type.xlsx';
module.exports = {
  attributes: {
  },
  ecritureOkKoSanteclair : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else
    {
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
         ligneDate1 = parseInt(rowNumber);
         var line = newworksheet.getRow(ligneDate1);
         var f = line.getCell(3).value;
         var mc1 = "^"+"SANTECLAIR"+"$";
         const regex = new RegExp(mc1,'i');
         if(regex.test(f))
         {
           ligneDate = parseInt(rowNumber);
         }
       }
     });
     var rowDate = newworksheet.getRow(ligneDate);
     var numeroLigne = rowDate;
     var iniValue = ReportingExcel.getIniValue(table);
     
     var a5;
 
     var rowm = newworksheet.getRow(1);
     var colonnne;
     var colDate1;
     /*rowm.eachCell(function(cell, colNumber) {
       if(cell.value == 'DOCUMENTS SAISIS')
       {
         colDate1 = parseInt(colNumber);
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
     console.log(" Colnumber2 "+collonne);*/
     numeroLigne.getCell('BP').value = parseInt(nombre_ok_ko.ok);
     numeroLigne.getCell('DR').value = parseInt(nombre_ok_ko.ko);
 
     
     await newWorkbook.xlsx.writeFile(path_reporting);
     sails.log("Ecriture OK KO termin??"); 
     return callback(null, "OK");
   
     }
     catch
     {
       console.log("Une erreur s'est produite");
       Reportinghtp.deleteToutHtp(table,3,callback);
     }
    }
   
    
     },
     ecritureOkKoEtatReste : async function (nombre_ok_ko, table,date_export,mois1,callback) {
      if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null)
      {
       console.log('ok' + nombre_ok_ko.ok);
       console.log('ko' + nombre_ok_ko.ko);
       return callback(null, "KO");
      }
      else
      {
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
           ligneDate1 = parseInt(rowNumber);
           var line = newworksheet.getRow(ligneDate1);
           var f = line.getCell(3).value;
           var mc1 = "reste";
           const regex = new RegExp(mc1,'i');
           if(regex.test(f))
           {
             ligneDate = parseInt(rowNumber);
           }
         }
       });
       var rowDate = newworksheet.getRow(ligneDate);
       var numeroLigne = rowDate;
       var iniValue = ReportingExcel.getIniValue(table);
       var a5;
       var rowm = newworksheet.getRow(1);
       numeroLigne.getCell('BG').value = parseInt(nombre_ok_ko.ok);
       numeroLigne.getCell('DI').value = parseInt(nombre_ok_ko.ko);
       await newWorkbook.xlsx.writeFile(path_reporting);
       sails.log("Ecriture OK KO termin??"); 
       return callback(null, "OK");
     
       }
       catch
       {
         console.log("Une erreur s'est produite");
         Reportinghtp.deleteToutHtp(table,3,callback);
       }
      }
     
      
       },
       /****************************************************/
       ecritureOkKoIllisible : async function (nombre_ok_ko, table,date_export,mois1,callback) {
        if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null)
        {
         console.log('ok' + nombre_ok_ko.ok);
         console.log('ko' + nombre_ok_ko.ko);
         return callback(null, "KO");
        }
        else
        {
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
             ligneDate1 = parseInt(rowNumber);
             var line = newworksheet.getRow(ligneDate1);
             var f = line.getCell(3).value;
             var mc1 = "^"+"ALMERYS"+"$";
             const regex = new RegExp(mc1,'i');
             if(regex.test(f))
             {
               ligneDate = parseInt(rowNumber);
             }
           }
         });
         var rowDate = newworksheet.getRow(ligneDate);
         var numeroLigne = rowDate;
         
         var a5;
     
         var rowm = newworksheet.getRow(1);
         var colonnne;
         var colDate1;
         numeroLigne.getCell('DA').value = parseInt(nombre_ok_ko.ok);
         numeroLigne.getCell('FC').value = parseInt(nombre_ok_ko.ko);
     
         
         await newWorkbook.xlsx.writeFile(path_reporting);
         sails.log("Ecriture OK KO termin??"); 
         return callback(null, "OK");
       
         }
         catch
         {
           console.log("Une erreur s'est produite");
           Reportinghtp.deleteToutHtp(table,3,callback);
         }
        }
       
        
         },
  /**********************************************************************/
         ecritureOkKoIllisiblecbtp : async function (nombre_ok_ko, table,date_export,mois1,callback) {
          if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null)
          {
           console.log('ok' + nombre_ok_ko.ok);
           console.log('ko' + nombre_ok_ko.ko);
           return callback(null, "KO");
          }
          else
          {
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
               ligneDate1 = parseInt(rowNumber);
               var line = newworksheet.getRow(ligneDate1);
               var f = line.getCell(3).value;
               var mc1 = "CBTP";
               const regex = new RegExp(mc1,'i');
               if(regex.test(f))
               {
                 ligneDate = parseInt(rowNumber);
               }
             }
           });
           var rowDate = newworksheet.getRow(ligneDate);
           var numeroLigne = rowDate;
           
           var a5;
       
           var rowm = newworksheet.getRow(1);
           var colonnne;
           var colDate1;
           numeroLigne.getCell('DA').value = parseInt(nombre_ok_ko.ok);
           numeroLigne.getCell('FC').value = parseInt(nombre_ok_ko.ko);
       
           
           await newWorkbook.xlsx.writeFile(path_reporting);
           sails.log("Ecriture OK KO termin??"); 
           return callback(null, "OK");
         
           }
           catch
           {
             console.log("Une erreur s'est produite");
             Reportinghtp.deleteToutHtp(table,3,callback);
           }
          }
         
          
           },
  // R??cuperer nombre okko
  countok : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlOk ="select okko from "+table; 
    var sqlOk ="select sum(okko::integer) from "+table;
    // var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'";
    console.log(sqlOk);
    
    // console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(res.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, res.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
        });
      },
      // function (callback) {
      //   ReportingInovcomExport.query(sqlKo, function(err, resKo){
      //     if (err) return res.badRequest(err);
      //     callback(null, resKo.rows[0].ko);
      //   });
      // },
    ],function(err,result){
      if(err) return res.badRequest(err);
      // console.log("Count OK ==> " + result[0]);
      // console.log("Count KO ==> " + result[1]);
      var okko = {};
      okko.ok = result[0];
      // okko.ko = result[1];
      return callback(null, okko);
    })
  },
  // R??cuperer nombre OK ou KO 
  countOkKo : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlOk ="select nbok from "+table; 
    var sqlOk ="select sum(nbok::integer) from "+table;
    // var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'";
    console.log(sqlOk);
    // console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(res.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, res.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
        });
      },
      // function (callback) {
      //   ReportingInovcomExport.query(sqlKo, function(err, resKo){
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
    })
  },
  // R??cupereation fueille 1
  countOkKo11 : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlOk ="select count(okko) as ok from "+table+" where okko='OK'"; //trameFlux
    // var sqlKo ="select nbko from "+table;
    var sqlKo ="select sum(nbko::integer) from "+table; 
    // console.log(sqlOk);
    console.log(sqlKo);
    async.series([
      // function (callback) {
      //   ReportingInovcomExport.query(sqlOk, function(err, res){
      //     if (err) return res.badRequest(err);
      //     callback(null, res.rows[0].ok);
      //   });
      // },
      function (callback) {
        ReportingInovcomExport.query(sqlKo, function(err, resKo){
          if (err) return res.badRequest(err);
          callback(null, resKo.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(resKo.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, resKo.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
        });
      },
    ],function(err,result){
      if(err) return res.badRequest(err);
      console.log("Count OK ==> " + result[0]);
      // console.log("Count KO ==> " + result[1]);
      var okko = {};
      okko.ko = result[0];
      // okko.ko = result[1];
      return callback(null, okko);
    })
  },
  // R??cupereation fueille 1
  countOkKofll1 : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlOk ="select nbok from "+table; 
    // var sqlKo ="select nbko from "+table;
    var sqlOk ="select sum(nbok::integer) from "+table; 
    var sqlKo ="select sum(nbko::integer) from  "+table;
    console.log(sqlOk);
    console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(res.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, res.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }  
        });
      },
      function (callback) {
        ReportingInovcomExport.query(sqlKo, function(err, resKo){
          if (err) return res.badRequest(err);
          callback(null, resKo.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(resKo.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, resKo.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
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
  // R??cupereation fueille 1
  countOkKofll11 : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlOk ="select nbrokrib from "+table; 
    var sqlOk ="select sum(nbrokrib::integer) from "+table;
    // var sqlKo ="select nbko from "+table;
    console.log(sqlOk);
    // console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(res.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, res.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
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
  // R??cupereation fueille 4
  countOkKofll4 : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlOk ="select nbok from "+table; //trameFlux
    // var sqlKo ="select nbko from "+table;
    var sqlOk ="select sum(nbok::integer) from "+table; 
    var sqlKo ="select sum(nbko::integer) from  "+table;
    console.log(sqlOk);
    console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(res.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, res.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
        });
      },
      function (callback) {
        ReportingInovcomExport.query(sqlKo, function(err, resKo){
          if (err) return res.badRequest(err);
          callback(null, resKo.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(resKo.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, resKo.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
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
  /*********************************************************************/
  // R??cuperer nombre okko dans inov 5 et 6
  countOkKo6 : function (table, callback) {
    const Excel = require('exceljs');
    var sqlOk ="select sum(nbok::integer) from "+table;
    var sqlKo ="select sum(nbko::integer) from "+table;
    console.log(sqlOk);
    console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(res.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, res.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
        });
      },
      function (callback) {
        ReportingInovcomExport.query(sqlKo, function(err, resKo){
          if (err) return res.badRequest(err);
          callback(null, resKo.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(resKo.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, resKo.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
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
  /************************************************************/
  // R??cupereation fueille 7
  countOkKofll7 : function (table, callback) {
    const Excel = require('exceljs');
    var sqlOk ="select sum(typologiedelademande::integer) from "+table;
    var sqlKo ="select sum(okko::integer) from "+table;
    console.log(sqlOk);
    console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(res.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, res.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
        });
      },
      function (callback) {
        ReportingInovcomExport.query(sqlKo, function(err, resKo){
          if (err) return res.badRequest(err);
          callback(null, resKo.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(resKo.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, resKo.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
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
  // R??cupereation fueille 8
  countOkKofll8 : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlKo ="select sum(typologiedelademande::integer) from "+table;
    var sqlOk ="select sum(okko::integer) from "+table;
    console.log(sqlOk);
    // console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(res.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, res.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
        });
      },
      // function (callback) {
      //   ReportingInovcomExport.query(sqlKo, function(err, resKo){
      //     if (err) return res.badRequest(err);
      //     callback(null, resKo.rows[0].sum);
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
  // R??cupereation fueille 9
  countOkKofll9 : function (table, callback) {
    const Excel = require('exceljs');
    // var sqlKo ="select sum(typologiedelademande::integer) from "+table;
    // var sqlOk ="select okko from "+table;
    var sqlOk ="select sum(okko::integer) from "+table;
    console.log(sqlOk);
    // console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].sum);
          // if (err) {
          //   console.log(err);
          //   //return null;
          // }
          // else
          // {
          //   if(res.rows[0])
          //   {
          //     console.log('ok');
          //     callback(null, res.rows[0].sum);
          //   }
          //   else
          //   {
          //     console.log("null");
          //     callback(null, 0);
          //   }
          // }
        });
      },
      // function (callback) {
      //   ReportingInovcomExport.query(sqlKo, function(err, resKo){
      //     if (err) return res.badRequest(err);
      //     callback(null, resKo.rows[0].sum);
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
  countOkKoTrameLamie2 : function (table, callback) {
    const Excel = require('exceljs');
    var sqlOk ="select count(okko) as ok from "+table+" where okko='OK' "; //trameFlux
    var sqlKo ="select count(okko) as ko from "+table+" where okko='KO' ";
    console.log(sqlOk);
    console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].ok);
        });
      },
      function (callback) {
        ReportingInovcomExport.query(sqlKo, function(err, resKo){
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
  countOkKoTrameLamie : function (table, callback) {
    const Excel = require('exceljs');
    var sqlOk ="select count(okko) as ok from "+table+" where okko='OK' AND typologiedelademande!='R??siliation' "; //trameFlux
    var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'  AND typologiedelademande!='R??siliation' ";
    console.log(sqlOk);
    console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].ok);
        });
      },
      function (callback) {
        ReportingInovcomExport.query(sqlKo, function(err, resKo){
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
  countOkKoTrameLamieResiliation : function (table, callback) {
    const Excel = require('exceljs');
    var sqlOk ="select count(okko) as ok from "+table+" where okko='OK' AND typologiedelademande='R??siliation' "; //trameFlux
    var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'  AND typologiedelademande='R??siliation' ";
    console.log(sqlOk);
    console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].ok);
        });
      },
      function (callback) {
        ReportingInovcomExport.query(sqlKo, function(err, resKo){
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
    var sqlOk ="select count(okko) as ok from "+table+" where okko='OK' AND typologiedelademande='R??siliation' "; 
    var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'  AND typologiedelademande='R??siliation' ";
    console.log(sqlOk);
    console.log(sqlKo);
    async.series([
      function (callback) {
        ReportingInovcomExport.query(sqlOk, function(err, res){
          if (err) return res.badRequest(err);
          callback(null, res.rows[0].ok);
        });
      },
      function (callback) {
        ReportingInovcomExport.query(sqlKo, function(err, resKo){
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
  // Convertion date
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
  /**********************************************************************************/
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 1
 ecritureOkKo : async function (nombre_ok_ko, table,date_export,mois1,callback) {
  if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
  {
   console.log('ok' + nombre_ok_ko.ok);
   console.log('ko' + nombre_ok_ko.ko);
   return callback(null, "KO");
  }
  else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }

    },
  /***********************************************************/  
  ecritureOkKo1 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    // var collonne;
    // var colDate2;
    // rowm.eachCell(function(cell, colNumber) {
    //   if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
    //   {
    //     colDate2 = parseInt(colNumber);
    //     var man = newworksheet.getRow(3);
    //     var f = man.getCell(colDate2).value;
    //     if(f == iniValue.ok)
    //     {
    //       collonne = parseInt(colNumber);
    //     }
    //   }
    // });
    // console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    // numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }

  }

    },
  /***********************************************************/  
  ecritureOkKo11 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    // var collonne;
    // var colDate2;
    // rowm.eachCell(function(cell, colNumber) {
    //   if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
    //   {
    //     colDate2 = parseInt(colNumber);
    //     var man = newworksheet.getRow(3);
    //     var f = man.getCell(colDate2).value;
    //     if(f == iniValue.ok)
    //     {
    //       collonne = parseInt(colNumber);
    //     }
    //   }
    // });
    // console.log(" Colnumber2"+collonne);
    // numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }

    },
  /***********************************************************/  
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 2
  ecritureOkKo2 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
        //console.log();
        //console.log(iniValue.ok);
        var getko_ini = man.getCell(colDate1).address;
          // console.log(getko_ini);
        if(getko_ini == iniValue.ko+3 && f == iniValue.ok)
        {
          colonnne = parseInt(colNumber);
        }
        }
    });
    console.log(" Colnumber"+colonnne);
    /*var collonne;
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
    console.log(" Colnumber2"+collonne);*/
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    //numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }

  } 
    },
  /**********************************************************/
  ecritureOkKo21 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();

          try{
            console.log('ecriture ok ko 21');
            await newWorkbook.xlsx.readFile(path_reporting);
            const newworksheet = newWorkbook.getWorksheet(mois1);
            var colonneDate = newworksheet.getColumn('A');
            var ligneDate1;
            var ligneDate;
            colonneDate.eachCell(function(cell, rowNumber) {
                var dateExcel = ReportingInovcomExport.convertDate(cell.text);
                if(dateExcel==date_export)
                {
                    ligneDate1 = parseInt(rowNumber);
                    var line = newworksheet.getRow(ligneDate1);
                    var f = line.getCell(3).value;
                    // console.log(f);
                    // var ai = "SANTECLAIR";
                    // const regex = new RegExp(ai,'i');
                    // // if(regex.test(f) == true)
                    // if(regex.test(f))
                    if(f == "SANTECLAIR" || f == "SANTECLAIR ")
                    {
                      ligneDate = parseInt(rowNumber);
                    }
                  }
              });
              console.log("LIGNE DATE ===> "+ ligneDate);
              var rowDate = newworksheet.getRow(ligneDate);
              var numeroLigne = rowDate;
              var iniValue = ReportingInovcomExport.getIniValue(table);
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
                      // console.log(getko_ini);
                    if(getko_ini == iniValue.ko+3 && f == iniValue.ok)
                    {
                      colonnne = parseInt(colNumber);
                    }
                  }
              });
              console.log(" Colnumber "+colonnne);
              numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
              await newWorkbook.xlsx.writeFile(path_reporting);
              sails.log("Ecriture OK KO termin??"); 
              return callback(null, "OK");

          }
          catch
          {
            console.log("Une erreur s'est produite");
            Reportinghtp.deleteToutHtp(table,3,callback);
          }

        }



  },

  /**********************************************************/
 
  ecritureOkKo22 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
      console.log('ecriture ok ko 22');
      await newWorkbook.xlsx.readFile(path_reporting);
      const newworksheet = newWorkbook.getWorksheet(mois1);
      var colonneDate = newworksheet.getColumn('A');
      var ligneDate1;
      var ligneDate;
      colonneDate.eachCell(function(cell, rowNumber) {
        var dateExcel = ReportingInovcomExport.convertDate(cell.text);
        if(dateExcel==date_export)
        {
          ligneDate1 = parseInt(rowNumber);
          var line = newworksheet.getRow(ligneDate1);
          var f = line.getCell(3).value;
          //console.log();
          if(f == "MGEFI")
          {
            ligneDate = parseInt(rowNumber);
          }
        }
      });
      console.log("LIGNE DATE ===> "+ ligneDate);
      var rowDate = newworksheet.getRow(ligneDate);
      var numeroLigne = rowDate;
      var iniValue = ReportingInovcomExport.getIniValue(table);
      
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }

  }


  },

  /**********************************************************/
  ecritureOkKo23 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
      console.log('ecriture ok ko 23');
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        var ai = "PUBLIPOSTAGE";
        const regex = new RegExp(ai,'i');
        if(regex.test(f))
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }

  }


    },
  
  /**********************************************************/
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 3
  ecritureOkKo3 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
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
        var a = iniValue.ok;
        // console.log('a'+a);
        // console.log('f'+f);
        const regex = new RegExp(a,'i');
        if(regex.test(f) == true)
        // if(f == iniValue.ok)
        {
          colonnne = parseInt(colNumber);
        }
        }
    });
    console.log(" Colnumber"+colonnne);
   /* var collonne;
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
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);*/
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    // numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }


  }

    },
  /**********************************************************/
  ecritureOkKo31 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
    var a5;

    var rowm = newworksheet.getRow(1);
   /* var colonnne;
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
    console.log(" Colnumber"+colonnne);*/
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        var a = iniValue.ok;
        const regex = new RegExp(a,'i');
        var getko_ini = man.getCell(colDate2).address;
        if(getko_ini == iniValue.ko+3 && regex.test(f) == true)
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    // numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }

  }

    },
  /**********************************************************/
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 4
  ecritureOkKo4 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }

  }


    },
  /**********************************************************/
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 4
  ecritureOkKo4etat : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        //console.log();
        if(f == "Etat des restes " || f == "Etat des restes")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }  
  
  
    },
  /**********************************************************/
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 4
  ecritureOkKo4pack : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        //console.log();
        if(f == "Pack Sp??. CBTP")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }




    },
  /**********************************************************/
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 5
  ecritureOkKo5 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }




    },
  /**********************************************************/
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 6
  ecritureOkKo6 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }



    },
  /**********************************************************/
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 7
  ecritureOkKo7 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }



    },
  /**********************************************************/
  ecritureOkKo7bis : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }


    },
  /**********************************************************/
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 8
  ecritureOkKo8 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
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
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
    var a5;

    var rowm = newworksheet.getRow(1);
   /* var colonnne;
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
    console.log(" Colnumber"+colonnne);*/
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
    // numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }


    },
  /**********************************************************/
  ecritureOkKo81 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        //console.log();
        if(f == "Pack Sp??. CBTP")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
    var a5;

    var rowm = newworksheet.getRow(1);
   /* var colonnne;
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
    console.log(" Colnumber"+colonnne);*/
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
    // numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }



    },
  /**********************************************************/
  //FONCTION POUR REMPLIR LE FICHIER EXCEL INOVCOM 9
  ecritureOkKo9 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        //console.log();
        if(f == "Pack Sp?? .ALMERYS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
    var a5;

    var rowm = newworksheet.getRow(1);
   /* var colonnne;
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
    console.log(" Colnumber"+colonnne);*/
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        var a = iniValue.ok;
        const regex = new RegExp(a,'i');
        var getko_ini = man.getCell(colDate2).address;
        if(getko_ini == iniValue.ko+3 && regex.test(f) == true)
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    // numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }



    },
   /**********************************************************/
 
   ecritureOkKofll11 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
      console.log('ecriture ok ko feuille 11');
      await newWorkbook.xlsx.readFile(path_reporting);
      const newworksheet = newWorkbook.getWorksheet(mois1);
      var colonneDate = newworksheet.getColumn('A');
      var ligneDate1;
      var ligneDate;
      colonneDate.eachCell(function(cell, rowNumber) {
        var dateExcel = ReportingInovcomExport.convertDate(cell.text);
        if(dateExcel==date_export)
        {
          ligneDate1 = parseInt(rowNumber);
          var line = newworksheet.getRow(ligneDate1);
          var f = line.getCell(3).value;
          //console.log();
          if(f == "Pack Sp?? .ALMERYS")
          {
            ligneDate = parseInt(rowNumber);
          }
        }
      });
      console.log("LIGNE DATE ===> "+ ligneDate);
      var rowDate = newworksheet.getRow(ligneDate);
      var numeroLigne = rowDate;
      var iniValue = ReportingInovcomExport.getIniValue(table);
      
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }



  },

  /**********************************************************/
  ecritureOkKofll11cbtp : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
      console.log('ecriture feuille 11 CBTP');
      await newWorkbook.xlsx.readFile(path_reporting);
      const newworksheet = newWorkbook.getWorksheet(mois1);
      var colonneDate = newworksheet.getColumn('A');
      var ligneDate1;
      var ligneDate;
      colonneDate.eachCell(function(cell, rowNumber) {
        var dateExcel = ReportingInovcomExport.convertDate(cell.text);
        if(dateExcel==date_export)
        {
          ligneDate1 = parseInt(rowNumber);
          var line = newworksheet.getRow(ligneDate1);
          var f = line.getCell(3).value;
          //console.log();
          if(f == "Pack Sp??. CBTP")
          {
            ligneDate = parseInt(rowNumber);
          }
        }
      });
      console.log("LIGNE DATE ===> "+ ligneDate);
      var rowDate = newworksheet.getRow(ligneDate);
      var numeroLigne = rowDate;
      var iniValue = ReportingInovcomExport.getIniValue(table);
      
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }


  },

  /**********************************************************/
  ecritureOkKofll11sante : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        if(f == "SANTECLAIR" || f == "SANTECLAIR ")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
    var a5;

    var rowm = newworksheet.getRow(1);
    var colonnne;
    var colDate1;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS SAISIS')
      {
        colDate1 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate1).value;
        if(f == iniValue.ok)
        {
          colonnne = parseInt(colNumber);
        }
        }
    });
    console.log(" Colnumber"+colonnne);
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }

  }


    },
   /**********************************************************/
  ecritureOkKofll12 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        if(f == "ALMERYS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
    var a5;

    var rowm = newworksheet.getRow(1);
    var colonnne;
    var colDate1;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS SAISIS')
      {
        colDate1 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate1).value;
        var a = iniValue.ok;
        const regex = new RegExp(a,'i');
        var getko_ini = man.getCell(colDate1).address;
        if(getko_ini == iniValue.ko+3 && regex.test(f) == true)
        {
          colonnne = parseInt(colNumber);
        }
        }
    });
    console.log(" Colnumber"+colonnne);
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
  }


    },
  /**********************************************************/
  ecritureOkKofll12retours : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        if(f == "ALMERYS")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }

  }


    },
     /***********************************************************/  
     ecritureOkKoIllisiblecbtp__ : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    if(nombre_ok_ko.ok==null && nombre_ok_ko.ko==null || nombre_ok_ko.ok==null && nombre_ok_ko.ko==undefined)
    {
     console.log('ok' + nombre_ok_ko.ok);
     console.log('ko' + nombre_ok_ko.ko);
     return callback(null, "KO");
    }
    else{
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        //console.log();
        if(f == "Pack Sp??. CBTP")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingInovcomExport.getIniValue(table);
    
    var a5;

    var rowm = newworksheet.getRow(1);
    var colonnne;
    var colDate1;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS ENVOYES')
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
    numeroLigne.getCell(colonnne).value = parseInt(nombre_ok_ko.ok);
    numeroLigne.getCell(collonne).value = parseInt(nombre_ok_ko.ko);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO termin??"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(table,3,callback);
    }

  }

    },
   /**********************************************************/
  //CONFIGURATION DU FICHIER INI
  getConfigIni : function() {
    const fs = require('fs');
    const ini = require('ini');
    const config = ini.parse(fs.readFileSync('./config_excelInovcom.ini', 'utf-8'));
    // console.log(config);
    return config;
  },

  //PRENDRE LA CONFIGURATION DU FICHIER INI
  getIniValue : function(table) {
    var iniValue = ReportingInovcomExport.getConfigIni();
    var numeroColonneOk,numeroColonneKo;

    
    if(table == "retourconventionsaisiedesconventions"){
      numeroColonneOk = iniValue.retourconventionsaisiedesconventions.ok;
      numeroColonneKo = iniValue.retourconventionsaisiedesconventions.ko;
    }
    if(table == "conventions"){
      numeroColonneOk = iniValue.conventions.ok;
      numeroColonneKo = iniValue.conventions.ko;
    }
    if(table == "ribtpmep"){
      numeroColonneOk = iniValue.ribtpmep.ok;
      numeroColonneKo = iniValue.ribtpmep.ko;
    }
    if(table == "tpmep"){
      numeroColonneOk = iniValue.tpmep.ok;
      numeroColonneKo = iniValue.tpmep.ko;
    }
    if(table == "curethermale"){
      numeroColonneOk = iniValue.curethermale.ok;
      numeroColonneKo = iniValue.curethermale.ko;
    }
    if(table == "dentaireretourfacturedentaireetcds"){
      numeroColonneOk = iniValue.dentaireretourfacturedentaireetcds.ok;
      numeroColonneKo = iniValue.dentaireretourfacturedentaireetcds.ko;
    }
    if(table == "optiqueretourpublipostage"){
      numeroColonneOk = iniValue.optiqueretourpublipostage.ok;
      numeroColonneKo = iniValue.optiqueretourpublipostage.ko;
    }
    if(table == "factureaudio"){
      numeroColonneOk = iniValue.factureaudio.ok;
      numeroColonneKo = iniValue.factureaudio.ko;
    }
    if(table == "retourhospipec"){
      numeroColonneOk = iniValue.retourhospipec.ok;
      numeroColonneKo = iniValue.retourhospipec.ko;
    }
    if(table == "retourpecdentaire"){
      numeroColonneOk = iniValue.retourpecdentaire.ok;
      numeroColonneKo = iniValue.retourpecdentaire.ko;
    }
    if(table == "retourpecoptique"){
      numeroColonneOk = iniValue.retourpecoptique.ok;
      numeroColonneKo = iniValue.retourpecoptique.ko;
    }
    if(table == "retourpecaudio"){
      numeroColonneOk = iniValue.retourpecaudio.ok;
      numeroColonneKo = iniValue.retourpecaudio.ko;
    }
    if(table == "santeclairtableauretourgeneral"){
      numeroColonneOk = iniValue.santeclairtableauretourgeneral.ok;
      numeroColonneKo = iniValue.santeclairtableauretourgeneral.ko;
    }
    if(table == "santeclairoptique"){
      numeroColonneOk = iniValue.santeclairoptique.ok;
      numeroColonneKo = iniValue.santeclairoptique.ko;
    }
    if(table == "noemiehtpmgefi"){
      numeroColonneOk = iniValue.noemiehtpmgefi.ok;
      numeroColonneKo = iniValue.noemiehtpmgefi.ko;
    }
    if(table == "mgefigtomgefirejetsaisienoemiehtp"){
      numeroColonneOk = iniValue.mgefigtomgefirejetsaisienoemiehtp.ok;
      numeroColonneKo = iniValue.mgefigtomgefirejetsaisienoemiehtp.ko;
    }
    if(table == "retourreclamtramereclamationtiers"){
      numeroColonneOk = iniValue.retourreclamtramereclamationtiers.ok;
      numeroColonneKo = iniValue.retourreclamtramereclamationtiers.ko;
    }
    if(table == "reclamsetramereclamationse"){
      numeroColonneOk = iniValue.reclamsetramereclamationse.ok;
      numeroColonneKo = iniValue.reclamsetramereclamationse.ko;
    }
    if(table == "reclamhospi"){
      numeroColonneOk = iniValue.reclamhospi.ok;
      numeroColonneKo = iniValue.reclamhospi.ko;
    }
    if(table == "dentairereclamationdentaire"){
      numeroColonneOk = iniValue.dentairereclamationdentaire.ok;
      numeroColonneKo = iniValue.dentairereclamationdentaire.ko;
    }
    if(table == "optiquetramereclamationoptique"){
      numeroColonneOk = iniValue.optiquetramereclamationoptique.ok;
      numeroColonneKo = iniValue.optiquetramereclamationoptique.ko;
    }
    if(table == "reclamationaudio"){
      numeroColonneOk = iniValue.reclamationaudio.ok;
      numeroColonneKo = iniValue.reclamationaudio.ko;
    }    
    if(table == "majribcbtp"){
      numeroColonneOk = iniValue.majribcbtp.ok;
      numeroColonneKo = iniValue.majribcbtp.ko;
    }
    if(table == "majagapsinteramc"){
      numeroColonneOk = iniValue.majagapsinteramc.ok;
      numeroColonneKo = iniValue.majagapsinteramc.ko;
    }
    if(table == "hospidemat"){
      numeroColonneOk = iniValue.hospidemat.ok;
      numeroColonneKo = iniValue.hospidemat.ko;
    }
    if(table == "extractionrcforce"){
      numeroColonneOk = iniValue.extractionrcforce.ok;
      numeroColonneKo = iniValue.extractionrcforce.ko;
    }
    if(table == "faveole"){
      numeroColonneOk = iniValue.faveole.ok;
      numeroColonneKo = iniValue.faveole.ko;
    }
    if(table == "favmgefi"){
      numeroColonneOk = iniValue.favmgefi.ok;
      numeroColonneKo = iniValue.favmgefi.ko;
    }
    if(table == "favbalma"){
      numeroColonneOk = iniValue.favbalma.ok;
      numeroColonneKo = iniValue.favbalma.ko;
    }
    if(table == "rcindeterminable"){
      numeroColonneOk = iniValue.rcindeterminable.ok;
      numeroColonneKo = iniValue.rcindeterminable.ko;
    }   
    if(table == "fav"){
      numeroColonneOk = iniValue.fav.ok;
      numeroColonneKo = iniValue.fav.ko;
    }
    if(table == "retourcmuc"){
      numeroColonneOk = iniValue.retourcmuc.ok;
      numeroColonneKo = iniValue.retourcmuc.ko;
    }
    if(table == "hospidematrejetprive"){
      numeroColonneOk = iniValue.hospidematrejetprive.ok;
      numeroColonneKo = iniValue.hospidematrejetprive.ko;
    }
    if(table == "defraiment"){
      numeroColonneOk = iniValue.defraiment.ok;
      numeroColonneKo = iniValue.defraiment.ko;
    }
    if(table == "retouravisannulationtramealmerys"){
      numeroColonneOk = iniValue.retouravisannulationtramealmerys.ok;
      numeroColonneKo = iniValue.retouravisannulationtramealmerys.ko;
    }
    if(table == "retouravisannulationcbtp"){
      numeroColonneOk = iniValue.retouravisannulationcbtp.ok;
      numeroColonneKo = iniValue.retouravisannulationcbtp.ko;
    }
    if(table == "recherchefactureinteriale"){
      numeroColonneOk = iniValue.recherchefactureinteriale.ok;
      numeroColonneKo = iniValue.recherchefactureinteriale.ko;
    }
    if(table == "inovgestionillisible"){
      numeroColonneOk = iniValue.inovgestionillisible.ok;
      numeroColonneKo = iniValue.inovgestionillisible.ko;
    }
    // if(table == "inovgestionillisiblecbtp"){
    //   numeroColonneOk = iniValue.inovgestionillisiblecbtp.ok;
    //   numeroColonneKo = iniValue.inovgestionillisiblecbtp.ko;
    // }
    if(table == "inovtpsalmerys"){
      numeroColonneOk = iniValue.inovtpsalmerys.ok;
      numeroColonneKo = iniValue.inovtpsalmerys.ko;
    }
    if(table == "inovsealmerys"){
      numeroColonneOk = iniValue.inovsealmerys.ok;
      numeroColonneKo = iniValue.inovsealmerys.ko;
    }
    if(table == "inovspehospi"){
      numeroColonneOk = iniValue.inovspehospi.ok;
      numeroColonneKo = iniValue.inovspehospi.ko;
    }
    if(table == "inovpackspedentaire"){
      numeroColonneOk = iniValue.inovpackspedentaire.ok;
      numeroColonneKo = iniValue.inovpackspedentaire.ko;
    }
    if(table == "inovpackspeoptique"){
      numeroColonneOk = iniValue.inovpackspeoptique.ok;
      numeroColonneKo = iniValue.inovpackspeoptique.ko;
    }
    if(table == "inovspeaudio"){
      numeroColonneOk = iniValue.inovspeaudio.ok;
      numeroColonneKo = iniValue.inovspeaudio.ko;
    }
    if(table == "inovetatdesreste"){
      numeroColonneOk = iniValue.inovetatdesreste.ok;
      numeroColonneKo = iniValue.inovetatdesreste.ko;
    }
    if(table == "inovtpscbtp"){
      numeroColonneOk = iniValue.inovtpscbtp.ok;
      numeroColonneKo = iniValue.inovtpscbtp.ko;
    }
    if(table == "inovsecbtp"){
      numeroColonneOk = iniValue.inovsecbtp.ok;
      numeroColonneKo = iniValue.inovsecbtp.ko;
    }
    if(table == "santeclairaudio"){
      numeroColonneOk = iniValue.santeclairaudio.ok;
      numeroColonneKo = iniValue.santeclairaudio.ko;
    }
    if(table == "inovaglaesynthese"){
      numeroColonneOk = iniValue.inovaglaesynthese.ok;
      numeroColonneKo = iniValue.inovaglaesynthese.ko;
    }
    if(table == "inovaglaefraudemms"){
      numeroColonneOk = iniValue.inovaglaefraudemms.ok;
      numeroColonneKo = iniValue.inovaglaefraudemms.ko;
    }
    if(table == "inovaglaeag2r"){
      numeroColonneOk = iniValue.inovaglaeag2r.ok;
      numeroColonneKo = iniValue.inovaglaeag2r.ko;
    }
    if(table == "inovaglaefraudeinteriale"){
      numeroColonneOk = iniValue.inovaglaefraudeinteriale.ok;
      numeroColonneKo = iniValue.inovaglaefraudeinteriale.ko;
    }
    if(table == "inovaglaefraudemg"){
      numeroColonneOk = iniValue.inovaglaefraudemg.ok;
      numeroColonneKo = iniValue.inovaglaefraudemg.ko;
    }
    if(table == "inovaglae100"){
      numeroColonneOk = iniValue.inovaglae100.ok;
      numeroColonneKo = iniValue.inovaglae100.ko;
    }
    if(table == "psfemajagaps"){
      numeroColonneOk = iniValue.psfemajagaps.ok;
      numeroColonneKo = iniValue.psfemajagaps.ko;
    }
    if(table == "favpharma"){
      numeroColonneOk = iniValue.favpharma.ok;
      numeroColonneKo = iniValue.favpharma.ko;
    }
    if(table == "favnument"){
      numeroColonneOk = iniValue.favnument.ok;
      numeroColonneKo = iniValue.favnument.ko;
    }
    
    var ok_ko = {};
    ok_ko.ok = numeroColonneOk;
    ok_ko.ko = numeroColonneKo;

    console.log("INI OK = "+ok_ko.ok);
    console.log("INI KO = "+ok_ko.ko);
    return ok_ko;
  },
};

