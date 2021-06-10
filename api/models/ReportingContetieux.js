/**
 * ReportingContetieux.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
const path_reporting = 'D:/Reporting/Reporting/REPORTING CONTENTIEUX type.xlsx';

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
          // if (err) return res.badRequest(err);
          // // callback(null, res.rows[0].ok);
          // console.log(res.rows[0].nb);
          // if(res.rows[0].nb != undefined){
          //   callback(null, res.rows[0].nb);
          // }
          // else{
          //   return res.rows[0].nb = 0;
          // }
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
          // if (err) return res.badRequest(err);
          // // callback(null, res.rows[0].ok);
          // console.log(res.rows[0].sum);
          // if(res.rows[0].sum != undefined){
          //   callback(null, res.rows[0].sum);
          // }
          // else{
          //   return res.rows[0].sum = 0;
          // }
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


   /* Import */
  importTrameFlux929type2 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
    var tab = [];
    tab = ReportingContetieux.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      ReportingContetieux.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<18;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log(j);
      ReportingContetieux.lectureEtInsertiontype2( trameflux,feuil,cellule,table,cellule2,y,numligne,callback);
    }
    };
  },
  lectureEtInsertiontype2:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = parseInt(numligne[nb]);
    try{
      var nbr = 0;
      const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var col = 0;
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
         
          var sql = "insert into "+table[nbe]+" (nb) values ('"+nbr+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err) return console.log(err);
                        else return callback(null, true);        
                                            });
          console.log("nombreeeeebr"+ nbr);
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
deleteFromChemin : function (table,callback) {
      var sql = "delete from chemincontetieux ";
      ReportingContetieux.getDatastore().sendNativeQuery(sql, function(err, res){
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
    importEssai4: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,callback) {
      const fs = require('fs');
      var re  = 'a';
      var tab = [];
      var a = table[0]+date+table2[nb];
      //var a ='\\\\10.128.1.2\\almerys-out\\Retour_Easytech_20210512\\TRAITEMENT_RETOUR_OTD_N2\\' ;
      var b = option[nb];
      //var b = 'OTD_ALMERYS SATD';
      //var c = 'vrai';
      //console.log(a);
      var nomTable = nomtable;
      var numLigne= numligne;
      var numFeuille = numfeuille;
      var nomColonne = nomcolonne;
      var c = Reportinghtp.existenceFichier(a);
      console.log(c);
      if(c=='vrai')
      {
        fs.readdir(a, (err, files) => {
          console.log(a);
              files.forEach(file => {
                const regex = new RegExp(b+'*');
                if(regex.test(file))
                {
                   //re = a+'\\'+file;
                   re = a+'/'+file;
                   var sql = "insert into chemincontetieux (chemin,nomtable,numligne,numfeuile,colonnecible) values ('"+re+"','"+nomTable[nb]+"','"+numLigne[nb]+"','"+numFeuille[nb]+"','"+nomColonne[nb]+"','"+colonnecible2+"') ";
                   Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
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
                else
                {
                 var sql = "insert into chemintsisy (typologiedelademande) values ('"+re+"') ";
                 Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
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
        Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
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
    },
    importEssai: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,callback) {
      const fs = require('fs');
      var re  = 'a';
      var tab = [];
      var a = table[0]+date+table2[nb];
      //var a ='\\\\10.128.1.2\\almerys-out\\Retour_Easytech_20210512\\TRAITEMENT_RETOUR_OTD_N2\\' ;
      var b = option[nb];
      //var b = 'OTD_ALMERYS SATD';
      //var c = 'vrai';
      //console.log(a);
      var nomTable = nomtable;
      var numLigne= numligne;
      var numFeuille = numfeuille;
      var nomColonne = nomcolonne;
      var c = ReportingInovcom.existenceFichier(a);
      console.log(c);
      if(c=='vrai')
      {
        fs.readdir(a, (err, files) => {
          console.log(a);
              files.forEach(file => {
                const regex = new RegExp(b+'*');
                if(regex.test(file))
                {
                   re = a+'\\'+file;
                   //console.log(re);  
                   var sql = "insert into chemincontetieux (chemin,nomtable,numligne,numfeuile,colonnecible) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                     if(err) return console.log(err);
                     else return callback(null, true);        
                                         }) 
                     
                } 
                else
                {
                 var sql = "insert into chemintsisy (typologiedelademande) values ('"+re+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                     if(err) return console.log(err);
                     else return callback(null, true);        
                                         }) 
                }
               
               
            });
            
           
          });
      }
      else
      {
        var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
        ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
          if(err) return console.log(err);
          else return callback(null, true);        
                              })   
      }   
    },
    deleteToutHtp : function (table,nb,callback) {
      var sql = "delete from "+table+" ";
      ReportingContetieux.getDatastore().sendNativeQuery(sql, function(err, res){
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
        var a = ReportingContetieux.existenceFichier(trameflux[i]);
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
        ReportingContetieux.deleteFichier(table,i,callback);
      };
    },
    deleteFichier: function (table,nb,callback) {
      var tab= '';
      console.log(tab);
      const fs = require('fs');
      fs.writeFile(table[nb]+'.txt', tab, (err) => {
        var sql = "insert into trame (typologiedelademande) values ('k') ";
        ReportingContetieux.getDatastore().sendNativeQuery(sql, function(err,res){
          if(err) return console.log(err);
          else return callback(null, true);        
                              })      
      });
    },
    
    deleteReportingHtp : function (table,nb,callback) {
      var sql = "delete from "+table[nb]+" ";
      ReportingContetieux.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return console.log(err); }
        return callback(null, true);
        });
    },
    deleteHtp : function (table,nb,callback) {
      var j;
      var i = parseInt(j);
      for(i=0;i<nb;i++)
      {
        ReportingContetieux.deleteReportingHtp(table,i,callback);
      };
    },

};
