/**
 * ReportingInovcom.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
 module.exports = {
  attributes: {},
  importTrameFlux929type2 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
    var tab = [];
    tab = ReportingInovcom.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      ReportingInovcom.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log(j);
      ReportingInovcom.lectureEtInsertiontype2( trameflux,feuil,cellule,table,cellule2,j,numligne,callback)
    // ReportingInovcom.lectureEtInsertion(trameflux,feuil,cellule,table,cellule2,j,callback)
    }
    };
  },
  lectureEtInsertiontype2:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = numligne[nb];
    try{
      var nbr = 0;
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
        /*console.log('cc' + cellule[nb]);
        console.log('dv'+ desired_value);*/
        if(desired_value==cellule[nb])
        {
          col=ra;
        }
      };
      console.log('colonne cible' +col);
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
          }
          console.log("nombreeeeebr"+ nbr);
          var sql = "insert into "+table[nb]+" (typologiedelademande,okko) values ('"+nbr+"','"+nbr+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err) return console.log(err);
                        else return callback(null, true);        
                                            })
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
      console.log('cc'+cellule[nb]);
      for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
       
        console.log('valeur'+desired_value);
        if(desired_value==cellule[nb])
        {
          col=ra;
        }
      };
      console.log('colonne cible' +col);
      var tab = [];
      var tabl = [];
      //console.log('table1' + tab);
      var col2;
      for(var ra=0;ra<=range.e.c;ra++)
        {
          var address_of_cell = {c:ra, r:numeroligne};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          //console.log(desired_cell.v);
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

            if(desired_value1!=undefined)
            {
              var m = desired_value1;
              m=m.replace("'", "''");
              var sql = "insert into "+table[nb]+" (typologiedelademande) values ('"+m+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err) return console.log(err);
                        else return callback(null, true);        
                                            })
            }
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
  lectureEtInsertiontype4:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[0];
    var numeroligne = numligne[0];
    console.log(trameflux[nb]);
    console.log(numeroligne);
    console.log(numerofeuille);
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
        if(desired_value==cellule[0])
        {
          col=ra;
        }
      };
      console.log('colonne cible' +col);
      var tab = [];
      var tabl = [];
      //console.log('table1' + tab);
      /*var col2;
      for(var ra=0;ra<=range.e.c;ra++)
        {
          var address_of_cell = {c:ra, r:numeroligne};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          //console.log(desired_cell.v);
          if(desired_value==cellule2[nb])
          {
            col2=ra;
          }
        };
      console.log('colonne cible2' +col2);*/
      if(col!=undefined)
      {
        for(var a=0;a<=range.e.r;a++)
          {
            var address_of_cell = {c:col, r:a};
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            var desired_value1 = (desired_cell ? desired_cell.v : undefined);
          

           /* var address_of_cell2 = {c:col2, r:a};
            var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
            var desired_cell2 = sheet[cell_ref2];
            var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);*/

            console.log(desired_value1);

            var sql = "insert into extractionrcforce (typologiedelademande,okko) values ('"+desired_value1+"','"+desired_value1+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
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
  importTrameFlux929type4 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
    var tab = [];
    tab = ReportingInovcom.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      ReportingInovcom.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log(j);
      ReportingInovcom.lectureEtInsertiontype4( trameflux,feuil,cellule,table,cellule2,j,numligne,callback)
    // ReportingInovcom.lectureEtInsertion(trameflux,feuil,cellule,table,cellule2,j,callback)
    }
    };
  },

  lectureEtInsertiontype5:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[0]);
    var numerofeuille = feuil[nb];
    var numeroligne = numligne[0];
    console.log(trameflux[0]);
    console.log(numeroligne);
    console.log(numerofeuille);
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[nb]];
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
        if(desired_value==cellule[0])
        {
          col=ra;
        }
      };
      console.log('colonne cible' +col);
      var tab = [];
      var tabl = [];
      //console.log('table1' + tab);
      var col2;
      for(var ra=0;ra<=range.e.c;ra++)
        {
          var address_of_cell = {c:ra, r:numeroligne};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          //console.log(desired_cell.v);
          if(desired_value==cellule2[0])
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

            //console.log(desired_value1);

            var sql = "insert into fav (typologiedelademande,okko) values ('"+desired_value1+"','"+desired_value1+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
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
  lectureEtInsertiontype8:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = numligne[nb];
    try{
      var nbr = 0;
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
            else
            {
              console.log('non trouvé');
            }
          }
         
         /* */
      }

      else
      {
        console.log('Colonne non trouvé');
      }
      console.log("nombreeeeebr"+ nbr);
      var sql = "insert into "+table[nb]+" (typologiedelademande,okko) values ('"+nbr+"','"+nbr+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err) return console.log(err);
                        else return callback(null, true);        
                                            })
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
    
  },

  importTrameFlux929type5 : async function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
    var tab = [];
    tab = ReportingInovcom.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      ReportingInovcom.deleteToutHtp(table,3,callback);
    }
    else{
      const Excel = require('exceljs');
      const cmd=require('node-cmd');
      const newWorkbook = new Excel.Workbook();
      try{
      await newWorkbook.xlsx.readFile(trameflux[0]);
      // var feuille = newWorkbook.getWorksheet();
      var test = newWorkbook.worksheets;
      var essaie = parseInt(test.length);
      for(var y=0;y<essaie;y++) //parcours anle dossier rehetra
      {
        /*var j = parseInt(tab[y]);*/
        console.log(y);
        ReportingInovcom.lectureEtInsertiontype5( trameflux,feuil,cellule,table,cellule2,y,numligne,callback);
      // ReportingInovcom.lectureEtInsertion(trameflux,feuil,cellule,table,cellule2,j,callback)
      }
    }
    catch
    {
      console.log('ko');
    }
    
  
    };
  },
  importTrameFlux929type6 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback) {
    var tab = [];
    tab = ReportingInovcom.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      ReportingInovcom.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log(j);
      ReportingInovcom.lectureEtInsertiontype6( trameflux,feuil,cellule,table,cellule2,j,numligne,dernierl,callback)
      //ReportingInovcom.lectureEtInsertion(trameflux,feuil,cellule,table,cellule2,j,callback)
    }
    };
  },
  lectureEtInsertiontype6:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback){
    var Excel = require('exceljs');
    var numerofeuille = feuil[nb];
    var numeroligne = numligne[nb];
    var workbook = new Excel.Workbook();
    try{
      workbook.xlsx.readFile(trameflux[nb])
        .then(function() {
            var newworksheet = workbook.getWorksheet(numerofeuille);
            var row = newworksheet.getRow(numeroligne);
            var a;
            row.eachCell(function(cell, colNumber) {
              if(cell.text==dernierl[nb])
              {
                a = parseInt(colNumber);
              }
            });
            var b;
            row.eachCell(function(cell, colNumber) {
              if(cell.text==cellule[nb])
              {
                b = parseInt(colNumber);
              }
            });
            var c;
            row.eachCell(function(cell, colNumber) {
              if(cell.text==cellule2[nb])
              {
                c = parseInt(colNumber);
              }
            });
            if(a!=undefined && b!=undefined && c!=undefined)
            {
              var col = newworksheet.getColumn(a);
              //var tab = [];
              console.log('col' + col);
              var date = 'Sat Sep 14 2020 03:00:00 GMT+0300 (GMT+03:00)';
              var daty ;
              col.eachCell(function(cell, rowNumber) {
                var date1 = cell.text ;
                if(date1>date)
                {
                  daty=rowNumber;
                }
              });

            var row1 = newworksheet.getRow(daty);
            var ok = row1.getCell(b);
            var ko = row1.getCell(c);
            
            console.log('ok' + ok.text);
            console.log('ko' + ko.text);
            var sql = "insert into retourcmuc (typologiedelademande,okko) values ('"+ok+"','"+ko+"') ";
                    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if(err) return console.log(err);
                      else return callback(null, true);        
                                          }) 
            }
            else
            {
              console.log("Nom de colonne non trouvé");
            }
           
        });
    }
    catch
    {
      console.log("Erreur trouvé");
    }
  },
  importTrameFlux929type7 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback) {
    var tab = [];
    tab = ReportingInovcom.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      ReportingInovcom.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log(j);
      ReportingInovcom.lectureEtInsertiontype7( trameflux,feuil,cellule,table,cellule2,j,numligne,dernierl,callback)
    // ReportingInovcom.lectureEtInsertion(trameflux,feuil,cellule,table,cellule2,j,callback)
    }
    };
  },
  lectureEtInsertiontype7:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = parseInt(numligne[nb]);
    console.log('lign' +numeroligne);
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var col ;
      var col1;
      var col2;
      console.log('Nombre de colonne' + range.e.c);
      console.log('Nombre de ligne' + range.e.r);
      console.log(numeroligne + 'numlign');
      console.log(numerofeuille + 'numfeuille');
      console.log(cellule2[nb] + 'c2');
      console.log(cellule[nb] + 'c1');
      console.log(dernierl[nb] + 'c3');
      console.log(table[nb] + 'table');
      for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(desired_value==dernierl[nb])
        {
          col=ra;
        }
      };
      for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(desired_value==cellule[nb])
        {
          col1=ra;
        }
      };
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
      console.log('colonne cible' +col + col1 +col2);
      if(col!=undefined && col1!=undefined  && col2!=undefined) 
      {
        var tabok = 0;
        var taboki = 0;
        for(var a=0;a<=range.e.r;a++)
          {
            var address_of_cell = {c:col, r:a};
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            var desired_value1 = (desired_cell ? desired_cell.v : undefined);
            var bi = 'Total';
            const regex = new RegExp(bi+'*');
            if(regex.test(desired_value1))
            {
              var z = parseInt(a) - 1;
              var address_of_cell2 = {c:col1, r:z};
              var cell_refs = XLSX.utils.encode_cell(address_of_cell2);
              var desired_cell2 = sheet[cell_refs];
              var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);

              var address_of_cell21 = {c:col2, r:z};
              var cell_refs1 = XLSX.utils.encode_cell(address_of_cell21);
              var desired_cell21 = sheet[cell_refs1];
              var desired_value21 = (desired_cell21 ? desired_cell21.v : undefined);

              if(desired_value2!=undefined)
              {
                tabok= tabok + 1;
                //console.log('ok');
              }
              else
              {
                //console.log('ko');
              };

              if(desired_value21!=undefined)
              {
                taboki = taboki +1;
                //console.log('ok2');
              }
              else
              {
                //console.log('ko2');
              };

            };
          };
          console.log('nb =' + tabok);
          console.log('nb2 =' + taboki);
          var sql = "insert into hospidematrejetprive (typologiedelademande,okko) values ('"+tabok+"','"+taboki+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err) return console.log(err);
                        else return callback(null, true);        
                                           });
      }
      else
      {
        console.log('Colonne non trouvé');
      };
      
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
  },
  importTrameFlux929type8 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback) {
    var tab = [];
    tab = ReportingInovcom.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      ReportingInovcom.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log(j);
      ReportingInovcom.lectureEtInsertiontype8( trameflux,feuil,cellule,table,cellule2,j,numligne,dernierl,callback)
    // ReportingInovcom.lectureEtInsertion(trameflux,feuil,cellule,table,cellule2,j,callback)
    }
    };
  },
  importTrameFlux929 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
    var tab = [];
    tab = ReportingInovcom.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      ReportingInovcom.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log(j);
      ReportingInovcom.lectureEtInsertion2( trameflux,feuil,cellule,table,cellule2,j,numligne,callback)
    // ReportingInovcom.lectureEtInsertion(trameflux,feuil,cellule,table,cellule2,j,callback)
    }
    };
  },
 deleteFromChemin : function (table,callback) {
      var sql = "delete from chemininovcom ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    deleteFromChemin2 : function (table,callback) {
      var sql = "delete from chemininovcomtype2 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    deleteFromChemin3 : function (table,callback) {
      var sql = "delete from chemininovcomtype3 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    deleteFromChemin4 : function (table,callback) {
      var sql = "delete from chemininovcomtype4 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    deleteFromChemin5 : function (table,callback) {
      var sql = "delete from chemininovcomtype5 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    deleteFromChemin6 : function (table,callback) {
      var sql = "delete from chemininovcomtype6 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    deleteFromChemin7 : function (table,callback) {
      var sql = "delete from chemininovcomtype7 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    deleteFromChemin8 : function (table,callback) {
      var sql = "delete from chemininovcomtype8 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    deleteFromChemin9 : function (table,callback) {
      var sql = "delete from chemininovcomtype9 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    deletetype9 : function (table,callback) {
      var sql = "delete from "+table+" ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
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

    importEssaitype9: function (table,table2,date,option,nb,callback) {
      const fs = require('fs');
      var re  = 'a';
      //var a = '\\\\10.128.1.2\\almerys-out\\Retour_Easytech_20210428\\RETOUR_RECHERCHE FACTURE INTERIALE\\INTERIALE';
      var a = table[0]+date+table2[nb];
      var c = ReportingInovcom.existenceFichier(a);
      console.log(c);
      if(c=='vrai')
      {
        var re = 0;
        fs.readdir(a, (err, files) => {
          console.log(a);
              files.forEach(file => {
                const regex = new RegExp('.pdf');
  
                if(regex.test(file))
                {
                   re = re + 1;
                   
                } 
            });
            console.log(re); 
            var sql = "insert into recherchefactureinteriale (typologiedelademande,okko) values ("+re+","+re+") ";
             ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
              if(err) return console.log(err);
              else return callback(null, true);        
                                  })   
          });
      }
      else
      {
        var sql = "insert into recherchefactureinteriale (typologiedelademande,okko) values (0,0)";
        ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
          if(err) return console.log(err);
          else return callback(null, true);        
                              })   
      }   
  
    },

    importEssai: function (table,table2,date,option,nb,callback) {
      const fs = require('fs');
      var re  = 'a';
      var tab = [];
      var a = table[0]+date+table2[nb];
      var b = option[nb];
      //console.log(a);
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
                   re = file;
                   console.log(re);  
                } 
            });
            var sql = "insert into chemininovcom (typologiedelademande) values ('"+re+"') ";
                    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if(err) return console.log(err);
                      else return callback(null, true);        
                                          })  
            console.log('ato anatiny'+re);
          });
      }
      else
      {
        var sql = "insert into chemininovcom (typologiedelademande) values ('k') ";
        ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
          if(err) return console.log(err);
          else return callback(null, true);        
                              })   
      }   
    },
    importEssaitype2: function (table,table2,date,option,nb,callback) {
      const fs = require('fs');
     var re  = 'a';
     var tab = [];
     var a = table[0]+date+table2[nb];
     var b = option[nb];
     //console.log(a);
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
                  re = file;
                  console.log(re);  
               } 
           });
           var sql = "insert into chemininovcomtype2 (typologiedelademande) values ('"+re+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                     if(err) return console.log(err);
                     else return callback(null, true);        
                                         })  
           console.log('ato anatiny'+re);
         });
     }
     else
     {
       var sql = "insert into chemininovcomtype2 (typologiedelademande) values ('k') ";
       ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
         if(err) return console.log(err);
         else return callback(null, true);        
                             })   
     }   
   },
   importEssaitype3: function (table,table2,date,option,nb,date2,callback) {
    const fs = require('fs');
   var re  = 'a';
   var tab = [];
   var a = table[0]+date+table2[nb]+date2;
   var b = option[nb];
   //console.log(a);
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
                re = file;
                console.log(re);  
             } 
         });
         var sql = "insert into chemininovcomtype3 (typologiedelademande) values ('"+re+"') ";
                 ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                   if(err) return console.log(err);
                   else return callback(null, true);        
                                       })  
         console.log('ato anatiny'+re);
       });
   }
   else
   {
     var sql = "insert into chemininovcomtype3 (typologiedelademande) values ('k') ";
     ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
       if(err) return console.log(err);
       else return callback(null, true);        
                           })   
   }   
 },
   importEssaitype4: function (table,table2,date,option,nb,callback) {
    const fs = require('fs');
   var re  = 'a';
   var tab = [];
   var a = table[0]+date+table2[nb];
   var b = option[nb];
   //console.log(a);
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
                re = file;
                console.log(re);  
             } 
             var sql = "insert into chemininovcomtype4 (typologiedelademande) values ('"+re+"') ";
             ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
               if(err) return console.log(err);
               else return callback(null, true);        
                                   })  
             console.log('ato anatiny'+re);
         });
         
        
       });
   }
   else
   {
     var sql = "insert into chemininovcomtype4 (typologiedelademande) values ('k') ";
     ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
       if(err) return console.log(err);
       else return callback(null, true);        
                           })   
   }   
 },
 importEssaitype5: function (table,table2,date,option,nb,callback) {
  const fs = require('fs');
 var re  = 'a';
 var tab = [];
 var a = table[0]+date+table2[nb];
 var b = option[nb];
 //console.log(a);
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
              re = file;
              console.log(re);  
           } 
         
          
       });
       var sql = "insert into chemininovcomtype5 (typologiedelademande) values ('"+re+"') ";
       ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
        if(err) return console.log(err);
        else return callback(null, true);        
                            })  
      console.log('ato anatiny'+re);
       
      
     });
 }
 else
 {
   var sql = "insert into chemininovcomtype5 (typologiedelademande) values ('k') ";
   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
     if(err) return console.log(err);
     else return callback(null, true);        
                         })   
 }   
},
importEssaitype6: function (table,table2,date,option,nb,callback) {
  const fs = require('fs');
 var re  = 'a';
 var tab = [];
 var a = table[0]+date+table2[nb];
 var b = option[nb];
 //console.log(a);
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
              re = file;
              console.log(re);  
           } 
         
          
       });
       var sql = "insert into chemininovcomtype6 (typologiedelademande) values ('"+re+"') ";
       ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
        if(err) return console.log(err);
        else return callback(null, true);        
                            })  
      console.log('ato anatiny'+re);
       
      
     });
 }
 else
 {
   var sql = "insert into chemininovcomtype (typologiedelademande) values ('k') ";
   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
     if(err) return console.log(err);
     else return callback(null, true);        
                         })   
 }   
},
importEssaitype7: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,callback) {
 const fs = require('fs');
 var re  = 'a';
 var tab = [];
 var ab = table[0]+date+table2[nb];
 var b = option[nb];
 //console.log(a);
 var c = ReportingInovcom.existenceFichier(ab);
 console.log(c);
 if(c=='vrai')
 {
  fs.readdir(ab, (err, files) => {
    if(err){
      console.log('ito le erreur : '+err);
    }
    else{
      var a;
      files.forEach(file =>{
        for(var i = 0; i < files.length; i++){
              if(file == files[i]){
              const test1 = ab +files[i];
              fs.readdir(test1, (err, files1) => {
                if(err){
                  console.log(err);
                }
                else{
                  //console.log(file +" " +  files1[files1.length-1]);
                  //var cible = "MASQUE SAISIE";
                  const regex = new RegExp(b+'*');
                  for(var i = 0; i < files1.length; i++){
                    if(regex.test(files1[i]))
                    {
                      var a =ab + file +"\\" + files1[i];
                      console.log('*****************');
                      console.log(a);  
                      var sql = "insert into chemininovcomtype7 (chemin,nomtable,numligne,numfeuile,colonnecible,colonnecible2,colonnecible3) values ('"+a+"','"+nomtable+"','"+numligne+"','"+numfeuille+"','"+nomcolonne+"','"+nomcolonne2+"','"+nomcolonne3+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err) return console.log(err);
                        else return callback(null, true);        
                                            }) 
                    } 
                    else
                    {
                      var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err) return console.log(err);
                        else return callback(null, true);        
                                            }) 
                    };
                  };
                }
              });
            }
            
        }

      })
    };
 });
}
 else
 {
   var sql = "insert into chemintsisy (typologiedelademande) values ('k') ";
   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
     if(err) return console.log(err);
     else return callback(null, true);        
                         })   
 }   
},
importEssaitype8: function (table,table2,date,option,nb,type,type2,callback) {
  const fs = require('fs');
  var re  = 'a';
  var tab = [];
  var ab = table[0]+date+table2[nb];
  var b = option[nb];
  //console.log(a);
  var c = ReportingInovcom.existenceFichier(ab);
  console.log(c);
  if(c=='vrai')
  {
    fs.readdir(ab, (err, files) => {
      if(err){
        console.log('ito le erreur : '+err);
      }
      else{
      files.forEach(file =>{
        
        for(var i = 0; i < files.length; i++){
              if(file == files[i]){
              const test1 = ab +files[i] + type[nb] ;
              var m = ReportingInovcom.existenceFichier(test1);
              if(m=='vrai')
              {
                fs.readdir(test1, (err, files1) => {
                  if(err){
                    console.log(err);
  
                  }
                  else{
                    console.log(file +" " +  files1[files1.length-1]);
                    console.log('ok');
    
                    var a = file  + type2[nb] + files1[files1.length-1];  
                    console.log("haha"+a);
                    var sql = "insert into chemininovcomtype8 (typologiedelademande) values ('"+a+"') ";
                    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if(err) return console.log(err);
                      else return callback(null, true);        
                                          }) 
    
                  }
                  //console.log("haha"+a);
                 
                })
              }
              else
              {
                var sql = "insert into chemininovcomtype8 (typologiedelademande) values ('k') ";
                ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                  if(err) return console.log(err);
                  else return callback(null, true);        
                                      })   
              }
             
            }
            else
            {
             
            }
            
        }

      })
    }
 });
}
 else
 {
   var sql = "insert into chemininovcomtype8 (typologiedelademande) values ('k') ";
   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
     if(err) return console.log(err);
     else return callback(null, true);        
                         })   
 }   
},
    importInovcom: function (trameflux,feuil,cellule,table,cellule2,numligne,nb,callback) {
      var tab = [];
      tab = ReportingInovcom.totalFichierExistant(trameflux,nb,callback);
      console.log(tab);
      if(tab.length==0)
      {
        console.log('Aucune reporting pour ce date');
        ReportingInovcom.deleteToutHtp(table,3,callback);
        
      }
      else{
        for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
      {
        var j = parseInt(tab[y]);
        console.log(j);
        ReportingInovcom.insertion(trameflux,feuil,cellule,table,cellule2,j,numligne,callback);
      }
      };
    },
    insertion:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
      var tab= ReportingInovcom.lectureEtInsertionModifie(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
      console.log(tab);
      const fs = require('fs');
      fs.writeFile(table[nb]+'.txt', tab, (err) => {
              
              var sql = "insert into trame (typologiedelademande) values ('ok') ";
                  ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                    if(err) return console.log(err);
                    else return callback(null, true);        
                                        })
             
      });
    },
    lectureEtInsertionModifie:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
      XLSX = require('xlsx');
      var workbook = XLSX.readFile(trameflux[nb]);
      var first_sheet_name = workbook.SheetNames;
      var numerofeuille = feuil[nb];
      var numeroligne = numligne[nb];
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
      for(var a=0;a<=range.e.r;a++)
      {
        var address_of_cell = {c:col, r:a};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value1 = (desired_cell ? desired_cell.v : undefined);
        tab.push(desired_value1);
  
        var address_of_cell2 = {c:col2, r:a};
        var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
        var desired_cell2 = sheet[cell_ref2];
        var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
        tabl.push(desired_value2);
      }
      var com = [];
      for(var i=0;i<tab.length;i++)
      {
          com.push(tab[i]+';'+tabl[i]+'\n');
      };
      console.log(com);
      return com;
    },
    deleteToutHtp : function (table,nb,callback) {
      var sql = "delete from "+table+" ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    importTout: function (trameflux,table,nb,callback) {
      var tab = [];
      tab = ReportingInovcom.totalFichierExistant(trameflux,nb,callback);
      console.log('table miexiste'+tab);
      for(var i=0;i<nb;i++){
        var j = parseInt(tab[i]);
        ReportingInovcom.importFinal(table,i,callback);
      };
    },
    importFinal: function (table,nb,callback) {
      var tablem = table[nb];
      console.log('tablem'+tablem);
      var chemin = 'D:/projet/'+tablem+'.txt';
      console.log(chemin);
      var sql = " COPY "+tablem+" FROM '"+chemin+"'  (DELIMITER(';') ) ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
        if(err) return console.log('erreur');
        else return callback(null, true);        
                            });
    },
    totalFichierExistant : function (trameflux,nb,callback) {
      var tab = [];
      var j;
      var i = parseInt(j);
      for(i=0;i<nb;i++)
      {
        var a = ReportingInovcom.existenceFichier(trameflux[i]);
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
        ReportingInovcom.deleteFichier(table,i,callback);
      };
    },
    deleteFichier: function (table,nb,callback) {
      var tab= '';
      console.log(tab);
      const fs = require('fs');
      fs.writeFile(table[nb]+'.txt', tab, (err) => {
        var sql = "insert into trame (typologiedelademande) values ('k') ";
        ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
          if(err) return console.log(err);
          else return callback(null, true);        
                              })      
      });
    },
    
    deleteReportingHtp : function (table,nb,callback) {
      var sql = "delete from "+table[nb]+" ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return console.log(err); }
        return callback(null, true);
        });
    },
    deleteHtp : function (table,nb,callback) {
      var j;
      var i = parseInt(j);
      for(i=0;i<nb;i++)
      {
        ReportingInovcom.deleteReportingHtp(table,i,callback);
      };
    },
    insertnbInovcom : async function (path,callback) {
      const Excel = require('exceljs');
      const cmd=require('node-cmd');
      const newWorkbook = new Excel.Workbook();
      try{
      await newWorkbook.xlsx.readFile(path);
      // var feuille = newWorkbook.getWorksheet();
      var test = newWorkbook.worksheets;
      var essaie = test.length;
      console.log(essaie);
      var sql = " INSERT INTO nbinovcomtype5 (nb) VALUES ("+essaie+"); ";
      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    }
    catch
    {
      console.log('ko');
    }
      
    },
};

