/**
 * ReportingInovcom.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
 module.exports = {
  attributes: {},
  insert2ko : function (table,nb,callback) {
    //var nbr = parseInt(nb);
    var sql = "insert into indufactstc VALUES (0, 0); ";
    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
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
  insert2 : function (table,nb,callback) {
    var nbr = parseInt(nb);
    var sql = "insert into "+table[nbr]+" VALUES (0, 0); ";
    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log("Une erreur supprooo?");
        return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
      });
  },
  insert2text : function (table,nb,callback) {
    var nbr = parseInt(nb);
    var sql = "insert into "+table[nbr]+" VALUES ('0', '0'); ";
    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log("Une erreur insert?");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
      });
  },
  insert3text : function (table,nb,callback) {
    var nbr = parseInt(nb);
    var sql = "insert into ribtpmep VALUES ('0', '0', '0'); ";
    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log("Une erreur supprooo?");
        return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
      });
  },
  importTrameFlux929type2 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
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
      tab = ReportingInovcom.lectureEtInsertiontype2( trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
      var nbe= parseInt(nb);
      console.log(tab);
      var sql = "insert into "+table[nbe]+" (okko) values ('"+tab[0]+"') ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
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
  lectureEtInsertion4:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
  var workbook = XLSX.readFile(trameflux[nb]);
  var numerofeuille = feuil[nb];
  var numeroligne = parseInt(numligne[nb]);
  try{
    var nbr = 0;
    var nbrko = 0;
    var nbrtsisy = 0;
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col ;
    var nbe = parseInt(nb);
    for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(desired_value==cellule[nb])
        {
          col=ra;
        };
      };
      console.log("colonne"+col);
      
   if(col!=undefined)
    {
      console.log('tafa');

      var debutligne = numeroligne + 1;
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:col, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.v : undefined);
          if(desired_value1!=undefined)
          {
            if(desired_value1=='Saisir rib PS' || desired_value1=='Saisir envoyer convention au PS' || desired_value1=="Saisir mise à jour informations PS")
            {
              nbr=nbr + 1;
            }
            else if(desired_value1=='Saisir dossier de conventionnement PS')
            {
              nbrko =nbrko + 1;
            }
            else
            {
              var a = 1;
            };

          }
         else
         {
           var b =4;
         }
          
        };
          
    }
    else
    {
      console.log('Colonne non trouvé');
    };
    console.log("nombreeeeebr"+ nbr + 'et' + nbrko );
    var tab = [nbr,nbrko];
    return tab;
  }
  catch
  {
    console.log("erreur absolu haaha");
  };
  },
  lectureEtInsertion3:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
  var workbook = XLSX.readFile(trameflux[nb]);
  var numerofeuille = feuil[nb];
  var numeroligne = parseInt(numligne[nb]);
  try{
    var nbr = 0;
    var nbrko = 0;
    var nbrokrib = 0;
    var nbrtsisy = 0;
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col ;
    //var col = 16;
    var nbe = parseInt(nb);
    for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(desired_value==cellule[nb])
        {
          col=ra;
        };
      };
      console.log("colonne"+col );
      
   if(col!=undefined )
    {
      var debutligne = numeroligne + 1;
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:col, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.v : undefined);
          if(desired_value1=='facture saisie ce jour' || desired_value1=='Facture réglée')
          {
            nbr=nbr + 1;
          }
          else
          {
            if(desired_value1!=undefined)
            {
              nbrko =nbrko + 1;
            }
          };
        };  
    }
    else
    {
      console.log('Colonne non trouvé');
    };
    console.log("nombreeeeebr"+ nbr + 'et' + nbrko);
    var tab = [nbr,nbrko];
    return tab;
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
  var numeroligne = parseInt(numligne[nb]);
  try{
    var nbr = 0;
    var nbrko = 0;
    var nbrokrib = 0;
    var nbrtsisy = 0;
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col ;
    //var col = 16;
    var nbe = parseInt(nb);
    for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(desired_value==cellule[nb])
        {
          col=ra;
        };
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
        };
      };
      console.log("colonne"+col + 'et' + col2);
      
   if(col!=undefined )
    {
      var debutligne = numeroligne + 1;
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:col, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.v : undefined);
          //console.log(desired_value1);
          if(desired_value1=='Facture saisie ' || desired_value1=='Facture réglée le' || desired_value1=="Demande d'accord" || desired_value1=="Demande d'accord (A contrôler)" )
          {
            nbr=nbr + 1;
          }
          else
          {
            if(desired_value1!=undefined)
            {
              nbrko =nbrko + 1;
            }
          };
        };  
    }
    if(col2!=undefined )
    {
      var debutligne = numeroligne + 1;
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:col2, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.v : undefined);
          if(desired_value1=='Saisie Rib et mise en pratique' || desired_value1=='Modification Rib')
          {
            nbrokrib=nbrokrib + 1;
          }
          else
          {
              nbrtsisy = 1;
          }
        };
    }
    else
    {
      console.log('Colonne non trouvé');
    };
    console.log("nombreeeeebr"+ nbr + 'et' + nbrko + 'et' + nbrokrib);
    var tab = [nbr,nbrko,nbrokrib];
    return tab;
  }
  catch
  {
    console.log("erreur absolu haaha");
  };
  },
/* type 4 */


importTrameFlux929type4 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
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
    tab = ReportingInovcom.lectureEtInsertiontype4(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
    var nbe= parseInt(nb);
    console.log(tab);
    var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+tab[0]+"','"+tab[1]+"')";
    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
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
  };
},
lectureEtInsertiontype4:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
  XLSX = require('xlsx');
  var workbook = XLSX.readFile(trameflux[nb]);
  var numerofeuille = feuil[nb];
  var numeroligne = parseInt(numligne[nb]);
  try{
    var nbr = 0;
    var nbrko = 0;
    var nbrtsisy = 0;
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col ;
    //var col = 16;
    var nbe = parseInt(nb);
    for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(desired_value==cellule[nb])
        {
          col=ra;
        };
      };
      console.log("colonne"+col);
   if(col!=undefined)
    {
      var debutligne = numeroligne + 1;
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:col, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.v : undefined);
          //console.log(desired_value1);

          var ok = 'OK';
          var ko = 'KO';
          const regex = new RegExp(ok,'i');
          const regex1 = new RegExp(ko,'i');
          if(regex.test(desired_value1))
          {
            nbr=nbr + 1;
          }
          else if(regex1.test(desired_value1))
          {
            nbrko=nbrko + 1;
          }
          else
          {
            nbrtsisy = 1;
          };
        };
       
       /*var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+nbr+"','"+nbrko+"') ";
                    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if(err) return console.log(err);
                      else return callback(null, true);        
                                          });*/
        
    }
    else
    {
      console.log('Colonne non trouvé');
    }
    console.log("nombreeeeebr"+ nbr);
        var tab = [nbr,nbrko];
        return tab;

  }
  catch
  {
    console.log("erreur absolu haaha");
  }
  
},

/* type fin 4 */
  lectureEtInsertiontype5:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[0]);
    var numerofeuille = feuil[nb];
    var numeroligne = numligne[0];
    console.log(trameflux[0]);
    console.log(numeroligne);
    console.log(numerofeuille);
    var nbr = 0;
    var nbrko = 0;
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
      if(col!=undefined)
      {
        for(var a=0;a<=range.e.r;a++)
          {
            var address_of_cell = {c:col, r:a};
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            var desired_value1 = (desired_cell ? desired_cell.v : undefined);
            if(desired_value1=='OK' || desired_value1=='ok')
            {
              nbr=nbr + 1;
            }
            if(desired_value1=='KO' || desired_value1=='ko')
            {
              nbrko=nbrko + 1;
            }
            else
            {
             var nbrtsisy = 1;
            };

            
          };
          console.log(nbr + 'et' + nbrko);
          var sql = "insert into fav (nbok,nbko) values ('"+nbr+"','"+nbrko+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve fav?");
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
        console.log('Colonne non trouvé');
      }
      
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
    
  },
  lectureEtInsertiontype8:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = parseInt(numligne[nb]);
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
              var amp = 8;
            }
          };
      }

      else
      {
        console.log('Colonne non trouvé');
      }
      console.log("nombreeeeebr"+ nbr);
      var tab = [nbr];
      return tab;
      /*var sql = "insert into "+table[nb]+" (typologiedelademande,okko) values ('"+nbr+"','"+nbr+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err) return console.log(err);
                        else return callback(null, true);        
                                            });*/
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
      var essaie = parseInt(test.length) - 1;
      for(var y=0;y<essaie;y++) //parcours anle dossier rehetra
      {
        /*var j = parseInt(tab[y]);*/
        console.log(y);
        ReportingInovcom.lectureEtInsertiontype5(trameflux,feuil,cellule,table,cellule2,y,numligne,callback);
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
            var sql = "insert into retourcmuc (nbok,nbko) values ('"+ok+"','"+ko+"') ";
                    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log(err);
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
    console.log('ito le numerofeuille');
    console.log(numerofeuille);
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
        
        var mc1 = dernierl[nb];
        const regex = new RegExp(mc1,'i');
        if(regex.test(desired_value))
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
        var mc1 = cellule[nb];
        const regex = new RegExp(mc1,'i');
        if(regex.test(desired_value))
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
        var mc1 = cellule2[nb];
        const regex = new RegExp(mc1,'i');
        if(regex.test(desired_value))
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
            const regex = new RegExp(bi,'i');
            if(regex.test(desired_value1))
            {
              var z = parseInt(a) - 1;
              var address_of_cell2 = {c:22, r:z};
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
                var mm =1;
                //console.log('ko');
              };

              if(desired_value21!=undefined)
              {
                taboki = taboki +1;
                //console.log('ok2');
              }
              else
              {
                var mm =1;
                //console.log('ko2');
              };

            };
          };
          console.log('nb =' + tabok);
          console.log('nb2 =' + taboki);
          var sql = "insert into hospidematrejetprive (typologiedelademande,okko) values ('"+tabok+"','"+taboki+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve hospidematrejetprive?");
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
        console.log('Colonne non trouvé');
      };
      
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
  },
  importTrameFlux929type8 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
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
      tab = ReportingInovcom.lectureEtInsertiontype8(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
      var sql = "insert into "+table[nb]+" (typologiedelademande,okko) values ('"+tab[0]+"','"+tab[0]+"') ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
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
    };
  },
  importTrameFlux929 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
    console.log(table[nb]);
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
    else
    {
      var nbe= parseInt(nb);
      var tab = [];
      if(table[nbe]=="ribtpmep")
      {
        console.log('ribtpmep');
        tab = ReportingInovcom.lectureEtInsertion2(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
        console.log(tab);
        var sql = "insert into "+table[nbe]+" (nbok,nbko,nbrokrib) values ('"+tab[0]+"','"+tab[1]+"','"+tab[2]+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve insertion?");
                        //return callback(err);
                       }
                      else
                      {
                        console.log(sql);
                        return callback(null, true);
                      };     
                                          });
      }
      else if(table[nbe]=="curethermale")
      {
        console.log('cure');
        tab = ReportingInovcom.lectureEtInsertion3(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
        console.log(tab);
        var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+tab[0]+"','"+tab[1]+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve insertion?");
                        //return callback(err);
                       }
                      else
                      {
                        console.log(sql);
                        return callback(null, true);
                      };     
                                          });
      }
      else if(table[nbe]=="retourconventionsaisiedesconventions")
      {
        console.log('conv');
        tab = ReportingInovcom.lectureEtInsertion4(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
        console.log(tab);
        var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+tab[0]+"','"+tab[1]+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve insertion?");
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
        var sql = "delete from chemintsisy ";
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
    };
  };
  },
  importTrameFlux929type3 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
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
      tab = ReportingInovcom.lectureEtInsertiontype3( trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
      var nbe= parseInt(nb);
      console.log(tab);
      var sql = "insert into "+table[nbe]+" (okko) values ('"+tab[0]+"') ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
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
    };
  },
  deletealmerys : function (table,callback) {
    var sql = "delete from retouravisannulationcbtp ";
    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        //return callback(err); 
        console.log('une erreur de suppression');
      }
      return callback(null, true);
      });
  },
  deletecbtp : function (table,callback) {
    var sql = "delete from retouravisannulationtramealmerys ";
    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        //return callback(err); 
        console.log('une erreur de suppression');
      }
      return callback(null, true);
      });
  },
 deleteFromChemin : function (table,callback) {
      var sql = "delete from "+table+" ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { 
          console.log('une erreur de suppression');
          //return callback(err); 
        }
        return callback(null, true);
        });
    },
    deleteFromChemin2 : function (table,callback) {
      var sql = "delete from chemininovcomtype2 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) {
           //return callback(err); 
           console.log('une erreur de suppression');
          }
        return callback(null, true);
        });
    },
    deleteFromChemin3 : function (table,callback) {
      var sql = "delete from chemininovcomtype3 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { 
          //return callback(err); 
          console.log('une erreur de suppression');
        }
        return callback(null, true);
        });
    },
    deleteFromChemin4 : function (table,callback) {
      var sql = "delete from chemininovcomtype4";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) {
          // return callback(err); 
          console.log('une erreur de suppression');
        }
        return callback(null, true);
        });
    },
    deleteFromChemin5 : function (table,callback) {
      var sql = "delete from chemininovcomtype5 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { 
          console.log('une erreur de suppression');
          //return callback(err); 
        }
        return callback(null, true);
        });
    },
    deleteFromChemin6 : function (table,callback) {
      var sql = "delete from chemininovcomtype6 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { 
          //return callback(err); 
          console.log('une erreur de suppression');
        }
        return callback(null, true);
        });
    },
    deleteFromChemin7 : function (table,callback) {
      var sql = "delete from chemininovcomtype7 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { 
          //return callback(err); 
          console.log('une erreur de suppression');
        }
        return callback(null, true);
        });
    },
    deleteFromChemin8 : function (table,callback) {
      var sql = "delete from chemininovcomtype8 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { 
          //return callback(err); 
          console.log('une erreur de suppression');
        }
        return callback(null, true);
        });
    },
    deleteFromChemin9 : function (table,callback) {
      var sql = "delete from chemininovcomtype9 ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { 
          //return callback(err); 
          console.log('une erreur de suppression');
        }
        return callback(null, true);
        });
    },
    deletetype9 : function (table,callback) {
      var sql = "delete from recherchefactureinteriale ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { 
          //return callback(err); 
          console.log('une erreur de suppression');
        }
        return callback(null, true);
        });
    },
    existenceRoute : function (trameflux,callback) {
      var sql= "select count(chemin) as ok from "+trameflux+" ";
      Reportinghtp.getDatastore().sendNativeQuery(sql ,function(err, nc) {
            if (err){
              //console.log(err);
              //return callback(err);
              console.log('erreur trouvé dans existenceRoute');
            }
            else
            {
              return callback(null, nc);
            };
        });
    },
   
    existenceFichier : function (pathparam) {
      const fs = require('fs');
  
        var existe ='vrai';
        try{
          fs.accessSync(pathparam, fs.constants.F_OK);
        
        }catch(e){
          //console.log(e);
          existe = 'faux';
          console.log('chemin diso');
        }
        return existe;
    },

    importEssaitype9: function (table,table2,date,option,nb,callback) {
      const fs = require('fs');
      var re  = 'a';
      //var a = '\\\\10.128.1.2\\almerys-out\\Retour_Easytech_20210428\\RETOUR_RECHERCHE FACTURE INTERIALE\\INTERIALE';
      var a = table[0]+date+table2[nb];
      console.log('ch' +a);
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
              if (err) { 
                //console.log(err);
                console.log('une erreur trouvé');
                //return callback(err);
               }
              else
              {
                console.log(sql);
                return callback(null, true);
              };        
                                  }); 
          });
      }
      else
      {
        var sql = "insert into chemintsisy (typologiedelademande,okko) values (0,0)";
        ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
          if (err) { 
            //console.log(err);
            console.log('une erreur trouvé');
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
    importEssai: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomBase,chemin,option2,callback) {
      const fs = require('fs');
      var re  = 'a';
      var tab = [];
      var a = table[0]+date+table2[nb];
      var a1 = table[0]+date+chemin[nb];
      var b = option[nb];
      var b2 = option2[nb];
      console.log(b);
      var c = ReportingInovcom.existenceFichier(a);
      var d = ReportingInovcom.existenceFichier(a1);
      console.log(c);
      if(c=='vrai')
      {
        var nomCol = nomcolonne[nb].replace("'", "''"); 
        var nomCol2 = nomcolonne2[nb].replace("'", "''"); 
        var p = a.replace("'", "''"); 
        fs.readdir(a, (err, files) => {
          console.log(a);
              files.forEach(file => {
                var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
                var m2 = '^[^~]';
                const regex1 = new RegExp(m1,'i');
                const regex2 = new RegExp(m2);
                const regex = new RegExp(b,'i');
                const regex4 = new RegExp(b2,'i');
                console.log(b);
                if((regex.test(file) || regex4.test(file))  && regex1.test(file) && regex2.test(file))
                {
                  console.log(file);
                   var file1 = file.replace("'", "''");
                   //re = p + '\\' + file1;
                   re = p + '/' + file1;
                   //re=re.replace("'", "''");
                   console.log('ato'+re);
                   var sql = "insert into "+nomBase+" (chemin,nomtable,numligne,numfeuile,colonnecible,colonnecible2) values ('"+re+"','"+nomtable[nb]+"','"+numligne[nb]+"','"+numfeuille[nb]+"','"+nomCol+"','"+nomCol2+"') ";
                    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        //console.log(err);
                        console.log('une erreur trouvé');
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
                 };
            
                                          }); 
          });
      }
      else if(d=='vrai')
      {
        var nomCol = nomcolonne[nb].replace("'", "''"); 
        var nomCol2 = nomcolonne2[nb].replace("'", "''"); 
        var p = a1.replace("'", "''"); 
        fs.readdir(a1, (err, files) => {
          console.log(a1);
              files.forEach(file => {
                var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
                var m2 = '^[^~]';
                const regex1 = new RegExp(m1,'i');
                const regex2 = new RegExp(m2);
                const regex = new RegExp(b,'i');
                const regex4 = new RegExp(b2,'i');
                console.log(b);
                if((regex.test(file) || regex4.test(file)) && regex1.test(file) && regex2.test(file))
                {
                  console.log(file);
                   var file1 = file.replace("'", "''");
                   //re = p + '\\' + file1;
                   re = p + '/' + file1;

                   //re=re.replace("'", "''");
                   console.log('ato'+re);
                   var sql = "insert into "+nomBase+" (chemin,nomtable,numligne,numfeuile,colonnecible,colonnecible2) values ('"+re+"','"+nomtable[nb]+"','"+numligne[nb]+"','"+numfeuille[nb]+"','"+nomCol+"','"+nomCol2+"') ";
                    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        //console.log(err);
                        console.log('une erreur trouvé');
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
                 };
            
                                          }); 
          });
      }
      else
      {
        var sql = "insert into chemintsisy (typologiedelademande) values ('k') ";
        ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
          if(err) return console.log(err);
          else return callback(null, true);        
                              });   
      };
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
               var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
               var m2 = '^[^~]';
               const regex1 = new RegExp(m1,'i');
               const regex2 = new RegExp(m2);
               if(regex.test(file) && regex1.test(file) && regex2.test(file))
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
               var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
               var m2 = '^[^~]';
               const regex1 = new RegExp(m1,'i');
               const regex2 = new RegExp(m2);
               if(regex.test(file) && regex1.test(file) && regex2.test(file))
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
 importEssaitype4: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,nomBase,chemin,option2,callback) {
  const fs = require('fs');
  var re  = 'a';
  var a = table[0]+date+table2[nb];
  var b = option[nb];
  var nomTable = nomtable[nb];
  var numLigne= numligne[nb];
  var numFeuille = numfeuille[nb];
  var nomColonne = nomcolonne[nb];
  var c = ReportingInovcom.existenceFichier(a);
  var a1 = table[0]+date+chemin[nb];
  var b2 = option2[nb];
  var d = ReportingInovcom.existenceFichier(a1);
  console.log(c);
  if(c=='vrai')
  {
    fs.readdir(a, (err, files) => {
      console.log(a);
          files.forEach(file => {
            var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
            var m2 = '^[^~]';
            const regex = new RegExp(b,'i');
            const regex1 = new RegExp(m1,'i');
            const regex2 = new RegExp(m2);
            const regex4 = new RegExp(b2,'i');
            if((regex.test(file)  || regex4.test(file)) && regex1.test(file) && regex2.test(file))
            {
               //re = a+'\\'+file;
               re = a+'/'+file;
               var sql = "insert into "+nomBase+" (chemin,nomtable,numligne,numfeuile,colonnecible) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"') ";
               ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 4");
                  //return callback(err);
                 }
                else
                {
                  console.log(sql);
                  return callback(null, true);
                };          
                                     }) ;    
            } 
            else
            {
             var sql = "delete from chemintsisy ";
               ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log(sql);
                  return callback(null, true);
                };       
                                     }) ;
            };
           
           
        });
      });
  }
  else if(d=='vrai')
  {
    fs.readdir(a1, (err, files) => {
      console.log(a1);
          files.forEach(file => {
            var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
            var m2 = '^[^~]';
            const regex = new RegExp(b,'i');
            const regex1 = new RegExp(m1,'i');
            const regex2 = new RegExp(m2);
            const regex4 = new RegExp(b2,'i');
            if((regex.test(file)  || regex4.test(file)) && regex1.test(file) && regex2.test(file))
            {
               //re = a1+'\\'+file;
               re = a+'/'+file;
               var sql = "insert into "+nomBase+" (chemin,nomtable,numligne,numfeuile,colonnecible) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"') ";
               ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 4");
                  //return callback(err);
                 }
                else
                {
                  console.log(sql);
                  return callback(null, true);
                };          
                                     }) ;    
            } 
            else
            {
             var sql = "delete from chemintsisy ";
               ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log(sql);
                  return callback(null, true);
                };       
                                     }) ;
            };
           
           
        });
      });
 
  }
  else
  {
    var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log("Une erreur ve? import 1");
        //return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
                          }) ;  
  };   
},
  
 importEssaitype5: function (table,table2,date,option,nb,chem2,option2,callback) {
  const fs = require('fs');
 var re  = 'a';
 var tab = [];
 var a = table[0]+date+table2[nb];
 var a1 = table[0]+date+chem2[nb];
 var b = option[nb];
 var b1 = option2[nb];
 var c = ReportingInovcom.existenceFichier(a);
 var d = ReportingInovcom.existenceFichier(a1);
 console.log(c);
 if(c=='vrai')
 {
   fs.readdir(a, (err, files) => {
     console.log(a);
         files.forEach(file => {
          const regex = new RegExp(b,'i');
          const regex4 = new RegExp(b1,'i');
          var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
          var m2 = '^[^~]';
          const regex1 = new RegExp(m1,'i');
          const regex2 = new RegExp(m2);
          if( (regex.test(file) || regex4.test(file)) && regex1.test(file) && regex2.test(file))
           {
              re = file;
              console.log(re);  
           } 
       });
       var sql = "insert into chemininovcomtype5 (typologiedelademande) values ('"+re+"') ";
       ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log(err);
          //return callback(err);
         }
        else
        {
          console.log(sql);
          return callback(null, true);
        }; 
                            });
      console.log('ato anatiny'+re);
       
      
     });
 }
 else if(d=='vrai')
 {
   fs.readdir(a1, (err, files) => {
     console.log(a1);
         files.forEach(file => {
          const regex = new RegExp(b,'i');
          const regex4 = new RegExp(b1,'i');
          var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
          var m2 = '^[^~]';
          const regex1 = new RegExp(m1,'i');
          const regex2 = new RegExp(m2);
          if( (regex.test(file) || regex4.test(file)) && regex1.test(file) && regex2.test(file))
           {
              re = file;
              console.log(re);  
           } 
       });
       var sql = "insert into chemininovcomtype5 (typologiedelademande) values ('"+re+"') ";
       ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log(err);
          //return callback(err);
         }
        else
        {
          console.log(sql);
          return callback(null, true);
        }; 
                            });
      console.log('ato anatiny'+re);
       
      
     });
 }
 else
 {
  var sql = "delete from chemintsisy ";
  ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
   if (err) { 
     console.log("Une erreur ve? import 1");
     //return callback(err);
    }
   else
   {
     console.log(sql);
     return callback(null, true);
   };       
                        }) ;
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
          var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
          var m2 = '^[^~]';
          const regex1 = new RegExp(m1,'i');
          const regex2 = new RegExp(m2);
          if(regex.test(file) && regex1.test(file) && regex2.test(file))
           {
              re = file;
              console.log(re);  
           } 
         
          
       });
       var sql = "insert into chemininovcomtype6 (typologiedelademande) values ('"+re+"') ";
       ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log("Une erreur ve? import 6");
          //return callback(err);
         }
        else
        {
          console.log(sql);
          return callback(null, true);
        };          
                            }) ; 
      console.log('ato anatiny'+re);
       
      
     });
 }
 else
 {
   var sql = "delete from chemintsisy ";
   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
    if (err) { 
      console.log("Une erreur ve? import 1");
      //return callback(err);
     }
    else
    {
      console.log(sql);
      return callback(null, true);
    };       
                         }) ;
 }   
},
importEssaitype7: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,chem2,option2,callback) {
  const fs = require('fs');
  var re  = 'a';
  var tab = [];
  var ab = table[0]+date+table2[nb];
  var b = option[nb];
  var ab1 = table[0]+date+chem2[nb];
  var b1 = option2[nb];
  
  console.log('ch1' + ab);
  console.log('ch2' + ab1);
  
  var c = ReportingInovcom.existenceFichier(ab);
  var d = ReportingInovcom.existenceFichier(ab1);
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
                   const regex = new RegExp(b,'i');
                   const regex4 = new RegExp(b1,'i');
                   for(var i = 0; i < files1.length; i++){
                   
                    var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
                    var m2 = '^[^~]';
                    const regex1 = new RegExp(m1,'i');
                    const regex2 = new RegExp(m2);
                    
                     if( ( regex.test(files1[i]) || regex4.test(files1[i]) ) && regex1.test(files1[i]) && regex2.test(files1[i]))
                     {
                       //var a =ab + file +"\\" + files1[i];
                       var a =ab + file +"/" + files1[i];
                       console.log('*****************');
                       console.log(a);  
                       var sql = "insert into chemininovcomtype7 (chemin,nomtable,numligne,numfeuile,colonnecible,colonnecible2,colonnecible3) values ('"+a+"','"+nomtable+"','"+numligne+"','"+numfeuille+"','"+nomcolonne+"','"+nomcolonne2+"','"+nomcolonne3+"') ";
                       ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err)
                        {
                          console.log('une erreur');
                          //console.log(err);
                        }
                        else 
                        {
                          return callback(null, true); 
                        }          
                                             });
                     } 
                     else
                     {
                       var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
                       ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err)
                        {
                          console.log('une erreur');
                          //console.log(err);
                        }
                        else 
                        {
                          console.log(sql);
                          return callback(null, true); 
                        }          
                                             });
                     };
                   };
                 }
               });
             }
             
         }
 
       });
     };
  });
 }
 else if(d=='vrai')
  {
   fs.readdir(ab1, (err, files) => {
     if(err){
       console.log('ito le erreur : '+err);
     }
     else{
       var a;
       files.forEach(file =>{
         for(var i = 0; i < files.length; i++){
               if(file == files[i]){
               const test1 = ab1 +files[i];
               fs.readdir(test1, (err, files1) => {
                 if(err){
                   console.log(err);
                 }
                 else{
                   //console.log(file +" " +  files1[files1.length-1]);
                   //var cible = "MASQUE SAISIE";
                   const regex = new RegExp(b,'i');
                   const regex4 = new RegExp(b1,'i');
                   for(var i = 0; i < files1.length; i++){
                   
                    var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
                    var m2 = '^[^~]';
                    const regex1 = new RegExp(m1,'i');
                    const regex2 = new RegExp(m2);
                    
                     if( ( regex.test(files1[i]) || regex4.test(files1[i]) ) && regex1.test(files1[i]) && regex2.test(files1[i]))
                     {
                       //var a =ab1 + file +"\\" + files1[i];
                       var a =ab + file +"/" + files1[i];
                       console.log('*****************');
                       console.log(a);  
                       var sql = "insert into chemininovcomtype7 (chemin,nomtable,numligne,numfeuile,colonnecible,colonnecible2,colonnecible3) values ('"+a+"','"+nomtable+"','"+numligne+"','"+numfeuille+"','"+nomcolonne+"','"+nomcolonne2+"','"+nomcolonne3+"') ";
                       ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                         if(err)
                         {
                           console.log('une erreur');
                           //console.log(err);
                         }
                         else 
                         {
                           console.log(sql);
                           return callback(null, true); 
                         }       
                                             });
                     } 
                     else
                     {
                       var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
                       ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err)
                        {
                          console.log('une erreur');
                          //console.log(err);
                        }
                        else 
                        {
                          console.log(sql);
                          return callback(null, true); 
                        }          
                                             });
                     };
                   };
                 }
               });
             }
             
         }
 
       });
     };
  });
 }
  else
  {
    var sql = "insert into chemintsisy (typologiedelademande) values ('k') ";
    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err)
                        {
                          console.log('une erreur');
                          //console.log(err);
                        }
                        else 
                        {
                          console.log(sql);
                          return callback(null, true); 
                        }             
                          });  
  }   
 },
importEssaitype8: function (table,table2,date,option,nb,type,type2,nomtable,numligne,numfeuille,nomcolonne,callback) {
  const fs = require('fs');
  var re  = 'a';
  var tab = [];
  var chemin = table[0]+date;
  var ab = table[0]+date+table2[nb];
  console.log('chemin'+ab);
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
              console.log('chemin2' + test1);
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
                    var a = chemin+ "/RETOUR_HOSPI/RETOUR_AVIS D''ANNULATION/" + file  + type2[nb] + files1[files1.length-1];  
                    //var a = chemin+ "\\RETOUR_HOSPI\\RETOUR_AVIS D''ANNULATION\\" + file  + type2[nb] + files1[files1.length-1];  
                    console.log("haha"+a);
                    var sql = "insert into chemininovcomtype8 (chemin,nomtable,numligne,numfeuile,colonnecible) values ('"+a+"','"+nomtable+"','"+numligne+"','"+numfeuille+"','"+nomcolonne+"') ";
                    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve oui?");
                        //return callback(err);
                       }
                      else
                      {
                        console.log(sql);
                        return callback(null, true);
                      };    
                      /*if(err) return console.log(err);
                      else return callback(null, true);   */     
                                          }) ;
    
                  }
                 
                })
              }
              else
              {
                var sql = "insert into chemintsisy (typologiedelademande) values ('k') ";
                ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                  if (err) { 
                    console.log("Une erreur ve oui?");
                    //return callback(err);
                   }
                  else
                  {
                    console.log(sql);
                    return callback(null, true);
                  };    
                  /*if(err) return console.log(err);
                  else return callback(null, true);  */      
                                      }) ;  
              }
             
            }
            else
            {
             
            }
            
        };

      })
    }
 });
}
 else
 {
   var sql = "insert into chemintsisy (typologiedelademande) values ('k') ";
   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
    if (err) { 
      console.log("Une erreur ve oui?");
      //return callback(err);
     }
    else
    {
      console.log(sql);
      return callback(null, true);
    };    
     /*if(err) return console.log(err);
     else return callback(null, true); */       
                         });  
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
    delete : function (table,nb,callback) {
      var nbr = parseInt(nb);
      var sql = "delete from "+table[nbr]+" ";
      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
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
    lectureEtInsertiontype3:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
      XLSX = require('xlsx');
      var workbook = XLSX.readFile(trameflux[nb]);
      var numerofeuille = feuil[nb];
      var numeroligne = parseInt(numligne[nb]);
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
        /*var sql = "insert into "+table[nb]+" (typologiedelademande,okko) values ('"+nbr+"','"+nbr+"') ";
                        ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                          if(err) return console.log(err);
                          else return callback(null, true);        
                                              })*/
          var tab = [nbr];
          return tab;
      }
      catch
      {
        console.log("erreur absolu haaha");
      }
      
    },
};

