/**
 * ReportingIndu.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */

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
};

