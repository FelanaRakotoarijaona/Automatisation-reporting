/**
 * ReportingRetour.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */

const { parse } = require('path');

module.exports = {

  attributes: {
  },
  lectureEtInsertiontype22:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
   XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = parseInt(numligne[nb]);
    console.log(trameflux[nb] + numerofeuille + numerofeuille + 'hihi');
    try{
      var nbrok = 0;
      var nbrko = 0;
      const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var col = 0;
      var nbe = parseInt(nb);
      if(col!=undefined)
      {
        var debutligne = numeroligne + 1;
        for(var a=debutligne;a<=range.e.r;a++)
          {
            var address_of_cell = {c:1, r:a};
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            var desired_value1 = (desired_cell ? desired_cell.w : undefined);

            var address_of_cell1 = {c:23, r:a};
            var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
            var desired_cell1 = sheet[cell_ref1];
            var desired_value2 = (desired_cell1 ? desired_cell1.w : undefined);

            //console.log('mba ato ar ve e');
            //console.log(desired_value1 + desired_value2);
            if(desired_value2!=undefined && (desired_value1<desired_value2))
            {
              nbrok=nbrok + 1;
              //console.log('aryy atoo');

            }
            else if(desired_value2==undefined && (desired_value1!=undefined) || (desired_value2!=undefined && (desired_value1>desired_value2)))
            {
              nbrko=nbrko + 1;
              //console.log('aryy atoo 2');
            }
            else
            {
              var s = 1;
            }
          };
          console.log("nombreeeeebr"+ nbrok + nbrko);
          var tab = [nbrok,nbrko];
          return tab;
         /* var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+nbrok+"','"+nbrko+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err)
                        {
                          console.log(err);
                        }
                        else
                        {
                          console.log(sql);
                          return callback(null, true);  
                        }       
                                            });
          console.log("nombreeeeebr"+ nbrok + 'h' + nbrko);*/
      }
      else
      {
        console.log('Colonne non trouvé');
      }
      /*var tab = ['0','5'];
      return tab;*/
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
    
  },
  //import du chemin dans le serveur
  importTrameFlux929type2 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
    if(trameflux[nb]==undefined)
    {
      console.log('trame undefined');
      var sql = "insert into chemintsisy(typologiedelademande) values ('ko') ";
      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log("Une erreur ve ok?");
          return callback(err);
         }
        else
        {
          console.log(sql);
          return callback(null, true);
        };
      });
    }
   else if(table[nb]=="coldrcbtppublic")
    {
      console.log('hehe coldrcbtppublic ato v oooooooo');
      var tab = [];
      tab = ReportingRetour.lectureEtInsertiontype22( trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
      console.log('tab' + tab);
      var sql = "insert into "+table[nb]+" (nbok,nbko) values ('"+tab[0]+"','"+tab[1]+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err)
                        {
                          console.log(err);
                        }
                        else
                        {
                          console.log(sql);
                          return callback(null, true);  
                        }       
                                            });
      
    }
    else if(table[nb]=="trpecaudio" || table[nb]=="trpecdentaire")
    {
      console.log('hehe trpecdentaire');
      var tab = [];
      tab=ReportingRetour.lectureEtInsertiontype21( trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
      var nbe= parseInt(nb);
      console.log(tab);
      var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
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
    }
    else{
      var tab = [];
      tab=ReportingRetour.lectureEtInsertiontype2( trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
      var nbe= parseInt(nb);
      console.log(tab);
      var sql = "insert into "+table[nbe]+" (nb) values ('"+tab[0]+"') ";
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
  //lecture du chemin
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
         
         /* var sql = "insert into "+table[nbe]+" (nb) values ('"+nbr+"') ";
                      ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if(err) return console.log(err);
                        else return callback(null, true);        
                                            });*/
          console.log("nombreeeeebr"+ nbr);
          var tab = [nbr];
          return tab;
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
   lectureEtInsertiontype21:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = parseInt(numligne[nb]);
    console.log(trameflux[nb]);
    console.log(numerofeuille);
    console.log(numeroligne);
    try{
      var nbr = 0;
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var col = 0;
      var colnonvide;
      var nbe = parseInt(nb);
      console.log('tafiditra ato v?');
      var bi = 'FIN DE TRAITEMENT';
      const regex = new RegExp(bi,'i');
      for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(regex.test(desired_value))
        {
          colnonvide=ra;
        }
        else
        {
          colnonvide = 13;
        }
        console.log('tafiditra ato v 2?');
      };
      console.log("colonnevide"+colnonvide);
      if(col!=undefined && colnonvide!=undefined)
      {
        console.log('tafiditra ato v? 3');
        var debutligne = numeroligne + 1;
        for(var a=debutligne;a<=range.e.r;a++)
          {
            var address_of_cell = {c:col, r:a};
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            var desired_value1 = (desired_cell ? desired_cell.v : undefined);

            var address_of_cell2 = {c:colnonvide, r:a};
            var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
            var desired_cell2 = sheet[cell_ref2];
            var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);

            console.log("valeur = " + desired_value2 + desired_value1);
            if(desired_value1!=undefined && desired_value2!=undefined )
            {
              nbr=nbr + 1;
            }
          };
          console.log("nombreeeeebr"+ nbr);
          var tab = [nbr];
          return tab;
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
  //effacement du chemin dans la base pour eviter le doublon
deleteFromChemin : function (table,callback) {
      var sql = "delete from cheminretourvrai ";
      ReportingRetour.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return callback(err); }
        return callback(null, true);
        });
    },
    //test existence du fichier
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
    //import
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
                var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
                var m2 = '^[^~]';
                const regex1 = new RegExp(m1,'i');
                const regex2 = new RegExp(m2);
                if(regex.test(file) && regex1.test(file) && regex2.test(file))
                {
                   re = a+'\\'+file;
                   //console.log(re);  
                   var sql = "insert into cheminretourvrai (chemin,nomtable,numligne,numfeuile,colonnecible) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"') ";
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
      ReportingRetour.getDatastore().sendNativeQuery(sql, function(err, res){
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
        var a = ReportingRetour.existenceFichier(trameflux[i]);
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
        ReportingRetour.deleteFichier(table,i,callback);
      };
    },
    deleteFichier: function (table,nb,callback) {
      var tab= '';
      console.log(tab);
      const fs = require('fs');
      fs.writeFile(table[nb]+'.txt', tab, (err) => {
        var sql = "insert into trame (typologiedelademande) values ('k') ";
        ReportingRetour.getDatastore().sendNativeQuery(sql, function(err,res){
          if(err) return console.log(err);
          else return callback(null, true);        
                              })      
      });
    },
    
    deleteReportingHtp : function (table,nb,callback) {
      var sql = "delete from "+table[nb]+" ";
      ReportingRetour.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { return console.log(err); }
        return callback(null, true);
        });
    },
    deleteHtp : function (table,nb,callback) {
      var j;
      var i = parseInt(j);
      for(i=0;i<nb;i++)
      {
        ReportingRetour.deleteReportingHtp(table,i,callback);
      };
    },
};

