/**
 * Garantie.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */

module.exports = {

  attributes: {

  },
  /*deleteFromChemin : function (table,callback) {
    var sql = "delete from chemingarantie ";
    console.log(sql);
    Garantie.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return callback(err); }
      return callback(null, true);
      });
  },
  // deleteFromChemin2 : function (table,callback) {
  //   var sql = "delete from cheminhtp2 ";
  //   Garantie.getDatastore().sendNativeQuery(sql, function(err, res){
  //     if (err) { return callback(err); }
  //     return callback(null, true);
  //     });
  // },

  importEssaidemat: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,nomcolonne3,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    var ab = table[0]+date+table2[nb];
    var b = option[nb];
    console.log(ab);
    var c = Garantie.existenceFichier(ab);
    console.log(c);
    if(c=='vrai')
    {
     fs.readdir(ab, (err, files) => {
       if(err){
         console.log('ito le erreur : '+err);
       }
       else{
         console.log('tena ts erruer kay lait');
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
                     
                      var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
                      var m2 = '^[^~]';
                      const regex1 = new RegExp(m1,'i');
                      const regex2 = new RegExp(m2);
                      
                       if(regex.test(files1[i]) && regex1.test(files1[i]) && regex2.test(files1[i]))
                       {
                         var a =ab + file +"\\" + files1[i]; 
                         var sql = "insert into chemingarantie (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+a+"','"+nomtable+"','"+numligne+"','"+numfeuille+"','"+nomcolonne+"','"+nomcolonne2+"','"+nomcolonne3+"') ";
                        console.log(sql);
                         Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                           if(err) 
                           {
                             console.log(err);
                           }
                           else {
                            return callback(null, true);
                           }        
                                               });
                       } 
                       else
                       {
                         var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
                         Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                          if(err) 
                           {
                             console.log("erreur");
                           }
                           else {
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
      Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
        if(err) 
        {
          console.log("erreur");
        }
        else return callback(null, true);        
      })   
    }   
   },
/************************************************************************/

 /* importEssaidemat1: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    var a = table[0]+date+table2[nb];
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
    console.log(c);    
    if(c=='vrai')
    {
      fs.readdir(a, (err, files) => {
        // console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              if(regex.test(file))
              {
                 //re = a+'\\'+file;
                 re = a+''+file;
                 console.log('ato le sql mipoitra am rcindeterminable');
                 var sql = "insert into chemingarantie (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"','"+colonnecible2+"','"+colonnecible3+"') ";
                //  console.log(sql);
                 Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
                  if (err) { 
                    console.log("Une erreur ve? import 1");
                    //return callback(err);
                   }
                  else
                  {
                    console.log('ito kay le ataony an :'+sql);
                    return callback(null, true);
                  };
                   
              });
             }
              else
              {
               var sql = "insert into chemintsisy (typologiedelademande) values ('"+re+"') ";
               Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log('ito le tsy tokony ho ataony '+sql);
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
      Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
  /*******************************************************/
 /* importEssaidematbpo: function (table_1,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,callback) {
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
    console.log('ccccccccccccccccccccccccccccccc');
    console.log(c);    
    if(c=='vrai')
    {
      console.log('fidirana voalohany');
      fs.readdir(a, (err, files) => {
        console.log('mbola fidirana voalohany ihany');
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              if(regex.test(file))
              {
                 //re = a+'\\'+file;
                 re = a+''+file;
                 console.log('fidirana faharoa');
                 var sql = "insert into chemingarantie (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"','"+colonnecible2+"','"+colonnecible3+"') ";
                //  console.log(sql);
                 Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
               Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
      Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
  /*importEssaiavis: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,Sup,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    // var a = table[0]+date+table2[nb];
    var chem = table[0]+date+"\\RETOUR_HOSPI\\RETOUR_AVIS D''ANNULATION\\";
    var a = table[0]+date+table2[nb]+date+Sup;
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
    console.log('ccccccccccccccccccccccccccccccc');
    console.log(c);    
    if(c=='vrai')
    {
      console.log('fidirana voalohany');
      fs.readdir(a, (err, files) => {
        console.log('mbola fidirana voalohany ihany');
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              if(regex.test(file))
              {
                console.log(file);
                 a = chem+date+Sup;
                 re = a+''+file;
                //  re = a+''+file;
                 console.log('fidirana faharoa');
                 console.log(re);
                 var sql = "insert into chemingarantie (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"','"+colonnecible2+"','"+colonnecible3+"') ";
                 console.log(sql);
                 Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
               Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log("ato v iz no manao console : "+sql);
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
      Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
 /* importEssaidematconvention: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    // var a = table[0]+date+table2[nb];
    var a = table[0]+date+table2[nb];
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
    console.log('ccccccccccccccccccccccccccccccc');
    console.log(c);    
    if(c=='vrai')
    {
      console.log('fidirana voalohany');
      fs.readdir(a, (err, files) => {
        console.log('mbola fidirana voalohany ihany');
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              if(regex.test(file))
              {
                 //re = a+'\\'+file;
                 re = a+''+file;
                 console.log('fidirana faharoa');
                 var sql = "insert into chemingarantie (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"','"+colonnecible2+"','"+colonnecible3+"') ";
                //  console.log(sql);
                 Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
               Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
      Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
 /* importEssaiencindus: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,Sup,date_indus,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    // var a = table[0]+date+table2[nb];
    var a = table[0]+date+table2[nb]+Sup+date_indus;
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
    console.log('ccccccccccccccccccccccccccccccc');
    console.log(c);    
    if(c=='vrai')
    {
      console.log('fidirana voalohany');
      fs.readdir(a, (err, files) => {
        console.log('mbola fidirana voalohany ihany');
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              if(regex.test(file))
              {
                console.log(file);
                 re = a+''+file;
                //  re = a+''+file;
                 console.log('fidirana faharoa');
                 console.log(re);
                 var sql = "insert into chemingarantie (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"','"+colonnecible2+"','"+colonnecible3+"') ";
                 console.log(sql);
                 Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
               Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
                if (err) { 
                  console.log("Une erreur ve? import 1");
                  //return callback(err);
                 }
                else
                {
                  console.log("ato v iz no manao console : "+sql);
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
      Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
 /* importEssaidematretention: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,colonnecible3,datej_1,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    // var a = table[0]+date+table2[nb];
    var a = table[0]+datej_1+table2[nb];
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
    console.log('ccccccccccccccccccccccccccccccc');
    console.log(c);    
    if(c=='vrai')
    {
      console.log('fidirana voalohany');
      fs.readdir(a, (err, files) => {
        console.log('mbola fidirana voalohany ihany');
        console.log(a);
            files.forEach(file => {
              const regex = new RegExp(b+'*');
              if(regex.test(file))
              {
                 //re = a+'\\'+file;
                 re = a+''+file;
                 console.log('fidirana faharoa');
                 var sql = "insert into chemingarantie (chemin,nomtable,numligne,numfeuille,colonnecible,colonnecible2,colonnecible3) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"','"+colonnecible2+"','"+colonnecible3+"') ";
                //  console.log(sql);
                 Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
               Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
      Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
  /*existenceFichier : function (pathparam) {
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
  totalFichierExistant : function (trameflux,nb,callback) {
    var tab = [];
    var j;
    var i = parseInt(j);
    for(i=0;i<nb;i++)
    {
      var a = Garantie.existenceFichier(trameflux[i]);
      console.log('ito le manome valeur an le tab');
      console.log(a);
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
      Garantie.deleteFichier(table,i,callback);
    };
  },
  deleteToutHtp : function (table,nb,callback) {
    var sql = "delete from "+table+" ";
    Garantie.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return callback(err); }
      return callback(null, true);
      });
  },
  deleteFichier: function (table,nb,callback) {
    var tab= '';
    console.log(tab);
    const fs = require('fs');
    fs.writeFile(table[nb]+'.txt', tab, (err) => {
      var sql = "insert into trame (typologiedelademande) values ('k') ";
      Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
        if(err) return console.log(err);
        else return callback(null, true);        
                            })      
    });
  },
  deleteReportingHtp : function (table,nb,callback) {
    var sql = "delete from "+table[nb]+" ";
    Garantie.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return console.log(err); }
      return callback(null, true);
      });
  },
  deleteHtp : function (table,nb,callback) {
    var j;
    var i = parseInt(j);
    for(i=0;i<nb;i++)
    {
      Garantie.deleteReportingHtp(table,i,callback);
    };
  },
/****************************************************************************************/
   /*importTrameDemat : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback) {
    // var tab = [];
    // tab = Garantie.totalFichierExistant(trameflux,nb,callback);
    // console.log(tab);
    // if(tab.length==0)
    // {
    //   console.log('Aucune reporting pour ce date');
    //   Garantie.deleteToutHtp(table,3,callback);
    // }
    // else{
    //   for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    // {
    //   var j = parseInt(tab[y]);
    //   console.log('jjjjjjjjjjjjjjjjjjjjjjj');
    //   console.log(j);
    //   Garantie.lectureEtInsertiongarantie( trameflux,feuil,cellule,table,cellule2,j,numligne,dernierl,callback)
    // // ReportingInovcom.lectureEtInsertion(trameflux,feuil,cellule,table,cellule2,j,callback)
    // }
    // };
    console.log('****************');
    console.log(nb);
    console.log(trameflux[0]);
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
      
      var tab1 = [];
      tab1 = Garantie.lectureEtInsertiongarantieRcindeterminable( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback);
      console.log(tab1);
      var sql = "insert into garantiercindeterminable (nb) values ('"+tab1[0]+"') ";
      console.log('mety aloha hatreto '+sql);

      var tab = [];
      tab = Garantie.lectureEtInsertiongarantie( trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback);
      var nbe= parseInt(nb);
      console.log(tab);
      var sql = "insert into garantiedemat (nb) values ('"+tab[0]+"') ";

      Garantie.getDatastore().sendNativeQuery(sql, function(err,res){
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
  /****************************************************************************/
 /* lectureEtInsertiongarantie:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[1]);
    var numerofeuille = feuil[1];
    var numeroligne = parseInt(numligne[1]);
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
        if(desired_value==dernierl[1])
        {
          col=ra;
        }
      };
      // for(var ra=0;ra<=range.e.c;ra++)
      // {
      //   var address_of_cell = {c:ra, r:numeroligne};
      //   var cell_ref = XLSX.utils.encode_cell(address_of_cell);
      //   var desired_cell = sheet[cell_ref];
      //   var desired_value = (desired_cell ? desired_cell.v : undefined);
      //   // console.log(desired_value);
      //   if(desired_value==cellule[nb])
      //   {
      //     col1=ra;
      //   }
      // };
      // for(var ra=0;ra<=range.e.c;ra++)
      // {
      //   var address_of_cell = {c:ra, r:numeroligne};
      //   var cell_ref = XLSX.utils.encode_cell(address_of_cell);
      //   var desired_cell = sheet[cell_ref];
      //   var desired_value = (desired_cell ? desired_cell.v : undefined);
      //   // console.log(desired_value);
      //   if(desired_value==cellule2[nb])
      //   {
      //     col2=ra;
      //   }
      // };
      // console.log('colonne cible' +col + col1 +col2);
      console.log('colonne cible : ' +col);
      // if(col!=undefined && col1!=undefined  && col2!=undefined) 
      if(col!=undefined)
      {
        var nbr = 0;
        // var tabok = 0;
        // var taboki = 0;
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
              // var z = parseInt(a) - 1;
              // var address_of_cell2 = {c:col1, r:z};
              // var cell_refs = XLSX.utils.encode_cell(address_of_cell2);
              // var desired_cell2 = sheet[cell_refs];
              // var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);

              // var address_of_cell21 = {c:col2, r:z};
              // var cell_refs1 = XLSX.utils.encode_cell(address_of_cell21);
              // var desired_cell21 = sheet[cell_refs1];
              // var desired_value21 = (desired_cell21 ? desired_cell21.v : undefined);

              // if(desired_value2!=undefined)
              // {
              //   tabok= tabok + 1;
              //   //console.log('ok');
              // }
              // else
              // {
              //   //console.log('ko');
              // };

              // if(desired_value21!=undefined)
              // {
              //   taboki = taboki +1;
              //   //console.log('ok2');
              // }
              // else
              // {
              //   //console.log('ko2');
              // };
              // tabok= tabok + 1;
              nbr=nbr + 1;
            };
          };
          // console.log('nb =' + tabok);
          // var data = tabok-2;
          // console.log('nb2 =' + taboki);
          // var sql = "insert into garantiedemat (typologiedelademande,okko) values ('"+tabok+"','"+taboki+"') ";
          // var sql = "insert into garantiedemat (nb) values ('"+data+"') ";
          //             ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
          //               if (err) { 
          //                 console.log("Une erreur ve garantiedemat ?");
          //                 //return callback(err);
          //                }
          //               else
          //               {
          //                 console.log(sql);
          //                 return callback(null, true);
          //               };          
          //                                  });
      }

      // else
      // {
      //   console.log('Colonne non trouvé');
      // };

      
      else
      {
        console.log('Colonne non trouvé');
      }
      var tab = [nbr-2];
          console.log("nombreeeeebr__quotidien__"+ nbr);
          return tab; 
      
    }
   
    catch
    {
      console.log("erreur absolu haaha");
    }


  },
  /***************************************************************************/
  /*importTrameRcindeterminable : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback) {
    var tab = [];
    tab = Garantie.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      Garantie.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log('j222222222222222222222');
      console.log(j);
      Garantie.lectureEtInsertiongarantieRcindeterminable( trameflux,feuil,cellule,table,cellule2,j,numligne,dernierl,callback)
    // ReportingInovcom.lectureEtInsertion(trameflux,feuil,cellule,table,cellule2,j,callback)
    }
    };
  },
  /***************************************************************************/
/*lectureEtInsertiongarantieRcindeterminable:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dernierl,callback){
  XLSX = require('xlsx');
  var workbook = XLSX.readFile(trameflux[0]);
  var numerofeuille = feuil[0];
  var numeroligne = parseInt(numligne[0]);
  console.log('lign ato am lecture et insertion : ' +numeroligne);
//   try{
//     console.log('miditra ato am try v iz?');
//     console.log(numerofeuille);
//     const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
//     var range = XLSX.utils.decode_range(sheet['!ref']);
//     var col ;
//     var col1;
//     var col2;
//     console.log('Nombre de colonne' + range.e.c);
//     console.log('Nombre de ligne' + range.e.r);
//     console.log(numeroligne + 'numlign');
//     console.log(numerofeuille + 'numfeuille');
//     console.log(cellule2[nb] + 'c2');
//     console.log(cellule[nb] + 'c1');
//     console.log(dernierl[nb] + 'c3');
//     console.log(table[nb] + 'table');
//     for(var ra=0;ra<=range.e.c;ra++)
//     {
//       var address_of_cell = {c:ra, r:numeroligne};
//       // console.log(address_of_cell);//c:5 r:0
//       var cell_ref = XLSX.utils.encode_cell(address_of_cell);
//       var desired_cell = sheet[cell_ref];
//       // console.log(desired_cell);
//       var desired_value = (desired_cell ? desired_cell.v : undefined);
//       console.log(desired_value);//Etat fin de traitement
//       if(desired_value==dernierl[nb])
//       {
//         col=ra;
//       }
//     };
//     console.log('colonne cible : ' +col);
   
//     if(col!=undefined)
//     {
//       var tabok = 0;
//       for(var a=0;a<=range.e.r;a++)
//         {
//           var address_of_cell = {c:col, r:a};
//           var cell_ref = XLSX.utils.encode_cell(address_of_cell);
//           var desired_cell = sheet[cell_ref];
//           var desired_value1 = (desired_cell ? desired_cell.v : undefined);
//           var bi = 'Total';
//           const regex = new RegExp(bi+'*');
//           if(regex.test(desired_value1))
//           {
//             tabok= tabok + 1;
//           };
//         };
//         console.log('nb =' + tabok);
//         var data = tabok-1;
//         var sql = "insert into garantiercindeterminable (nb) values ('"+data+"') ";
//                     ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
//                       if (err) { 
//                         console.log("Une erreur ve garantiercindeterminable ?");
                        
//                        }
//                       else
//                       {
//                         console.log(sql);
//                         return callback(null, true);
//                       };          
//                                          });
//     }
//     else
//     {
//       console.log('Colonne non trouvé');
//     };
    
//   }
//   catch
//   {
//     console.log("erreur absolu haaha");
//   }

try{
  var nbr = 0;
  const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
  var range = XLSX.utils.decode_range(sheet['!ref']);
  var col=0;
  var nbe = parseInt(0);
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
  var tab1 = [nbr];
      console.log("nombreeeeebr"+ nbr);
      return tab1;
}
catch
{
  console.log("erreur absolu haaha");
}

 },*/

};


