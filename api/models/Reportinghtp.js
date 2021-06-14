/**
 * Reportinghtp.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
module.exports = {
  attributes: {
  },
  importTrameFlux929type2 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {

    if(trameflux[nb]==undefined)
    {
      var sql = "insert into chemintsisy(typologiedelademande) values ('tsisy') ";
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
    else
    {
      var tab = [];
      tab = Reportinghtp.lectureEtInsertion2( trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
      var nbe = parseInt(nb);
      var sql = "insert into "+table[nbe]+" (nbok,nbko,nbokresiliation,nbkoresiliation) values ("+tab[0]+","+tab[1]+","+tab[2]+","+tab[3]+") ";
      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
        if (err) { 
          console.log("Une erreur ve type2?");
          return callback(err);
        }
        else
        {
          console.log(sql);
          return callback(null, true);
        };

      });
    };
  },
  lectureEtInsertion2:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    console.log('ok v');
   XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = parseInt(numligne[nb]);
    try{
      var nbr = 0;
      var nbrko = 0;
      var nbrokresiliation = 0;
      var nbrkoresiliation = 0;
      var nbrtsisy = 0;
      const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var col = 0;
      var col2 =0;
      var nbe = parseInt(nb);
      console.log(cellule[nb]);
      for(var ra=0;ra<=range.e.c;ra++)
        {
          var address_of_cell = {c:ra, r:numeroligne};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          console.log(desired_value);
          if(desired_value=='SAISIE OK/KO' || desired_value=='Saisie OK/KO' || desired_value=='saisie OK/KO' || desired_value=='saisie ok/ko' || desired_value=='OK/KO' || desired_value=='ETAT' || desired_value=='etat' || desired_value=='Etat' || desired_value=='ETAT OK/KO' || desired_value=='etat OK/KO' || desired_value=='Etat OK/KO' )
          {
            col=ra;
          };
       };
        console.log("colonne"+col);
       for(var ra=0;ra<=range.e.c;ra++)
        {
          var address_of_cell = {c:ra, r:numeroligne};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          if(desired_value=='Typologie de la demande')
          {
            col2=ra;
          };
        };
     if(col!=undefined && col2!=undefined)
      {
        var debutligne = numeroligne + 1;
        for(var a=debutligne;a<=range.e.r;a++)
          {
            var address_of_cell = {c:col, r:a};
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            var desired_value1 = (desired_cell ? desired_cell.v : undefined);

            var address_of_cell2 = {c:col2, r:a};
            var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
            var desired_cell2 = sheet[cell_ref2];
            var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
            if(desired_value1=='OK' && ( desired_value2=='Adhésion' ||desired_value2=='Espace adhérent'|| desired_value2=="Ajout d'Ayant droit" || desired_value2=='Autres' || desired_value2=='Carte TP' || desired_value2=='Cotisations' || desired_value2=='MAJ N1' || desired_value2=='Demande attestation' || desired_value2=='MAJ Régime / Caisse' || desired_value2=='MAJ RIB' || desired_value2=='Modification formule' || desired_value2=='Prestations' || desired_value2=='Prévoyance'  ) )
            {

              nbr=nbr + 1;
            }
            if(desired_value1=='KO' && ( desired_value2=='Adhésion' ||desired_value2=='Espace adhérent' || desired_value2=="Ajout d'Ayant droit" || desired_value2=='Autres' || desired_value2=='Carte TP' || desired_value2=='Cotisations' || desired_value2=='MAJ N1' || desired_value2=='Demande attestation' || desired_value2=='MAJ Régime / Caisse' || desired_value2=='MAJ RIB' || desired_value2=='Modification formule' || desired_value2=='Prestations' || desired_value2=='Prévoyance'  ))
            {
              nbrko=nbrko + 1;
            }
            if(desired_value1=='OK' && ( desired_value2=='Résiliation') )
            {

              nbrokresiliation=nbrokresiliation + 1;
            }
            if(desired_value1=='KO' && ( desired_value2=='Résiliation' ))
            {
              nbrkoresiliation=nbrkoresiliation + 1;
            }
            else
            {
              var nbrtsisy = 1;
            };
          };
         
          var tab = [nbr,nbrko,nbrokresiliation,nbrkoresiliation];
          console.log("nombreeeeebr"+ nbr);
          return tab;
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
  importTrameFlux929type4 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,dateexport,callback) {
    if(trameflux[nb]==undefined)
    {
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
      var tab = [];
      tab = Reportinghtp.lectureEtInsertiontype4( trameflux,feuil,cellule,table,cellule2,nb,numligne,dateexport,callback);
      var nbe= parseInt(nb);
      var sql = "insert into "+table[nbe]+" (nbok,nbko) values ("+tab[0]+","+tab[1]+") ";
                        Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                          if (err) { 
                            console.log("Une erreur ve?");
                            return callback(err);
                           }
                          else
                          {
                            console.log(sql);
                            return callback(null, true);
                          };
                         
                                              });
    };
    
    /*tab = Reportinghtp.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<nb;y++) //parcours anle dossier rehetra
      {
        var j = parseInt(y);
        console.log(j);
        Reportinghtp.lectureEtInsertiontype4( trameflux,feuil,cellule,table,cellule2,j,numligne,dateexport,callback);
      };
    };*/
  },
  lectureEtInsertiontype4:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,dateexport,callback){
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
      var col = 0;
      var col = 0;
      var nbe = parseInt(nb);
      console.log(cellule[nb]);
      for(var ra=0;ra<=range.e.c;ra++)
        {
          var address_of_cell = {c:ra, r:numeroligne};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          if(desired_value==cellule[nb] || desired_value=='SAISIE OK/KO' || desired_value=='Saisie OK/KO' || desired_value=='saisie OK/KO' || desired_value=='saisie ok/ko' || desired_value=='OK/KO' || desired_value=='ETAT' || desired_value=='etat' || desired_value=='Etat' || desired_value=='ETAT OK/KO' || desired_value=='etat OK/KO' || desired_value=='Etat OK/KO' )
          {
            col=ra;
          }
          else{
            var a=10;
            //console.log('colonne introuvable');
          };
        };
        for(var ra=0;ra<=range.e.c;ra++)
        {
          var address_of_cell = {c:ra, r:numeroligne};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          if(desired_value=='Date de traitement' || desired_value=='Date de traitemet')
          {
            col2=ra;
          }
          else{
            var a=10;
            //console.log('colonne introuvable');
          };
        };
        console.log("colonne"+col);
     if(col!=undefined && col2!=undefined )
      {
        var debutligne = numeroligne + 1;
        for(var a=debutligne;a<=range.e.r;a++)
          {
            var address_of_cell = {c:col, r:a};
            var cell_ref = XLSX.utils.encode_cell(address_of_cell);
            var desired_cell = sheet[cell_ref];
            var desired_value1 = (desired_cell ? desired_cell.v : undefined);

            var address_of_cell2 = {c:col2, r:a};
            var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
            var desired_cell2 = sheet[cell_ref2];
            var desired_value2 = (desired_cell2 ? desired_cell2.w : undefined); 

           var dateExcel = dateexport;
           if(desired_value2==dateExcel && (desired_value1=='OK' || desired_value1=='ok'))
            {
              nbr=nbr + 1;
            }
            else
            {
              nbrtsisy = 1;
            };
            if(desired_value2==dateExcel && (desired_value1=='KO' || desired_value1=='ko'))
            {
              nbrko=nbrko + 1;
            }
            else
            {
              nbrtsisy = 1;
            };
          };
         var a = table[nbe];
         console.log('nbe'+a);
         console.log(desired_value1);
         /*var sql = "insert into "+table[nbe]+" (nbok,nbko) values ("+nbr+","+nbrko+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
                          return callback(err); }
                        else
                        {
                          console.log(sql);
                          return callback(null, true);
                        }
                       
                                            });*/
          var tab = [nbr,nbrko];
          console.log("nombreeeeebr"+ nbr);
          return tab;
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
importEssai: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,callback) {
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
                 var sql = "insert into cheminhtp (chemin,nomtable,numligne,numfeuile,colonnecible,colonnecible2) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"','"+colonnecible2+"') ";
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
  importEssaitype2: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,colonnecible2,callback) {
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
                 re = a+'/'+file;  
                 var sql = "insert into cheminhtp2 (chemin,nomtable,numligne,numfeuile,colonnecible,colonnecible2) values ('"+re+"','"+nomTable+"','"+numLigne+"','"+numFeuille+"','"+nomColonne+"','"+colonnecible2+"') ";
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
                                       }) ;
                   
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
    };   
  },
  deleteFromChemin : function (table,callback) {
    var sql = "delete from cheminhtp ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return callback(err); }
      return callback(null, true);
      });
  },
  deleteFromChemin2 : function (table,callback) {
    var sql = "delete from cheminhtp2 ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
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
  
  importInovcom: function (trameflux,feuil,cellule,table,cellule2,numligne,nb,callback) {
    var tab = [];
    tab = Reportinghtp.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      Reportinghtp.deleteToutHtp(table,3,callback);
      
    }
    else{
      for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log(j);
      Reportinghtp.insertion(trameflux,feuil,cellule,table,cellule2,j,numligne,callback);
    }
    };
  },
  insertion:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    var tab= Reportinghtp.lectureEtInsertionModifie(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
    console.log(tab);
    const fs = require('fs');
    fs.writeFile(table[nb]+'.txt', tab, (err) => {
            
            var sql = "insert into trame (typologiedelademande) values ('ok') ";
                Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
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
    //var data = XLSX.utils.sheet_to_json(workbook.Sheets[first_sheet_name[0]]);
    //console.log('longueur data' + data.length);
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

      tab.push(desired_value2);
      tabl.push(desired_value1);
    }
    var com = [];
    for(var i=0;i<tab.length;i++)
    {
        com.push(tab[i]+';'+tabl[i]+'\n');
    };
    console.log(com);
    return com;
  },
  
  importTout: function (trameflux,table,nb,callback) {
    var tab = [];
    console.log('table miexiste'+tab);
    for(var i=0;i<5;i++){
      //var j = parseInt(tab[i]);
      Reportinghtp.importFinal(table,i,callback);
    };
  },
  importFinal: function (table,nb,callback) {
    var tablem = table[nb];
    console.log('tablem'+tablem);
    var chemin = 'D:/projet/'+tablem+'.txt';
    console.log(chemin);
    var sql = " COPY "+tablem+" FROM '"+chemin+"'  (DELIMITER(';') ) ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
      if(err) return console.log('erreur');
      else return callback(null, true);        
                          });
  },
  deleteTout: function (table,nb,callback) {
    for(var i=0;i<nb;i++){
      Reportinghtp.deleteFichier(table,i,callback);
    };
  },
  deleteFichier: function (table,nb,callback) {
    var tab= '';
    console.log(tab);
    const fs = require('fs');
    fs.writeFile(table[nb]+'.txt', tab, (err) => {
      var sql = "insert into trame (typologiedelademande) values ('k') ";
      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
        if(err) return console.log(err);
        else return callback(null, true);        
                            })      
    });
  },
  existenceRoute : function (trameflux,callback) {
    var sql= "select count(chemin) as ok from cheminhtp ";
    Reportinghtp.getDatastore().sendNativeQuery(sql ,function(err, nc) {
          if (err){
            console.log(err);
            return callback(err);
          }
          else
          {
            return callback(null, nc);
          };
      });
  },
  existenceRoute2 : function (trameflux,callback) {
    var sql= "select count(chemin) as ok from cheminhtp2 ";
    Reportinghtp.getDatastore().sendNativeQuery(sql ,function(err, nc) {
          if (err){
            console.log(err);
            return callback(err);
          }
          else
          {
            return callback(null, nc);
          };
      });
  },
  importTrameFlux929 : function (trameflux,feuil,cellule,table,cellule2,nb,callback) {
    var tab = [];
    tab = Reportinghtp.totalFichierExistant(trameflux,nb,callback);
    console.log(tab);
    if(tab.length==0)
    {
      console.log('Aucune reporting pour ce date');
      Reportinghtp.deleteToutHtp(table,3,callback);
    }
    else{
      for(var y=0;y<tab.length;y++) //parcours anle dossier rehetra
    {
      var j = parseInt(tab[y]);
      console.log(j);
      //Reportinghtp.lectureEtInsertion2( trameflux,feuil,cellule,table,cellule2,j,numligne,callback)
     Reportinghtp.lectureEtInsertion(trameflux,feuil,cellule,table,cellule2,j,callback)
    }
    };
  },
  totalFichierExistant : function (trameflux,nb,callback) {
    var tab = [];
    var j;
    var i = parseInt(j);
    for(i=0;i<nb;i++)
    {
      var a = Reportinghtp.existenceFichier(trameflux[i]);
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
  lectureEtInsertion:function(trameflux,feuil,cellule,table,cellule2,nb,callback){
    var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      try{
        workbook.xlsx.readFile(trameflux[nb])
          .then(function() {
              var newworksheet = workbook.getWorksheet(feuil[nb]);
              var row = newworksheet.getRow(1);
              var a;
              row.eachCell(function(cell, colNumber) {
                if(cell.text==cellule[nb])
                {
                  a = parseInt(colNumber);
                }
              });
              var b;
              row.eachCell(function(cell, colNumber) {
                if(cell.text==cellule2[nb])
                {
                  b = parseInt(colNumber);
                }
              });
              if(a!=undefined && b!=undefined)
              {
                var col = newworksheet.getColumn(a);
                var tab = [];
                console.log('col' + col);
                col.eachCell(function(cell, rowNumber) {
                  tab.push(cell.text);
                });
                var col2 = newworksheet.getColumn(b);
                var tabl = [];
                col2.eachCell(function(cell, rowNumber) {
                  var m = cell.text;
                  m=m.replace("'", "''");
                  console.log(m);
                  tabl.push(m);
                });
                var i;
                var j = parseInt(i);
                var cell;
                var cell2;
                var myJsonString = JSON.stringify(tabl);
                //console.log(myJsonString.length);
                var f = JSON.parse(myJsonString);
                //console.log(f);
               for(j=1;j<tab.length;j++)
                {
                  cell = tab[j];
                  cell2 = tabl[j];
                  var sql = "insert into "+table[nb]+" (typologiedelademande,okko) values ('"+cell2+"','"+cell+"') ";
                  Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                    if(err) return console.log(err);
                    else return callback(null, true);        
                                        })
                };
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
  
  lecture:function(option,callback){

      var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('D:/Reporting/Reporting/Essai.xlsx')
          .then(function() {
              var newworksheet = workbook.getWorksheet('Feuil1');
              var row = newworksheet.getRow(1);
              var a;
              row.eachCell(function(cell, colNumber) {
                if(cell.value=='cellule')
                {
                  a = parseInt(colNumber);
                }
              });
              var col = newworksheet.getColumn(a);
              var tab = [];
              col.eachCell(function(cell, rowNumber) {
                tab.push(cell.value);
              });
          });
  },
  del1:function(req,callback){
    var Excel = require('exceljs');
      var workbook = new Excel.Workbook();
      workbook.xlsx.readFile('//10.128.1.2/almerys-out/Retour_Easytech_20210112/RETOUR_ADHESION_CSS/adhesion ITE/Suivi Saisie Prod  20210112- ITE.xlsx')
          .then(function() {
              var newworksheet = workbook.getWorksheet('Feuil1');
              var row = newworksheet.getRow(1);
              var a;
              row.eachCell(function(cell, colNumber) {
                if(cell.value=='saisie OK/KO')
                {
                  a = parseInt(colNumber);
                }
              });
              var col = newworksheet.getColumn(a);
              var tab = [];
              col.eachCell(function(cell, rowNumber) {
                tab.push(cell.value);
              });
              var i;
              var j = parseInt(i);
              for(j=1;j<tab.length;j++)
              {
                cell = tab[j];
                var sql = "insert into suivisaisieprodite (okko) values ('"+cell+"') ";
                Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                  if(err) return console.log(err);
                  else return callback(null, true);        
                                      })
              };      
          });
  },
  deleteToutHtp : function (table,nb,callback) {
    var sql = "delete from "+table+" ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return callback(err); }
      return callback(null, true);
      });
  },
  deleteReportingHtp : function (table,nb,callback) {
    var sql = "delete from "+table[nb]+" ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log("Une erreur suppr?");
        return callback(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };
      });
  },
  deleteHtp : function (table,nb,callback) {
    var j;
    var i = parseInt(j);
    for(i=0;i<5;i++)
    {
      Reportinghtp.deleteReportingHtp(table,i,callback);
    };
  },
  insertResiliationHtp : function (table,nb,callback) {
    var sql = "insert into "+table[nb]+" (nbok,nbko,nbokresiliation,nbkoresiliation) values (0,0,0,0)";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return console.log(err); }
      return callback(null, true);
      });
  },
  insertReportingHtp : function (table,nb,callback) {
    var sql = "insert into "+table[nb]+" (nbok,nbko) values (0,0)";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return console.log(err); }
      return callback(null, true);
      });
  },
  insertHtp : function (table,nb,callback) {
    var j;
    var i = parseInt(j);
    for(i=0;i<nb;i++)
    {
      Reportinghtp.insertReportingHtp(table,i,callback);
    };
  },
  insertTrameFlux : function (param,callback) {
    var sql = "insert into trame(okko) values ('k') ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return callback(err); }
      return callback(null, true);
      });
  },
  };



