/**
 * ReportingIndu.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */

module.exports = {

  attributes: {
  },
  importEssaitype7: function (table,table2,date,option,nb,nomtable,numligne,numfeuille,nomcolonne,nomcolonne2,callback) {
    const fs = require('fs');
    var re  = 'a';
    var tab = [];
    var ab = table[0]+date+table2[nb];
    var b = option[nb];
    console.log('tonga ato v?');
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
                 console.log('chem'+ files[i]);
                 var chemi = files[i];
                 var alm = 'ALMERYS';
                 const regex2 = new RegExp(alm+'*');
                 var alm = 'CBTP';
                 const regex3 = new RegExp(alm+'*');
                
                 if(regex2.test(chemi))
                 {

                 fs.readdir(test1, (err, files1) => {
                   if(err){
                     console.log(err);
                   }
                   else{
                     const regex = new RegExp(b+'*');
                     for(var i = 0; i < files1.length; i++){
                      const regex = new RegExp(b+'*');
                      var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
                      var m2 = '^[^~]';
                      const regex11 = new RegExp(m1,'i');
                      const regex12 = new RegExp(m2);
                       if(regex.test(files1[i]) && regex11.test(files1[i]) && regex12.test(files1[i]))
                       {
                         var a =ab + file +"\\" + files1[i];
                         console.log('*****************');
                         console.log(a);  
                         var sql = "insert into cheminindu2 (chemin,nomtable,numligne,numfeuile,colonnecible,colonnecible2) values ('"+a+"','indurelevedecomptealmerys','"+numligne[nb]+"','"+numfeuille[nb]+"','"+nomcolonne[nb]+"','"+nomcolonne2[nb]+"') ";
                         ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                          if (err) { 
                            console.log("Une erreur ve oki1?");
                            //console.log(err);
                            
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
                          var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
                          ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                            if (err) { 
                              console.log("Une erreur ve ok2?");
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
                   }
                 });
                }
                else if(regex3.test(chemi)) {
                  
                 fs.readdir(test1, (err, files1) => {
                  if(err){
                    console.log(err);
                  }
                  else{
                    const regex = new RegExp(b+'*');
                    for(var i = 0; i < files1.length; i++){
                      if(regex.test(files1[i]))
                      {
                        var a =ab + file +"\\" + files1[i];
                        console.log('*****************');
                        console.log(a);  
                        var sql = "insert into cheminindu2 (chemin,nomtable,numligne,numfeuile,colonnecible,colonnecible2) values ('"+a+"','indurelevedecomptecbtp','"+numligne[nb]+"','"+numfeuille[nb]+"','"+nomcolonne[nb]+"','"+nomcolonne2[nb]+"') ";
                        ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                         if (err) { 
                           console.log("Une erreur ve oki1?");
                           //console.log(err);
                           
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
                         var sql = "insert into chemintsisy(typologiedelademande) values ('k') ";
                         ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                           if (err) { 
                             console.log("Une erreur ve ok2?");
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
                  }
                });
                }
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
        if (err) { 
          console.log("Une erreur ve ok3?");
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

  importTrameFlux929 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,date2,callback) {
    console.log(table[nb]);
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
    else
    {
      var nbe= parseInt(nb);
      var tab = [];
      if(table[nbe]=="indufraudelmg")
      {
        console.log('indufraudelmg');
        tab = ReportingIndu.lectureEtInsertion2(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
        console.log(tab);
        var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+tab[0]+"','"+tab[1]+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve insertion?");
                        return callback(err);
                       }
                      else
                      {
                        console.log(sql);
                        return callback(null, true);
                      };     
                                          });
      }
      else if(table[nbe]=="induentrain")
      {
        console.log('induentrain');
        tab = ReportingIndu.lectureEtInsertion3(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
        console.log(tab);
        var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+tab[0]+"','"+tab[1]+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve insertion?");
                        return callback(err);
                       }
                      else
                      {
                        console.log(sql);
                        return callback(null, true);
                      };     
                                          });
      }
      else if(table[nbe]=="indudentaire" || table[nbe]=="induoptique" || table[nbe]=="induaudio" ||table[nbe]=="induhospi" ||table[nbe]=="induse")
      {
        console.log('type2');
        tab = ReportingIndu.lectureEtInsertion4(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
        console.log(tab);
        var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+tab[0]+"','"+tab[1]+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve insertion?");
                        return callback(err);
                       }
                      else
                      {
                        console.log(sql);
                        return callback(null, true);
                      };     
                                          });
      }
      else if(table[nbe]=="inducodelisftp" || table[nbe]=="induinterialepost" || table[nbe]=="induinterialeaudio" ||table[nbe]=="induinterialepre" ||table[nbe]=="inducodelismail" ||table[nbe]=="inducodelisappel")
      {
        console.log('type2');
        tab = ReportingIndu.lectureEtInsertion5(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
        console.log(tab);
        var sql = "insert into "+table[nbe]+" (nbok) values ('"+tab[0]+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve insertion?");
                        return callback(err);
                       }
                      else
                      {
                        console.log(sql);
                        return callback(null, true);
                      };     
                                          });
      }
      else if(table[nbe]=="indupecrefus")
      {
        console.log('type2');
        for(var m=0;m<2;m++)
        {
          tab = ReportingIndu.lectureEtInsertion6(trameflux,m,cellule,table,cellule2,nb,numligne,callback);
          console.log(tab);
          var sql = "insert into "+table[nbe]+" (nbok) values ('"+tab[0]+"') ";
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
        };
        
      }
      else if(table[nbe]=="indusansnotif")
      {
        console.log('type7');
          var tab = [];
          tab = ReportingIndu.lectureEtInsertion7(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
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
      else if(table[nbe]=="induvalidation")
      {
        console.log('type8');
          var tab = [];
          tab = ReportingIndu.lectureEtInsertion8(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
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
      else if(table[nbe]=="inducheque")
      {
        console.log('type8');
          var tab = [];
          tab = ReportingIndu.lectureEtInsertion9(trameflux,feuil,cellule,table,cellule2,nb,numligne,date2,callback);
          console.log(tab);
         var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+tab[0]+"','"+tab[1]+"') ";
                     ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve insertion hoe?");
                          //return callback(err);
                         }
                        else
                        {
                          console.log(sql);
                          return callback(null, true);
                        };     
                                            });
        
      }
      else if(table[nbe]=="inducontestation")
      {
        console.log('inducontestation');
          var tab = [];
          tab = ReportingIndu.lectureEtInsertion10(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
          console.log(tab);
          for(var i=0;i<tab.length;i++)
          {
            var sql = "insert into "+table[nbe]+" (nbok) values ('"+tab[i]+"') ";
            ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
               if (err) { 
                 console.log("Une erreur ve insertion hoe?");
                 //return callback(err);
                }
               else
               {
                 console.log(sql);
                 return callback(null, true);
               };     
                                   });
          }
         
        
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
  importTrameFlux9292 : function (trameflux,feuil,cellule,table,cellule2,nb,numligne,callback) {
    console.log(table[nb]);
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
    else
    {
      var nbe= parseInt(nb);
      var tab = [];
      if(table[nbe]=="indurelevedecomptealmerys" || table[nbe]=="indurelevedecomptecbtp" )
      {
        console.log('relevedecompte');
        tab = ReportingIndu.lectureEtInsertiontype2(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
        console.log(tab);
        var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+tab[0]+"','"+tab[1]+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve insertion?");
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
    var col = 0;
    var nbe = parseInt(nb);
   if(col!=undefined)
    {
      console.log('tafa');

      var debutligne = numeroligne + 1;
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:6, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.v : undefined);
          var address_of_cell2 = {c:7, r:a};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);

          if(desired_value1!=undefined)
          {
            nbr=nbr + 1;
          }
          if(desired_value2!=undefined)
          {
            nbrko=nbrko + 1;
          }
          else
          {
            var an = 1;
          };
        };
        console.log("nombreeeeebr"+ nbr + 'et' + nbrko );
        var tab = [nbr,nbrko];
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
  };
  },
  lectureEtInsertiontype2:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
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
    var col = 6;
    //var col = 16;
    var nbe = parseInt(nb);
    /*for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(desired_value==cellule[nb])
        {
          col=ra;
        };
      };*/
      console.log("colonne"+col);
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
          if(desired_value1!=undefined)
          {
            if(desired_value1=='OUI' || desired_value1=='oui')
            {
              nbr=nbr + 1;
            }
            else if(desired_value1=='NON' || desired_value1=='non')
            {
              nbrko = nbrko +1;
            }
            else
            {
              var ap =0;
            };
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
    /*var tab = [nbr,nbrko,nbrokrib];
    return tab;*/
  }
  catch
  {
    console.log("erreur absolu haaha");
  };
  },
  lectureEtInsertion10:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
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
    var col = 0;
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
      var tab  = [];
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
          if(desired_value1!=undefined)
          {
            tab.push(desired_value1);
          }
          
        };  
    }
    else
    {
      console.log('Colonne non trouvé');
    };
    console.log("nombreeeeebr"+ tab);
    return tab;
    /*var tab = [nbr,nbrko,nbrokrib];
    return tab;*/
  }
  catch
  {
    console.log("erreur absolu haaha");
  };
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
    var col = 0;
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
          if(desired_value1!=undefined)
          {
            if(desired_value1=='PRE')
            {
              nbr=nbr + 1;
            }
            else if(desired_value1=='POST')
            {
              nbrko = nbrko +1;
            }
            else
            {
              var ap =0;
            };
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
    /*var tab = [nbr,nbrko,nbrokrib];
    return tab;*/
  }
  catch
  {
    console.log("erreur absolu haaha");
  };
  },
  lectureEtInsertion5:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
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
    var col = 0;
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
          if(desired_value1!=undefined)
          {
            nbr=nbr + 1;
          }
          else
          {
            var an = 1;
          };
          
        };  
    }
    else
    {
      console.log('Colonne non trouvé');
    };
    console.log("nombreeeeebr"+ nbr );
    var tab = [nbr,nbrko];
    return tab;
    /*var tab = [nbr,nbrko,nbrokrib];
    return tab;*/
  }
  catch
  {
    console.log("erreur absolu haaha");
  };
  },
  lectureEtInsertion6:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
  var workbook = XLSX.readFile(trameflux[nb]);
  var numerofeuille = parseInt(feuil);
  var numeroligne = parseInt(numligne[nb]);
  try{
    var nbr = 0;
    var nbrko = 0;
    var nbrokrib = 0;
    var nbrtsisy = 0;
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col = 0;
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
          if(desired_value1!=undefined)
          {
            nbr=nbr + 1;
          }
          else
          {
            var an = 1;
          };
          
        };  
    }
    else
    {
      console.log('Colonne non trouvé');
    };
    console.log("nombreeeeebr"+ nbr );
    var tab = [nbr];
    return tab;
    /*var tab = [nbr,nbrko,nbrokrib];
    return tab;*/
  }
  catch
  {
    console.log("erreur absolu haaha");
  };
  },
  lectureEtInsertion7:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
  var workbook = XLSX.readFile(trameflux[nb]);
  var numerofeuille = parseInt(feuil);
  var numeroligne = parseInt(numligne[nb]);
  try{
    var nbr = 0;
    var nbrko = 0;
    var nbrokrib = 0;
    var nbrtsisy = 0;
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col = 0;
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
   if(col!=undefined )
    {
      var debutligne = numeroligne + 1;
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:col, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.v : undefined);

          var address_of_cell2 = {c:0, r:a};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
          //console.log(desired_value1);
          if(desired_value1!=undefined)
          {
            nbr=nbr + 1;
          }
          if(desired_value2!=undefined)
          {
            nbrko=nbrko + 1;
          }
          else
          {
            var an = 1;
          };
          
        };  
    }
    else
    {
      console.log('Colonne non trouvé');
    };
    
    var nb= 0;
    nb = parseInt(nbrko) - parseInt(nbr);
    console.log("nombreeeeebr"+ nbr + 'et' + nb);
    var tab = [nbr,nb];
    return tab;
    /*var tab = [nbr,nbrko,nbrokrib];
    return tab;*/
  }
  catch
  {
    console.log("erreur absolu haaha");
  };
  },
  lectureEtInsertion8:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
  var workbook = XLSX.readFile(trameflux[nb]);
  var numerofeuille = parseInt(feuil);
  var numeroligne = parseInt(numligne[nb]);
  try{
    var nbr = 0;
    var nbrko = 0;
    var nbrokrib = 0;
    var nbrtsisy = 0;
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col = 0;
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
          if(desired_value1!=undefined)
          {
            nbrko=nbrko + 1;
          }
          else
          {
            var an = 1;
          };
          
        };  
    }
    else
    {
      console.log('Colonne non trouvé');
    };
    
    var nb= 0;
  
    console.log("nombreeeeebr"+ nbr + 'et' + nbrko);
    var tab = [nbr,nbrko];
    return tab;
    /*var tab = [nbr,nbrko,nbrokrib];
    return tab;*/
  }
  catch
  {
    console.log("erreur absolu haaha");
  };
  },
  lectureEtInsertion9:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,date2,callback){
    XLSX = require('xlsx');
  var workbook = XLSX.readFile(trameflux[nb]);
  var numerofeuille = parseInt(feuil);
  var numeroligne = parseInt(numligne[nb]);
  try{
    var nbr = 0;
    var nbrko = 0;
    var nbrokrib = 0;
    var nbrtsisy = 0;
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col = 0;
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
   if(col!=undefined )
    {
      var debutligne = numeroligne + 1;
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:1, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.w : undefined);

          var address_of_cell2 = {c:13, r:a};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.w : undefined);
          if(desired_value1==date2 && (desired_value2=="OUI" || desired_value2=='oui'))
          {
            nbr=nbr + 1;
          }
          else
          {
            var an = 1;
          };
          
        };  
    }
    else
    {
      console.log('Colonne non trouvé');
    };
    
    var nb= 0;
  
    console.log("nombreeeeebr"+ nbr + 'et' + nbrko);
    var tab = [nbr,nbrko];
    return tab;
    /*var tab = [nbr,nbrko,nbrokrib];
    return tab;*/
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
    var col = 0;
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
          if(desired_value1!=undefined)
          {
           nbr=nbr + 1;
           if(desired_value1=='D' || desired_value1=='R')
            {
              nbrko = nbrko +1;
            } 
            else
            {
              var ap =0;
            };
          }
          else
          {
            var f = 4;
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
    /*var tab = [nbr,nbrko,nbrokrib];
    return tab;*/
  }
  catch
  {
    console.log("erreur absolu haaha");
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

