/**
 * ReportingIndu.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
//const path_reporting = 'D:/Reporting/Reporting/REPORTING INDU Type.xlsx';
//const path_reporting = 'D:/LDR8_1421_nouv/PROJET_FELANA/REPORTING INDU Type.xlsx';
const path_reporting = '/dev/prod/00-TOUS/TestReporting/REPORTING INDU Type.xlsx';
module.exports = {

  attributes: {
  },
  lectureEtInsertion4:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
  var workbook = XLSX.readFile(trameflux[nb]);
  var numerofeuille = feuil[nb];
  var numeroligne = parseInt(numligne[nb]);
  try{
    var nbr = 0;
    var nbrko = 0;
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col=0;
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
                         //var a =ab + file +"\\" + files1[i];
                         var a =ab + file +"/" + files1[i];
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
                        //var a =ab + file +"\\" + files1[i];
                        var a =ab + file +"/" + files1[i];
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
      else if(table[nbe]=="indudentaire" || table[nbe]=="induoptique" || table[nbe]=="induaudio" ||table[nbe]=="induhospi")
      {
        console.log('type4');
        tab = ReportingIndu.lectureEtInsertion4(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
        console.log(tab);
        var sql = "insert into "+table[nbe]+" (nbok,nbko) values ('"+tab[0]+"','"+tab[1]+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve insertion?");
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
      else if(table[nbe]=="indutiers" || table[nbe]=="induse" )
      {
        console.log('indutiers ou induse');
        tab = ReportingIndu.lectureEtInsertion41(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
        console.log(tab);
        var sql = "insert into "+table[nbe]+" (nbok,nbko,nbokcbtp,nbkocbtp) values ('"+tab[0]+"','"+tab[1]+"','"+tab[2]+"','"+tab[3]+"') ";
                   ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                      if (err) { 
                        console.log("Une erreur ve insertion?");
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
      else if(table[nbe]=="inducodelisftp" || table[nbe]=="induinterialepost" || table[nbe]=="induinterialeaudio" ||table[nbe]=="induinterialepre" ||table[nbe]=="inducodelismail" ||table[nbe]=="inducodelisappel")
      {
        console.log('type2');
        tab = ReportingIndu.lectureEtInsertion5(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
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
      else if(table[nbe]=="inducheque")
      {
        console.log('inducheque');
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
      if(table[nbe]=="indurelevedecomptealmerys" || table[nbe]=="indurelevedecomptecbtp" )
      {
        console.log('relevedecompte');
        tab = ReportingIndu.lectureEtInsertiontype2(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback);
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
  lectureEtInsertion41:function(trameflux,feuil,cellule,table,cellule2,nb,numligne,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux[nb]);
    var numerofeuille = feuil[nb];
    var numeroligne = parseInt(numligne[nb]);
    try{
      var nbr = 0;
      var nbrko = 0;
      var nbrokcbtp = 0;
      var nbrkocbtp = 0;
      const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var col;
      var nbe = parseInt(nb);
    
      var colb ;
      for(var ra=0;ra<=range.e.c;ra++)

        {
          var bb = 'Spécialité';
          const regex4 = new RegExp(bb,'i');
          var address_of_cell = {c:ra, r:numeroligne};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          console.log(desired_value);
          if(regex4.test(desired_value))
          {
            colb=ra;
          };
        };
        console.log("colonneb"+colb);
    if(colb!=undefined)
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

            var address_of_cell3 = {c:colb, r:a};
            var cell_ref3 = XLSX.utils.encode_cell(address_of_cell3);
            var desired_cell3 = sheet[cell_ref3];
            var desired_value3 = (desired_cell3 ? desired_cell3.v : undefined);

            var cbtp = 'CBTP';
            const regex = new RegExp(cbtp,'i');
            if(regex.test(desired_value3))
            {
              if(desired_value1!=undefined)
            {
              nbrokcbtp=nbrokcbtp + 1;
            }
            if(desired_value2!=undefined)
            {
              nbrkocbtp=nbrkocbtp + 1;
            }
            else
            {
              var an = 1;
            };
            }
            else
            {
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
            }

            
          };
          
            
      }
      else
      {
        console.log('Colonne non trouvé');
      };
      console.log("nombreeeeebr"+ nbr + 'et' + nbrko + 'et' + nbrokcbtp + 'et' + nbrokcbtp);
          var tab = [nbr,nbrko,nbrokcbtp,nbrkocbtp];
          return tab;
      
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
    var nbe = parseInt(nb);
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
    var col;
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
    var col;
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
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col=10;
    //var col = 16;
    var nbe = parseInt(nb);
   /* for(var ra=0;ra<=range.e.c;ra++)
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
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col;
 
    for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        var bb = 'Finess PS';
        const regex = new RegExp(bb,'i');
        if(regex.test(desired_value))
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
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
   
    var indunonsaisi = 'INDU_NON_SAISI_MOTIF';
    const regexns1= new RegExp(indunonsaisi,'i');
    var indunonsaisi2 = 'Indu non saisi motif';
    const regexns2= new RegExp(indunonsaisi2,'i');
   
    
    var indusaisi = 'INDU_SANS_NOTIFICATION_SAISI_LE';
    const regexs1= new RegExp(indusaisi ,'i');
    var indusaisi2 = 'Indu sans notification saisi le';
    const regexs2= new RegExp(indusaisi2 ,'i');

    var debutligne;
    var cols;
    var colns ;
    for(var ra=0;ra<=range.e.c;ra++)
    {
      var m = 0;
      var address_of_cell = {c:ra, r:0};
      var cell_ref = XLSX.utils.encode_cell(address_of_cell);
      var desired_cell = sheet[cell_ref];
      var desired_value = (desired_cell ? desired_cell.v : undefined);
      if(regexs1.test(desired_value) || regexs2.test(desired_value) )
      {
        cols=ra;
        debutligne = m + 1;
      };
    };
  for(var ra=0;ra<=range.e.c;ra++)
    {
      var m = 0;
      var address_of_cell = {c:ra, r:0};
      var cell_ref = XLSX.utils.encode_cell(address_of_cell);
      var desired_cell = sheet[cell_ref];
      var desired_value = (desired_cell ? desired_cell.v : undefined);
      if(regexns1.test(desired_value) ||regexns2.test(desired_value) )
      {
        colns=ra;
        debutligne = m + 1;
      };
    };
    for(var ra=0;ra<=range.e.c;ra++)
      {
        var m = 1;
        var address_of_cell = {c:ra, r:1};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(regexs1.test(desired_value) || regexs2.test(desired_value) )
        {
          cols=ra;
          var debutligne = m + 1;
        };
      };
    for(var ra=0;ra<=range.e.c;ra++)
      {
        var m = 1;
        var address_of_cell = {c:ra, r:1};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(regexns1.test(desired_value) ||regexns2.test(desired_value) )
        {
          colns=ra;
          debutligne = m + 1;
        };
      };
    for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(regexs1.test(desired_value) || regexs2.test(desired_value) )
        {
          cols=ra;
          debutligne = numeroligne + 1;
        };
      };
    for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(regexns1.test(desired_value) ||regexns2.test(desired_value) )
        {
          colns=ra;
          debutligne = numeroligne + 1;
        };
      };
      console.log("colonnes"+cols);
      console.log("colonnens"+colns);
      console.log("debut"+debutligne);
    if(cols!=undefined && colns!=undefined)
    {
      
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:cols, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.v : undefined);

          var address_of_cell2 = {c:colns, r:a};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
          if(desired_value1!=undefined)
          {
            nbr=nbr + 1;
          }
          else if(desired_value2!=undefined && desired_value1==undefined)
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
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var nbe = parseInt(nb);
    var bb = 'Etat de fin de traitement';
    const regex4 = new RegExp(bb,'i');
    var colb = 6;
    /*for(var ra=0;ra<=range.e.c;ra++)

      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        if(regex4.test(desired_value))
        {
          colb=ra;
        };
      };*/
    console.log(colb);
   if(colb!=undefined )
    {
      var debutligne = numeroligne + 1;
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:colb, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.v : undefined);
          var ok = 'ok';
          const regexok = new RegExp(ok,'i');
          var ko= 'ko';
          const regexko = new RegExp(ko,'i');
          if(regexok.test(desired_value1))
          {
            nbr=nbr + 1;
          }
          else if(regexko.test(desired_value1))
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
    const sheet = workbook.Sheets[workbook.SheetNames[numerofeuille]];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var col;
    var col2=13; 
    var nbe = parseInt(nb);
    for(var ra=0;ra<=range.e.c;ra++)
      {
        var address_of_cell = {c:ra, r:numeroligne};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);
        var ko = cellule[nb];
        const regex1 = new RegExp(ko,'i');
        if(regex1.test(desired_value))
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
        var ok = cellule2[nb];
        console.log(desired_value);
        const regex1 = new RegExp(ok,'i');
        if(regex1.test(desired_value))
        {
          //col2=ra;
        };
      };
      console.log("colonne"+col + 'g' + col2);

   if(col!=undefined && col2!=undefined)
    {
      var debutligne = numeroligne + 1;
      for(var a=debutligne;a<=range.e.r;a++)
        {
          var address_of_cell = {c:col, r:a};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value1 = (desired_cell ? desired_cell.w : undefined);

          var address_of_cell2 = {c:col2, r:a};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);

          if(desired_value1==date2 && (desired_value2=="OUI" || desired_value2=='oui'))
          {
            nbr=nbr + 1;
          }
          else if(desired_value1==date2 && (desired_value2=="NON" || desired_value2=='non'))
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
    var col;
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
        if (err) { //
          //return callback(err); 
          console.log("erreur de suppression");
        }
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
        if (err) { 
          //return callback(err);
          console.log("erreur de suppression");
         }
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


/**********************************************************************************/
 // Récuperer nombre OK ou KO
 countOkKoDouble : function (table, callback) {
  const Excel = require('exceljs');
  var sqlOk ="select nbok from "+table; 
  var sqlKo ="select nbko from "+table;
 
  console.log(sqlOk);
  console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sqlOk, function(err, res){
        // if (err) return res.badRequest(err);
        // callback(null, res.rows[0].nbok);
        if (err) {
          console.log(err);
          //return null;
        }
        else
        {
          if(res.rows[0])
          {
            console.log('ok');
            callback(null, res.rows[0].nbok);
          }
          else
          {
            console.log("null");
            callback(null, 0);
          }
        }
      });
    },
    function (callback) {
      ReportingIndu.query(sqlKo, function(err, resKo){
        // if (err) return res.badRequest(err);
        // callback(null, resKo.rows[0].nbko);
        if (err) {
          console.log(err);
          //return null;
        }
        else
        {
          if(resKo.rows[0])
          {
            console.log('ok');
            callback(null, resKo.rows[0].nbko);
          }
          else
          {
            console.log("null");
            callback(null, 0);
          }
        }
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
/******************************************************************************************/
countOkKoDoubleSum : function (table, callback) {
  const Excel = require('exceljs');
  var sqlOk ="select sum(nbok::integer) from "+table; 
  var sqlKo ="select sum(nbko::integer) from  "+table;
 
  console.log(sqlOk);
  console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sqlOk, function(err, res){
        // if (err) return res.badRequest(err);
        // callback(null, res.rows[0].sum);
        if (err) {
          console.log(err);
          //return null;
        }
        else
        {
          if(res.rows[0])
          {
            console.log('ok');
            if(res.rows[0].sum == null){
              callback(null, 0);
            }
            else{
              callback(null, res.rows[0].sum);
            }
            
          }
          else
          {
            console.log("null");
            callback(null, 0);
          }
        }
      });
    },
    function (callback) {
      ReportingIndu.query(sqlKo, function(err, resKo){
        // if (err) return res.badRequest(err);
        // callback(null, resKo.rows[0].sum);
        if (err) {
          console.log(err);
          //return null;
        }
        else
        {
          if(resKo.rows[0])
          {
            console.log('ok');
            if(resKo.rows[0].sum == null){
              callback(null, 0);
            }
            else{
              callback(null, resKo.rows[0].sum);
            }
          }
          else
          {
            console.log("null");
            callback(null, 0);
          }
        }
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
/**********************************************************************************************/
countOkKoDoubleSumcbtp : function (table, callback) {
  const Excel = require('exceljs');
  var sqlOk ="select sum(nbokcbtp::integer) from "+table; 
  var sqlKo ="select sum(nbkocbtp::integer) from  "+table;
 
  console.log(sqlOk);
  console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sqlOk, function(err, res){
        // if (err) return res.badRequest(err);
        // callback(null, res.rows[0].sum);
        if (err) {
          console.log(err);
          //return null;
        }
        else
        {
          if(res.rows[0])
          {
            console.log('ok');
            if(res.rows[0].sum == null){
              callback(null, 0);
            }
            else{
              callback(null, res.rows[0].sum);
            }
            
          }
          else
          {
            console.log("null");
            callback(null, 0);
          }
        }
      });
    },
    function (callback) {
      ReportingIndu.query(sqlKo, function(err, resKo){
        // if (err) return res.badRequest(err);
        // callback(null, resKo.rows[0].sum);
        if (err) {
          console.log(err);
          //return null;
        }
        else
        {
          if(resKo.rows[0])
          {
            console.log('ok');
            if(resKo.rows[0].sum == null){
              callback(null, 0);
            }
            else{
              callback(null, resKo.rows[0].sum);
            }
          }
          else
          {
            console.log("null");
            callback(null, 0);
          }
        }
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
/**********************************************************************************************/
countOkKo : function (table, callback) {
  const Excel = require('exceljs');
  var sqlOk ="select nbok from "+table; //trameFlux
  // var sqlKo ="select nbko from "+table;
  console.log(sqlOk);
  // console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingInovcomExport.query(sqlOk, function(err, res){
        // if (err) return res.badRequest(err);
        // callback(null, res.rows[0].nbok);
        if (err) {
          console.log(err);
          //return null;
        }
        else
        {
          if(res.rows[0])
          {
            console.log('ok');
            callback(null, res.rows[0].nbok);
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
/********************************************************************************************/
countOkKoSum : function (table, callback) {
  const Excel = require('exceljs');
  // var sqlOk ="select count(okko) as ok from "+table+" where okko='OK'"; //trameFlux
  // var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'";
  var sql ="select sum(nbok::integer) from "+table; 
 
  console.log(sql);
  // console.log(sqlOk);
  // console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sql, function(err, res){
        // if (err) return res.badRequest(err);
        // // callback(null, res.rows[0].ok);
        // console.log(res.rows[0].sum);
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
        // if(res.rows[0].sum != undefined){
        //   callback(null, res.rows[0].sum);
        // }
        // else{
        //   return res.rows[0].sum = 0;
        // }
        
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
    console.log("Count OK somme ==> " + result[0]);
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
countOkKoSumko : function (table, callback) {
  const Excel = require('exceljs');
  // var sqlOk ="select count(okko) as ok from "+table+" where okko='OK'"; //trameFlux
  // var sqlKo ="select count(okko) as ko from "+table+" where okko='KO'";
  var sql ="select sum(nbko::integer) from "+table; 
 
  console.log(sql);
  // console.log(sqlOk);
  // console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sql, function(err, res){
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
/***************************************************************************/
countOkKoIndu2 : function (table, callback) {
  const Excel = require('exceljs');
  // var sqlOk ="select nbok from "+table; 
  // var sqlKo ="select nbko from "+table;
  var sqlOk ="select sum(nbok::integer) from "+table; 
  var sqlKo ="select sum(nbko::integer) from  "+table;

  console.log(sqlOk);
  console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sqlOk, function(err, res){
        // if (err) return res.badRequest(err);
        // callback(null, res.rows[0].nbok);
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
    function (callback) {
      ReportingIndu.query(sqlKo, function(err, resKo){
        // if (err) return res.badRequest(err);
        // callback(null, resKo.rows[0].nbko);
        if (err) {
          console.log(err);
          //return null;
        }
        else
        {
          if(resKo.rows[0])
          {
            console.log('ok');
            callback(null, resKo.rows[0].sum);
          }
          else
          {
            console.log("null");
            callback(null, 0);
          }
        }
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
/************************************************************************* */
countOkKoContest : function (table, callback) {
  const Excel = require('exceljs');
  // var sqlOk ="select nbok from "+table; 
  // var sqlKo ="select nbko from "+table;
  var sqlOk ="SELECT COUNT(DISTINCT nbok) AS Count FROM "+table;
  console.log(sqlOk);
  // console.log(sqlKo);
  async.series([
    function (callback) {
      ReportingIndu.query(sqlOk, function(err, res){
        // if (err) return res.badRequest(err);
        // callback(null, res.rows[0].count);
        if (err) {
          console.log(err);
          //return null;
        }
        else
        {
          if(res.rows[0])
          {
            console.log('ok');
            callback(null, res.rows[0].count);
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
    //   ReportingIndu.query(sqlKo, function(err, resKo){
    //     if (err) return res.badRequest(err);
    //     callback(null, resKo.rows[0].nbko);
    //   });
    // },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK CONTEST ==> " + result[0]);
    // console.log("Count KO ==> " + result[1]);
    var okko = {};
    okko.ok = result[0];
    // okko.ko = result[1];      
    return callback(null, okko);
  })
},
/***********************************************************************/
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
/************************************************************************* */
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
    var dateExcel = ReportingIndu.convertDate(cell.text);
    // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
    // var valiny = Retour.convertDate(andro);
    // console.log(valiny);
    if(dateExcel==date_export)
    {
      ligneDate1 = parseInt(rowNumber);
      var line = newworksheet.getRow(ligneDate1);
      var f = line.getCell(3).value;
      // console.log(f);
      if(f == "almerys" || f == "Almerys")
      {
        ligneDate = parseInt(rowNumber);
      }
    }
  });
  console.log("LIGNE DATE ===> "+ ligneDate);
  var rowDate = newworksheet.getRow(ligneDate);
  var numeroLigne = rowDate;
  var iniValue = ReportingIndu.getIniValue(table);
  
  var a5;

  var rowm = newworksheet.getRow(1);
 

  var collonne;
  var colDate2;
  rowm.eachCell(function(cell, colNumber) {
    if(cell.value == 'DOCUMENTS SAISIS')
    {
      colDate2 = parseInt(colNumber);
      var man = newworksheet.getRow(3);
      var f = man.getCell(colDate2).value;
      var getko_ini = man.getCell(colDate2).address;
      if(getko_ini == iniValue.ko+3 && f == iniValue.ok)
      {
        collonne = parseInt(colNumber);
      }
    }
  });
  console.log(" Colnumber2"+collonne);
  numeroLigne.getCell(collonne).value = nombre_ok_ko.ok;
  
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
  ecritureOkKoko : async function (nombre_ok_ko, table,date_export,mois1,callback) {
    const Excel = require('exceljs');
    const cmd=require('node-cmd');
    const newWorkbook = new Excel.Workbook();
    
    try{
    
     console.log('miditra ato am ecritureOkkoko');
      await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet(mois1);
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    var ligneDate;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingIndu.convertDate(cell.text);
      // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
      // var valiny = Retour.convertDate(andro);
      // console.log(valiny);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        // console.log(f);
        if(f == "almerys" || f == "Almerys")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingIndu.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
   
  
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
    //numeroLigne.getCell(collonne).value = nombre_ok_ko.ok;
    if(numeroLigne.getCell(collonne).value == undefined || numeroLigne.getCell(collonne).value == undefined){
      nombre_ok_ko.ko = 0;
    }
    else{
      numeroLigne.getCell(collonne).value = nombre_ok_ko.ko;
    }
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
  ecritureOkKoDouble : async function (nombre_ok_ko, table,date_export,mois1,callback) {
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
      var dateExcel = ReportingIndu.convertDate(cell.text);
      // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
      // var valiny = Retour.convertDate(andro);
      // console.log(valiny);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        // console.log(f);
        if(f == "almerys" || f == "Almerys")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingIndu.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = nombre_ok_ko.ok;
    numeroLigne.getCell(collonne).value = nombre_ok_ko.ko;
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
    ecritureOkKoDoublecbtp : async function (nombre_ok_ko, table,date_export,mois1,callback) {
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
        var dateExcel = ReportingIndu.convertDate(cell.text);
        // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
        // var valiny = Retour.convertDate(andro);
        // console.log(valiny);
        if(dateExcel==date_export)
        {
          ligneDate1 = parseInt(rowNumber);
          var line = newworksheet.getRow(ligneDate1);
          var f = line.getCell(3).value;
          // console.log(f);
          if(f == "cbtp")
          {
            ligneDate = parseInt(rowNumber);
          }
        }
      });
      console.log("LIGNE DATE ===> "+ ligneDate);
      var rowDate = newworksheet.getRow(ligneDate);
      var numeroLigne = rowDate;
      var iniValue = ReportingIndu.getIniValue(table);
      
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
      numeroLigne.getCell(colonnne).value = nombre_ok_ko.ok;
      numeroLigne.getCell(collonne).value = nombre_ok_ko.ko;
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
  ecritureOkKoIndu2 : async function (nombre_ok_ko, table,date_export,mois1,callback) {
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
      var dateExcel = ReportingIndu.convertDate(cell.text);
      // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
      // var valiny = Retour.convertDate(andro);
      // console.log(dateExcel);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        // console.log(f);
        if(f == "almerys" || f == "Almerys")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingIndu.getIniValue(table);
    
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
    numeroLigne.getCell(colonnne).value = nombre_ok_ko.ok;
    numeroLigne.getCell(collonne).value = nombre_ok_ko.ko;
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
    ecritureOkKoIndu2cbtp : async function (nombre_ok_ko, table,date_export,mois1,callback) {
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
        var dateExcel = ReportingIndu.convertDate(cell.text);
        // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
        // var valiny = Retour.convertDate(andro);
        // console.log(dateExcel);
        if(dateExcel==date_export)
        {
          ligneDate1 = parseInt(rowNumber);
          var line = newworksheet.getRow(ligneDate1);
          var f = line.getCell(3).value;
          // console.log(f);
          if(f == "cbtp")
          {
            ligneDate = parseInt(rowNumber);
          }
        }
      });
      console.log("LIGNE DATE ===> "+ ligneDate);
      var rowDate = newworksheet.getRow(ligneDate);
      var numeroLigne = rowDate;
      var iniValue = ReportingIndu.getIniValue(table);
      
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
      numeroLigne.getCell(colonnne).value = nombre_ok_ko.ok;
      numeroLigne.getCell(collonne).value = nombre_ok_ko.ko;
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
      ecritureOkKoContest : async function (nombre_ok_ko, table,date_export,mois1,callback) {
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
          var dateExcel = ReportingIndu.convertDate(cell.text);
          // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
          // var valiny = Retour.convertDate(andro);
          // console.log(dateExcel);
          if(dateExcel==date_export)
          {
            ligneDate1 = parseInt(rowNumber);
            var line = newworksheet.getRow(ligneDate1);
            var f = line.getCell(3).value;
            // console.log(f);
            if(f == "almerys" || f == "Almerys")
            {
              ligneDate = parseInt(rowNumber);
            }
          }
        });
        console.log("LIGNE DATE ===> "+ ligneDate);
        var rowDate = newworksheet.getRow(ligneDate);
        var numeroLigne = rowDate;
        var iniValue = ReportingIndu.getIniValue(table);
        
        var a5;
      
        var rowm = newworksheet.getRow(1);
        // var colonnne;
        // var colDate1;
        // rowm.eachCell(function(cell, colNumber) {
        //   if(cell.value == 'DOCUMENTS SAISIS')
        //   {
        //     colDate1 = parseInt(colNumber);
        //     //var col = newworksheet.getColumn(colDate1);
        //     var man = newworksheet.getRow(3);
        //     var f = man.getCell(colDate1).value;
        //     if(f == iniValue.ok)
        //     {
        //       colonnne = parseInt(colNumber);
        //     }
        //     }
        // });
        // console.log(" Colnumber"+colonnne);
      
        var collonne;
        var colDate2;
        rowm.eachCell(function(cell, colNumber) {
          if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
          {
            colDate2 = parseInt(colNumber);
            var man = newworksheet.getRow(3);
            var f = man.getCell(colDate2).value;
            var getko_ini = man.getCell(colDate2).address;
          // console.log(getko_ini);
          if(getko_ini == iniValue.ko+3 && f == iniValue.ok)
            {
              collonne = parseInt(colNumber);
            }
          }
        });
        console.log(" Colnumber2"+collonne);
        // numeroLigne.getCell(colonnne).value = nombre_ok_ko.ok;
        numeroLigne.getCell(collonne).value = nombre_ok_ko.ok;
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
  ecritureOkKoSante : async function (nombre_ok_ko, table,date_export,mois1,callback) {
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
      var dateExcel = ReportingIndu.convertDate(cell.text);
      // var andro = "Wed May 12 2021 03:00:00 GMT+0300 (heure normale de l’Arabie)";
      // var valiny = Retour.convertDate(andro);
      // console.log(dateExcel);
      if(dateExcel==date_export)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(3).value;
        // console.log(f);
        if(f == "Santéclair" || f == "Santeclair")
        {
          ligneDate = parseInt(rowNumber);
        }
      }
    });
    console.log("LIGNE DATE ===> "+ ligneDate);
    var rowDate = newworksheet.getRow(ligneDate);
    var numeroLigne = rowDate;
    var iniValue = ReportingIndu.getIniValue(table);
    
    var a5;
  
    var rowm = newworksheet.getRow(1);
    // var colonnne;
    // var colDate1;
    // rowm.eachCell(function(cell, colNumber) {
    //   if(cell.value == 'DOCUMENTS SAISIS')
    //   {
    //     colDate1 = parseInt(colNumber);
    //     //var col = newworksheet.getColumn(colDate1);
    //     var man = newworksheet.getRow(3);
    //     var f = man.getCell(colDate1).value;
    //     if(f == iniValue.ok)
    //     {
    //       colonnne = parseInt(colNumber);
    //     }
    //     }
    // });
    // console.log(" Colnumber"+colonnne);
  
    var collonne;
    var colDate2;
    rowm.eachCell(function(cell, colNumber) {
      if(cell.value == 'DOCUMENTS TRAITES NON SAISIS (RETOURS)')
      {
        colDate2 = parseInt(colNumber);
        var man = newworksheet.getRow(3);
        var f = man.getCell(colDate2).value;
        var getko_ini = man.getCell(colDate2).address;
      // console.log(getko_ini);
      if(getko_ini == iniValue.ko+3 && f == iniValue.ok)
        {
          collonne = parseInt(colNumber);
        }
      }
    });
    console.log(" Colnumber2"+collonne);
    numeroLigne.getCell(collonne).value = nombre_ok_ko.ok;
    // numeroLigne.getCell(collonne).value = nombre_ok_ko.ko;
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
    const config = ini.parse(fs.readFileSync('./config_excelIndu.ini', 'utf-8'));
    // console.log(config);
    return config;
  },

  getIniValue : function(table) {
    var iniValue = ReportingIndu.getConfigIni();
    var numeroColonneOk,numeroColonneKo;
    
    if(table == "indurelevedecomptealmerys"){
      numeroColonneOk = iniValue.indurelevedecomptealmerys.ok;
      numeroColonneKo = iniValue.indurelevedecomptealmerys.ko;
    }
    if(table == "indurelevedecomptecbtp"){
      numeroColonneOk = iniValue.indurelevedecomptecbtp.ok;
      numeroColonneKo = iniValue.indurelevedecomptecbtp.ko;
    }
    if(table == "induse"){
      numeroColonneOk = iniValue.induse.ok;
      numeroColonneKo = iniValue.induse.ko;
    }
    if(table == "induhospi"){
      numeroColonneOk = iniValue.induhospi.ok;
      numeroColonneKo = iniValue.induhospi.ko;
    }
    if(table == "indusansnotif"){
      numeroColonneOk = iniValue.indusansnotif.ok;
      numeroColonneKo = iniValue.indusansnotif.ko;
    }
   if(table == "indutiers"){
      numeroColonneOk = iniValue.indutiers.ok;
      numeroColonneKo = iniValue.indutiers.ko;
    }
    if(table == "indufraudelmg"){
      numeroColonneOk = iniValue.indufraudelmg.ok;
      numeroColonneKo = iniValue.indufraudelmg.ko;
    }
    if(table == "indufraudelmg"){
      numeroColonneOk = iniValue.indufraudelmg.ok;
      numeroColonneKo = iniValue.indufraudelmg.ko;
    }
    if(table == "induinterialepre"){
      numeroColonneOk = iniValue.induinterialepre.ok;
      numeroColonneKo = iniValue.induinterialepre.ko;
    }
    if(table == "induinterialepost"){
      numeroColonneOk = iniValue.induinterialepost.ok;
      numeroColonneKo = iniValue.induinterialepost.ko;
    }
    if(table == "inducodelisftp"){
      numeroColonneOk = iniValue.inducodelisftp.ok;
      numeroColonneKo = iniValue.inducodelisftp.ko;
    }
    if(table == "inducodelismail"){
      numeroColonneOk = iniValue.inducodelismail.ok;
      numeroColonneKo = iniValue.inducodelismail.ko;
    }
    if(table == "inducodelisappel"){
      numeroColonneOk = iniValue.inducodelisappel.ok;
      numeroColonneKo = iniValue.inducodelisappel.ko;
    }
    if(table == "inducheque"){
      numeroColonneOk = iniValue.inducheque.ok;
      numeroColonneKo = iniValue.inducheque.ko;
    }
    if(table == "indupecrefus"){
      numeroColonneOk = iniValue.indupecrefus.ok;
      numeroColonneKo = iniValue.indupecrefus.ko;
    }
    if(table == "inducontestation"){
      numeroColonneOk = iniValue.inducontestation.ok;
      numeroColonneKo = iniValue.inducontestation.ko;
    }
    if(table == "induinterialeaudio"){
      numeroColonneOk = iniValue.induinterialeaudio.ok;
      numeroColonneKo = iniValue.induinterialeaudio.ko;
    }
    if(table == "indufactstc"){
      numeroColonneOk = iniValue.indufactstc.ok;
      numeroColonneKo = iniValue.indufactstc.ko;
    }
    if(table == "indufraudelmgdent"){
      numeroColonneOk = iniValue.indufraudelmgdent.ok;
      numeroColonneKo = iniValue.indufraudelmgdent.ko;
    }
    
    // if(table == ""){
    //   numeroColonneOk = iniValue..ok;
    //   numeroColonneKo = iniValue..ko;
    // }
   
    var ok_ko = {};
    ok_ko.ok = numeroColonneOk;
    ok_ko.ko = numeroColonneKo;

    console.log("INI OK = "+ok_ko.ok);
    console.log("INI KO = "+ok_ko.ko);
    return ok_ko;
  },


};

