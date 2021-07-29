/**
 * TpsGrs.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
 const path_reporting = '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/TestReporting/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
//const path_reporting = '//10.128.1.2/bpo_almerys/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/TestReporting/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
module.exports = {
  attributes: {
  },
  ecritureDate : async function (tab,date_export,row,callback) {
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    //var path_reporting = 'D:/Reporting Engagement/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet('202106_GRS');
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    console.log(path_reporting);
    var ligne = 0;
    var max = parseInt(row);
    var min = max - 15;
    console.log('min' + min + 'max' + max);
    colonneDate.eachCell(function(cell, rowNumber) {
     
      if(rowNumber>=min && rowNumber<=max)
      {
        console.log('row'+rowNumber);
        var m = newworksheet.getRow(rowNumber);
        m.getCell(2).value = date_export;
      }
    });
    console.log(ligne);
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO terminé"); 
    return callback(null, "OK");

    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(tab,3,callback);
    }
    },
    ecritureDate1 : async function (tab,date_export,row,callback) {
      const Excel = require('exceljs');
      const newWorkbook = new Excel.Workbook();
      try{
      //var path_reporting = 'D:/Reporting Engagement/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
      await newWorkbook.xlsx.readFile(path_reporting);
      console.log(path_reporting);
      const newworksheet = newWorkbook.getWorksheet('202106_GRS');
      var colonneDate = newworksheet.getColumn('A');
      var ligneDate1;
      console.log(path_reporting);
      var ligne = 0;
      var max = parseInt(row);
      var min = max - 15;
      console.log('max' + min);
      colonneDate.eachCell(function(cell, rowNumber) {
       //console.log(date_export);
        if(rowNumber==min)
        {
          console.log('row'+rowNumber);
          var m = newworksheet.getRow(rowNumber);
          m.getCell(1).value = date_export;
        }
      });
      //console.log(ligne);
      await newWorkbook.xlsx.writeFile(path_reporting);
      sails.log("Ecriture OK KO terminé"); 
      return callback(null, "OK");
  
      }
      catch
      {
        console.log("Une erreur s'est produite");
        Reportinghtp.deleteToutHtp(tab,3,callback);
      }
      },
  ecritureEtp : async function (tab,row,date_export,motcle,nb,callback) {
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    //var path_reporting = 'D:/Reporting Engagement/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet('202106_GRS');
    var colonneDate = newworksheet.getColumn('B');
    var ligneDate1;
    //var date_export='14/06/2021';
    console.log(date_export);
    var ligne = 0;
    var max = parseInt(row);
    var min = max - 15;
    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      if(rowNumber>=min && rowNumber<=max)
      {
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        var bi = motcle[nb];
        const regex = new RegExp(bi,'i');
        if(regex.test(f))
        {
          console.log(rowNumber);
          ligne = rowNumber;
        }
      }
    });
    console.log(ligne);
    
    var m = newworksheet.getRow(ligne);
    //console.log(m);

    m.getCell(5).value = tab[0].nb;
    /*m.getCell(6).value = parseFloat(tab[0].tt16h);
    m.getCell(7).value = parseFloat(tab[0].tt23h);
    m.getCell(9).value = parseFloat(tab[0].ttj2);
    m.getCell(11).value = parseFloat(tab[0].ttj5);
    m.getCell(16).value = parseFloat(tab[0].stock16h);
    m.getCell(20).value = parseFloat(tab[0].bonj);
    m.getCell(21).value = parseFloat(tab[0].bonj1);
    m.getCell(22).value = parseFloat(tab[0].bonj2);
    m.getCell(23).value = parseFloat(tab[0].bonj5);*/
   
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO terminé"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(tab,3,callback);
    }
    },
  countEtp : function (nomcolonne, callback) {
    var sql ="select sum("+nomcolonne+"::float) as nb from tpsgrsetp";
    //,sum(trinument) as trinument,sum(sdpnument) as sdpnument,sum(sdmnument) as sdmnument,sum(factse) as factse,sum(facttiers) as facttiers,sum(factoptique) as factoptique,sum(factaudio) as factaudio,sum(factdentaire) as factdentaire, sum(facthospi) as facthospi,sum(santeclair) as santeclair,sum(pecoptique) as pecoptique,sum(pecaudio) as pecaudio,sum(pecdentaire) as pecdentaire,sum(pechospi) as pechospi from tpsgrsetp2 ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        console.log(sql);
        result = res.rows;
        console.log(result[0].nb);
        return callback(null,result);
      };
      });
  },
  ecriture : async function (tab,row,date_export,motcle,nb,callback) {
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    //var path_reporting = 'D:/Reporting Engagement/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
    console.log(path_reporting);
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet('202106_GRS');
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    //var date_export='14/06/2021';
    console.log(date_export);
    var ligne = 0;
    var max = parseInt(row);
    var min = max - 15;
    console.log(max);
    colonneDate.eachCell(function(cell, rowNumber) {
      if(rowNumber>=min && rowNumber<=max)
      {
        console.log('rownumber' + rowNumber);
        ligneDate1 = parseInt(rowNumber);
        var line = newworksheet.getRow(ligneDate1);
        var f = line.getCell(4).value;
        var bi = motcle[nb];
        const regex = new RegExp(bi,'i');
        if(regex.test(f))
        {
          console.log('row'+rowNumber);
          ligne = rowNumber;
        }
      }
    });
    //console.log(ligne);
    
    var m = newworksheet.getRow(ligne);
   //m.getCell(5).value = tab[0].tt16h;
    m.getCell(6).value = parseFloat(tab[0].tt16h);
    m.getCell(7).value = parseFloat(tab[0].tt23h);
    m.getCell(9).value = parseFloat(tab[0].ttj2);
    m.getCell(11).value = parseFloat(tab[0].ttj5);
    m.getCell(16).value = parseFloat(tab[0].stock16h);
    m.getCell(20).value = parseFloat(tab[0].bonj);
    m.getCell(21).value = parseFloat(tab[0].bonj1);
    m.getCell(22).value = parseFloat(tab[0].bonj2);
    m.getCell(23).value = parseFloat(tab[0].bonj5);
   
    await newWorkbook.xlsx.writeFile(path_reporting);
    sails.log("Ecriture OK KO terminé"); 
    return callback(null, "OK");
  
    }
    catch
    {
      console.log("Une erreur s'est produite");
      Reportinghtp.deleteToutHtp(tab,3,callback);
    }
    },
    ecriture3 : async function (tab,row,date_export,motcle,nb,callback) {
      const Excel = require('exceljs');
      const newWorkbook = new Excel.Workbook();
      try{
      //var path_reporting = 'D:/Reporting Engagement/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
      console.log(path_reporting);
      await newWorkbook.xlsx.readFile(path_reporting);
      const newworksheet = newWorkbook.getWorksheet('202106_GRS');
      var colonneDate = newworksheet.getColumn('A');
      var ligneDate1;
      //var date_export='14/06/2021';
      console.log(date_export);
      var ligne = 0;
      var max = parseInt(row);
      var min = max - 15;
      console.log(max);
      colonneDate.eachCell(function(cell, rowNumber) {
        if(rowNumber>=min && rowNumber<=max)
        {
          console.log('rownumber' + rowNumber);
          ligneDate1 = parseInt(rowNumber);
          var line = newworksheet.getRow(ligneDate1);
          var f = line.getCell(4).value;
          var bi = motcle[nb];
          const regex = new RegExp(bi,'i');
          if(regex.test(f))
          {
            console.log('row'+rowNumber);
            ligne = rowNumber;
          }
        }
      });
      //console.log(ligne);
      
      var m = newworksheet.getRow(ligne);
     
      if(motcle[nb]=='Tri TP' || motcle[nb]=='Tri Nument'  || motcle[nb]=='SDP')
      {
        m.getCell(6).value = parseFloat(tab[0].tt16h);
        m.getCell(7).value = parseFloat(tab[0].tt23h);
        m.getCell(16).value = parseFloat(tab[0].stock16h);
        m.getCell(20).value = parseFloat(tab[0].bonj);
        m.getCell(21).value = parseFloat(tab[0].bonj1);
        m.getCell(22).value = parseFloat(tab[0].bonj2);
        m.getCell(23).value = parseFloat(tab[0].bonj5);
      }
      else
      {
        m.getCell(6).value = parseFloat(tab[0].tt16h);
        m.getCell(7).value = parseFloat(tab[0].tt23h);
        m.getCell(9).value = parseFloat(tab[0].ttj2);
        m.getCell(11).value = parseFloat(tab[0].ttj5);
        m.getCell(16).value = parseFloat(tab[0].stock16h);
        m.getCell(20).value = parseFloat(tab[0].bonj);
        m.getCell(21).value = parseFloat(tab[0].bonj1);
        m.getCell(22).value = parseFloat(tab[0].bonj2);
        m.getCell(23).value = parseFloat(tab[0].bonj5);
      }
     
     
      await newWorkbook.xlsx.writeFile(path_reporting);
      sails.log("Ecriture OK KO terminé"); 
      return callback(null, "OK");
    
      }
      catch
      {
        console.log("Une erreur s'est produite");
        Reportinghtp.deleteToutHtp(tab,3,callback);
      }
      },
  traitementInsertionstockbonJ5:function(ast,traitement,motcle1,motcle2,motcle3,motcle4,nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    try{
      //console.log(ast[nb]);
      //console.log(traitement[nb]);
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      if(jour=='Friday')
      {
        console.log('Friday');
        for(var ra=0;ra<=range.e.r;ra++)
        {
          //nature de tache
          var address_of_cell = {c:1, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:6, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          //etat de la tache
          var address_of_cell2 = {c:0, r:ra};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
          var b2 = "ETAT1";
          var b3= "ETAT4";
          const regex2 = new RegExp(b2,'i');
          const regex3 = new RegExp(b3,'i');

          //etat facture any amle undefined
          var address_of_cell4 = {c:9, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

           //date de maturit5
           var address_of_cell5 = {c:4, r:ra};
           var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
           var desired_cell5 = sheet[cell_ref5];
           var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);

          if(regex2.test(desired_value2))
          {
            if(desired_value4!=undefined)
            {
              var j = 1;
            }
            else
            {
              var b = traitement[nb];
              const regex = new RegExp(b,'i');
              var conv = parseInt(date);
              var j2 = conv - 5;
              if(regex.test(desired_value) && parseInt(desired_value5)<=j2 )
              {
              
                var c = motcle1[nb];
                var c1 = motcle2[nb];
                var c2 = motcle3[nb];
                var c5 = ast[nb];

                const regex21 = new RegExp(c1,'i');
                const regex31 = new RegExp(c2,'i');
                const regex1 = new RegExp(c,'i');
                const regex41 = new RegExp(c5,'i');

             
                if(motcle4[nb]=='a')
                {
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex41.test(desired_value)) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
                else if(motcle4[nb]=='b')
                {
                  somme=somme+1;
                }
                else
                {
                  var c4 = motcle4[nb];
                  const regex4 = new RegExp(c4,'i');
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value)  || regex41.test(desired_value) ) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
              }
            }
          }
            else
            {
              var r= 'a';
            }

            }
           // console.log(somme);
            var sql = "insert into "+table[nb]+" (bonj5) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log(err);
                          //return callback(err); 
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
              console.log('hafa');
              for(var ra=0;ra<=range.e.r;ra++)
              {
                //nature de tache
                var address_of_cell = {c:1, r:ra};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value = (desired_cell ? desired_cell.v : undefined);

                // identification de la tache : le isaina
                var address_of_cell1 = {c:6, r:ra};
                var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
                var desired_cell1 = sheet[cell_ref1];
                var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

                //etat de la tache
                var address_of_cell2 = {c:0, r:ra};
                var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
                var desired_cell2 = sheet[cell_ref2];
                var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
                var b2 = "ETAT1";
                var b3= "ETAT4";
                const regex2 = new RegExp(b2,'i');
                //const regex3 = new RegExp(b3,'i');

                //etat facture any amle undefined
                var address_of_cell4 = {c:9, r:ra};
                var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
                var desired_cell4 = sheet[cell_ref4];
                var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

                //date de maturit5
                var address_of_cell5 = {c:4, r:ra};
                var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
                var desired_cell5 = sheet[cell_ref5];
                var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);

                if(regex2.test(desired_value2) )
                {
                  if(desired_value4!=undefined)
                  {
                    var j = 1;
                  }
                  else
                  {
                    var b = traitement[nb];
                    const regex = new RegExp(b,'i');
                    var conv = parseInt(date);
                    var j2 = conv - 7;
                    if(regex.test(desired_value) && parseInt(desired_value5)<=j2 )
                    {
                     
                      var c = motcle1[nb];
                      var c1 = motcle2[nb];
                      var c2 = motcle3[nb];
                      var c5 = ast[nb];
      
                      const regex21 = new RegExp(c1,'i');
                      const regex31 = new RegExp(c2,'i');
                      const regex1 = new RegExp(c,'i');
                      const regex41 = new RegExp(c5,'i');
                  
                      if(motcle4[nb]=='a')
                      {
                        if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value)  || regex41.test(desired_value)) 
                        {
                          var a = '1';
                        }
                        else
                        {
                          if(desired_value1!=undefined)
                          {
                            somme=somme+1;
                          }
                          else
                          {
                            var p = 0;
                          }
                        
                        };
                      }
                      else if(motcle4[nb]=='b')
                      {
                        somme=somme+1;
                      }
                      else
                      {
                        var c4 = motcle4[nb];
                        const regex4 = new RegExp(c4,'i');
                        if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value) || regex41.test(desired_value) ) 
                        {
                          var a = '1';
                        }
                        else
                        {
                          if(desired_value1!=undefined)
                          {
                            somme=somme+1;
                          }
                          else
                          {
                            var p = 0;
                          }
                        
                        };
                      }
                    }
                  }
                }
                  else
                  {
                    var r= 'a';
                  }
                  }
                  console.log(somme);
                  var sql = "insert into "+table[nb]+" (bonj5) values ("+somme+") ";
                            Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                              if (err) { 
                                console.log(err);
                                //return callback(err); 
                              }
                              else
                              {
                                console.log(sql);
                                return callback(null, true);
                              }
                            
                                                  });
                                                }
          
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
  },
  traitementInsertionstockbonJ2:function(ast,traitement,motcle1,motcle2,motcle3,motcle4,nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    try{
      //console.log(ast[nb]);
      //console.log(traitement[nb]);
      
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      
      if(jour=='Monday' || jour=='Tuesday')
      {
        console.log('Monday');
        for(var ra=0;ra<=range.e.r;ra++)
        {
          //nature de tache
          var address_of_cell = {c:1, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:6, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          //etat de la tache
          var address_of_cell2 = {c:0, r:ra};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
          var b2 = "ETAT1";
          var b3= "ETAT4";
          const regex2 = new RegExp(b2,'i');
          //const regex3 = new RegExp(b3,'i');

          //etat facture any amle undefined
          var address_of_cell4 = {c:9, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

           //date de maturit5
           var address_of_cell5 = {c:4, r:ra};
           var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
           var desired_cell5 = sheet[cell_ref5];
           var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);

          if(regex2.test(desired_value2) )
          {
            if(desired_value4!=undefined)
            {
              var j = 1;
            }
            else
            {
              var b = traitement[nb];
              const regex = new RegExp(b,'i');
              var conv = parseInt(date);
              var j2 = conv - 4;
              if(regex.test(desired_value) && parseInt(desired_value5)<=j2 )
              {
              
                var c = motcle1[nb];
                var c1 = motcle2[nb];
                var c2 = motcle3[nb];
                var c5 = ast[nb];

                const regex21 = new RegExp(c1,'i');
                const regex31 = new RegExp(c2,'i');
                const regex1 = new RegExp(c,'i');
                const regex41 = new RegExp(c5,'i');
                if(motcle4[nb]=='a')
                {
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex41.test(desired_value) ) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
                else if(motcle4[nb]=='b')
                {
                  somme=somme+1;
                }
                else
                {
                  var c4 = motcle4[nb];
                  const regex4 = new RegExp(c4,'i');
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value)  || regex41.test(desired_value)  ) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
              }
            }
          }
            else
            {
              var r= 'a';
            }

            }
            console.log(somme);
            var sql = "insert into "+table[nb]+" (bonj2) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log(err);
                          //return callback(err); 
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
              console.log('hafa');
              for(var ra=0;ra<=range.e.r;ra++)
              {
                //nature de tache
                var address_of_cell = {c:1, r:ra};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value = (desired_cell ? desired_cell.v : undefined);

                // identification de la tache : le isaina
                var address_of_cell1 = {c:6, r:ra};
                var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
                var desired_cell1 = sheet[cell_ref1];
                var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

                //etat de la tache
                var address_of_cell2 = {c:0, r:ra};
                var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
                var desired_cell2 = sheet[cell_ref2];
                var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
                var b2 = "ETAT1";
                var b3= "ETAT4";
                const regex2 = new RegExp(b2,'i');
                const regex3 = new RegExp(b3,'i');

                //etat facture any amle undefined
                var address_of_cell4 = {c:9, r:ra};
                var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
                var desired_cell4 = sheet[cell_ref4];
                var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

                //date de maturit5
                var address_of_cell5 = {c:4, r:ra};
                var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
                var desired_cell5 = sheet[cell_ref5];
                var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);

                if(regex2.test(desired_value2) )
                {
                  if(desired_value4!=undefined)
                  {
                    var j = 1;
                  }
                  else
                  {
                    var b = traitement[nb];
                    const regex = new RegExp(b,'i');
                    var conv = parseInt(date);
                    var j2 = conv - 2;
                    if(regex.test(desired_value) && parseInt(desired_value5)<=j2 )
                    {
                      var c = motcle1[nb];
                      var c1 = motcle2[nb];
                      var c2 = motcle3[nb];
                      var c5 = ast[nb];
                      const regex21 = new RegExp(c1,'i');
                      const regex31 = new RegExp(c2,'i');
                      const regex1 = new RegExp(c,'i');
                      const regex41 = new RegExp(c5,'i');
                  
                      if(motcle4[nb]=='a')
                      {
                        if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex41.test(desired_value)) 
                        {
                          var a = '1';
                        }
                        else
                        {
                          if(desired_value1!=undefined)
                          {
                            somme=somme+1;
                          }
                          else
                          {
                            var p = 0;
                          }
                        
                        };
                      }
                      else if(motcle4[nb]=='b')
                      {
                        somme=somme+1;
                      }
                      else
                      {
                        var c4 = motcle4[nb];
                        const regex4 = new RegExp(c4,'i');
                        if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value) || regex41.test(desired_value)) 
                        {
                          var a = '1';
                        }
                        else
                        {
                          if(desired_value1!=undefined)
                          {
                            somme=somme+1;
                          }
                          else
                          {
                            var p = 0;
                          }
                        
                        };
                      }
                    }
                  }
                }
                  else
                  {
                    var r= 'a';
                  }
                  }
                  console.log(somme);
                  var sql = "insert into "+table[nb]+" (bonj2) values ("+somme+") ";
                            Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                              if (err) { 
                                console.log(err);
                                //return callback(err); 
                              }
                              else
                              {
                                console.log(sql);
                                return callback(null, true);
                              }
                            
                                                  });
                                                }
          
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
  },
  traitementInsertionstockbonJ1:function(ast,traitement,motcle1,motcle2,motcle3,motcle4,nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      
      if(jour=='Tuesday')
      {
        console.log('Tuesday');
        for(var ra=0;ra<=range.e.r;ra++)
        {
          //nature de tache
          var address_of_cell = {c:1, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:6, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          //etat de la tache
          var address_of_cell2 = {c:0, r:ra};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
          var b2 = "ETAT1";
          var b3= "ETAT4";
          const regex2 = new RegExp(b2,'i');
          //const regex3 = new RegExp(b3,'i');

          //etat facture any amle undefined
          var address_of_cell4 = {c:9, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

           //date de maturit5
           var address_of_cell5 = {c:4, r:ra};
           var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
           var desired_cell5 = sheet[cell_ref5];
           var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);

          if(regex2.test(desired_value2) )
          {
            if(desired_value4!=undefined)
            {
              var j = 1;
            }
            else
            {
              var b = traitement[nb];
              const regex = new RegExp(b,'i');
              
              var conv = parseInt(date);
              var j1 = conv - 1;
              var j2 = conv -2;
              var j3 = conv - 3;
              if(regex.test(desired_value) && (desired_value5==j1 || desired_value5==j2 || desired_value5==j3) )
              {
              
                var c = motcle1[nb];
                var c1 = motcle2[nb];
                var c2 = motcle3[nb];
                var c5 = ast[nb];

                const regex21 = new RegExp(c1,'i');
                const regex31 = new RegExp(c2,'i');
                const regex1 = new RegExp(c,'i');
                const regex41 = new RegExp(c5,'i');

             
                if(motcle4[nb]=='a')
                {
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex41.test(desired_value)) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
                else if(motcle4[nb]=='b')
                {
                  somme=somme+1;
                }
                else
                {
                  var c4 = motcle4[nb];
                  const regex4 = new RegExp(c4,'i');
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value)  || regex41.test(desired_value) ) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
              }
            }
          }
            else
            {
              var r= 'a';
            }

            }
            console.log(somme);
            var sql = "insert into "+table[nb]+" (bonj1) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log(err);
                          //return callback(err); 
                        }
                        else
                        {
                          console.log(sql);
                          return callback(null, true);
                        }
                       
                                            });
          }
          else if(jour=='Monday')   
          {
            console.log('Monday');
              for(var ra=0;ra<=range.e.r;ra++)
              {
                //nature de tache
                var address_of_cell = {c:1, r:ra};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value = (desired_cell ? desired_cell.v : undefined);

                // identification de la tache : le isaina
                var address_of_cell1 = {c:6, r:ra};
                var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
                var desired_cell1 = sheet[cell_ref1];
                var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

                //etat de la tache
                var address_of_cell2 = {c:0, r:ra};
                var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
                var desired_cell2 = sheet[cell_ref2];
                var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
                var b2 = "ETAT1";
                var b3= "ETAT4";
                const regex2 = new RegExp(b2,'i');
                const regex3 = new RegExp(b3,'i');

                //etat facture any amle undefined
                var address_of_cell4 = {c:9, r:ra};
                var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
                var desired_cell4 = sheet[cell_ref4];
                var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

                //date de maturit5
                var address_of_cell5 = {c:4, r:ra};
                var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
                var desired_cell5 = sheet[cell_ref5];
                var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);

                if(regex2.test(desired_value2) || regex3.test(desired_value2) )
                {
                  if(desired_value4!=undefined)
                  {
                    var j = 1;
                  }
                  else
                  {
                    var b = traitement[nb];
                    const regex = new RegExp(b,'i');
                    var conv = parseInt(date);
                    var j1 = conv - 3;
                    if(regex.test(desired_value) && desired_value5==j1 )
                    {
                    
                      var c = motcle1[nb];
                      var c1 = motcle2[nb];
                      var c2 = motcle3[nb];
                      var c5 = ast[nb];
                      const regex21 = new RegExp(c1,'i');
                      const regex31 = new RegExp(c2,'i');
                      const regex1 = new RegExp(c,'i');
                      const regex41 = new RegExp(c5,'i');
                  
                      if(motcle4[nb]=='a')
                      {
                        if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex41.test(desired_value)) 
                        {
                          var a = '1';
                        }
                        else
                        {
                          if(desired_value1!=undefined)
                          {
                            somme=somme+1;
                          }
                          else
                          {
                            var p = 0;
                          }
                        
                        };
                      }
                      else if(motcle4[nb]=='b')
                      {
                        somme=somme+1;
                      }
                      else
                      {
                        var c4 = motcle4[nb];
                        const regex4 = new RegExp(c4,'i');
                        if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value) || regex41.test(desired_value)) 
                        {
                          var a = '1';
                        }
                        else
                        {
                          if(desired_value1!=undefined)
                          {
                            somme=somme+1;
                          }
                          else
                          {
                            var p = 0;
                          }
                        
                        };
                      }
                    }
                  }
                }
                  else
                  {
                    var r= 'a';
                  }
                  }
                  console.log(somme);
                  var sql = "insert into "+table[nb]+" (bonj1) values ("+somme+") ";
                            Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                              if (err) { 
                                console.log(err);
                                //return callback(err); 
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
              console.log('hafa');
              for(var ra=0;ra<=range.e.r;ra++)
              {
                //nature de tache
                var address_of_cell = {c:1, r:ra};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value = (desired_cell ? desired_cell.v : undefined);

                // identification de la tache : le isaina
                var address_of_cell1 = {c:6, r:ra};
                var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
                var desired_cell1 = sheet[cell_ref1];
                var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

                //etat de la tache
                var address_of_cell2 = {c:0, r:ra};
                var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
                var desired_cell2 = sheet[cell_ref2];
                var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
                var b2 = "ETAT1";
                var b3= "ETAT4";
                const regex2 = new RegExp(b2,'i');
                const regex3 = new RegExp(b3,'i');

                //etat facture any amle undefined
                var address_of_cell4 = {c:9, r:ra};
                var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
                var desired_cell4 = sheet[cell_ref4];
                var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

                //date de maturit5
                var address_of_cell5 = {c:4, r:ra};
                var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
                var desired_cell5 = sheet[cell_ref5];
                var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);

                if(regex2.test(desired_value2) || regex3.test(desired_value2) )
                {
                  if(desired_value4!=undefined)
                  {
                    var j = 1;
                  }
                  else
                  {
                    var b = traitement[nb];
                    const regex = new RegExp(b,'i');
                    var conv = parseInt(date);
                    var j1 = conv - 1;
                    if(regex.test(desired_value) && desired_value5==j1 )
                    {
                    
                      var c = motcle1[nb];
                      var c1 = motcle2[nb];
                      var c2 = motcle3[nb];
                      var c5 = ast[nb];
                      const regex21 = new RegExp(c1,'i');
                      const regex31 = new RegExp(c2,'i');
                      const regex1 = new RegExp(c,'i');
                      const regex41 = new RegExp(c5,'i');
                  
                      if(motcle4[nb]=='a')
                      {
                        if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex41.test(desired_value)) 
                        {
                          var a = '1';
                        }
                        else
                        {
                          if(desired_value1!=undefined)
                          {
                            somme=somme+1;
                          }
                          else
                          {
                            var p = 0;
                          }
                        
                        };
                      }
                      else if(motcle4[nb]=='b')
                      {
                        somme=somme+1;
                      }
                      else
                      {
                        var c4 = motcle4[nb];
                        const regex4 = new RegExp(c4,'i');
                        if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value) || regex41.test(desired_value)) 
                        {
                          var a = '1';
                        }
                        else
                        {
                          if(desired_value1!=undefined)
                          {
                            somme=somme+1;
                          }
                          else
                          {
                            var p = 0;
                          }
                        
                        };
                      }
                    }
                  }
                }
                  else
                  {
                    var r= 'a';
                  }
                  }
                  console.log(somme);
                  var sql = "insert into "+table[nb]+" (bonj1) values ("+somme+") ";
                            Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                              if (err) { 
                                console.log(err);
                                //return callback(err); 
                              }
                              else
                              {
                                console.log(sql);
                                return callback(null, true);
                              }
                            
                                                  });
                                                }
          
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
  },
  traitementInsertionstockbonJ:function(ast,traitement,motcle1,motcle2,motcle3,motcle4,nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    try{
      //console.log(ast[nb]);
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;

      if(jour=='Monday')
      {
        for(var ra=0;ra<=range.e.r;ra++)
        {
          //nature de tache
          var address_of_cell = {c:1, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:6, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          //etat de la tache
          var address_of_cell2 = {c:0, r:ra};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
          var b2 = "ETAT1";
          var b3= "ETAT4";
          const regex2 = new RegExp(b2,'i');
          //const regex3 = new RegExp(b3,'i');

          //etat facture any amle undefined
          var address_of_cell4 = {c:9, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

          //date de maturit5
          var address_of_cell5 = {c:4, r:ra};
          var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
          var desired_cell5 = sheet[cell_ref5];
          var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);

          if(regex2.test(desired_value2) )
          {
            if(desired_value4!=undefined)
            {
              var j = 1;
            }
            else
            {
              var b = traitement[nb];
              const regex = new RegExp(b,'i');
              var j1 = date;
              var conv = parseInt(date);
              var j2 = conv - 1;
              var j3 = conv -2;
              if(regex.test(desired_value) && (desired_value5==j1 || desired_value5==j2 || desired_value5==j3 ) )
              {
              
                var c = motcle1[nb];
                var c1 = motcle2[nb];
                var c2 = motcle3[nb];
                var c5 = ast[nb];
                const regex21 = new RegExp(c1,'i');
                const regex31 = new RegExp(c2,'i');
                const regex1 = new RegExp(c,'i');
                const regex41 = new RegExp(c5,'i');
            
                if(motcle4[nb]=='a')
                {
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex41.test(desired_value)) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
                else if(motcle4[nb]=='b')
                {
                  somme=somme+1;
                }
                else
                {
                  var c4 = motcle4[nb];
                  const regex4 = new RegExp(c4,'i');
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value) || regex41.test(desired_value)) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
              }
            }
          }
            else
            {
              var r= 'a';
            }
            }
            console.log(somme);
            var sql = "insert into "+table[nb]+" (bonj) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log(err);
                          //return callback(err); 
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
        for(var ra=0;ra<=range.e.r;ra++)
        {
          //nature de tache
          var address_of_cell = {c:1, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:6, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          //etat de la tache
          var address_of_cell2 = {c:0, r:ra};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
          var b2 = "ETAT1";
          var b3= "ETAT4";
          const regex2 = new RegExp(b2,'i');
          //const regex3 = new RegExp(b3,'i');

          //etat facture any amle undefined
          var address_of_cell4 = {c:9, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

          //date de maturit5
          var address_of_cell5 = {c:4, r:ra};
          var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
          var desired_cell5 = sheet[cell_ref5];
          var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);

          if(regex2.test(desired_value2) )
          {
            if(desired_value4!=undefined)
            {
              var j = 1;
            }
            else
            {
              var b = traitement[nb];
              const regex = new RegExp(b,'i');
              var j1 = date;
              var conv = parseInt(date);
              var j2 = conv - 1;
              var j3 = conv -2;
              if(regex.test(desired_value) && desired_value5==j1 )
              {
              
                var c = motcle1[nb];
                var c1 = motcle2[nb];
                var c2 = motcle3[nb];
                var c5 = ast[nb];
                const regex21 = new RegExp(c1,'i');
                const regex31 = new RegExp(c2,'i');
                const regex1 = new RegExp(c,'i');
                const regex41 = new RegExp(c5,'i');
            
                if(motcle4[nb]=='a')
                {
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex41.test(desired_value)) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
                else if(motcle4[nb]=='b')
                {
                  somme=somme+1;
                }
                else
                {
                  var c4 = motcle4[nb];
                  const regex4 = new RegExp(c4,'i');
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value) || regex41.test(desired_value)) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
              }
            }
          }
            else
            {
              var r= 'a';
            }
            }
            console.log(somme);
            var sql = "insert into "+table[nb]+" (bonj) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log(err);
                          //return callback(err); 
                        }
                        else
                        {
                          console.log(sql);
                          return callback(null, true);
                        }
                      
                                            });
                                          
      }
             
          
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
  },
  traitementInsertionstock16h:function(ast,traitement,motcle1,motcle2,motcle3,motcle4,nb,date,jour,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    try{
     
      /*console.log(traitement[nb]);
      console.log(motcle3[nb]);
      console.log(motcle1[nb]);
      console.log(motcle2[nb]);*/
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      for(var ra=0;ra<=range.e.r;ra++)
        {
          //nature de tache
          var address_of_cell = {c:1, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:6, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          //etat de la tache
          var address_of_cell2 = {c:0, r:ra};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
          var b2 = "ETAT1";
          var b3= "ETAT4";
          const regex2 = new RegExp(b2,'i');
          //const regex3 = new RegExp(b3,'i');

          //etat facture any amle undefined
          var address_of_cell4 = {c:9, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

          if(regex2.test(desired_value2))
          {
            if(desired_value4!=undefined)
            {
              var j = 1;
            }
            else
            {
              var b = traitement[nb];
              const regex = new RegExp(b,'i');
              if(regex.test(desired_value) )
              {
              
                var c = motcle1[nb];
                var c1 = motcle2[nb];
                var c2 = motcle3[nb];
                var c5 = ast[nb];
                const regex21 = new RegExp(c1,'i');
                const regex31 = new RegExp(c2,'i');
                const regex1 = new RegExp(c,'i');
                const regex41 = new RegExp(c5,'i');
             
                if(motcle4[nb]=='a')
                {
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex41.test(desired_value) ) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
                else if(motcle4[nb]=='b')
                {
                  somme=somme+1;
                }
                else
                {
                  var c4 = motcle4[nb];
                  const regex4 = new RegExp(c4,'i');
                  if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value) || regex41.test(desired_value)) 
                  {
                    var a = '1';
                  }
                  else
                  {
                    if(desired_value1!=undefined)
                    {
                      somme=somme+1;
                    }
                    else
                    {
                      var p = 0;
                    }
                  
                  };
                }
              }
            }
          }
            else
            {
              var r= 'a';
            }

            }
            console.log(somme);
            var sql = "insert into "+table[nb]+" (stock16h) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log(err);
                          //return callback(err); 
                        }
                        else
                        {
                          console.log(sql);
                          return callback(null, true);
                        }
                       
                                            });
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
  },

  traitementInsertion:function(ast,traitement,motcle1,motcle2,motcle3,motcle4,nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    //var trameflux= "D:/Reporting Engagement/Prod_16HF.xlsb";
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      for(var ra=0;ra<=range.e.r;ra++)
        {
          var address_of_cell = {c:2, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          var address_of_cell1 = {c:1, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          var address_of_cell2 = {c:0, r:ra};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);

          var address_of_cell4 = {c:4, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);
          
          var b2 = "ALMERYS";
          var b3= "CBTP";
          var b4 = traitement[nb];
          const regex2 = new RegExp(b2,'i');
          const regex3 = new RegExp(b3,'i');
          const regex4 = new RegExp(b4,'i');

          var b1 = 'ASTT6';
          
          const regex1 = new RegExp(b1);
          if(regex1.test(desired_value1) && (regex2.test(desired_value2) || regex3.test(desired_value2))  )
          {
           var b = traitement[nb];
           const regex = new RegExp(b,'i');
            if(regex.test(desired_value) )
            {
              
              var c = motcle1[nb];
              var c1 = motcle2[nb];
              var c2 = motcle3[nb];
              
              const regex21 = new RegExp(c1,'i');
              const regex31 = new RegExp(c2,'i');
              const regex1 = new RegExp(c,'i');
             
              if(motcle4[nb]=='a')
              {
                if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) ) 
                {
                  var a = '1';
                }
                else
                {
                  if(desired_value4!=undefined)
                  {
                    somme=somme+1;
                  }
                  else
                  {
                    var p = 0;
                  }
                 
                };
              }
              else if(motcle4[nb]=='b')
              {
                somme=somme+1;
              }
              else
              {
                var c4 = motcle4[nb];
                const regex4 = new RegExp(c4,'i');
                if(regex1.test(desired_value) || regex21.test(desired_value) || regex31.test(desired_value) || regex4.test(desired_value) ) 
                {
                  var a = '1';
                }
                else
                {
                  if(desired_value4!=undefined)
                  {
                    somme=somme+1;
                  }
                  else
                  {
                    var p = 0;
                  }
                 
                };
              }
              
            }
            else
            {
              var r= 'a';
            }
          }
          else if(regex1.test(desired_value1) && regex4.test(desired_value2) )
          {
            somme=somme+1;
          }
          else
          {
           var a = 'l';
          }
        };
        console.log(somme);
        var sql = "insert into "+table[nb]+" (tt16h) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
                          console.log(err);
                          //return callback(err); 
                        }
                        else
                        {
                          console.log(sql);
                          return callback(null, true);
                        }
                       
                                            });
      
    }
    catch
    {
      console.log("erreur absolu haaha");
    }
  },
  delete : function (table,callback) {
    var sql = "delete from "+table+" ";
    ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        //return callback(err); 
        console.log('une erreur de suppression');
      }
      else{
        console.log(sql);
        return callback(null, true);
      }
    
      });
  },
  copieEtp:function(date,nb,trameflux,nomColonne,callback){
    XLSX = require('xlsx');
    var workbook = XLSX.readFile(trameflux);
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var ligne = 0;
      for(var ra=0;ra<=range.e.r;ra++)
        {
          //nature de tache
          var address_of_cell = {c:0, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.w : undefined);

          if(desired_value==date)
          {
            console.log('vaelur :' + ra);
            col = parseInt(ra) + parseInt(nb);
          }
          else{
              var b = 41;
          }
        }
        var address_of_cell1 = {c:4, r:col};
        var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
        var desired_cell1 = sheet[cell_ref1];
        var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

        console.log(desired_value1);
        var sql = "insert into tpsgrsetp ("+nomColonne[nb]+") values ("+desired_value1+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log(err);
                          //return callback(err); 
                        }
                        else
                        {
                          console.log(sql);
                          return callback(null, true);
                        }
                       
                                            });
    }
    catch
    {
      console.log("erreur absolu haaha");
    } 
  },
};

