/**
 * TpsGrs.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */

module.exports = {

  attributes: {
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

