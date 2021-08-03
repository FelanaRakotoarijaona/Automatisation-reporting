/**
 * Tpstpc.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */
 const path_reporting = '/dev/prod/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/TestReporting/Copie de TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
 //const path_reporting = '//10.128.1.2/bpo_almerys/03-POLE_TPS-TPC/00-PILOTAGE/09-REPORTING ENGAGEMENT/TestReporting/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
 module.exports = {

  datastore : 'easy',
  attributes: {
  },
  ecritureDate : async function (tab,date_export,callback) {
    const Excel = require('exceljs');
    const newWorkbook = new Excel.Workbook();
    try{
    //var path_reporting = 'D:/Reporting Engagement/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
    await newWorkbook.xlsx.readFile(path_reporting);
    const newworksheet = newWorkbook.getWorksheet('202106_Easy');
    var colonneDate = newworksheet.getColumn('A');
    var ligneDate1;
    //var date_export='14/06/2021';
    console.log(date_export);
    var ligne = 0;

    colonneDate.eachCell(function(cell, rowNumber) {
      var dateExcel = ReportingInovcomExport.convertDate(cell.text);
      //console.log(dateExcel + 'date')
      //if(rowNumber==3685)
      //if(rowNumber>=3685 && rowNumber<=3700)
      if(dateExcel==date_export)
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
      return callback(null,'KO');
    }
    },
  importfichier: function (cheminfinal,motcle,table,nb,callback) {
   const fs = require('fs');
   var re  = 'a';
   var tab = [];
   var chemin = cheminfinal[nb];
   console.log(chemin);
   var c = ReportingInovcom.existenceFichier(chemin);
   var motcle1 = motcle[nb];
   var tab = table[nb];
   console.log(c);
   var cheminbase ;
   if(c=='vrai')
   {
     fs.readdir(chemin, (err, files) => {
       console.log(chemin);
           files.forEach(file => {
             const regex = new RegExp(motcle1,'i');
             var m1 = '.xlsx|.xls|.xlsm|.xlsb$';
             var m2 = '^[^~]';
             const regex1 = new RegExp(m1,'i');
             const regex2 = new RegExp(m2);
             if(regex.test(file) && regex1.test(file) && regex2.test(file))
             {
              cheminbase = chemin + '/' + file;
              console.log(cheminbase);  
             } 
         });
         console.log('table' +tab);
         var sql = "insert into "+tab+" (chemin) values ('"+cheminbase+"') ";
                 ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
                  if (err) { 
                    //console.log(err);
                    return callback(err);
                   }
                  else
                  {
                    console.log(sql);
                    return callback(null, true);
                  };
                                       }) ; 
         console.log('ato anatiny'+cheminbase);
       });
   }
   else
   {
     var sql = "insert into chemintsisy  values ('k') ";
     ReportingInovcom.getDatastore().sendNativeQuery(sql, function(err,res){
      if (err) { 
        console.log(err);
       }
      else
      {
        console.log(sql);
        return callback(null, true);
      };    
                           })   
   }   
 },
  selection: function (identifiant,identifiant1,identifiant2,date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = " select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree,  p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = "+identifiant+"  AND almerys_lien_ss_spe.id_alm_ss_spe="+identifiant1+" AND almerys_lien_ss_spe.id_lien_ss_spe2="+identifiant2+" AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers";
    //var sql = "select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree  from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = "+identifiant+"  AND almerys_lien_ss_spe.id_alm_ss_spe="+identifiant1+"  AND almerys_lien_ss_spe.id_lien_ss_spe2="+identifiant2+"   AND date_deb_ldt = '"+date+"' AND id_type_ldt = 0 group by almerys_lien_ss_spe.id_almerys ";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
  selectionSanteclair : function (date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = "select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree,  p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = 36137  AND almerys_lien_ss_spe.id_alm_ss_spe=929  AND almerys_lien_ss_spe.id_lien_ss_spe2=1199  AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers ";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
  selectionFactTiers : function (date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = "select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree,  p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = 36140  AND (almerys_lien_ss_spe.id_alm_ss_spe=939 OR almerys_lien_ss_spe.id_alm_ss_spe=941)  AND (almerys_lien_ss_spe.id_lien_ss_spe2=1229 OR almerys_lien_ss_spe.id_lien_ss_spe2=1232) AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
  selectionSE : function (date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = " select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree,  p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = 36141  AND (almerys_lien_ss_spe.id_alm_ss_spe=942 OR almerys_lien_ss_spe.id_alm_ss_spe=945)  AND (almerys_lien_ss_spe.id_lien_ss_spe2=1236 OR almerys_lien_ss_spe.id_lien_ss_spe2=1235 OR almerys_lien_ss_spe.id_lien_ss_spe2=1240) AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers ";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
  selectionFactDentaire : function (date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = "select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree, p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = 36139  AND (almerys_lien_ss_spe.id_alm_ss_spe=935 OR almerys_lien_ss_spe.id_alm_ss_spe=938)  AND (almerys_lien_ss_spe.id_lien_ss_spe2=1219 OR almerys_lien_ss_spe.id_lien_ss_spe2=1225) AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
  selectionFactHospi : function (date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = "select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree,  p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = 36142  AND (almerys_lien_ss_spe.id_alm_ss_spe=946 OR almerys_lien_ss_spe.id_alm_ss_spe=950)  AND (almerys_lien_ss_spe.id_lien_ss_spe2=1244 OR almerys_lien_ss_spe.id_lien_ss_spe2=1243 OR almerys_lien_ss_spe.id_lien_ss_spe2=1251)  AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
  selectionNument : function (date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = "select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree,    p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = 36141  AND almerys_lien_ss_spe.id_alm_ss_spe=954  AND almerys_lien_ss_spe.id_lien_ss_spe2=1269  AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
  selectionFactOpt : function (date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = "select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree,    p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = 36137  AND almerys_lien_ss_spe.id_alm_ss_spe=925  AND almerys_lien_ss_spe.id_lien_ss_spe2=1189  AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
  selectionPecOptique : function (date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = "select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree,    p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = 36137  AND almerys_lien_ss_spe.id_alm_ss_spe=926  AND almerys_lien_ss_spe.id_lien_ss_spe2=1194  AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
  selectionPecAudio : function (date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = "select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree,    p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = 36138  AND almerys_lien_ss_spe.id_alm_ss_spe=932  AND almerys_lien_ss_spe.id_lien_ss_spe2=1208  AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
  selectionPecHospi : function (date,callback) {
    //var nbr = parseInt(nb);
    console.log(date);
    var sql = "select DISTINCT SUM(DATE_PART('epoch', to_timestamp(p_ldt.date_fin_ldt||' '||p_ldt.h_fin, 'YYYYMMDD HH24:MI:SS') -  to_timestamp(p_ldt.date_deb_ldt||' '||p_ldt.h_deb, 'YYYYMMDD HH24:MI:SS') ))/3600 as duree,    p_ldt.id_pers from p_ldt LEFT join almerys_lien_ss_spe ON p_ldt.id_ldt= almerys_lien_ss_spe.id_ldt where p_ldt.id_lotclient = 36142  AND almerys_lien_ss_spe.id_alm_ss_spe=949  AND almerys_lien_ss_spe.id_lien_ss_spe2=1248  AND date_deb_ldt = '"+date+"' AND id_type_ldt=0 group by almerys_lien_ss_spe.id_almerys,p_ldt.id_pers order by p_ldt.id_pers";
    Tpstpc.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { 
        console.log(err);
        //return callback(err);
       }
      else
      {
        var somme = 0;
        console.log(sql);
        result = res.rows;
        /*console.log(result[0].duree);
        console.log(result.length);*/
        var f = 0.00;
        for(var i=0;i<result.length;i++)
        {
          var m = parseFloat(result[i].duree);
          f = m.toFixed(2);
          somme= somme + parseFloat(f);
        }
        console.log(somme);
        return callback(null,somme);
      };
      });
  },
		/***********************************************************/ 
    countOkKo : function (table,nb, callback) {
      var sql ="select sum(tt16h) as tt16h,sum(tt23h) as tt23h,sum(ttj2) as ttj2,sum(ttj5) as ttj5,sum(stock16h) as stock16h,sum(bonj) as bonj,sum(bonj1) as bonj1,sum(bonj2) as bonj2,sum(bonj5) as bonj5 from "+table[nb]+" ";
      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { 
          console.log(err);
          //return callback(err);
         }
        else
        {
          console.log(sql);
          result = res.rows;
          console.log(result[0].ttj2);
          return callback(null,result);
        };
        });
    },
    countErreur : function (table,nb, callback) {
      var sql ="SELECT erreureasy from tpserreur; ";
      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
        if (err) { 
          console.log(err);
          //return callback(err);
         }
        else
        {
          console.log(sql);
          result = res.rows;
          var resultat = result[0].erreureasy;
         // console.log(result[0] +'e');
          return callback(null,resultat);
        };
        });
    },
    ecriture : async function (tab,date_export,motcle,nb,callback) {
      const Excel = require('exceljs');
      const newWorkbook = new Excel.Workbook();
      try{
      //var path_reporting = 'D:/Reporting Engagement/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
      console.log(path_reporting);
      await newWorkbook.xlsx.readFile(path_reporting);
      const newworksheet = newWorkbook.getWorksheet('202106_Easy');
      var colonneDate = newworksheet.getColumn('A');
      var ligneDate1;
      //var date_export='14/06/2021';
      console.log(date_export);
      var ligne = 0;

      colonneDate.eachCell(function(cell, rowNumber) {
        var dateExcel = ReportingInovcomExport.convertDate(cell.text);
        if(dateExcel==date_export)
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
        return callback(null,'KO');
        //Reportinghtp.deleteToutHtp(tab,3,callback);
      }
      },


      ecritureEtp : async function (tab,date_export,motcle,nb,callback) {
        const Excel = require('exceljs');
        const newWorkbook = new Excel.Workbook();
        try{
        //var path_reporting = 'D:/Reporting Engagement/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
        await newWorkbook.xlsx.readFile(path_reporting);
        const newworksheet = newWorkbook.getWorksheet('202106_Easy');
        var colonneDate = newworksheet.getColumn('A');
        var ligneDate1;
        //var date_export='14/06/2021';
        console.log(date_export);
        var ligne = 0;
  
        colonneDate.eachCell(function(cell, rowNumber) {
          var dateExcel = ReportingInovcomExport.convertDate(cell.text);
          if(dateExcel==date_export)
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
        var valeur = parseFloat(tab)
        m.getCell(5).value = valeur;
       
        await newWorkbook.xlsx.writeFile(path_reporting);
        sails.log("Ecriture OK KO terminé"); 
        return callback(null, "OK");
      
        }
        catch
        {
          console.log("Une erreur s'est produite");
          return callback(null,'KO');
        }
        },

        ecritureEtp2 : async function (tab,date_export,motcle,nb,callback) {
          
            console.log('erreurrrrr');
            const Excel = require('exceljs');
            const newWorkbook = new Excel.Workbook();
            try{
            //var path_reporting = 'D:/Reporting Engagement/TPS-TPC_Reporting-Traitement-J-SLA_V12.xlsx';
            await newWorkbook.xlsx.readFile(path_reporting);
            const newworksheet = newWorkbook.getWorksheet('202106_Easy');
            var colonneDate = newworksheet.getColumn('A');
            var ligneDate1;
            //var date_export='14/06/2021';
            console.log(date_export);
            var ligne = 0;
      
            colonneDate.eachCell(function(cell, rowNumber) {
              var dateExcel = ReportingInovcomExport.convertDate(cell.text);
              if(dateExcel==date_export)
              {
                ligneDate1 = parseInt(rowNumber);
                var line = newworksheet.getRow(ligneDate1);
                var f = line.getCell(4).value;
                var bi = 'Erreur applicative';
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
            var valeur = parseFloat(tab);
            m.getCell(25).value = valeur;
           
            await newWorkbook.xlsx.writeFile(path_reporting);
            sails.log("Ecriture OK KO terminé"); 
            return callback(null, "OK");
          
            }
            catch
            {
              console.log("Une erreur s'est produite");
              return callback(null,'KO');
            }
        
         
          },
    /***********************************************************/  

  traitementInsertionEtp:function(table,callback){
    XLSX = require('xlsx');
    var trameflux= "D:/Reporting Engagement/TDB Reporting Almerys.xlsx";
    var workbook = XLSX.readFile(trameflux);
    
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      console.log('ok v');
      for(var ra=1;ra<=range.e.r;ra++)
        {
         // console.log(ra);
          var address_of_cell = {c:3, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          if(desired_value!=undefined)
          {
            var a = parseFloat(desired_value);
            somme = somme + a;
          }
          else{
           var a = 1;
          }
        };
        //console.log(somme);
        var sommefinal = parseInt(somme) / 7.5;
        var sql = "insert into "+table+" (erreureasy) values ("+sommefinal+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
  traitementInsertionErreur:function(table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      console.log('ok v');
      for(var ra=1;ra<=range.e.r;ra++)
        {
         // console.log(ra);
          var address_of_cell = {c:0, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
     

          var address_of_cell1 = {c:1, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          var address_of_cell2 = {c:2, r:ra};
          var cell_ref2 = XLSX.utils.encode_cell(address_of_cell2);
          var desired_cell2 = sheet[cell_ref2];
          var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);

          var a1 = 'SUPPORTN2-370';
          var a2 = 'ECH2';
          var a3 = 'SUPPORTN2-720';
          var a4= 'EA13';
          var a5 = 'ECH1';
          var a6 = 'ECH2ACS';
          var a7 = 'EA13ACS';
          var a8 = 'SUPPORTCLT-1178';
          var a9 = 'ECH1MGEFI';
          var a10 = 'ECH2MGEFI';
          var a11 = 'EA13MGEFI';
          var a12 = 'SUPPORTN2-448';
          var a13 = 'ECH1ACS';

          var b1 = 'FIN OK';
          var b2 = 'erreur conversion image';
          var b3 = 'factures a destination du client';
          var b4= 'Erreur interne de servlet';
          var b5 = 'Erreur au chargement de l';
          var b6 = 'ACS Erreur conversion image';
          var b7 = 'HTP acs erreur interne servlet';
          var b8 = 'Generali Erreur visualisation des documents';
          var b9 = 'MGEFI Erreur au chargement de l';
          var b10 = 'MGEFI Erreur conversion image';
          var b11 = 'HTP MGEFI Erreur interne servlet';
          var b12 = 'Titre de Recette MACIF 1500';
          var b13 = 'ACS Erreur au chargement de l';

          const regexa1 = new RegExp(a1,'i');
          const regexa2 = new RegExp(a2,'i');
          const regexa3 = new RegExp(a3,'i');
          const regexa4 = new RegExp(a4,'i');
          const regexa5 = new RegExp(a5,'i');
          const regexa6 = new RegExp(a6,'i');
          const regexa7 = new RegExp(a7,'i');
          const regexa8 = new RegExp(a8,'i');
          const regexa9 = new RegExp(a9,'i');
          const regexa10 = new RegExp(a10,'i');
          const regexa11 = new RegExp(a11,'i');
          const regexa12 = new RegExp(a12,'i');
          const regexa13 = new RegExp(a13,'i');

          const regexb1 = new RegExp(b1,'i');
          const regexb2 = new RegExp(b2,'i');
          const regexb3 = new RegExp(b3,'i');
          const regexb4 = new RegExp(b4,'i');
          const regexb5 = new RegExp(b5,'i');
          const regexb6 = new RegExp(b6,'i');
          const regexb7 = new RegExp(b7,'i');
          const regexb8 = new RegExp(b8,'i');
          const regexb9 = new RegExp(b9,'i');
          const regexb10 = new RegExp(b10,'i');
          const regexb11 = new RegExp(b11,'i');
          const regexb12 = new RegExp(b12,'i');
          const regexb13 = new RegExp(b13,'i');

          if( (regexa1.test(desired_value) && regexb1.test(desired_value1)) || (regexa2.test(desired_value) && regexb2.test(desired_value1)) || (regexa3.test(desired_value) && regexb3.test(desired_value1)) || (regexa4.test(desired_value) && regexb4.test(desired_value1)) || (regexa5.test(desired_value) && regexb5.test(desired_value1)) || (regexa6.test(desired_value) && regexb6.test(desired_value1)) || (regexa7.test(desired_value) && regexb7.test(desired_value1)) || (regexa8.test(desired_value) && regexb8.test(desired_value1)) || (regexa9.test(desired_value) && regexb9.test(desired_value1)) || (regexa10.test(desired_value) && regexb10.test(desired_value1)) || (regexa11.test(desired_value) && regexb11.test(desired_value1) || (regexa12.test(desired_value) && regexb12.test(desired_value1)) || (regexa13.test(desired_value) && regexb13.test(desired_value1))  )  )
          {
           var am =1;
          }
          else{
           somme = somme + parseInt(desired_value2);
          }
        };
        console.log(somme);
        var sql = "insert into "+table+" (erreureasy) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
  traitementInsertionstocksanteclairJ5:function(nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      if(jour=='Friday')
      {
        console.log('Fridayday');
      for(var ra=2;ra<=range.e.r;ra++)
        {
         // console.log(ra);
          var address_of_cell = {c:7, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          /*var bi = 'IDENTIFIANT DE LA TACHE';
          const regex = new RegExp(bi,'i');*/

          var address_of_cell1 = {c:5, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          var conv = parseInt(date);
          var j2 = conv - 5;
          if(desired_value!=undefined && parseInt(desired_value1)<=j2)
          {
           somme = somme +1;  
          }
          else{
           var a = 'l';
          }
        };
        console.log(somme);
        var sql = "insert into "+table[nb]+" (bonj5) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
        console.log('No Monday');
      for(var ra=2;ra<=range.e.r;ra++)
        {
         // console.log(ra);
          var address_of_cell = {c:7, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          /*var bi = 'IDENTIFIANT DE LA TACHE';
          const regex = new RegExp(bi,'i');*/

          var address_of_cell1 = {c:5, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          var conv = parseInt(date);
          var j2 = conv - 5;
          if(desired_value!=undefined && parseInt(desired_value1)<=j2)
          {
           somme = somme +1;  
          }
          else{
           var a = 'l';
          }
        };
        console.log(somme);
        var sql = "insert into "+table[nb]+" (bonj5) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
  traitementInsertionstocksanteclairJ2:function(nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      if(jour=='Monday' || jour=='Tuesday')
      {
        console.log('Monday');
      for(var ra=2;ra<=range.e.r;ra++)
        {
         // console.log(ra);
          var address_of_cell = {c:7, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          /*var bi = 'IDENTIFIANT DE LA TACHE';
          const regex = new RegExp(bi,'i');*/

          var address_of_cell1 = {c:5, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          var conv = parseInt(date);
          var j2 = conv - 4;
          if(desired_value!=undefined && parseInt(desired_value1)<=j2)
          {
           somme = somme +1;  
          }
          else{
           var a = 'l';
          }
        };
        console.log(somme);
        var sql = "insert into "+table[nb]+" (bonj2) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
        console.log('No Monday');
      for(var ra=2;ra<=range.e.r;ra++)
        {
         // console.log(ra);
          var address_of_cell = {c:7, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          /*var bi = 'IDENTIFIANT DE LA TACHE';
          const regex = new RegExp(bi,'i');*/

          var address_of_cell1 = {c:5, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          var conv = parseInt(date);
          var j2 = conv - 2;
          if(desired_value!=undefined && parseInt(desired_value1)<=j2)
          {
           somme = somme +1;  
          }
          else{
           var a = 'l';
          }
        };
        console.log(somme);
        var sql = "insert into "+table[nb]+" (bonj2) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
  traitementInsertionstocksanteclairJ1:function(nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      if(jour=='Monday')
      {
        console.log('Monday');
      for(var ra=2;ra<=range.e.r;ra++)
        {
         // console.log(ra);
          var address_of_cell = {c:7, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          /*var bi = 'IDENTIFIANT DE LA TACHE';
          const regex = new RegExp(bi,'i');*/

          var address_of_cell1 = {c:5, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          var conv = parseInt(date);
              var j1 = conv - 1;
              var j2 = conv -2;
              var j3 = conv - 3;
          if(desired_value!=undefined && (desired_value1==j1 || desired_value1==j2 || desired_value1==j3 ))
          {
           somme = somme +1;  
          }
          else{
           var a = 'l';
          }
        };
        console.log(somme);
        var sql = "insert into "+table[nb]+" (bonj1) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
        console.log('No Monday');
      for(var ra=2;ra<=range.e.r;ra++)
        {
         // console.log(ra);
          var address_of_cell = {c:7, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          /*var bi = 'IDENTIFIANT DE LA TACHE';
          const regex = new RegExp(bi,'i');*/

          var address_of_cell1 = {c:5, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          var conv = parseInt(date);
          var j1 = conv - 1;
          if(desired_value!=undefined && desired_value1==j1)
          {
           somme = somme +1;  
          }
          else{
           var a = 'l';
          }
        };
        console.log(somme);
        var sql = "insert into "+table[nb]+" (bonj1) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
  traitementInsertionstocksanteclairJ:function(nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      for(var ra=2;ra<=range.e.r;ra++)
        {
         // console.log(ra);
          var address_of_cell = {c:7, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          /*var bi = 'IDENTIFIANT DE LA TACHE';
          const regex = new RegExp(bi,'i');*/

          var address_of_cell1 = {c:5, r:ra};
          var cell_ref1 = XLSX.utils.encode_cell(address_of_cell1);
          var desired_cell1 = sheet[cell_ref1];
          var desired_value1 = (desired_cell1 ? desired_cell1.v : undefined);

          var j1 = date;
          if(desired_value!=undefined && desired_value1==j1)
          {
           somme = somme +1;  
          }
          else{
           var a = 'l';
          }
        };
        console.log(somme);
        var sql = "insert into "+table[nb]+" (bonj) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
  traitementInsertionstocksanteclair:function(nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      for(var ra=2;ra<=range.e.r;ra++)
        {
         // console.log(ra);
          var address_of_cell = {c:7, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);
          /*var bi = 'IDENTIFIANT DE LA TACHE';
          const regex = new RegExp(bi,'i');*/
          if(desired_value!=undefined)
          {
           somme = somme +1;  
          }
          else{
           var a = 'l';
          }
        };
        console.log(somme);
        var sql = "insert into "+table[nb]+" (stock16h) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
      /*if(jour=='Friday')
      {*/
        console.log('Friday');
        for(var ra=0;ra<=range.e.r;ra++)
        {
          //nature de tache
          var address_of_cell = {c:2, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:7, r:ra};
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
          var address_of_cell4 = {c:10, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

           //date de maturit5
           var address_of_cell5 = {c:5, r:ra};
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
                                          //}
         /* else
          {
              console.log('hafa');
              for(var ra=0;ra<=range.e.r;ra++)
              {
                //nature de tache
                var address_of_cell = {c:2, r:ra};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value = (desired_cell ? desired_cell.v : undefined);

                // identification de la tache : le isaina
                var address_of_cell1 = {c:7, r:ra};
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
                var address_of_cell4 = {c:10, r:ra};
                var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
                var desired_cell4 = sheet[cell_ref4];
                var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

                //date de maturit5
                var address_of_cell5 = {c:5, r:ra};
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
                                                }*/
          
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
          var address_of_cell = {c:2, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:7, r:ra};
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
          var address_of_cell4 = {c:10, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

           //date de maturit5
           var address_of_cell5 = {c:5, r:ra};
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
                var address_of_cell = {c:2, r:ra};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value = (desired_cell ? desired_cell.v : undefined);

                // identification de la tache : le isaina
                var address_of_cell1 = {c:7, r:ra};
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
                var address_of_cell4 = {c:10, r:ra};
                var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
                var desired_cell4 = sheet[cell_ref4];
                var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

                //date de maturit5
                var address_of_cell5 = {c:5, r:ra};
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
          var address_of_cell = {c:2, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:7, r:ra};
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
          var address_of_cell4 = {c:10, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

          //date de maturit5
          var address_of_cell5 = {c:5, r:ra};
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
          var address_of_cell = {c:2, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:7, r:ra};
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
          var address_of_cell4 = {c:10, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

          //date de maturit5
          var address_of_cell5 = {c:5, r:ra};
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
          var address_of_cell = {c:2, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:7, r:ra};
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
          var address_of_cell4 = {c:10, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

           //date de maturit5
           var address_of_cell5 = {c:5, r:ra};
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
        var address_of_cell = {c:2, r:ra};
        var cell_ref = XLSX.utils.encode_cell(address_of_cell);
        var desired_cell = sheet[cell_ref];
        var desired_value = (desired_cell ? desired_cell.v : undefined);

        // identification de la tache : le isaina
        var address_of_cell1 = {c:7, r:ra};
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
        var address_of_cell4 = {c:10, r:ra};
        var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
        var desired_cell4 = sheet[cell_ref4];
        var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

        //date de maturit5
        var address_of_cell5 = {c:5, r:ra};
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
                var address_of_cell = {c:2, r:ra};
                var cell_ref = XLSX.utils.encode_cell(address_of_cell);
                var desired_cell = sheet[cell_ref];
                var desired_value = (desired_cell ? desired_cell.v : undefined);

                // identification de la tache : le isaina
                var address_of_cell1 = {c:7, r:ra};
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
                var address_of_cell4 = {c:10, r:ra};
                var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
                var desired_cell4 = sheet[cell_ref4];
                var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

                //date de maturit5
                var address_of_cell5 = {c:5, r:ra};
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
          var address_of_cell = {c:2, r:ra};
          var cell_ref = XLSX.utils.encode_cell(address_of_cell);
          var desired_cell = sheet[cell_ref];
          var desired_value = (desired_cell ? desired_cell.v : undefined);

          // identification de la tache : le isaina
          var address_of_cell1 = {c:7, r:ra};
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
          var address_of_cell4 = {c:10, r:ra};
          var cell_ref4 = XLSX.utils.encode_cell(address_of_cell4);
          var desired_cell4 = sheet[cell_ref4];
          var desired_value4 = (desired_cell4 ? desired_cell4.v : undefined);

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
            var sql = "insert into "+table[nb]+" (stock16h ) values ("+somme+") ";
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
  traitementInsertion23h:function(ast,traitement,motcle1,motcle2,motcle3,motcle4,nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    try{
      /*console.log(ast[nb]);
      console.log(traitement[nb]);
      console.log(motcle3[nb]);
      console.log(motcle1[nb]);
      console.log(motcle2[nb]);*/
      //console.log(ast[nb]);
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

          var b1 = ast[nb];
          
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
        var sql = "insert into "+table[nb]+" (tt23h) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
      /*console.log(ast[nb]);
      console.log(traitement[nb]);
      console.log(motcle3[nb]);
      console.log(motcle1[nb]);
      console.log(motcle2[nb]);*/
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

          var b1 = ast[nb];
          
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
  traitementInsertionJ2:function(ast,traitement,motcle1,motcle2,motcle3,motcle4,nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
    var trameflux= chemin;
    var workbook = XLSX.readFile(trameflux);
    try{
      const sheet = workbook.Sheets[workbook.SheetNames[1]];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var somme = 0;
      if(jour=='Monday' || jour=='Tuesday' || jour=='Wednesday')
      {
        console.log('Monday ou Tuesday ou Wednes');
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

          var address_of_cell5 = {c:6, r:ra};
          var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
          var desired_cell5 = sheet[cell_ref5];
          var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);
          
          var b2 = "ALMERYS";
          var b3= "CBTP";
        
          var b4 = traitement[nb];
          const regex2 = new RegExp(b2,'i');
          const regex3 = new RegExp(b3,'i');
          const regex4 = new RegExp(b4,'i');

          var b1 = ast[nb];

          var j1 = date;
          var conv = parseInt(date);
            
            var j2 = conv - 1;
            var j3 = conv - 2;
            var j4 = conv - 3;
            var j5 = conv - 4;
            //console.log(j1+j2+j3+j4+j5);
            const regex1 = new RegExp(b1);
            if(regex1.test(desired_value1) &&  ( regex2.test(desired_value2) || regex3.test(desired_value2) )  &&  ( desired_value5==j1 ||  desired_value5==j2 || desired_value5==j3 ||  desired_value5==j4 || desired_value5==j5))
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
          }
        
        var sql = "insert into "+table[nb]+" (ttj2) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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
        console.log('tsy Monday');
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

          var address_of_cell5 = {c:6, r:ra};
          var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
          var desired_cell5 = sheet[cell_ref5];
          var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);
          
          var b2 = "ALMERYS";
          var b3= "CBTP";
        
          var b4 = traitement[nb];
          const regex2 = new RegExp(b2,'i');
          const regex3 = new RegExp(b3,'i');
          const regex4 = new RegExp(b4,'i');

          var b1 = ast[nb];

          var j1 = date;
          var conv = parseInt(date);
            
            var j2 = conv - 1;
            var j3 = conv -2;
           
            const regex1 = new RegExp(b1);
            if(regex1.test(desired_value1) &&  ( regex2.test(desired_value2) || regex3.test(desired_value2) )  &&  ( desired_value5==j1 ||  desired_value5==j2 || desired_value5==j3  ))
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
          }
        
        var sql = "insert into "+table[nb]+" (ttj2) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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

  traitementInsertionJ5:function(ast,traitement,motcle1,motcle2,motcle3,motcle4,nb,jour,date,table,chemin,callback){
    XLSX = require('xlsx');
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

          var address_of_cell5 = {c:6, r:ra};
          var cell_ref5 = XLSX.utils.encode_cell(address_of_cell5);
          var desired_cell5 = sheet[cell_ref5];
          var desired_value5 = (desired_cell5 ? desired_cell5.v : undefined);
          
          var b2 = "ALMERYS";
          var b3= "CBTP";
        
          var b4 = traitement[nb];
          const regex2 = new RegExp(b2,'i');
          const regex3 = new RegExp(b3,'i');
          const regex4 = new RegExp(b4,'i');

          var b1 = ast[nb];

          var j1 = date;
          var conv = parseInt(date);
            
            var j2 = conv - 1;
            var j3 = conv -2;
            var j4 = conv -3;
            var j5 = conv -4;
            var j6 = conv - 5;
            var j7 = conv -6;
            var j8 = conv - 7;
            const regex1 = new RegExp(b1);
            if(regex1.test(desired_value1) &&  ( regex2.test(desired_value2) || regex3.test(desired_value2) )  &&  ( desired_value5==j1 ||  desired_value5==j2 || desired_value5==j3 ||  desired_value5==j4 || desired_value5==j5 || desired_value5==j6 || desired_value5==j7 || desired_value5==j8  ))
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
                  el
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
          }
        
        var sql = "insert into "+table[nb]+" (ttj5) values ("+somme+") ";
                      Reportinghtp.getDatastore().sendNativeQuery(sql, function(err,res){
                        if (err) { 
                          console.log("Une erreur ve?");
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

