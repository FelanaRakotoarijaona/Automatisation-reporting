/**
 * ReportingInovcomExportController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */
module.exports = {
    accueil : function(req,res)
    {
      var html= req.param("html");
      if(html=='o')
      {
        html= '<script>'+
        '$( document ).ready(function() {'+
        'tost();'+
        '});'+
        "function tost(){$('#toastk').toast('show');}</script>";
      };
      if(html=='x')
      {
        html= '<script>'+
        '$( document ).ready(function() {'+
        'tost();'+
        '});'+
        "function tost(){$('#toastd').toast('show');}</script>";
      };
      var jour = req.param("jour");
      var mois = req.param("mois");
      var annee = req.param("annee");
      var dateexport = jour + '/' + mois + '/' +annee;
      return res.view('Inovcom/exportErica', {date : dateexport , html : html});
    },
    rechercheColonne : function (req, res) {
      var jour = req.param("jour");
      var mois = req.param("mois");
      var annee = req.param("annee");
      var mois1 = 'Janvier' ;
      if(mois==01)
      {
        mois1= 'Janvier';
      };
      if(mois==02)
      {
        mois1= 'Fevrier';
      };
      if(mois==03)
      {
        mois1= 'Mars';
      };
      if(mois==04)
      {
        mois1= 'Avril';
      };
      if(mois==05)
      {
        mois1= 'Mai';
      };
      console.log(mois1);
      var date_export = jour + '/' + mois + '/' +annee;
      console.log("RECHERCHE COLONNE");
      async.series([
        function (callback) {
          ReportingInovcomExport.countOkKo("extractionrcforce",callback);
        },
        function (callback) {
          ReportingInovcomExport.countOkKo("favmgefi",callback);
        },
       /* function (callback) {
          ReportingInovcomExport.countOkKo("suivisaisiemgas",callback);
        },
        function (callback) {
          ReportingInovcomExport.countOkKo("suivisaisieprodite",callback);
        },
        function (callback) {
          ReportingInovcomExport.countOkKoTrameLamie("tramelamiestock",callback);
        },
        function (callback) {
          ReportingInovcomExport.countOkKoTrameLamieResiliation("tramelamiestock",callback);
        }*/
      ],function(err,result){
        if(err) return res.badRequest(err);
        console.log("Count OK extractionrcforce ==> " + result[0].ok + " / " + result[0].ko);
        console.log("Count OK favmgefi ==> " + result[1].ok + " / " + result[1].ko);
        async.series([
          function (callback) {
            ReportingInovcomExport.ecritureOkKo(result[0],"extractionrcforce",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo(result[1],"favmgefi",date_export,mois1,callback);
          },
         /* function (callback) {
            ReportingInovcomExport.ecritureOkKo(result[2],"suivisaisiemgas",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo(result[3],"suivisaisieprodite",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo(result[4],"tramelamiestocknr",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo(result[5],"tramelamiestock",date_export,mois1,callback);
          },*/

        ],function(err,resultExcel){
       
            if(resultExcel[0]==true)
            {
              console.log("true zn");
              res.view('Inovcom/erera');
            }
            if(resultExcel[0]=='OK')
            {
              res.redirect('/exportInovcom/'+date_export+'/x')
            }
        })
      })
    },
};

