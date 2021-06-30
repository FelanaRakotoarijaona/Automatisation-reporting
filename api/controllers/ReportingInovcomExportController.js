/**
 * ReportingInovcomExportController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */
module.exports = {
    accueilInov1 : function(req,res)
    {
      return res.view('Inovcom/exportexcelinovcom1');
    },
    accueilInov2 : function(req,res)
    {
      return res.view('Inovcom/exportexcelinovcom2');
    },
    accueilInov3 : function(req,res)
    {
      return res.view('Inovcom/exportexcelinovcom3');
    },
    accueilInov4 : function(req,res)
    {
      return res.view('Inovcom/exportexcelinovcom4');
    },
    accueilInov5 : function(req,res)
    {
      return res.view('Inovcom/exportexcelinovcom5');
    },
    accueilInov6 : function(req,res)
    {
      return res.view('Inovcom/exportexcelinovcom6');
    },
    accueilInov7 : function(req,res)
    {
      return res.view('Inovcom/exportexcelinovcom7');
    },
    accueilInov8 : function(req,res)
    {
      return res.view('Inovcom/exportexcelinovcom8');
    },
    accueilInov9 : function(req,res)
    {
      return res.view('Inovcom/exportexcelinovcom9');
    },
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
    rechercheColonne1: function (req, res) {
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
      // var jour = req.param("jour");
      // var mois = req.param("mois");
      // var annee = req.param("annee");
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
      if(mois==06)
      {
        mois1= 'Juin';
      };
      if(mois==07)
      {
        mois1= 'Juillet';
      };
      if(mois==08)
      {
        mois1= 'Aout';
      };
      if(mois==09)
      {
        mois1= 'Septembre';
      };
      if(mois==10)
      {
        mois1= 'Octobre';
      };
      if(mois==11)
      {
        mois1= 'Novembre';
      };
      if(mois==12)
      {
        mois1= 'Decembre';
     };
      console.log(mois1);
      var date_export = jour + '/' + mois + '/' +annee;
      console.log("RECHERCHE COLONNE");
      async.series([
        // function (callback) {
        //   ReportingInovcomExport.countOkKo("extractionrcforce",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countOkKo("favmgefi",callback);
        // },
        function (callback) {
          ReportingInovcomExport.countOkKo("retourconventionsaisiedesconventions",callback);
        },
        function (callback) {
          ReportingInovcomExport.countOkKofll1("ribtpmep",callback);
        },
        function (callback) {
          ReportingInovcomExport.countOkKofll11("ribtpmep",callback);
        },
        function (callback) {
          ReportingInovcomExport.countOkKofll1("curethermale",callback);
        },
        function (callback) {
          ReportingInovcomExport.countOkKo11("retourconventionsaisiedesconventions",callback);
        },
 
      ],function(err,result){
        if(err) return res.badRequest(err);
        console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
        console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
        console.log("Count OK 2 ==> " + result[2].ok + " / " + result[2].ko);
        console.log("Count OK 4==> " + result[4].ok + " / " + result[4].ko);
        async.series([
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo(result[0],"extractionrcforce",date_export,mois1,callback);
          // },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo(result[1],"favmgefi",date_export,mois1,callback);
          // },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo1(result[0],"retourconventionsaisiedesconventions",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo(result[1],"ribtpmep",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo(result[2],"tpmep",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo(result[3],"curethermale",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo11(result[4],"conventions",date_export,mois1,callback);
          },
        ],function(err,resultExcel){
       
            if(resultExcel[0]==true)
            {
              console.log("true zn");
              res.view('Inovcom/erera');
            }
            if(resultExcel[0]=='OK')
            {
              // res.redirect('/exportInovcom/'+date_export+'/x')
              res.view('Contentieux/succes');
            }
        })
      })
    },
    /******************************************************************************/
    rechercheColonne2: function (req, res) {
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
    
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
      if(mois==06)
      {
        mois1= 'Juin';
      };
      if(mois==07)
      {
        mois1= 'Juillet';
      };
      if(mois==08)
      {
        mois1= 'Aout';
      };
      if(mois==09)
      {
        mois1= 'Septembre';
      };
      if(mois==10)
      {
        mois1= 'Octobre';
      };
      if(mois==11)
      {
        mois1= 'Novembre';
      };
      if(mois==12)
      {
        mois1= 'Decembre';
      };
      console.log(mois1);
      var date_export = jour + '/' + mois + '/' +annee;
      console.log("RECHERCHE COLONNE");
      async.series([
       
        function (callback) {
          ReportingInovcomExport.countok("dentaireretourfacturedentaireetcds",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("optiqueretourpublipostage",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("factureaudio",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("retourhospipec",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("retourpecdentaire",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("retourpecoptique",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("retourpecaudio",callback);
        },
        // function (callback) {
        //   ReportingInovcomExport.countok("santeclairtableauretourgeneral",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countok("santeclairoptique",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countok("noemiehtpmgefi",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countok("mgefigtomgefirejetsaisienoemiehtp",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countok("retourreclamtramereclamationtiers",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countok("reclamsetramereclamationse",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countok("reclamhospi",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countok("dentairereclamationdentaire",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countok("optiquetramereclamationoptique",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countok("reclamationaudio",callback);
        // },

      ],function(err,result){
        if(err) return res.badRequest(err);
        
        else{
          console.log("Count OK suivant 0 ==> " + result[0].ok + " / " + result[0].ko);
        console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
        console.log("Count OK 2 ==> " + result[2].ok + " / " + result[2].ko);
        console.log("Count OK 3 ==> " + result[3].ok + " / " + result[3].ko);
        console.log("Count OK 4 ==> " + result[4].ok + " / " + result[4].ko);
        console.log("Count OK 5 ==> " + result[5].ok + " / " + result[5].ko);
        console.log("Count OK 6 ==> " + result[6].ok + " / " + result[6].ko);
       
        async.series([
         
         function (callback) {
            ReportingInovcomExport.ecritureOkKo2(result[0],"dentaireretourfacturedentaireetcds",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo2(result[1],"optiqueretourpublipostage",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo2(result[2],"factureaudio",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo2(result[3],"retourhospipec",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo2(result[4],"retourpecdentaire",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo2(result[5],"retourpecoptique",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo2(result[6],"retourpecaudio",date_export,mois1,callback);
          },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo21(result[7],"santeclairtableauretourgeneral",date_export,mois1,callback);
          // },  
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo21(result[8],"santeclairoptique",date_export,mois1,callback);
          // },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo22(result[7],"noemiehtpmgefi",date_export,mois1,callback);
          // },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo22(result[8],"mgefigtomgefirejetsaisienoemiehtp",date_export,mois1,callback);
          // },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo23(result[7],"retourreclamtramereclamationtiers",date_export,mois1,callback);
          // },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo23(result[8],"reclamsetramereclamationse",date_export,mois1,callback);
          // },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo23(result[9],"reclamhospi",date_export,mois1,callback);
          // },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo23(result[10],"dentairereclamationdentaire",date_export,mois1,callback);
          // },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo23(result[11],"optiquetramereclamationoptique",date_export,mois1,callback);
          // },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo23(result[12],"reclamationaudio",date_export,mois1,callback);
          // },



          ],function(err,resultExcel){
            console.log('**************');
            console.log(resultExcel);
            console.log('**************');
            if(resultExcel[0]==true)
            {
              console.log("true zn");
              res.view('Inovcom/erera');
            }
            else
            {
              // return res.view('Inovcom/exportsuivantinovcom2', {date: datetest});
              res.view('reporting/succes');

            }
        });//fermeture async 2

      };

    });//fermeture async 1
    
    },
    /*********************************************************************************/
    rechercheColonne2suivant: function (req, res) {
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
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
      if(mois==06)
      {
        mois1= 'Juin';
      };
      if(mois==07)
      {
        mois1= 'Juillet';
      };
      if(mois==08)
      {
        mois1= 'Aout';
      };
      if(mois==09)
      {
        mois1= 'Septembre';
      };
      if(mois==10)
      {
        mois1= 'Octobre';
      };
      if(mois==11)
      {
        mois1= 'Novembre';
      };
      if(mois==12)
      {
        mois1= 'Decembre';
      };
      console.log(mois1);
      var date_export = jour + '/' + mois + '/' +annee;
      console.log("RECHERCHE COLONNE");
      async.series([
        function (callback) {
          ReportingInovcomExport.countok("santeclairoptique",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("noemiehtpmgefi",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("mgefigtomgefirejetsaisienoemiehtp",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("retourreclamtramereclamationtiers",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("reclamsetramereclamationse",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("reclamhospi",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("dentairereclamationdentaire",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("optiquetramereclamationoptique",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("reclamationaudio",callback);
        },
      ],function(err,result){
        if(err) return res.badRequest(err);
        console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
        console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
        console.log("Count OK 2 ==> " + result[2].ok + " / " + result[2].ko);
        console.log("Count OK 3 ==> " + result[3].ok + " / " + result[3].ko);
        console.log("Count OK 4 ==> " + result[4].ok + " / " + result[4].ko);
        console.log("Count OK 5 ==> " + result[5].ok + " / " + result[5].ko);
        console.log("Count OK 6 ==> " + result[6].ok + " / " + result[6].ko);
        async.series([
          function (callback) {
            ReportingInovcomExport.ecritureOkKo21(result[0],"santeclairoptique",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo22(result[1],"noemiehtpmgefi",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo22(result[2],"mgefigtomgefirejetsaisienoemiehtp",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[3],"retourreclamtramereclamationtiers",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[4],"reclamsetramereclamationse",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[5],"reclamhospi",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[6],"dentairereclamationdentaire",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[7],"optiquetramereclamationoptique",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[8],"reclamationaudio",date_export,mois1,callback);
          },

        ],function(err,resultExcel){
            console.log('**************');
            console.log(resultExcel);
            console.log('**************');
            if(resultExcel[0]==true)
            {
              console.log("true zn");
              res.view('Inovcom/erera');
            }
            else
            {
              
              res.view('reporting/succes');
            }
        });
      });
    },
    /*********************************************************************************/
    rechercheColonne3: function (req, res) {
      var datetest = req.param("date",0);
      var annee = datetest.substr(0, 4);
      var mois = datetest.substr(5, 2);
      var jour = datetest.substr(8, 2);
      // var jour = req.param("jour");
      // var mois = req.param("mois");
      // var annee = req.param("annee");
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
      if(mois==06)
      {
        mois1= 'Juin';
      };
      if(mois==07)
      {
        mois1= 'Juillet';
      };
      if(mois==08)
      {
        mois1= 'Aout';
      };
      if(mois==09)
      {
        mois1= 'Septembre';
      };
      if(mois==10)
      {
        mois1= 'Octobre';
      };
      if(mois==11)
      {
        mois1= 'Novembre';
      };
      if(mois==12)
      {
        mois1= 'Decembre';
    };
      console.log(mois1);
      var date_export = jour + '/' + mois + '/' +annee;
      console.log("RECHERCHE COLONNE");
      async.series([
        function (callback) {
          ReportingInovcomExport.countok("majribcbtp",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("majagapsinteramc",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("hospidemat",callback);
        },
      ],function(err,result){
        if(err) return res.badRequest(err);
        console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
        console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
        console.log("Count OK 2 ==> " + result[2].ok + " / " + result[2].ko);
        async.series([          
          function (callback) {
            ReportingInovcomExport.ecritureOkKo3(result[0],"majribcbtp",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo3(result[1],"majagapsinteramc",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo31(result[2],"hospidemat",date_export,mois1,callback);
          },

        ],function(err,resultExcel){
       
            if(resultExcel[0]==true)
            {
              console.log("true zn");
              res.view('Inovcom/erera');
            }
            if(resultExcel[0]=='OK')
            {
              // res.redirect('/exportInovcom/'+date_export+'/x')
              res.view('reporting/succes');
            }
        })
      })
    },
/***********************************************************************/
rechercheColonne4: function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  // var jour = req.param("jour");
  // var mois = req.param("mois");
  // var annee = req.param("annee");
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
      if(mois==06)
      {
        mois1= 'Juin';
      };
      if(mois==07)
      {
        mois1= 'Juillet';
      };
      if(mois==08)
      {
        mois1= 'Aout';
      };
      if(mois==09)
      {
        mois1= 'Septembre';
      };
      if(mois==10)
      {
        mois1= 'Octobre';
      };
      if(mois==11)
      {
        mois1= 'Novembre';
      };
      if(mois==12)
      {
        mois1= 'Decembre';
    };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      ReportingInovcomExport.countOkKofll4("extractionrcforce",callback);
    },
    function (callback) {
      ReportingInovcomExport.countOkKofll4("faveole",callback);
    },
    function (callback) {
      ReportingInovcomExport.countOkKofll4("favmgefi",callback);
    },
    function (callback) {
      ReportingInovcomExport.countOkKofll4("favbalma",callback);
    },
    function (callback) {
      ReportingInovcomExport.countOkKofll4("rcindeterminable",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
    console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
    console.log("Count OK 2 ==> " + result[2].ok + " / " + result[2].ko);
    async.series([          
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4(result[0],"extractionrcforce",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4(result[1],"faveole",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4(result[2],"favmgefi",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4(result[3],"favbalma",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4(result[4],"rcindeterminable",date_export,mois1,callback);
      },
      // function (callback) {
      //   ReportingInovcomExport.ecritureOkKo4(result[0].ko,"extractionrcforce",date_export,mois1,callback);
      // // },
      // function (callback) {
      //   ReportingInovcomExport.ecritureOkKo4(result[1].ko,"faveole",date_export,mois1,callback);
      // },
      // function (callback) {
      //   ReportingInovcomExport.ecritureOkKo4(result[2].ko,"favmgefi",date_export,mois1,callback);
      // },
      // function (callback) {
      //   ReportingInovcomExport.ecritureOkKo4(result[3].ko,"favbalma",date_export,mois1,callback);
      // },
      // function (callback) {
      //   ReportingInovcomExport.ecritureOkKo4(result[4].ko,"rcindeterminable",date_export,mois1,callback);
      // },

    ],function(err,resultExcel){
   
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Inovcom/erera');
        }
        if(resultExcel[0]=='OK')
        {
          // res.redirect('/exportInovcom/'+date_export+'/x')
          res.view('reporting/succes');
        }
    })
  })
},
/**********************************************************************/
rechercheColonne5: function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  // var jour = req.param("jour");
  // var mois = req.param("mois");
  // var annee = req.param("annee");
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
      if(mois==06)
      {
        mois1= 'Juin';
      };
      if(mois==07)
      {
        mois1= 'Juillet';
      };
      if(mois==08)
      {
        mois1= 'Aout';
      };
      if(mois==09)
      {
        mois1= 'Septembre';
      };
      if(mois==10)
      {
        mois1= 'Octobre';
      };
      if(mois==11)
      {
        mois1= 'Novembre';
      };
      if(mois==12)
      {
        mois1= 'Decembre';
    };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      ReportingInovcomExport.countOkKo("fav",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
    async.series([          
      function (callback) {
        ReportingInovcomExport.ecritureOkKo5(result[0],"fav",date_export,mois1,callback);
      },
    ],function(err,resultExcel){
   
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Inovcom/erera');
        }
        if(resultExcel[0]=='OK')
        {
          // res.redirect('/exportInovcom/'+date_export+'/x')
          res.view('reporting/succes');
        }
    })
  })
},
/**********************************************************************/
rechercheColonne6: function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  // var jour = req.param("jour");
  // var mois = req.param("mois");
  // var annee = req.param("annee");
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
      if(mois==06)
      {
        mois1= 'Juin';
      };
      if(mois==07)
      {
        mois1= 'Juillet';
      };
      if(mois==08)
      {
        mois1= 'Aout';
      };
      if(mois==09)
      {
        mois1= 'Septembre';
      };
      if(mois==10)
      {
        mois1= 'Octobre';
      };
      if(mois==11)
      {
        mois1= 'Novembre';
      };
      if(mois==12)
      {
        mois1= 'Decembre';
    };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      ReportingInovcomExport.countOkKo("retourcmuc",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
    async.series([          
      function (callback) {
        ReportingInovcomExport.ecritureOkKo6(result[0],"retourcmuc",date_export,mois1,callback);
      },
    ],function(err,resultExcel){
   
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Inovcom/erera');
        }
        if(resultExcel[0]=='OK')
        {
          // res.redirect('/exportInovcom/'+date_export+'/x')
          res.view('reporting/succes');
        }
    })
  })
},
/**********************************************************************/
rechercheColonne7: function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  // var jour = req.param("jour");
  // var mois = req.param("mois");
  // var annee = req.param("annee");
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
      if(mois==06)
      {
        mois1= 'Juin';
      };
      if(mois==07)
      {
        mois1= 'Juillet';
      };
      if(mois==08)
      {
        mois1= 'Aout';
      };
      if(mois==09)
      {
        mois1= 'Septembre';
      };
      if(mois==10)
      {
        mois1= 'Octobre';
      };
      if(mois==11)
      {
        mois1= 'Novembre';
      };
      if(mois==12)
      {
        mois1= 'Decembre';
    };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      ReportingInovcomExport.countOkKofll7("hospidematrejetprive",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
    async.series([          
      function (callback) {
        ReportingInovcomExport.ecritureOkKo7(result[0],"hospidematrejetprive",date_export,mois1,callback);
      },
    ],function(err,resultExcel){
   
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Inovcom/erera');
        }
        if(resultExcel[0]=='OK')
        {
          // res.redirect('/exportInovcom/'+date_export+'/x')
          res.view('reporting/succes');
        }
    })
  })
},
/**********************************************************************/
rechercheColonne8: function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  // var jour = req.param("jour");
  // var mois = req.param("mois");
  // var annee = req.param("annee");
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
      if(mois==06)
      {
        mois1= 'Juin';
      };
      if(mois==07)
      {
        mois1= 'Juillet';
      };
      if(mois==08)
      {
        mois1= 'Aout';
      };
      if(mois==09)
      {
        mois1= 'Septembre';
      };
      if(mois==10)
      {
        mois1= 'Octobre';
      };
      if(mois==11)
      {
        mois1= 'Novembre';
      };
      if(mois==12)
      {
        mois1= 'Decembre';
    };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      ReportingInovcomExport.countOkKofll8("retouravisannulationtramealmerys",callback);
    },
    function (callback) {
      ReportingInovcomExport.countOkKofll8("retouravisannulationcbtp",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
    console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
    async.series([          
      function (callback) {
        ReportingInovcomExport.ecritureOkKo8(result[0],"retouravisannulationtramealmerys",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo81(result[1],"retouravisannulationcbtp",date_export,mois1,callback);
      },
    ],function(err,resultExcel){
   
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Inovcom/erera');
        }
        if(resultExcel[0]=='OK')
        {
          // res.redirect('/exportInovcom/'+date_export+'/x')
          res.view('reporting/succes');
        }
    })
  })
},
/**********************************************************************/
rechercheColonne9: function (req, res) {
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  // var jour = req.param("jour");
  // var mois = req.param("mois");
  // var annee = req.param("annee");
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
      if(mois==06)
      {
        mois1= 'Juin';
      };
      if(mois==07)
      {
        mois1= 'Juillet';
      };
      if(mois==08)
      {
        mois1= 'Aout';
      };
      if(mois==09)
      {
        mois1= 'Septembre';
      };
      if(mois==10)
      {
        mois1= 'Octobre';
      };
      if(mois==11)
      {
        mois1= 'Novembre';
      };
      if(mois==12)
      {
        mois1= 'Decembre';
    };
  console.log(mois1);
  var date_export = jour + '/' + mois + '/' +annee;
  console.log("RECHERCHE COLONNE");
  async.series([
    function (callback) {
      ReportingInovcomExport.countOkKofll9("recherchefactureinteriale",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
    async.series([          
      function (callback) {
        ReportingInovcomExport.ecritureOkKo9(result[0],"recherchefactureinteriale",date_export,mois1,callback);
      },
    ],function(err,resultExcel){
   
        if(resultExcel[0]==true)
        {
          console.log("true zn");
          res.view('Inovcom/erera');
        }
        if(resultExcel[0]=='OK')
        {
          // res.redirect('/exportInovcom/'+date_export+'/x')
          res.view('reporting/succes');
        }
    })
  })
},
};

