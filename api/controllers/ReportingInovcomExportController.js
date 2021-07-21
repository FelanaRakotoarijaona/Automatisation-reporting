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
     

      ],function(err,result){
        if(err) return res.badRequest(err);
        
        else{
        console.log("Count OK 0 ==> " + result[0].ok);  
        console.log("Count OK 1 ==> " + result[1].ok);
        console.log("Count OK 2 ==> " + result[2].ok);
        console.log("Count OK 3 ==> " + result[3].ok);
        console.log("Count OK 4 ==> " + result[4].ok);
        console.log("Count OK 5 ==> " + result[5].ok);
        console.log("Count OK 6 ==> " + result[6].ok);
       
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
        

          ],function(err,resultExcel){
            console.log('**************');
            console.log(resultExcel[0]);
            console.log('**************');
            if(resultExcel[0]==true)
            {
              console.log("true zn");
              res.view('Inovcom/erera');
            }
            else
            {
              return res.view('Inovcom/exportsuivantinovcom2', {date: datetest});
              // res.view('reporting/succes');

            }
        });//fermeture async 2

      };

    });//fermeture async 1
    
    },
    /*********************************************************************************/
    rechercheColonne2suivant1: function (req, res) {
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
          ReportingInovcomExport.countok("santeclairtableauretourgeneral",callback);
        },
        function (callback) {
          ReportingInovcomExport.countok("santeclairoptique",callback);
        },
        //tranferer dans le boutton Nombre de ligne(3)
        // function (callback) {
        //   ReportingInovcomExport.countok("noemiehtpmgefi",callback);
        // },
        // function (callback) {
        //   ReportingInovcomExport.countok("mgefigtomgefirejetsaisienoemiehtp",callback);
        // },
      ],function(err,result){
        if(err) return res.badRequest(err);
        console.log("Count OK suiv_1 0 ==> " + result[0].ok);
        console.log("Count OK suiv_1 1 ==> " + result[1].ok);
        async.series([
          function (callback) {
            ReportingInovcomExport.ecritureOkKo21(result[0],"santeclairtableauretourgeneral",date_export,mois1,callback);
          },  
          function (callback) {
            ReportingInovcomExport.ecritureOkKo21(result[1],"santeclairoptique",date_export,mois1,callback);
          },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo22(result[2],"noemiehtpmgefi",date_export,mois1,callback);
          // },
          // function (callback) {
          //   ReportingInovcomExport.ecritureOkKo22(result[3],"mgefigtomgefirejetsaisienoemiehtp",date_export,mois1,callback);
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
              return res.view('Inovcom/exportsuivantinovcom3', {date: datetest});
              // res.view('reporting/succes');
            }
        });
      });
    },
    /*********************************************************************************/
    rechercheColonne2suivant2: function (req, res) {
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
        console.log("Count OK suiv_2 0 ==> " + result[0].ok);
        console.log("Count OK suiv_2 1 ==> " + result[1].ok);
        console.log("Count OK suiv_2 2 ==> " + result[2].ok);
        console.log("Count OK suiv_2 3 ==> " + result[3].ok);
        console.log("Count OK suiv_2 4 ==> " + result[4].ok);
        console.log("Count OK suiv_2 5 ==> " + result[5].ok);
        async.series([
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[0],"retourreclamtramereclamationtiers",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[1],"reclamsetramereclamationse",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[2],"reclamhospi",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[3],"dentairereclamationdentaire",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[4],"optiquetramereclamationoptique",date_export,mois1,callback);
          },
          function (callback) {
            ReportingInovcomExport.ecritureOkKo23(result[5],"reclamationaudio",date_export,mois1,callback);
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
        function (callback) {
          ReportingInovcomExport.countok("psfemajagaps",callback);
        },
      ],function(err,result){
        if(err) return res.badRequest(err);
        console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
        console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
        console.log("Count OK 2 ==> " + result[2].ok + " / " + result[2].ko);
        console.log("Count OK 3 ==> " + result[3].ok + " / " + result[3].ko);
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
          function (callback) {
            ReportingInovcomExport.ecritureOkKo3(result[3],"psfemajagaps",date_export,mois1,callback);
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
    // function (callback) {
    //   ReportingInovcomExport.countOkKofll4("extractionrcforce",callback);
    // },
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
      ReportingInovcomExport.countOkKofll4("favpharma",callback);
    },
     function (callback) {
      ReportingInovcomExport.countOkKofll4("favnument",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
    console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
    console.log("Count OK 2 ==> " + result[2].ok + " / " + result[2].ko);
    console.log("Count OK 3 ==> " + result[3].ok + " / " + result[3].ko);
    console.log("Count OK 4 ==> " + result[4].ok + " / " + result[4].ko);
    async.series([  
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4(result[0],"faveole",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4(result[1],"favmgefi",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4(result[2],"favbalma",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4etat(result[3],"favpharma",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4pack(result[4],"favnument",date_export,mois1,callback);
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
//RECHERCHE DU COLONNE DU RETOUR FAV
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
      ReportingInovcomExport.countOkKo6("fav",callback);
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
//RECHERCHE DU COLONNE DU RETOUR CMUC
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
      ReportingInovcomExport.countOkKo6("retourcmuc",callback);
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
    function (callback) {
      ReportingInovcomExport.countOkKofll7("defraiment",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 0 ==> " + result[0].ok + " / " + result[0].ko);
    console.log("Count OK 1 ==> " + result[1].ok + " / " + result[1].ko);
    async.series([          
      function (callback) {
        ReportingInovcomExport.ecritureOkKo7bis(result[0],"hospidematrejetprive",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo7(result[1],"defraiment",date_export,mois1,callback);
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
    console.log("Count OK 0 ==> " + result[0].ok);
    console.log("Count OK 1 ==> " + result[1].ok);
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
/**********************************************************************************************************/
//EXPORT EXCEL nombre de ligne(3)
rechercheColonne10: function (req, res) {
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
    //tranferer dans le boutton Nombre de ligne(3)
    function (callback) {
      ReportingInovcomExport.countok("noemiehtpmgefi",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("mgefigtomgefirejetsaisienoemiehtp",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK Rojo 0 ==> " + result[0].ok);
    console.log("Count OK Rojo 1 ==> " + result[1].ok);
    async.series([
      function (callback) {
        ReportingInovcomExport.ecritureOkKo22(result[0],"noemiehtpmgefi",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKo22(result[1],"mgefigtomgefirejetsaisienoemiehtp",date_export,mois1,callback);
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
          // return res.view('Inovcom/exportsuivantinovcom3', {date: datetest});
          res.view('reporting/succes');
        }
    });
  });
},
/********************************************************************************************************/
//EXPORT EXCEL nombre de ligne(2) ATTENTE CONSIGNE POUR LES 2 EXPORTS
rechercheColonne11: function (req, res) {
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
      ReportingInovcomExport.countok("inovtpsalmerys",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovsealmerys",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovspehospi",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovpackspedentaire",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovpackspeoptique",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovspeaudio",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("santeclairaudio",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK fll_11 0 ==> " + result[0].ok);
    console.log("Count OK fll_11 1 ==> " + result[1].ok);
    console.log("Count OK fll_11 2 ==> " + result[2].ok);
    console.log("Count OK fll_11 3 ==> " + result[3].ok);
    console.log("Count OK fll_11 4 ==> " + result[4].ok);
    console.log("Count OK fll_11 5 ==> " + result[5].ok);
    async.series([
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll11(result[0],"inovtpsalmerys",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll11(result[1],"inovsealmerys",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll11(result[2],"inovspehospi",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll11(result[3],"inovpackspedentaire",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll11(result[4],"inovpackspeoptique",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll11(result[5],"inovspeaudio",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll11sante(result[6],"santeclairaudio",date_export,mois1,callback);
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
          return res.view('Inovcom/exportsuivantinovcom11', {date: datetest});
          // res.view('reporting/succes');
        }
    });
  });
},

/********************************************************************************************************/
rechercheColonne11cbtp: function (req, res) {
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
      ReportingInovcomExport.countok("inovtpscbtp",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovsecbtp",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK fll_11cbtp 0 ==> " + result[0].ok);
    console.log("Count OK fll_11cbtp 1 ==> " + result[1].ok);
    async.series([
     
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll11cbtp(result[0],"inovtpscbtp",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll11cbtp(result[1],"inovsecbtp",date_export,mois1,callback);
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
          // return res.view('Inovcom/exportsuivantinovcom3', {date: datetest});
          res.view('reporting/succes');
        }
    });
  });
},

/********************************************************************************************************/
rechercheColonne12: function (req, res) {
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
      ReportingInovcomExport.countok("inovaglaesynthese",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovaglaefraudemms",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovaglaeag2r",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovaglaefraudeinteriale",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovaglaefraudemg",callback);
    },
    function (callback) {
      ReportingInovcomExport.countok("inovaglae100",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK fll_12 Aglae 0 ==> " + result[0].ok);
    console.log("Count OK fll_12 Aglae 1 ==> " + result[1].ok);
    console.log("Count OK fll_12 Aglae 2 ==> " + result[2].ok);
    console.log("Count OK fll_12 Aglae 3 ==> " + result[3].ok);
    console.log("Count OK fll_12 Aglae 4 ==> " + result[4].ok);
    console.log("Count OK fll_12 Aglae 5 ==> " + result[5].ok);
    async.series([
     
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll12(result[0],"inovaglaesynthese",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll12(result[1],"inovaglaefraudemms",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll12(result[2],"inovaglaeag2r",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll12(result[3],"inovaglaefraudeinteriale",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll12(result[4],"inovaglaefraudemg",date_export,mois1,callback);
      },
      function (callback) {
        ReportingInovcomExport.ecritureOkKofll12retours(result[5],"inovaglae100",date_export,mois1,callback);
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
          // return res.view('Inovcom/exportsuivantinovcom3', {date: datetest});
          res.view('reporting/succes');
        }
    });
  });
},
/**********************************************************************************/
rechercheColonne14: function (req, res) {
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
      ReportingInovcomExport.countOkKofll4("rcindeterminable",callback);
    },
  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK fll-14 0 ==> " + result[0].ok + " / " + result[0].ko);
    console.log("Count OK fll-14 1 ==> " + result[1].ok + " / " + result[1].ko);
    async.series([          
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4(result[0],"extractionrcforce",date_export,mois1,callback);
      },     
      function (callback) {
        ReportingInovcomExport.ecritureOkKo4(result[1],"rcindeterminable",date_export,mois1,callback);
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
};

