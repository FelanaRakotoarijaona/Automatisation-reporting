/**
 * RetourController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

module.exports = {
  
//AJOUT FONCTION RECHERCHECOLONNE POUR RETOUR
  rechercheColonne : function (req, res) {
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
    console.log(date_export);
    console.log("RECHERCHE COLONNE");
    async.series([
      function (callback) {
        Retour.countOkKo("trhospimulti",callback);
      },
   function (callback) {
        Retour.countOkKo("trstcdentaire",callback);
      },
      function (callback) {
        Retour.countOkKo("trstcoptique",callback);
      },
      function (callback) {
        Retour.countOkKo("trstcaudio",callback);
      },
     function (callback) {
        Retour.countOkKoSum("trretourfacttiers",callback);
      },
      // function (callback) {
      //   Retour.countOkKo("retouralmgto2",callback);
      // },
      function (callback) {
        Retour.countOkKo("trffacturehospi",callback);
      },
      function (callback) {
        Retour.countOkKo("trffacturedentaire",callback);
      },
      function (callback) {
        Retour.countOkKo("trfactureoptique",callback);
      },
      // function (callback) {
      //   Retour.countOkKo("retouralmgto6",callback);
      // },
      function (callback) {
        Retour.countOkKoSum("trhospi",callback);
      },
      function (callback) {
        Retour.countOkKo("trtramepecdentaire",callback);
      },
      // function (callback) {
      //   Retour.countOkKo("retouralmgto9",callback);
      // },
      function (callback) {
        Retour.countOkKo("trpecaudio",callback);
      },
      function (callback) {
        Retour.countOkKo("trldralmerys",callback);
      },
    //  function (callback) {
    //     Retour.countOkKo("retouralmftp1",callback);
    //   },
      function (callback) {
        Retour.countOkKo("trretourotdn2",callback);
      },
      function (callback) {
        Retour.countOkKo("trtre",callback);
      },
      // function (callback) {
      //   Retour.countOkKo("retouralmftp4",callback);
      // },
      // function (callback) {
      //   Retour.countOkKo("retouralmpackspe1",callback);
      // },
      // function (callback) {
      //   Retour.countOkKo("retouralmpackspe2",callback);
      // },
      // function (callback) {
      //   Retour.countOkKo("retouralmpackspe3",callback);
      // },
      function (callback) {
        Retour.countOkKo("trindunoehtp",callback);
      },
      // function (callback) {
      //   Retour.countOkKo("retouralmcbtpgto",callback);
      // },
      // function (callback) {
      //   Retour.countOkKo("retouretat1",callback);
      // },
      function (callback) {
        Retour.countOkKo("trcentredesoin",callback);
      },
      // function (callback) {
      //   Retour.countOkKo("retouretat3",callback);
      // },
      // function (callback) {
      //   Retour.countOkKo("retouretat4",callback);
      // },
      // function (callback) {
      //   Retour.countOkKo("retouretat5",callback);
      // },
      // function (callback) {
      //   Retour.countOkKo("retourpublipostage1",callback);
      // },
      // function (callback) {
      //   Retour.countOkKo("retourpublipostage2",callback);
      // },
      function (callback) {
        Retour.countOkKo("trhospimulti",callback);
      },
      // function (callback) {
      //   Retour.countOkKo("retourpublipostage4",callback);
      // },
      // function (callback) {
      //   Retour.countOkKo("retourpublipostage5",callback);
      // },
      // function (callback) {
      //   Retour.countOkKo("retourpublipostage6",callback);
      // },
     
    ],function(err,result){
      if(err) return res.badRequest(err);
      console.log("Count OK 1==> " + result[0].ok);
      console.log("Count OK 2 ==> " + result[1].ok );
      console.log("Count OK 3 ==> " + result[2].ok );
      console.log("Count OK 4 ==> " + result[3].ok );
      console.log("Count OK 1==> " + result[4].ok );
      console.log("Count OK 2 ==> " + result[5].ok );
      console.log("Count OK 3 ==> " + result[6].ok );
      console.log("Count OK 4 ==> " + result[7].ok );
      console.log("Count OK 1==> " + result[8].ok );
      console.log("Count OK 2 ==> " + result[9].ok );
      console.log("Count OK 3 ==> " + result[10].ok );
      console.log("Count OK 4 ==> " + result[11].ok );
      console.log("Count OK 1==> " + result[12].ok );
      console.log("Count OK 2 ==> " + result[13].ok );
      console.log("Count OK 3 ==> " + result[14].ok );
      console.log("Count OK 3 ==> " + result[15].ok );
      // console.log("Count OK tramelamiestock ==> " + result[4].ok + " / " + result[4].ko);
      // console.log("Count OK tramelamiestockResiliation ==> " + result[5].ok + " / " + result[5].ko);
      async.series([
        function (callback) {
          Retour.ecritureOkKo(result[0],"trhospimulti",date_export,mois1,callback);
        },
       function (callback) {
          Retour.ecritureOkKo2(result[1],"trstcdentaire",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo2(result[2],"trstcoptique",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo2(result[3],"trstcaudio",date_export,mois1,callback);
        },
      function (callback) {
          Retour.ecritureOkKo3(result[4],"trretourfacttiers",date_export,mois1,callback);
        },
        // function (callback) {
        //    Retour.ecritureOkKo3(result[5],"retouralmgto2",date_export,mois1,callback);
        // },
        function (callback) {
          Retour.ecritureOkKo3(result[5],"trffacturehospi",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo3(result[6],"trffacturedentaire",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo3(result[7],"trfactureoptique",date_export,mois1,callback);
        },
        // function (callback) {
        //   Retour.ecritureOkKo3(result[9],"retouralmgto6",date_export,mois1,callback);
        // },
        function (callback) {
          Retour.ecritureOkKo3(result[8],"trhospi",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo3(result[9],"trtramepecdentaire",date_export,mois1,callback);
        },
        // function (callback) {
        //   Retour.ecritureOkKo3(result[12],"retouralmgto9",date_export,mois1,callback);
        // },
        function (callback) {
          Retour.ecritureOkKo3(result[10],"trpecaudio",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo3(result[11],"trldralmerys",date_export,mois1,callback);
        },
      //  function (callback) {
      //     Retour.ecritureOkKo4(result[12],"retouralmftp1",date_export,mois1,callback);
      //   },
        function (callback) {
          Retour.ecritureOkKo4(result[12],"trretourotdn2",date_export,mois1,callback);
        },
        function (callback) {
          Retour.ecritureOkKo4(result[13],"trtre",date_export,mois1,callback);
        },
        // function (callback) {
        //   Retour.ecritureOkKo4(result[15],"retouralmftp4",date_export,mois1,callback);
        // },
        // function (callback) {
        //  Retour.ecritureOkKo5(result[19],"retouralmpackspe1",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo5(result[20],"retouralmpackspe2",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo5(result[21],"retouralmpackspe3",date_export,mois1,callback);
        // },
        function (callback) {
          Retour.ecritureOkKo5(result[14],"trindunoehtp",date_export,mois1,callback);
        },
      //  function (callback) {
      //     Retour.ecritureOkKo6(result[23],"retouralmcbtpgto",date_export,mois1,callback);
      //   },
      //  function (callback) {
      //     Retour.ecritureOkKo7(result[24],"retouretat1",date_export,mois1,callback);
      //   },
        function (callback) {
          Retour.ecritureOkKo7(result[15],"trcentredesoin",date_export,mois1,callback);
        },
        // function (callback) {
        //   Retour.ecritureOkKo7(result[26],"retouretat3",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo7(result[27],"retouretat4",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo7(result[28],"retouretat5",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[29],"retourpublipostage1",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[30],"retourpublipostage2",date_export,mois1,callback);
        // },
        function (callback) {
          Retour.ecritureOkKo8(result[16],"trhospimulti",date_export,mois1,callback);
        },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[32],"retourpublipostage4",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[33],"retourpublipostage5",date_export,mois1,callback);
        // },
        // function (callback) {
        //   Retour.ecritureOkKo8(result[34],"retourpublipostage6",date_export,mois1,callback);
        // },
      
      ],function(err,resultExcel){
     console.log(resultExcel[0]);
          if(resultExcel[0]==true)
          {
            console.log("true zn");
            res.view('reporting/erera');
          }
          if(resultExcel[0]=='OK')
          {
            // res.redirect('/exportRetour/'+date_export+'/x')
            res.view('reporting/succes');
          }


          /*console.log("Traitement terminé ===> "+ resultExcel[0]);
          console.log("Traitement terminé ===> "+ resultExcel[1]);
          console.log("Traitement terminé ===> "+ resultExcel[2]);
          console.log("Traitement terminé ===> "+ resultExcel[3]);
          console.log("Traitement terminé ===> "+ resultExcel[4]);
          var html = "Echec d'enregistrement";
          return res.redirect('/accueil');*/
          
        
        
      })
    })
  },
};

