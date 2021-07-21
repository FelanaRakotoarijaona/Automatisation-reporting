/**
 * EngagementhtpController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

module.exports = {
    //fonction n'est pas encore en service
    exporthtpengagement1: function(req, res){
        console.log('commencer');
        var datetest = req.param("date",0);
        var annee = datetest.substr(0, 4);
        var mois = datetest.substr(5, 2);
        var jour = datetest.substr(8, 2);
        var feuille = annee+mois+'_EASY';
        console.log(feuille);

        var date_export = jour + '/' + mois + '/' +annee;
        console.log("RECHERCHE FEUILLE EXCEL");

        async.series([
            function (callback) {
                Engagementhtp.recupdata("trhospimulti",callback);
            },
           
          ],function(err,result){
            if(err) return res.badRequest(err);
            console.log("Count OK 0 ==> " + result[0].ok);
            async.series([
              function (callback) {
                Engagementhtp.ecriture(result[0],"trhospimulti",date_export,feuille,callback);
              },
        
            
            ],function(err,resultExcel){
           console.log(resultExcel[0]);
                if(resultExcel[0]==true)
                {
                  console.log("true zn");
                  res.view('Retour/erera');
                }
                if(resultExcel[0]=='OK')
                {
                  // res.redirect('/exportRetour/'+date_export+'/x')
                  res.view('Retour/succes');
                }
                             
              
            })
          })

    },
    //FONCTION POUR L'EXPORT TACHES TRAITEES
    exporthtpengagement : function (req, res) {
        var datetest = req.param("date",0);
        var annee = datetest.substr(0, 4);
        var mois = datetest.substr(5, 2);
        var jour = datetest.substr(8, 2);
        var feuille = annee+mois+'_EASY';
        console.log(feuille);
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
          mois1= 'octobre';
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
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
          function (callback) {
            Engagementhtp.recupdata("trhospimulti",callback);
          },
         
        ],function(err,result){
          if(err) return res.badRequest(err);
          console.log("Count OK 1==> " + result[0].ok);
          async.series([
            function (callback) {
              Engagementhtp.ecrituredata16tri(result[0],"trhospimulti",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16facM(result[1],"trhospimulti",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16devi(result[2],"trhospimulti",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16sales(result[3],"trhospimulti",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16flux(result[4],"trhospimulti",date_export,feuille,callback);
            },
            function (callback) {
                Engagementhtp.ecrituredata16rejet(result[5],"trhospimulti",date_export,feuille,callback);
              },
            function (callback) {
            Engagementhtp.ecrituredata16cotlamie(result[6],"trhospimulti",date_export,feuille,callback);
            },
            function (callback) {
                Engagementhtp.ecrituredatafinptri(result[7],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                Engagementhtp.ecrituredatafinpfacM(result[8],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                Engagementhtp.ecrituredatafinpdevi(result[9],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                Engagementhtp.ecrituredatafinpsales(result[10],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                Engagementhtp.ecrituredatafinpflux(result[11],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                  Engagementhtp.ecrituredatafinprejet(result[12],"trhospimulti",date_export,feuille,callback);
                },
              function (callback) {
              Engagementhtp.ecrituredatafinpcotlamie(result[13],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                Engagementhtp.ecrituredata16cotite(result[14],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                Engagementhtp.ecrituredatafinpcotite(result[15],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                Engagementhtp.ecrituredata16faclamie(result[16],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                Engagementhtp.ecrituredata16acs(result[17],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                Engagementhtp.ecrituredatafinpfaclamie(result[18],"trhospimulti",date_export,feuille,callback);
              },
              function (callback) {
                Engagementhtp.ecrituredatafinpacs(result[19],"trhospimulti",date_export,feuille,callback);
              },
          
          ],function(err,resultExcel){
         console.log(resultExcel[0]);
              if(resultExcel[0]==true)
              {
                console.log("true zn");
                res.view('Retour/erera');
              }
              if(resultExcel[0]=='OK')
              {
                // res.redirect('/exportRetour/'+date_export+'/x')
                res.view('HTPengagement/exportHTPengagementsuivant_1', {date : datetest});
              }

          })
        })
      },
/**********************************************************************************************************/ 
//FONCTION EXPORT SUIVANT 1 (taches traitees)
exporthtpengagementsuivant_1 : function (req, res) {
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var feuille = annee+mois+'_EASY';
    console.log(feuille);
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
      mois1= 'octobre';
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
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },

    ],function(err,result){
      if(err) return res.badRequest(err);
      console.log("Count OK 1==> " + result[0].ok);
      async.series([
        function (callback) {
          Engagementhtp.ecrituredataj2tri(result[0],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataj2facM(result[1],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataj2devi(result[2],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataj2sales(result[3],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataj2flux(result[4],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredataj2rejet(result[5],"trhospimulti",date_export,feuille,callback);
          },
        function (callback) {
            Engagementhtp.ecrituredataj2cotlamie(result[6],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredataj5tri(result[7],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredataj5facM(result[8],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredataj5devi(result[9],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredataj5sales(result[10],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredataj5flux(result[11],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredataj5rejet(result[12],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredataj5cotlamie(result[13],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredataj2cotite(result[14],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredataj5cotite(result[15],"trhospimulti",date_export,feuille,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredataj2faclamie(result[16],"trhospimulti",date_export,feuille,callback);
          },
        function (callback) {
          Engagementhtp.ecrituredataj2acs(result[17],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataj5faclamie(result[18],"trhospimulti",date_export,feuille,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredataj5acs(result[19],"trhospimulti",date_export,feuille,callback);
        },
      
      
      ],function(err,resultExcel){
     console.log(resultExcel[0]);
          if(resultExcel[0]==true)
          {
            console.log("true zn");
            res.view('Retour/erera');
          }
          if(resultExcel[0]=='OK')
          {
            // res.redirect('/exportRetour/'+date_export+'/x')
            res.view('HTPengagement/exportHTPengagementsuivant_2', {date : datetest});
          }

      })
    })
  },

/**********************************************************************************************************/ 
//FONCTION EXPORT STOCKS
exporthtpengagementsuivant_2 : function (req, res) {
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var feuille = annee+mois+'_EASY';
    console.log(feuille);
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
      mois1= 'octobre';
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
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
    

    ],function(err,result){
      if(err) return res.badRequest(err);
      console.log("Count OK 1==> " + result[0].ok);
      async.series([
        function (callback) {
          Engagementhtp.ecrituredatastock16tri(result[0],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatastock16facM(result[1],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatastock16devi(result[2],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatastock16sales(result[3],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredatastock16flux(result[4],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredatastock16rejet(result[5],"trhospimulti",date_export,feuille,callback);
          },
        function (callback) {
            Engagementhtp.ecrituredatastock16cotlamie(result[6],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredatastock16cotite(result[7],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredatastock16faclamie(result[8],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
            Engagementhtp.ecrituredatastock16acs(result[9],"trhospimulti",date_export,feuille,callback);
        },
        
      
      ],function(err,resultExcel){
     console.log(resultExcel[0]);
          if(resultExcel[0]==true)
          {
            console.log("true zn");
            res.view('Retour/erera');
          }
          if(resultExcel[0]=='OK')
          {
            // res.redirect('/exportRetour/'+date_export+'/x')
            res.view('HTPengagement/exportHTPengagementsuivant_3', {date : datetest});
          }

      })
    })
  },

/**********************************************************************************************************/ 
//FONCTION EXPORT ETP
exporthtpengagementsuivant_3 : function (req, res) {
    var datetest = req.param("date",0);
    var annee = datetest.substr(0, 4);
    var mois = datetest.substr(5, 2);
    var jour = datetest.substr(8, 2);
    var feuille = annee+mois+'_EASY';
    console.log(feuille);
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
      mois1= 'octobre';
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
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      function (callback) {
        Engagementhtp.recupdata("trhospimulti",callback);
      },
      
    

    ],function(err,result){
      if(err) return res.badRequest(err);
      console.log("Count OK 1==> " + result[0].ok);
      async.series([
        function (callback) {
          Engagementhtp.ecrituredataetptri(result[0],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataetpfacM(result[1],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataetpdevi(result[2],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataetpfaclamie(result[3],"trhospimulti",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataetpcotlamie(result[4],"trhospimulti",date_export,feuille,callback);
        },
        
        
      
      ],function(err,resultExcel){
     console.log(resultExcel[0]);
          if(resultExcel[0]==true)
          {
            console.log("true zn");
            res.view('Retour/erera');
          }
          if(resultExcel[0]=='OK')
          {
            // res.redirect('/exportRetour/'+date_export+'/x')
            res.view('Retour/succes');
          }

      })
    })
  },



};

