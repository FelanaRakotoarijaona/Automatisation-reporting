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
  /*********************************/
  accueilengagementhtp : async function(req,res)
  {
    return res.view('HTPengagement/accueilreportingengagementhtp');
  },
  /***********************************************************************************/
  //FONCTION POUR L'IMPORT DU CHEMIN UTILISER (22-07-2021)
insertcheminengagementhtp : function(req,res)
{
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  var table = ['\\\\10.128.1.2\\bpo_almerys'];
  var datetest = req.param("date",0);
  var annee = datetest.substr(0, 4);
  var mois = datetest.substr(5, 2);
  var jour = datetest.substr(8, 2);
  var date = annee+mois+jour;
  var date_indus = jour+'.'+mois+'.'+annee;
  var datej_1 = annee+mois+jour -1;
  console.log(datej_1);
  var nomtable = [];
  var numligne = [];
  var numfeuille = [];
  var nomcolonne = [];
  var nomcolonne2 = [];
  var nomcolonne3 = [];
  console.log(date);
  var cheminp = [];
  var MotCle= [];
  var Sup= [];
  var nomBase = "cheminengagementhtp";
  workbook.xlsx.readFile('engagementhtp.xlsx')
      .then(function() {
        var newworksheet = workbook.getWorksheet('Feuil2');
        var nomColonne3 = newworksheet.getColumn(3);
        var numFeuille = newworksheet.getColumn(4);
        var nomColonne = newworksheet.getColumn(5);
        var nomTable = newworksheet.getColumn(6);
        var nomColonne2 = newworksheet.getColumn(7);
        var numLigne = newworksheet.getColumn(8);
        var cheminparticulier = newworksheet.getColumn(9);
        var motcle = newworksheet.getColumn(10);
        var suppleant = newworksheet.getColumn(11);

      numFeuille.eachCell(function(cell, rowNumber) {
        numfeuille.push(cell.value);
      });
      nomColonne.eachCell(function(cell, rowNumber) {
        nomcolonne.push(cell.value);
      });
      nomColonne2.eachCell(function(cell, rowNumber) {
        nomcolonne2.push(cell.value);
      });
      nomColonne3.eachCell(function(cell, rowNumber) {
        nomcolonne3.push(cell.value);
      });
      nomTable.eachCell(function(cell, rowNumber) {
        nomtable.push(cell.value);
      });
      numLigne.eachCell(function(cell, rowNumber) {
        numligne.push(cell.value);
      });
        cheminparticulier.eachCell(function(cell, rowNumber) {
          cheminp.push(cell.value);
      });
        motcle.eachCell(function(cell, rowNumber) {
          MotCle.push(cell.value);
      });
      suppleant.eachCell(function(cell, rowNumber) {
        Sup.push(cell.value);
      });

        
          console.log(cheminp[0]);
          console.log(MotCle[0]);
          async.series([  
              function(cb){
                  Garantie.deleteFromChemin_bpo1(table,cb);
                },
           function(cb){
                  Garantie.importdatacheminhtp(table,cheminp,date,MotCle,0,nomtable[0],numligne[0],numfeuille[0],nomcolonne[0],nomcolonne2[0],nomcolonne3[0],cb);
                },
                
          ],
          function(err, resultat){
            var sql4= "select count(chemin) as ok from "+nomBase+" ";
                    console.log(sql4);
                    Reportinghtp.getDatastore().sendNativeQuery(sql4 ,function(err, nc) {
                       nc = nc.rows;
                       console.log(nc);
                       console.log('nc'+nc[0].ok);
                       var f = parseInt(nc[0].ok);
                       console.log(f);
                          if (err){
                            return res.view('Inovcom/erreur');
                          }
                        if(f==0)
                          {
                            return res.view('Inovcom/erreur');
                          }
                          else
                          {
                            return res.view('Garantie/importGarantiesansdouble', {date : datetest});
                                                          
                          };
                      });
        });
      });
},
  /***********************************************************************************/
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
          Engagementhtp.recupdata("htptri16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpfacmg16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpdevi16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpsales16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpflux16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htprejet16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpcotlamie16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htptrifin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpfacmgfin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpdevifin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpsalesfin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpfluxfin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htprejetfin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpcotlamiefin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpcotite16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpcotitefin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpfaclamie16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpacs16",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpfaclamiefin",callback);
        },
        function (callback) {
          Engagementhtp.recupdata("htpacsfin",callback);
        },
       
      ],function(err,result){
        if(err) return res.badRequest(err);
        console.log("Count OK 1==> " + result[0].ok);
        async.series([
          function (callback) {
            Engagementhtp.ecrituredata16tri(result[0],"htptri16",date_export,feuille,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16facM(result[1],"htpfacmg16",date_export,feuille,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16devi(result[2],"htpdevi16",date_export,feuille,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16sales(result[3],"htpsales16",date_export,feuille,callback);
          },
          function (callback) {
            Engagementhtp.ecrituredata16flux(result[4],"htpflux16",date_export,feuille,callback);
          },
          function (callback) {
              Engagementhtp.ecrituredata16rejet(result[5],"htprejet16",date_export,feuille,callback);
            },
          function (callback) {
          Engagementhtp.ecrituredata16cotlamie(result[6],"htpcotlamie16",date_export,feuille,callback);
          },
          function (callback) {
              Engagementhtp.ecrituredatafinptri(result[7],"htptrifin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpfacM(result[8],"htpfacmgfin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpdevi(result[9],"htpdevifin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpsales(result[10],"htpsalesfin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpflux(result[11],"htpfluxfin",date_export,feuille,callback);
            },
            function (callback) {
                Engagementhtp.ecrituredatafinprejet(result[12],"htprejetfin",date_export,feuille,callback);
              },
            function (callback) {
            Engagementhtp.ecrituredatafinpcotlamie(result[13],"htpcotlamiefin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16cotite(result[14],"htpcotite16",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpcotite(result[15],"htpcotitefin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16faclamie(result[16],"htpfaclamie16",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredata16acs(result[17],"htpacs16",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpfaclamie(result[18],"htpfaclamiefin",date_export,feuille,callback);
            },
            function (callback) {
              Engagementhtp.ecrituredatafinpacs(result[19],"htpacsfin",date_export,feuille,callback);
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
      Engagementhtp.recupdata("htptrij2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevij2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpsalesj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfluxj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htprejetj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamiej2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htptrij5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevij5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpsalesj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfluxj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htprejetj5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamiej5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotitej2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotitej5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamiej2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacsj2",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamiej5",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacsj5",callback);
    },

  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredataj2tri(result[0],"htptrij2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataj2facM(result[1],"htpfacmgj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataj2devi(result[2],"htpdevij2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataj2sales(result[3],"htpsalesj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataj2flux(result[4],"htpfluxj2",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj2rejet(result[5],"htprejetj2",date_export,feuille,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredataj2cotlamie(result[6],"htpcotlamiej2",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5tri(result[7],"htptrij5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5facM(result[8],"htpfacmgj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5devi(result[9],"htpdevij5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5sales(result[10],"htpsalesj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5flux(result[11],"htpfluxj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5rejet(result[12],"htprejetj5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5cotlamie(result[13],"htpcotlamiej5",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj2cotite(result[14],"htpcotitej2",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredataj5cotite(result[15],"htpcotitej5",date_export,feuille,callback);
        },
        function (callback) {
          Engagementhtp.ecrituredataj2faclamie(result[16],"htpfaclamiej2",date_export,feuille,callback);
        },
      function (callback) {
        Engagementhtp.ecrituredataj2acs(result[17],"htpacsj2",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataj5faclamie(result[18],"htpfaclamiej5",date_export,feuille,callback);
      },
    function (callback) {
        Engagementhtp.ecrituredataj5acs(result[19],"htpacsj5",date_export,feuille,callback);
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
      Engagementhtp.recupdata("htptristock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgstock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevistock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpsalesstock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfluxstock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htprejetstock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamiestock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotitestock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamiestock",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpacsstock",callback);
    },
  

  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredatastock16tri(result[0],"htptristock",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastock16facM(result[1],"htpfacmgstock",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastock16devi(result[2],"htpdevistock",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastock16sales(result[3],"htpsalesstock",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredatastock16flux(result[4],"htpfluxstock",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastock16rejet(result[5],"htprejetstock",date_export,feuille,callback);
        },
      function (callback) {
          Engagementhtp.ecrituredatastock16cotlamie(result[6],"htpcotlamiestock",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastock16cotite(result[7],"htpcotitestock",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastock16faclamie(result[8],"htpfaclamiestock",date_export,feuille,callback);
      },
      function (callback) {
          Engagementhtp.ecrituredatastock16acs(result[9],"htpacsstock",date_export,feuille,callback);
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
      Engagementhtp.recupdata("htptrietp",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfacmgetp",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpdevietp",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpfaclamieetp",callback);
    },
    function (callback) {
      Engagementhtp.recupdata("htpcotlamieetp",callback);
    },
    
  

  ],function(err,result){
    if(err) return res.badRequest(err);
    console.log("Count OK 1==> " + result[0].ok);
    async.series([
      function (callback) {
        Engagementhtp.ecrituredataetptri(result[0],"htptrietp",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataetpfacM(result[1],"htpfacmgetp",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataetpdevi(result[2],"htpdevietp",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataetpfaclamie(result[3],"htpfaclamieetp",date_export,feuille,callback);
      },
      function (callback) {
        Engagementhtp.ecrituredataetpcotlamie(result[4],"htpcotlamieetp",date_export,feuille,callback);
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

