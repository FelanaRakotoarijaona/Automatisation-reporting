/**
 * Route Mappings
 * (sails.config.routes)
 *
 * Your routes tell Sails what to do each time it receives a request.
 *
 * For more information on configuring custom routes, check out:
 * https://sailsjs.com/anatomy/config/routes-js
 */

module.exports.routes = {


     //route tps-tpc grs
     '/accueilstockj1et2et5' : 'TpsGrsController.accueilstockj1et2et5',
     '/traitementstockj1et2et5' : 'TpsGrsController.traitementstockj1et2et5',
     '/accueilEtpGrs' : 'TpsGrsController.accueilEtp',
     '/copieEtp' : 'TpsGrsController.copieEtp',
     '/accueilGrs' : 'TpsGrsController.accueil',
     '/traitementTacheTraiteGrs' : 'TpsGrsController.traitementTacheTraite',
     '/accueilrecherchefichier' : 'TpsGrsController.accueilrecherchefichier',
     '/recherchefichiertpsgrs' : 'TpsGrsController.recherchefichier',
     '/accueilstock16h' : 'TpsGrsController.accueilstock16h',
     '/traitementgrsstock16h' : 'TpsGrsController.traitementgrsstock16h',


     '/recherchefichiertpstpc' : 'TpstpcController.recherchefichiertpstpc',
     '/accueilTpstpc' : 'TpstpcController.accueil1',
     '/accueilstocketbonj' : 'TpstpcController.accueil2',
     '/accueiltachenontraite' : 'TpstpcController.accueil3',
     '/accueiletp' : 'TpstpcController.accueil',
     '/accueilecriture' : 'TpstpcController.accueil6',
     '/traitementetp' : 'TpstpcController.selection',
     '/accueilecritureetp' : 'TpstpcController.accueiletp',
     '/ecritureetp' : 'TpstpcController.ecritureEtp',
     '/traitementfinal' :'TpstpcController.ecriture3',
     '/ecritureetp2' : 'TpstpcController.ecritureEtp2',
     '/ecritureerreur' : 'TpstpcController.ecritureErreur',
   
     '/traitementtachenontraite' : 'TpstpcController.traitementBonJ1',
     '/accueilsanteclair' : 'TpstpcController.accueil4',
     '/traitementSanteclair' : 'TpstpcController.traitementSanteclair',
     '/accueilerreur' : 'TpstpcController.accueil5',
     '/traitementerreureasy' : 'TpstpcController.traitementErreurEasy',
     '/traitementTpstpc' : 'TpstpcController.traitementTacheTraite',
     '/traitementsuivant' : 'TpstpcController.ecritureExcel',
     
     '/traitementJ2' : 'TpstpcController.traitementStocketBonJ',
  
   //route tps-tpc
   '/accueilrecherchefichier' : 'TpstpcController.accueilrecherchefichier',
   '/recherchefichiertpstpc' : 'TpstpcController.recherchefichiertpstpc',
   '/accueilTpstpc' : 'TpstpcController.accueil1',
   '/accueilstocketbonj' : 'TpstpcController.accueil2',
   '/accueiltachenontraite' : 'TpstpcController.accueil3',
   '/accueiletp' : 'TpstpcController.accueil',
   '/accueilecriture' : 'TpstpcController.accueil6',
   '/traitementetp' : 'TpstpcController.selection',
   '/accueilecritureetp' : 'TpstpcController.accueiletp',
   '/ecritureetp' : 'TpstpcController.ecritureEtp',
   '/traitementfinal' :'TpstpcController.ecriture3',
   '/ecritureetp2' : 'TpstpcController.ecritureEtp2',
   '/ecritureerreur' : 'TpstpcController.ecritureErreur',
 
   '/traitementtachenontraite' : 'TpstpcController.traitementBonJ1',
   '/accueilsanteclair' : 'TpstpcController.accueil4',
   '/traitementSanteclair' : 'TpstpcController.traitementSanteclair',
   '/accueilerreur' : 'TpstpcController.accueil5',
   '/traitementerreureasy' : 'TpstpcController.traitementErreurEasy',
   '/traitementTpstpc' : 'TpstpcController.traitementTacheTraite',
   '/traitementsuivant' : 'TpstpcController.ecritureExcel',
   
   '/traitementJ2' : 'TpstpcController.traitementStocketBonJ',


  //Route Login
  '/': 'AuthentificationController.loginSimple',
  '/login' : 'AuthentificationController.loginLdap',
  '/logout' : 'AuthentificationController.logout',
  '/test': { view: 'pages/navbar' },

  //AJOUT PAGES DEBUT
  '/accueildebut1': { view: 'reporting/accueil1' },
  '/accueildebut2': { view: 'reporting/accueilG' },
  
  //Route HTP
  '/accueilhtp' : 'ReportinghtpController.accueil1',
  '/essai' : 'ReportinghtpController.Essaii',
  '/accueil/:date' : 'ReportinghtpController.accueil',
  //'/import' : 'ReportinghtpController.import',
  '/reportinghtp' : 'ReportinghtpController.essaiExcel',
  //Route HTP Export
  //'/export/:jour/:mois/:annee/:html' : 'ReportingExcelController.accueil',
  //'/exportReporting/:jour/:mois/:annee' : 'ReportingExcelController.rechercheColonne',
  '/exportExcel' : 'ReportingExcelController.rechercheColonne',
  '/exportExcelHTP' : 'ReportingExcelController.rechercheColonne',
  '/exportExcelH' : 'ReportingExcelController.accueilHTP',

  //Route HTP2
  '/accueil2' : 'ReportinghtpController.accueiltype2',
  '/essai2' : 'ReportinghtpController.Essaiitype2',
  '/accueil2/:date' : 'ReportinghtpController.accueiltype2',
  //'/import' : 'ReportinghtpController.import',
  '/reportinghtp2' : 'ReportinghtpController.essaiExcel2',
  //Route HTP Export
  //'/export/:jour/:mois/:annee/:html' : 'ReportingExcelController.accueil',
  //'/exportReporting/:jour/:mois/:annee' : 'ReportingExcelController.rechercheColonne',
  '/exportExcel' : 'ReportingExcelController.rechercheColonne',
 

  //Route INOVCOM
  '/accueilInovcom' : 'ReportingInovcomController.accueil1',
  '/essaiInovcom' : 'ReportingInovcomController.Essaii',
  '/accueil2Inovcom/:date' : 'ReportingInovcomController.accueil',
  '/reportinginovcom' : 'ReportingInovcomController.essaiExcel',

  //type2
  '/accueilInovcomtype2' : 'ReportingInovcomController.accueil1type2',
  '/essaiInovcomtype2' : 'ReportingInovcomController.Essaiitype2',
  
   //type4
   '/accueilInovcomtype4' : 'ReportingInovcomController.accueil1type4',
   '/essaiInovcomtype4' : 'ReportingInovcomController.Essaiitype4',
   '/reportinginovcomtype4' : 'ReportingInovcomController.essaiExceltype4',


    //type14
    '/accueilInovcomtype14' : 'ReportingInovcomController.accueil1type14',
    '/essaiInovcomtype14' : 'ReportingInovcomController.Essaiitype14',
    '/reportinginovcomtype14' : 'ReportingInovcomController.essaiExceltype14',


   //type5
   '/accueilInovcomtype5' : 'ReportingInovcomController.accueil1type5',
   '/essaiInovcomtype5' : 'ReportingInovcomController.Essaiitype5',
   '/reportinginovcomtype5' : 'ReportingInovcomController.essaiExceltype5',

    //type6
    '/accueilInovcomtype6' : 'ReportingInovcomController.accueil1type6',
    '/essaiInovcomtype6' : 'ReportingInovcomController.Essaiitype6',
    '/reportinginovcomtype6' : 'ReportingInovcomController.essaiExceltype6',

   //type3
   '/accueilInovcomtype3' : 'ReportingInovcomController.accueil1type3',
   '/essaiInovcomtype3' : 'ReportingInovcomController.Essaiitype3',
   '/reportinginovcomtype3' : 'ReportingInovcomController.essaiExceltype3',

    //type7
    '/accueilInovcomtype7' : 'ReportingInovcomController.accueil1type7',
    '/essaiInovcomtype7' : 'ReportingInovcomController.Essaiitype7',
    '/reportinginovcomtype7' : 'ReportingInovcomController.essaiExceltype7',

     //type8
     '/accueilInovcomtype8' : 'ReportingInovcomController.accueil1type8',
     '/essaiInovcomtype8' : 'ReportingInovcomController.Essaiitype8',
     '/reportinginovcomtype8' : 'ReportingInovcomController.essaiExceltype8',

     //type9
     '/accueilInovcomtype9' : 'ReportingInovcomController.accueil1type9',
     '/essaiInovcomtype9' : 'ReportingInovcomController.Essaiitype9',
     '/reportinginovcomtype8' : 'ReportingInovcomController.essaiExceltype8',
     '/reportinginovcomtype2' : 'ReportingInovcomController.essaiExceltype2',

     //type11
     '/accueilInovcomtype11' : 'ReportingInovcomController.accueil1type11',
     '/essaiInovcomtype11' : 'ReportingInovcomController.Essaiitype11',
     '/reportinginovcomtype11' : 'ReportingInovcomController.essaiExceltype11',

      //type10
      '/accueilInovcomtype10' : 'ReportingInovcomController.accueil1type10',
      '/essaiInovcomtype10' : 'ReportingInovcomController.Essaiitype10',
      '/reportinginovcomtype10' : 'ReportingInovcomController.essaiExceltype10',

      //type12
     '/accueilInovcomtype12' : 'ReportingInovcomController.accueil1type12',
     '/essaiInovcomtype12' : 'ReportingInovcomController.Essaiitype12',
     '/reportinginovcomtype12' : 'ReportingInovcomController.essaiExceltype12',
  


  //Route INOVCOM Export
  // '/exportInovcom/:jour/:mois/:annee/:html' : 'ReportingInovcomExportController.accueil',
  // '/exportReportingInovcom/:jour/:mois/:annee' : 'ReportingInovcomExportController.rechercheColonne',

  '/exportExcelInovcom1' : 'ReportingInovcomExportController.rechercheColonne1',
  '/exportExcelInov1' : 'ReportingInovcomExportController.accueilInov1',

  '/exportExcelInovcom2' : 'ReportingInovcomExportController.rechercheColonne2',
  '/exportExcelInov2' : 'ReportingInovcomExportController.accueilInov2',
  '/exportExcelInovcom2suivant' : 'ReportingInovcomExportController.rechercheColonne2suivant1',//ajout suivant inov2
  '/exportExcelInovcom2suivant2' : 'ReportingInovcomExportController.rechercheColonne2suivant2',

  '/exportExcelInovcom3' : 'ReportingInovcomExportController.rechercheColonne3',
  '/exportExcelInov3' : 'ReportingInovcomExportController.accueilInov3',

  '/exportExcelInovcom4' : 'ReportingInovcomExportController.rechercheColonne4',
  '/exportExcelInov4' : 'ReportingInovcomExportController.accueilInov4',

  

  '/exportExcelInovcom5' : 'ReportingInovcomExportController.rechercheColonne5',
  '/exportExcelInov5' : 'ReportingInovcomExportController.accueilInov5',

  '/exportExcelInovcom6' : 'ReportingInovcomExportController.rechercheColonne6',
  '/exportExcelInov6' : 'ReportingInovcomExportController.accueilInov6',

  '/exportExcelInovcom7' : 'ReportingInovcomExportController.rechercheColonne7',
  '/exportExcelInov7' : 'ReportingInovcomExportController.accueilInov7',

  '/exportExcelInovcom8' : 'ReportingInovcomExportController.rechercheColonne8',
  '/exportExcelInov8' : 'ReportingInovcomExportController.accueilInov8',

  '/exportExcelInovcom9' : 'ReportingInovcomExportController.rechercheColonne9',
  '/exportExcelInov9' : 'ReportingInovcomExportController.accueilInov9',

  //EXPORT INOVCOM Nombre de ligne(2,3,4)
  '/exportExcelInovcom10' : 'ReportingInovcomExportController.rechercheColonne10',
  '/exportExcelInovcom11' : 'ReportingInovcomExportController.rechercheColonne11',
  '/exportExcelInovcom2suivant11' : 'ReportingInovcomExportController.rechercheColonne11cbtp',
  '/exportExcelInovcom12' : 'ReportingInovcomExportController.rechercheColonne12',
  '/exportExcelInovcom14' : 'ReportingInovcomExportController.rechercheColonne14',

   //Route INDU
   '/accueilIndu' : 'ReportingInduController.accueil1',
   '/essaiIndu' : 'ReportingInduController.Essaii',
   '/accueil2Indu/:date' : 'ReportingInduController.accueil',
   '/reportingindu' : 'ReportingInduController.essaiExcel',

   '/accueilIndu2' : 'ReportingInduController.accueiltype2',
   '/essaiIndu2' : 'ReportingInduController.Essaii2',
   '/accueil2Indu/:date' : 'ReportingInduController.accueil2',
   '/reportingindu2' : 'ReportingInduController.essaiExcel2',

   
   '/accueilIndu3' : 'ReportingInduController.accueiltype3',
   '/essaiIndu3' : 'ReportingInduController.Essaii3',
   '/accueil3Indu/:date' : 'ReportingInduController.accueil3',
   '/reportingindu3' : 'ReportingInduController.essaiExcel3',


   //Route INDU Export mbola tsy traité
  //  '/exportIndu/:jour/:mois/:annee/:html' : 'ReportingInovcomExportController.accueil',
  //  '/exportReportingIndu:jour/:mois/:annee' : 'ReportingInovcomExportController.rechercheColonne',
   '/exportExcelIndu' : 'ReportingInduController.rechercheColonne',
   '/exportExcelIndusuivant' : 'ReportingInduController.rechercheColonneindusuivant',
   '/exportExcelin' : 'ReportingInduController.accueilI',

   '/exportExcelIndu2' : 'ReportingInduController.rechercheColonne2',
   '/exportExcelin2' : 'ReportingInduController.accueilI2',

   '/exportExcelIndu3' : 'ReportingInduController.rechercheColonne3',
   '/exportExcelin3' : 'ReportingInduController.accueilI3',

  
   //Route RETOUR
   '/accueilRetour' : 'ReportingRetourController.accueil1',//1
   '/essaiRetour' : 'ReportingRetourController.Essaii',//2
   '/accueil2Retour/:date' : 'ReportingRetourController.accueil',
   '/reportingretour' : 'ReportingRetourController.essaiExcel',//3
   //Route RETOUR Export mbola tsy traité
  //  '/exportRetour/:jour/:mois/:annee/:html' : 'ReportingInovcomExportController.accueil',
  //  '/exportReportingIndu:jour/:mois/:annee' : 'ReportingInovcomExportController.rechercheColonne',
   //ajout routes pour RETOUR
   '/exportExcelRetour' : 'ReportingRetourController.rechercheColonne',//A DECOMMENTER SI EXCEL RETOUR CORRECTE
   '/exportExcelRet' : 'ReportingRetourController.accueilR',//4
   '/exportretoursuivant' : 'ReportingRetourController.rechercheColonne_suivant',
  //  '/exportExcelRetour' : 'ReportingRetourController.rechercheColonnetest',//5 test
 
     //Route CONTENTIEUX
    //  '/accueilContentieux' : 'ReportingContetieuxController.accueil1',
    //  '/essaiContentieux' : 'ReportingContetieuxController.Essaii',
    //  '/accueil2Contentieux/:date' : 'ReportingContetieuxController.accueil',
    //  '/reportingContentieux' : 'ReportingContetieuxController.essaiExcel',
     //Route CONTENTIEUX Export mbola tsy traité
    //  '/exportRetour/:jour/:mois/:annee/:html' : 'ReportingInovcomExportController.accueil',
    //  '/exportReportingIndu:jour/:mois/:annee' : 'ReportingInovcomExportController.rechercheColonne',
     //ajout routes pour RETOUR
     '/exportExcelContentieux' : 'ReportingContetieuxController.rechercheColonne',
     '/exportExcelCont' : 'ReportingContetieuxController.accueilCont',
    //Route CONTETIEUX
    '/accueilContetieux' : 'ReportingContetieuxController.accueil1',
    '/essaiContetieux' : 'ReportingContetieuxController.Essaii',
    '/accueil2Retour/:date' : 'ReportingRetourController.accueil',
    '/reportingcontetieux' : 'ReportingContetieuxController.essaiExcel',
    //Route RETOUR Export mbola tsy traité
    // '/exportRetour/:jour/:mois/:annee/:html' : 'ReportingInovcomExportController.accueil',
    // '/exportReportingIndu:jour/:mois/:annee' : 'ReportingInovcomExportController.rechercheColonne',
    //ROUTES TEST CONTENTIEUX
    '/testcont' : 'ReportingContetieuxController.testcont',
  
  // ROUTES REPORTING ALMERYS ENGAGEMENT
  '/accueilGarantie' : 'GarantieController.accueilGarantie',//accueilGarantie
  '/insertGarantiesansdouble' : 'GarantieController.insertChemingarantiesansdouble',
  '/insertGarantieligne' : 'GarantieController.insertChemingarantieligne',
  '/insertGarantiebpo1' : 'GarantieController.insertChemingarantiebpo1',
  '/reportingGarantiesansdouble' : 'GarantieController.importGarantiesansdouble',
  '/reportingGarantieligne' : 'GarantieController.importGarantieligne',
  '/reportingGarantiebpo1' : 'GarantieController.importGarantiebpo1',

   //ROUTES HTP REPORTING ENGAGEMENT
   '/accueilengagementhtp' : 'EngagementhtpController.accueilengagementhtp',
   '/insertengagementhtp' : 'EngagementhtpController.insertcheminengagementhtp',
   '/reportingengagementhtp' : 'EngagementhtpController.importengagementhtp',


   //'/cheminHTPengagement' : { view: 'HTPengagement/exportHTPengagement' }, EXPORT
   '/exportHTPengagement' : 'EngagementhtpController.exporthtpengagement',
   '/exportHTPengagementsuivant_1' : 'EngagementhtpController.exporthtpengagementsuivant_1',
   '/exportHTPengagementsuivant_2' : 'EngagementhtpController.exporthtpengagementsuivant_2',
   '/exportHTPengagementsuivant_3' : 'EngagementhtpController.exporthtpengagementsuivant_3',

  /***************************************************************************
  *                                                                          *
  * More custom routes here...                                               *
  * (See https://sailsjs.com/config/routes for examples.)                    *
  *                                                                          *
  * If a request to a URL doesn't match any of the routes in this file, it   *
  * is matched against "shadow routes" (e.g. blueprint routes).  If it does  *
  * not match any of those, it is matched against static assets.             *
  *                                                                          *
  ***************************************************************************/


};
