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
  //Route Login
  '/': 'AuthentificationController.loginSimple',
  '/login' : 'AuthentificationController.loginLdap',
  '/logout' : 'AuthentificationController.logout',
  
  //Route HTP
  '/accueil1' : 'ReportinghtpController.accueil1',
  '/essai' : 'ReportinghtpController.Essaii',
  '/accueil/:date' : 'ReportinghtpController.accueil',
  //'/import' : 'ReportinghtpController.import',
  '/reportinghtp' : 'ReportinghtpController.essaiExcel',
  //Route HTP Export
  //'/export/:jour/:mois/:annee/:html' : 'ReportingExcelController.accueil',
  //'/exportReporting/:jour/:mois/:annee' : 'ReportingExcelController.rechercheColonne',
  '/exportExcel' : 'ReportingExcelController.rechercheColonne',

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
  //Route INOVCOM Export
  '/exportInovcom/:jour/:mois/:annee/:html' : 'ReportingInovcomExportController.accueil',
  '/exportReportingInovcom/:jour/:mois/:annee' : 'ReportingInovcomExportController.rechercheColonne',

   //Route INDU
   '/accueilIndu' : 'ReportingInduController.accueil1',
   '/essaiIndu' : 'ReportingInduController.Essaii',
   '/accueil2Indu/:date' : 'ReportingInduController.accueil',
   '/reportingindu' : 'ReportingInduController.essaiExcel',
   //Route INDU Export mbola tsy traité
   '/exportIndu/:jour/:mois/:annee/:html' : 'ReportingInovcomExportController.accueil',
   '/exportReportingIndu:jour/:mois/:annee' : 'ReportingInovcomExportController.rechercheColonne',

   //Route RETOUR
   '/accueilRetour' : 'ReportingRetourController.accueil1',
   '/essaiRetour' : 'ReportingRetourController.Essaii',
   '/accueil2Retour/:date' : 'ReportingRetourController.accueil',
   '/reportingretour' : 'ReportingRetourController.essaiExcel',
   //Route RETOUR Export mbola tsy traité
   '/exportRetour/:jour/:mois/:annee/:html' : 'ReportingInovcomExportController.accueil',
   '/exportReportingIndu:jour/:mois/:annee' : 'ReportingInovcomExportController.rechercheColonne',
 
    //Route CONTETIEUX
    '/accueilContetieux' : 'ReportingContetieuxController.accueil1',
    '/essaiContetieux' : 'ReportingContetieuxController.Essaii',
    '/accueil2Retour/:date' : 'ReportingRetourController.accueil',
    '/reportingcontetieux' : 'ReportingContetieuxController.essaiExcel',
    //Route RETOUR Export mbola tsy traité
    '/exportRetour/:jour/:mois/:annee/:html' : 'ReportingInovcomExportController.accueil',
    '/exportReportingIndu:jour/:mois/:annee' : 'ReportingInovcomExportController.rechercheColonne',
  


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
