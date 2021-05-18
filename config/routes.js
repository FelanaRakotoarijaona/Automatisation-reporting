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
 

  //Route INOVCOM
  '/accueilInovcom' : 'ReportingInovcomController.accueil1',
  '/essaiInovcom' : 'ReportingInovcomController.Essaii',
  '/accueil2Inovcom/:date' : 'ReportingInovcomController.accueil',
  '/reportinginovcom' : 'ReportingInovcomController.essaiExcel',
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
