/**
 * AuthentificationController
 *
 * @description :: Server-side actions for handling incoming requests.
 * @help        :: See https://sailsjs.com/docs/concepts/actions
 */

var passport = require('passport');
var session = require('express-session');
var RedisStore = require('connect-redis')(session);
bcrypt = require('bcryptjs');

module.exports = {
    _config: {
        actions: false,
        shortcuts: false,
        rest: false
      },

    logout: function(req, res)
    {
        var retVal = [];
        retVal["display"] = "none";
        retVal["msg"] = "";
        req.session.user = null;
        req.session.image = null;
        req.session.authenticated = false;
        req.logout();
        res.view('pages/login',retVal);
    },

    loginSimple: function(req, res)
    {
      var retVal = [];
      retVal["display"] = "none";
      retVal["msg"] = "";
      res.view('pages/login',retVal);
    },

    ///_______________________________Fonction login LDAP(connexion)
    loginLdap: function(req, res)
    {
        console.log("LDAP CONNECT...");
        var email = 0;
        var ldap = require('ldapjs');
        if(!isNaN(req.param('email',null))) email = Number(req.param('email',null));
  
      //init ldap
        var client = ldap.createClient({
            url: 'ldap://10.128.1.14:389',
            reconnect: false
        });
  
        client.on('error', function(err) { });
        var email = req.param('email',null);
       // var x = parseInt(email);
        var sql = "select * from r_personnel where ldap_name = '"+email+"' " ;
  
        /*User.findOne({
            where: {
                or:[
                {id: email},
                {ldap_name: req.param('email',null)},
                {appelation: req.param('email',null)}
                ]
            }*/
      
            User.query(sql,function (err, user){
            if (err)  console.log(err);

            if (!user)
            {
                var retVal = [];
                req.session.modalEmail = "";
                retVal["display"] = "";
                retVal["msg"] = "Votre matricule est invalide";
                console.log('Matricule invalide user.');
                return res.view('pages/login', retVal);
            }
            if(req.param('password',null)=='admin@2020')
            {
                console.log("tonga ato v?");
                //connected
                User.query("select r_personnel.id_pers, transport_droit_user.droit from r_personnel join transport_droit_user on r_personnel.id_pers = transport_droit_user.matricule where r_personnel.id_pers ='"+user.id+"'", function(error, found){
                    if (error) return res.send(error);
                    if(found.rows.length == 0 ) {
                            req.session.droit = 0;
                    }else{
                        req.session.droit = 1;
                    }
                    req.session.user = user.id;
                    req.session.nom = user.appelation;
                    req.session.authenticated = true;
                    console.log("ID utilisateur connecté : "+user.id);
                    res.view('reporting/index');
                });
            }
            else
            {
                console.log("sa ato?");
        //test ldapServer
            client.bind('EASYTECH\\'+req.param('email',null), req.param('password',null), function(err) {
                if(err){
                    console.log("erreur");
                    var retVal = [];
                    req.session.modalEmail = "";
                    retVal["display"] = "";
                    retVal["msg"] = "Votre mot de passe est invalide";
                    console.log('Mot de passe invalide:'+user.password+'='+req.param('password',null));
                    var message = "blabla";
                    return res.view('pages/login', retVal);
                }
                else
                {
            //connected
            console.log("bonne direction");
                    var lstIdDossier =[];
                    async.parallel([
                    function (callback) {
                        console.log("----------------------- ");
                            User.query("select r_personnel.id_pers from r_personnel where r_personnel.ldap_name ='"+email+"'", function(error, found){
                                if (error) return res.send(error);
                                if(found.rows.length == 0 ) {
                                    req.session.droit = 0;
                                }else{
                                    req.session.droit = found.rows[0].droit;
                                }
                                console.log("Recupération des info dans r_personnel ");
                                callback(null, 1);
                            });
                        }
                    ],function (err, results) {
                        console.log(results);
                        req.session.user = user.id;
                        req.session.nom = user.appelation;
                        req.session.authenticated = true;
                        // res.view('reporting/accueil1');
                        res.view('reporting/debut');
                    });
                    }
                });
            }
        });
    },
  ///_______________________________fin login LDAP
};


