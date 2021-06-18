/**
 * Garantie.js
 *
 * @description :: A model definition represents a database table/collection.
 * @docs        :: https://sailsjs.com/docs/concepts/models-and-orm/models
 */

module.exports = {

  attributes: {

  },
  deleteFromChemin : function (table,callback) {
    var sql = "delete from cheminhtp ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return callback(err); }
      return callback(null, true);
      });
  },
  deleteFromChemin2 : function (table,callback) {
    var sql = "delete from cheminhtp2 ";
    Reportinghtp.getDatastore().sendNativeQuery(sql, function(err, res){
      if (err) { return callback(err); }
      return callback(null, true);
      });
  },
};

