/**
 * Module dependencies.
 */
var passport = require('passport')
  , url = require('url')
  , https = require('https')
  , util = require('util')
  , utils = require('./utils')
  , jwt = require('jwt-simple')
  , OAuth2 = require('oauth').OAuth2
  , InternalOAuthError = require('./internaloautherror');;


/**
 * `Strategy` constructor.
 *
 * @param {Object} options
 * @param {Function} verify
 * @api public
 */
function Strategy(options, verify) {
  options = options || {}
  passport.Strategy.call(this);
  this.name = 'sharepoint';
  this._verify = verify;
  
  if (!options.clientID) throw new Error('SharePointStrategy requires a clientID option');
  if (!options.clientSecret) throw new Error('SharePointStrategy requires a clientSecret option');
  
  this._clientID = options.clientID;
  this._clientSecret = options.clientSecret;
  this._callbackURL = options.callbackURL;
  this._spHostUrl = options.spHostUrl;
  this._scope = options.scope;
  this._scopeSeparator = options.scopeSeparator || ' ';
  this._passReqToCallback = options.passReqToCallback;
  this._skipUserProfile = (options.skipUserProfile === undefined) ? false : options.skipUserProfile;
}

/**
 * Inherit from `passport.Strategy`.
 */
util.inherits(Strategy, passport.Strategy);


/**
 * Authenticate request by delegating to a service provider using OAuth 2.0.
 *
 * @param {Object} req
 * @api protected
 */
Strategy.prototype.authenticate = function(req, options) {
  options = options || {};
  var self = this;
  
  if (req.query && req.query.error) {
    // TODO: Error information pertaining to OAuth 2.0 flows is encoded in the
    //       query parameters, and should be propagated to the application.
    return this.fail();
  }
  
  var callbackURL = options.callbackURL || this._callbackURL;
  if (callbackURL) {
    var parsed = url.parse(callbackURL);
    if (!parsed.protocol) {
      // The callback URL is relative, resolve a fully qualified URL from the
      // URL of the originating request.
      callbackURL = url.resolve(utils.originalURL(req), callbackURL);
    }
  }
  // check if there is a app token present
  if (req.query && req.body.SPAppToken) {
    if (req.query.SPHostUrl)
      this._spHostUrl = req.query.SPHostUrl;
    if (!this._spHostUrl)
      return self.error(new InternalOAuthError('SharePoint host url not detected or specified.'));
      
    token = eval(jwt.decode(req.body.SPAppToken, '', true));
    splitApptxSender = token.appctxsender.split("@");
    sharepointServer = url.parse(this._spHostUrl)
    resource = splitApptxSender[0]+"/"+sharepointServer.host+"@"+splitApptxSender[1];
    spClientID = this._clientID+"@"+splitApptxSender[1];
    appctx = JSON.parse(token.appctx);
    tokenServiceUri = url.parse(appctx.SecurityTokenServiceUri);
    authorizationURL = req.query.SPHostUrl + '/_layouts/15/OAuthAuthorize.aspx';
    tokenURL = tokenServiceUri.protocol+'//'+tokenServiceUri.host+'/'+splitApptxSender[1]+tokenServiceUri.path;
    
    this._oauth2 = new OAuth2(spClientID,  this._clientSecret, '', authorizationURL, tokenURL);
    this._oauth2.getOAuthAccessToken(
        token.refreshtoken,
        {grant_type: 'refresh_token', refresh_token: token.refreshtoken, resource: resource},
        function (err, accessToken, refreshToken) {
            if (err) { return self.error(new InternalOAuthError('failed to obtain access token', err)); }
            
            self._loadUserProfile(accessToken, function(err, profile) {
            if (err) { return self.error(err); };
          
            function verified(err, user, info) {
              if (err) { return self.error(err); }
              if (!user) { return self.fail(info); }
              self.success(user, info);
            }
            
            if (self._passReqToCallback) {
              var arity = self._verify.length;
              if (arity == 6) {
                self._verify(req, accessToken, refreshToken, params, profile, verified);
              } else { // arity == 5
                self._verify(req, accessToken, refreshToken, profile, verified);
              }
            } else {
              var arity = self._verify.length;
              if (arity == 5) {
                self._verify(accessToken, refreshToken, params, profile, verified);
              } else { // arity == 4
                self._verify(accessToken, refreshToken, profile, verified);
              }
            }
          });  
        });
  } else {
    var params = this.authorizationParams(options);
    params['response_type'] = 'code';
    params['IsDlg'] = 1;
    params['redirect_uri'] = callbackURL;
    var scope = options.scope || this._scope;
    if (scope) {
      if (Array.isArray(scope)) { scope = scope.join(this._scopeSeparator); }
      params.scope = scope;
    }
    var state = options.state;
    if (state) { params.state = state; }
    
    var location = this._oauth2.getAuthorizeUrl(params);
    this.redirect(location);
  }
}

/**
 * Retrieve user profile from SharePoint.
 *
 * @param {String} accessToken
 * @param {Function} done
 * @api protected
 */
Strategy.prototype.userProfile = function(accessToken, done) {
  sharepointServer = url.parse(this._spHostUrl)
  var headers = {
    'ACCEPT': 'application/json; odata=verbose',
    'Authorization' : 'Bearer ' + accessToken
  };
  var options = {
    host : sharepointServer.hostname, 
    port : sharepointServer.port,
    path : sharepointServer.path + '_api/web/currentuser',
    method : 'GET',
    headers : headers
  };
  
  var req = https.get(options, function(res) {
    var userData = "";
    res.on('data', function(data) {
        userData += data;
    });
    
    res.on('end', function() {
      var json = JSON.parse(userData);
      var profile = { provider: 'sharepoint' };
      profile.id = json.d.Id;
      profile.username = json.d.LoginName;
      profile.displayName = json.d.Title;
      profile.emails = [{ value: json.d.Email }];
      
      profile._raw = userData;
      profile._json = json;
      
      done(null, profile);
    });
  }).on('error', function(e) {
    return done(null, {});
  });
}

/**
 * Return extra parameters to be included in the authorization request.
 *
 * @param {Object} options
 * @return {Object}
 * @api protected
 */
Strategy.prototype.authorizationParams = function(options) {
  return {};
}

/**
 * Load user profile, contingent upon options.
 *
 * @param {String} accessToken
 * @param {Function} done
 * @api private
 */
Strategy.prototype._loadUserProfile = function(accessToken, done) {
  var self = this;
  
  function loadIt() {
    return self.userProfile(accessToken, done);
  }
  function skipIt() {
    return done(null);
  }
  
  if (typeof this._skipUserProfile == 'function' && this._skipUserProfile.length > 1) {
    // async
    this._skipUserProfile(accessToken, function(err, skip) {
      if (err) { return done(err); }
      if (!skip) { return loadIt(); }
      return skipIt();
    });
  } else {
    var skip = (typeof this._skipUserProfile == 'function') ? this._skipUserProfile() : this._skipUserProfile;
    if (!skip) { return loadIt(); }
    return skipIt();
  }
}


/**
 * Expose `Strategy`.
 */ 
module.exports = Strategy;

