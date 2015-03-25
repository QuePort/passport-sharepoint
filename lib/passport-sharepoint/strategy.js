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


var SP_AUTH_PREFIX = '/_layouts/15/OAuthAuthorize.aspx';
var SP_REDIRECT_PREFIX = '/_layouts/15/appredirect.aspx';

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
    
  this._appId = options.appId;
  this._appSecret = options.appSecret;
  this._callbackURL = options.callbackURL;
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
 * Authenticate request by delegating to the SharePoint OAuth 2.0 provider.
 *
 * @param {Object} req
 * @api protected
 */
Strategy.prototype.authenticate = function(req, options) {
  options = options || {};
  var self = this;
  
  if (req != undefined && req.query && req.query.error) {
    // TODO: Error information pertaining to OAuth 2.0 flows is encoded in the
    //       query parameters, and should be propagated to the application.
    return this.fail();
  }
  
  var callbackURL = options.callbackURL || this._callbackURL;
  if (callbackURL && req != undefined) {
    var parsed = url.parse(callbackURL);
    if (!parsed.protocol) {
      // The callback URL is relative, resolve a fully qualified URL from the
      // URL of the originating request.
      callbackURL = url.resolve(utils.originalURL(req), callbackURL);
    }
  }
  
  var spLanguage = options.spLanguage;
  
  var spAppToken = undefined;
  var spSiteUrl = undefined;
      
  // load token from request
  if (req != undefined && req.body && req.body.SPAppToken)
  	spAppToken = req.body.SPAppToken;
  if(req != undefined && req.body && req.body.SPSiteUrl)
  	spSiteUrl = req.body.SPSiteUrl;
  
  if(spSiteUrl == undefined && req != undefined && req.query && req.query.SPHostURL)
  	spSiteUrl = req.query.SPHostURL;
  	
  //fallback to optional values
  if (spAppToken == undefined)
      spAppToken = options.spAppToken;
  if(spSiteUrl == undefined)
      spSiteUrl = options.spSiteUrl;
  
  // you can pass the appId and Secret in every round
  if (options.appId)
    this._appId = options.appId;
  if (options.appSecret)
    this._appSecret = options.appSecret;
    
  if (!this._appId) throw new Error('SharePointStrategy requires a appId.');
  if (!this._appSecret) throw new Error('SharePointStrategy requires a appSecret.');

  var authorizationURL = spSiteUrl + SP_AUTH_PREFIX;
  var appRedirectURL = spSiteUrl + SP_REDIRECT_PREFIX;
  
  // check if there is a app token present
  if (spAppToken && spSiteUrl) {
    var token = eval(jwt.decode(spAppToken, '', true));
    var splitApptxSender = token.appctxsender.split("@");
    var sharepointServer = url.parse(spSiteUrl)
    var resource = splitApptxSender[0]+"/"+sharepointServer.host+"@"+splitApptxSender[1];
    var spappId = this._appId+"@"+splitApptxSender[1];
    var appctx = JSON.parse(token.appctx);
    var tokenServiceUri = url.parse(appctx.SecurityTokenServiceUri);
    var tokenURL = tokenServiceUri.protocol+'//'+tokenServiceUri.host+'/'+splitApptxSender[1]+tokenServiceUri.path;

    this._oauth2 = new OAuth2(spappId,  this._appSecret, '', authorizationURL, tokenURL);
    this._oauth2.getOAuthAccessToken(
        token.refreshtoken,
        {grant_type: 'refresh_token', refresh_token: token.refreshtoken, resource: resource},
        function (err, accessToken, refreshToken, params) {
            if (err) { return self.error(new InternalOAuthError('failed to obtain access token', err)); }
            if (!refreshToken)
              refreshToken = spAppToken;
              
            self._loadUserProfile(accessToken, spSiteUrl, function(err, profile) {
              if (err) { return self.error(err); };
            
              function verified(err, user, info) {
                if (err) { return self.error(err); }
                if (!user) { return self.fail(info); }
                self.success(user, info);
              }
              
              profile.cacheKey = appctx.CacheKey;
              profile.language = spLanguage;
              
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
  } else if (req != undefined && req.query && req.query.code && authorizationURL && req.query.tokenURL) {
    this._oauth2 = new OAuth2(this._appId, this._appSecret, '', authorizationURL, req.query.tokenURL);
    var code = req.query.code;
    
    // NOTE: The module oauth (0.9.5), which is a dependency, automatically adds
    //       a 'type=web_server' parameter to the percent-encoded data sent in
    //       the body of the access token request.  This appears to be an
    //       artifact from an earlier draft of OAuth 2.0 (draft 22, as of the
    //       time of this writing).  This parameter is not necessary, but its
    //       presence does not appear to cause any issues.
    this._oauth2.getOAuthAccessToken(code, { grant_type: 'authorization_code', resource: req.query.resource, redirect_uri: callbackURL },
      function(err, accessToken, refreshToken, params) {
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
      }
    );
  } else if (appRedirectURL) {
    this._oauth2 = new OAuth2(this._appId,  this._appSecret, '', appRedirectURL, '');
    var params = this.authorizationParams(options);
    params['response_type'] = 'code';
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
  } else
    return self.error(new InternalOAuthError('failed to obtain access token')); 
}

/**
 * Retrieve user profile from SharePoint.
 *
 * @param {String} accessToken
 * @param {Function} done
 * @api protected
 */
Strategy.prototype.userProfile = function(accessToken, spSiteUrl, done) {
  if (spSiteUrl)
    sharepointServer = url.parse(spSiteUrl)
  else
    return done(null, {});
  if (sharepointServer.path.length > 1)
    sharepointServer.path = sharepointServer.path + '/';
  
  var headers = {
    'Accept': 'application/json;odata=verbose',
    'Authorization' : 'Bearer ' + accessToken
  };
  var options = {
    host : sharepointServer.hostname, 
    port : sharepointServer.port || 443,
    path : sharepointServer.path + '_api/web/currentuser',
    method : 'GET',
    headers : headers,
    agent: false,
    secureOptions: require('constants').SSL_OP_NO_TLSv1_2
  };
  
  var req = https.get(options, function(res) {
    res.setEncoding('utf8');
    var userData = '';
    
    res.on('data', function(data) {
        userData += data;
    });
    
    res.on('end', function() {
      var json = JSON.parse(userData);
      if (json.d) {
        var profile = { provider: 'sharepoint' };
        profile.id = json.d.Id;
        profile.username = json.d.LoginName;
        profile.displayName = json.d.Title;
        profile.emails = [{ value: json.d.Email }];
        siteUrl = url.parse(spSiteUrl);
        profile.host = siteUrl.protocol + '//' + siteUrl.host;
        profile.site = siteUrl.pathname;
        if (profile.site.length > 1)
          profile.site = profile.site + '/';
        profile._raw = json;
        
        done(null, profile);
      } else if (json.error) {
        return done ('Authentication failed: ' + json.error.code + ' at ' + options.host + ':' + options.port + options.path, null);
      } else {
        return done('Authentication failed: Unknown exception at' 
          + options.host + ':' + options.port + options.path , null);
      }
    });
  }).on('error', function(e) {
    return done('Authentication failed: ' + e + ' at '
      + options.host + ':' + options.port + options.path , null);
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
Strategy.prototype._loadUserProfile = function(accessToken, spSiteUrl, done) {
  var self = this;
  
  function loadIt() {
    return self.userProfile(accessToken, spSiteUrl, done);
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

