# Passport-SharePoint

[Passport](http://passportjs.org/) strategy for authenticating with [SharePoint 2013](http://sharepoint.microsoft.com/.com/) OnPremise and O365 using the OAuth 2.0 API.

This module lets you authenticate using SharePoint 2013 OnPremise or O365 in your Node.js applications.
By plugging into Passport, SharePoint authentication can be easily and unobtrusively integrated into any application or framework that supports [Connect](http://www.senchalabs.org/connect/)-style middleware, including [Express](http://expressjs.com/).

## Installation

    $ npm install passport-sharepoint

## Usage

#### Configure Strategy

The SharePoint authentication strategy authenticates users using a SharePoint 2013 OnPremise or O365
account using OAuth 2.0.  The strategy requires a `verify` callback, which
accepts these credentials and calls `done` providing a user, as well as
`options` specifying a app ID, app secret, and callback URL.

    passport.use(new SharePointStrategy({
        appId: SHAREPOINT_APP_ID ,
        appSecret: SHAREPOINT_APP_SECRET ,
        callbackURL: "http://localhost:3000/auth/sharepoint/callback"
      },
      function(accessToken, refreshToken, profile, done) {
        User.findOrCreate({ userID: profile.id }, function (err, user) {
          return done(err, user);
        });
      }
    ));
    
#### Configure SharePoint AppPart

On the SharePoint side you need a provider hosted AppPart that talks to you Node.JS server and you must register your Node.JS server as a app.
The AppPart you can simply create via the VS2012 AppPart wizard.
These AppPart must define the `StandardTokens` as the url parameter so that the strategy can work.

    <Content Type="html" Src="https://nodeserver:3000/auth/sharepoint?{StandardTokens}" />

The Node.JS Server you can register as an app at
`https://sharepoint/_layouts/15/AppRegNew.aspx`
The app id and app secret you specify here is used in our strategy.

#### App Permission Request

To load the user profile from the current user automatically, you should add the following permission request to you app manifest or register manually the permission via https://your-tenant.sharepoint.com/_layouts/15/appinv.aspx

    <AppPermissionRequests AllowAppOnlyPolicy="true" >
      <AppPermissionRequest Scope="http://sharepoint/social/tenant" Right="Read" />
    </AppPermissionRequests>

#### Authenticate Requests

Use `passport.authenticate()`, specifying the `'sharepoint'` strategy, to
authenticate requests.

For example, as route middleware in an [Express](http://expressjs.com/)
application:

    app.get('/auth/sharepoint',
      passport.authenticate('sharepoint'),
      function(req, res){
        // The request will be redirected to SharePoint for authentication, so
        // this function will not be called.
      });

    app.get('/auth/sharepoint/callback', 
      passport.authenticate('sharepoint', { failureRedirect: '/login' }),
      function(req, res) {
        // Successful authentication, redirect home.
        res.redirect('/');
      });
      
## Credits

  - [QuePort](https://github.com/QuePort)
  - [Thomas Herbst](https://github.com/macrauder)

## License

(The MIT License)

Copyright (c) 2013 Thomas Herbst / QuePort

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.