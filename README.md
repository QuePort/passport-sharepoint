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

    passport.use(new PayPalStrategy({
        clientID: SHAREPOINT_APP_ID ,
    	clientSecret: SHAREPOINT_APP_SECRET ,
        callbackURL: "http://localhost:3000/auth/sharepoint/callback"
      },
      function(accessToken, refreshToken, profile, done) {
        User.findOrCreate({ userID: profile.id }, function (err, user) {
          return done(err, user);
        });
      }
    ));

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