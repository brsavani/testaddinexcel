// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the root of the repo.

/* 
    This file configures auth for the add-in. 
*/

using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin;
using System.Configuration;
using Microsoft.Owin.Security.OpenIdConnect;
using System.Threading.Tasks;
using System.Globalization;
using Microsoft.Owin.Security;
using Owin;

namespace Office_Add_in_ASPNET_SSO_WebAPI
{
    public partial class Startup
    {
        private static string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static string appKey = ConfigurationManager.AppSettings["ida:AppKey"];
        private static string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private static string tenant = ConfigurationManager.AppSettings["ida:Tenant"];
        private static string audience = ConfigurationManager.AppSettings["ida:Audience"];
        private static string postLogoutRedirectUri = ConfigurationManager.AppSettings["ida:PostLogoutRedirectUri"];
        private static string authority = string.Format(CultureInfo.InvariantCulture, aadInstance, tenant);
        
        public void ConfigureAuth(IAppBuilder app)
        {
                       
            
            app.SetDefaultSignInAsAuthenticationType(CookieAuthenticationDefaults.AuthenticationType);
            app.UseCookieAuthentication(new CookieAuthenticationOptions());
            app.UseOpenIdConnectAuthentication(new OpenIdConnectAuthenticationOptions
            {
                ClientId = clientId,
                Authority = authority,
                PostLogoutRedirectUri = postLogoutRedirectUri,
                Notifications = new OpenIdConnectAuthenticationNotifications
                {
                    AuthenticationFailed = context =>
                    {
                        context.HandleResponse();
                        context.Response.Redirect("Home/Error");
                        return Task.FromResult(0);
                    }
                }
            });
            
            
            
            
            
            //var tvps = new TokenValidationParameters
            //{
            //    // Set the strings to validate against. (Scopes, which should be 
            //   // simply "access_as_user" in this sample, is validated inside the controller.)
            //    ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            //    ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],

                // Save the raw token recieved from the Office host, so it can be 
                // used in the "on behalf of" flow.
            //    SaveSigninToken = true
            };

            // The more familiar UseWindowsAzureActiveDirectoryBearerAuthentication does not work
            // with the Azure AD V2 endpoint, so use UseOAuthBearerAuthentication instead.
           // app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
            //{
          //      AccessTokenFormat = new JwtFormat(tvps,
//
                    // Specify the discovery endpoint, also called the "metadata address".
           //         new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
           // });
        }
    }
}
