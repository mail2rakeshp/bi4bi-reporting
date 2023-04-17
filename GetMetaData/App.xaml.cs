﻿using Microsoft.Identity.Client;
using System;
using System.Windows;

namespace GetMetaData
{
    public partial class App : Application
    {
        public static string InitialStartDirectory;
        static App()
        {

            //SplashScreen splash = new SplashScreen("/Images/logo.png");
            //splash.Show(false);
            //splash.Close(TimeSpan.FromSeconds(10));
            InitialStartDirectory = Environment.CurrentDirectory;
            _clientApp = PublicClientApplicationBuilder.Create(ClientId)
            .WithAuthority($"{Instance}{Tenant}")
            .WithDefaultRedirectUri()
            .Build();
            TokenCacheHelper.EnableSerialization(_clientApp.UserTokenCache);
        }

        // Below are the clientId (Application Id) of your app registration and the tenant information.
        // You have to replace:
        // - the content of ClientID with the Application Id for your app registration
        // - The content of Tenant by the information about the accounts allowed to sign-in in your application:
        //   - For Work or School account in your org, use your tenant ID, or domain
        //   - for any Work or School accounts, use organizations
        //   - for any Work or School accounts, or Microsoft personal account, use common
        //   - for Microsoft Personal account, use consumers
        private static string ClientId = "4a1aa1d5-c567-49d0-ad0b-cd957a47f842";

        // Note: Tenant is important for the quickstart. We'd need to check with Andre/Portal if we
        // want to change to the AadAuthorityAudience.
        private static string Tenant = "common";
        private static string Instance = "https://login.microsoftonline.com/";
        private static IPublicClientApplication _clientApp;

        public static IPublicClientApplication PublicClientApp { get { return _clientApp; } }
    }
}