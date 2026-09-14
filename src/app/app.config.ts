import { ApplicationConfig } from '@angular/core';
import { provideRouter } from '@angular/router';
import { provideHttpClient, withInterceptorsFromDi } from '@angular/common/http';
import { HTTP_INTERCEPTORS } from '@angular/common/http';

import {
  MsalGuard,
  MsalInterceptor,
  MsalInterceptorConfiguration,
  MsalGuardConfiguration,
  MsalService,
  MsalBroadcastService,
  MSAL_INSTANCE,
  MSAL_GUARD_CONFIG,
  MSAL_INTERCEPTOR_CONFIG
} from '@azure/msal-angular';

import {
  PublicClientApplication,
  IPublicClientApplication,
  InteractionType,
  BrowserCacheLocation
} from '@azure/msal-browser';

import { routes } from './app.routes';

const READ_SCOPE =
  'api://ef080cc3-9d30-4a1a-89b5-c5236f9bd00d/read';

export function MSALInstanceFactory(): IPublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: '7c78da7c-5f00-4191-83c2-2f59e990246e',
      authority:
        'https://login.microsoftonline.com/fd969d4c-5256-48f4-899c-b33c192dc051',
      redirectUri: 'http://localhost:4200/',
      postLogoutRedirectUri: 'http://localhost:4200/'
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage
    }
  });
}

export function MSALGuardConfigFactory(): MsalGuardConfiguration {
  return {
    interactionType: InteractionType.Redirect
  };
}

export function MSALInterceptorConfigFactory(): MsalInterceptorConfiguration {
  const protectedResourceMap = new Map<string, Array<string>>();

  protectedResourceMap.set(
    'https://7o73gqbh0b.execute-api.us-east-1.amazonaws.com/api/pedidos',
    [READ_SCOPE]
  );

  protectedResourceMap.set(
    'https://7o73gqbh0b.execute-api.us-east-1.amazonaws.com/api/notificaciones/estado',
    [READ_SCOPE]
  );

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap
  };
}

export const appConfig: ApplicationConfig = {
  providers: [
    provideRouter(routes),

    provideHttpClient(
      withInterceptorsFromDi()
    ),

    {
      provide: HTTP_INTERCEPTORS,
      useClass: MsalInterceptor,
      multi: true
    },

    {
      provide: MSAL_INSTANCE,
      useFactory: MSALInstanceFactory
    },

    {
      provide: MSAL_GUARD_CONFIG,
      useFactory: MSALGuardConfigFactory
    },

    {
      provide: MSAL_INTERCEPTOR_CONFIG,
      useFactory: MSALInterceptorConfigFactory
    },

    MsalService,
    MsalBroadcastService,
    MsalGuard
  ]
};