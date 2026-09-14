import { Routes } from '@angular/router';
import { MsalGuard } from '@azure/msal-angular';
import { App } from './app';

export const routes: Routes = [
  {
    path: '',
    component: App,
    canActivate: [MsalGuard]
  }
];