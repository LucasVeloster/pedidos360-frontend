import { Component, OnInit, signal } from '@angular/core';
import { HttpClient } from '@angular/common/http';

import {
  MsalService,
  MsalBroadcastService
} from '@azure/msal-angular';

import {
  AuthenticationResult,
  InteractionStatus,
  SilentRequest
} from '@azure/msal-browser';

import { filter } from 'rxjs/operators';
import { PedidoService, Pedido } from './pedido.service';

@Component({
  selector: 'app-root',
  imports: [],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App implements OnInit {

  isLoggedIn = signal(false);
  username = signal('');
  pedidos = signal<Pedido[]>([]);
  mensaje = signal('');

  // Claims obtenidos desde el Access Token
  scopeToken = signal('');
  rolesToken = signal<string[]>([]);

  // Estado del segundo microservicio
  estadoNotificaciones = signal('');

  private notificacionesUrl =
    'https://7o73gqbh0b.execute-api.us-east-1.amazonaws.com/api/notificaciones/estado';

  constructor(
    private msalService: MsalService,
    private msalBroadcastService: MsalBroadcastService,
    private pedidoService: PedidoService,
    private http: HttpClient
  ) {}

  ngOnInit(): void {

    this.msalService.handleRedirectObservable().subscribe({
      next: (result: AuthenticationResult | null) => {

        console.log('Resultado MSAL:', result);

        if (result?.account) {
          this.msalService.instance.setActiveAccount(result.account);
        }

        this.checkAccount();
      },

      error: (error) => {
        console.error('Error MSAL:', error);
      }
    });

    this.msalBroadcastService.inProgress$
      .pipe(
        filter(
          (status: InteractionStatus) =>
            status === InteractionStatus.None
        )
      )
      .subscribe(() => {

        console.log(
          'Cuentas encontradas:',
          this.msalService.instance.getAllAccounts()
        );

        this.checkAccount();
      });
  }

  login(): void {
    this.msalService.loginRedirect();
  }

  logout(): void {
    this.msalService.logoutRedirect({
      postLogoutRedirectUri: 'http://localhost:4200/'
    });
  }

  cargarPedidos(): void {

    this.mensaje.set('Cargando pedidos...');

    this.pedidoService.obtenerPedidos().subscribe({
      next: (pedidos) => {

        this.pedidos.set(pedidos);

        this.mensaje.set(
          'Pedidos cargados correctamente.'
        );

        console.log(
          'Pedidos recibidos:',
          pedidos
        );
      },

      error: (error) => {

        console.error(
          'Error al obtener pedidos:',
          error
        );

        this.mensaje.set(
          'No se pudieron obtener los pedidos.'
        );
      }
    });
  }

  consultarNotificaciones(): void {

    this.estadoNotificaciones.set(
      'Consultando microservicio...'
    );

    this.http.get<{
      servicio: string;
      estado: string;
    }>(this.notificacionesUrl)
      .subscribe({

        next: (respuesta) => {

          this.estadoNotificaciones.set(
            `${respuesta.servicio}: ${respuesta.estado}`
          );

          console.log(
            'Respuesta microservicio notificaciones:',
            respuesta
          );
        },

        error: (error) => {

          console.error(
            'Error microservicio notificaciones:',
            error
          );

          this.estadoNotificaciones.set(
            'Error al consultar notificaciones'
          );
        }
      });
  }

  private checkAccount(): void {

    const accounts =
      this.msalService.instance.getAllAccounts();

    if (accounts.length > 0) {

      const account = accounts[0];

      this.msalService.instance.setActiveAccount(
        account
      );

      this.isLoggedIn.set(true);

      this.username.set(
        account.name ?? account.username
      );

      console.log(
        'USUARIO LOGUEADO:',
        this.username()
      );

      const request: SilentRequest = {
        account,
        scopes: [
          'api://ef080cc3-9d30-4a1a-89b5-c5236f9bd00d/read'
        ]
      };

      this.msalService
        .acquireTokenSilent(request)
        .subscribe({

          next: (
            result: AuthenticationResult
          ) => {

            const tokenParts =
              result.accessToken.split('.');

            const payload = JSON.parse(
              atob(tokenParts[1])
            );

            console.log(
              'CLAIMS DEL ACCESS TOKEN:',
              payload
            );

            this.scopeToken.set(
              payload.scp ?? ''
            );

            if (Array.isArray(payload.roles)) {

              this.rolesToken.set(
                payload.roles
              );

            } else {

              this.rolesToken.set([]);
            }

            console.log(
              'SCOPES:',
              this.scopeToken()
            );

            console.log(
              'ROLES:',
              this.rolesToken()
            );
          },

          error: (error) => {

            console.error(
              'Error obteniendo Access Token:',
              error
            );
          }
        });

    } else {

      this.isLoggedIn.set(false);
      this.username.set('');

      this.scopeToken.set('');
      this.rolesToken.set([]);
      this.estadoNotificaciones.set('');

      console.log('NO HAY USUARIO');
    }
  }
}