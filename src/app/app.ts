import { Component, OnInit, signal } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { FormsModule } from '@angular/forms';

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
  imports: [FormsModule],
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

  // Formulario de pedidos para demostracion en vivo
  nuevoPedido: Pedido = {
    cliente: '',
    producto: '',
    cantidad: 1,
    estado: 'CREADO'
  };

  creandoPedido = signal(false);

  crearNuevoPedido(): void {

    if (
      !this.nuevoPedido.cliente.trim() ||
      !this.nuevoPedido.producto.trim() ||
      !Number.isInteger(Number(this.nuevoPedido.cantidad)) ||
      Number(this.nuevoPedido.cantidad) < 1
    ) {
      this.mensaje.set('Completa correctamente los datos del pedido.');
      return;
    }

    this.creandoPedido.set(true);
    this.mensaje.set('Creando pedido...');

    this.pedidoService.crearPedido({ ...this.nuevoPedido })
      .subscribe({
        next: (pedido) => {
          this.creandoPedido.set(false);
          this.mensaje.set(
            `Pedido #${pedido.id} creado correctamente. Evento enviado al backend.`
          );

          this.nuevoPedido = {
            cliente: '',
            producto: '',
            cantidad: 1,
            estado: 'CREADO'
          };

          this.cargarPedidos();
        },
        error: (error) => {
          this.creandoPedido.set(false);
          this.mensaje.set(
            `Error al crear pedido. HTTP ${error.status}`
          );
          console.error('Error creando pedido:', error);
        }
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

  // Prueba de autorizaciÃƒÆ’Ã‚Â³n:
  // intenta crear un pedido usando un token que solo tiene scope "read".
  // El backend exige "write" para POST, por lo que esperamos HTTP 403.
  probar403(): void {

    const pedidoPrueba: Pedido = {
      cliente: 'Prueba 403',
      producto: 'Sin permiso write',
      cantidad: 1,
      estado: 'Prueba'
    };

    this.pedidoService.crearPedido(pedidoPrueba)
      .subscribe({

        next: (respuesta) => {

          console.log(
            'POST permitido:',
            respuesta
          );

          alert(
            'La peticiÃƒÆ’Ã‚Â³n fue permitida. No se obtuvo 403.'
          );
        },

        error: (error) => {

          console.error(
            'PRUEBA 403:',
            error
          );

          alert(
            `Prueba de autorizaciÃƒÆ’Ã‚Â³n: HTTP ${error.status} ${error.statusText}`
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

  probarAdminRabbit(): void {
    this.mensaje.set('Probando administracion de RabbitMQ...');

    this.http.post(
      'http://localhost:8083/api/admin/colas/prueba-presentacion-360',
      null,
      { responseType: 'text' }
    ).subscribe({
      next: (respuesta) => {
        this.mensaje.set(respuesta);
        alert('RabbitMQ Admin: ' + respuesta);
      },
      error: (error) => {
        this.mensaje.set('Error Admin HTTP ' + error.status);
        console.error('Error en Admin:', error);
        alert('Error Admin HTTP ' + error.status);
      }
    });
  }
}