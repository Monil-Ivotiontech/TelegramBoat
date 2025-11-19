import { Injectable } from '@angular/core';
import {
  HttpRequest,
  HttpHandler,
  HttpEvent,
  HttpInterceptor
} from '@angular/common/http';
import { Observable } from 'rxjs';
import { StorageService } from '../service/storage.service';

@Injectable()
export class TokenInterceptor implements HttpInterceptor {

  constructor(private storage: StorageService) { }

  intercept(request: HttpRequest<unknown>, next: HttpHandler): Observable<HttpEvent<unknown>> {
    let token = this.storage.getToken() || '';
    let client_key = this.storage.getClientKey() || '';
    console.log('token', token);
    console.log('client_key', client_key);
    request = request.clone({
      headers: request.headers.set('Content-Type', 'application/json')
    });
    if (token != '') {
      request = request.clone({
        headers: request.headers.set('authtoken', token)
      });
    }
    if (client_key != '') {
      request = request.clone({
        headers: request.headers.set('clientkey', client_key)
      })
    }
    return next.handle(request);
  }
}
