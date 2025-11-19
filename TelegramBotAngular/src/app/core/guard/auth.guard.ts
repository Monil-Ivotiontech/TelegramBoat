import { Injectable } from '@angular/core';
import { CanActivate, Router, UrlTree } from '@angular/router';
import { Observable } from 'rxjs';
import { StorageService } from '../service/storage.service';

@Injectable({
    providedIn: 'root'
})
export class AuthGuard implements CanActivate {
    constructor(private storageService: StorageService,
        private router: Router) { }
    canActivate(): Observable<boolean> | Promise<boolean> | boolean | UrlTree {
        let token = this.storageService.getToken() || '';
        if (token != '') {
            return true;
        }
        else {
            return this.router.parseUrl("/auth/login");
        }
    }

}
