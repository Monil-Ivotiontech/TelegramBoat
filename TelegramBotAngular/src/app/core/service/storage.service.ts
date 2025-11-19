import { Injectable } from '@angular/core';
import { Router } from '@angular/router';
import { CookieService } from 'ngx-cookie-service';
import { UtilService } from './util.service';

@Injectable({
    providedIn: 'root',
})
export class StorageService {
    constructor(
        private cookies: CookieService,
        private _util: UtilService,
        private router: Router
    ) { }

    saveToken(token: string): void {
        let expiredDate = new Date();
        expiredDate.setDate(expiredDate.getDate() + 1);
        let globleVarible = this._util.getGlobleEnvironments();
        this.cookies.set(
            'venus-token',
            this._util.encrypt(token),
            expiredDate,
            '/',
            globleVarible.hostname
        );
    }

    getToken() {
        return this._util.decrypt(this.cookies.get('venus-token'));
    }
    removeCookies() {
        let globleVarible = this._util.getGlobleEnvironments();
        this.cookies.deleteAll('/', globleVarible.hostname);
    }

    saveClientKey(token: string): void {
        let expiredDate = new Date();
        expiredDate.setDate(expiredDate.getDate() + 1);
        let globleVarible = this._util.getGlobleEnvironments();
        this.cookies.set(
            'venus-clientKey',
            this._util.encrypt(token),
            expiredDate,
            '/',
            globleVarible.hostname
        );
        this.router.navigate(['./authentication/signin']);
    }

    getClientKey() {
        return this._util.decrypt(this.cookies.get('venus-clientKey'));
    }
    removeClientKey() {
        this.cookies.delete('venus-clientKey');
    }

    setLocalObject(key: string, value: any) {
        localStorage.setItem(key, this._util.encrypt(JSON.stringify(value)));
    }

    getLocalObject(key: string): any {
        return JSON.parse(this._util.decrypt(localStorage.getItem(key) || {}));
    }

    removeLocalObject(key: string) {
        localStorage.removeItem(key);
    }
}
