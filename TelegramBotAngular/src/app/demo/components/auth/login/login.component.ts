import { Component, OnInit } from '@angular/core';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { MessageService } from 'primeng/api';
import { Subscription } from 'rxjs';
import { EndPoints, ApiMethod } from 'src/app/core/const';
import { HttpService } from 'src/app/core/service/http.service';
import { StorageService } from 'src/app/core/service/storage.service';
import { HttpClient } from '@angular/common/http';


@Component({
    selector: 'app-login',
    templateUrl: './login.component.html',
    styles: [`
        i {
            opacity: 0.6;
            transition-duration: .12s;
            
            &:hover {
                opacity: 1;
            }
        }
    `],
    providers: [MessageService]
})
export class LoginComponent implements OnInit {
    // global variables 
    ProgressSpinnerDlg: boolean = false;
    sub = new Subscription();

    // login variables
    loginForm!: FormGroup;
    login_request_resp: any;
    login_creds: Array<any> = [];
    img_url: any;
    ipAddress: any;

    constructor(
        private router: Router,
        private formBuilder: FormBuilder,
        private storageService: StorageService,
        private httpService: HttpService,
        private messageService: MessageService,
        private http: HttpClient
    ) {
    }

    ngOnInit(): void {
        this.img_url = 'assets/images/vlogo-red.png'
        this.http.get<{ ip: string }>('https://jsonip.com')
            .subscribe(data => {
                console.log('th data', data);
                this.ipAddress = data
            })
        const token = this.storageService.getToken() || '';
        if (token != '') {
            this.router.navigate(['configuration/platform']).then(() => { });
        }
        else {
            this.loginForm = this.formBuilder.group({
                useremail: ['', Validators.required],
                password: ['', Validators.required]
            })
        }
    }

    forgot_password() {
        this.router.navigate(['/auth/forgotpassword']);
    }

    login_admin_request() {
        if (this.loginForm.invalid) {
            this.messageService.add({ severity: 'error', summary: 'Error Message', detail: 'Invalid Form Credentials!' });
            return;
        }
        else {
            this.ProgressSpinnerDlg = true;
            let username = this.loginForm.controls.useremail.value;
            let pass = this.loginForm.controls.password.value;
            const credentials = { email: username, password: pass };
            this.httpService.frontendRequestCall(EndPoints.login_admin_request, ApiMethod.POST, credentials)
                .subscribe(response => {
                    response = response || {};
                    let message = response.message || '';
                    let status = response.status || 'false';
                    if (status == 'true') {
                        this.login_request_resp = response.auth_code;
                        this.ProgressSpinnerDlg = false;
                        this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
                        this.login_creds = [{ "email": username, "password": pass, "auth_code": this.login_request_resp }];
                        this.storageService.setLocalObject('logincreds', this.login_creds);
                        this.router.navigate(['auth/verification']);
                    }
                    else {
                        this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
                        this.ProgressSpinnerDlg = false;
                    }
                },
                    error => {
                        this.ProgressSpinnerDlg = false;
                        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong!' });
                    });
        }
    }
}
