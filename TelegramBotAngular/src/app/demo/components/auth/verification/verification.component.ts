import { Component, OnInit, ViewChild } from '@angular/core';
import { Router } from '@angular/router';
import { Config } from 'ng-otp-input/lib/models/config';
import { MessageService } from 'primeng/api';
import { Observable, Subscription, throwError } from 'rxjs';
import { EndPoints, ApiMethod } from 'src/app/core/const';
import { HttpService } from 'src/app/core/service/http.service';
import { StorageService } from 'src/app/core/service/storage.service';
import { UtilService } from 'src/app/core/service/util.service';
import { LoginComponent } from '../login/login.component';
import { GMapModule } from 'primeng/gmap';

@Component({
    templateUrl: './verification.component.html',
    styles: [`
        :host ::ng-deep .otp-input {
            font-size: 20px !important;
            
            &:focus-visible {
                outline-color: #c5c5c5 !important;
            }
        }
    `],
    providers: [LoginComponent, MessageService]
})
export class VerificationComponent implements OnInit {
    // global variables 
    ProgressSpinnerDlg: boolean = false;
    sub = new Subscription();
    user_details: any;
    otp!: string;
    showOtpComponent = true;
    @ViewChild('ngOtpInput', { static: false }) ngOtpInput: any;
    config: Config = {
        allowNumbersOnly: false,
        length: 8,
        isPasswordInput: false,
        disableAutoFocus: false,
        placeholder: "",
        inputStyles: {
            width: '11%',
            color: "#495057",
        },
    };


    // Verification variables
    login_credentials: Array<any> = [];
    email_address!: any;
    infoWindow: any;
    img_url: any;

    constructor(public loginPage: LoginComponent,
        private router: Router,
        private storageService: StorageService,
        private httpService: HttpService,
        private utilService: UtilService,
        private messageService: MessageService,) {
    }

    ngOnInit(): void {
        this.img_url = 'assets/images/vlogo-red.png';
        const token = this.storageService.getToken() || '';
        if (token != '') {
            this.router.navigate(['configuration/platform']).then(() => { });
        }

        else {
            this.login_credentials = this.storageService.getLocalObject("logincreds");
            this.email_address = this.login_credentials[0].email.split('@')[0].substring(0, 2) + "**@" + this.login_credentials[0].email.split('@').pop();
        }
        this.login_admin_cookies();
    }

    onOtpChange(otp: any) {
        this.otp = otp;
        if (otp.length == 4) {
            this.validateOtp();
        }
    }
    setVal(val: any) {
        this.ngOtpInput.setValue(val);
    }

    onConfigChange() {
        this.showOtpComponent = false;
        this.otp = '';
        setTimeout(() => {
            this.showOtpComponent = true;
        }, 0);
    }
    validateOtp() {
        if (this.otp != this.login_credentials[0].auth_code) {
            this.messageService.add({ severity: 'error', summary: 'Error Message', detail: 'Invalid Auth Code!' });
            return;
        }
        else {
            this.login_admin_success();
        }
    }

    login_admin_success() {
        this.ProgressSpinnerDlg = true;
        let username = this.login_credentials[0].email;
        let pass = this.login_credentials[0].password;
        const credentials = { email: username, password: pass };
        this.sub = this.httpService.frontendRequestCall(EndPoints.login_admin_success, ApiMethod.POST, credentials)
            .subscribe(response => {
                response = response || {};
                let message = response.message || '';
                let status = response.status || 'false';
                if (status == 'true') {
                    this.ProgressSpinnerDlg = false;
                    this.otp = '';
                    this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
                    this.user_details = response.user_detail || {};
                    this.storageService.removeLocalObject('logincreds');
                    this.storageService.setLocalObject('userdetails', this.user_details);
                    this.storageService.saveToken(this.user_details.auth_token);
                    this.storageService.saveClientKey(this.user_details.client_key);
                    this.router.navigate(['configuration/platform']);
                    this.login_admin_cookies();
                }
                else {
                    this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
                    this.ProgressSpinnerDlg = false;
                }
            },
                error => {
                    this.ProgressSpinnerDlg = false;
                    this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong!' });
                }
            );
    }

    login_admin_cookies() {
        let browserInfo = this.utilService.getBrowserInfo();
        this.utilService.getIpAddress().subscribe((res: any) => {
            let ipAddress = res.ip;
            this.utilService.getGEOLocation(ipAddress).subscribe((geores: any) => {
                let city = geores['city'];
                let country = geores['country_name'];
                let state = geores['state_prov'];
                const details = {
                    "userid": this.user_details.userid,
                    "email": this.user_details.email,
                    "location": city,
                    "ipadd": ipAddress,
                    "browsername": browserInfo['browserName'],
                    "browserver": browserInfo['appName'],
                    "useros": browserInfo['userAgent'],
                    "time": new Date().toString().split(' ')[4]
                }
                this.httpService.frontendRequestCall(EndPoints.login_admin_cookies, ApiMethod.POST, details)
                    .subscribe(response => {
                        response = response || {};
                        let message = response.message || '';
                        let status = response.status || 'false';
                        if (status == "true") {
                        } else {
                            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
                        }
                    },
                        error => {
                            this.messageService.add({ severity: 'error', summary: 'Error', detail: error.message });
                        });
            });
        });
    }
}

