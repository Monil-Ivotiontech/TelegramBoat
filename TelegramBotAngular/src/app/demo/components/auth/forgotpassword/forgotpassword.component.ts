import { Component, OnInit } from '@angular/core';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { MessageService } from 'primeng/api';
import { Subscription } from 'rxjs';
import { EndPoints, ApiMethod } from 'src/app/core/const';
import { HttpService } from 'src/app/core/service/http.service';
import { StorageService } from 'src/app/core/service/storage.service';
import { UtilService } from 'src/app/core/service/util.service';

@Component({
    templateUrl: './forgotpassword.component.html',
    providers: [MessageService]
})
export class ForgotPasswordComponent implements OnInit {
    // global variables 
    ProgressSpinnerDlg: boolean = false;
    sub = new Subscription();

    // forgot password variables
    forgot_pass_form!: FormGroup;
    img_url: any;

    constructor(
        private router: Router,
        private formBuilder: FormBuilder,
        private storageService: StorageService,
        private httpService: HttpService,
        private utilService: UtilService,
        private messageService: MessageService) { }

    ngOnInit(): void {
        this.img_url = 'assets/images/vlogo-red.png';
        this.forgot_pass_form = this.formBuilder.group({
            email: ['', [Validators.required, Validators.email, Validators.pattern('^[a-z0-9._%+-]+@[a-z0-9.-]+\\.[a-z]{2,4}$')]]
        })
    }

    // **************** forgot password methods ****************
    forgot_pass() {
        if (this.forgot_pass_form.invalid) {
            this.messageService.add({ severity: 'error', summary: 'Error Message', detail: 'Please enter valid email' });
            return;
        }
        else if (this.forgot_pass_form.controls.email.value == '') {
            this.messageService.add({ severity: 'error', summary: 'Error Message', detail: 'Please enter email' });
            return;
        }
        else {
            this.ProgressSpinnerDlg = true;
            let browserInfo = this.utilService.getBrowserInfo();
            this.utilService.getIpAddress().subscribe((res: any) => {
                let ipAddress = res.ip;
                this.utilService.getGEOLocation(ipAddress).subscribe((geores: any) => {
                    let email = this.forgot_pass_form.controls.email.value;
                    let city = geores['city'];
                    const details = {
                        "email": email,
                        "location": city,
                        "ipadd": ipAddress,
                        "browsername": browserInfo['browserName'],
                        "browserver": browserInfo['appName'],
                        "useros": browserInfo['userAgent']
                    }
                    this.httpService.frontendRequestCall(EndPoints.admin_forgot_password, ApiMethod.POST, details)
                        .subscribe(response => {
                            response = response || {};
                            let message = response.message || '';
                            let status = response.status || 'false';
                            if (status == "true") {
                                this.ProgressSpinnerDlg = false;
                                this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
                                this.router.navigate(['/auth/newpassword']);
                            } else {
                                this.ProgressSpinnerDlg = false;
                                this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
                            }
                        },
                            error => {
                                this.ProgressSpinnerDlg = false;
                                this.messageService.add({ severity: 'error', summary: 'Error', detail: error.message });
                            });

                });
            })
        }

    }

}
