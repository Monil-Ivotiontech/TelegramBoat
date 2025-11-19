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
	templateUrl: './newpassword.component.html',
	providers: [MessageService]
})
export class NewPasswordComponent {
	// global variables 
	ProgressSpinnerDlg: boolean = false;
	sub = new Subscription();

	// reset password variables
	reset_pass_form!: FormGroup;
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

		this.reset_pass_form = this.formBuilder.group({
			temp_password: ['', Validators.required],
			new_password: ['', Validators.required],
			confirm_password: ['', Validators.required]
		})
	}

	// ************** Reset password methods ******************
	reset_pass() {
		if (this.reset_pass_form.invalid) {
			this.messageService.add({ severity: 'error', summary: 'Error Message', detail: 'Please enter valid password' });
			return;
		}
		else if (this.reset_pass_form.controls.new_password.value != this.reset_pass_form.controls.confirm_password.value) {
			this.messageService.add({ severity: 'error', summary: 'Error Message', detail: 'New Password and Confirm Password Should be same!' });
			return;
		}
		else {
			this.ProgressSpinnerDlg = true;
			let temppass = this.reset_pass_form.controls.temp_password.value;
			let newpassword = this.reset_pass_form.controls.new_password.value;
			const credentials = { "temppass": temppass, "newpassword": newpassword };
			this.httpService.frontendRequestCall(EndPoints.admin_reset_password, ApiMethod.POST, credentials)
				.subscribe(response => {
					response = response || {};
					let message = response.message || '';
					let status = response.status || 'false';
					if (status == "true") {
						this.ProgressSpinnerDlg = false;
						this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
						this.router.navigate(['']);
					} else {
						this.ProgressSpinnerDlg = false;
						this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
					}
				},
					error => {
						this.ProgressSpinnerDlg = false;
						this.messageService.add({ severity: 'error', summary: 'Error', detail: error.message });
					}
				);
		}
	}
}
