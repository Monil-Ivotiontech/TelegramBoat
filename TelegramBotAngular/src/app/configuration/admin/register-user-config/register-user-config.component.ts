import { Component, OnInit } from '@angular/core';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { MessageService } from 'primeng/api';
import { EndPoints, ApiMethod } from 'src/app/core/const';
import { HttpService } from 'src/app/core/service/http.service';
import { StorageService } from 'src/app/core/service/storage.service';

@Component({
  selector: 'app-register-user-config',
  templateUrl: './register-user-config.component.html',
  styleUrls: ['./register-user-config.component.scss'],
  providers: [MessageService]
})
export class RegisterUserConfigComponent implements OnInit {
  // Global Variables
  ProgressSpinnerDlg = false;
  user_details: any;

  // Register user variables
  user_register_form!: FormGroup;
  roles_list: Array<any> = [];
  selected_role_name: any;
  role_details: boolean = false;

  constructor(
    private router: Router,
    private formBuilder: FormBuilder,
    private messageService: MessageService,
    private storageService: StorageService,
    private httpService: HttpService
  ) { }

  ngOnInit(): void {
    this.user_details = this.storageService.getLocalObject("userdetails");
    this.user_register_form = this.formBuilder.group({
      firstname: ['', Validators.required],
      lastname: ['', Validators.required],
      email: ['', Validators.required],
      mobile_no: ['', Validators.required],
      password: ['', Validators.required],
      confirmpassword: ['', Validators.required],
      country_name: ['', Validators.required],
      role_list: ['', Validators.required]
    });

    this.user_register_form.reset();
  }

  // ********************** GLOBAL METHODS ****************************
  back_to_previous_page() {
    this.router.navigate(['/configuration/platform']);
  }

  refreshTable() {
    let currentUrl = this.router.url;
    this.router.routeReuseStrategy.shouldReuseRoute = () => false;
    this.router.onSameUrlNavigation = 'reload';
    this.router.navigate([currentUrl]);
  }


  // ********************** REGISTER USER METHODS ****************************
  resetForm() {
    this.user_register_form.reset();
  }

  onCountrySelect(event: any) {
    this.role_details = true;
    this.roles_list = ["SUPERUSER", "ADMIN", "STAFF"];
  }

  onRoleChange($event: { value: never[]; }) {
    this.selected_role_name = $event.value || [];
  }

  register_user() {
    if (this.user_register_form.invalid) {
      this.messageService.add({ severity: 'error', summary: 'Error Message', detail: 'Please fill all the required fields' });
      return;
    }
    else if (this.selected_role_name.length == 0) {
      this.messageService.add({ severity: 'error', summary: 'Error Message', detail: 'Please select atleast one role' });
      return;
    }
    else if (this.user_register_form.value.password != this.user_register_form.value.confirmpassword) {
      this.messageService.add({ severity: 'error', summary: 'Error Message', detail: 'Password and Confirm Password should be same' });
      return;
    }
    else if (this.user_register_form.value.password.length < 6) {
      this.messageService.add({ severity: 'error', summary: 'Error Message', detail: 'Password should be atleast 6 characters' });
      return;
    }
    else {
      let reg_detail = {
        "firstname": this.user_register_form.value.firstname,
        "lastname": this.user_register_form.value.lastname,
        "email": this.user_register_form.value.email,
        "password": this.user_register_form.value.password,
        "confpassword": this.user_register_form.value.confirmpassword,
        "mobile": this.user_register_form.value.mobile_no,
        "country": this.user_register_form.value.country_name,
        "rolename": this.selected_role_name
      }
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let useremail = this.user_details.email;
      let regdetail = reg_detail;
      const credentials = { "userid": userid, "useremail": useremail, "regdetail": regdetail };
      this.httpService.frontendRequestCall(EndPoints.register_admin_user, ApiMethod.POST, credentials)
        .subscribe((response: { message?: any; status?: any; }) => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || 'false';
          if (status == "true") {
            this.ProgressSpinnerDlg = false;
            this.messageService.add({ severity: 'success', summary: 'Success Message', detail: message });
            this.user_register_form.reset();
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.messageService.add({ severity: 'error', summary: 'Error Message', detail: message });
          }
        }, (error: any) => {
          this.ProgressSpinnerDlg = false;
          this.messageService.add({ severity: 'error', summary: 'Error Message', detail: error });
        }
        );
    }
  }

}
