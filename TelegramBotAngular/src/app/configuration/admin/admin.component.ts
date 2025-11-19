import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { MessageService } from 'primeng/api';
import * as FileSaver from 'file-saver';
import { Table } from 'primeng/table';
import { HttpService } from 'src/app/core/service/http.service';
import { StorageService } from 'src/app/core/service/storage.service';
import { ApiMethod, EndPoints } from 'src/app/core/const';

@Component({
  selector: 'app-admin',
  templateUrl: './admin.component.html',
  styleUrls: ['./admin.component.scss'],
  providers: [MessageService]
})
export class AdminComponent implements OnInit {
  // Global Variables
  ProgressSpinnerDlg = false;
  user_details: any;

  // Admin Users List Variables
  registered_admin_row_data: Array<any> = [];
  registered_admin_col_header: Array<any> = [];
  selected_deletable_reg_admin_id: Array<any> = [];
  status_list: Array<any> = [];
  status_selected_value: string = 'active';
  registered_admin_list_resp: any;

  delete_user_dialog: boolean = false;
  selected_deletable_data: any;
  selected_first_name: string = '';
  selected_last_name: string = '';

  constructor(
    private messageService: MessageService,
    private storageService: StorageService,
    private httpService: HttpService,
    private router: Router
  ) { }

  // Global Methods
  ngOnInit(): void {
    this.user_details = this.storageService.getLocalObject("userdetails");

    this.status_list = [
      { label: 'Active', value: 'active' },
      { label: 'Inactive', value: 'inactive' }
    ];
    this.get_registered_admin_list();

  }

  onTabChange(event: { index: number; }) {
    if (event.index == 0) {
      this.get_registered_admin_list();
    }
  }

  refreshTable() {
    let currentUrl = this.router.url;
    this.router.routeReuseStrategy.shouldReuseRoute = () => false;
    this.router.onSameUrlNavigation = 'reload';
    this.router.navigate([currentUrl]);
  }
  onGlobalFilter(table: Table, event: Event) {
    table.filterGlobal((event.target as HTMLInputElement).value, 'contains');
  }

  // Export table data to csv
  admin_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.registered_admin_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "registered_admin_data");
    });
  }

  saveAsCSVFile(buffer: any, fileName: string): void {
    let CSV_TYPE = 'text/csv;charset=utf-8;';
    let CSV_EXTENSION = '.csv';
    const data: Blob = new Blob([buffer], {
      type: CSV_TYPE
    });
    FileSaver.saveAs(data, fileName + '_export_' + new Date().getTime() + CSV_EXTENSION);
  }

  //**** Admin Users List Main Methods */

  get_registered_admin_list() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_registered_admin_list, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.registered_admin_list_resp = response;
          this.registered_admin_row_data = this.registered_admin_list_resp.admin_data[this.status_selected_value];
          this.registered_admin_col_header = Object.keys(this.registered_admin_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.registered_admin_row_data = [];
          this.registered_admin_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.registered_admin_row_data = [];
        this.registered_admin_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  // status method
  on_status_select(event: any) {
    this.status_selected_value = event.value;
    this.registered_admin_row_data = this.registered_admin_list_resp.admin_data[this.status_selected_value];
    this.registered_admin_col_header = Object.keys(this.registered_admin_row_data[0]);
  }

  // new user creation method
  open_new_user_config() {
    this.router.navigate(['/configuration/platform/register-new-user']);
  }

  // Delete selected users in single
  delete_selected_reg_admin_data(data: any) {
    this.selected_deletable_data = data;
    this.selected_first_name = data.firstname;
    this.selected_last_name = data.lastname;
    this.delete_user_dialog = true;
  }

  cancel_user_delete() {
    this.delete_user_dialog = false;
    this.selected_deletable_data = [];
  }

  confirm_user_delete() {
    this.delete_user_dialog = false;
    if (this.user_details.role == "SUPERUSER") {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let adminuserid = this.selected_deletable_data.userid;
      const credentials = { "userid": userid, "email": email, "adminuserid": adminuserid };
      this.httpService.frontendRequestCall(EndPoints.delete_registered_admin_user, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || 'false';
          if (status == "true") {
            this.get_registered_admin_list();
            this.ProgressSpinnerDlg = false;
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
            this.selected_deletable_data = [];
            this.selected_first_name = '';
            this.selected_last_name = '';
          } else {
            this.ProgressSpinnerDlg = false;
            this.get_registered_admin_list();
            this.selected_deletable_data = [];
            this.selected_first_name = '';
            this.selected_last_name = '';
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.get_registered_admin_list();
          this.selected_deletable_data = [];
          this.selected_first_name = '';
          this.selected_last_name = '';
          this.messageService.add({ severity: 'error', summary: 'Something Went Wrong!', detail: error });
        }
        );
    }
    else {
      this.selected_deletable_data = [];
      this.selected_first_name = '';
      this.selected_last_name = '';
      this.get_registered_admin_list();
      this.messageService.add({ severity: 'warn', summary: 'Warning', detail: 'Only Superuser can delete registered users!' });
    }
  }

}
