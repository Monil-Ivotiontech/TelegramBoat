import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { MessageService } from 'primeng/api';
import * as FileSaver from 'file-saver';
import { Table } from 'primeng/table';
import { HttpService } from 'src/app/core/service/http.service';
import { StorageService } from 'src/app/core/service/storage.service';
import { ApiMethod, EndPoints } from 'src/app/core/const';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';

@Component({
  selector: 'app-transaction',
  templateUrl: './transaction.component.html',
  styleUrls: ['./transaction.component.scss'],
  providers: [MessageService]
})
export class TransactionComponent implements OnInit {
  // Global Variables
  ProgressSpinnerDlg = false;
  user_details: any;

  // Deal Variables
  deal_row_data: Array<any> = [];
  deal_col_header: Array<any> = [];
  deal_groups_tables_list: Array<any> = [];
  deal_groups_switched_table_list_name: string = "mt5_manager";
  mt5_manager_name_table: boolean = true;
  venus_manager_name_table: boolean = false;
  deal_groups_radio_value: any;

  deal_deatils_summary_row_data: Array<any> = [];
  deal_deatils_summary_col_header: Array<any> = [];
  deal_types_tables_list: Array<any> = [];
  deal_types_switched_table_list_name: string = "details";
  deal_summary_table: boolean = false;
  deal_details_table: boolean = false;

  // Position Variables
  position_group_row_data: Array<any> = [];
  position_group_col_header: Array<any> = [];
  position_groups_tables_list: Array<any> = [];
  position_groups_switched_table_list_name: string = "mt5_manager";
  position_groups_radio_value: any;
  position_mt5_manager_name_table: boolean = true;
  position_venus_manager_name_table: boolean = false;
  position_details_table: boolean = false;
  position_row_data: Array<any> = [];
  position_col_header: Array<any> = [];

  // Trade Variables
  trade_row_data: Array<any> = [];
  trade_col_header: Array<any> = [];

  constructor(
    private messageService: MessageService,
    private storageService: StorageService,
    private httpService: HttpService,
    private router: Router,
    private formBuilder: FormBuilder
  ) { }

  ngOnInit(): void {
    this.user_details = this.storageService.getLocalObject("userdetails");
    this.deal_groups_tables_list = [{ label: 'MT5 Manager', value: 'mt5_manager' }, { label: 'Venus Manager', value: 'venus_manager' }];
    this.position_groups_tables_list = [{ label: 'MT5 Manager', value: 'mt5_manager' }, { label: 'Venus Manager', value: 'venus_manager' }];
    if (this.deal_groups_switched_table_list_name == 'mt5_manager') {
      this.get_mt5_manager_detail();
    }
    else if (this.deal_groups_switched_table_list_name == 'venus_manager') {
      this.get_venus_manager_detail();
    }
    this.deal_types_tables_list = [{ label: 'Details', value: 'details' }, { label: 'Summary', value: 'summary' }];

  }

  // Global Methods

  onTabChange(event: { index: number; }) {
    if (event.index == 0) {
      if (this.deal_groups_switched_table_list_name == 'mt5_manager') {
        this.get_mt5_manager_detail();
      }
      else if (this.deal_groups_switched_table_list_name == 'venus_manager') {
        this.get_venus_manager_detail();
      }
    }
    else if (event.index == 1) {
      if (this.position_groups_switched_table_list_name == 'mt5_manager') {
        this.get_position_mt5_manager_detail();
      }
      else if (this.position_groups_switched_table_list_name == 'venus_manager') {
        this.get_position_venus_manager_detail();
      }
    }
    else if (event.index == 2) { }
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
  deal_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.deal_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "deal_groups_data");
    });
  }

  deal_details_summary_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.deal_deatils_summary_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "deal_data");
    });
  }

  position_group_export() { }

  position_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.position_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "position_data");
    });
  }

  trade_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.trade_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "trade_data");
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

  // *************************************** DEAL METHODS ********************************
  deal_groups_table_name_switched(event: any) {
    this.deal_groups_switched_table_list_name = event.option.value;
    if (this.deal_groups_switched_table_list_name == 'mt5_manager') {
      this.get_mt5_manager_detail();
    }
    else if (this.deal_groups_switched_table_list_name == 'venus_manager') {
      this.get_venus_manager_detail();
    }
    this.deal_groups_radio_value = undefined;
  }

  get_mt5_manager_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_mt5_manager_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.deal_row_data = response.manager_detail;
          this.deal_col_header = Object.keys(this.deal_row_data[0]);
          this.mt5_manager_name_table = true;
          this.venus_manager_name_table = false;
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.deal_row_data = [];
          this.deal_col_header = [];
          this.mt5_manager_name_table = false;
          this.venus_manager_name_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.deal_row_data = [];
        this.deal_col_header = [];
        this.mt5_manager_name_table = false;
        this.venus_manager_name_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  get_venus_manager_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_venus_manager_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.deal_row_data = response.manager_data;
          this.deal_col_header = Object.keys(this.deal_row_data[0]);
          this.venus_manager_name_table = true;
          this.mt5_manager_name_table = false;
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.deal_row_data = [];
          this.deal_col_header = [];
          this.venus_manager_name_table = false;
          this.mt5_manager_name_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.deal_row_data = [];
        this.deal_col_header = [];
        this.venus_manager_name_table = false;
        this.mt5_manager_name_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  deal_groups_radio_checked(data: any) {
    this.deal_groups_radio_value = undefined;
    this.deal_groups_radio_value = data;
  }

  deal_types_table_name_switched(event: any) {
    this.deal_types_switched_table_list_name = event.option.value;
    if (this.deal_types_switched_table_list_name == 'details') {
      this.get_mt5_managerwise_user_deal_detail();
    }
    else if (this.deal_types_switched_table_list_name == 'summary') {
      this.get_mt5_managerwise_userwise_deal_summary();
    }
  }

  get_mt5_managerwise_user_deal_detail() {
    if (this.deal_groups_radio_value == null || this.deal_groups_radio_value == undefined) {
      this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Please Select Group Name First!' });
      return;
    }
    else {
      let command_id;
      if (this.deal_groups_switched_table_list_name == 'mt5_manager') {
        command_id = this.deal_groups_radio_value.Login;
      }
      else if (this.deal_groups_switched_table_list_name == 'venus_manager') {
        command_id = this.deal_groups_radio_value.managerid;
      }
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let commandid = command_id;
      const credentials = { "userid": userid, "email": email, "commandid": commandid };
      this.httpService.frontendRequestCall(EndPoints.get_mt5_groupwise_user_deal_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.deal_deatils_summary_row_data = response.deal_detail;
            this.deal_deatils_summary_col_header = Object.keys(this.deal_deatils_summary_row_data[0]);
            this.deal_details_table = true;
            this.deal_summary_table = false;
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.deal_groups_radio_value = {};
            this.deal_deatils_summary_row_data = [];
            this.deal_deatils_summary_col_header = [];
            this.deal_details_table = false;
            this.deal_summary_table = false;
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.deal_groups_radio_value = {};
          this.deal_deatils_summary_row_data = [];
          this.deal_deatils_summary_col_header = [];
          this.deal_details_table = false;
          this.deal_summary_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
  }

  get_mt5_managerwise_userwise_deal_summary() {
    if (this.deal_groups_radio_value == null || this.deal_groups_radio_value == undefined) {
      this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Please Select Group Name First!' });
      return;
    }
    else {
      let group_name, group_type;
      if (this.deal_groups_switched_table_list_name == 'mt5_manager') {
        group_name = this.deal_groups_radio_value.Group;
        group_type = "mt5";
      }
      else if (this.deal_groups_switched_table_list_name == 'venus_manager') {
        group_name = this.deal_groups_radio_value.groupname;
        group_type = "venus";
      }
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let groupname = group_name;
      let grouptype = group_type;
      const credentials = { "userid": userid, "email": email, "groupname": groupname, "grouptype": grouptype };
      this.httpService.frontendRequestCall(EndPoints.get_mt5_groupwise_userwise_deal_summary, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.deal_deatils_summary_row_data = response.deal_detail;
            this.deal_deatils_summary_col_header = Object.keys(this.deal_deatils_summary_row_data[0]);
            this.deal_details_table = false;
            this.deal_summary_table = true;
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.deal_groups_radio_value = {};
            this.deal_deatils_summary_row_data = [];
            this.deal_deatils_summary_col_header = [];
            this.deal_details_table = false;
            this.deal_summary_table = false;
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.deal_groups_radio_value = {};
          this.deal_deatils_summary_row_data = [];
          this.deal_deatils_summary_col_header = [];
          this.deal_details_table = false;
          this.deal_summary_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
  }

  // *************************************** POSITION METHODS ********************************
  open_new_position_config() { }

  position_groups_table_name_switched(event: any) {
    this.position_group_row_data = [];
    this.position_group_col_header = [];
    this.position_groups_switched_table_list_name = event.option.value;
    if (this.position_groups_switched_table_list_name == 'mt5_manager') {
      this.get_position_mt5_manager_detail();
    }
    else if (this.position_groups_switched_table_list_name == 'venus_manager') {
      this.get_position_venus_manager_detail();
    }
  }

  get_position_mt5_manager_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_mt5_manager_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.position_group_row_data = response.manager_detail;
          this.position_group_col_header = Object.keys(this.position_group_row_data[0]);
          this.position_mt5_manager_name_table = true;
          this.position_venus_manager_name_table = false;
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.position_group_row_data = [];
          this.position_group_col_header = [];
          this.position_mt5_manager_name_table = false;
          this.position_venus_manager_name_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.position_group_row_data = [];
        this.position_group_col_header = [];
        this.position_mt5_manager_name_table = false;
        this.position_venus_manager_name_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  get_position_venus_manager_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_venus_manager_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.position_group_row_data = response.manager_data;
          this.position_col_header = Object.keys(this.position_group_row_data[0]);
          this.position_venus_manager_name_table = true;
          this.position_mt5_manager_name_table = false;
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.position_group_row_data = [];
          this.position_group_col_header = [];
          this.position_mt5_manager_name_table = false;
          this.position_venus_manager_name_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.position_group_row_data = [];
        this.position_group_col_header = [];
        this.position_mt5_manager_name_table = false;
        this.position_venus_manager_name_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  get_mt5_managerwise_user_position_detail() {
    if (this.position_groups_radio_value == null || this.position_groups_radio_value == undefined) {
      this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Please Select Group Name First!' });
      return;
    }
    else {
      let command_id;
      if (this.position_groups_switched_table_list_name == 'mt5_manager') {
        command_id = this.position_groups_radio_value.Login;
      }
      else if (this.position_groups_switched_table_list_name == 'venus_manager') {
        command_id = this.position_groups_radio_value.managerid;
      }
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let commandid = command_id;
      const credentials = { "userid": userid, "email": email, "commandid": commandid };
      this.httpService.frontendRequestCall(EndPoints.get_mt5_groupwise_user_position_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.position_row_data = response.position_detail;
            this.position_col_header = Object.keys(this.position_row_data[0]);
            this.position_details_table = true;
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.position_groups_radio_value = {};
            this.position_row_data = [];
            this.position_col_header = [];
            this.position_details_table = false;
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.position_groups_radio_value = {};
          this.position_row_data = [];
          this.position_col_header = [];
          this.position_details_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
  }

  position_groups_radio_checked(data: any) {
    this.position_groups_radio_value = undefined;
    this.position_groups_radio_value = data;
  }

  // *************************************** TRADE METHODS ********************************
  open_new_trade_config() { }

}
