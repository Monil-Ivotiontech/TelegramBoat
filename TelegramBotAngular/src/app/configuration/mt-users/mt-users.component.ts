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
  selector: 'app-mt-users',
  templateUrl: './mt-users.component.html',
  styleUrls: ['./mt-users.component.scss'],
  providers: [MessageService]
})
export class MtUsersComponent implements OnInit {
  // Global Variables
  ProgressSpinnerDlg = false;
  user_details: any;

  //  MT5 MANAGER VARIABLES 
  manager_row_data: Array<any> = [];
  manager_col_header: Array<any> = [];

  manager_group_row_data: Array<any> = [];
  manager_group_col_header: Array<any> = [];
  manager_group_list_table: boolean = false;
  selected_manager_table_type: string = 'Selected Managers Group List';

  selected_clonnable_manager_record_data: any;
  clone_mt5_manager_dialog: boolean = false;
  clone_managaer_form!: FormGroup;
  create_manager_chat_dialog: boolean = false;
  manager_chat_form!: FormGroup;
  selected_manager_chat_data: any;
  selected_available_Chatid: any;
  available_chatid: any;
  selected_cmd_access: boolean = false;
  selected_dw_access: boolean = false;
  selected_exp_date: any;

  // MT5 Groups Variables
  groups_row_data: Array<any> = [];
  groups_col_header: Array<any> = [];
  groups_tables_list: Array<any> = [];
  groups_switched_table_list_name: string = "details";

  groupwise_user_row_data: Array<any> = [];
  groupwise_user_col_header: Array<any> = [];
  count_details_table: boolean = false;

  clone_mt5_group_dialog: boolean = false;
  new_group_name: any;
  ref_group_name: any;

  // MT5 Users Variables
  users_row_data: Array<any> = [];
  users_col_header: Array<any> = [];
  clonnable_mt5_new_user_data: any;
  clone_mt5_user_dialog: boolean = false;
  create_new_mt5_user_dialog: boolean = false;
  clone_user_form!: FormGroup;
  create_new_mt5_user_form!: FormGroup;
  update_mt5_user_dialog: boolean = false;
  dummy_id_confirmation_dialog: boolean = false;
  update_mt5_user_form!: FormGroup;
  selected_mt5_user_data: any;
  selected_dummy_id_raw_data: any;

  // Venus Manager Variables
  venus_manager_row_data: Array<any> = [];
  venus_manager_col_header: Array<any> = [];

  venus_manager_mapping_row_data: Array<any> = [];
  venus_manager_mapping_col_header: Array<any> = [];
  selected_mapping_data: Array<any> = [];
  venus_manager_mapping_table: boolean = false;
  selected_venus_manager_data: any;
  selected_vm_mapping_data: any;
  venus_manager_mapping_selected_details: Array<any> = [];

  venus_manager_mapping_tables_list: Array<any> = [];
  venus_manager_mapping_switched_table_list_name: string = "mt5_users";

  view_venus_manager_mapping_row_data: Array<any> = [];
  view_venus_manager_mapping_col_header: Array<any> = [];
  view_venus_manager_mapping_table: boolean = false;

  add_new_venus_manager_dialog: boolean = false;
  update_venus_manager_dialog: boolean = false;
  venus_manager_delete_dialog: boolean = false;
  venus_manager_mapped_data_delete_dialog: boolean = false;
  venus_manager_form!: FormGroup;
  update_venus_manager_form!: FormGroup;
  selected_updatable_venus_manager_data: any;
  selected_deletable_row_data: any;
  selected_deletable_mapped_row_data: any;
  selected_view_manager_mapping_details: any;
  venus_manager_view_mapping_switched_table_list_name: string = "mt5_users";
  prev_chat_id: any;
  year: any;
  month: any;
  date: any;
  expdate: any;

  // Venus Group Variables
  venus_group_row_data: Array<any> = [];
  venus_group_col_header: Array<any> = [];
  venus_group_mapping_row_data: Array<any> = [];
  venus_group_mapping_col_header: Array<any> = [];
  selected_group_mapping_data: Array<any> = [];
  venus_group_mapping_selected_details: Array<any> = [];
  view_venus_group_mapping_row_data: Array<any> = [];
  view_venus_group_mapping_col_header: Array<any> = [];
  add_new_venus_group_dialog: boolean = false;
  update_venus_group_dialog: boolean = false;
  venus_group_delete_dialog: boolean = false;
  venus_group_mapping_table: boolean = false;
  view_venus_group_mapping_table: boolean = false;
  venus_group_mapped_data_delete_dialog: boolean = false;
  venus_group_form!: FormGroup;
  update_venus_group_form!: FormGroup;
  selected_updatable_venus_group_data: any;
  selected_deletable_venus_group_row_data: any;
  selected_venus_group_data: any;
  selected_vg_mapping_data: any;
  selected_view_group_mapping_details: any;
  selected_deletable_group_mapped_row_data: any;

  // Bot Access Variables
  bot_access_row_data: Array<any> = [];
  bot_access_col_header: Array<any> = [];
  bot_access_details_row_data: Array<any> = [];
  bot_access_details_col_header: Array<any> = [];
  bot_types_list: Array<any> = [];
  access_detail: Array<any> = [];
  bot_access_switched_table_list_name: string = "mt5_manager";
  bot_access_selected_check_value: Array<any> = [];
  selected_bot_access_type: any;
  selected_bot_access: any;
  bot_access_mt5_manager_table: boolean = false;
  bot_access_venus_manager_table: boolean = false;
  view_bot_access_details_table: boolean = false;
  add_new_bot_access_dialog: boolean = false;
  bot_access_form!: FormGroup;
  update_bot_access_dialog: boolean = false;
  update_bot_access_form!: FormGroup;
  bot_access_delete_dialog: boolean = false;
  selected_bot_access_data_for_update: any;
  selected_deletable_bot_access_row_data: any;
  selected_bot_access_manager_details: any;

  // Command Name Variables
  command_name_row_data: Array<any> = [];
  command_name_col_header: Array<any> = [];
  commission_types: Array<any> = [];
  command_types: Array<any> = [];
  add_new_command_name_dialog: boolean = false;
  venus_command_name_delete_dialog: boolean = false;
  update_venus_cmd_dialog: boolean = false;
  selected_deletable_command_name_data: any;
  selected_updatable_venus_cmd_data: any;
  selected_cmd_no: any;
  selected_commission_type: any;
  command_name_form!: FormGroup;
  update_venus_cmd_form!: FormGroup;

  // Command Mapping Variables
  command_mapping_row_data: Array<any> = [];
  command_mapping_col_header: Array<any> = [];
  group_names_mapping_tables_list: Array<any> = [];
  mt5_group_name_mapping_table: boolean = true;
  venus_manager_name_mapping_table: boolean = false;
  group_names_switched_table_list_name: string = "mt5_manager";
  methods_list: Array<any> = [];
  selected_method: any;
  mapping_selected_check_value: Array<any> = [];
  cmd_name_mapping_check_value: Array<any> = [];
  mapping_details_list: Array<any> = [];
  selected_manager_group_data: any;
  selected_command_name_data: any;

  venus_command_mapping_row_data: Array<any> = [];
  venus_command_mapping_col_header: Array<any> = [];

  venus_cgm_row_data: Array<any> = [];
  venus_cgm_col_header: Array<any> = [];
  venus_cgm_delete_dialog: boolean = false;
  selected_deletable_cgm_data: any;
  newSelectedChatIds: any = [];

  constructor(
    private messageService: MessageService,
    private storageService: StorageService,
    private httpService: HttpService,
    private router: Router,
    private formBuilder: FormBuilder
  ) { }

  ngOnInit(): void {
    this.user_details = this.storageService.getLocalObject("userdetails");
    this.load();
  }

  // ********************************* Global Methods ********************************

  validateNumber(event: any) {
    const charCode = (event.which) ? event.which : event.keyCode;
    if (charCode > 31 && (charCode < 48 || charCode > 57)) {
      return false;
    }
    return true;
  }


  public load(): void {
    this.get_mt5_manager_detail();
    this.clone_managaer_form = this.formBuilder.group({
      loginid: ['', Validators.required],
      username: ['', Validators.required],
      firstname: ['', Validators.required],
      lastname: ['', Validators.required],
      useremail: ['', Validators.required],
      state: ['', Validators.required],
      zipcode: ['', Validators.required],
    });

    this.manager_chat_form = this.formBuilder.group({
      chatid: ['', Validators.required],
      cmd_access: ['', Validators.required],
      dw_access: ['', Validators.required],
    });

    this.manager_chat_form.reset();

    this.clone_user_form = this.formBuilder.group({
      loginid: ['', Validators.required],
      username: ['', Validators.required],
      firstname: ['', Validators.required],
      lastname: ['', Validators.required],
      useremail: ['', Validators.required],
    });

    this.create_new_mt5_user_form = this.formBuilder.group({
      loginid: ['', Validators.required],
      username: ['', Validators.required],
      firstname: ['', Validators.required],
      lastname: ['', Validators.required],
      useremail: ['', Validators.required],
      state: ['', Validators.required],
      zipcode: ['', Validators.required],
      grpname: ['', Validators.required],
      leverage: ['', Validators.required]
    });

    this.update_mt5_user_form = this.formBuilder.group({
      state: ['', Validators.required],
      zipcode: ['', Validators.required],
      leverage: ['', Validators.required]
    });

    this.venus_manager_form = this.formBuilder.group({
      manager_name: ['', Validators.required],
      manager_desc: ['', Validators.required],
      chatid: ['', Validators.required],
      cmd_access: ['', Validators.required],
      exp_date: ['', Validators.required],
    });

    this.venus_manager_form.reset();

    this.update_venus_manager_form = this.formBuilder.group({
      chatid: ['', Validators.required],
      cmd_access: ['', Validators.required],
      exp_date: ['', Validators.required],
    });

    this.update_venus_manager_form.reset();

    this.venus_group_form = this.formBuilder.group({
      group_master: ['', Validators.required],
      group_name: ['', Validators.required],
      group_desc: ['', Validators.required]
    });

    this.update_venus_group_form = this.formBuilder.group({
      group_master: ['', Validators.required],
      group_name: ['', Validators.required],
      group_desc: ['', Validators.required]
    });

    this.bot_access_form = this.formBuilder.group({
      bot_type: ['', Validators.required],
    });

    this.bot_access_form.reset();

    this.update_bot_access_form = this.formBuilder.group({
      bot_type: ['', Validators.required],
    });

    this.update_bot_access_form.reset();

    this.command_name_form = this.formBuilder.group({
      command_types: ['', Validators.required],
      command_name: ['', Validators.required],
      command_desc: ['', Validators.required],
      commission_type: ['', Validators.required]
    });

    this.update_venus_cmd_form = this.formBuilder.group({
      commission_type: ['', Validators.required],
      cmd_seq_no: ['', Validators.required]
    });
  }

  onTabChange(event: { index: number; }) {
    if (event.index == 0) {
      this.get_mt5_manager_detail();
    }
    else if (event.index == 1) {
      this.groups_tables_list = [{ label: 'Summary', value: 'summary' }, { label: 'Details', value: 'details' }];
      if (this.groups_switched_table_list_name == 'details') {
        this.get_mt5_group_detail();
        this.count_details_table = false;
      }
      else {
        this.get_mt5_groupwise_user_count();
      }
    }
    else if (event.index == 2) {
      this.get_mt5_user_detail();
    }

    else if (event.index == 3) {
      this.get_venus_group_detail();
      this.view_venus_group_mapping_table = false;
      this.venus_group_mapping_table = false;
    }
    else if (event.index == 4) {
      this.venus_manager_mapping_tables_list = [{ label: 'MT5 Users', value: 'mt5_users' }, { label: 'Venus Group', value: 'venus_group' }];
      this.get_venus_manager_detail();
      this.view_venus_manager_mapping_table = false;
      this.venus_manager_mapping_table = false;
    }
    else if (event.index == 5) {
      this.group_names_mapping_tables_list = [{ label: 'MT5 Manager', value: 'mt5_manager' }, { label: 'Venus Manager', value: 'venus_manager' }];
      if (this.bot_access_switched_table_list_name == 'mt5_manager') {
        this.bot_access_mt5_manager();
      }
      else if (this.bot_access_switched_table_list_name == 'venus_manager') {
        this.bot_access_venus_manager();
      }

      this.bot_types_list = ['ZIP', 'EMAIL', 'PHONE', 'STATE', 'CITY', 'IDNUMBER'];
      this.view_bot_access_details_table = false;
    }

    else if (event.index == 6) {
      this.commission_types = ['ours', 'partner'];
      this.command_types = ['mt5', 'venus'];
      this.get_venus_command_detail();
    }
    else if (event.index == 7) {
      this.selected_manager_group_data = [];
      this.selected_command_name_data = [];
      this.get_venus_command_list_for_mapping();
      this.group_names_mapping_tables_list = [{ label: 'MT5 Manager', value: 'mt5_manager' }, { label: 'Venus Manager', value: 'venus_manager' }];
      this.methods_list = [
        { label: 'M2M', value: 'M2M' },
        { label: 'COM POS', value: 'COM POS' },
        { label: 'TOTAL POS', value: 'TOTAL POS' },
        { label: 'Update All', value: 'UPDATE ALL' },
        { label: 'TOP 5', value: 'TOP 5' },
        { label: 'TOP 10', value: 'TOP 10' }
      ];

      if (this.group_names_switched_table_list_name == 'mt5_manager') {
        this.get_mt5_group_mapping_detail();
        this.get_mt5_command_list_for_mapping();
      }
      else if (this.group_names_switched_table_list_name == 'venus_manager') {
        this.get_venus_mapping_group_detail();
        this.get_venus_command_list_for_mapping();
      }
      this.get_venus_command_mapping_detail();
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
  saveAsCSVFile(buffer: any, fileName: string): void {
    let CSV_TYPE = 'text/csv;charset=utf-8;';
    let CSV_EXTENSION = '.csv';
    const data: Blob = new Blob([buffer], {
      type: CSV_TYPE
    });
    FileSaver.saveAs(data, fileName + '_export_' + new Date().getTime() + CSV_EXTENSION);
  }

  // ********************************* MT5 MANAGER TAB METHODS ********************************
  manager_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.manager_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "manager_data");
    });
  }

  get_mt5_manager_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.getMatUserData(credentials);
  }

  getMatUserData(data: any) {
    this.httpService.frontendRequestCall(EndPoints.get_mt5_manager_detail, ApiMethod.POST, data).subscribe({
      next: (response: any) => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {    
          this.ProgressSpinnerDlg = false;
          this.manager_row_data = response.manager_detail;
          this.manager_col_header = Object.keys(this.manager_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.manager_row_data = [];
          this.manager_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error: (error: any) => {
        this.ProgressSpinnerDlg = false;
        this.manager_row_data = [];
        this.manager_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later111111!' });
      }
    });
  }

  manager_group_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.manager_group_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "manager_groups");
    });
  }

  clone_mt5_manager_record(data: any) {
    this.selected_clonnable_manager_record_data = data;
    this.clone_mt5_manager_dialog = true;
  }

  cancel_mt5_manager_clonning() {
    this.clone_mt5_manager_dialog = false;
    this.clone_managaer_form.reset();
  }

  confirm_mt5_manager_clonning() {
    this.clone_mt5_manager_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let managerloginid = this.selected_clonnable_manager_record_data.Login;
    let loginid = this.clone_managaer_form.controls.loginid.value;
    let username = this.clone_managaer_form.controls.username.value;
    let firstname = this.clone_managaer_form.controls.firstname.value;
    let lastname = this.clone_managaer_form.controls.lastname.value;
    let useremail = this.clone_managaer_form.controls.useremail.value;
    let state = this.clone_managaer_form.controls.state.value;
    let zipcode = this.clone_managaer_form.controls.zipcode.value;
    const credentials = {
      "userid": userid, "email": email, "managerloginid": managerloginid, "loginid": loginid, "username": username, "firstname": firstname, "lastname": lastname, "useremail": useremail, "state": state, "zipcode": zipcode
    };
    this.httpService.frontendRequestCall(EndPoints.clone_mt5_manager_to_new, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.get_mt5_manager_detail();
          this.clone_managaer_form.reset();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_mt5_manager_detail();
          this.clone_managaer_form.reset();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.get_mt5_manager_detail();
        this.clone_managaer_form.reset()
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  open_manager_chat_dlg(data: any) {
    this.selected_manager_chat_data = data;
    this.get_mt5_manager_chatid();
    this.create_manager_chat_dialog = true;
  }

  get_mt5_manager_chatid() {
    this.newSelectedChatIds = []
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let managerloginid = this.selected_manager_chat_data.Login;
    const credentials = { "userid": userid, "email": email, "managerloginid": managerloginid };
    this.httpService.frontendRequestCall(EndPoints.get_mt5_manager_chatid, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.available_chatid = response.chatid;
          this.newSelectedChatIds = response?.chatids.map(String);  
          this.selected_available_Chatid = this.available_chatid;
          if (response.commandaccess == "ENABLE") {
            this.selected_cmd_access = true;
          }
          else {
            this.selected_cmd_access = false;
          }
          if (response.dwaccess == "ENABLE") {
            this.selected_dw_access = true;
          }
          else {
            this.selected_dw_access = false;
          }
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.available_chatid = "";
          this.selected_available_Chatid = "";
          if (response.commandaccess == "ENABLE") {
            this.selected_cmd_access = true;
          }
          else {
            this.selected_cmd_access = false;
          }
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.available_chatid = "";
        this.selected_available_Chatid = "";
        this.selected_cmd_access = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  cmd_access_switch_change() {
    this.selected_cmd_access == !this.selected_cmd_access;
  }

  dw_access_switch_change() {
    this.selected_dw_access == !this.selected_dw_access;
  }

  cancel_manager_chat() {
    this.create_manager_chat_dialog = false;
  }

  confirm_manager_chat() {
    this.create_manager_chat_dialog = false;
    let cmd_access, dw_access
    if (this.selected_cmd_access == true) {
      cmd_access = "ENABLE";
    }
    else if (this.selected_cmd_access == false) {
      cmd_access = "DISABLE";
    }
    if (this.selected_dw_access == true) {
      dw_access = "ENABLE";
    }
    else if (this.selected_dw_access == false) {
      dw_access = "DISABLE";
    }
    if (this.available_chatid !== "" || this.available_chatid === 0) {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let managerloginid = this.selected_manager_chat_data.Login;
      let chatid = this.selected_available_Chatid;
      let commandaccess = cmd_access;
      let dwaccess = dw_access;
      const credentials = { "userid": userid, "email": email, "managerloginid": managerloginid, "chatid": this.manager_chat_form?.value?.chatid?.toString(), "commandaccess": commandaccess, "dwaccess": dwaccess };
      this.httpService.frontendRequestCall(EndPoints.update_mt5_manager_chatid, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.manager_chat_form.reset();
            this.get_mt5_manager_detail();
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.manager_chat_form.reset();
            this.get_mt5_manager_detail();
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.manager_chat_form.reset();
          this.get_mt5_manager_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
    else if (this.available_chatid === "") {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let managerloginid = this.selected_manager_chat_data.Login;
      let chatid = this.selected_available_Chatid;
      let commandaccess = cmd_access;
      let dwaccess = dw_access;
      const credentials = { "userid": userid, "email": email, "managerloginid": managerloginid, "chatid": this.manager_chat_form?.value?.chatid?.toString(), "commandaccess": commandaccess, "dwaccess": dwaccess };
      this.httpService.frontendRequestCall(EndPoints.save_mt5_manager_chatid, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.manager_chat_form.reset();
            this.get_mt5_manager_detail();
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.manager_chat_form.reset();
            this.get_mt5_manager_detail();
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.manager_chat_form.reset();
          this.get_mt5_manager_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
  }

  get_mt5_group_list_for_selected_manager(data: any) {
    this.selected_manager_table_type = 'Selected Manager Group List';
    this.manager_group_row_data = [];
    this.manager_group_col_header = [];
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let managerloginid = data.Login;
    const credentials = { "userid": userid, "email": email, "managerloginid": managerloginid };
    this.httpService.frontendRequestCall(EndPoints.get_mt5_group_list_for_selected_manager, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.manager_group_row_data = response.group_list;
          this.manager_group_col_header = Object.keys(this.manager_group_row_data[0]);
          this.manager_group_list_table = true;
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.manager_group_row_data = [];
          this.manager_group_col_header = [];
          this.manager_group_list_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.manager_group_row_data = [];
        this.manager_group_col_header = [];
        this.manager_group_list_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  get_mt5_user_list_for_selected_manager(data: any) {
    this.selected_manager_table_type = 'Selected Manager Users List';
    this.manager_group_row_data = [];
    this.manager_group_col_header = [];
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let managerloginid = data.Login;
    const credentials = { "userid": userid, "email": email, "managerloginid": managerloginid };
    this.httpService.frontendRequestCall(EndPoints.get_mt5_user_list_for_selected_manager, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.manager_group_row_data = response.user_detail;
          this.manager_group_col_header = Object.keys(this.manager_group_row_data[0]);
          this.manager_group_list_table = true;
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.manager_group_row_data = [];
          this.manager_group_col_header = [];
          this.manager_group_list_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.manager_group_row_data = [];
        this.manager_group_col_header = [];
        this.manager_group_list_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  // ********************************* Group Methods **********************************
  groups_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.groups_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "groups_data");
    });
  }

  groupwise_user_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.groupwise_user_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "groupwise_user_data");
    });
  }

  get_mt5_group_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_mt5_group_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.groups_row_data = response.group_detail;
          this.groups_col_header = Object.keys(this.groups_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.groups_row_data = [];
          this.groups_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.groups_row_data = [];
        this.groups_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  get_mt5_groupwise_user_count() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_mt5_groupwise_user_count, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.groups_row_data = response.group_detail;
          this.groups_col_header = Object.keys(this.groups_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.groups_row_data = [];
          this.groups_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.groups_row_data = [];
        this.groups_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  groups_table_name_switched(event: any) {
    this.groups_switched_table_list_name = event.option.value;
    if (this.groups_switched_table_list_name == 'details') {
      this.get_mt5_group_detail();
      this.count_details_table = false;
    }
    else {
      this.get_mt5_groupwise_user_count();
    }
  }

  get_mt5_groupwise_user_detail(data: any) {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let groupname = data.Group;
    const credentials = { "userid": userid, "email": email, "groupname": groupname };
    this.httpService.frontendRequestCall(EndPoints.get_mt5_groupwise_user_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.count_details_table = true;
          this.groupwise_user_row_data = response.user_detail;
          this.groupwise_user_col_header = Object.keys(this.groupwise_user_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.groupwise_user_row_data = [];
          this.count_details_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.groupwise_user_row_data = [];
        this.count_details_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  cloned_mt5_groupwise_dialog(data: any) {
    this.new_group_name = data.Group;
    this.ref_group_name = data.Group;
    this.clone_mt5_group_dialog = true;
  }

  cancel_clone_mt5_group_clonning() {
    this.clone_mt5_group_dialog = false;
  }

  confirm_clone_mt5_group_clonning() {
    this.clone_mt5_group_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let newgrpname = this.new_group_name;
    let refgrpname = this.ref_group_name;
    const credentials = { "userid": userid, "email": email, "newgrpname": newgrpname, "refgrpname": refgrpname };
    this.httpService.frontendRequestCall(EndPoints.clone_mt5_group_to_new_group, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.get_mt5_group_detail();
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_mt5_group_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.get_mt5_group_detail();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  // ************************************** USERS METHODS ********************************
  users_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.users_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "users_data");
    });
  }

  get_mt5_user_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_mt5_user_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.users_row_data = response.user_detail;
          this.users_col_header = Object.keys(this.users_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.users_row_data = [];
          this.users_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.users_row_data = [];
        this.users_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  cancel_dummy_id_confirmation() {
    this.dummy_id_confirmation_dialog = false;
    this.selected_dummy_id_raw_data = "";
  }

  confirm_dummy_id_confirmation() {
    this.ProgressSpinnerDlg = true;
    this.dummy_id_confirmation_dialog = false;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let managerid = this.selected_dummy_id_raw_data.Login;
    let state = this.selected_dummy_id_raw_data.State;
    let zipcode = this.selected_dummy_id_raw_data.ZipCode;
    let grpname = this.selected_dummy_id_raw_data.Group;
    let leverage = this.selected_dummy_id_raw_data.Leverage;
    const credentials = { "userid": userid, "email": email, "managerid": managerid, "state": state, "zipcode": zipcode, "grpname": grpname, "leverage": leverage };
    this.httpService.frontendRequestCall(EndPoints.create_mt5_dummy_user, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          this.get_mt5_user_detail();
          this.selected_dummy_id_raw_data = "";
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.users_row_data = [];
          this.users_col_header = [];
          this.selected_dummy_id_raw_data = "";
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.users_row_data = [];
        this.users_col_header = [];
        this.selected_dummy_id_raw_data = "";
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  create_mt5_dummy_user(data: any) {
    this.dummy_id_confirmation_dialog = true;
    this.selected_dummy_id_raw_data = data;
  }

  clone_mt5_new_user(data: any) {
    this.clonnable_mt5_new_user_data = data;
    this.clone_mt5_user_dialog = true;
  }

  cancel_mt5_user_clonning() {
    this.clone_mt5_user_dialog = false;
    this.clone_user_form.reset();
  }

  confirm_mt5_user_clonning() {
    this.clone_mt5_user_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let loginid = this.clone_user_form.controls.loginid.value;
    let username = this.clone_user_form.controls.username.value;
    let firstname = this.clone_user_form.controls.firstname.value;
    let lastname = this.clone_user_form.controls.lastname.value;
    let useremail = this.clone_user_form.controls.useremail.value;
    let state = this.clonnable_mt5_new_user_data.State;
    let zipcode = this.clonnable_mt5_new_user_data.ZipCode;
    let grpname = this.clonnable_mt5_new_user_data.Group;
    let leverage = this.clonnable_mt5_new_user_data.Leverage;
    const credentials = { "userid": userid, "email": email, "loginid": loginid, "username": username, "firstname": firstname, "lastname": lastname, "useremail": useremail, "state": state, "zipcode": zipcode, "grpname": grpname, "leverage": leverage };
    this.httpService.frontendRequestCall(EndPoints.create_mt5_new_or_clone_user, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          this.get_mt5_user_detail();
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_mt5_user_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.get_mt5_user_detail();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  create_new_mt5_user() {
    this.create_new_mt5_user_dialog = true;
  }

  cancel_mt5_user_creation() {
    this.create_new_mt5_user_dialog = false;
    this.create_new_mt5_user_form.reset();
  }

  confirm_mt5_user_creation() {
    this.create_new_mt5_user_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let loginid = this.create_new_mt5_user_form.controls.loginid.value;
    let username = this.create_new_mt5_user_form.controls.username.value;
    let firstname = this.create_new_mt5_user_form.controls.firstname.value;
    let lastname = this.create_new_mt5_user_form.controls.lastname.value;
    let useremail = this.create_new_mt5_user_form.controls.useremail.value;
    let state = this.create_new_mt5_user_form.controls.state.value;
    let zipcode = this.create_new_mt5_user_form.controls.zipcode.value;
    let grpname = this.create_new_mt5_user_form.controls.grpname.value;
    let leverage = this.create_new_mt5_user_form.controls.leverage.value;
    const credentials = { "userid": userid, "email": email, "loginid": loginid, "username": username, "firstname": firstname, "lastname": lastname, "useremail": useremail, "state": state, "zipcode": zipcode, "grpname": grpname, "leverage": leverage };
    this.httpService.frontendRequestCall(EndPoints.create_mt5_new_or_clone_user, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          this.get_mt5_user_detail();
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_mt5_user_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.get_mt5_user_detail();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  update_selected_mt5_user(data: any) {
    this.update_mt5_user_dialog = true;
    this.selected_mt5_user_data = data;
  }

  cancel_mt5_user_updation() {
    this.update_mt5_user_dialog = false;
    this.update_mt5_user_form.reset();
  }

  confirm_mt5_user_updation() {
    this.update_mt5_user_dialog = false;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let loginid = this.selected_mt5_user_data.Login;
    let username = this.selected_mt5_user_data.Name;
    let useremail = this.selected_mt5_user_data.Email;
    let state = this.update_mt5_user_form.controls.state.value;
    let zipcode = this.update_mt5_user_form.controls.zipcode.value;
    let grpname = this.selected_mt5_user_data.Group;
    let leverage = this.update_mt5_user_form.controls.leverage.value;
    const credentials = { "userid": userid, "email": email, "loginid": loginid, "username": username, "useremail": useremail, "state": state, "zipcode": zipcode, "grpname": grpname, "leverage": leverage };
    this.httpService.frontendRequestCall(EndPoints.update_mt5_user, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          this.get_mt5_user_detail();
          this.update_mt5_user_form.reset();
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_mt5_user_detail();
          this.update_mt5_user_form.reset();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.get_mt5_user_detail();
        this.update_mt5_user_form.reset();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  // ************************************** VENUS MANAGER METHODS ********************************
  venus_manager_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.venus_manager_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "Venus_group");
    });
  }

  venus_manager_mapping_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.venus_manager_mapping_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "Venus_group_mapped_data");
    });
  }

  venus_manager_mapping_table_name_switched(event: any) {
    this.venus_manager_mapping_switched_table_list_name = event.option.value;
    this.map_venus_manager_detail(this.selected_venus_manager_data);
    this.selected_mapping_data = [];
    this.selected_vm_mapping_data = [];
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
          this.venus_manager_row_data = response.manager_data;
          this.venus_manager_col_header = Object.keys(this.venus_manager_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.venus_manager_row_data = [];
          this.venus_manager_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.venus_manager_row_data = [];
        this.venus_manager_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  open_new_venus_manager_config() {
    this.add_new_venus_manager_dialog = true;
  }

  cancel_new_venus_manager_charges() {
    this.add_new_venus_manager_dialog = false;
    this.venus_manager_form.reset();
  }

  confirm_new_venus_manager_charges() {
    this.add_new_venus_manager_dialog = false;
    let cmd_access
    if (this.selected_cmd_access = true) {
      cmd_access = "ENABLE";
    }
    else if (this.selected_cmd_access = false) {
      cmd_access = "DISABLE";
    }
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let managername = this.venus_manager_form.controls.manager_name.value;
    let managerdesc = this.venus_manager_form.controls.manager_desc.value;
    let chatid = this.venus_manager_form.controls.chatid.value;
    let commandaccess = cmd_access;
    let expdate = this.expdate;
    const credentials = { "userid": userid, "email": email, "managername": managername, "managerdesc": managerdesc, "chatid": chatid, "commandaccess": commandaccess, "expdate": expdate };
    this.httpService.frontendRequestCall(EndPoints.save_venus_manager_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.venus_manager_form.reset();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          this.get_venus_manager_detail();
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_venus_manager_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.get_venus_manager_detail();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  venus_manager_view_mapping_table_name_switched(event: any) {
    this.venus_manager_view_mapping_switched_table_list_name = event.option.value;
    this.get_venus_manager_mapping_detail(this.selected_view_manager_mapping_details)
  }

  update_selected_venus_manager_data(data: any) {
    this.update_venus_manager_dialog = true;
    this.selected_updatable_venus_manager_data = data;
    this.prev_chat_id = data.chatid;
    this.expdate = data.expdate;
    if (data.commandaccess == "ENABLE") {
      this.selected_cmd_access = true;
    }
    else {
      this.selected_cmd_access = false;
    }
  }

  confirm_venus_manager_updation() {
    let cmd_access
    if (this.selected_cmd_access == true) {
      cmd_access = "ENABLE";
    }
    else if (this.selected_cmd_access == false) {
      cmd_access = "DISABLE";
    }

    this.update_venus_manager_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let managerid = this.selected_updatable_venus_manager_data.managerid;
    let chatid = this.update_venus_manager_form.controls.chatid.value;
    let commandaccess = cmd_access;
    let expdate = this.expdate;
    const credentials = { "userid": userid, "email": email, "managerid": managerid, "chatid": chatid, "commandaccess": commandaccess, "expdate": expdate };
    this.httpService.frontendRequestCall(EndPoints.update_venus_manager_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.update_venus_manager_form.reset();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          this.get_venus_manager_detail();
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_venus_manager_detail();
          this.update_venus_manager_form.reset();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.get_venus_manager_detail();
        this.update_venus_manager_form.reset();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  cancel_venus_manager_updation() {
    this.update_venus_manager_dialog = false;
    this.update_venus_manager_form.reset();
  }

  delete_venus_manager_data(data: any) {
    this.selected_deletable_row_data = data;
    this.venus_manager_delete_dialog = true;
  }

  cancel_venus_manager_delete() {
    this.venus_manager_delete_dialog = false;
  }

  confirm_venus_manager_delete() {
    this.venus_manager_delete_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let managerid = this.selected_deletable_row_data.managerid;
    const credentials = { "userid": userid, "email": email, "managerid": managerid };
    this.httpService.frontendRequestCall(EndPoints.delete_venus_manager_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || 'false';
        if (status == "true") {
          this.ProgressSpinnerDlg = false;
          this.get_venus_manager_detail();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });

        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_venus_manager_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      },
        error => {
          this.ProgressSpinnerDlg = false;
          this.get_venus_manager_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
        })
  }

  map_venus_manager_detail(data: any) {
    this.selected_venus_manager_data = data;
    if (this.venus_manager_mapping_switched_table_list_name == 'mt5_users') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      const credentials = { "userid": userid, "email": email };
      this.httpService.frontendRequestCall(EndPoints.get_mt5_user_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.venus_manager_mapping_row_data = response.user_detail;
            this.venus_manager_mapping_col_header = Object.keys(this.venus_manager_mapping_row_data[0]);
            this.venus_manager_mapping_table = true;
            this.view_venus_manager_mapping_table = false;
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.venus_manager_mapping_row_data = [];
            this.venus_manager_mapping_col_header = [];
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
            this.venus_manager_mapping_table = false;
            this.view_venus_manager_mapping_table = false;
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.venus_manager_mapping_row_data = [];
          this.venus_manager_mapping_col_header = [];
          this.venus_manager_mapping_table = false;
          this.view_venus_manager_mapping_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
    else if (this.venus_manager_mapping_switched_table_list_name == 'venus_group') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      const credentials = { "userid": userid, "email": email };
      this.httpService.frontendRequestCall(EndPoints.get_venus_group_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.venus_manager_mapping_row_data = response.group_data;
            this.venus_manager_mapping_col_header = Object.keys(this.venus_manager_mapping_row_data[0]);
            this.venus_manager_mapping_table = true;
            this.view_venus_manager_mapping_table = false;
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.venus_manager_mapping_row_data = [];
            this.venus_manager_mapping_col_header = [];
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
            this.venus_manager_mapping_table = false;
            this.view_venus_manager_mapping_table = false;
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.venus_manager_mapping_row_data = [];
          this.venus_manager_mapping_col_header = [];
          this.venus_manager_mapping_table = false;
          this.view_venus_manager_mapping_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }

  }

  save_venus_manager_mapping_detail() {
    if (this.venus_manager_mapping_switched_table_list_name == 'mt5_users') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      for (let i = 0; i < this.selected_mapping_data.length; i++) {
        let mapping_detail =
        {
          "managerid": this.selected_venus_manager_data.managerid,
          "loginid": this.selected_mapping_data[i].Login
        }
        this.venus_manager_mapping_selected_details.push(mapping_detail);
      }
      let mappingdetail = this.venus_manager_mapping_selected_details;
      const credentials = { "userid": userid, "email": email, "mappingdetail": mappingdetail }
      this.httpService.frontendRequestCall(EndPoints.save_venus_manager_mapping_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.selected_mapping_data = [];
            this.selected_vm_mapping_data = [];
            this.venus_manager_mapping_selected_details = [];
            this.venus_manager_mapping_table = false;
            this.get_venus_manager_detail();
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.selected_mapping_data = [];
            this.selected_vm_mapping_data = [];
            this.venus_manager_mapping_selected_details = [];
            this.venus_manager_mapping_table = false;
            this.get_venus_manager_detail();
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.selected_mapping_data = [];
          this.selected_vm_mapping_data = [];
          this.venus_manager_mapping_selected_details = [];
          this.venus_manager_mapping_table = false;
          this.get_venus_manager_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
    else if (this.venus_manager_mapping_switched_table_list_name == 'venus_group') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      for (let i = 0; i < this.selected_mapping_data.length; i++) {
        let mapping_detail =
        {
          "managerid": this.selected_venus_manager_data.managerid,
          "groupid": this.selected_mapping_data[i].groupid
        }
        this.venus_manager_mapping_selected_details.push(mapping_detail);
      }
      let mappingdetail = this.venus_manager_mapping_selected_details;
      const credentials = { "userid": userid, "email": email, "mappingdetail": mappingdetail }
      this.httpService.frontendRequestCall(EndPoints.save_venus_manager_group_mapping_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.selected_mapping_data = [];
            this.selected_vm_mapping_data = [];
            this.venus_manager_mapping_selected_details = [];
            this.venus_manager_mapping_table = false;
            this.get_venus_manager_detail();
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.selected_mapping_data = [];
            this.selected_vm_mapping_data = [];
            this.venus_manager_mapping_selected_details = [];
            this.venus_manager_mapping_table = false;
            this.get_venus_manager_detail();
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.selected_mapping_data = [];
          this.selected_vm_mapping_data = [];
          this.venus_manager_mapping_selected_details = [];
          this.venus_manager_mapping_table = false;
          this.get_venus_manager_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
  }

  manager_mapping_row_check(rowData: any) {
    const index: number = this.selected_mapping_data.indexOf(rowData);
    if (index != -1) {
      this.selected_mapping_data.splice(index, 1);
    }
    else {
      this.selected_mapping_data.push(rowData);
    }
  }

  manager_mapping_select_all(rowData: any) {
    const index: number = this.selected_mapping_data.indexOf(rowData);
    if (index != -1) {
      this.selected_mapping_data.splice(index, 1);
    }
    else {
      this.selected_mapping_data = rowData;
    }
  }

  view_venus_manager_detail(data: any) {
    this.selected_view_manager_mapping_details = data;
    this.get_venus_manager_mapping_detail(data);
  }

  get_venus_manager_mapping_detail(data: any) {
    if (data === undefined || data === null) {
      this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Please Select a Venus Manager first.' })
    }
    else {
      this.view_venus_manager_mapping_table = true;
      if (this.venus_manager_view_mapping_switched_table_list_name == "mt5_users") {
        this.ProgressSpinnerDlg = true;
        let userid = this.user_details.userid;
        let email = this.user_details.email;
        let managerid = data.managerid;
        const credentials = { "userid": userid, "email": email, "managerid": managerid };
        this.httpService.frontendRequestCall(EndPoints.get_venus_manager_mapping_detail, ApiMethod.POST, credentials)
          .subscribe(response => {
            response = response || {};
            let message = response.message || '';
            let status = response.status || '';
            if (response.status == "true") {
              this.ProgressSpinnerDlg = false;
              this.view_venus_manager_mapping_row_data = response.manager_mapping_data;
              this.view_venus_manager_mapping_col_header = Object.keys(this.view_venus_manager_mapping_row_data[0]);
              this.venus_manager_mapping_table = false;
            }
            else {
              this.ProgressSpinnerDlg = false;
              this.view_venus_manager_mapping_row_data = [];
              this.view_venus_manager_mapping_col_header = [];
              this.venus_manager_mapping_table = false;
              this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
            }
          }, error => {
            this.ProgressSpinnerDlg = false;
            this.view_venus_manager_mapping_row_data = [];
            this.view_venus_manager_mapping_col_header = [];
            this.venus_manager_mapping_table = false;
            this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
          }
          );
      }
      else if (this.venus_manager_view_mapping_switched_table_list_name == 'venus_group') {
        this.ProgressSpinnerDlg = true;
        let userid = this.user_details.userid;
        let email = this.user_details.email;
        let managerid = data.managerid;
        const credentials = { "userid": userid, "email": email, "managerid": managerid };
        this.httpService.frontendRequestCall(EndPoints.get_venus_manager_group_mapping_detail, ApiMethod.POST, credentials)
          .subscribe(response => {
            response = response || {};
            let message = response.message || '';
            let status = response.status || '';
            if (response.status == "true") {
              this.ProgressSpinnerDlg = false;
              this.view_venus_manager_mapping_row_data = response.manager_mapping_data;
              this.view_venus_manager_mapping_col_header = Object.keys(this.view_venus_manager_mapping_row_data[0]);
              this.venus_manager_mapping_table = false;
            }
            else {
              this.ProgressSpinnerDlg = false;
              this.view_venus_manager_mapping_row_data = [];
              this.view_venus_manager_mapping_col_header = [];
              this.venus_manager_mapping_table = false;
              this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
            }
          }, error => {
            this.ProgressSpinnerDlg = false;
            this.view_venus_manager_mapping_row_data = [];
            this.view_venus_manager_mapping_col_header = [];
            this.venus_manager_mapping_table = false;
            this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
          }
          );
      }
    }
  }

  delete_venus_manager_mapped_data(data: any) {
    this.selected_deletable_mapped_row_data = data;
    this.venus_manager_mapped_data_delete_dialog = true;
  }

  cancel_venus_manager_mapped_data_delete() {
    this.venus_manager_mapped_data_delete_dialog = false;
  }

  confirm_venus_manager_mapped_data_delete() {
    this.venus_manager_mapped_data_delete_dialog = false;
    if (this.venus_manager_view_mapping_switched_table_list_name == "mt5_users") {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let mappingid = this.selected_deletable_mapped_row_data.id;
      const credentials = { "userid": userid, "email": email, "mappingid": mappingid };
      this.httpService.frontendRequestCall(EndPoints.delete_user_from_venus_manager_mapping, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || 'false';
          if (status == "true") {
            this.ProgressSpinnerDlg = false;
            this.get_venus_manager_mapping_detail(this.selected_view_manager_mapping_details);
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });

          }
          else {
            this.ProgressSpinnerDlg = false;
            this.get_venus_manager_mapping_detail(this.selected_view_manager_mapping_details);
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        },
          error => {
            this.ProgressSpinnerDlg = false;
            this.get_venus_manager_mapping_detail(this.selected_view_manager_mapping_details);
            this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
          })
    }
    else if (this.venus_manager_view_mapping_switched_table_list_name == 'venus_group') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let mappingid = this.selected_deletable_mapped_row_data.id;
      const credentials = { "userid": userid, "email": email, "mappingid": mappingid };
      this.httpService.frontendRequestCall(EndPoints.delete_group_from_venus_manager_group_mapping, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || 'false';
          if (status == "true") {
            this.ProgressSpinnerDlg = false;
            this.get_venus_manager_mapping_detail(this.selected_view_manager_mapping_details);
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });

          }
          else {
            this.ProgressSpinnerDlg = false;
            this.get_venus_manager_mapping_detail(this.selected_view_manager_mapping_details);
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        },
          error => {
            this.ProgressSpinnerDlg = false;
            this.get_venus_manager_mapping_detail(this.selected_view_manager_mapping_details);
            this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
          })
    }
  }

  onExpSelect(event: any) {
    let d = new Date(Date.parse(event));
    let m = d.getMonth() + 1;
    let day = d.getDate();
    this.year = d.getFullYear();

    if (m < 10) {
      this.month = "0" + m;
    }
    else {
      this.month = m;
    }

    if (day < 10) {
      this.date = "0" + day;
    }
    else {
      this.date = day;
    }
    this.expdate = this.date + '/' + this.month + '/' + this.year;
  }

  // ************************************** VENUS GROUP METHODS *******************************
  open_new_venus_group_config() {
    this.add_new_venus_group_dialog = true;
  }

  venus_group_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.venus_group_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "venus_group_details");
    });
  }

  map_venus_group_detail(data: any) {
    this.selected_venus_group_data = data;
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
          this.venus_group_mapping_row_data = response.manager_detail;
          this.venus_group_mapping_col_header = Object.keys(this.venus_group_mapping_row_data[0]);
          this.venus_group_mapping_table = true;
          this.view_venus_group_mapping_table = false;
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.venus_group_mapping_row_data = [];
          this.venus_group_mapping_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          this.venus_group_mapping_table = false;
          this.view_venus_group_mapping_table = false;
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.venus_group_mapping_row_data = [];
        this.venus_group_mapping_col_header = [];
        this.venus_group_mapping_table = false;
        this.view_venus_group_mapping_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  group_mapping_row_check(rowData: any) {
    const index: number = this.selected_group_mapping_data.indexOf(rowData);
    if (index != -1) {
      this.selected_group_mapping_data.splice(index, 1);
    }
    else {
      this.selected_group_mapping_data.push(rowData);
    }
  }

  group_mapping_select_all(rowData: any) {
    const index: number = this.selected_group_mapping_data.indexOf(rowData);
    if (index != -1) {
      this.selected_group_mapping_data.splice(index, 1);
    }
    else {
      this.selected_group_mapping_data = rowData;
    }
  }

  save_venus_group_mapping_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    for (let i = 0; i < this.selected_group_mapping_data.length; i++) {
      let mapping_detail =
      {
        "groupid": this.selected_venus_group_data.groupid,
        "loginid": this.selected_group_mapping_data[i].Login
      }
      this.venus_group_mapping_selected_details.push(mapping_detail);
    }
    let mappingdetail = this.venus_group_mapping_selected_details;
    const credentials = { "userid": userid, "email": email, "mappingdetail": mappingdetail }
    this.httpService.frontendRequestCall(EndPoints.save_venus_group_mapping_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.selected_mapping_data = [];
          this.selected_vg_mapping_data = [];
          this.venus_group_mapping_selected_details = [];
          this.venus_group_mapping_table = false;
          this.selected_group_mapping_data = [];
          this.selected_venus_group_data = [];
          this.get_venus_group_detail();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.selected_mapping_data = [];
          this.selected_vg_mapping_data = [];
          this.venus_group_mapping_selected_details = [];
          this.venus_group_mapping_table = false;
          this.get_venus_group_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.selected_mapping_data = [];
        this.selected_vg_mapping_data = [];
        this.venus_group_mapping_selected_details = [];
        this.venus_group_mapping_table = false;
        this.get_venus_group_detail();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  venus_group_mapping_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.venus_group_mapping_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "Venus_group_mapped_data");
    });
  }

  view_venus_group_detail(data: any) {
    this.selected_view_group_mapping_details = data;
    this.get_venus_group_mapping_detail(data);
  }

  get_venus_group_mapping_detail(data: any) {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let groupid = data.groupid;
    const credentials = { "userid": userid, "email": email, "groupid": groupid };
    this.httpService.frontendRequestCall(EndPoints.get_venus_group_mapping_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.view_venus_group_mapping_row_data = response.group_mapping_data;
          this.view_venus_group_mapping_col_header = Object.keys(this.view_venus_group_mapping_row_data[0]);
          this.view_venus_group_mapping_table = true;
          this.venus_group_mapping_table = false;
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.view_venus_group_mapping_row_data = [];
          this.view_venus_group_mapping_col_header = [];
          this.view_venus_group_mapping_table = false;
          this.venus_group_mapping_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.view_venus_group_mapping_row_data = [];
        this.view_venus_group_mapping_col_header = [];
        this.view_venus_group_mapping_table = false;
        this.venus_group_mapping_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  delete_venus_group_mapped_data(data: any) {
    this.selected_deletable_group_mapped_row_data = data;
    this.venus_group_mapped_data_delete_dialog = true;
  }

  cancel_venus_group_mapped_data_delete() {
    this.venus_group_mapped_data_delete_dialog = false;
  }

  confirm_venus_group_mapped_data_delete() {
    this.venus_group_mapped_data_delete_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let mappingid = this.selected_deletable_group_mapped_row_data.id;
    const credentials = { "userid": userid, "email": email, "mappingid": mappingid };
    this.httpService.frontendRequestCall(EndPoints.delete_user_from_venus_group_mapping, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || 'false';
        if (status == "true") {
          this.ProgressSpinnerDlg = false;
          this.get_venus_group_mapping_detail(this.selected_view_group_mapping_details);
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });

        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_venus_group_mapping_detail(this.selected_view_group_mapping_details);
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      },
        error => {
          this.ProgressSpinnerDlg = false;
          this.get_venus_group_mapping_detail(this.selected_view_group_mapping_details);
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
        })
  }

  cancel_new_venus_group_charges() {
    this.add_new_venus_group_dialog = false;
    this.venus_group_form.reset();
  }

  confirm_new_venus_group_charges() {
    this.add_new_venus_group_dialog = false;
    if (this.venus_group_form.invalid) {
      this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Please enter all required details.' });
    }
    else {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let master = this.venus_group_form.controls.group_master.value;
      let name = this.venus_group_form.controls.group_name.value;
      let desc = this.venus_group_form.controls.group_desc.value;
      const credentials = { "userid": userid, "email": email, "master": master, "name": name, "desc": desc };
      this.httpService.frontendRequestCall(EndPoints.save_venus_group_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.venus_group_form.reset();
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
            this.get_venus_group_detail();
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.get_venus_group_detail();
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.get_venus_group_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
  }

  get_venus_group_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_venus_group_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.venus_group_row_data = response.group_data;
          this.venus_group_col_header = Object.keys(this.venus_group_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.venus_group_row_data = [];
          this.venus_group_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.venus_group_row_data = [];
        this.venus_group_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  update_selected_venus_group_data(data: any) {
    this.update_venus_group_dialog = true;
    this.selected_updatable_venus_group_data = data;
    this.update_venus_group_form.patchValue({
      group_master: data.master,
      group_name: data.name,
      group_desc: data.desc
    })
  }

  cancel_venus_group_updation() {
    this.update_venus_group_dialog = false;
    this.update_venus_group_form.reset();
  }

  confirm_venus_group_updation() {
    this.update_venus_group_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let groupid = this.selected_updatable_venus_group_data.groupid;
    let master = this.update_venus_group_form.controls.group_master.value;
    let name = this.update_venus_group_form.controls.group_name.value;
    let desc = this.update_venus_group_form.controls.group_desc.value;
    const credentials = { "userid": userid, "email": email, "groupid": groupid, "master": master, "name": name, "desc": desc }
    this.httpService.frontendRequestCall(EndPoints.update_venus_group_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.update_venus_group_form.reset();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          this.get_venus_group_detail();
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_venus_group_detail();
          this.update_venus_group_form.reset();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.get_venus_group_detail();
        this.update_venus_group_form.reset();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  delete_venus_group_data(data: any) {
    this.venus_group_delete_dialog = true;
    this.selected_deletable_venus_group_row_data = data;
  }

  cancel_venus_group_delete() {
    this.venus_group_delete_dialog = false;
  }

  confirm_venus_group_delete() {
    this.venus_group_delete_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let groupid = this.selected_deletable_venus_group_row_data.groupid;
    const credentials = { "userid": userid, "email": email, "groupid": groupid };
    this.httpService.frontendRequestCall(EndPoints.delete_venus_group_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || 'false';
        if (status == "true") {
          this.ProgressSpinnerDlg = false;
          this.get_venus_group_detail();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });

        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_venus_group_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      },
        error => {
          this.ProgressSpinnerDlg = false;
          this.get_venus_group_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
        })
  }

  // ************************************** BOT ACCESS METHODS ********************************
  bot_access_table_switched(event: any) {
    this.bot_access_switched_table_list_name = event.option.value;
    this.view_bot_access_details_table = false;
    if (this.bot_access_switched_table_list_name == 'mt5_manager') {
      this.bot_access_mt5_manager();
    }
    else if (this.bot_access_switched_table_list_name == 'venus_manager') {
      this.bot_access_venus_manager();
    }
  }

  bot_access_mt5_manager() {
    this.ProgressSpinnerDlg = true;
    this.bot_access_row_data = [];
    this.bot_access_col_header = [];
    this.selected_bot_access = [];
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
          this.bot_access_row_data = response.manager_detail;
          this.bot_access_col_header = Object.keys(this.bot_access_row_data[0]);
          this.bot_access_mt5_manager_table = true;
          this.bot_access_venus_manager_table = false;
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.bot_access_row_data = [];
          this.bot_access_col_header = [];
          this.bot_access_mt5_manager_table = false;
          this.bot_access_venus_manager_table = false;
          this.bot_access_switched_table_list_name = 'venus_manager';
          this.bot_access_venus_manager();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.bot_access_row_data = [];
        this.bot_access_col_header = [];
        this.bot_access_mt5_manager_table = false;
        this.bot_access_venus_manager_table = false;
        this.bot_access_switched_table_list_name = 'venus_manager';
        this.bot_access_venus_manager();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  bot_access_venus_manager() {
    this.ProgressSpinnerDlg = true;
    this.bot_access_row_data = [];
    this.bot_access_col_header = [];
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
          this.bot_access_row_data = response.manager_data;
          this.bot_access_col_header = Object.keys(this.bot_access_row_data[0]);
          this.bot_access_mt5_manager_table = false;
          this.bot_access_venus_manager_table = true;
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.bot_access_row_data = [];
          this.bot_access_col_header = [];
          this.bot_access_mt5_manager_table = false;
          this.bot_access_venus_manager_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          this.bot_access_switched_table_list_name = 'mt5_manager';
          this.bot_access_mt5_manager();
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.bot_access_row_data = [];
        this.bot_access_col_header = [];
        this.bot_access_mt5_manager_table = false;
        this.bot_access_venus_manager_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        this.bot_access_switched_table_list_name = 'mt5_manager';
        this.bot_access_mt5_manager();
      }
      );
  }
  view_bot_access_detail(data: any) {
    this.get_mt5_manager_bot_access_detail(data);
  }

  get_mt5_manager_bot_access_detail(data: any) {
    this.selected_bot_access_manager_details = data;
    if (this.bot_access_switched_table_list_name == 'mt5_manager') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let managerloginid = data.Login;
      const credentials = { "userid": userid, "email": email, "managerloginid": managerloginid };
      this.httpService.frontendRequestCall(EndPoints.get_mt5_manager_bot_access_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.bot_access_details_row_data = response.manager_data;
            this.bot_access_details_col_header = Object.keys(this.bot_access_details_row_data[0]);
            this.view_bot_access_details_table = true;
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.bot_access_details_row_data = [];
            this.bot_access_details_col_header = [];
            this.view_bot_access_details_table = false;
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.bot_access_details_row_data = [];
          this.bot_access_details_col_header = [];
          this.view_bot_access_details_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
    else if (this.bot_access_switched_table_list_name == 'venus_manager') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let managerid = data.managerid;
      const credentials = { "userid": userid, "email": email, "managerid": managerid };
      this.httpService.frontendRequestCall(EndPoints.get_venus_manager_bot_access_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.bot_access_details_row_data = response.manager_data;
            this.bot_access_details_col_header = Object.keys(this.bot_access_details_row_data[0]);
            this.view_bot_access_details_table = true;
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.bot_access_details_row_data = [];
            this.bot_access_details_col_header = [];
            this.view_bot_access_details_table = false;
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.bot_access_details_row_data = [];
          this.bot_access_details_col_header = [];
          this.view_bot_access_details_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }

  }

  bot_access_check_checked(data: any) {
    const index: number = this.bot_access_selected_check_value.indexOf(data);
    if (index != -1) {
      this.bot_access_selected_check_value.splice(index, 1);
    }
    else {
      this.bot_access_selected_check_value.push(data);
    }
  }

  on_bot_type_select(event: any) {
    this.selected_bot_access_type = event.value;
  }

  new_bot_access_detail() {
    this.add_new_bot_access_dialog = true;
    this.bot_access_form.reset();
  }

  cancel_new_bot_access_charges() {
    this.add_new_bot_access_dialog = false;
    this.bot_access_form.reset();
  }

  confirm_new_bot_access_charges() {
    this.add_new_bot_access_dialog = false;
    this.ProgressSpinnerDlg = true;
    if (this.bot_access_switched_table_list_name == 'mt5_manager') {
      for (let i = 0; i < this.bot_access_selected_check_value.length; i++) {
        for (let j = 0; j < this.selected_bot_access_type.length; j++) {
          let accessdetails = {
            "Login": this.bot_access_selected_check_value[i].Login,
            "bottype": this.selected_bot_access_type[j],
          }
          this.access_detail.push(accessdetails);
        }

      }
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let accessdetail = this.access_detail;
      const credentials = { "userid": userid, "email": email, "accessdetail": accessdetail };
      this.httpService.frontendRequestCall(EndPoints.save_mt5_manager_bot_access_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.bot_access_form.reset();
            if (this.bot_access_switched_table_list_name == 'mt5_manager') {
              this.bot_access_mt5_manager();
            }
            else if (this.bot_access_switched_table_list_name == 'venus_manager') {
              this.bot_access_venus_manager();
            }
            this.selected_bot_access = [];
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false
            this.bot_access_form.reset();
            if (this.bot_access_switched_table_list_name == 'mt5_manager') {
              this.bot_access_mt5_manager();
            }
            else if (this.bot_access_switched_table_list_name == 'venus_manager') {
              this.bot_access_venus_manager();
            }
            this.selected_bot_access = [];
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.bot_access_form.reset();
          if (this.bot_access_switched_table_list_name == 'mt5_manager') {
            this.bot_access_mt5_manager();
          }
          else if (this.bot_access_switched_table_list_name == 'venus_manager') {
            this.bot_access_venus_manager();
          }
          this.selected_bot_access = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
      this.selected_bot_access = [];
    }
    else if (this.bot_access_switched_table_list_name == 'venus_manager') {
      for (let i = 0; i < this.bot_access_selected_check_value.length; i++) {
        for (let j = 0; j < this.selected_bot_access_type.length; j++) {
          let accessdetails = {
            "managerid": this.bot_access_selected_check_value[i].managerid,
            "bottype": this.selected_bot_access_type[j],
          }
          this.access_detail.push(accessdetails);
        }
      }
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let accessdetail = this.access_detail;
      const credentials = { "userid": userid, "email": email, "accessdetail": accessdetail };
      this.httpService.frontendRequestCall(EndPoints.save_venus_manager_bot_access_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.bot_access_form.reset();
            if (this.bot_access_switched_table_list_name == 'mt5_manager') {
              this.bot_access_mt5_manager();
            }
            else if (this.bot_access_switched_table_list_name == 'venus_manager') {
              this.bot_access_venus_manager();
            }
            this.selected_bot_access = [];
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false
            this.bot_access_form.reset();
            if (this.bot_access_switched_table_list_name == 'mt5_manager') {
              this.bot_access_mt5_manager();
            }
            else if (this.bot_access_switched_table_list_name == 'venus_manager') {
              this.bot_access_venus_manager();
            }
            this.selected_bot_access = [];
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.bot_access_form.reset();
          if (this.bot_access_switched_table_list_name == 'mt5_manager') {
            this.bot_access_mt5_manager();
          }
          else if (this.bot_access_switched_table_list_name == 'venus_manager') {
            this.bot_access_venus_manager();
          }
          this.selected_bot_access = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
    this.selected_bot_access = [];
    this.selected_manager_group_data = [];
    this.bot_access_selected_check_value = [];
    this.access_detail = [];
  }

  update_bot_access_data(data: any) {
    this.update_bot_access_dialog = true;
    this.selected_bot_access_data_for_update = data;
  }

  cancel_bot_access_updation() {
    this.update_bot_access_dialog = false;
    this.update_bot_access_form.reset();
  }

  confirm_bot_access_updation() {
    this.update_bot_access_dialog = false;
    if (this.bot_access_switched_table_list_name == 'mt5_manager') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let accessid = this.selected_bot_access_data_for_update.accessid;
      let Login = this.selected_bot_access_data_for_update.Login;
      let bottype = this.update_bot_access_form.controls.bot_type.value;
      const credentials = { "userid": userid, "email": email, "accessid": accessid, "Login": Login, "bottype": bottype };
      this.httpService.frontendRequestCall(EndPoints.update_mt5_manager_bot_access_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
            this.update_bot_access_form.reset();
            this.selected_bot_access = [];
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
            this.update_bot_access_form.reset();
            this.selected_bot_access = [];
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
          this.update_bot_access_form.reset();
          this.selected_bot_access = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
    else if (this.bot_access_switched_table_list_name == 'venus_manager') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let accessid = this.selected_bot_access_data_for_update.accessid;
      let managerid = this.selected_bot_access_data_for_update.managerid;
      let bottype = this.update_bot_access_form.controls.bot_type.value;
      const credentials = { "userid": userid, "email": email, "accessid": accessid, "managerid": managerid, "bottype": bottype };
      this.httpService.frontendRequestCall(EndPoints.update_venus_manager_bot_access_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
            this.update_bot_access_form.reset();
            this.selected_bot_access = [];
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
            this.update_bot_access_form.reset();
            this.selected_bot_access = [];
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
          this.update_bot_access_form.reset();
          this.selected_bot_access = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }

    this.selected_bot_access = [];
  }

  delete_bot_access_data(data: any) {
    this.selected_deletable_bot_access_row_data = data;
    this.bot_access_delete_dialog = true;
  }

  cancel_bot_access_delete() {
    this.bot_access_delete_dialog = false;
    this.selected_deletable_bot_access_row_data = [];
  }

  confirm_bot_access_delete() {
    this.bot_access_delete_dialog = false;
    if (this.bot_access_switched_table_list_name == 'mt5_manager') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let accessid = this.selected_deletable_bot_access_row_data.accessid;
      const credentials = { "userid": userid, "email": email, "accessid": accessid };
      this.httpService.frontendRequestCall(EndPoints.delete_mt5_manager_bot_access_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || 'false';
          if (status == "true") {
            this.ProgressSpinnerDlg = false;
            this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });

          }
          else {
            this.ProgressSpinnerDlg = false;
            this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        },
          error => {
            this.ProgressSpinnerDlg = false;
            this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
            this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
          })
    }
    else if (this.bot_access_switched_table_list_name == 'venus_manager') {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let accessid = this.selected_deletable_bot_access_row_data.accessid;
      const credentials = { "userid": userid, "email": email, "accessid": accessid };
      this.httpService.frontendRequestCall(EndPoints.delete_venus_manager_bot_access_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || 'false';
          if (status == "true") {
            this.ProgressSpinnerDlg = false;
            this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });

          }
          else {
            this.ProgressSpinnerDlg = false;
            this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        },
          error => {
            this.ProgressSpinnerDlg = false;
            this.get_mt5_manager_bot_access_detail(this.selected_bot_access_manager_details);
            this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
          })
    }
  }

  bot_access_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.bot_access_details_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "bot_access_details");
    });
  }

  // ************************************** COMMAND NAME METHODS ********************************
  open_new_command_name_config() {
    this.command_name_form.reset();
    this.add_new_command_name_dialog = true;
  }

  cancel_new_command_name_charges() {
    this.add_new_command_name_dialog = false;
    this.command_name_form.reset();
  }

  confirm_new_command_name_charges() {
    this.add_new_command_name_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let commandtype = this.command_name_form.controls.command_types.value;
    let commandname = this.command_name_form.controls.command_name.value;
    let commanddesc = this.command_name_form.controls.command_desc.value;
    let commissiontype = this.command_name_form.controls.commission_type.value;
    const credentials = { "userid": userid, "email": email, "commandtype": commandtype, "commandname": commandname, "commanddesc": commanddesc, "commissiontype": commissiontype };
    this.httpService.frontendRequestCall(EndPoints.save_venus_command_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || 'false';
        if (status == "true") {
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_detail();
          this.command_name_form.reset();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      },
        error => {
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
        })
  }

  get_venus_command_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_venus_command_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.command_name_row_data = response.command_data;
          this.command_name_col_header = Object.keys(this.command_name_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.command_name_row_data = [];
          this.command_name_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.command_name_row_data = [];
        this.command_name_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  delete_venus_command_name_data(data: any) {
    this.selected_deletable_command_name_data = data;
    this.venus_command_name_delete_dialog = true;
  }

  cancel_venus_command_name_delete() {
    this.venus_command_name_delete_dialog = false;
  }

  confirm_venus_command_name_delete() {
    this.venus_command_name_delete_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let commandid = this.selected_deletable_command_name_data.commandid;
    const credentials = { "userid": userid, "email": email, "commandid": commandid };
    this.httpService.frontendRequestCall(EndPoints.delete_venus_command_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || 'false';
        if (status == "true") {
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_detail();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });

        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      },
        error => {
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
        })
  }

  update_venus_cmd_name_data(data: any) {
    this.update_venus_cmd_dialog = true;
    this.selected_updatable_venus_cmd_data = data;
    this.selected_commission_type = data.commission_type;
    this.selected_cmd_no = data.command_seq_no;
  }

  cancel_venus_cmd_updation() {
    this.update_venus_cmd_dialog = false;
    this.update_venus_cmd_form.reset();
  }

  confirm_venus_cmd_updation() {
    this.update_venus_cmd_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let commandid = this.selected_updatable_venus_cmd_data.commandid;
    let commissiontype = this.selected_commission_type;
    let commandseqno = this.selected_cmd_no;
    const credentials = { "userid": userid, "email": email, "commandid": commandid, "commissiontype": commissiontype, "commandseqno": commandseqno };
    this.httpService.frontendRequestCall(EndPoints.update_venus_command_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.update_venus_cmd_form.reset();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
          this.get_venus_command_detail();
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_detail();
          this.update_venus_cmd_form.reset();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.get_venus_command_detail();
        this.update_venus_cmd_form.reset();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  command_name_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.command_name_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "command_names_data");
    });
  }

  // ************************************** COMMAND MAPPING METHODS *****************************
  get_venus_command_list_for_mapping() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_venus_command_list_for_mapping, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.venus_command_mapping_row_data = response.command_data;
          this.venus_command_mapping_col_header = Object.keys(this.venus_command_mapping_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.venus_command_mapping_row_data = [];
          this.venus_command_mapping_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.venus_command_mapping_row_data = [];
        this.venus_command_mapping_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  get_mt5_command_list_for_mapping() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_mt5_command_list_for_mapping, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.venus_command_mapping_row_data = response.command_data;
          this.venus_command_mapping_col_header = Object.keys(this.venus_command_mapping_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.venus_command_mapping_row_data = [];
          this.venus_command_mapping_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.venus_command_mapping_row_data = [];
        this.venus_command_mapping_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  group_names_mapping_table_switched(event: any) {
    this.selected_manager_group_data = [];
    this.selected_command_name_data = [];
    this.group_names_switched_table_list_name = event.option.value;
    if (this.group_names_switched_table_list_name == 'mt5_manager') {
      this.get_mt5_group_mapping_detail();
    }
    else if (this.group_names_switched_table_list_name == 'venus_manager') {
      this.get_venus_mapping_group_detail();
    }
  }

  get_mt5_group_mapping_detail() {
    this.ProgressSpinnerDlg = true;
    this.command_mapping_row_data = [];
    this.command_mapping_col_header = [];
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
          this.command_mapping_row_data = response.manager_detail;
          this.command_mapping_col_header = Object.keys(this.command_mapping_row_data[0]);
          this.mt5_group_name_mapping_table = true;
          this.venus_manager_name_mapping_table = false;
          this.get_mt5_command_list_for_mapping();
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.command_mapping_row_data = [];
          this.command_mapping_col_header = [];
          this.venus_cgm_row_data = [];
          this.venus_cgm_col_header = [];
          this.venus_manager_name_mapping_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.command_mapping_row_data = [];
        this.command_mapping_col_header = [];
        this.venus_cgm_row_data = [];
        this.venus_cgm_col_header = [];
        this.venus_manager_name_mapping_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  get_venus_mapping_group_detail() {
    this.ProgressSpinnerDlg = true;
    this.command_mapping_row_data = [];
    this.command_mapping_col_header = [];
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
          this.command_mapping_row_data = response.manager_data;
          this.command_mapping_col_header = Object.keys(this.command_mapping_row_data[0]);
          this.venus_manager_name_mapping_table = true;
          this.mt5_group_name_mapping_table = false;
          this.get_venus_command_list_for_mapping();
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.command_mapping_row_data = [];
          this.command_mapping_col_header = [];
          this.venus_cgm_row_data = [];
          this.venus_cgm_col_header = [];
          this.venus_manager_name_mapping_table = false;
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.command_mapping_row_data = [];
        this.command_mapping_col_header = [];
        this.venus_cgm_row_data = [];
        this.venus_cgm_col_header = [];
        this.venus_manager_name_mapping_table = false;
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  command_name_mapping_check_checked(data: any) {
    const index: number = this.cmd_name_mapping_check_value.indexOf(data);
    if (index != -1) {
      this.cmd_name_mapping_check_value.splice(index, 1);
    }
    else {
      this.cmd_name_mapping_check_value.push(data);
    }
  }

  command_name_mapping_check_select_all(data: any) {
    const index: number = this.cmd_name_mapping_check_value.indexOf(data);
    if (index != -1) {
      this.cmd_name_mapping_check_value.splice(index, 1);
    }
    else {
      this.cmd_name_mapping_check_value = data;
    }
  }

  command_group_mapping_check_checked(data: any) {
    const index: number = this.mapping_selected_check_value.indexOf(data);
    if (index != -1) {
      this.mapping_selected_check_value.splice(index, 1);
    }
    else {
      this.mapping_selected_check_value.push(data);
    }
  }

  save_venus_command_mapping_detail() {
    for (let i = 0; i < this.cmd_name_mapping_check_value.length; i++) {
      for (let j = 0; j < this.mapping_selected_check_value.length; j++) {
        let manager_id, manager_type;
        if (this.group_names_switched_table_list_name == 'mt5_manager') {
          manager_id = this.mapping_selected_check_value[j].Login;
          manager_type = "mt5";
        }
        else if (this.group_names_switched_table_list_name == 'venus_manager') {
          manager_id = this.mapping_selected_check_value[j].managerid;
          manager_type = "venus";
        }

        let mapping_detail = {
          "command_type": this.cmd_name_mapping_check_value[i].command_type,
          "commandid": this.cmd_name_mapping_check_value[i].commandid,
          "managerid": manager_id,
          "managertype": manager_type,
          "method": this.selected_method["value"]
        }
        this.mapping_details_list.push(mapping_detail);
      }
    }

    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let mappingdetail = this.mapping_details_list;
    const credentials = { "userid": userid, "email": email, "mappingdetail": mappingdetail };
    this.httpService.frontendRequestCall(EndPoints.save_venus_command_mapping_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.mapping_details_list = [];
          this.cmd_name_mapping_check_value = [];
          this.mapping_selected_check_value = [];
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_list_for_mapping();
          this.get_venus_command_mapping_detail();
          if (this.group_names_switched_table_list_name == 'mt5_manager') {
            this.get_mt5_group_mapping_detail();
            this.get_mt5_command_list_for_mapping();
          }
          else if (this.group_names_switched_table_list_name == 'venus_manager') {
            this.get_venus_mapping_group_detail();
            this.get_venus_command_list_for_mapping();
          }
          this.selected_manager_group_data = [];
          this.selected_command_name_data = [];
          this.selected_method = "";
          this.selected_method = undefined;
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.mapping_details_list = [];
          this.selected_manager_group_data = [];
          this.selected_command_name_data = [];
          this.get_venus_command_list_for_mapping();
          this.get_venus_command_mapping_detail();
          this.mapping_details_list = [];
          this.cmd_name_mapping_check_value = [];
          this.mapping_selected_check_value = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.mapping_details_list = [];
        this.selected_manager_group_data = [];
        this.selected_command_name_data = [];
        this.get_venus_command_list_for_mapping();
        this.get_venus_command_mapping_detail();
        this.mapping_details_list = [];
        this.cmd_name_mapping_check_value = [];
        this.mapping_selected_check_value = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  get_venus_command_mapping_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_venus_command_mapping_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.venus_cgm_row_data = response.command_mapping_data;
          this.venus_cgm_col_header = Object.keys(this.venus_cgm_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.venus_cgm_row_data = [];
          this.venus_cgm_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.venus_cgm_row_data = [];
        this.venus_cgm_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  delete_venus_cgm_data(data: any) {
    this.selected_deletable_cgm_data = data;
    this.venus_cgm_delete_dialog = true;
  }

  cancel_venus_cgm_delete() {
    this.venus_cgm_delete_dialog = false;
  }

  confirm_venus_cgm_delete() {
    this.venus_cgm_delete_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let mappingid = this.selected_deletable_cgm_data.id;
    const credentials = { "userid": userid, "email": email, "mappingid": mappingid };
    this.httpService.frontendRequestCall(EndPoints.delete_manager_from_venus_command_mapping, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || 'false';
        if (status == "true") {
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_list_for_mapping();
          this.get_venus_command_mapping_detail();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });

        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_mapping_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      },
        error => {
          this.ProgressSpinnerDlg = false;
          this.get_venus_command_mapping_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
        })
  }
  command_mapping_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.command_mapping_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "command_mapping_data");
    });
  }

  venus_cgm_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.venus_cgm_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "command_cgm_data");
    });
  }

}
