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
  selector: 'app-symbol',
  templateUrl: './symbol.component.html',
  styleUrls: ['./symbol.component.scss'],
  providers: [MessageService]
})
export class SymbolComponent implements OnInit {
  // Global Variables
  ProgressSpinnerDlg = false;
  user_details: any;

  // Symbol Variables
  symbol_row_data: Array<any> = [];
  symbol_col_header: Array<any> = [];
  symbol_catg_list: Array<any> = [];
  add_new_symbol_dialog: boolean = false;
  update_symbol_dialog: boolean = false;
  symbol_delete_dialog: boolean = false;
  selected_deletable_symbol_row_data: any;
  selected_symbol_data_for_update: any;
  symbol_form!: FormGroup;
  update_symbol_form!: FormGroup;

  constructor(
    private messageService: MessageService,
    private storageService: StorageService,
    private httpService: HttpService,
    private router: Router,
    private formBuilder: FormBuilder
  ) { }

  ngOnInit(): void {
    this.user_details = this.storageService.getLocalObject("userdetails");
    this.get_symbol_category_detail();
    this.symbol_form = this.formBuilder.group({
      symbol_name: ['', Validators.required],
      symbol_catg: ['', Validators.required]
    })
    this.symbol_form.reset();
    this.symbol_catg_list = ["MCX", "NSE", "CMX"];

    this.update_symbol_form = this.formBuilder.group({
      symbol_catg: ['', Validators.required]
    })
  }

  // ********************************* Global Methods ********************************

  onTabChange(event: { index: number; }) {
    if (event.index == 0) {
      this.get_symbol_category_detail();
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

  symbol_export() {
    import("xlsx").then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.symbol_row_data);
      const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
      const csvBuffer: any = xlsx.write(workbook, { bookType: 'csv', type: 'array' });
      this.saveAsCSVFile(csvBuffer, "symbol_data");
    });
  }

  // *************************** SYMBOL METHODS ***************************
  get_symbol_category_detail() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    const credentials = { "userid": userid, "email": email };
    this.httpService.frontendRequestCall(EndPoints.get_symbol_category_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.symbol_row_data = response.symbol_data;
          this.symbol_col_header = Object.keys(this.symbol_row_data[0]);
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.symbol_row_data = [];
          this.symbol_col_header = [];
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.symbol_row_data = [];
        this.symbol_col_header = [];
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

  new_symbol_category_detail() {
    this.add_new_symbol_dialog = true;
    this.symbol_form.reset();
  }

  cancel_new_symbol_charges() {
    this.add_new_symbol_dialog = false;
    this.symbol_form.reset();
  }

  confirm_new_symbol_charges() {
    this.add_new_symbol_dialog = false;
    if (this.symbol_form.invalid) {
      this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Please enter all required details.' });
    }
    else {
      this.ProgressSpinnerDlg = true;
      let userid = this.user_details.userid;
      let email = this.user_details.email;
      let symbolname = this.symbol_form.controls.symbol_name.value.toUpperCase();
      let symbolcategory = this.symbol_form.controls.symbol_catg.value;
      const credentials = { "userid": userid, "email": email, "symbolname": symbolname, "symbolcategory": symbolcategory };
      this.httpService.frontendRequestCall(EndPoints.save_symbol_category_detail, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || '';
          if (response.status == "true") {
            this.ProgressSpinnerDlg = false;
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
            this.get_symbol_category_detail();
          }
          else {
            this.ProgressSpinnerDlg = false;
            this.get_symbol_category_detail();
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.get_symbol_category_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
        }
        );
    }
  }

  delete_symbol_data(data: any) {
    this.selected_deletable_symbol_row_data = data;
    this.symbol_delete_dialog = true;
  }

  cancel_symbol_delete() {
    this.symbol_delete_dialog = false;
    this.selected_deletable_symbol_row_data = [];
  }

  confirm_symbol_delete() {
    this.symbol_delete_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let symbolid = this.selected_deletable_symbol_row_data.symbolid;
    const credentials = { "userid": userid, "email": email, "symbolid": symbolid };
    this.httpService.frontendRequestCall(EndPoints.delete_symbol_category_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || 'false';
        if (status == "true") {
          this.ProgressSpinnerDlg = false;
          this.get_symbol_category_detail();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });

        }
        else {
          this.ProgressSpinnerDlg = false;
          this.get_symbol_category_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      },
        error => {
          this.ProgressSpinnerDlg = false;
          this.get_symbol_category_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong' });
        })
  }

  update_symbol_data(data: any) {
    this.update_symbol_dialog = true;
    this.selected_symbol_data_for_update = data;
  }

  cancel_symbol_updation() {
    this.update_symbol_dialog = false;
    this.update_symbol_form.reset();
  }

  confirm_symbol_updation() {
    this.update_symbol_dialog = false;
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let symbolid = this.selected_symbol_data_for_update.symbolid;
    let symbolname = this.selected_symbol_data_for_update.symbol;
    let symbolcategory = this.update_symbol_form.controls.symbol_catg.value;
    const credentials = { "userid": userid, "email": email, "symbolid": symbolid, "symbolname": symbolname, "symbolcategory": symbolcategory };
    this.httpService.frontendRequestCall(EndPoints.update_symbol_category_detail, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || '';
        if (response.status == "true") {
          this.ProgressSpinnerDlg = false;
          this.update_symbol_form.reset();
          this.selected_symbol_data_for_update = [];
          this.get_symbol_category_detail();
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
        }
        else {
          this.ProgressSpinnerDlg = false
          this.update_symbol_form.reset();
          this.selected_symbol_data_for_update = [];
          this.get_symbol_category_detail();
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.update_symbol_form.reset();
        this.selected_symbol_data_for_update = [];
        this.get_symbol_category_detail();
        this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong. Please try again Later!' });
      }
      );
  }

}
