import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { MessageService } from 'primeng/api';
import * as FileSaver from 'file-saver';
import { Table } from 'primeng/table';
import { HttpService } from 'src/app/core/service/http.service';
import { StorageService } from 'src/app/core/service/storage.service';
import { ApiMethod, EndPoints } from 'src/app/core/const';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { S3UtilService } from 'src/app/core/service/s3-util.service';

@Component({
  selector: 'app-duplicate-ip',
  templateUrl: './duplicate-ip.component.html',
  styleUrls: ['./duplicate-ip.component.scss'],
  providers: [MessageService]
})
export class DuplicateIpComponent implements OnInit {
  // Global Variables
  ProgressSpinnerDlg = false;
  user_details: any;
  filetype = '.log';
  history_file_type = '.csv';
  trans_file_type = '.zip';

  // Duplicate IP Variables
  duplicate_ip_upload_dialog: boolean = false;
  duplicate_ip_file_name: any;
  duplicate_ip_files: any;
  duplicate_ip_file_upload: any;
  duplicate_ip_row_data: Array<any> = [];
  duplicate_ip_col_header: Array<any> = [];
  bucket_folder_url: any;

  // Summary Variables
  summary_csv_upload_dialog: boolean = false;
  summary_csv_file_name: any;
  summary_csv_files: any;
  summary_csv_file_upload: any;
  summary_csv_file_path: any;
  summary_zip_upload_dialog: boolean = false;
  summary_zip_file_name: any;
  summary_zip_files: any;
  summary_zip_file_upload: any;
  summary_zip_file_path: any;


  constructor(
    private messageService: MessageService,
    private storageService: StorageService,
    private httpService: HttpService,
    private router: Router,
    private formBuilder: FormBuilder,
    private s3UtilService: S3UtilService
  ) { }

  ngOnInit(): void {
    this.user_details = this.storageService.getLocalObject("userdetails");
    this.get_bucket_folder_name();
  }

  // ********************************* Global Methods ********************************
  onTabChange(event: { index: number; }) {
    if (event.index == 0) { }
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

  save_renamed_file(buffer: any, fileName: string): void {
    let CSV_TYPE = 'text/csv;charset=utf-8;';
    const data: Blob = new Blob([buffer], {
      type: CSV_TYPE
    });
    FileSaver.saveAs(data, fileName);
  }

  // ********************************* DUPLICATE IP METHODS ********************************
  cancel_duplicate_ip_upload() {
    this.duplicate_ip_upload_dialog = false;
    this.duplicate_ip_files = null;
  }

  async confirm_duplicate_ip_upload() {
    this.duplicate_ip_upload_dialog = false;
    this.ProgressSpinnerDlg = true;
    let file = this.duplicate_ip_files;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let isUpload = await this.s3UtilService.uploadfile(file, this.bucket_folder_url, file.name.replaceAll(' ', '_'));
    if (isUpload) {
      let filepath = this.bucket_folder_url + file.name.replaceAll(' ', '_');
      const credentials = { "userid": userid, "email": email, "filepath": filepath };
      this.httpService.frontendRequestCall(EndPoints.get_data_duplication_analysis, ApiMethod.POST, credentials)
        .subscribe(response => {
          response = response || {};
          let message = response.message || '';
          let status = response.status || 'false';
          if (status == "true") {
            this.ProgressSpinnerDlg = false;
            this.messageService.add({ severity: 'success', summary: 'Success', detail: message, life: 5000 });
            this.duplicate_ip_file_upload.clear();
            this.duplicate_ip_files = null;
            let dw_file = this.s3UtilService.getFileFromS3(response.file1path);
            let d_file = this.s3UtilService.getFileFromS3(response.file2path);
            FileSaver.saveAs(dw_file);
            FileSaver.saveAs(d_file);
            this.duplicate_ip_file_name = "";

          } else {
            this.ProgressSpinnerDlg = false;
            this.duplicate_ip_files = null;
            this.duplicate_ip_file_upload.clear();
            this.duplicate_ip_file_name = "";
            this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
          }
        }, error => {
          this.ProgressSpinnerDlg = false;
          this.duplicate_ip_files = null;
          this.duplicate_ip_file_name = "";
          this.messageService.add({ severity: 'error', summary: 'Something Went Wrong!', detail: error });
        });
    }
    else {
      this.ProgressSpinnerDlg = false;
      this.messageService.add({ severity: 'error', summary: 'Something Went Wrong!', detail: 'File Upload Failed' });
    }
  }

  duplicate_ip_Upload_File(event: any, DuplicateIpFile: any) {
    for (const file of event.files) {
      this.duplicate_ip_files = file;
    }
    this.duplicate_ip_file_name = this.duplicate_ip_files.name;
    this.duplicate_ip_upload_dialog = true;
    this.duplicate_ip_file_upload = DuplicateIpFile;
  }


  // ********************************** SUMMARY METHODS ********************************

  cancel_summary_csv_upload() {
    this.summary_csv_upload_dialog = false;
    this.summary_csv_files = null;
  }

  get_bucket_folder_name() {
    this.ProgressSpinnerDlg = true;
    this.httpService.frontendRequestCall(EndPoints.get_bucket_folder_name, ApiMethod.GET)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || 'false';
        if (status == "true") {
          this.ProgressSpinnerDlg = false;
          this.bucket_folder_url = response.bucket_folder_name;
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message, life: 5000 });
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.bucket_folder_url = "";
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.bucket_folder_url = "";
        this.messageService.add({ severity: 'error', summary: 'Something Went Wrong!', detail: error });
      });
  }

  async confirm_summary_csv_upload() {
    this.ProgressSpinnerDlg = true;
    this.summary_csv_upload_dialog = false;
    let file = this.summary_csv_files;
    let isUpload = await this.s3UtilService.uploadfile(file, this.bucket_folder_url, file.name.replaceAll(' ', '_'));
    if (isUpload) {
      this.ProgressSpinnerDlg = false;
      this.summary_csv_file_path = this.bucket_folder_url + file.name.replaceAll(' ', '_');
    }
    else {
      this.ProgressSpinnerDlg = false;
      this.messageService.add({ severity: 'error', summary: 'Something Went Wrong!', detail: 'File Upload Failed' });
    }
  }

  cancel_summary_zip_upload() {
    this.summary_zip_upload_dialog = false;
    this.summary_zip_files = null;
  }

  async confirm_summary_zip_upload() {
    this.ProgressSpinnerDlg = true;
    this.summary_zip_upload_dialog = false;
    let file = this.summary_zip_files;
    let isUpload = await this.s3UtilService.uploadfile(file, this.bucket_folder_url, file.name.replaceAll(' ', '_'));
    if (isUpload) {
      this.ProgressSpinnerDlg = false;
      this.summary_zip_file_path = this.bucket_folder_url + file.name.replaceAll(' ', '_');
    }
    else {
      this.ProgressSpinnerDlg = false;
      this.messageService.add({ severity: 'error', summary: 'Something Went Wrong!', detail: 'File Upload Failed' });
    }
  }

  summary_csv_Upload_File(event: any, summaryCSVFile: any) {
    for (const file of event.files) {
      this.summary_csv_files = file;
    }
    this.summary_csv_file_name = this.summary_csv_files.name;
    this.summary_csv_upload_dialog = true;
    this.summary_csv_file_upload = summaryCSVFile;
  }

  summary_zip_Upload_File(event: any, summaryZIPFile: any) {
    for (const file of event.files) {
      this.summary_zip_files = file;
    }
    this.summary_zip_file_name = this.summary_zip_files.name;
    this.summary_zip_upload_dialog = true;
    this.summary_zip_file_upload = summaryZIPFile;
  }

  save_summary_details() {
    this.ProgressSpinnerDlg = true;
    let userid = this.user_details.userid;
    let email = this.user_details.email;
    let file1path = this.summary_csv_file_path;
    let file2path = this.summary_zip_file_path;
    const credentials = { "userid": userid, "email": email, "file1path": file1path, "file2path": file2path };
    this.httpService.frontendRequestCall(EndPoints.generate_data_summary_analysis, ApiMethod.POST, credentials)
      .subscribe(response => {
        response = response || {};
        let message = response.message || '';
        let status = response.status || 'false';
        if (status == "true") {
          this.ProgressSpinnerDlg = false;
          let summarized_file = this.s3UtilService.getFileFromS3(response.filepath);
          FileSaver.saveAs(summarized_file);
          this.summary_csv_file_path = "";
          this.summary_zip_file_path = "";
          this.summary_csv_file_name = "";
          this.summary_zip_file_name = "";
          this.messageService.add({ severity: 'success', summary: 'Success', detail: message, life: 5000 });
        }
        else {
          this.ProgressSpinnerDlg = false;
          this.summary_csv_file_path = "";
          this.summary_zip_file_path = "";
          this.summary_csv_file_name = "";
          this.summary_zip_file_name = "";
          this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
        }
      }, error => {
        this.ProgressSpinnerDlg = false;
        this.summary_csv_file_path = "";
        this.summary_zip_file_path = "";
        this.summary_csv_file_name = "";
        this.summary_zip_file_name = "";
        this.messageService.add({ severity: 'error', summary: 'Something Went Wrong!', detail: error });
      });
  }

}
