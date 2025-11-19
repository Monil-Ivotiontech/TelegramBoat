import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { DuplicateIpRoutingModule } from './duplicate-ip-routing.module';
import { FileUploadModule } from 'primeng/fileupload';


@NgModule({
  declarations: [],
  imports: [
    CommonModule,
    DuplicateIpRoutingModule,
    FileUploadModule,
  ]
})
export class DuplicateIpModule { }
