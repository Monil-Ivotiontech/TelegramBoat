import { CUSTOM_ELEMENTS_SCHEMA, NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { ConfigurationRoutingModule } from './configuration-routing.module';
import { AdminComponent } from './admin/admin.component';
import { ReactiveFormsModule, FormsModule } from '@angular/forms';
import { ButtonModule } from 'primeng/button';
import { DropdownModule } from 'primeng/dropdown';
import { InputTextModule } from 'primeng/inputtext';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { TableModule } from 'primeng/table';
import { TabViewModule } from 'primeng/tabview';
import { ToastModule } from 'primeng/toast';
import { DialogModule } from 'primeng/dialog';
import { MtUsersComponent } from './mt-users/mt-users.component';
import { TransactionComponent } from './transaction/transaction.component';
import { SelectButtonModule } from 'primeng/selectbutton';
import { RadioButtonModule } from 'primeng/radiobutton';
import { TestingComponent } from './testing/testing.component';
import { InputNumberModule } from 'primeng/inputnumber';
import { DuplicateIpComponent } from './duplicate-ip/duplicate-ip.component';
import { FileUploadModule } from 'primeng/fileupload';
import { PanelModule } from 'primeng/panel';
import { CalendarModule } from 'primeng/calendar';
import { InputSwitchModule } from 'primeng/inputswitch';
import { MultiSelectModule } from 'primeng/multiselect';
import { TagInputModule } from 'ngx-chips';



@NgModule({
  declarations: [
    AdminComponent,
    MtUsersComponent,
    TransactionComponent,
    TestingComponent,
    DuplicateIpComponent
  ],
  imports: [
    CommonModule,
    ConfigurationRoutingModule,
    ToastModule,
    ProgressSpinnerModule,
    TabViewModule,
    TableModule,
    ButtonModule,
    DropdownModule,
    ReactiveFormsModule,
    FormsModule,
    InputTextModule,
    DialogModule,
    SelectButtonModule,
    RadioButtonModule,
    InputNumberModule,
    FileUploadModule,
    PanelModule,
    CalendarModule,
    InputSwitchModule,
    MultiSelectModule,
    TagInputModule
  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA]
})
export class ConfigurationModule { }
