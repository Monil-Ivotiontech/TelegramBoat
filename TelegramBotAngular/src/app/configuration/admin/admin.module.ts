import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { AdminRoutingModule } from './admin-routing.module';
import { ToastModule } from 'primeng/toast';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { PanelModule } from 'primeng/panel';
import { ToolbarModule } from 'primeng/toolbar';
import { ButtonModule } from 'primeng/button';
import { DropdownModule } from 'primeng/dropdown';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { RegisterUserConfigComponent } from './register-user-config/register-user-config.component';
import { InputTextModule } from 'primeng/inputtext';


@NgModule({
  declarations: [

    RegisterUserConfigComponent
  ],
  imports: [
    CommonModule,
    AdminRoutingModule,
    ToastModule,
    ProgressSpinnerModule,
    ToolbarModule,
    PanelModule,
    ButtonModule,
    DropdownModule,
    ReactiveFormsModule,
    FormsModule,
    InputTextModule
  ]
})
export class AdminModule { }
