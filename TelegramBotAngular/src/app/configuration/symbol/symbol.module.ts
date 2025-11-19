import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { SymbolRoutingModule } from './symbol-routing.module';
import { SymbolComponent } from './symbol.component';
import { ToastModule } from 'primeng/toast';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { TabViewModule } from 'primeng/tabview';
import { TableModule } from 'primeng/table';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { DialogModule } from 'primeng/dialog';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { DropdownModule } from 'primeng/dropdown';


@NgModule({
  declarations: [
    SymbolComponent
  ],
  imports: [
    CommonModule,
    SymbolRoutingModule,
    ToastModule,
    ProgressSpinnerModule,
    TabViewModule,
    TableModule,
    ButtonModule,
    InputTextModule,
    DialogModule,
    ReactiveFormsModule,
    FormsModule,
    DropdownModule
  ]
})
export class SymbolModule { }
