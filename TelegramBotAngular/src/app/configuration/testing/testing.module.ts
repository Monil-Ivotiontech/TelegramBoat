import { NgModule, CUSTOM_ELEMENTS_SCHEMA } from '@angular/core';
import { CommonModule } from '@angular/common';

import { TestingRoutingModule } from './testing-routing.module';
import { StyleClassModule } from 'primeng/styleclass';
import { InputSwitchModule } from 'primeng/inputswitch';


@NgModule({
  declarations: [],
  imports: [
    CommonModule,
    TestingRoutingModule,
    StyleClassModule,
    InputSwitchModule
  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA]
})
export class TestingModule { }
