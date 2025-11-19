import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { MtUsersRoutingModule } from './mt-users-routing.module';
import { CalendarModule } from 'primeng/calendar';
import { StyleClassModule } from 'primeng/styleclass';


@NgModule({
  declarations: [],
  imports: [
    CommonModule,
    MtUsersRoutingModule,
    CalendarModule,
    StyleClassModule
  ]
})
export class MtUsersModule { }
