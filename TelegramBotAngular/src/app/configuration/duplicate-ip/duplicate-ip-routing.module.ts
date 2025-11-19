import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { DuplicateIpComponent } from './duplicate-ip.component';

const routes: Routes = [];

@NgModule({
  imports: [RouterModule.forChild([{ path: '', component: DuplicateIpComponent }])],
  exports: [RouterModule]
})
export class DuplicateIpRoutingModule { }
