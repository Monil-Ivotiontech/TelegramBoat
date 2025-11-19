import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { MtUsersComponent } from './mt-users.component';

const routes: Routes = [];

@NgModule({
  imports: [RouterModule.forChild([{ path: '', component: MtUsersComponent }])],
  exports: [RouterModule]
})
export class MtUsersRoutingModule { }
