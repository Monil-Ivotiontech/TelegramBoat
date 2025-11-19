import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { RegisterUserConfigComponent } from './register-user-config.component';

const routes: Routes = [];

@NgModule({
  imports: [RouterModule.forChild([{ path: '', component: RegisterUserConfigComponent },])],
  exports: [RouterModule]
})
export class RegisterUserConfigRoutingModule { }
