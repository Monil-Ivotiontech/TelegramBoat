import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { AdminComponent } from './admin.component';

const routes: Routes = [];

@NgModule({
  imports: [RouterModule.forChild([
    { path: '', component: AdminComponent },
    { path: 'register-new-user', data: { breadcrumb: 'Register User' }, loadChildren: () => import('./register-user-config/register-user-config.module').then(m => m.RegisterUserConfigModule) },
  ])],
  exports: [RouterModule]
})
export class AdminRoutingModule { }
