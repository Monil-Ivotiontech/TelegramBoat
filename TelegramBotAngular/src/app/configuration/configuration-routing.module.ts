import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';

const routes: Routes = [];

@NgModule({
  imports: [RouterModule.forChild(
    [
      { path: 'platform', data: { breadcrumb: 'Platform Users' }, loadChildren: () => import('./admin/admin.module').then(m => m.AdminModule) },
      { path: 'symbol', data: { breadcrumb: 'Symbol' }, loadChildren: () => import('./symbol/symbol.module').then(m => m.SymbolModule) },
      { path: 'mt-users', data: { breadcrumb: 'MT Users' }, loadChildren: () => import('./mt-users/mt-users.module').then(m => m.MtUsersModule) },
      { path: 'transactions', data: { breadcrumb: 'Transaction' }, loadChildren: () => import('./transaction/transaction.module').then(m => m.TransactionModule) },
      { path: 'testing', data: { breadcrumb: 'Testing' }, loadChildren: () => import('./testing/testing.module').then(m => m.TestingModule) },
      { path: 'duplicate-ip', data: { breadcrumb: 'Duplicate IP' }, loadChildren: () => import('./duplicate-ip/duplicate-ip.module').then(m => m.DuplicateIpModule) },
    ]
  )],
  exports: [RouterModule]
})
export class ConfigurationRoutingModule { }

