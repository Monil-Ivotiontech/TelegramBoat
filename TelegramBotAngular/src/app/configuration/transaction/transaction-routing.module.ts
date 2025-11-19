import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { TransactionComponent } from './transaction.component';

const routes: Routes = [];

@NgModule({
  imports: [RouterModule.forChild([{ path: '', component: TransactionComponent }])],
  exports: [RouterModule]
})
export class TransactionRoutingModule { }
