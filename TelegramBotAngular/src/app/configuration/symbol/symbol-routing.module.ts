import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { SymbolComponent } from './symbol.component';

const routes: Routes = [];

@NgModule({
  imports: [RouterModule.forChild([
    { path: '', component: SymbolComponent }
  ])],
  exports: [RouterModule]
})
export class SymbolRoutingModule { }
