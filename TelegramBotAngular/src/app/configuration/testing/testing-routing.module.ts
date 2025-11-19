import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { TestingComponent } from './testing.component';

const routes: Routes = [];

@NgModule({
  imports: [RouterModule.forChild([{ path: '', component: TestingComponent }])],
  exports: [RouterModule]
})
export class TestingRoutingModule { }
