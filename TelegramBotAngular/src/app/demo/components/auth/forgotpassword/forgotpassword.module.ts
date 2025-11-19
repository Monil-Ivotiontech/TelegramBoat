import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ForgotPasswordRoutingModule } from './forgotpassword-routing.module';
import { ForgotPasswordComponent } from './forgotpassword.component';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { AppConfigModule } from 'src/app/layout/config/app.config.module';
import { ReactiveFormsModule } from '@angular/forms';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { ToastModule } from 'primeng/toast';

@NgModule({
    imports: [
        CommonModule,
        ButtonModule,
        InputTextModule,
        ForgotPasswordRoutingModule,
        AppConfigModule,
        ReactiveFormsModule,
        ProgressSpinnerModule,
        ToastModule,
    ],
    declarations: [ForgotPasswordComponent]
})
export class ForgotPasswordModule { }