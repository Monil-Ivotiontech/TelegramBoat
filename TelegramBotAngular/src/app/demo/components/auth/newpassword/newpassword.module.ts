import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { NewPasswordComponent } from './newpassword.component';
import { NewPasswordRoutingModule } from './newpassword-routing.module';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { RippleModule } from 'primeng/ripple';
import { AppConfigModule } from 'src/app/layout/config/app.config.module';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { ToastModule } from 'primeng/toast';

@NgModule({
    imports: [
        CommonModule,
        NewPasswordRoutingModule,
        FormsModule,
        ButtonModule,
        InputTextModule,
        RippleModule,
        AppConfigModule,
        ReactiveFormsModule,
        ProgressSpinnerModule,
        ToastModule,
    ],
    declarations: [NewPasswordComponent]
})
export class NewPasswordModule { }
