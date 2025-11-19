import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { VerificationRoutingModule } from './verification-routing.module';
import { VerificationComponent } from './verification.component';
import { InputNumberModule } from 'primeng/inputnumber';
import { ButtonModule } from 'primeng/button';
import { RippleModule } from 'primeng/ripple';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { CheckboxModule } from 'primeng/checkbox';
import { ToastModule } from 'primeng/toast';
import { NgOtpInputModule } from 'ng-otp-input';

@NgModule({
    imports: [
        CommonModule,
        VerificationRoutingModule,
        FormsModule,
        InputNumberModule,
        ButtonModule,
        RippleModule,
        CheckboxModule,
        NgOtpInputModule,
        ToastModule,
        ProgressSpinnerModule,
        ReactiveFormsModule
    ],
    declarations: [VerificationComponent]
})
export class VerificationModule { }
