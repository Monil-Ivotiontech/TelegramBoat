import { Component, ElementRef, ViewChild, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { MessageService } from 'primeng/api';
import { Subscription } from 'rxjs';
import { LayoutService } from 'src/app/layout/service/app.layout.service';
import { EndPoints, ApiMethod } from '../core/const';
import { HttpService } from '../core/service/http.service';
import { StorageService } from '../core/service/storage.service';
import { AppSidebarComponent } from './app.sidebar.component';

@Component({
    selector: 'app-topbar',
    templateUrl: './app.topbar.component.html',
    providers: [MessageService]

})
export class AppTopbarComponent implements OnInit {

    @ViewChild('menubutton') menuButton!: ElementRef;

    @ViewChild(AppSidebarComponent) appSidebar!: AppSidebarComponent;

    user_details: any;
    loginUser!: string;
    sub = new Subscription();
    img_url: any;

    constructor(public layoutService: LayoutService, public el: ElementRef,
        private router: Router,
        private storageService: StorageService,
        private httpService: HttpService,
        private messageService: MessageService) { }

    ngOnInit(): void {
        this.img_url = 'assets/images/vlogo-red.png';

        this.user_details = this.storageService.getLocalObject('userdetails');
        // this.login_admin_cookies();
        this.loginUser = this.user_details.firstname + ' ' + this.user_details.lastname;
    }


    onMenuButtonClick() {
        this.layoutService.onMenuToggle();
    }

    onProfileButtonClick() {
        this.layoutService.showRightMenu();
    }

    onSearchClick() {
        this.layoutService.toggleSearchBar();
    }

    onRightMenuClick() {
        this.layoutService.showRightMenu();
    }

    get logo() {
        const logo = this.layoutService.config.menuTheme === 'white' || this.layoutService.config.menuTheme === 'orange' ? 'dark' : 'white';
        return logo;
    }

    onConfigButtonClick() {
        this.layoutService.showConfigSidebar();
    }

    // profile section opens
    profile() {
        this.router.navigate(['/profile/create'])
    }

    // Logout from session by clearing cache and cookies 
    logout() {
        let userid = this.user_details.userid;
        let email = this.user_details.email;
        const credentials = { userid: userid, email: email };
        this.sub = this.httpService.frontendRequestCall(EndPoints.admin_logout, ApiMethod.POST, credentials)
            .subscribe(response => {
                response = response || {};
                let message = response.message || '';
                let status = response.status || 'false';
                if (status == 'true') {
                    this.storageService.removeCookies();
                    this.storageService.removeLocalObject('userdetails');
                    this.router.navigate(['auth/login']);
                    this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
                }
                else {
                    this.messageService.add({ severity: 'error', summary: 'Error', detail: message });
                }
            },
                error => {
                    this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Something went wrong!' });
                }
            );
    }

}
