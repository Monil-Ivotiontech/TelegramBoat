import { OnInit } from '@angular/core';
import { Component } from '@angular/core';

@Component({
    selector: 'app-menu',
    templateUrl: './app.menu.component.html'
})
export class AppMenuComponent implements OnInit {

    model: any[] = [];

    ngOnInit() {
        this.model = [
            {
                label: 'Configuration',
                icon: 'fa-duotone fa-cog',
                items: [
                    {
                        label: 'Platform Users',
                        icon: 'fa-solid fa-user-gear',
                        routerLink: ['configuration/platform']
                    },
                    {
                        label: 'Symbol',
                        icon: 'fa-solid fa-symbols',
                        routerLink: ['configuration/symbol']
                    },
                    {
                        label: 'MT Users',
                        icon: 'fa-solid fa-arrows-down-to-people',
                        routerLink: ['configuration/mt-users']
                    },
                    {
                        label: 'transactions',
                        icon: 'fa-solid fa-hand-holding-dollar',
                        routerLink: ['configuration/transactions']
                    },
                    // {
                    //     label: 'Testing',
                    //     icon: 'fa-duotone fa-chart-scatter',
                    //     routerLink: ['configuration/testing']
                    // },
                    {
                        label: 'Duplicate IP',
                        icon: 'fa-solid fa-location-dot',
                        routerLink: ['configuration/duplicate-ip']
                    },
                ]
            },
        ];
    }
}
