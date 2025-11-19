import { Component } from '@angular/core';
import { LayoutService } from './service/app.layout.service';

@Component({
    selector: 'app-footer',
    templateUrl: './app.footer.component.html'
})
export class AppFooterComponent {
    img_url: any;

    constructor(public layoutService: LayoutService) {
        this.img_url = 'assets/images/vlogo-red.png';
    }

    get logo() {
        return this.layoutService.config.colorScheme === 'light' ? 'dark' : 'white';
    }

}
