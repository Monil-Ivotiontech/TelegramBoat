import { ComponentFixture, TestBed } from '@angular/core/testing';

import { RegisterUserConfigComponent } from './register-user-config.component';

describe('RegisterUserConfigComponent', () => {
  let component: RegisterUserConfigComponent;
  let fixture: ComponentFixture<RegisterUserConfigComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ RegisterUserConfigComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(RegisterUserConfigComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
