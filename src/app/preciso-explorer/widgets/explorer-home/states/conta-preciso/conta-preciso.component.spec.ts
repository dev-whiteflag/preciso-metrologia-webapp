import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { ContaPrecisoComponent } from './conta-preciso.component';

describe('ContaPrecisoComponent', () => {
  let component: ContaPrecisoComponent;
  let fixture: ComponentFixture<ContaPrecisoComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ ContaPrecisoComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(ContaPrecisoComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
