import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { ExplorerHomeComponent } from './explorer-home.component';

describe('ExplorerHomeComponent', () => {
  let component: ExplorerHomeComponent;
  let fixture: ComponentFixture<ExplorerHomeComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ ExplorerHomeComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(ExplorerHomeComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
