import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';

import { AppComponent } from './explorer.main';
import { ExplorerHeaderModule } from './widgets/explorer-header/explorer.header.module';
import { ExplorerBodyModule } from './widgets/explorer-body/explorer.body.module';
import { UserFormComponent } from './widgets/explorer-login/users/login-form/login-form.component'

import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { AuthService } from "./widgets/explorer-login/core/auth.service";
import { AngularFireAuth } from "angularfire2/auth"
import { RouterModule, Routes } from '@angular/router';
import { MatCardModule } from '@angular/material/card';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatButtonModule } from '@angular/material/button';

const routes: Routes = [
  { path: 'home', component: ExplorerBodyModule },
  { path: 'login', component: UserFormComponent },
];

@NgModule({
  declarations: [
    AppComponent,
    UserFormComponent,
  ],
  imports: [
    ExplorerHeaderModule,
    ExplorerBodyModule,
    BrowserModule,
    BrowserAnimationsModule,
    FormsModule,
    ReactiveFormsModule,
    RouterModule.forRoot(routes),
    MatCardModule,
    MatFormFieldModule,
    MatInputModule,
    MatButtonModule,
  ],
  exports: [],  
  providers: [AuthService, AngularFireAuth, RouterModule],
  bootstrap: [AppComponent]
})
export class AppModule { }
