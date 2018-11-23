import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';

import { AppComponent } from './explorer.main';
import { UserFormComponent } from './widgets/explorer-login/users/login-form/login-form.component'

import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { AuthService } from "./widgets/explorer-login/core/auth.service";
import { AngularFireAuth } from "angularfire2/auth"
import { RouterModule, Routes } from '@angular/router';
import { MatCardModule } from '@angular/material/card';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatButtonModule } from '@angular/material/button';
import { HeaderComponent } from './widgets/explorer-home/header/header.component';
import { FooterComponent } from './widgets/explorer-home/footer/footer.component';
import { MatToolbarModule } from '@angular/material/toolbar';
import { ExplorerHomeComponent } from './widgets/explorer-home/explorer-home.component';
import { MatIconModule } from '@angular/material/icon';
import { CertificadosComponent } from './widgets/explorer-home/states/certificados/certificados.component';
import { MatTabsModule } from '@angular/material/tabs';
import { MatTableModule } from '@angular/material/table';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { AngularFirestore } from 'angularfire2/firestore';
import { AngularFireModule } from 'angularfire2';
import { environment } from '../../environments/environment';
import { ContaPrecisoComponent } from './widgets/explorer-home/states/conta-preciso/conta-preciso.component';
import {MatExpansionModule} from '@angular/material/expansion';

const routes: Routes = [
  { path: '', redirectTo: 'login', pathMatch: 'full' },
  { path: 'home', component: ExplorerHomeComponent },
  { path: 'login', component: UserFormComponent, },
];

@NgModule({
  declarations: [
    AppComponent,
    UserFormComponent,
    HeaderComponent,
    FooterComponent,
    ExplorerHomeComponent,
    CertificadosComponent,
    ContaPrecisoComponent,
  ],
  imports: [
    BrowserModule,
    BrowserAnimationsModule,
    FormsModule,
    ReactiveFormsModule,
    RouterModule.forRoot(routes),
    MatCardModule,
    MatFormFieldModule,
    MatInputModule,
    MatButtonModule,
    MatToolbarModule,
    MatIconModule,
    MatTabsModule,
    MatFormFieldModule,
    MatCheckboxModule,
    MatTableModule,
    AngularFireModule.initializeApp(environment.firebase),
    MatExpansionModule,
  ],
  exports: [],
  providers: [AuthService, AngularFireAuth, RouterModule, AngularFirestore],
  bootstrap: [AppComponent]
})
export class AppModule { }
