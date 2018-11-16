import { NgModule } from '@angular/core';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { BrowserModule } from '@angular/platform-browser';

import { AngularFirestoreModule } from 'angularfire2/firestore';
import { AngularFireModule } from 'angularfire2';

import { MatCardModule } from '@angular/material/card';
import { MatTableModule } from '@angular/material/table';
import { MatInputModule } from '@angular/material';
import { MatFormFieldModule } from '@angular/material';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatButtonModule } from '@angular/material/button';

import { ExplorerBodyComponent } from './explorer.body';
import { environment } from '../../../../environments/environment';

@NgModule({
    imports: [
        MatButtonModule,
        MatFormFieldModule,
        MatCheckboxModule,
        MatInputModule,
        MatTableModule,
        MatCardModule,
        BrowserModule,
        BrowserAnimationsModule,
        AngularFireModule.initializeApp(environment.firebase),
        AngularFirestoreModule,
    ],
    declarations: [
        ExplorerBodyComponent,
    ],
    exports: [
        ExplorerBodyComponent
    ],
    bootstrap: [ExplorerBodyComponent],
    providers: []
})

export class ExplorerBodyModule { }