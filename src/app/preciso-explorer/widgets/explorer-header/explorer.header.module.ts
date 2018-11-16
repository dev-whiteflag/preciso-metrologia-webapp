import { NgModule } from '@angular/core';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { BrowserModule } from '@angular/platform-browser';

import { MatToolbarModule } from '@angular/material/toolbar';
import { MatButtonModule } from '@angular/material/button';

import { ExplorerHeaderComponent } from './explorer.header';

@NgModule({
    imports: [
        MatButtonModule,
        MatToolbarModule,
        BrowserModule,
        BrowserAnimationsModule,
    ],
    declarations: [
        ExplorerHeaderComponent,
    ],
    exports: [
        ExplorerHeaderComponent
    ],
    bootstrap: [ExplorerHeaderComponent],
    providers: []
})

export class ExplorerHeaderModule { }