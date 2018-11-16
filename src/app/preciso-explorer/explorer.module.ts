import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';

import { AppComponent } from './explorer.main';
import { ExplorerHeaderModule } from './widgets/explorer-header/explorer.header.module';
import { ExplorerBodyModule } from './widgets/explorer-body/explorer.body.module';

@NgModule({
  declarations: [
    AppComponent,
  ],
  imports: [
    ExplorerHeaderModule,
    ExplorerBodyModule,

    BrowserModule,
    BrowserAnimationsModule,
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
