import { Component, OnInit } from '@angular/core';
import { AuthService } from '../../explorer-login/core/auth.service';
import { Router } from '@angular/router';

@Component({
  selector: 'app-header',
  templateUrl: './header.component.html',
  styleUrls: ['./header.component.css']
})
export class HeaderComponent implements OnInit {

  constructor(private auth: AuthService) {}
    
   signOut(): void{
    this.auth.signOut();
   }

  ngOnInit() {  }
}
