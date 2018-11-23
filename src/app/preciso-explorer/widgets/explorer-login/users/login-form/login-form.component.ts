import { Component, OnInit } from '@angular/core';
import { AuthService } from "../../core/auth.service";
import { ReactiveFormsModule, FormGroup, FormBuilder, Validators } from '@angular/forms';

@Component({
  selector: 'login-form',
  templateUrl: './login-form.component.html',
  styleUrls: ['./login-form.component.css']
})
export class UserFormComponent implements OnInit {

  userForm: FormGroup;
  newUser: boolean = true; // to toggle login or signup form
  passReset: boolean = false;

  constructor(private fb: FormBuilder, private auth: AuthService) {}

   ngOnInit(): void {
     this.buildForm();
   }

   toggleForm(): void {
     this.newUser = !this.newUser;
   }

   login(): void {
     this.auth.emailLogin(this.userForm.value.email, this.userForm.value.password);
   }

   resetPassword() {
     this.auth.resetPassword(this.userForm.value['email'])
     .then(() => this.passReset = true)
   }

   buildForm(): void {
     this.userForm = this.fb.group({
       'email': ['', [
           Validators.required,
           Validators.email
         ]
       ],
       'password': ['', [
         Validators.pattern('^(?=.*[0-9])(?=.*[a-zA-Z])([a-zA-Z0-9]+)$'),
         Validators.minLength(6),
         Validators.maxLength(25)
       ]
     ],
     });

     this.userForm.valueChanges.subscribe(data => this.onValueChanged(data));
     this.onValueChanged(); // reset validation messages
   }

   // Updates validation state on form changes.
   onValueChanged(data?: any) {
     if (!this.userForm) { return; }
     const form = this.userForm;
     for (const field in this.formErrors) {
       // clear previous error message (if any)
       this.formErrors[field] = '';
       const control = form.get(field);
       if (control && control.dirty && !control.valid) {
         const messages = this.validationMessages[field];
         for (const key in control.errors) {
           this.formErrors[field] += messages[key] + ' ';
         }
       }
     }
   }

  formErrors = {
     'email': '',
     'password': ''
   };

   validationMessages = {
     'email': {
       'required':      'Inserir o Email é Obrigatório',
       'email':         'O Email precisa ser Válido.'
     },
     'password': {
       'required':      'Inserir a Senha é Obrigatório.',
       'pattern':       'A Senha precisa incluir uma letra e um número.',
       'minlength':     'A Senha precisa ter pelomenos 4 caracteres.',
       'maxlength':     'A Senha não pode ser maior que 40 caracteres.',
     }
   };

}