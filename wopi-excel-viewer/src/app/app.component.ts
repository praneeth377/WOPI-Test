import { HttpClient } from '@angular/common/http';
import { Component, OnInit } from '@angular/core';
import { RouterOutlet } from '@angular/router';

import { SafeUrlPipe } from './pipes/safe-url.pipe';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, SafeUrlPipe],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})

export class AppComponent implements OnInit {
  excelUrl!: string;

  constructor(private http: HttpClient) {}

  ngOnInit() {
    this.getExcelUrl();
  }

  getExcelUrl() {
    this.http.get<any>('http://localhost:3000/generate-token').subscribe(
        response => {
          this.excelUrl = response.excelUrl;
        },
        error => {
          console.error('Error fetching excelUrl:', error);
        }
      );
  }
  
}
