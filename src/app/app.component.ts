import { Component } from '@angular/core';
declare const Excel: any;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  colors = ['red', 'blue', 'yellow'];
  color = 'red';

  onChangeColor(color: string) {
    this.color = color;
  }

  onColor() {
    Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = this.color;
      await context.sync();
    });
  }
}
