import { Component, Injector } from '@angular/core';
import { PptGenerateService } from './services/ppt-generate.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'ppt-generation-in-angular';
  pptgenerateService: PptGenerateService;

  dataChartAreaLine = [
    {
      name: "Actual Sales",
      labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
      values: [1500, 4600, 5156, 3167, 8510, 8009, 6006, 7855, 12102, 12789, 10123, 15121],
    },
    {
      name: "Projected Sales",
      labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
      values: [1000, 2600, 3456, 4567, 5010, 6009, 7006, 8855, 9102, 10789, 11123, 12121],
    },
  ];

  constructor(_injector: Injector) {
    this.pptgenerateService = _injector.get(PptGenerateService);
  }

  generatePPT() {
    this.pptgenerateService.initializePPT();
    this.pptgenerateService.createStaticSlide1();
    this.pptgenerateService.createChartSlide(this.dataChartAreaLine);
    this.pptgenerateService.createTableSlide();
    this.pptgenerateService.addText();
    this.pptgenerateService.stylingTextInPPT();
    this.pptgenerateService.stylingTextInPPTCont();
    this.downloadPPT();
  }

  downloadPPT() {
    const ppt = this.pptgenerateService.downloadPpt();
    ppt.writeFile({fileName: 'PPT-Generation-In-Angular.pptx'});
  }

}
