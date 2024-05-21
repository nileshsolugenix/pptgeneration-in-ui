import { Injectable } from '@angular/core';
import pptxgen from 'pptxgenjs';
import { PptslideConfig } from '../config/ppt-slide-config';

@Injectable({
  providedIn: 'root'
})
export class PptGenerateService {
  ppt!: pptxgen;
  slide!: pptxgen.Slide;
  config: any;

  constructor() { }

  initializePPT() {
    this.config = PptslideConfig();
    this.ppt = new pptxgen();
    this.ppt.layout = this.config.layout;
    this.createMasterSlide();
  }

  downloadPpt() {
    return this.ppt;
  }

  createMasterSlide() {
    this.ppt.defineSlideMaster(this.config.masterSlide);
  }

  createStaticSlide1() {
    this.slide = this.ppt.addSlide();
    this.slide.addImage(this.config.slide1Content.imageOption);
    this.slide.addText('Welcome To PPT Generation Using PPTXGENJS', this.config.slide1Content.textOptions);
  }

  createChartSlide(chartDataSet: any) {
    this.ppt.defineSlideMaster(this.config.masterSlide);
    this.slide = this.ppt.addSlide({masterName: this.config.masterSlide.title});
    this.slide.addText('Line Chart', this.config.slide2Content.titleOptions);
    this.slide.addChart(this.ppt.ChartType.line, chartDataSet, this.config.slide2Content.chartOptions);
  }

  createTableSlide() {
    const arrTabRows1: any = [
      [
        { text: "Top Lft", options: { valign: "top", align: "left", fontFace: "Arial" } },
        { text: "Top Ctr", options: { valign: "top", align: "center", fontFace: "Courier" } },
        { text: "Top Rgt", options: { valign: "top", align: "right", fontFace: "Verdana" } },
      ],
      [
        { text: "Mdl Lft", options: { valign: "middle", align: "left" } },
        { text: "Mdl Ctr", options: { valign: "middle", align: "center" } },
        { text: "Mdl Rgt", options: { valign: "middle", align: "right" } },
      ],
      [
        { text: "Btm Lft", options: { valign: "bottom", align: "left" } },
        { text: "Btm Ctr", options: { valign: "bottom", align: "center" } },
        { text: "Btm Rgt", options: { valign: "bottom", align: "right" } },
      ],
    ];

    const arrTabRowsWithHeader: any = [
      [
        { text: "Name", options: { valign: "middle", align: "center", fontFace: "Arial", fill: {color: '#ADD8E6'} } },
        { text: "Technology", options: { valign: "middle", align: "center", fontFace: "Arial", fill: {color: '#ADD8E6'} } },
        { text: "Experience", options: { valign: "middle", align: "center", fontFace: "Arial", fill: {color: '#ADD8E6'} } },
      ],
      [
        { text: "Nilesh Pawar", options: { valign: "middle", align: "center", fontFace: "Arial" } },
        { text: "Angular", options: { valign: "middle", align: "center", fontFace: "Arial" } },
        { text: "10", options: { valign: "middle", align: "center", fontFace: "Arial" } },
      ],
      [
        { text: "Mayank Porwal", options: { valign: "middle", align: "center", fontFace: "Arial" } },
        { text: "Java", options: { valign: "middle", align: "center", fontFace: "Arial" } },
        { text: "14", options: { valign: "middle", align: "center", fontFace: "Arial" } },
      ],
      [
        { text: "Arvind Pentum", options: { valign: "middle", align: "center", fontFace: "Arial" } },
        { text: "Angular", options: { valign: "middle", align: "center", fontFace: "Arial" } },
        { text: "3", options: { valign: "middle", align: "center",fontFace: "Arial" } },
      ],
    ];

    this.ppt.defineSlideMaster(this.config.masterSlide);
    this.slide = this.ppt.addSlide({masterName: this.config.masterSlide.title});
    this.slide.addText('Data Table', this.config.slide2Content.titleOptions);
    this.slide.addTable(arrTabRows1, {
      x: 0.5,
      y: 1.1,
      w: 5.0,
      rowH: 0.75,
      fill: { color: "F7F7F7" },
      fontSize: 14,
      color: "363636",
      border: {pt: 1, color: '#000000'}
    });

    this.slide.addTable(arrTabRowsWithHeader, {
      x: 0.5,
      y: 3.9,
      w: 5.0,
      rowH: 0.40,
      fill: { color: "F7F7F7" },
      fontSize: 14,
      color: "363636",
      border: {pt: 1, color: '#000000'}
    });
  }

  addText() {
    this.ppt.defineSlideMaster(this.config.masterSlide);
    this.slide = this.ppt.addSlide({masterName: this.config.masterSlide.title});
    this.slide.addText('Adding Text Dynamically', this.config.slide2Content.titleOptions);
    this.slide.addText('Sample Text', { x: 0.5, y: 1.9, w: 5, color: "FF0000", fontSize: 54 });
    this.slide.addText(
      [
        { text: "word-level", options: { fontSize: 20, color: "99ABCC", align: "right", breakLine: true } },
        { text: "formatting", options: { fontSize: 20, color: "FFFF00", align: "center" } }, 
      ],
      { x: 0.5, y: 4.1, w: 8.5, h: 2.0, fill: { color: "F1F1F1" } }
    );

    const text =  [
      { text: "word-level"},
      { text: "formatting"}, 
    ]
    // text.forEach((obj: any) => {
    //   this.slide.addText(obj.text, dynamciText.textsConfig[i]);
    // })
  }

  stylingTextInPPT() {
    this.ppt.defineSlideMaster(this.config.masterSlide);
    this.slide = this.ppt.addSlide({masterName: this.config.masterSlide.title});
    this.slide.addText('Styling Text In PPT', this.config.slide2Content.titleOptions);
    // Line break character
    this.slide.addText("Line 1\nLine 2\nLine 3", { x: 0.8, y: 2.4, w: "30%", h: 1, fill: { color: "F2F2F2" }, bullet: { type: "number" } });
    // Bullets can also be applied on a per-line level
    this.slide.addText(
        [
            { text: "I have a star bullet", options: { bullet: { code: "2605" }, color: "CC0000" } },
            { text: "I have a triangle bullet", options: { bullet: { code: "25BA" }, color: "00CD00" } },
            { text: "no bullets on this line", options: { fontSize: 12 } },
            { text: "I have a normal bullet", options: { bullet: true, color: "0000AB" } },
        ],
        { x: 0.8, y: 5.0, w: "30%", h: 1.4, color: "ABABAB", margin: 1 }
    );
  }

  stylingTextInPPTCont() {
    this.ppt.defineSlideMaster(this.config.masterSlide);
    this.slide = this.ppt.addSlide({masterName: this.config.masterSlide.title});
    this.slide.addText('Styling Text In PPT Cont.', this.config.slide2Content.titleOptions);
    let arrTextObjs1 = [
        { text: "1st line", options: { fontSize: 24, color: "99ABCC", breakLine: true } },
        { text: "2nd line", options: { fontSize: 36, color: "FFFF00", breakLine: true } },
        { text: "3rd line", options: { fontSize: 48, color: "0088CC" } },
    ];
    this.slide.addText(arrTextObjs1, { x: 0.5, y: 1, w: 8, h: 2, fill: { color: "232323" } });
    
    let arrTextObjs2 = [
        { text: "1st line", options: { fontSize: 24, color: "99ABCC", breakLine: false } },
        { text: "2nd line", options: { fontSize: 36, color: "FFFF00", breakLine: false } },
        { text: "3rd line", options: { fontSize: 48, color: "0088CC" } },
    ];
    this.slide.addText(arrTextObjs2, { x: 0.5, y: 4, w: 8, h: 2, fill: { color: "232323" } });
  }
}
