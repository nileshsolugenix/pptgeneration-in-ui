import { PPT_COLORS, PPT_FontFace, PPT_MASTER_SLIDES } from "../enums/ppt.enums";

export const PptslideConfig = (): any => ({
    layout: 'LAYOUT_WIDE',
    masterSlide: {
        title: PPT_MASTER_SLIDES.MASTER_SLIDE,
        background: {color: PPT_COLORS.WHITE},
        objects: [
            {
                image: {
                    x: 0.41,
                    y: 6.88,
                    h: 0.28,
                    w: 1.78,
                    path: 'assets/images/solugenix.png',
                }
            }
        ],
        slideNumber: {
            x: 12.74,
            y: 6.98,
            h: 0.32,
            w: 0.5,
            fontSize: 12,
            color: PPT_COLORS.BLACK,
            FontFace: PPT_FontFace.ARIAL  
        }
    },
    slide1Content: {
        imageOption: {
            x: 0.41,
            y: 6.88,
            h: 0.28,
            w: 1.78,
            path: 'assets/images/solugenix.png',
        },
        textOptions: {
            x: 0.41,
            y: 1.8,
            h: 1.19,
            w: 10.9,
            fontSize: 40,
            bold: true,
            color: PPT_COLORS.BLACK,
            FontFace: PPT_FontFace.ARIAL
        }
    },
    slide2Content: {
        titleOptions: {
            x: 0.33,
            y: 0.21,
            h: 0.57,
            w: 10.9,
            fontSize: 22,
            bold: true,
            color: PPT_COLORS.BLUE,
            FontFace: PPT_FontFace.ARIAL
        },
        chartOptions: {
            x: 1, 
            y: 1.9, 
            w: 8, 
            h: 4 
        }
    },
    dynamciText: {
        textsConfig: [
            {x: 2, y: 4},
            {x: 4, y: 4.5}
        ]
    }
})